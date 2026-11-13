using System.Text.Json;

namespace SharePointDiff.Core;

internal sealed record ManifestEnvelope(int FormatVersion, long LastSequence, Dictionary<string, InstalledRecord>? Records);
internal sealed record CommitJournal(long Sequence, InstalledRecord Installed, FileHashes Hashes, TransferResult Result);
internal sealed record JournalRange(long First, long Last);
internal sealed record CommitIntent(string RelativePath, string? BackupPath, string? OldSha256, string NewSha256, RemoteFile Remote, string? ReplacedPath = null);

/// <summary>Opened under WorkspaceLock. A durable journal precedes every in-memory update.</summary>
public sealed class ManifestSession : IAsyncDisposable
{
    private readonly Workspace workspace;
    private readonly SemaphoreSlim gate = new(1, 1);
    private readonly Dictionary<string, InstalledRecord> records;
    private readonly Dictionary<JsonElement, string> schemaHashes = [];
    private readonly CancellationTokenSource timerStop = new();
    private Task timer = Task.CompletedTask;
    private Exception? timerFailure;
    private long lastSequence, nextSequence;
    private int dirty;
    private bool disposed;
    public List<OperationIssue> RecoveryIssues { get; } = [];
    internal Exception? PrimaryFailure { get; set; }
    private string PathName => Path.Combine(workspace.DataRoot, "installed-manifest.json");
    private ManifestSession(Workspace workspace, Dictionary<string, InstalledRecord> records, long sequence)
    {
        this.workspace = workspace; this.records = records; lastSequence = sequence;
    }
    public static async Task<ManifestSession> OpenAsync(Workspace workspace, CancellationToken ct, bool recover = true)
    {
        var (records, sequence, legacy) = Load(workspace);
        var session = new ManifestSession(workspace, records, sequence);
        try
        {
            if (recover) await session.RecoverAsync(ct);
            if (legacy)
            {
                string backup = session.PathName + ".v1.backup";
                if (!File.Exists(backup))
                {
                    // Preserve the exact legacy bytes before the first format change.
                    string temporary = backup + "." + Guid.NewGuid().ToString("N") + ".tmp";
                    await FileHashing.CopyVerifiedAsync(session.PathName, temporary, ct, FileMode.CreateNew);
                    File.Move(temporary, backup, false);
                }
                else if (await FileHashing.ComputeAsync(backup, ct) != await FileHashing.ComputeAsync(session.PathName, ct))
                    throw new OperationException("MetadataStoreCorrupt", backup);
                session.dirty++;
            }
            session.timer = session.CheckpointTimerAsync();
            return session;
        }
        catch { await session.StopTimerAsync(); session.gate.Dispose(); session.timerStop.Dispose(); throw; }
    }
    internal static (Dictionary<string, InstalledRecord> Records, long Sequence, bool Legacy) Load(Workspace workspace)
    {
        string path = Path.Combine(workspace.DataRoot, "installed-manifest.json");
        IoMetrics.Add("Manifest.Loads");
        var envelope = JsonFiles.Read<ManifestEnvelope>(path);
        if (envelope?.FormatVersion == 2 && envelope.Records != null) return (envelope.Records, envelope.LastSequence, false);
        if (envelope is { FormatVersion: > 2 }) throw new OperationException("MetadataStoreCorrupt", path);
        return (envelope == null ? [] : JsonFiles.Read<Dictionary<string, InstalledRecord>>(path) ?? [], 0, envelope != null);
    }
    public async Task ReserveAsync(string run, int count, CancellationToken ct)
    {
        string counterPath = Path.Combine(workspace.DataRoot, "commit-sequence.json");
        var counter = JsonFiles.Read<JournalRange>(counterPath);
        long highest = Math.Max(lastSequence, counter?.Last ?? 0);
        string runs = Path.Combine(workspace.DataRoot, "runs");
        foreach (string directory in RunDirectories(runs))
            highest = Math.Max(highest, JsonFiles.Read<JournalRange>(Path.Combine(directory, "journal-range.json"))?.Last ?? 0);
        nextSequence = highest;
        await JsonFiles.WriteAsync(counterPath, new JournalRange(highest + 1, checked(highest + count)), ct);
        await JsonFiles.WriteAsync(Path.Combine(run, "journal-range.json"), new JournalRange(highest + 1, checked(highest + count)), ct);
    }
    public async Task<InstalledRecord?> ReadAsync(string path, CancellationToken ct)
    {
        await gate.WaitAsync(ct);
        try { ThrowTimerFailure(); return records.GetValueOrDefault(SafePaths.Key(path)); }
        finally { gate.Release(); }
    }
    public async Task<IReadOnlyDictionary<string, InstalledRecord>> SnapshotAsync(CancellationToken ct)
    {
        await gate.WaitAsync(ct);
        try { ThrowTimerFailure(); return new Dictionary<string, InstalledRecord>(records); }
        finally { gate.Release(); }
    }
    public async Task RecordCommitAsync(string run, RemoteFile remote, FileHashes hashes, TransferResult result, string source)
    {
        await gate.WaitAsync(CancellationToken.None);
        try
        {
            ThrowTimerFailure();
            var installed = new InstalledRecord(remote.RelativePath, hashes.Sha256, remote.QuickXorHash, remote.ETag, DateTimeOffset.Now, remote, source);
            long sequence = ++nextSequence;
            await JsonFiles.WriteAsync(Path.Combine(run, "committed", JsonFiles.StableId(remote.RelativePath) + ".json"),
                new CommitJournal(sequence, installed, hashes, result), CancellationToken.None);
            records[SafePaths.Key(remote.RelativePath)] = installed; lastSequence = Math.Max(lastSequence, sequence); dirty++;
            if (dirty >= 50) await CheckpointInsideAsync();
        }
        finally { gate.Release(); }
    }
    public async Task ObserveMatchesAsync(IEnumerable<DiffEntry> entries, CancellationToken ct)
    {
        await gate.WaitAsync(ct);
        try
        {
            ThrowTimerFailure();
            foreach (var e in entries)
            {
                string key = SafePaths.Key(e.RelativePath);
                if (records.TryGetValue(key, out var previous) && previous.Sha256 == e.Local!.Hashes.Sha256 &&
                    previous.ETag == e.Remote!.ETag && previous.Remote.Metadata.ObservedAt == e.Remote.Metadata.ObservedAt &&
                    JsonElement.DeepEquals(previous.Remote.Basic, e.Remote.Basic) &&
                    JsonElement.DeepEquals(previous.Remote.Metadata.Fields, e.Remote.Metadata.Fields) &&
                    JsonElement.DeepEquals(previous.Remote.Metadata.DisplayValues, e.Remote.Metadata.DisplayValues) &&
                    SchemaHash(previous.Remote.Metadata.Columns) == SchemaHash(e.Remote.Metadata.Columns)) continue;
                string source = records.TryGetValue(key, out var old) && old.Sha256 == e.Local!.Hashes.Sha256 && old.Source != null ? old.Source : "Matched";
                records[key] = new(e.RelativePath, e.Local!.Hashes.Sha256, e.Remote!.QuickXorHash, e.Remote.ETag, DateTimeOffset.Now, e.Remote, source);
                dirty++;
            }
        }
        finally { gate.Release(); }
    }
    private string SchemaHash(JsonElement schema)
    {
        if (!schemaHashes.TryGetValue(schema, out var hash)) schemaHashes[schema] = hash = SchemaStorage.Hash(schema);
        return hash;
    }
    internal async Task PutAsync(InstalledRecord record, CancellationToken ct)
    {
        await gate.WaitAsync(ct);
        try { ThrowTimerFailure(); records[SafePaths.Key(record.Path)] = record; dirty++; }
        finally { gate.Release(); }
    }
    public async Task FlushAsync()
    {
        await gate.WaitAsync(CancellationToken.None);
        try { ThrowTimerFailure(); await CheckpointInsideAsync(); }
        finally { gate.Release(); }
    }
    private async Task CheckpointInsideAsync()
    {
        if (dirty == 0) return;
        try { await JsonFiles.WriteAsync(PathName, new ManifestEnvelope(2, lastSequence, records), CancellationToken.None); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { throw new OperationException("ManifestWriteFailed", PathName, ex); }
        IoMetrics.Add("Manifest.Checkpoints"); dirty = 0;
    }
    private async Task CheckpointTimerAsync()
    {
        using var ticks = new PeriodicTimer(TimeSpan.FromSeconds(5));
        try { while (await ticks.WaitForNextTickAsync(timerStop.Token)) await FlushAsync(); }
        catch (OperationCanceledException) when (timerStop.IsCancellationRequested) { }
        catch (Exception ex) { timerFailure = ex; }
    }
    private void ThrowTimerFailure()
    {
        if (timerFailure != null) throw new OperationException("ManifestWriteFailed", PathName, timerFailure);
    }
    private async Task RecoverAsync(CancellationToken ct)
    {
        var pending = new List<CommitJournal>();
        foreach (string run in RunDirectories(Path.Combine(workspace.DataRoot, "runs")))
        {
            ct.ThrowIfCancellationRequested();
            var range = JsonFiles.Read<JournalRange>(Path.Combine(run, "journal-range.json"));
            if (range == null || range.Last > lastSequence)
            {
                string commits = Path.Combine(run, "committed");
                if (Directory.Exists(commits) && (File.GetAttributes(commits) & FileAttributes.ReparsePoint) == 0)
                    foreach (string path in Directory.EnumerateFiles(commits, "*.json", NoLinks))
                    {
                        var journal = JsonFiles.Read<CommitJournal>(path);
                        if (journal is { Sequence: > 0, Installed: not null, Hashes: not null } && journal.Sequence > lastSequence)
                        {
                            if (journal.Hashes.Sha256 != journal.Installed.Sha256 || !journal.Installed.Remote.Metadata.Complete)
                                throw new OperationException("MetadataStoreCorrupt", path);
                            pending.Add(journal);
                        }
                    }
            }
            string intents = Path.Combine(run, "intents");
            string reviewed = Path.Combine(run, "recovery-reviewed.json");
            if (!Directory.Exists(intents) || File.Exists(reviewed) || (File.GetAttributes(intents) & FileAttributes.ReparsePoint) != 0) continue;
            var observations = new List<object>();
            foreach (string intentPath in Directory.EnumerateFiles(intents, "*.json", NoLinks))
            {
                if (File.Exists(Path.Combine(run, "committed", Path.GetFileName(intentPath)))) continue;
                var intent = JsonFiles.Read<CommitIntent>(intentPath);
                if (intent == null) continue;
                string destination = SafePaths.Resolve(workspace.LocalRoot, intent.RelativePath);
                FileHashes? actual = File.Exists(destination) ? await FileHashing.ComputeAsync(destination, ct) : null;
                observations.Add(new { intent.RelativePath, intent.BackupPath, intent.ReplacedPath, Actual = actual, intent.OldSha256, intent.NewSha256 });
                RecoveryIssues.Add(new("InterruptedCommit", intent.RelativePath, intentPath));
            }
            if (observations.Count > 0) await JsonFiles.WriteAsync(reviewed, observations, ct);
        }
        var duplicate = pending.GroupBy(x => x.Sequence).FirstOrDefault(x => x.Count() > 1);
        if (duplicate != null) throw new OperationException("MetadataStoreCorrupt", workspace.DataRoot);
        foreach (var journal in pending.OrderBy(x => x.Sequence))
        {
            records[SafePaths.Key(journal.Installed.Path)] = journal.Installed;
            lastSequence = journal.Sequence; dirty++; IoMetrics.Add("Manifest.ReplayedCommits");
        }
    }
    private static readonly EnumerationOptions NoLinks = new() { AttributesToSkip = FileAttributes.ReparsePoint, RecurseSubdirectories = false };
    private static IEnumerable<string> RunDirectories(string path) => Directory.Exists(path) ? Directory.EnumerateDirectories(path, "*", NoLinks) : [];
    private async Task StopTimerAsync() { timerStop.Cancel(); await timer; }
    public async ValueTask DisposeAsync()
    {
        if (disposed) return; disposed = true;
        await StopTimerAsync();
        try { await FlushAsync(); }
        catch (Exception ex) when (PrimaryFailure != null) { PrimaryFailure.Data["ManifestCheckpointFailure"] = ex.ToString(); }
        finally { gate.Dispose(); timerStop.Dispose(); }
    }
}
