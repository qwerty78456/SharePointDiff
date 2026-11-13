using System.Diagnostics;
using System.Text.Json;

namespace SharePointDiff.Core;

/// <summary>Exclusive, mirror-level lock held for the whole of a scan or an apply.
/// Windows enforces <see cref="FileShare.None"/> per handle, so it also excludes a second caller in the same process.
/// The file is never deleted: a leftover file after a crash is harmless because the lock is the open handle, not the file.</summary>
public sealed class WorkspaceLock : IDisposable
{
    private readonly FileStream handle;
    private WorkspaceLock(FileStream handle) => this.handle = handle;
    public static WorkspaceLock Acquire(Workspace workspace)
    {
        try { return new(new FileStream(workspace.LockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None)); }
        // ERROR_SHARING_VIOLATION (32) / ERROR_LOCK_VIOLATION (33): another scan or apply owns this mirror.
        catch (IOException ex) when ((ex.HResult & 0xFFFF) is 32 or 33) { throw new OperationException("WorkspaceBusy", workspace.LockPath, ex); }
    }
    public void Dispose() => handle.Dispose();
}

/// <summary>Counts transfers between the start of the final local checks and the committed record, so sign-out can let them finish.</summary>
public static class CommitGate
{
    private static int active;
    public static int Active => Volatile.Read(ref active);
    public static void Enter() => Interlocked.Increment(ref active);
    public static void Exit() => Interlocked.Decrement(ref active);
    public static bool WaitIdle(TimeSpan timeout)
    {
        var watch = Stopwatch.StartNew();
        while (Active > 0) { if (watch.Elapsed >= timeout) return false; Thread.Sleep(20); }
        return true;
    }
}

public static class LibrarySelectionStore
{
    public const string FileName = "library-selection.json";
    /// <summary>GUIDs compare in "D" lower case whatever casing SharePoint returned; anything else lower-invariant.</summary>
    public static string Id(string id) => Guid.TryParse(id, out var guid) ? guid.ToString("D") : id.ToLowerInvariant();
    public static LibrarySelection? Read(string dataRoot)
    {
        string path = Path.Combine(dataRoot, FileName);
        try { return JsonFiles.Read<LibrarySelection>(path); }
        // A damaged file is repaired by the next Compare; until then the agent is simply not ready.
        catch (JsonException ex) { throw new OperationException("AgentNotReady", path, ex); }
    }
    public static Task WriteAsync(Workspace workspace, LibrarySelection selection, CancellationToken ct) =>
        JsonFiles.WriteAsync(Path.Combine(workspace.DataRoot, FileName), selection, ct);
    /// <summary>Known = previously known ∪ listed; Included = ticked ∪ (previously included but not listed this time).</summary>
    public static LibrarySelection Merge(LibrarySelection? previous, IEnumerable<LibraryInfo> listed, IEnumerable<LibraryInfo> ticked, DateTimeOffset comparedAt)
    {
        var listedIds = listed.Select(l => Id(l.Id)).ToHashSet();
        var known = (previous?.Known ?? []).Select(Id).Concat(listedIds).Distinct().Order(StringComparer.Ordinal).ToArray();
        var included = ticked.Select(l => Id(l.Id)).Concat((previous?.Included ?? []).Select(Id).Where(id => !listedIds.Contains(id)))
            .Distinct().Order(StringComparer.Ordinal).ToArray();
        return new(known, included, comparedAt);
    }
}

/// <summary>Monotonic backup sequence numbers, reserved under the mirror lock.</summary>
public static class BackupSequence
{
    private sealed record State(long Next);
    /// <summary>Reserves <paramref name="count"/> numbers and returns the first. Seeds from existing records when the counter file is missing.</summary>
    public static async Task<long> ReserveAsync(Workspace workspace, int count, CancellationToken ct)
    {
        string path = Path.Combine(workspace.DataRoot, "backup-sequence.json");
        State? state;
        try { state = JsonFiles.Read<State>(path); } catch (JsonException) { state = null; }
        long next = state is { Next: > 0 } saved ? saved.Next : 1 + Retention.Records(workspace).Select(r => r.Record?.Sequence ?? 0).DefaultIfEmpty(0).Max();
        if (count > 0) await JsonFiles.WriteAsync(path, new State(next + count), ct);
        return next;
    }
}

/// <summary>Bounded growth: per file, the newest three distinct SharePoint versions among backups are kept and older ones deleted;
/// cache entries of superseded versions are deleted. Anything that might be the only copy of user content is kept.</summary>
public static class Retention
{
    public const int PristineVersionsPerFile = 3;
    private static readonly EnumerationOptions NoLinks = new() { AttributesToSkip = FileAttributes.ReparsePoint, RecurseSubdirectories = false };
    private static readonly System.Text.RegularExpressions.Regex CacheName = new(@"^(?<id>[0-9a-f]{16})\.bin(?<kind>\.verified\.json|\.partial|\.partial\.json)?$");
    internal sealed record Found(string Run, string MetaPath, BackupRecord? Record, string? Problem);
    /// <summary>Every backup-meta record under runs/ (reparse points skipped); unreadable ones come back with a problem.</summary>
    internal static List<Found> Records(Workspace workspace, List<PrunedItem>? failed = null)
    {
        var found = new List<Found>();
        string runs = Path.Combine(workspace.DataRoot, "runs");
        if (!Directory.Exists(runs)) return found;
        foreach (string run in Directory.EnumerateDirectories(runs, "*", NoLinks))
        {
            string meta = Path.Combine(run, "backup-meta");
            try
            {
                if (!Directory.Exists(meta) || (File.GetAttributes(meta) & FileAttributes.ReparsePoint) != 0) continue;
                foreach (string file in Directory.EnumerateFiles(meta, "*.json", NoLinks))
                {
                    try { found.Add(new(run, file, JsonFiles.Read<BackupRecord>(file), null)); }
                    catch (Exception ex) when (ex is JsonException or OperationException { Code: "MetadataStoreCorrupt" })
                    { found.Add(new(run, file, null, "unreadable: " + ex.Message)); }
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            { failed?.Add(new("run", run, null, null, 0, Path.GetFileName(run), null, null, null, null, null, null, null, ex.Message)); }
        }
        return found;
    }
    public static async Task<PruneResult> PruneAsync(Workspace workspace, string currentRun, ScanResult scan, int keep = PristineVersionsPerFile, IReadOnlyDictionary<string, InstalledRecord>? installed = null)
    {
        var deleted = new List<PrunedItem>(); var kept = new List<PrunedItem>(); var failed = new List<PrunedItem>();
        var candidates = new List<(PrunedItem Item, string File, string? Meta, FileHashes? Expected)>();
        // ---- Backups ----
        var eligible = new List<(Found F, BackupRecord R)>();
        foreach (var f in Records(workspace, failed))
        {
            string? reason = f.Problem ?? Ineligible(f);
            if (reason == null) eligible.Add((f, f.Record!));
            else kept.Add(Tomb("backup", f.Record?.BackupPath ?? f.MetaPath, f, reason));
        }
        foreach (var group in eligible.GroupBy(x => SafePaths.Key(x.R.RelativePath!)))
        {
            var retained = new HashSet<string>(StringComparer.Ordinal);
            foreach (var (f, r) in group.OrderByDescending(x => x.R.Sequence))
            {
                bool current = string.Equals(Path.GetFullPath(f.Run), Path.GetFullPath(currentRun), StringComparison.OrdinalIgnoreCase);
                bool duplicate = retained.Contains(r.Hashes!.Sha256);
                if (!duplicate && retained.Count < keep) { retained.Add(r.Hashes.Sha256); continue; }
                if (current) { retained.Add(r.Hashes.Sha256); continue; } // counted, never deleted
                candidates.Add((Tomb("backup", r.BackupPath!, f, duplicate ? "duplicate version" : "older than the newest " + keep + " versions"), r.BackupPath!, f.MetaPath, r.Hashes));
            }
        }
        // ---- Cache ----
        string? cacheSkipped = CacheCandidates(workspace, scan, candidates, kept, installed);
        if (candidates.Count == 0) return new(deleted, kept, failed, cacheSkipped);
        // The audit trail is written before anything is deleted; without it nothing is deleted.
        try { await JsonFiles.WriteAsync(Path.Combine(currentRun, "prune-plan.json"), candidates.Select(c => c.Item).ToArray(), CancellationToken.None); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        { failed.Add(new("plan", Path.Combine(currentRun, "prune-plan.json"), null, null, 0, null, null, null, null, null, null, null, null, ex.Message)); return new(deleted, kept, failed, cacheSkipped); }
        foreach (var (item, file, meta, expected) in candidates)
        {
            try
            {
                if (expected != null)
                {
                    // Delete only the exact bytes the record describes.
                    if (!File.Exists(file) || (File.GetAttributes(file) & FileAttributes.ReparsePoint) != 0) { kept.Add(item with { Reason = "missing or a link" }); continue; }
                    FileHashes actual;
                    try { actual = await FileHashing.ComputeAsync(file, CancellationToken.None); }
                    catch (OperationException ex) { kept.Add(item with { Reason = ex.Code }); continue; }
                    if (actual != expected) { kept.Add(item with { Reason = "bytes differ from the record" }); continue; }
                    File.Delete(file);
                    if (meta != null) File.Delete(meta);
                }
                else File.Delete(file);
                deleted.Add(item);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { failed.Add(item with { Reason = ex.Message }); }
        }
        try { await JsonFiles.WriteAsync(Path.Combine(currentRun, "pruned.json"), new { Deleted = deleted, Failed = failed }, CancellationToken.None); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        { failed.Add(new("report", Path.Combine(currentRun, "pruned.json"), null, null, 0, null, null, null, null, null, null, null, null, ex.Message)); }
        return new(deleted, kept, failed, cacheSkipped);
    }
    /// <summary>Null when the record may be pruned; otherwise why it is kept.</summary>
    private static string? Ineligible(Found f)
    {
        var r = f.Record!;
        if (r.RelativePath == null || r.BackupPath == null || r.Hashes == null || r.Sequence <= 0) return "incomplete record";
        try { SafePaths.Inside(Path.Combine(f.Run, "backups"), r.BackupPath); } catch (OperationException) { return "outside its run"; }
        if (!File.Exists(r.BackupPath)) return "backup missing";
        if ((File.GetAttributes(r.BackupPath) & FileAttributes.ReparsePoint) != 0) return "link";
        if (!File.Exists(Path.Combine(f.Run, "results.json"))) return "run not finished";
        string id = Path.GetFileName(f.MetaPath);
        if (File.Exists(Path.Combine(f.Run, "intents", id)) && !File.Exists(Path.Combine(f.Run, "committed", id))) return "interrupted replacement";
        // Pristine: the replaced bytes were a verified SharePoint version (still in SharePoint's version history).
        var m = r.InstalledMetadata;
        if (m == null || m.Sha256 != r.Hashes.Sha256 || string.IsNullOrEmpty(m.SharePointHash) || m.SharePointHash != r.Hashes.QuickXorHash) return "local edit or unknown origin";
        return null;
    }
    private static PrunedItem Tomb(string kind, string path, Found f, string reason)
    {
        var r = f.Record; var remote = r?.InstalledMetadata?.Remote;
        return new(kind, path, r?.RelativePath, r?.Hashes?.Sha256, r?.Hashes?.Size ?? 0, Path.GetFileName(f.Run), r?.Sequence,
            remote?.DriveId, remote?.ItemId, remote?.ETag, remote?.WebUrl, remote?.Modified, r?.BackedUpAt, reason);
    }
    private static string? CacheCandidates(Workspace workspace, ScanResult scan, List<(PrunedItem, string, string?, FileHashes?)> candidates, List<PrunedItem> kept, IReadOnlyDictionary<string, InstalledRecord>? installed)
    {
        string manifest = Path.Combine(workspace.DataRoot, "installed-manifest.json");
        if (!File.Exists(manifest)) return "no installed manifest yet";
        if (scan.Libraries == null) return "scan without library list";
        if (!Directory.Exists(workspace.CacheRoot)) return null;
        IReadOnlyDictionary<string, InstalledRecord>? records;
        try { records = installed ?? JsonFiles.Read<Dictionary<string, InstalledRecord>>(manifest); } catch (JsonException) { return "manifest unreadable"; }
        // Live: installed versions (offline restore) and everything the scan saw on SharePoint.
        var live = (records?.Values ?? []).Select(r => r.Remote).Concat(scan.Entries.Where(e => e.Remote != null).Select(e => e.Remote!))
            .Select(r => Path.GetFileName(workspace.CachePath(r))[..16]).ToHashSet(StringComparer.Ordinal);
        var complete = scan.Libraries.Where(l => scan.RemoteFolders.Contains(l.LocalFolder) && !scan.Issues.Any(i => i.Path == l.Title))
            .Select(l => LibrarySelectionStore.Id(l.Id)).ToHashSet();
        var files = Directory.EnumerateFiles(workspace.CacheRoot, "*", NoLinks).Select(p => (Path: p, Match: CacheName.Match(Path.GetFileName(p)))).Where(x => x.Match.Success).ToArray();
        foreach (var byId in files.GroupBy(x => x.Match.Groups["id"].Value))
        {
            var kinds = byId.ToDictionary(x => x.Match.Groups["kind"].Value, x => x.Path);
            // A partial without its state can never be resumed; nothing writes it while we hold the lock.
            if (kinds.TryGetValue(".partial", out string? orphan) && !kinds.ContainsKey(".partial.json"))
            { candidates.Add((Cache(orphan, "orphan partial"), orphan, null, null)); kinds.Remove(".partial"); }
            if (live.Contains(byId.Key) || kinds.Count == 0) continue;
            string? library = LibraryOf(kinds);
            if (library == null || !complete.Contains(LibrarySelectionStore.Id(library)))
            { kept.Add(Cache(kinds.Values.First(), library == null ? "no library recorded" : "library not fully read")); continue; }
            foreach (string path in kinds.Values) candidates.Add((Cache(path, "superseded version"), path, null, null));
        }
        return null;
    }
    private static PrunedItem Cache(string path, string reason) =>
        new("cache", path, null, null, File.Exists(path) ? new FileInfo(path).Length : 0, null, null, null, null, null, null, null, null, reason);
    private sealed record LibraryHint(string? LibraryId);
    private static string? LibraryOf(Dictionary<string, string> kinds)
    {
        foreach (string kind in new[] { ".verified.json", ".partial.json" })
            if (kinds.TryGetValue(kind, out string? path))
                try { if (JsonFiles.Read<LibraryHint>(path)?.LibraryId is { Length: > 0 } id) return id; } catch (JsonException) { }
        return null;
    }
}
