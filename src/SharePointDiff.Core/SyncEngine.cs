using System.Collections.Concurrent;
using System.Text.Json;

namespace SharePointDiff.Core;

public sealed class SyncEngine(IRemoteReader remote)
{
    // Deterministic fault/race injection at the actual filesystem boundary, never configured by the UI.
    internal Action<string, string>? Checkpoint { get; init; }
    public async Task<ApplyResult> ApplyAsync(ScanResult scan, SyncPlan plan, IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        using var measurements = IoMetrics.Begin("Apply");
        var workspace = new Workspace(scan.SiteUrl, scan.LocalRoot);
        using var gate = WorkspaceLock.Acquire(workspace);
        string run = workspace.CreateRun();
        await JsonFiles.WriteAsync(Path.Combine(run, "plan.json"), plan, ct);
        var results = new ConcurrentBag<TransferResult>();
        // Conservatively allow staging, new installation, and the immutable local backup.
        long required = plan.Items.Sum(x => x.Entry.Remote!.Size) * 2 + plan.BackupBytes;
        var drive = new DriveInfo(Path.GetPathRoot(scan.LocalRoot)!);
        if (drive.AvailableFreeSpace < required) throw new OperationException("DiskFull");
        await using var manifest = await ManifestSession.OpenAsync(workspace, ct);
        await manifest.ReserveAsync(run, plan.Items.Count, ct);
        await remote.PrepareApplyAsync(plan.Items.Select(w => w.Entry.Remote!).ToArray(), ct);
        // Backup order comes from a counter kept under the lock, never from the wall clock.
        long sequence = await BackupSequence.ReserveAsync(workspace, plan.Replacements, ct) - 1;
        long NextSequence() => Interlocked.Increment(ref sequence);
        int done = 0;
        using var stopOnFault = CancellationTokenSource.CreateLinkedTokenSource(ct);
        var workToken = stopOnFault.Token;
        using var parallel = new SemaphoreSlim(2, 2);
        var cacheLocks = new ConcurrentDictionary<string, SemaphoreSlim>();
        var reuseSources = plan.Items.GroupBy(x => ContentGrouping.Key(x.Entry.Remote!))
            .ToDictionary(g => g.Key, g => g.Where(x => x.Entry.ReusePath != null && File.Exists(x.Entry.ReusePath))
                .Select(x => new ReuseCandidate(x.Entry.ReusePath!, x.Entry.ReuseSha256)).FirstOrDefault());
        string? stopReason = null;
        async Task<TransferResult> FailAsync(WorkItem item, Exception ex, string code)
        {
            string? backup = null;
            if (item.Replace && Directory.Exists(Path.Combine(run, "backups")))
            {
                string candidate = SafePaths.Resolve(Path.Combine(run, "backups"), item.Entry.Local!.RelativePath);
                if (File.Exists(candidate)) backup = candidate;
            }
            var failure = new TransferResult(item.Entry.RelativePath, "Error", backup, item.Entry.Local?.Hashes.Sha256, null, null,
                ex.Data["NetworkBytes"] is long bytes ? bytes : 0, code,
                ex.Data["Attempts"] is int attempts ? attempts : 0, ex.Data["Stalls"] is int stalls ? stalls : 0,
                PreservedPath: ex.Data["PreservedPath"] as string);
            await JsonFiles.WriteAsync(Path.Combine(run, "errors", JsonFiles.StableId(item.Entry.RelativePath) + ".json"),
                new { Path = item.Entry.RelativePath, Code = code, ExceptionType = ex.GetType().Name, Message = ex.Message, Details = ex.Data, ex.StackTrace });
            return failure;
        }
        async Task Execute(WorkItem item)
        {
            await parallel.WaitAsync(workToken);
            try
            {
                TransferResult result;
                try { result = await TransferOneAsync(workspace, run, item, cacheLocks, NextSequence, manifest, reuseSources, progress, workToken); }
                // A cancellation nobody asked for is a timeout: report it as a per-file error, not as "Cancelled".
                catch (OperationCanceledException ex) when (!workToken.IsCancellationRequested) { result = await FailAsync(item, ex, "RequestTimeout"); }
                catch (OperationCanceledException) { result = new(item.Entry.RelativePath, "Cancelled", null, null, null, null, 0); }
                catch (Exception ex) when (ex is OperationException or IOException or UnauthorizedAccessException or System.Net.Http.HttpRequestException)
                {
                    result = await FailAsync(item, ex, ex is OperationException op ? op.Code : ex is IOException or UnauthorizedAccessException ? "FileIoError" : "TransferFailed");
                    // Without a token every remaining file would fail the same way: record it once and stop the run.
                    if (ex is OperationException { Code: "SignInRequired" } && Interlocked.CompareExchange(ref stopReason, "SignInRequired", null) == null) stopOnFault.Cancel();
                }
                results.Add(result);
                progress?.Report(new("Applying", Interlocked.Increment(ref done), plan.Items.Count, item.Entry.RelativePath));
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                ex.Data["Path"] = item.Entry.RelativePath; ex.Data["Operation"] = "Apply";
                // Stop both active transfers and the queue before releasing this worker slot.
                stopOnFault.Cancel();
                throw;
            }
            finally { parallel.Release(); }
        }
        var tasks = plan.Items.Select(Execute).ToArray();
        try { await Task.WhenAll(tasks); } catch (OperationCanceledException) { }
        var completedPaths = results.Select(r => SafePaths.Key(r.Path)).ToHashSet();
        foreach (var item in plan.Items)
            if (!completedPaths.Contains(SafePaths.Key(item.Entry.RelativePath))) results.Add(new(item.Entry.RelativePath, "Cancelled", null, null, null, null, 0));
        var ordered = results.OrderBy(x => x.Path).ToArray();
        var report = new ApplyResult(run, ordered, ct.IsCancellationRequested, stopReason);
        await manifest.FlushAsync();
        // Preserve an audit trail even if the user cancelled the operation.
        await JsonFiles.WriteAsync(Path.Combine(run, "results.json"), report);
        var csv = new List<IEnumerable<object?>> { new object?[] { "Path", "Status", "BackupPath", "OldSha256", "NewSha256", "SharePointHash", "BytesDownloaded", "Error", "Attempts", "Stalls" } };
        csv.AddRange(report.Results.Select(r => new object?[] { r.Path, r.Status, r.BackupPath, r.OldSha256, r.NewSha256, r.SharePointHash, r.BytesDownloaded, r.Error, r.Attempts, r.Stalls }));
        await JsonFiles.CsvAsync(Path.Combine(run, "results.csv"), csv);
        var installedMetadata = new List<RemoteFile>();
        var installedSnapshot = await manifest.SnapshotAsync(CancellationToken.None);
        foreach (var result in report.Results.Where(r => r.Status is "Verified" or "DownloadedUnverified"))
        {
            var installed = installedSnapshot.GetValueOrDefault(SafePaths.Key(result.Path));
            if (installed != null) installedMetadata.Add(installed.Remote);
        }
        await ReportWriter.WriteMetadataCsvAsync(Path.Combine(run, "remote-metadata.csv"), installedMetadata, CancellationToken.None);
        foreach (var semaphore in cacheLocks.Values) semaphore.Dispose();
        // Retention only after a complete, uncancelled run that was not stopped; still under the mirror lock.
        if (!ct.IsCancellationRequested && stopReason == null) report = report with { Pruned = await Retention.PruneAsync(workspace, run, scan, installed: installedSnapshot) };
        await JsonFiles.WriteAsync(Path.Combine(run, "io.json"), measurements.Snapshot());
        return report with { Statistics = measurements.Snapshot() };
    }
    private async Task<TransferResult> TransferOneAsync(Workspace workspace, string run, WorkItem work,
        ConcurrentDictionary<string, SemaphoreSlim> locks, Func<long> nextSequence, ManifestSession manifest,
        IReadOnlyDictionary<string, ReuseCandidate?> reuseSources, IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        using var phase = IoMetrics.Stage("Transfer");
        var entry = work.Entry;
        if (entry.Remote == null || entry.Status is DiffStatus.LocalOnly or DiffStatus.Error or DiffStatus.Unverified or DiffStatus.Match)
            throw new OperationException("NotEligible", entry.RelativePath);
        if (!entry.Remote.Metadata.Complete) throw new OperationException("MetadataIncomplete", entry.RelativePath);
        var file = entry.Remote;
        string destination = SafePaths.Resolve(workspace.LocalRoot, entry.RelativePath);
        if (!work.Replace && (File.Exists(destination) || Directory.Exists(destination))) throw new OperationException("FileChanged", entry.RelativePath);
        if (work.Replace)
        {
            var stat = new FileInfo(destination);
            if (!stat.Exists || stat.Length != entry.Local!.Hashes.Size || stat.LastWriteTimeUtc != entry.Local.LastWriteUtc)
                throw new OperationException("FileChanged", entry.RelativePath);
        }
        destination = SafePaths.Resolve(workspace.LocalRoot, entry.RelativePath, true);
        string install = destination + ".spdiff-" + Guid.NewGuid().ToString("N") + ".tmp";
        string cache = workspace.CachePath(file);
        var cacheLock = locks.GetOrAdd(cache, _ => new SemaphoreSlim(1, 1));
        long downloaded = 0; int attempts = 0, stalls = 0;
        string origin = "Downloaded";
        string? backup = null, preserved = null;
        bool commitAttempted = false;
        bool sourceObserved = false;
        FileHashes? old = null;
        try
        {
            FileHashes verified;
            FileHashes? expectedStage = null;
            await cacheLock.WaitAsync(ct);
            try
            {
                bool staged = false;
                var stamp = File.Exists(cache) ? JsonFiles.Read<CacheStamp>(cache + ".verified.json") : null;
                if (stamp != null && (file.QuickXorHash is { Length: > 0 } || stamp.ETag == file.ETag))
                {
                    var copied = await FileHashing.CopyHashAsync(cache, install, ct);
                    if (Matches(copied, file) && copied.Sha256 == stamp.Sha256) { staged = true; origin = "Cache"; expectedStage = copied; }
                    else File.Delete(install);
                }
                if (!staged)
                {
                    // A cheap Graph-only check prevents needless content transfer of an obsolete plan.
                    var start = await remote.ReadContentStateAsync(file, ct);
                    if (!SameContent(file, start)) throw new OperationException("RemoteChanged", entry.RelativePath);
                    file = start; sourceObserved = true;
                    string partial = cache + ".partial";
                    FileHashes? received = null;
                    if (reuseSources.GetValueOrDefault(ContentGrouping.Key(file)) is { } source && File.Exists(source.Path) && file.QuickXorHash is { Length: > 0 })
                    {
                        received = await FileHashing.CopyHashAsync(source.Path, partial, ct, FileMode.Create);
                        if (Matches(received, file) && (source.Sha256 == null || source.Sha256 == received.Sha256)) origin = "Reused";
                        else { File.Delete(partial); received = null; }
                    }
                    if (received == null)
                    {
                        var transfer = await remote.DownloadAsync(file, partial, progress, ct);
                        downloaded = transfer.NetworkBytes; attempts = transfer.Attempts; stalls = transfer.Stalls;
                        received = transfer.Hashes ?? await FileHashing.ComputeAsync(partial, ct);
                    }
                    if (!Matches(received, file))
                    {
                        File.Move(partial, partial + ".invalid_" + Guid.NewGuid().ToString("N"));
                        throw new OperationException("HashMismatch", entry.RelativePath);
                    }
                    File.Move(partial, cache, true);
                    var copied = await FileHashing.CopyHashAsync(cache, install, ct);
                    if (copied != received) throw new OperationException("HashMismatch", entry.RelativePath);
                    expectedStage = copied;
                    // The source was read back from cache before it is marked verified.
                    await JsonFiles.WriteAsync(cache + ".verified.json", new CacheStamp(copied.Sha256, file.ETag, file.LibraryId, file.DriveId, file.ItemId), ct);
                }
                // Independent readback of the actual staged bytes remains mandatory before touching the original.
                verified = await FileHashing.ComputeAsync(install, ct);
                if (verified != expectedStage || !Matches(verified, file))
                    throw new OperationException("HashMismatch", entry.RelativePath);
            }
            finally { cacheLock.Release(); }
            var current = await remote.RefreshAsync(file, ct);
            if (!SameContent(file, current) || ((sourceObserved || string.IsNullOrEmpty(file.QuickXorHash)) && current.ETag != file.ETag))
                throw new OperationException("RemoteChanged", entry.RelativePath);
            if (!current.Metadata.Complete) throw new OperationException("MetadataIncomplete", entry.RelativePath);
            file = current;
            ct.ThrowIfCancellationRequested();
            SafePaths.AssertDirectoryChain(workspace.LocalRoot, destination);
            using var guard = work.Replace ? new FileStream(destination, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete,
                1024 * 1024, FileOptions.Asynchronous | FileOptions.SequentialScan) : null;
            var guardedIdentity = guard != null ? NativeFileInfo.TryRead(guard) : null;
            if (work.Replace)
            {
                string backupRoot = Path.Combine(run, "backups"); Directory.CreateDirectory(backupRoot);
                backup = SafePaths.Resolve(backupRoot, entry.Local!.RelativePath, true);
                old = await FileHashing.CopyHashAsync(guard!, backup, ct);
                if (old != entry.Local.Hashes) throw new OperationException("FileChanged", entry.RelativePath);
                if (await FileHashing.ComputeAsync(backup, ct) != old) throw new OperationException("BackupFailed", entry.RelativePath);
                DateTime created = File.GetCreationTimeUtc(destination), modified = File.GetLastWriteTimeUtc(destination);
                File.SetCreationTimeUtc(backup, created); File.SetLastWriteTimeUtc(backup, modified);
                var oldMetadata = await manifest.ReadAsync(entry.RelativePath, ct);
                if (oldMetadata?.Sha256 != old.Sha256) oldMetadata = null;
                await JsonFiles.WriteNewAsync(Path.Combine(run, "backup-meta", JsonFiles.StableId(entry.RelativePath) + ".json"),
                    new BackupRecord(entry.RelativePath, backup, old, entry.Local.FullPath, created, modified, oldMetadata, DateTimeOffset.Now, nextSequence()), ct);
                // Keep the final live-path read: the guard allows rename/delete, not just our own ReplaceFile.
                if (await FileHashing.ComputeAsync(destination, ct) != old) throw new OperationException("FileChanged", entry.RelativePath);
                preserved = SafePaths.Resolve(run, "replaced/" + Guid.NewGuid().ToString("N") + ".bin", true);
            }
            SafePaths.AssertDirectoryChain(workspace.LocalRoot, destination);
            await JsonFiles.WriteAsync(Path.Combine(run, "intents", JsonFiles.StableId(entry.RelativePath) + ".json"),
                new CommitIntent(entry.RelativePath, backup, old?.Sha256, verified.Sha256, file, preserved), ct);
            CommitGate.Enter();
            try
            {
                ct.ThrowIfCancellationRequested(); Checkpoint?.Invoke("BeforeReplace", destination); commitAttempted = true;
                if (work.Replace) File.Replace(install, destination, preserved, true);
                else File.Move(install, destination, false);
                if (preserved != null)
                {
                    await using var actual = new FileStream(preserved, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);
                    var actualIdentity = NativeFileInfo.TryRead(actual);
                    bool conflict = guardedIdentity != null && actualIdentity != null
                        ? !NativeFileInfo.SameObject(guardedIdentity, actualIdentity)
                        : await FileHashing.ComputeAsync(preserved, CancellationToken.None) != old;
                    if (conflict)
                        throw new OperationException("CommitPathConflict", entry.RelativePath);
                }
                // From this point cancellation must not interrupt verification or durable audit records.
                File.SetCreationTimeUtc(destination, file.Created.UtcDateTime);
                File.SetLastWriteTimeUtc(destination, file.Modified.UtcDateTime);
                var installedRead = await FileHashing.ReadDetailedAsync(destination, CancellationToken.None, FileShare.Read, false);
                var installed = installedRead.Hashes;
                if (installed != verified) throw new OperationException("HashMismatch", entry.RelativePath);
                var local = new LocalFile(entry.RelativePath, destination, installed, installedRead.LastWriteUtc, HashBasis.Fresh, installedRead.ReadAt, installedRead.Identity);
                var result = new TransferResult(entry.RelativePath, file.QuickXorHash is { Length: > 0 } ? "Verified" : "DownloadedUnverified",
                    backup, old?.Sha256, installed.Sha256, file.QuickXorHash, downloaded, null, attempts, stalls, local, file);
                await manifest.RecordCommitAsync(run, file, installed, result with { InstalledLocal = null, InstalledRemote = null }, origin);
                if (preserved != null) File.Delete(preserved); // exact original retained independently in the verified backup
                return result;
            }
            finally { CommitGate.Exit(); }
        }
        catch (Exception ex)
        {
            ex.Data["NetworkBytes"] = downloaded; ex.Data["Attempts"] = attempts; ex.Data["Stalls"] = stalls;
            if (preserved != null) ex.Data["PreservedPath"] = preserved;
            throw;
        }
        finally
        {
            // Keep artifacts of an attempted commit for recovery; pre-commit staging is disposable.
            if (!commitAttempted && File.Exists(install)) File.Delete(install);
        }
    }
    private static bool SameContent(RemoteFile a, RemoteFile b) => a.ItemId == b.ItemId && a.DriveId == b.DriveId && a.Size == b.Size && a.QuickXorHash == b.QuickXorHash;
    private static bool Matches(FileHashes h, RemoteFile file) => h.Size == file.Size && (string.IsNullOrEmpty(file.QuickXorHash) || h.QuickXorHash == file.QuickXorHash);
    /// <param name="LibraryId">Lets retention delete a superseded entry only when its library was read completely.</param>
    internal sealed record CacheStamp(string Sha256, string ETag, string? LibraryId = null, string? DriveId = null, string? ItemId = null);
}
