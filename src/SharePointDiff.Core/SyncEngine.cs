using System.Collections.Concurrent;
using System.Text.Json;

namespace SharePointDiff.Core;

public sealed class SyncEngine(IRemoteReader remote)
{
    public async Task<ApplyResult> ApplyAsync(ScanResult scan, SyncPlan plan, IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        var workspace = new Workspace(scan.SiteUrl, scan.LocalRoot);
        string run = workspace.CreateRun();
        await JsonFiles.WriteAsync(Path.Combine(run, "plan.json"), plan, ct);
        var results = new ConcurrentBag<TransferResult>();
        // Conservatively allow staging, new installation, and the immutable local backup.
        long required = plan.Items.Sum(x => x.Entry.Remote!.Size) * 2 + plan.BackupBytes;
        var drive = new DriveInfo(Path.GetPathRoot(scan.LocalRoot)!);
        if (drive.AvailableFreeSpace < required) throw new OperationException("DiskFull");
        int done = 0;
        using var stopOnFault = CancellationTokenSource.CreateLinkedTokenSource(ct);
        var workToken = stopOnFault.Token;
        using var parallel = new SemaphoreSlim(2, 2);
        var cacheLocks = new ConcurrentDictionary<string, SemaphoreSlim>();
        async Task Execute(WorkItem item)
        {
            await parallel.WaitAsync(workToken);
            try
            {
                TransferResult result;
                try { result = await TransferOneAsync(workspace, run, item, cacheLocks, progress, workToken); }
                catch (OperationCanceledException) { result = new(item.Entry.RelativePath, "Cancelled", null, null, null, null, 0); }
                catch (Exception ex) when (ex is OperationException or IOException or UnauthorizedAccessException or System.Net.Http.HttpRequestException)
                {
                    string? backup = null;
                    if (item.Replace && Directory.Exists(Path.Combine(run, "backups")))
                    {
                        string candidate = SafePaths.Resolve(Path.Combine(run, "backups"), item.Entry.Local!.RelativePath);
                        if (File.Exists(candidate)) backup = candidate;
                    }
                    result = new(item.Entry.RelativePath, "Error", backup, item.Entry.Local?.Hashes.Sha256, null, null, 0, ex is OperationException op ? op.Code : ex is IOException or UnauthorizedAccessException ? "FileIoError" : "TransferFailed");
                    await JsonFiles.WriteAsync(Path.Combine(run, "errors", JsonFiles.StableId(item.Entry.RelativePath) + ".json"),
                        new { Path = item.Entry.RelativePath, Code = result.Error, ExceptionType = ex.GetType().Name, Message = ex.Message, ex.StackTrace });
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
        foreach (var item in plan.Items)
            if (!results.Any(r => SafePaths.Key(r.Path) == SafePaths.Key(item.Entry.RelativePath))) results.Add(new(item.Entry.RelativePath, "Cancelled", null, null, null, null, 0));
        var report = new ApplyResult(run, results.OrderBy(x => x.Path).ToArray(), ct.IsCancellationRequested);
        // Preserve an audit trail even if the user cancelled the operation.
        await JsonFiles.WriteAsync(Path.Combine(run, "results.json"), report);
        var csv = new List<IEnumerable<object?>> { new object?[] { "Path", "Status", "BackupPath", "OldSha256", "NewSha256", "SharePointHash", "BytesDownloaded", "Error" } };
        csv.AddRange(report.Results.Select(r => new object?[] { r.Path, r.Status, r.BackupPath, r.OldSha256, r.NewSha256, r.SharePointHash, r.BytesDownloaded, r.Error }));
        await JsonFiles.CsvAsync(Path.Combine(run, "results.csv"), csv);
        var installedMetadata = new List<RemoteFile>();
        foreach (var result in report.Results.Where(r => r.Status is "Verified" or "DownloadedUnverified"))
        {
            var installed = await ManifestStore.ReadAsync(workspace, result.Path, CancellationToken.None);
            if (installed != null) installedMetadata.Add(installed.Remote);
        }
        await ReportWriter.WriteMetadataCsvAsync(Path.Combine(run, "remote-metadata.csv"), installedMetadata, CancellationToken.None);
        foreach (var semaphore in cacheLocks.Values) semaphore.Dispose();
        return report;
    }
    private async Task<TransferResult> TransferOneAsync(Workspace workspace, string run, WorkItem work,
        ConcurrentDictionary<string, SemaphoreSlim> locks, IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        var entry = work.Entry;
        if (entry.Remote == null || entry.Status is DiffStatus.LocalOnly or DiffStatus.Error or DiffStatus.Unverified or DiffStatus.Match)
            throw new OperationException("NotEligible", entry.RelativePath);
        var file = await remote.RefreshAsync(entry.Remote, ct);
        if (!SameContent(entry.Remote, file)) throw new OperationException("RemoteChanged", entry.RelativePath);
        if (!file.Metadata.Complete) throw new OperationException("MetadataIncomplete", entry.RelativePath);
        string destination = SafePaths.Resolve(workspace.LocalRoot, entry.RelativePath);
        if (!work.Replace && (File.Exists(destination) || Directory.Exists(destination))) throw new OperationException("FileChanged", entry.RelativePath);
        FileHashes? old = null;
        if (work.Replace)
        {
            old = await FileHashing.ComputeAsync(destination, ct);
            if (old != entry.Local!.Hashes) throw new OperationException("FileChanged", entry.RelativePath);
        }
        string cache = workspace.CachePath(file);
        var cacheLock = locks.GetOrAdd(cache, _ => new SemaphoreSlim(1, 1));
        await cacheLock.WaitAsync(ct);
        long downloaded = 0; FileHashes verified;
        try
        {
            bool valid = false;
            if (File.Exists(cache))
            {
                verified = await FileHashing.ComputeAsync(cache, ct);
                valid = Matches(verified, file) && JsonFiles.Read<CacheStamp>(cache + ".verified.json") is { } stamp &&
                    stamp.Sha256 == verified.Sha256 && (file.QuickXorHash is { Length: > 0 } || stamp.ETag == file.ETag);
            }
            if (!valid)
            {
                string partial = cache + ".partial";
                bool reused = false;
                if (entry.ReusePath is { } source && File.Exists(source) && file.QuickXorHash is { Length: > 0 })
                {
                    var sourceHash = await FileHashing.ComputeAsync(source, ct);
                    if (Matches(sourceHash, file))
                    {
                        await CopyAsync(source, partial, ct); reused = true;
                    }
                }
                if (!reused)
                {
                    var transfer = await remote.DownloadAsync(file, partial, progress, ct); downloaded = transfer.NetworkBytes;
                }
                verified = await FileHashing.ComputeAsync(partial, ct);
                if (!Matches(verified, file))
                {
                    // Keep the bad bytes separate, so the next retry cannot resume corrupt content.
                    File.Move(partial, partial + ".invalid_" + Guid.NewGuid().ToString("N"));
                    throw new OperationException("HashMismatch", entry.RelativePath);
                }
                var after = await remote.RefreshAsync(file, ct);
                if (after.ETag != file.ETag || !SameContent(file, after)) throw new OperationException("RemoteChanged", entry.RelativePath);
                file = after;
                if (!file.Metadata.Complete) throw new OperationException("MetadataIncomplete", entry.RelativePath);
                File.Move(partial, cache, true);
                await JsonFiles.WriteAsync(cache + ".verified.json", new CacheStamp(verified.Sha256, file.ETag), ct);
            }
            verified = await FileHashing.ComputeAsync(cache, ct);
            if (!Matches(verified, file)) throw new OperationException("HashMismatch", entry.RelativePath);
        }
        finally { cacheLock.Release(); }
        // A cached/local reuse must also be guarded by the current remote eTag.
        var current = await remote.RefreshAsync(file, ct);
        if (current.ETag != file.ETag || !SameContent(file, current) || !current.Metadata.Complete)
            throw new OperationException("RemoteChanged", entry.RelativePath);
        file = current;
        ct.ThrowIfCancellationRequested();
        destination = SafePaths.Resolve(workspace.LocalRoot, entry.RelativePath, true);
        string? backup = null;
        if (work.Replace)
        {
            var check = await FileHashing.ComputeAsync(destination, ct);
            if (check != old) throw new OperationException("FileChanged", entry.RelativePath);
            string backupRoot = Path.Combine(run, "backups"); Directory.CreateDirectory(backupRoot);
            backup = SafePaths.Resolve(backupRoot, entry.Local!.RelativePath, true);
            await CopyAsync(destination, backup, ct, FileMode.CreateNew);
            var b = await FileHashing.ComputeAsync(backup, ct);
            if (b != old) throw new OperationException("BackupFailed", entry.RelativePath);
            DateTime created = File.GetCreationTimeUtc(destination), modified = File.GetLastWriteTimeUtc(destination);
            File.SetCreationTimeUtc(backup, created); File.SetLastWriteTimeUtc(backup, modified);
            var oldMetadata = await ManifestStore.ReadAsync(workspace, entry.RelativePath, ct);
            if (oldMetadata?.Sha256 != old!.Sha256) oldMetadata = null;
            await JsonFiles.WriteAsync(backup + ".metadata.json", new
            {
                entry.RelativePath,
                Hashes = old,
                Source = entry.Local.FullPath,
                OriginalCreationUtc = created,
                OriginalModifiedUtc = modified,
                InstalledMetadata = oldMetadata,
                BackedUpAt = DateTimeOffset.Now
            }, ct);
        }
        // Install temp is on the destination volume; backups are already durable and independently verified.
        string install = destination + ".spdiff-" + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            await CopyAsync(cache, install, ct, FileMode.CreateNew);
            var staged = await FileHashing.ComputeAsync(install, ct);
            if (staged != verified) throw new OperationException("HashMismatch", entry.RelativePath);
            ct.ThrowIfCancellationRequested();
            SafePaths.AssertDirectoryChain(workspace.LocalRoot, destination);
            if (work.Replace)
            {
                // Lock out writers while hashing and replacing; a locked user document is reported as an error.
                using var guard = new FileStream(destination, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);
                var check = await FileHashing.ComputeAsync(destination, ct);
                if (check != old) throw new OperationException("FileChanged", entry.RelativePath);
                await JsonFiles.WriteAsync(Path.Combine(run, "intents", JsonFiles.StableId(entry.RelativePath) + ".json"),
                    new { entry.RelativePath, BackupPath = backup, OldSha256 = old!.Sha256, NewSha256 = verified.Sha256, Remote = file }, ct);
                File.Replace(install, destination, null, true);
            }
            else File.Move(install, destination, false);
            // Cancellation never interrupts verification/auditing after a replacement has committed.
            File.SetCreationTimeUtc(destination, file.Created.UtcDateTime);
            File.SetLastWriteTimeUtc(destination, file.Modified.UtcDateTime);
            var installed = await FileHashing.ComputeAsync(destination, CancellationToken.None);
            if (installed != verified) throw new OperationException("HashMismatch", entry.RelativePath);
            await ManifestStore.SaveAsync(workspace, file, installed, CancellationToken.None);
            var result = new TransferResult(entry.RelativePath, file.QuickXorHash is { Length: > 0 } ? "Verified" : "DownloadedUnverified", backup, old?.Sha256, installed.Sha256, file.QuickXorHash, downloaded);
            await JsonFiles.WriteAsync(Path.Combine(run, "committed", JsonFiles.StableId(entry.RelativePath) + ".json"), new { Result = result, Remote = file }, CancellationToken.None);
            return result;
        }
        finally { if (File.Exists(install)) File.Delete(install); }
    }
    private static bool SameContent(RemoteFile a, RemoteFile b) => a.ItemId == b.ItemId && a.DriveId == b.DriveId && a.Size == b.Size && a.QuickXorHash == b.QuickXorHash;
    private static bool Matches(FileHashes h, RemoteFile file) => h.Size == file.Size && (string.IsNullOrEmpty(file.QuickXorHash) || h.QuickXorHash == file.QuickXorHash);
    private static async Task CopyAsync(string source, string target, CancellationToken ct, FileMode mode = FileMode.Create)
    {
        await using var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read, 1024 * 1024, true);
        await using var output = new FileStream(target, mode, FileAccess.Write, FileShare.None, 1024 * 1024, true);
        await input.CopyToAsync(output, 1024 * 1024, ct); await output.FlushAsync(ct); output.Flush(true);
    }
    private sealed record CacheStamp(string Sha256, string ETag);
}
