using System.Buffers;
using System.Security.Cryptography;

namespace SharePointDiff.Core;

public sealed record ScrubResult(long ReadBytes, int VerifiedBlocks, IReadOnlyList<OperationIssue> Issues, bool Paused);

public static class LocalScrubber
{
    public static async Task<ScrubResult> RunAsync(Workspace workspace, Func<bool> shouldPause, CancellationToken ct, TimeProvider? time = null)
    {
        using var measurements = IoMetrics.Begin("IdleScrub");
        using var gate = WorkspaceLock.Acquire(workspace);
        var cache = LocalHashCache.Load(workspace); var issues = cache.Issues.ToList();
        long bytes = 0; int verified = 0; bool paused = false;
        var clock = time ?? TimeProvider.System;
        byte[] buffer = ArrayPool<byte>.Shared.Rent(1024 * 1024);
        try
        {
            foreach (var candidate in cache.ScrubCandidates(clock.GetUtcNow()))
            {
                var block = candidate.Entry.Blocks[candidate.Index];
                if (cache.RemainingBudget(clock.GetUtcNow()) < block.Length) break;
                if (ct.IsCancellationRequested || shouldPause()) { paused = true; break; }
                long read = 0;
                try
                {
                    string path = SafePaths.Inside(workspace.LocalRoot, Path.Combine(workspace.LocalRoot, candidate.Entry.RelativePath.Replace('/', Path.DirectorySeparatorChar)));
                    await using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete,
                        1024 * 1024, FileOptions.Asynchronous | FileOptions.SequentialScan);
                    string? final = NativeFileInfo.FinalPath(input);
                    if (final == null) { cache.Invalidated(candidate.Key); continue; }
                    SafePaths.Inside(workspace.LocalRoot, final);
                    if (NativeFileInfo.TryRead(input) != candidate.Entry.Identity) { cache.Invalidated(candidate.Key); continue; }
                    input.Position = block.Offset;
                    using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
                    while (read < block.Length)
                    {
                        if (ct.IsCancellationRequested || shouldPause()) { paused = true; break; }
                        int n = await input.ReadAsync(buffer.AsMemory(0, (int)Math.Min(buffer.Length, block.Length - read)), ct);
                        if (n == 0) throw new OperationException("FileChanged", candidate.Entry.RelativePath);
                        hash.AppendData(buffer.AsSpan(0, n)); read += n; IoMetrics.Add("Content.ReadBytes", n);
                    }
                    if (paused) break;
                    await using var current = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                    if (NativeFileInfo.TryRead(input) != candidate.Entry.Identity || NativeFileInfo.TryRead(current) != candidate.Entry.Identity)
                    { cache.Invalidated(candidate.Key); continue; }
                    if (Convert.ToHexString(hash.GetHashAndReset()) != block.Sha256)
                    {
                        cache.Invalidated(candidate.Key);
                        issues.Add(new("HashCacheMismatch", candidate.Entry.RelativePath, "Block offset: " + block.Offset));
                        continue;
                    }
                    cache.Scrubbed(candidate.Key, candidate.Entry, candidate.Index, clock.GetUtcNow(), block.Length); verified++;
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested) { paused = true; break; }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or OperationException)
                {
                    cache.Invalidated(candidate.Key);
                    // An active writer is ordinary contention; its next Compare computes the hash again.
                    if (ex is not IOException io || (io.HResult & 0xFFFF) is not (32 or 33))
                        issues.Add(new(ex is OperationException op ? op.Code : "LocalReadFailed", candidate.Entry.RelativePath, ex.ToString()));
                }
                finally { bytes += read; cache.ChargeScrub(read); }
            }
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
            await cache.SaveAsync(CancellationToken.None);
        }
        var result = new ScrubResult(bytes, verified, issues, paused);
        if (bytes > 0 || issues.Count > 0) await JsonFiles.WriteAsync(Path.Combine(workspace.DataRoot, "last-idle-check.json"), new { CheckedAt = clock.GetUtcNow(), Result = result, Io = measurements.Snapshot() });
        return result;
    }
}
