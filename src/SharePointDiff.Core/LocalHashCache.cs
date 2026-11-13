using System.Collections.Concurrent;
using System.Text.Json;

namespace SharePointDiff.Core;

public sealed record HashBlock(long Offset, int Length, string Sha256, DateTimeOffset CheckedAt);
internal sealed record HashReadResult(FileHashes Hashes, FileIdentity? Identity, DateTime LastWriteUtc, IReadOnlyList<HashBlock> Blocks, DateTimeOffset ReadAt);
internal sealed record HashCacheEntry(string RelativePath, FileIdentity Identity, FileHashes Hashes, DateTimeOffset HashedAt, IReadOnlyList<HashBlock> Blocks);
internal sealed record HashCacheDocument(int FormatVersion, string Root, Dictionary<string, HashCacheEntry> Entries, DateOnly ScrubDay, long ScrubBytes);

/// <summary>Speculative scan acceleration only. Never used to authorize replacement or pruning.</summary>
public sealed class LocalHashCache
{
    public const int BlockSize = 8 * 1024 * 1024;
    private readonly Workspace workspace;
    private readonly ConcurrentDictionary<string, HashCacheEntry> entries;
    private readonly bool supported;
    private int dirty;
    private DateOnly scrubDay;
    private long scrubBytes;
    public List<OperationIssue> Issues { get; } = [];
    private string PathName => Path.Combine(workspace.DataRoot, "local-hash-cache.json");
    private LocalHashCache(Workspace workspace, HashCacheDocument? document)
    {
        this.workspace = workspace; supported = NativeFileInfo.SupportsCache(workspace.LocalRoot);
        entries = new(document?.Entries ?? new Dictionary<string, HashCacheEntry>(), StringComparer.Ordinal);
        scrubDay = document?.ScrubDay ?? default; scrubBytes = document?.ScrubBytes ?? 0;
    }
    public static LocalHashCache Load(Workspace workspace)
    {
        string path = Path.Combine(workspace.DataRoot, "local-hash-cache.json");
        HashCacheDocument? saved = null; string? problem = null;
        try
        {
            saved = JsonFiles.Read<HashCacheDocument>(path);
            if (saved != null && (saved.FormatVersion != 1 || saved.Root != workspace.LocalRoot || saved.Entries == null ||
                saved.Entries.Any(p => p.Value == null || p.Key != SafePaths.Key(p.Value.RelativePath) || !Valid(p.Value))))
            { saved = null; problem = "Invalid hash-cache format or entries"; }
        }
        catch (JsonException ex) { problem = ex.Message; }
        var cache = new LocalHashCache(workspace, saved);
        if (problem != null) { cache.Issues.Add(new("HashCacheInvalid", path, problem)); cache.dirty = 1; }
        return cache;
    }
    private static bool Valid(HashCacheEntry entry)
    {
        if (entry.Identity == null || entry.Hashes == null || entry.Blocks == null || entry.Identity.Size != entry.Hashes.Size ||
            entry.Hashes.Size < 0 || entry.Identity.FileId is not { Length: 32 } || !Hex(entry.Hashes.Sha256)) return false;
        try { if (Convert.FromBase64String(entry.Hashes.QuickXorHash).Length != 20) return false; } catch (FormatException) { return false; }
        long offset = 0;
        foreach (var block in entry.Blocks)
        {
            if (block == null || block.Offset != offset || block.Length <= 0 || block.Length > BlockSize || !Hex(block.Sha256)) return false;
            offset += block.Length;
        }
        return offset == entry.Hashes.Size;
    }
    private static bool Hex(string? value) => value is { Length: 64 } && value.All(char.IsAsciiHexDigit);
    public async Task<LocalFile> ReadAsync(string path, bool force, CancellationToken ct)
    {
        string relative = Path.GetRelativePath(workspace.LocalRoot, path).Replace('\\', '/'), key = SafePaths.Key(relative);
        if (supported && !force && entries.TryGetValue(key, out var saved))
        {
            try
            {
                // Briefly exclude writers/deleters while checking the handle; an open Office writer falls back to a real read.
                await using var handle = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                string? final = NativeFileInfo.FinalPath(handle);
                if (final != null) SafePaths.Inside(workspace.LocalRoot, final);
                if (final != null && NativeFileInfo.TryRead(handle) == saved.Identity)
                {
                    IoMetrics.Add("HashCache.Hits");
                    return new(relative, path, saved.Hashes, DateTime.FromFileTimeUtc(saved.Identity.LastWriteTime), HashBasis.Saved, saved.HashedAt, saved.Identity);
                }
            }
            catch (IOException ex) when ((ex.HResult & 0xFFFF) is 32 or 33) { }
        }
        IoMetrics.Add("HashCache.Misses");
        var fresh = await FileHashing.ReadDetailedAsync(path, ct, FileShare.ReadWrite | FileShare.Delete, supported, workspace.LocalRoot);
        if (supported && fresh.Identity != null)
        {
            entries[key] = new(relative, fresh.Identity, fresh.Hashes, fresh.ReadAt, fresh.Blocks);
            Interlocked.Exchange(ref dirty, 1);
        }
        return new(relative, path, fresh.Hashes, fresh.LastWriteUtc, HashBasis.Fresh, fresh.ReadAt, fresh.Identity);
    }
    public async Task SaveAsync(CancellationToken ct)
    {
        if (Volatile.Read(ref dirty) == 0) return;
        await JsonFiles.WriteAsync(PathName, new HashCacheDocument(1, workspace.LocalRoot, new(entries), scrubDay, scrubBytes), ct);
        Interlocked.Exchange(ref dirty, 0);
    }
    internal IReadOnlyList<(string Key, HashCacheEntry Entry, int Index)> ScrubCandidates(DateTimeOffset now) => entries
        .SelectMany(p => p.Value.Blocks.Select((_, i) => (Key: p.Key, Entry: p.Value, Index: i)))
        .Where(p => p.Entry.Blocks[p.Index].CheckedAt.UtcDateTime.Date < now.UtcDateTime.Date)
        .OrderBy(p => p.Entry.Blocks[p.Index].CheckedAt).ThenBy(p => p.Key, StringComparer.Ordinal).ThenBy(p => p.Index).ToArray();
    internal long RemainingBudget(DateTimeOffset now)
    {
        var today = DateOnly.FromDateTime(now.UtcDateTime);
        if (scrubDay != today) { scrubDay = today; scrubBytes = 0; }
        long total = entries.Values.Sum(e => e.Hashes.Size);
        long budget = total == 0 ? 0 : Math.Min(1024L * 1024 * 1024, Math.Max(BlockSize, (long)Math.Ceiling(total * .02)));
        return Math.Max(0, budget - scrubBytes);
    }
    internal void Scrubbed(string key, HashCacheEntry expected, int index, DateTimeOffset now, int bytes)
    {
        if (!entries.TryGetValue(key, out var current) || current.Identity != expected.Identity || current.Hashes != expected.Hashes) return;
        var blocks = current.Blocks.ToArray(); blocks[index] = blocks[index] with { CheckedAt = now };
        if (entries.TryUpdate(key, current with { Blocks = blocks }, current)) dirty = 1;
    }
    internal void ChargeScrub(long bytes) { if (bytes > 0) { scrubBytes += bytes; dirty = 1; } }
    internal void Invalidated(string key) { if (entries.TryRemove(key, out _)) dirty = 1; }
}
