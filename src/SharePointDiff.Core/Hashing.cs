using System.Buffers;
using System.Security.Cryptography;

namespace SharePointDiff.Core;

/// <summary>Microsoft QuickXorHash: 160-bit XOR, rotate 11 bits per byte; length XOR at byte 12.
/// Reference: https://learn.microsoft.com/onedrive/developer/code-snippets/quickxorhash
/// Accumulates 160 byte lanes, then folds them to the 160-bit output.</summary>
public sealed class QuickXorHasher
{
    private readonly byte[] lanes = new byte[160];
    private long length;
    public void Append(ReadOnlySpan<byte> data)
    {
        int offset = (int)(length % 160), i = 0;
        while (offset != 0 && i < data.Length) { lanes[offset] ^= data[i++]; offset = (offset + 1) % 160; }
        while (i + 160 <= data.Length)
        {
            for (int j = 0; j < 160; j++) lanes[j] ^= data[i + j];
            i += 160;
        }
        for (int j = 0; i < data.Length; i++, j++) lanes[j] ^= data[i];
        length += data.Length;
    }
    public string Finish()
    {
        var output = new byte[20];
        for (int i = 0; i < 160; i++)
        {
            int shift = i * 11 % 160, index = shift / 8, bit = shift % 8;
            output[index] ^= (byte)(lanes[i] << bit);
            if (bit != 0) output[(index + 1) % 20] ^= (byte)(lanes[i] >> (8 - bit));
        }
        for (int i = 0; i < 8; i++) output[12 + i] ^= (byte)((ulong)length >> (8 * i));
        return Convert.ToBase64String(output);
    }
}

public static class FileHashing
{
    public static async Task<FileHashes> CopyHashAsync(Stream input, string target, CancellationToken ct, FileMode mode = FileMode.CreateNew)
    {
        using var hashes = new HashAccumulator();
        byte[] buffer = ArrayPool<byte>.Shared.Rent(1024 * 1024);
        try
        {
            await using var output = new FileStream(target, mode, FileAccess.Write, FileShare.None, 1024 * 1024, true);
            int n;
            while ((n = await input.ReadAsync(buffer.AsMemory(), ct)) != 0)
            {
                hashes.Append(buffer.AsSpan(0, n)); IoMetrics.Add("Content.ReadBytes", n);
                await output.WriteAsync(buffer.AsMemory(0, n), ct); IoMetrics.Add("Content.WriteBytes", n);
            }
            await output.FlushAsync(ct); output.Flush(true);
            return hashes.Finish();
        }
        finally { ArrayPool<byte>.Shared.Return(buffer); }
    }
    public static async Task<FileHashes> CopyHashAsync(string source, string target, CancellationToken ct, FileMode mode = FileMode.CreateNew)
    {
        await using var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read, 1024 * 1024, FileOptions.Asynchronous | FileOptions.SequentialScan);
        return await CopyHashAsync(input, target, ct, mode);
    }
    public static async Task<FileHashes> CopyVerifiedAsync(string source, string target, CancellationToken ct, FileMode mode = FileMode.CreateNew)
    {
        var copied = await CopyHashAsync(source, target, ct, mode);
        if (copied != await ComputeAsync(target, ct)) throw new OperationException("HashMismatch", target);
        return copied;
    }
    /// <param name="share">Scans pass ReadWrite|Delete so Office can still save; the size/mtime post-check then reports FileChanged.
    /// Every verification before a write keeps the default Read share.</param>
    public static async Task<FileHashes> ComputeAsync(string path, CancellationToken ct = default, FileShare share = FileShare.Read) =>
        (await ReadDetailedAsync(path, ct, share, false)).Hashes;

    internal static async Task<HashReadResult> ReadDetailedAsync(string path, CancellationToken ct, FileShare share, bool blocks, string? root = null)
    {
        IoMetrics.Add("Hash.Calls");
        var beforePath = new FileInfo(path); long expected = beforePath.Length; DateTime stamp = beforePath.LastWriteTimeUtc;
        await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, share, 1024 * 1024,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        if (root != null)
        {
            if (NativeFileInfo.FinalPath(stream) is { } final) SafePaths.Inside(root, final);
            else SafePaths.Resolve(root, Path.GetRelativePath(root, path));
        }
        var before = NativeFileInfo.TryRead(stream);
        using var hashes = new HashAccumulator();
        using var chunkHash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var chunks = new List<HashBlock>(); long chunkOffset = 0; int chunkSize = 0;
        byte[] buffer = ArrayPool<byte>.Shared.Rent(1024 * 1024);
        DateTimeOffset now = DateTimeOffset.UtcNow;
        try
        {
            int n;
            while ((n = await stream.ReadAsync(buffer.AsMemory(), ct)) != 0)
            {
                ct.ThrowIfCancellationRequested(); hashes.Append(buffer.AsSpan(0, n)); IoMetrics.Add("Content.ReadBytes", n);
                if (blocks)
                {
                    int offset = 0;
                    while (offset < n)
                    {
                        int take = Math.Min(n - offset, LocalHashCache.BlockSize - chunkSize);
                        chunkHash.AppendData(buffer.AsSpan(offset, take)); offset += take; chunkSize += take;
                        if (chunkSize == LocalHashCache.BlockSize)
                        {
                            chunks.Add(new(chunkOffset, chunkSize, Convert.ToHexString(chunkHash.GetHashAndReset()), now));
                            chunkOffset += chunkSize; chunkSize = 0;
                        }
                    }
                }
            }
            if (blocks && chunkSize > 0) chunks.Add(new(chunkOffset, chunkSize, Convert.ToHexString(chunkHash.GetHashAndReset()), now));
        }
        finally { ArrayPool<byte>.Shared.Return(buffer); }
        var after = NativeFileInfo.TryRead(stream); beforePath.Refresh();
        if (hashes.Size != expected || beforePath.Length != expected || beforePath.LastWriteTimeUtc != stamp || before != after)
            throw new OperationException("FileChanged", path);
        // With a scan's permissive share, a pathname can be swapped while the original handle remains valid.
        if (after != null)
        {
            await using var current = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            if (after != NativeFileInfo.TryRead(current)) throw new OperationException("FileChanged", path);
        }
        return new(hashes.Finish(), after, after != null ? DateTime.FromFileTimeUtc(after.LastWriteTime) : stamp, chunks, now);
    }

}

internal sealed class HashAccumulator : IDisposable
{
    private readonly QuickXorHasher quick = new();
    private readonly IncrementalHash sha = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
    public long Size { get; private set; }
    public void Append(ReadOnlySpan<byte> bytes) { quick.Append(bytes); sha.AppendData(bytes); Size += bytes.Length; }
    public FileHashes Finish() => new(Size, quick.Finish(), Convert.ToHexString(sha.GetCurrentHash()));
    public void Dispose() => sha.Dispose();
}
