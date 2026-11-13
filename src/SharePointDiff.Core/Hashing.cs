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
    public static async Task<FileHashes> ComputeAsync(string path, CancellationToken ct = default, FileShare share = FileShare.Read)
    {
        IoMetrics.Add("Hash.Calls");
        var before = new FileInfo(path);
        long expected = before.Length;
        DateTime stamp = before.LastWriteTimeUtc;
        await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, share, 1024 * 1024,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        var qxh = new QuickXorHasher();
        using var sha = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        byte[] buffer = ArrayPool<byte>.Shared.Rent(1024 * 1024);
        long size = 0;
        try
        {
            int count;
            while ((count = await stream.ReadAsync(buffer.AsMemory(), ct)) != 0)
            {
                ct.ThrowIfCancellationRequested();
                qxh.Append(buffer.AsSpan(0, count));
                sha.AppendData(buffer.AsSpan(0, count));
                size += count;
                IoMetrics.Add("Content.ReadBytes", count);
            }
        }
        finally { ArrayPool<byte>.Shared.Return(buffer); }
        before.Refresh();
        if (size != expected || before.Length != expected || before.LastWriteTimeUtc != stamp)
            throw new OperationException("FileChanged", path);
        return new(size, qxh.Finish(), Convert.ToHexString(sha.GetHashAndReset()));
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
