using System.Collections.Concurrent;
using System.Net;

namespace SharePointDiff.Core;

internal static class HttpTransport
{
    internal static SocketsHttpHandler Handler(bool json) => new()
    {
        AllowAutoRedirect = false,
        AutomaticDecompression = json ? DecompressionMethods.GZip | DecompressionMethods.Deflate | DecompressionMethods.Brotli : DecompressionMethods.None,
        PooledConnectionLifetime = TimeSpan.FromMinutes(10),
        KeepAlivePingPolicy = HttpKeepAlivePingPolicy.WithActiveRequests,
        KeepAlivePingDelay = TimeSpan.FromSeconds(30),
        KeepAlivePingTimeout = TimeSpan.FromSeconds(15),
        ConnectTimeout = TimeSpan.FromSeconds(30),
        UseCookies = false
    };
    internal static HttpClient Client(bool json) => new(Handler(json))
    {
        DefaultRequestVersion = HttpVersion.Version20,
        DefaultVersionPolicy = HttpVersionPolicy.RequestVersionOrLower,
        Timeout = Timeout.InfiniteTimeSpan
    };
}

/// <summary>Workers share the service's not-before time. Never shorten Retry-After.</summary>
internal sealed class ServiceCooldown
{
    private readonly ConcurrentDictionary<string, DateTimeOffset> until = new(StringComparer.OrdinalIgnoreCase);
    public void Set(Uri uri, DateTimeOffset value) => until.AddOrUpdate(uri.Authority, value, (_, old) => old > value ? old : value);
    public async Task WaitAsync(Uri uri, TimeProvider time, TimeSpan budget, CancellationToken ct)
    {
        while (until.TryGetValue(uri.Authority, out var next))
        {
            var delay = next - time.GetUtcNow();
            if (delay <= TimeSpan.Zero) return;
            if (delay > budget) throw new OperationException("RetryLater", uri.GetLeftPart(UriPartial.Path),
                new HttpRequestException("Retry-After: " + next.ToString("O")));
            IoMetrics.Add("Network.RetryWaitMilliseconds", (long)Math.Ceiling(delay.TotalMilliseconds));
            await Task.Delay(delay, time, ct);
        }
    }
}

internal sealed class CountingReadStream(Stream inner, string counter) : Stream
{
    public override bool CanRead => inner.CanRead;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
    public override int Read(byte[] buffer, int offset, int count) { int n = inner.Read(buffer, offset, count); IoMetrics.Add(counter, n); return n; }
    public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken ct = default)
    { int n = await inner.ReadAsync(buffer, ct); IoMetrics.Add(counter, n); return n; }
    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken ct) => ReadAsync(buffer.AsMemory(offset, count), ct).AsTask();
    public override void Flush() => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    protected override void Dispose(bool disposing) { if (disposing) inner.Dispose(); base.Dispose(disposing); }
    public override ValueTask DisposeAsync() => inner.DisposeAsync();
}
