using System.Net;
using System.Text;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class OptimizationTests
{
    [Fact]
    public async Task FlatFolderIsEnumeratedOnceInsteadOfOncePerRemoteFile()
    {
        using var f = new Fixture(); var remote = new FakeRemote();
        Directory.CreateDirectory(Path.Combine(f.Root, "Docs"));
        for (int i = 0; i < 3000; i++)
        {
            string path = "Docs/f" + i + ".bin";
            await File.WriteAllBytesAsync(Path.Combine(f.Root, path), []);
            remote.Add(Fixture.Remote(path, []), []);
        }
        await new DiffService(remote).ScanAsync(Fixture.Site, f.Root, [Fixture.Library], null, default, report: false);
        var io = JsonFiles.Read<IoStatistics>(Path.Combine(new Workspace(Fixture.Site, f.Root).DataRoot, "last-scan-io.json"))!;
        Assert.Equal(2, io.Counters["Directory.Enumerations"]);
        Assert.Equal(3001, io.Counters["Directory.Entries"]);
        Assert.Equal(3000, io.Counters["Hash.Calls"]);
        Assert.Equal(0, remote.Downloads);
    }

    [Fact]
    public void CompressionIsOnlyEnabledForJsonAndRawClientsNeverRedirectTokens()
    {
        using var json = HttpTransport.Handler(true); using var raw = HttpTransport.Handler(false);
        Assert.Equal(DecompressionMethods.GZip | DecompressionMethods.Deflate | DecompressionMethods.Brotli, json.AutomaticDecompression);
        Assert.Equal(DecompressionMethods.None, raw.AutomaticDecompression);
        Assert.False(json.AllowAutoRedirect); Assert.False(raw.AllowAutoRedirect);
        Assert.Equal(TimeSpan.FromMinutes(10), json.PooledConnectionLifetime);
    }

    private sealed class Tokens : ITokenProvider
    {
        public Task<string> GetAsync(string resource, CancellationToken ct, bool refresh = false) => Task.FromResult("test");
    }
    private sealed class Handler(Func<HttpRequestMessage, HttpResponseMessage> send) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct) => Task.FromResult(send(request));
    }
    [Fact]
    public async Task DirectContentUsesRawAuthenticatedClientAndNeverJsonClient()
    {
        using var f = new Fixture(); var file = Fixture.Remote("Docs/raw", [1, 2, 3]);
        using var reader = new SharePointReader(new Tokens(),
            new HttpClient(new Handler(_ => throw new InvalidOperationException("JSON client must not receive content"))),
            new HttpClient(new Handler(_ => throw new InvalidOperationException("no redirect"))),
            new HttpClient(new Handler(req =>
            {
                Assert.Equal("graph.microsoft.com", req.RequestUri!.Host); Assert.Equal("test", req.Headers.Authorization!.Parameter);
                Assert.Contains(req.Headers.AcceptEncoding, x => x.Value == "identity");
                return new(HttpStatusCode.OK) { Content = new ByteArrayContent([1, 2, 3]) };
            })));
        await reader.DownloadAsync(file, Path.Combine(f.Root, "part"), null, default);
        Assert.Equal(new byte[] { 1, 2, 3 }, await File.ReadAllBytesAsync(Path.Combine(f.Root, "part")));
    }

    [Fact]
    public async Task LongRetryAfterIsReportedWithoutRetryingEarly()
    {
        int calls = 0;
        using var reader = new SharePointReader(new Tokens(), new HttpClient(new Handler(_ =>
        {
            calls++;
            return new(HttpStatusCode.TooManyRequests)
            {
                Content = new StringContent("busy", Encoding.UTF8),
                Headers = { RetryAfter = new System.Net.Http.Headers.RetryConditionHeaderValue(DateTimeOffset.UtcNow.AddMinutes(10)) }
            };
        })), new HttpClient()) { RetryWaitBudget = TimeSpan.FromSeconds(1) };
        var error = await Assert.ThrowsAsync<OperationException>(() => reader.GetLibrariesAsync(Fixture.Site, new Dictionary<string, string>(), default));
        Assert.Equal("RetryLater", error.Code); Assert.Equal(1, calls); Assert.Contains("Retry-After", error.InnerException!.Message);
    }
}
