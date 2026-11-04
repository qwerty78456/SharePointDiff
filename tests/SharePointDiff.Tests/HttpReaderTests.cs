using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class HttpReaderTests
{
    private sealed class Tokens : ITokenProvider
    {
        public HashSet<string> Resources { get; } = [];
        public Task<string> GetAsync(string resource, CancellationToken ct, bool refresh = false) { Resources.Add(resource); return Task.FromResult("TEST-ONLY-TOKEN"); }
    }
    private sealed class Handler(Func<HttpRequestMessage, HttpResponseMessage> action) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        { Assert.Equal(HttpMethod.Get, request.Method); return Task.FromResult(action(request)); }
    }
    private static HttpResponseMessage Json(object value) => new(HttpStatusCode.OK) { Content = new StringContent(JsonSerializer.Serialize(value), Encoding.UTF8, "application/json") };
    [Fact]
    public async Task ResumesPartialContentWithoutSendingBearerToDownloadHost()
    {
        using var fixture = new Fixture(); byte[] data = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]; var file = Fixture.Remote("Docs/video.bin", data);
        string partial = Path.Combine(fixture.Root, "part.bin"); await File.WriteAllBytesAsync(partial, data[..4]);
        await JsonFiles.WriteAsync(partial + ".json", new { file.ETag, file.Size });
        var api = new HttpClient(new Handler(req =>
        {
            Assert.Equal("TEST-ONLY-TOKEN", req.Headers.Authorization?.Parameter);
            return new(HttpStatusCode.Found) { Headers = { Location = new Uri("https://download.example.test/opaque") } };
        }));
        int calls = 0;
        var content = new HttpClient(new Handler(req =>
        {
            calls++; Assert.Null(req.Headers.Authorization); Assert.Equal(4, Assert.Single(req.Headers.Range!.Ranges).From);
            var response = new HttpResponseMessage(HttpStatusCode.PartialContent) { Content = new ByteArrayContent(data[4..]) };
            response.Content.Headers.ContentRange = new ContentRangeHeaderValue(4, 9, 10); return response;
        }));
        using var reader = new SharePointReader(new Tokens(), api, content);
        var result = await reader.DownloadAsync(file, partial, null, default);
        Assert.Equal(6, result.NetworkBytes); Assert.Equal(1, calls); Assert.Equal(data, await File.ReadAllBytesAsync(partial));
    }
    [Fact]
    public async Task ServerIgnoringRangeRestartsCleanlyInsteadOfAppending()
    {
        using var fixture = new Fixture(); byte[] data = [1, 2, 3, 4]; var file = Fixture.Remote("Docs/a", data);
        string partial = Path.Combine(fixture.Root, "part"); await File.WriteAllBytesAsync(partial, [1, 2]); await JsonFiles.WriteAsync(partial + ".json", new { file.ETag, file.Size });
        using var reader = new SharePointReader(new Tokens(),
            new HttpClient(new Handler(_ => new(HttpStatusCode.Found) { Headers = { Location = new Uri("https://download.example.test/a") } })),
            new HttpClient(new Handler(_ => new(HttpStatusCode.OK) { Content = new ByteArrayContent(data) })));
        var result = await reader.DownloadAsync(file, partial, null, default);
        Assert.Equal(data, await File.ReadAllBytesAsync(partial)); Assert.Equal(4, result.NetworkBytes);
    }
    [Fact]
    public async Task ExpiredPartialVersionIsNeverAppended()
    {
        using var fixture = new Fixture(); byte[] data = [1, 2, 3]; var file = Fixture.Remote("Docs/a", data); string partial = Path.Combine(fixture.Root, "part");
        await File.WriteAllBytesAsync(partial, [9, 9]); await JsonFiles.WriteAsync(partial + ".json", new { ETag = "old", file.Size });
        using var reader = new SharePointReader(new Tokens(), new HttpClient(new Handler(_ => new(HttpStatusCode.Found) { Headers = { Location = new Uri("https://download.example.test/a") } })),
            new HttpClient(new Handler(req => { Assert.Null(req.Headers.Range); return new(HttpStatusCode.OK) { Content = new ByteArrayContent(data) }; })));
        await reader.DownloadAsync(file, partial, null, default); Assert.Equal(data, await File.ReadAllBytesAsync(partial));
    }
    [Fact]
    public async Task ContentThrottlingRetriesWithoutLosingPartial()
    {
        using var fixture = new Fixture(); var file = Fixture.Remote("Docs/a", [1, 2, 3]); int calls = 0;
        using var reader = new SharePointReader(new Tokens(), new HttpClient(new Handler(_ => new(HttpStatusCode.Found) { Headers = { Location = new Uri("https://download.example.test/a") } })),
            new HttpClient(new Handler(_ => ++calls == 1 ? new((HttpStatusCode)429) { Headers = { RetryAfter = new RetryConditionHeaderValue(TimeSpan.Zero) } } : new(HttpStatusCode.OK) { Content = new ByteArrayContent([1, 2, 3]) })));
        await reader.DownloadAsync(file, Path.Combine(fixture.Root, "part"), null, default); Assert.Equal(2, calls);
    }
    [Fact]
    public async Task PagedInventoryPreservesMetadataAndAppliesLastDeltaRecord()
    {
        var tokens = new Tokens(); var file = Fixture.Remote("Docs/a.bin", [1, 2, 3]);
        object Root() => new { id = "root", name = "Docs", root = new { }, folder = new { } };
        object Item(string id, string name) => new
        {
            id,
            name,
            size = 3,
            file = new { hashes = new { quickXorHash = file.QuickXorHash } },
            parentReference = new { id = "root" },
            sharepointIds = new { listItemId = "1" },
            createdDateTime = "2025-01-01T00:00:00Z",
            lastModifiedDateTime = "2026-01-01T00:00:00Z",
            eTag = "e",
            webUrl = "https://contoso.sharepoint.com/a"
        };
        using var reader = new SharePointReader(tokens, new HttpClient(new Handler(req =>
        {
            string path = req.RequestUri!.AbsolutePath;
            if (path.EndsWith("/fields")) return Json(new { value = new[] { new { InternalName = "Revision", Title = "Revision", TypeAsString = "Text", Hidden = false, ReadOnlyField = false, FromBaseType = false } } });
            if (path.EndsWith("/items")) return Json(new { value = new[] { new { Id = 1, Revision = "B", ReviewersId = new[] { 12, 13 }, Taxonomy = new { Label = "Civil", TermGuid = "guid" }, FieldValuesAsText = new { Revision = "B", Reviewers = "Alice; Bob" } } } });
            if (path.EndsWith("/root/delta")) return Json(new Dictionary<string, object>{{"value",new[]{Root(),Item("removed","removed.bin"),Item("a","old.bin")}},
                {"@odata.nextLink","https://graph.microsoft.com/v1.0/page2"}});
            if (path.EndsWith("/page2")) return Json(new Dictionary<string, object>{{"value",new object[]{new{id="removed",deleted=new{}},Item("a","a.bin")}},
                {"@odata.deltaLink","https://graph.microsoft.com/v1.0/delta-finished"}});
            throw new InvalidOperationException(path);
        })), new HttpClient(new Handler(_ => throw new InvalidOperationException("No content downloads during inventory"))));
        var result = await reader.GetInventoryAsync(Fixture.Site, [Fixture.Library], null, default);
        Assert.Empty(result.Issues); var actual = Assert.Single(result.Files); Assert.Equal("Docs/a.bin", actual.RelativePath);
        Assert.True(actual.Metadata.Complete); Assert.Equal("B", actual.Metadata.Fields.GetProperty("Revision").GetString());
        Assert.Equal(2, actual.Metadata.Fields.GetProperty("ReviewersId").GetArrayLength());
        Assert.Equal("Civil", actual.Metadata.Fields.GetProperty("Taxonomy").GetProperty("Label").GetString());
        Assert.Contains("https://graph.microsoft.com", tokens.Resources); Assert.Contains("https://contoso.sharepoint.com", tokens.Resources);
    }
    [Fact]
    public async Task ApiErrorIncludesEndpointStatusServerMessageAndRequestId()
    {
        using var reader = new SharePointReader(new Tokens(), new HttpClient(new Handler(_ => new(HttpStatusCode.Forbidden) { Content = new StringContent("{\"error\":{\"code\":\"accessDenied\",\"message\":\"Read permission required\"}}"), Headers = { { "request-id", "test-request-id" } } })), new HttpClient());
        var ex = await Assert.ThrowsAsync<OperationException>(() => reader.GetLibrariesAsync(Fixture.Site, new Dictionary<string, string>(), default));
        Assert.Equal("AccessDenied", ex.Code); Assert.Contains("graph.microsoft.com", ex.ItemPath);
        Assert.Contains("403", ex.InnerException!.Message); Assert.Contains("Read permission required", ex.InnerException.Message); Assert.Contains("test-request-id", ex.InnerException.Message);
    }
}
