using System.Net;
using System.Net.Http.Headers;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Time.Testing;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

/// <summary>Streams that misbehave like a throttled line; all waits honour the read's cancellation token.</summary>
internal sealed class ScriptedStream(byte[] prefix, Func<CancellationToken, Task<int>>? afterPrefix = null, TimeSpan? perByte = null, TimeProvider? time = null) : Stream
{
    private int position;
    public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken ct = default)
    {
        if (position < prefix.Length)
        {
            if (perByte is { } delay) { await Task.Delay(delay, time ?? TimeProvider.System, ct); buffer.Span[0] = prefix[position++]; return 1; }
            int n = Math.Min(buffer.Length, prefix.Length - position); prefix.AsSpan(position, n).CopyTo(buffer.Span); position += n; return n;
        }
        return afterPrefix == null ? 0 : await afterPrefix(ct);
    }
    public static Func<CancellationToken, Task<int>> Stall => async ct => { await Task.Delay(Timeout.Infinite, ct); return 0; };
    public static Func<CancellationToken, Task<int>> Throw(Exception ex) => _ => Task.FromException<int>(ex);
    public override int Read(byte[] buffer, int offset, int count) => ReadAsync(buffer.AsMemory(offset, count)).AsTask().GetAwaiter().GetResult();
    public override bool CanRead => true; public override bool CanSeek => false; public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException(); public override long Position { get => position; set => throw new NotSupportedException(); }
    public override void Flush() { }
    public override long Seek(long o, SeekOrigin s) => throw new NotSupportedException();
    public override void SetLength(long v) => throw new NotSupportedException(); public override void Write(byte[] b, int o, int c) => throw new NotSupportedException();
}
internal sealed class RecordingHandler(Func<HttpRequestMessage, int, CancellationToken, Task<HttpResponseMessage>> respond) : HttpMessageHandler
{
    public List<HttpRequestMessage> Requests { get; } = [];
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
    {
        Assert.Equal(HttpMethod.Get, request.Method); int n; lock (Requests) { Requests.Add(request); n = Requests.Count; }
        return respond(request, n, ct);
    }
}

public sealed class DownloadResilienceTests
{
    private sealed class Tokens(Exception? error = null) : ITokenProvider
    { public Task<string> GetAsync(string resource, CancellationToken ct, bool refresh = false) => error == null ? Task.FromResult("TEST-ONLY-TOKEN") : Task.FromException<string>(error); }
    private static readonly byte[] Data = Enumerable.Range(0, 20).Select(i => (byte)(i + 1)).ToArray();
    /// <summary>The Graph /content call always redirects, to a fresh URL each time.</summary>
    private static RecordingHandler Api() => new((req, n, _) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.Found) { Headers = { Location = new Uri($"https://download.example.test/url-{n}") } }));
    private static HttpResponseMessage Body(byte[] data, long from, Stream? stream = null)
    {
        var slice = data[(int)from..];
        var response = new HttpResponseMessage(from > 0 ? HttpStatusCode.PartialContent : HttpStatusCode.OK) { Content = new StreamContent(stream ?? new MemoryStream(slice)) };
        if (from > 0) response.Content.Headers.ContentRange = new ContentRangeHeaderValue(from, data.Length - 1, data.Length);
        return response;
    }
    private static long From(HttpRequestMessage req) => req.Headers.Range?.Ranges.Single().From ?? 0;
    private static SharePointReader Reader(RecordingHandler api, RecordingHandler content, TimeProvider time, int attempts = 8) =>
        new(new Tokens(), new HttpClient(api), new HttpClient(content)) { Time = time, ResumeAttempts = attempts };
    private static async Task<T> DriveAsync<T>(FakeTimeProvider time, Task<T> run, TimeSpan step, int maxSteps = 5000)
    {
        for (int i = 0; i < maxSteps && !run.IsCompleted; i++) { time.Advance(step); await Task.Delay(1); }
        Assert.True(run.IsCompleted, "did not finish within the fake-time bound");
        return await run;
    }

    [Fact]
    public async Task StalledBodyResumesWithRangeAndFreshDownloadUrl()
    {
        using var f = new Fixture(); var file = Fixture.Remote("Docs/a.bin", Data); string partial = Path.Combine(f.Root, "a.partial"); var time = new FakeTimeProvider();
        var api = Api();
        var content = new RecordingHandler((req, n, _) => Task.FromResult(n == 1 ? Body(Data, 0, new ScriptedStream(Data[..4], ScriptedStream.Stall)) : Body(Data, From(req))));
        using var reader = Reader(api, content, time);
        var result = await DriveAsync(time, reader.DownloadAsync(file, partial, null, default), TimeSpan.FromSeconds(1));
        Assert.Equal((20L, 2, 1), (result.NetworkBytes, result.Attempts, result.Stalls));
        Assert.Equal(Data, await File.ReadAllBytesAsync(partial));
        Assert.Equal(4, From(content.Requests[1])); Assert.NotEqual(content.Requests[0].RequestUri, content.Requests[1].RequestUri);
        Assert.Equal(2, api.Requests.Count); Assert.All(content.Requests, r => Assert.Null(r.Headers.Authorization));
    }
    [Fact]
    public async Task StallBeforeFirstByteIsDetected()
    {
        using var f = new Fixture(); var file = Fixture.Remote("Docs/a.bin", Data); var time = new FakeTimeProvider();
        var content = new RecordingHandler((req, n, _) => Task.FromResult(n == 1 ? Body(Data, 0, new ScriptedStream([], ScriptedStream.Stall)) : Body(Data, From(req))));
        using var reader = Reader(Api(), content, time);
        var result = await DriveAsync(time, reader.DownloadAsync(file, Path.Combine(f.Root, "a.partial"), null, default), TimeSpan.FromSeconds(1));
        Assert.Equal((2, 1), (result.Attempts, result.Stalls)); Assert.Equal(0, From(content.Requests[1]));
    }
    [Fact]
    public async Task SlowSteadyBodyIsNeverStalled()
    {
        using var f = new Fixture(); var file = Fixture.Remote("Docs/a.bin", Data[..6]); var time = new FakeTimeProvider();
        var content = new RecordingHandler((_, _, _) => Task.FromResult(Body(Data[..6], 0, new ScriptedStream(Data[..6], perByte: TimeSpan.FromSeconds(50), time: time))));
        using var reader = Reader(Api(), content, time);
        var result = await DriveAsync(time, reader.DownloadAsync(file, Path.Combine(f.Root, "a.partial"), null, default), TimeSpan.FromSeconds(5));
        Assert.Equal((6L, 1, 0), (result.NetworkBytes, result.Attempts, result.Stalls)); Assert.Single(content.Requests);
    }
    [Fact]
    public async Task RetryAfterLongerThanHeaderTimeoutIsNotAStall()
    {
        using var f = new Fixture(); var file = Fixture.Remote("Docs/a.bin", Data); var time = new FakeTimeProvider();
        var content = new RecordingHandler((req, n, _) => Task.FromResult(n == 1
            ? new HttpResponseMessage(HttpStatusCode.ServiceUnavailable) { Headers = { RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromSeconds(150)) } }
            : Body(Data, From(req))));
        using var reader = Reader(Api(), content, time);
        var result = await DriveAsync(time, reader.DownloadAsync(file, Path.Combine(f.Root, "a.partial"), null, default), TimeSpan.FromSeconds(5));
        Assert.Equal((1, 0), (result.Attempts, result.Stalls)); Assert.Equal(2, content.Requests.Count);
    }
    [Fact]
    public async Task PersistentStallFailsWithDownloadStalledThenNextReaderResumes()
    {
        using var f = new Fixture(); var file = Fixture.Remote("Docs/a.bin", Data); string partial = Path.Combine(f.Root, "a.partial"); var time = new FakeTimeProvider();
        // Every attempt delivers one more byte, then stalls.
        var content = new RecordingHandler((req, _, _) => { long from = From(req); return Task.FromResult(Body(Data, from, new ScriptedStream(Data[(int)from..((int)from + 1)], ScriptedStream.Stall))); });
        using (var reader = Reader(Api(), content, time))
        {
            var failure = DriveAsync(time, reader.DownloadAsync(file, partial, null, default), TimeSpan.FromSeconds(5));
            Assert.Equal("DownloadStalled", (await Assert.ThrowsAsync<OperationException>(() => failure)).Code);
        }
        Assert.Equal(8, content.Requests.Count); Assert.Equal(8, new FileInfo(partial).Length); Assert.True(File.Exists(partial + ".json"));
        var healthy = new RecordingHandler((req, _, _) => Task.FromResult(Body(Data, From(req))));
        using var next = Reader(Api(), healthy, time);
        var result = await DriveAsync(time, next.DownloadAsync(file, partial, null, default), TimeSpan.FromSeconds(1));
        Assert.Equal(8, From(healthy.Requests.Single())); Assert.Equal(12, result.NetworkBytes); Assert.Equal(Data, await File.ReadAllBytesAsync(partial));
    }
    [Fact]
    public async Task SocketResetIsRetriedButLocalWriteFailureIsNot()
    {
        using var f = new Fixture(); var file = Fixture.Remote("Docs/a.bin", Data); var time = new FakeTimeProvider();
        var reset = new IOException("connection reset (test)", new SocketException(10054));
        var content = new RecordingHandler((req, n, _) => Task.FromResult(n == 1 ? Body(Data, 0, new ScriptedStream(Data[..3], ScriptedStream.Throw(reset))) : Body(Data, From(req))));
        using (var reader = Reader(Api(), content, time))
        {
            var result = await DriveAsync(time, reader.DownloadAsync(file, Path.Combine(f.Root, "a.partial"), null, default), TimeSpan.FromSeconds(1));
            Assert.Equal((2, 0), (result.Attempts, result.Stalls)); Assert.Equal(3, From(content.Requests[1]));
        }
        // The partial file cannot be written: a local failure is reported as it is and never retried.
        string locked = Path.Combine(f.Root, "locked.partial"); await File.WriteAllBytesAsync(locked, Data[..2]);
        await JsonFiles.WriteAsync(locked + ".json", new { file.ETag, file.Size, file.QuickXorHash }); File.SetAttributes(locked, FileAttributes.ReadOnly);
        var once = new RecordingHandler((req, _, _) => Task.FromResult(Body(Data, From(req))));
        using var second = Reader(Api(), once, time);
        await Assert.ThrowsAsync<UnauthorizedAccessException>(() => DriveAsync(time, second.DownloadAsync(file, locked, null, default), TimeSpan.FromSeconds(1)));
        Assert.Empty(once.Requests); // local destination is checked before spending any download bandwidth
    }
    [Fact]
    public async Task UserCancellationIsNotTreatedAsStall()
    {
        using var f = new Fixture(); var file = Fixture.Remote("Docs/a.bin", Data); var time = new FakeTimeProvider(); using var cancel = new CancellationTokenSource();
        var content = new RecordingHandler((_, _, _) => { cancel.CancelAfter(50); return Task.FromResult(Body(Data, 0, new ScriptedStream(Data[..2], ScriptedStream.Stall))); });
        using var reader = Reader(Api(), content, time);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reader.DownloadAsync(file, Path.Combine(f.Root, "a.partial"), null, cancel.Token).WaitAsync(TimeSpan.FromSeconds(10)));
        Assert.Single(content.Requests);
    }
    [Fact]
    public async Task PartialResumesAfterMetadataOnlyETagChange()
    {
        using var f = new Fixture(); var file = Fixture.Remote("Docs/a.bin", Data, etag: "etag-new"); string partial = Path.Combine(f.Root, "a.partial");
        await File.WriteAllBytesAsync(partial, Data[..4]); await JsonFiles.WriteAsync(partial + ".json", new { ETag = "etag-old", file.Size, file.QuickXorHash });
        var content = new RecordingHandler((req, _, _) => Task.FromResult(Body(Data, From(req))));
        using var reader = Reader(Api(), content, TimeProvider.System);
        var result = await reader.DownloadAsync(file, partial, null, default);
        Assert.Equal(4, From(content.Requests.Single())); Assert.Equal(16, result.NetworkBytes); Assert.Equal(Data, await File.ReadAllBytesAsync(partial));
    }
    [Fact]
    public async Task DownloadHost404AndSignInRequiredAreNotRetried()
    {
        using var f = new Fixture(); var file = Fixture.Remote("Docs/a.bin", Data);
        var missing = new RecordingHandler((_, _, _) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound)));
        using (var reader = Reader(Api(), missing, TimeProvider.System))
            Assert.Equal("RemoteMissing", (await Assert.ThrowsAsync<OperationException>(() => reader.DownloadAsync(file, Path.Combine(f.Root, "a.partial"), null, default))).Code);
        Assert.Single(missing.Requests);
        var api = Api(); var unused = new RecordingHandler((_, _, _) => throw new InvalidOperationException("no content request expected"));
        using var signedOut = new SharePointReader(new Tokens(new OperationException("SignInRequired", "graph")), new HttpClient(api), new HttpClient(unused));
        Assert.Equal("SignInRequired", (await Assert.ThrowsAsync<OperationException>(() => signedOut.DownloadAsync(file, Path.Combine(f.Root, "b.partial"), null, default))).Code);
        Assert.Empty(api.Requests);
    }
}

/// <summary>A small in-memory SharePoint for the real reader: one library "Docs" (list id "list", drive "drive") with scripted content.</summary>
internal sealed class FakeSharePoint
{
    public const string Site = Fixture.Site;
    public readonly Dictionary<string, (byte[] Bytes, string ETag, string Revision)> Files = [];
    /// <summary>Optional per-request override for the download host: (item id, request number for that item, range start) → response.</summary>
    public Func<string, int, long, HttpResponseMessage?>? OnContent { get; set; }
    public RecordingHandler Api { get; }
    public RecordingHandler Content { get; }
    private readonly Dictionary<string, int> contentCalls = [];
    public FakeSharePoint() { Api = new((req, _, _) => Task.FromResult(Route(req))); Content = new((req, _, _) => Task.FromResult(Download(req))); }
    public void Put(string name, byte[] bytes, string etag = "e1", string revision = "B") => Files[name] = (bytes, etag, revision);
    private static string Id(string name) => "id-" + name;
    private static HttpResponseMessage Json(object value) => new(HttpStatusCode.OK) { Content = new StringContent(JsonSerializer.Serialize(value), Encoding.UTF8, "application/json") };
    private object Item(string name)
    {
        var (bytes, etag, _) = Files[name]; var q = new QuickXorHasher(); q.Append(bytes);
        return new
        {
            id = Id(name),
            name,
            size = bytes.Length,
            eTag = etag,
            file = new { hashes = new { quickXorHash = q.Finish() } },
            parentReference = new { id = "root", path = "/drives/drive/root:" },
            sharepointIds = new { listItemId = (Files.Keys.ToList().IndexOf(name) + 1).ToString() },
            createdDateTime = "2025-01-02T03:04:05Z",
            lastModifiedDateTime = "2026-01-02T03:04:05Z",
            webUrl = "https://contoso.sharepoint.com/" + name
        };
    }
    private object Metadata(string name) => new { Id = Files.Keys.ToList().IndexOf(name) + 1, Revision = Files[name].Revision, FieldValuesAsText = new { Revision = Files[name].Revision } };
    private HttpResponseMessage Route(HttpRequestMessage req)
    {
        string path = Uri.UnescapeDataString(req.RequestUri!.AbsolutePath);
        if (path == "/v1.0/sites/contoso.sharepoint.com:/sites/review") return Json(new { id = "site-id" });
        if (path.EndsWith("/_api/web/lists")) return Json(new { value = new[] { new { Id = "list", Title = "Docs", Hidden = false, BaseTemplate = 101, ItemCount = Files.Count, RootFolder = new { ServerRelativeUrl = "/sites/review/Docs" } } } });
        if (path == "/v1.0/sites/site-id/lists/list/drive") return Json(new { id = "drive" });
        if (path.EndsWith("/fields")) return Json(new { value = new[] { new { InternalName = "Revision", Title = "Revision", TypeAsString = "Text", Hidden = false, ReadOnlyField = false, FromBaseType = false } } });
        if (path.EndsWith("/items")) return Json(new { value = Files.Keys.Select(Metadata).ToArray() });
        if (path.Contains("/items(")) { int id = int.Parse(path[(path.LastIndexOf('(') + 1)..path.LastIndexOf(')')]); return Json(Metadata(Files.Keys.ElementAt(id - 1))); }
        if (path == "/v1.0/drives/drive/root/delta")
            return Json(new Dictionary<string, object> { ["value"] = new object[] { new { id = "root", name = "root", root = new { }, folder = new { } } }.Concat(Files.Keys.Select(Item)).ToArray(), ["@odata.deltaLink"] = "https://graph.microsoft.com/v1.0/done" });
        if (path.StartsWith("/v1.0/drives/drive/items/") && path.EndsWith("/content"))
            return new(HttpStatusCode.Found) { Headers = { Location = new Uri("https://download.example.test/" + path.Split('/')[^2] + "?n=" + Guid.NewGuid().ToString("N")) } };
        if (path.StartsWith("/v1.0/drives/drive/items/")) return Json(Item(path["/v1.0/drives/drive/items/".Length..][3..]));
        throw new InvalidOperationException("unexpected request " + path);
    }
    private HttpResponseMessage Download(HttpRequestMessage req)
    {
        string id = req.RequestUri!.AbsolutePath.Trim('/'); string name = id[3..]; var bytes = Files[name].Bytes; long from = req.Headers.Range?.Ranges.Single().From ?? 0;
        int n = contentCalls[id] = contentCalls.GetValueOrDefault(id) + 1;
        if (OnContent?.Invoke(id, n, from) is { } custom) return custom;
        if (from >= bytes.Length) return new(HttpStatusCode.RequestedRangeNotSatisfiable);
        var response = new HttpResponseMessage(from > 0 ? HttpStatusCode.PartialContent : HttpStatusCode.OK) { Content = new ByteArrayContent(bytes[(int)from..]) };
        if (from > 0) response.Content.Headers.ContentRange = new ContentRangeHeaderValue(from, bytes.Length - 1, bytes.Length);
        return response;
    }
}

public sealed class OfflineEndToEndTests
{
    private sealed class Tokens : ITokenProvider { public Task<string> GetAsync(string resource, CancellationToken ct, bool refresh = false) => Task.FromResult("TEST-ONLY-TOKEN"); }
    private static async Task<SharePointReader> PrepareAsync(Fixture f, FakeSharePoint sp, TimeProvider? time = null)
    {
        await LibrarySelectionStore.WriteAsync(new Workspace(Fixture.Site, f.Root), new(["list"], ["list"], DateTimeOffset.Now), default);
        return new SharePointReader(new Tokens(), new HttpClient(sp.Api), new HttpClient(sp.Content)) { Time = time ?? TimeProvider.System };
    }
    [Fact]
    public async Task CheckThenSyncAllInstallsAndIssuesOnlyGets()
    {
        using var f = new Fixture(); var sp = new FakeSharePoint();
        sp.Put("a.bin", [1, 2, 3]); sp.Put("b.bin", [4, 5, 6, 7]); sp.Put("c.bin", [8, 9]);
        await f.LocalAsync("Docs/a.bin", [1, 2, 3]); var edited = await f.LocalAsync("Docs/b.bin", [0, 0]);
        using var reader = await PrepareAsync(f, sp);
        var outcome = await new DiffService(reader).CheckAsync(Fixture.Site, f.Root, null, default);
        Assert.Equal((1, 1, 1), (outcome.Summary.NewFiles, outcome.Summary.Updates, outcome.Summary.Unknown));
        var result = await new SyncEngine(reader).ApplyAsync(outcome.Scan, DiffService.SyncAll(outcome.Scan.Entries), null, default);
        Assert.All(result.Results, r => Assert.Equal("Verified", r.Status));
        Assert.Equal(new byte[] { 4, 5, 6, 7 }, await File.ReadAllBytesAsync(edited.FullPath));
        Assert.Equal(new byte[] { 8, 9 }, await File.ReadAllBytesAsync(Path.Combine(f.Root, "Docs", "c.bin")));
        Assert.Equal(new byte[] { 0, 0 }, await File.ReadAllBytesAsync(result.Results.Single(r => r.Path == "Docs/b.bin").BackupPath!));
        Assert.All(sp.Api.Requests.Concat(sp.Content.Requests), r => Assert.Equal(HttpMethod.Get, r.Method));
        Assert.All(sp.Content.Requests, r => Assert.Null(r.Headers.Authorization));
    }
    [Fact]
    public async Task StalledFirstAttemptResumesWithinRun()
    {
        using var f = new Fixture(); var sp = new FakeSharePoint(); var time = new FakeTimeProvider();
        byte[] big = Enumerable.Range(0, 64).Select(i => (byte)i).ToArray(); sp.Put("big.bin", big);
        sp.OnContent = (_, n, from) => n == 1 ? new HttpResponseMessage(HttpStatusCode.OK) { Content = new StreamContent(new ScriptedStream(big[..10], ScriptedStream.Stall)) } : null;
        using var reader = await PrepareAsync(f, sp, time);
        var outcome = await new DiffService(reader).CheckAsync(Fixture.Site, f.Root, null, default);
        var run = new SyncEngine(reader).ApplyAsync(outcome.Scan, DiffService.SyncAll(outcome.Scan.Entries), null, default);
        for (int i = 0; i < 2000 && !run.IsCompleted; i++) { time.Advance(TimeSpan.FromSeconds(1)); await Task.Delay(1); }
        var row = Assert.Single((await run).Results);
        Assert.Equal(("Verified", 2, 1, 64L), (row.Status, row.Attempts, row.Stalls, row.BytesDownloaded));
        Assert.Equal(10, sp.Content.Requests[1].Headers.Range!.Ranges.Single().From);
        Assert.Equal(big, await File.ReadAllBytesAsync(Path.Combine(f.Root, "Docs", "big.bin")));
        Assert.Contains("Attempts", await File.ReadAllTextAsync(Path.Combine((await run).RunDirectory, "results.csv")));
    }
    [Fact]
    public async Task CorruptResumedPrefixIsQuarantinedAndNextRunRestartsFromZero()
    {
        using var f = new Fixture(); var sp = new FakeSharePoint(); byte[] bytes = [10, 20, 30, 40, 50, 60]; sp.Put("x.bin", bytes);
        using var reader = await PrepareAsync(f, sp);
        var outcome = await new DiffService(reader).CheckAsync(Fixture.Site, f.Root, null, default);
        var remote = outcome.Scan.Entries.Single().Remote!;
        string partial = new Workspace(Fixture.Site, f.Root).CachePath(remote) + ".partial";
        await File.WriteAllBytesAsync(partial, [99, 99, 99]); await JsonFiles.WriteAsync(partial + ".json", new { remote.ETag, remote.Size, remote.QuickXorHash });
        var first = await new SyncEngine(reader).ApplyAsync(outcome.Scan, DiffService.SyncAll(outcome.Scan.Entries), null, default);
        Assert.Equal("HashMismatch", Assert.Single(first.Results).Error);
        Assert.Single(Directory.GetFiles(Path.GetDirectoryName(partial)!, Path.GetFileName(partial) + ".invalid_*"));
        int before = sp.Content.Requests.Count;
        var second = await new SyncEngine(reader).ApplyAsync(outcome.Scan, DiffService.SyncAll(outcome.Scan.Entries), null, default);
        Assert.Equal("Verified", Assert.Single(second.Results).Status);
        Assert.Null(sp.Content.Requests[before].Headers.Range);
        Assert.Equal(bytes, await File.ReadAllBytesAsync(Path.Combine(f.Root, "Docs", "x.bin")));
    }
    [Fact]
    public async Task MetadataOnlyChangeAfterDownloadGivesRemoteChangedThenZeroBytes()
    {
        using var f = new Fixture(); var sp = new FakeSharePoint(); byte[] bytes = [1, 1, 2, 3, 5, 8]; sp.Put("m.bin", bytes);
        using var reader = await PrepareAsync(f, sp);
        var outcome = await new DiffService(reader).CheckAsync(Fixture.Site, f.Root, null, default);
        // A column is edited in the browser while the bytes are downloading: the ETag changes, the content does not.
        sp.OnContent = (_, n, _) => { if (n == 1) sp.Files["m.bin"] = (bytes, "e2", "C"); return null; };
        var first = await new SyncEngine(reader).ApplyAsync(outcome.Scan, DiffService.SyncAll(outcome.Scan.Entries), null, default);
        Assert.Equal("RemoteChanged", Assert.Single(first.Results).Error);
        sp.OnContent = null;
        var again = await new DiffService(reader).CheckAsync(Fixture.Site, f.Root, null, default);
        var second = await new SyncEngine(reader).ApplyAsync(again.Scan, DiffService.SyncAll(again.Scan.Entries), null, default);
        var row = Assert.Single(second.Results);
        Assert.Equal(("Verified", 0L), (row.Status, row.BytesDownloaded));
        Assert.Equal(bytes, await File.ReadAllBytesAsync(Path.Combine(f.Root, "Docs", "m.bin")));
    }
}
