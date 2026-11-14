using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using SharePointDiff.Core;

if (args.Length < 3) throw new ArgumentException("Use --crash <fixture mirror> <stage>, --live/--live-inventory <fixture mirror> <settings> <site> <tiny-file-json>, or --normalize <fixture mirror> <inventory-json>.");
string root = Path.GetFullPath(args[1]);
var parent = new DirectoryInfo(root).Parent;
if (Path.GetFileName(root) != "mirror" || parent == null || !Guid.TryParseExact(parent.Name, "N", out _))
    throw new ArgumentException("A generated GUID/mirror test folder is required.");
Directory.CreateDirectory(root);
if (args[0] == "--crash")
{
    const string site = "https://contoso.sharepoint.com/sites/crash";
    string path = SafePaths.Resolve(root, "Docs/a", true); await File.WriteAllBytesAsync(path, [1, 2, 3]);
    var local = new LocalFile("Docs/a", path, await FileHashing.ComputeAsync(path), File.GetLastWriteTimeUtc(path));
    var remote = new ProbeRemote(site); var entry = new DiffEntry { RelativePath = "Docs/a", Remote = remote.File, Local = local, Status = DiffStatus.DifferentHash };
    var engine = new SyncEngine(remote)
    {
        Checkpoint = (stage, _) => { if (stage == args[2]) { Process.GetCurrentProcess().Kill(); Thread.Sleep(Timeout.Infinite); } }
    };
    await engine.ApplyAsync(new(site, root, [entry], ["Docs"], [], DateTimeOffset.Now), new([new(entry, true)], 0, 3, 3), null, default);
    throw new InvalidOperationException("The requested crash boundary was not reached.");
}
if (args[0] == "--normalize")
{
    var inventory = JsonFiles.Read<RemoteInventory>(args[2]) ?? throw new InvalidOperationException("Missing inventory");
    var workspace = new Workspace(inventory.Files.First().SiteUrl, root);
    string output = Path.Combine(workspace.DataRoot, "normalized-inventory.json");
    await JsonFiles.WriteAsync(output, inventory);
    var roundTrip = JsonFiles.Read<RemoteInventory>(output)!;
    if (roundTrip.Files.Count != inventory.Files.Count || roundTrip.Files.Zip(inventory.Files).Any(pair =>
        !pair.First.Metadata.Complete || pair.First.RelativePath != pair.Second.RelativePath || pair.First.Size != pair.Second.Size || pair.First.QuickXorHash != pair.Second.QuickXorHash ||
        !JsonElement.DeepEquals(pair.First.Metadata.Fields, pair.Second.Metadata.Fields) || !JsonElement.DeepEquals(pair.First.Metadata.DisplayValues, pair.Second.Metadata.DisplayValues) ||
        !JsonElement.DeepEquals(pair.First.Metadata.Columns, pair.Second.Metadata.Columns) || !JsonElement.DeepEquals(pair.First.Basic, pair.Second.Basic)))
        throw new InvalidOperationException("Typed roundtrip failed");
    await JsonFiles.WriteAsync(Path.Combine(parent.FullName, "normalization.json"), new
    {
        Files = inventory.Files.Count,
        BeforeBytes = new FileInfo(args[2]).Length,
        AfterBytes = new FileInfo(output).Length,
        SchemaBytes = Directory.GetFiles(Path.Combine(workspace.DataRoot, "schemas")).Sum(p => new FileInfo(p).Length),
        Schemas = Directory.GetFiles(Path.Combine(workspace.DataRoot, "schemas")).Length,
        TypedRoundTrip = true
    });
    return;
}
if (args[0] is not ("--live" or "--live-inventory") || args.Length < 5) throw new ArgumentException("Unknown probe mode");
var configuration = JsonFiles.Read<UserSettings>(args[2])?.Normalize() ?? throw new InvalidOperationException("Missing test configuration");
string liveSite = DiffService.NormalizeSite(args[3]);
var profile = configuration.Profiles[configuration.SelectedProfile];
var authentication = new BrowserAuthentication(profile, true); // deliberately no persisted auth cache
Console.WriteLine("SIGN_IN_REQUIRED");
await authentication.SignInAsync(["https://graph.microsoft.com", new Uri(liveSite).GetLeftPart(UriPartial.Authority)], default);
Console.WriteLine("SIGNED_IN");
using var audit = new InventoryAudit(HttpTransport.Handler(true));
HttpClient ContentClient() => args[0] == "--live-inventory" ? new(new ForbiddenInventoryContent()) : HttpTransport.Client(false);
using var reader = new SharePointReader(authentication.Interactive, new HttpClient(audit), ContentClient(), ContentClient());
var service = new DiffService(reader);
var libraries = await service.ConnectAsync(liveSite, root, default);
if (args[0] == "--live-inventory")
{
    var examined = new List<object>();
    foreach (var candidate in libraries.Where(l => l.ItemCount > 0).OrderBy(l => l.ItemCount))
    {
        audit.Clear();
        Console.WriteLine("READING_FILE_LIBRARY_" + candidate.ItemCount);
        var cold = await service.ScanAsync(liveSite, root, [candidate], null, default, report: false, options: new(ReadAllRemote: true));
        int files = cold.Entries.Count(e => e.Remote != null);
        examined.Add(new { candidate.ItemCount, Files = files, cold.Issues });
        if (cold.Issues.Count != 0) throw new InvalidOperationException("Live inventory incomplete: " + JsonSerializer.Serialize(cold.Issues));
        if (files == 0) continue;
        var coldRequests = audit.Snapshot();
        audit.Clear();
        Console.WriteLine("READING_FILE_DELTA");
        var warm = await service.ScanAsync(liveSite, root, [candidate], null, default, report: false);
        var warmRequests = audit.Snapshot();
        audit.Clear();
        Console.WriteLine("READING_FILE_FULL_REFERENCE");
        var full = await service.ScanAsync(liveSite, root, [candidate], null, default, report: false, options: new(ReadAllRemote: true));
        var fullRequests = audit.Snapshot();
        static void EqualInventory(ScanResult a, ScanResult b)
        {
            var left = a.Entries.Where(e => e.Remote != null).OrderBy(e => e.RelativePath, StringComparer.Ordinal).Select(e => e.Remote!).ToArray();
            var right = b.Entries.Where(e => e.Remote != null).OrderBy(e => e.RelativePath, StringComparer.Ordinal).Select(e => e.Remote!).ToArray();
            if (a.Issues.Count != 0 || b.Issues.Count != 0 || left.Length != right.Length ||
                !a.RemoteFolders.Order().SequenceEqual(b.RemoteFolders.Order()) || left.Zip(right).Any(p =>
                p.First.RelativePath != p.Second.RelativePath || p.First.Size != p.Second.Size || p.First.QuickXorHash != p.Second.QuickXorHash ||
                p.First.ETag != p.Second.ETag || !JsonElement.DeepEquals(p.First.Basic, p.Second.Basic) ||
                !JsonElement.DeepEquals(p.First.Metadata.Fields, p.Second.Metadata.Fields) ||
                !JsonElement.DeepEquals(p.First.Metadata.DisplayValues, p.Second.Metadata.DisplayValues) ||
                !JsonElement.DeepEquals(p.First.Metadata.Columns, p.Second.Metadata.Columns)))
                throw new InvalidOperationException("Live delta and full inventory differ; retain evidence and inspect a concurrent source change.");
        }
        EqualInventory(cold, warm); EqualInventory(warm, full);
        if (warm.RemoteReads?.Single().Incremental != true || warmRequests.GetValueOrDefault("REST.Items") != 0 ||
            warm.Statistics?.Counters.GetValueOrDefault("Content.ReadBytes") != 0)
            throw new InvalidOperationException("Warm live scan did not use the expected shortcuts.");
        await JsonFiles.WriteAsync(Path.Combine(parent.FullName, "live-inventory-results.json"), new
        {
            LibrariesListed = libraries.Count,
            Examined = examined,
            Files = files,
            candidate.ItemCount,
            ColdIo = cold.Statistics,
            WarmIo = warm.Statistics,
            FullIo = full.Statistics,
            ColdRequests = coldRequests,
            WarmRequests = warmRequests,
            FullRequests = fullRequests,
            WarmRead = warm.RemoteReads,
            TypedInventoryEquivalent = true,
            PayloadDownloaded = 0,
            AuthCacheFile = authentication.CacheFilePath,
            Note = "Read-only initial/delta/full inventory of the smallest library containing files; no content request or SharePoint write."
        });
        Console.WriteLine("LIVE_INVENTORY_PASSED");
        return;
    }
    throw new InvalidOperationException("No document library containing files was available.");
}
var selected = libraries.Where(l => l.ItemCount > 0).OrderBy(l => l.ItemCount).Take(1).ToArray();
if (selected.Length == 0) selected = libraries.Take(1).ToArray();
Console.WriteLine("READING_SMALLEST_LIBRARY");
var first = await service.ScanAsync(liveSite, root, selected, null, default, report: false);
Console.WriteLine("READING_DELTA");
var second = await service.ScanAsync(liveSite, root, selected, null, default, report: false);
var tinyEntry = JsonFiles.Read<DiffEntry>(args[4]);
var known = tinyEntry?.Remote ?? throw new InvalidOperationException("Missing tiny-file fixture");
if (known.SiteUrl != liveSite || known.Size != 40 || !libraries.Any(l => l.Id == known.LibraryId)) throw new InvalidOperationException("Tiny-file scope mismatch");
var tiny = await reader.RefreshAsync(known, default);
if (tiny.Size != known.Size || !tiny.Metadata.Complete) throw new InvalidOperationException("Tiny-file size/metadata changed");
var missing = new DiffEntry { RelativePath = tiny.RelativePath, Status = DiffStatus.MissingLocal, Remote = tiny };
ScanResult Thin(DiffEntry entry) => new(liveSite, root, [entry], [], [], DateTimeOffset.Now);
SyncPlan Plan(DiffEntry entry, bool replace) => new([new(entry, replace)], 0, tiny.Size, replace ? tiny.Size : 0);
Console.WriteLine("DOWNLOADING_TINY_FILE");
var installed = await new SyncEngine(reader).ApplyAsync(Thin(missing), Plan(missing, false), null, default);
var downloaded = installed.Results.Single();
if (downloaded.Status != "Verified" || downloaded.InstalledLocal == null || downloaded.InstalledRemote == null) throw new InvalidOperationException("Tiny download failed: " + downloaded.Error);
byte[] dirty = await File.ReadAllBytesAsync(downloaded.InstalledLocal.FullPath); dirty[0] ^= 0x55;
await File.WriteAllBytesAsync(downloaded.InstalledLocal.FullPath, dirty);
string extra = Path.Combine(root, "local-only-test.txt"); await File.WriteAllTextAsync(extra, "local only; preserve");
var edited = new LocalFile(tiny.RelativePath, downloaded.InstalledLocal.FullPath, await FileHashing.ComputeAsync(downloaded.InstalledLocal.FullPath), File.GetLastWriteTimeUtc(downloaded.InstalledLocal.FullPath));
var different = new DiffEntry { RelativePath = tiny.RelativePath, Status = DiffStatus.DifferentHash, Remote = downloaded.InstalledRemote, Local = edited };
Console.WriteLine("REPLACING_FROM_VERIFIED_CACHE");
var replacement = await new SyncEngine(reader).ApplyAsync(Thin(different), Plan(different, true), null, default);
var replaced = replacement.Results.Single();
if (replaced.Status != "Verified" || replaced.BackupPath == null || !File.ReadAllBytes(replaced.BackupPath).SequenceEqual(dirty) || replaced.BytesDownloaded != 0)
    throw new InvalidOperationException("Replacement/backup/cache verification failed: " + replaced.Error);
string reuse = Path.Combine(root, "reuse-source.bin"); File.Move(edited.FullPath, reuse);
// Hold only this test's generated cache aside so that the third transfer must use the local candidate.
var workspaceLive = new Workspace(liveSite, root); string cache = workspaceLive.CachePath(tiny);
File.Move(cache, Path.Combine(parent.FullName, "held-cache.bin"));
var reuseEntry = new DiffEntry { RelativePath = tiny.RelativePath, Status = DiffStatus.MissingLocal, Remote = replaced.InstalledRemote!, ReusePath = reuse, ReuseSha256 = replaced.NewSha256 };
Console.WriteLine("REUSING_LOCAL_BYTES");
var reused = await new SyncEngine(reader).ApplyAsync(Thin(reuseEntry), Plan(reuseEntry, false), null, default);
if (reused.Results.Single().Status != "Verified" || reused.Results.Single().BytesDownloaded != 0 || !File.Exists(reuse) || File.ReadAllText(extra) != "local only; preserve")
    throw new InvalidOperationException("Local reuse/preservation failed");
await JsonFiles.WriteAsync(Path.Combine(parent.FullName, "live-results.json"), new
{
    LibrariesListed = libraries.Count,
    SelectedLibraryCount = selected.Length,
    SelectedItemCount = selected.Sum(l => l.ItemCount),
    FirstFiles = first.Entries.Count(e => e.Remote != null),
    FirstIssues = first.Issues,
    FirstIo = first.Statistics,
    SecondIssues = second.Issues,
    SecondIo = second.Statistics,
    SecondRead = second.RemoteReads,
    TinyBytes = tiny.Size,
    PayloadDownloaded = installed.Results.Sum(r => r.BytesDownloaded) + replacement.Results.Sum(r => r.BytesDownloaded) + reused.Results.Sum(r => r.BytesDownloaded),
    BackupVerified = true,
    LocalReuseVerified = true,
    LocalOnlyPreserved = true,
    Permissions = authentication.PermissionReports,
    AuthCacheFile = authentication.CacheFilePath,
    authentication.PersistenceBlocked,
    Note = "Read-only SharePoint operations; inventory on the smallest nonempty library, engine transfer fixture on an existing tiny file; no Windows logoff or physical power-loss test."
});
Console.WriteLine("LIVE_TEST_PASSED");

internal sealed class ProbeRemote : IRemoteReader
{
    public RemoteFile File { get; }
    public ProbeRemote(string site)
    {
        var q = new QuickXorHasher(); q.Append([4, 5, 6]);
        File = new(site, "list", "drive", "file", "1", "Docs/a", "a", 3, q.Finish(), "etag", DateTimeOffset.UtcNow, DateTimeOffset.UtcNow, site + "/a", JsonFiles.Empty,
            new(JsonSerializer.SerializeToElement(new { Revision = "B" }), JsonFiles.Empty, JsonFiles.EmptyArray, true));
    }
    public Task<IReadOnlyList<LibraryInfo>> GetLibrariesAsync(string site, IReadOnlyDictionary<string, string> maps, CancellationToken ct) => Task.FromResult<IReadOnlyList<LibraryInfo>>([]);
    public Task<RemoteInventory> GetInventoryAsync(string site, IReadOnlyList<LibraryInfo> libs, IProgress<OperationProgress>? progress, CancellationToken ct) => Task.FromResult(new RemoteInventory([File], ["Docs"], []));
    public Task<RemoteFile> RefreshAsync(RemoteFile file, CancellationToken ct) => Task.FromResult(File);
    public async Task<DownloadResult> DownloadAsync(RemoteFile file, string path, IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        byte[] bytes = [4, 5, 6]; await System.IO.File.WriteAllBytesAsync(path, bytes, ct);
        return new(3, Hashes: new(3, File.QuickXorHash!, Convert.ToHexString(SHA256.HashData(bytes))));
    }
}

internal sealed class InventoryAudit(HttpMessageHandler inner) : DelegatingHandler(inner)
{
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, int> counts = new();
    public void Clear() => counts.Clear();
    public Dictionary<string, int> Snapshot() => new(counts);
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
    {
        if (request.Method != HttpMethod.Get) throw new InvalidOperationException("The live probe permits GET only.");
        string path = request.RequestUri!.AbsolutePath;
        if (path.EndsWith("/content", StringComparison.Ordinal)) throw new InvalidOperationException("Inventory must not download content.");
        string kind = path.Contains("/_api/", StringComparison.Ordinal) ? path.EndsWith("/items", StringComparison.Ordinal) ? "REST.Items" :
            path.EndsWith("/fields", StringComparison.Ordinal) ? "REST.Fields" : "REST.Count" :
            path.Contains("/lists/", StringComparison.Ordinal) ? "Graph.ListDelta" : path.Contains("/drives/", StringComparison.Ordinal) ? "Graph.DriveDelta" : "Graph.Other";
        counts.AddOrUpdate(kind, 1, (_, n) => n + 1);
        return base.SendAsync(request, ct);
    }
}

internal sealed class ForbiddenInventoryContent : HttpMessageHandler
{
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct) =>
        throw new InvalidOperationException("The inventory acceptance probe must not request file content.");
}
