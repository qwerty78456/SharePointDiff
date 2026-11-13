using System.Net;
using System.Text;
using System.Text.Json;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class RemoteDeltaTests
{
    private sealed class Tokens : ITokenProvider
    {
        public string Account { get; set; } = "account-a";
        public string CacheIdentity => Account;
        public Task<string> GetAsync(string resource, CancellationToken ct, bool refresh = false) => Task.FromResult("test-only");
    }
    private sealed class Server : HttpMessageHandler
    {
        public readonly Dictionary<string, JsonElement> Items = [];
        public readonly Dictionary<string, string> Revisions = [];
        private readonly Dictionary<string, (long Version, JsonElement Item)> driveChanges = [];
        private readonly Dictionary<string, (long Version, JsonElement Item)> listChanges = [];
        public readonly List<string> Requests = [];
        public long Version;
        public bool ExtraColumn, MissingColumn, MissingMetadata, ExpireOnce, AlwaysExpire, BadParentOnce, EmptyPageOnce, MaliciousLink, ListUnavailable;
        public Server()
        {
            Items["root"] = JsonSerializer.SerializeToElement(new { id = "root", root = new { }, folder = new { } });
            Items["file"] = File("file", "one.bin", "root", "1"); Revisions["1"] = "A";
        }
        public int MetadataCalls => Requests.Count(r => r.Contains("/_api/", StringComparison.Ordinal) && r.Contains("/items?", StringComparison.Ordinal));
        public int FullDriveCalls => Requests.Count(r => r.Contains("/drives/drive/root/delta?", StringComparison.Ordinal) && !r.Contains("cursor=", StringComparison.Ordinal) && !r.Contains("page=", StringComparison.Ordinal));
        private static JsonElement File(string id, string name, string parent, string listId) => JsonSerializer.SerializeToElement(new
        {
            id,
            name,
            size = 3,
            file = new { hashes = new { quickXorHash = Fixture.Remote("x", [1, 2, 3]).QuickXorHash } },
            parentReference = new { id = parent },
            sharepointIds = new { listItemId = listId },
            eTag = "content-etag",
            createdDateTime = "2025-01-01T00:00:00Z",
            lastModifiedDateTime = "2026-01-01T00:00:00Z",
            webUrl = Fixture.Site + "/" + name
        });
        public void MetadataOnly(string revision) { Version++; Revisions["1"] = revision; listChanges["1"] = (Version, JsonSerializer.SerializeToElement(new { id = "1", eTag = "list-" + Version })); }
        public void PutFileInFolder()
        {
            Version++;
            Items["folder"] = JsonSerializer.SerializeToElement(new { id = "folder", name = "Folder", folder = new { }, parentReference = new { id = "root" } });
            Items["file"] = File("file", "one.bin", "folder", "1");
            Revisions["2"] = "folder";
            driveChanges["folder"] = (Version, Items["folder"]); driveChanges["file"] = (Version, Items["file"]);
            listChanges["2"] = (Version, JsonSerializer.SerializeToElement(new { id = "2", eTag = "folder-etag" }));
        }
        public void RenameFolder()
        {
            Version++;
            Items["folder"] = JsonSerializer.SerializeToElement(new { id = "folder", name = "Renamed", folder = new { }, parentReference = new { id = "root" } });
            driveChanges["folder"] = (Version, Items["folder"]);
        }
        public void DeleteFile()
        {
            Version++; Items.Remove("file"); Revisions.Remove("1");
            driveChanges["file"] = (Version, JsonSerializer.SerializeToElement(new { id = "file", deleted = new { } }));
            listChanges["1"] = (Version, JsonSerializer.SerializeToElement(new { id = "1", deleted = new { } }));
        }
        private object Row(string id)
        {
            var row = new Dictionary<string, object?> { ["Id"] = int.Parse(id), ["Revision"] = Revisions[id], ["FieldValuesAsText"] = new { Revision = Revisions[id] } };
            if (ExtraColumn && !MissingColumn) row["Code"] = 42;
            return row;
        }
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            Assert.Equal(HttpMethod.Get, request.Method); Assert.Equal("test-only", request.Headers.Authorization!.Parameter);
            var uri = request.RequestUri!; string path = uri.AbsolutePath; string query = Uri.UnescapeDataString(uri.Query);
            Requests.Add(uri.ToString());
            HttpResponseMessage Json(object value) => new(HttpStatusCode.OK) { Content = new StringContent(JsonSerializer.Serialize(value), Encoding.UTF8, "application/json") };
            if (path.Contains("/sites/contoso.sharepoint.com:", StringComparison.Ordinal)) return Task.FromResult(Json(new { id = "site" }));
            if (path.EndsWith("/fields", StringComparison.Ordinal))
            {
                object Column(string name) => new { Id = name, InternalName = name, EntityPropertyName = name, Title = name, TypeAsString = "Text", Hidden = false, ReadOnlyField = false, FromBaseType = false };
                return Task.FromResult(Json(new { value = ExtraColumn ? new[] { Column("Revision"), Column("Code") } : new[] { Column("Revision") } }));
            }
            if (path.Contains("/delta", StringComparison.Ordinal))
            {
                bool drive = path.Contains("/drives/", StringComparison.Ordinal);
                if (!drive && ListUnavailable) return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound) { Content = new StringContent("unsupported") });
                long since = query.Contains("cursor=", StringComparison.Ordinal) ? long.Parse(query[(query.IndexOf("cursor=", StringComparison.Ordinal) + 7)..].Split('&')[0]) : -1;
                if (drive && since >= 0 && (ExpireOnce || AlwaysExpire)) { ExpireOnce = false; return Task.FromResult(new HttpResponseMessage(HttpStatusCode.Gone) { Content = new StringContent("expired") }); }
                var endpoint = "https://graph.microsoft.com" + path;
                if (drive && EmptyPageOnce && !query.Contains("page=", StringComparison.Ordinal))
                { EmptyPageOnce = false; return Task.FromResult(Json(new Dictionary<string, object> { ["value"] = Array.Empty<object>(), ["@odata.nextLink"] = endpoint + "?page=1" })); }
                JsonElement[] values = since < 0
                    ? drive ? Items.Values.ToArray() : Revisions.Keys.Select(id => JsonSerializer.SerializeToElement(new { id, eTag = "list-0" })).ToArray()
                    : (drive ? driveChanges : listChanges).Values.Where(x => x.Version > since).Select(x => x.Item).ToArray();
                if (drive && BadParentOnce && since >= 0)
                {
                    BadParentOnce = false;
                    values = [File("file", "one.bin", "missing-parent", "1")];
                }
                return Task.FromResult(Json(new Dictionary<string, object> { ["value"] = values, ["@odata.deltaLink"] = MaliciousLink ? "https://evil.example/delta" : endpoint + "?cursor=" + Version }));
            }
            if (path.EndsWith("/items", StringComparison.Ordinal))
            {
                IEnumerable<string> ids = Revisions.Keys;
                if (query.Contains("$filter=", StringComparison.Ordinal))
                    ids = query[(query.IndexOf("$filter=", StringComparison.Ordinal) + 8)..].Split(" or ").Select(x => x.Replace("Id eq ", "", StringComparison.Ordinal));
                return Task.FromResult(Json(new { value = MissingMetadata ? Array.Empty<object>() : ids.Where(Revisions.ContainsKey).Select(Row).ToArray() }));
            }
            if (path.EndsWith("/lists(guid'list')", StringComparison.Ordinal)) return Task.FromResult(Json(new { ItemCount = Items.Count - 1 }));
            throw new InvalidOperationException("Unexpected test URL: " + uri);
        }
    }
    private static SharePointReader Reader(Tokens tokens, Server server) => new(tokens, new HttpClient(server, false), new HttpClient());
    private static Task<RemoteInventory> Scan(SharePointReader reader, Workspace workspace, bool all = false) =>
        reader.GetInventoryAsync(Fixture.Site, [Fixture.Library], null, default, new InventoryReadOptions(workspace.DataRoot, all));

    [Fact]
    public async Task UnchangedDeltaAcrossReaderInstancesDoesNotFetchMetadataOrRewriteLargeSnapshot()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root); var tokens = new Tokens(); using var server = new Server();
        using (var reader = Reader(tokens, server)) Assert.Empty((await Scan(reader, workspace)).Issues);
        string snapshot = Directory.GetFiles(Path.Combine(workspace.DataRoot, "remote-state"), "*.json", SearchOption.AllDirectories).Single(p => p.Contains("snapshots"));
        byte[] bytes = await File.ReadAllBytesAsync(snapshot); var stamp = File.GetLastWriteTimeUtc(snapshot); int metadata = server.MetadataCalls;
        using (var reader = Reader(tokens, server))
        {
            var second = await Scan(reader, workspace);
            Assert.Empty(second.Issues); Assert.False(second.SnapshotChanged); Assert.True(second.Reads!.Single().Incremental);
        }
        Assert.Equal(metadata, server.MetadataCalls); Assert.Equal(stamp, File.GetLastWriteTimeUtc(snapshot)); Assert.Equal(bytes, await File.ReadAllBytesAsync(snapshot));
    }
    [Fact]
    public async Task ListDeltaFindsMetadataOnlyEditsEvenWhenDriveDoesNotChange()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root); using var server = new Server(); using var reader = Reader(new(), server);
        Assert.Equal("A", (await Scan(reader, workspace)).Files.Single().Metadata.Fields.GetProperty("Revision").GetString());
        int before = server.MetadataCalls; server.MetadataOnly("B");
        var changed = await Scan(reader, workspace);
        Assert.Empty(changed.Issues); Assert.Equal(before + 1, server.MetadataCalls);
        Assert.Equal("B", changed.Files.Single().Metadata.Fields.GetProperty("Revision").GetString());
        Assert.Contains(server.Requests, r => Uri.UnescapeDataString(r).Contains("$filter=Id eq 1", StringComparison.Ordinal));
    }
    [Fact]
    public async Task FolderRenameRebuildsDescendantPathsWithoutDescendantsInDriveDelta()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root); using var server = new Server(); using var reader = Reader(new(), server);
        server.PutFileInFolder(); Assert.Empty((await Scan(reader, workspace)).Issues);
        server.RenameFolder(); var changed = await Scan(reader, workspace);
        Assert.Empty(changed.Issues); Assert.Equal("Docs/Renamed/one.bin", changed.Files.Single().RelativePath);
        server.DeleteFile(); var deleted = await Scan(reader, workspace);
        Assert.Empty(deleted.Issues); Assert.Empty(deleted.Files); Assert.Contains("Docs/Renamed", deleted.Folders);
    }
    [Fact]
    public async Task ExpiredTokenAndMissingParentEachFallBackToOnlyOneFullRead()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root); using var server = new Server(); using var reader = Reader(new(), server);
        await Scan(reader, workspace); int full = server.FullDriveCalls;
        server.ExpireOnce = true; Assert.Empty((await Scan(reader, workspace)).Issues); Assert.Equal(full + 1, server.FullDriveCalls);
        full = server.FullDriveCalls; server.BadParentOnce = true;
        Assert.Empty((await Scan(reader, workspace)).Issues); Assert.Equal(full + 1, server.FullDriveCalls);
        full = server.FullDriveCalls; server.AlwaysExpire = true;
        Assert.NotEmpty((await Scan(reader, workspace)).Issues); Assert.Equal(full + 1, server.FullDriveCalls);
    }
    [Fact]
    public async Task SchemaChangesRefreshAllMetadataAndMissingValuesNeverPromoteState()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root); using var server = new Server(); using var reader = Reader(new(), server);
        await Scan(reader, workspace);
        string cursor = Directory.GetFiles(Path.Combine(workspace.DataRoot, "remote-state"), "cursor.json", SearchOption.AllDirectories).Single();
        byte[] before = await File.ReadAllBytesAsync(cursor);
        server.ExtraColumn = true; server.MissingColumn = true;
        var incomplete = await Scan(reader, workspace);
        Assert.Contains(incomplete.Issues, i => i.Code == "MetadataIncomplete"); Assert.Equal(before, await File.ReadAllBytesAsync(cursor));
        server.MissingColumn = false;
        var complete = await Scan(reader, workspace); Assert.Empty(complete.Issues);
        Assert.Equal(42, complete.Files.Single().Metadata.Fields.GetProperty("Code").GetInt32());
    }
    [Fact]
    public async Task ChangedAccountAndExplicitFullReadDoNotReuseCursors()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root); var tokens = new Tokens(); using var server = new Server(); using var reader = Reader(tokens, server);
        await Scan(reader, workspace); int full = server.FullDriveCalls;
        tokens.Account = "account-b"; Assert.False((await Scan(reader, workspace)).Reads!.Single().Incremental); Assert.Equal(full + 1, server.FullDriveCalls);
        Assert.False((await Scan(reader, workspace, all: true)).Reads!.Single().Incremental); Assert.Equal(full + 2, server.FullDriveCalls);
    }
    [Fact]
    public async Task MissingMetadataRowsDoNotReuseOldValuesOrAdvanceCursor()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root); using var server = new Server(); using var reader = Reader(new(), server);
        await Scan(reader, workspace);
        string cursor = Directory.GetFiles(Path.Combine(workspace.DataRoot, "remote-state"), "cursor.json", SearchOption.AllDirectories).Single();
        byte[] before = await File.ReadAllBytesAsync(cursor);
        server.MetadataOnly("B"); server.MissingMetadata = true;
        var missing = await Scan(reader, workspace);
        Assert.Contains(missing.Issues, i => i.Code == "MetadataIncomplete"); Assert.False(missing.Files.Single().Metadata.Complete);
        Assert.Equal(before, await File.ReadAllBytesAsync(cursor));
        server.MissingMetadata = false;
        Assert.Equal("B", (await Scan(reader, workspace)).Files.Single().Metadata.Fields.GetProperty("Revision").GetString());
    }
    [Fact]
    public async Task UnavailableListDeltaIsNotProbedWithAnExtraFullReadOnEveryScan()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root); using var server = new Server { ListUnavailable = true }; using var reader = Reader(new(), server);
        Assert.Empty((await Scan(reader, workspace)).Issues);
        int before = server.FullDriveCalls;
        Assert.Empty((await Scan(reader, workspace)).Issues);
        Assert.Equal(before + 1, server.FullDriveCalls);
        server.ListUnavailable = false;
        Assert.Empty((await Scan(reader, workspace, all: true)).Issues);
        Assert.True((await Scan(reader, workspace)).Reads!.Single().Incremental);
    }
    [Fact]
    public async Task EmptyPageIsFollowedAndUntrustedDeltaUrlIsNeverRequested()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root); using var server = new Server { EmptyPageOnce = true }; using var reader = Reader(new(), server);
        Assert.Empty((await Scan(reader, workspace)).Issues); Assert.Contains(server.Requests, r => r.Contains("page=1", StringComparison.Ordinal));
        server.MaliciousLink = true;
        var bad = await Scan(reader, workspace);
        Assert.Contains(bad.Issues, i => i.Code == "InvalidDeltaLink"); Assert.DoesNotContain(server.Requests, r => r.Contains("evil", StringComparison.Ordinal));
    }
}
