using System.Text.Json;

namespace SharePointDiff.Core;

public sealed partial class SharePointReader
{
    private const int StateVersion = 2;
    private readonly Dictionary<string, string> siteIds = new(StringComparer.OrdinalIgnoreCase);
    private sealed record RemoteState(int FormatVersion, string Generation, string Account, string Site, string Library, string Drive, string Folder,
        string SchemaId, Dictionary<string, JsonElement> Items, Dictionary<string, JsonElement> ListItems,
        Dictionary<string, MetadataSnapshot> Metadata, DateTimeOffset FullReadAt);
    private sealed record RemoteCursor(int FormatVersion, string Generation, string Account, string DriveLink, string ListLink, string? PreviousGeneration = null);
    private sealed record SelectionState(string Account, string Selection);
    private sealed record DeltaPage(Dictionary<string, JsonElement> Items, HashSet<string> Changed, string Link);
    private sealed record LibraryPass(RemoteState State, string DriveLink, string ListLink, IReadOnlyList<RemoteFile> Files,
        IReadOnlyList<string> Folders, IReadOnlyList<OperationIssue> Issues, bool Changed, bool Incremental);

    public async Task<RemoteInventory> GetInventoryAsync(string siteUrl, IReadOnlyList<LibraryInfo> libraries,
        IProgress<OperationProgress>? progress, CancellationToken ct, InventoryReadOptions options)
    {
        // No stable account identity means no cross-session shortcuts at all.
        await tokens.GetAsync(Graph, ct);
        if (tokens.CacheIdentity is not { Length: > 0 } account) return await GetInventoryAsync(siteUrl, libraries, progress, ct);
        columns.Clear();
        if (!siteIds.TryGetValue(siteUrl, out string? siteId))
        {
            var uri = SafePaths.SiteUri(siteUrl);
            siteId = Text(await GetAsync($"{Graph}/v1.0/sites/{uri.Host}:{uri.AbsolutePath}", ct), "id") ?? throw new OperationException("InvalidSite");
            siteIds[siteUrl] = siteId;
        }
        var files = new List<RemoteFile>(); var folders = new List<string>(); var issues = new List<OperationIssue>(); var reads = new List<LibraryReadInfo>();
        string selectionPath = Path.Combine(options.StateDirectory, "remote-selection.json");
        string selection = string.Join('|', libraries.OrderBy(l => l.Id, StringComparer.Ordinal).Select(l => l.Id + ":" + l.DriveId + ":" + l.LocalFolder));
        SelectionState? previousSelection;
        try { previousSelection = JsonFiles.Read<SelectionState>(selectionPath); } catch (JsonException) { previousSelection = null; }
        bool changed = options.ReadAll || previousSelection != new SelectionState(account, selection);
        int completed = 0;
        foreach (var library in libraries)
        {
            ct.ThrowIfCancellationRequested(); progress?.Report(new("ReadLibrary", completed, libraries.Count, library.Title));
            try
            {
                string directory = Path.Combine(options.StateDirectory, "remote-state", JsonFiles.StableId(library.Id + "|" + library.DriveId));
                var (saved, cursor) = options.ReadAll ? (null, null) : LoadState(directory, account, siteUrl, library);
                LibraryPass pass;
                bool fallbackUsed = false;
                try { pass = await ReadLibraryPassAsync(siteUrl, siteId, library, saved, cursor, ct); }
                catch (OperationException ex) when (saved != null && IsResync(ex))
                {
                    IoMetrics.Add("Remote.FullFallbacks");
                    pass = await ReadLibraryPassAsync(siteUrl, siteId, library, null, null, ct);
                    fallbackUsed = true;
                }
                if (saved != null && !fallbackUsed && pass.Issues.Any(i => i.Code == "CountChanged"))
                {
                    IoMetrics.Add("Remote.FullFallbacks");
                    pass = await ReadLibraryPassAsync(siteUrl, siteId, library, null, null, ct);
                }
                files.AddRange(pass.Files); folders.AddRange(pass.Folders); issues.AddRange(pass.Issues);
                reads.Add(new(library.Id, pass.Incremental, pass.State.FullReadAt));
                changed |= pass.Changed || pass.Issues.Count > 0;
                if (pass.Issues.Count == 0 && pass.Files.All(f => f.Metadata.Complete))
                    await SaveStateAsync(directory, pass, cursor, ct);
            }
            catch (OperationCanceledException) { throw; }
            catch (OperationException ex) when (ex.Code is "SignInRequired" or "TokenCacheUnavailable" or "AuthenticationFailed" or "AuthNetworkFailed") { throw; }
            catch (OperationException ex) when (IsUnsupportedDelta(ex))
            {
                // Some services do not expose list delta. A full, read-only scan remains usable there.
                var full = await GetInventoryAsync(siteUrl, [library], progress, ct);
                files.AddRange(full.Files); folders.AddRange(full.Folders); issues.AddRange(full.Issues); changed = true;
                reads.Add(new(library.Id, false, Time.GetUtcNow()));
            }
            catch (Exception ex) when (ex is OperationException or HttpRequestException or IOException or UnauthorizedAccessException or JsonException)
            { issues.Add(new(ex is OperationException op ? op.Code : "ReadLibraryFailed", library.Title, ex.ToString())); changed = true; }
            completed++;
        }
        if (previousSelection != new SelectionState(account, selection)) await JsonFiles.WriteAsync(selectionPath, new SelectionState(account, selection), ct);
        return new(files, folders, issues, changed, reads);
    }
    private static bool IsUnsupportedDelta(OperationException ex) => ex.Code == "RemoteMissing" ||
        ex.Code == "HttpError" && (ex.InnerException?.Message.Contains("HTTP 405", StringComparison.Ordinal) == true || ex.InnerException?.Message.Contains("HTTP 501", StringComparison.Ordinal) == true);
    private static bool IsResync(OperationException ex) => ex.Code is "DeltaExpired" or "IncompleteInventory" or "MetadataStoreCorrupt" or "InvalidDeltaLink";
    private static (RemoteState? State, RemoteCursor? Cursor) LoadState(string directory, string account, string site, LibraryInfo library)
    {
        try
        {
            string cursorPath = Path.Combine(directory, "cursor.json");
            var cursor = JsonFiles.Read<RemoteCursor>(cursorPath);
            if (cursor?.FormatVersion != StateVersion || cursor.Account != account || !Guid.TryParseExact(cursor.Generation, "N", out _)) return (null, null);
            var state = JsonFiles.Read<RemoteState>(Path.Combine(directory, "snapshots", cursor.Generation + ".json"));
            if (state?.FormatVersion != StateVersion || state.Generation != cursor.Generation || state.Account != account || state.Site != site ||
                state.Library != library.Id || state.Drive != library.DriveId || state.Folder != library.LocalFolder || state.Items == null || state.ListItems == null || state.Metadata == null) return (null, null);
            return (state, cursor);
        }
        catch (Exception ex) when (ex is JsonException or OperationException { Code: "MetadataStoreCorrupt" }) { IoMetrics.Add("Remote.InvalidState"); return (null, null); }
    }
    private async Task<LibraryPass> ReadLibraryPassAsync(string site, string siteId, LibraryInfo library, RemoteState? saved, RemoteCursor? cursor, CancellationToken ct)
    {
        var schema = await ColumnsAsync(site, library.Id, ct);
        string schemaId = SchemaStorage.Hash(schema);
        bool fullMetadata = saved == null || saved.SchemaId != schemaId;
        string driveEndpoint = $"{Graph}/v1.0/drives/{library.DriveId}/root/delta";
        string listEndpoint = $"{Graph}/v1.0/sites/{siteId}/lists/{library.Id}/items/delta";
        var drive = await DeltaAsync(cursor?.DriveLink ?? driveEndpoint + "?$select=" + Selection, driveEndpoint, saved?.Items, ct);
        var list = await DeltaAsync(cursor?.ListLink ?? listEndpoint + "?$select=id,eTag,deleted", listEndpoint, saved?.ListItems, ct);
        var metadata = fullMetadata ? new Dictionary<string, MetadataSnapshot>() : new(saved!.Metadata);
        var oldPaths = saved == null ? new Dictionary<string, string>() : ItemPaths(saved.Items, library.Title);
        var paths = ItemPaths(drive.Items, library.Title);
        var wanted = MetadataIds(drive, list, paths, oldPaths);
        bool fetched = fullMetadata || wanted.Count > 0;
        if (fetched) await UpdateMetadataAsync(site, library.Id, schema, metadata, fullMetadata ? null : wanted, ct);
        if (fetched)
        {
            // Close the enumeration/fetch window once, without polling a busy library indefinitely.
            var tailDrive = await DeltaAsync(drive.Link, driveEndpoint, drive.Items, ct);
            var tailList = await DeltaAsync(list.Link, listEndpoint, list.Items, ct);
            var tailPaths = ItemPaths(tailDrive.Items, library.Title);
            var tailWanted = MetadataIds(tailDrive, tailList, tailPaths, paths);
            if (tailWanted.Count > 0) await UpdateMetadataAsync(site, library.Id, schema, metadata, tailWanted, ct);
            drive = tailDrive with { Changed = drive.Changed.Concat(tailDrive.Changed).ToHashSet() };
            list = tailList with { Changed = list.Changed.Concat(tailList.Changed).ToHashSet() };
            paths = tailPaths;
        }
        var files = new List<RemoteFile>(); var folders = new List<string> { library.LocalFolder }; var issues = new List<OperationIssue>();
        foreach (var (id, item) in drive.Items)
        {
            if (Has(item, "root")) continue;
            string relative = library.LocalFolder + "/" + paths[id];
            try { SafePaths.Parts(relative); } catch (OperationException ex) { issues.Add(new(ex.Code, relative)); continue; }
            if (Has(item, "folder")) { folders.Add(relative); continue; }
            if (!Has(item, "file")) { issues.Add(new("UnsupportedItem", relative)); continue; }
            string listId = item.TryGetProperty("sharepointIds", out var ids) ? Text(ids, "listItemId") ?? "" : "";
            var meta = metadata.TryGetValue(listId, out var existing) ? existing : new MetadataSnapshot(JsonFiles.Empty, JsonFiles.Empty, schema, false, "Item metadata missing");
            if (!meta.Complete) issues.Add(new("MetadataIncomplete", library.Title, relative + ": " + meta.Error));
            files.Add(ToFile(site, library.Id, library.DriveId, relative, item, meta));
        }
        var count = await GetAsync($"{site}/_api/web/lists(guid'{library.Id}')?$select=ItemCount", ct);
        if (!count.TryGetProperty("ItemCount", out var total) || drive.Items.Count(p => !Has(p.Value, "root")) != total.GetInt32() || list.Items.Count != total.GetInt32())
            issues.Add(new("CountChanged", library.Title));
        // A reference missing from either complete feed is uncertainty, never an empty library.
        if (files.Any(f => !list.Items.ContainsKey(f.ListItemId))) throw new OperationException("IncompleteInventory", library.Title);
        var liveListIds = list.Items.Keys.ToHashSet();
        foreach (string deleted in metadata.Keys.Where(k => !liveListIds.Contains(k)).ToArray()) metadata.Remove(deleted);
        bool changed = saved == null || fullMetadata || drive.Changed.Count > 0 || list.Changed.Count > 0;
        var state = new RemoteState(StateVersion, changed ? Guid.NewGuid().ToString("N") : saved!.Generation, tokens.CacheIdentity!, site, library.Id,
            library.DriveId, library.LocalFolder, schemaId, drive.Items, list.Items, metadata, saved == null ? Time.GetUtcNow() : saved.FullReadAt);
        return new(state, drive.Link, list.Link, files, folders, issues, changed, saved != null);
    }
    private static HashSet<string> MetadataIds(DeltaPage drive, DeltaPage list, IReadOnlyDictionary<string, string> paths, IReadOnlyDictionary<string, string> oldPaths)
    {
        var wanted = list.Changed.Where(list.Items.ContainsKey).ToHashSet();
        foreach (var (id, item) in drive.Items)
            if (Has(item, "file") && (drive.Changed.Contains(id) || !oldPaths.TryGetValue(id, out var old) || old != paths[id]) &&
                item.TryGetProperty("sharepointIds", out var ids) && Text(ids, "listItemId") is { } listId) wanted.Add(listId);
        return wanted;
    }
    private async Task UpdateMetadataAsync(string site, string library, JsonElement schema, Dictionary<string, MetadataSnapshot> metadata, HashSet<string>? wanted, CancellationToken ct)
    {
        string query = $"{site}/_api/web/lists(guid'{library}')/items?$select={MetadataSelect(schema)}&$expand=FieldValuesAsText&$top=500";
        List<JsonElement> rows;
        if (wanted == null || wanted.Count > 100) rows = await PagesAsync(query, ct);
        else
        {
            rows = [];
            foreach (var batch in wanted.Chunk(20))
            {
                if (batch.Any(id => !long.TryParse(id, out long value) || value <= 0)) throw new OperationException("IncompleteInventory", library);
                try { rows.AddRange(await PagesAsync(query + "&$filter=" + Uri.EscapeDataString(string.Join(" or ", batch.Select(id => "Id eq " + id))), ct)); }
                catch (OperationException ex) when (ex.Code == "HttpError") { rows = await PagesAsync(query, ct); break; }
            }
        }
        foreach (var row in rows)
        {
            string id = Text(row, "Id") ?? Text(row, "ID") ?? throw new OperationException("IncompleteInventory", library);
            metadata[id] = MakeMetadata(row, schema) with { ObservedAt = Time.GetUtcNow() };
        }
        if (wanted != null)
        {
            var returned = rows.Select(r => Text(r, "Id") ?? Text(r, "ID")!).ToHashSet();
            if (wanted.Any(id => !returned.Contains(id))) throw new OperationException("IncompleteInventory", library);
        }
    }
    private async Task<DeltaPage> DeltaAsync(string url, string endpoint, Dictionary<string, JsonElement>? saved, CancellationToken ct)
    {
        var items = saved == null ? new Dictionary<string, JsonElement>() : new(saved);
        var changed = new HashSet<string>(); var seen = new HashSet<string>(); string? next = url, final = null;
        while (next != null)
        {
            ValidateDeltaLink(next, endpoint);
            if (!seen.Add(next) || seen.Count > 100000) throw new OperationException("IncompleteInventory", endpoint);
            var page = await GetAsync(next, ct);
            if (!page.TryGetProperty("value", out var values) || values.ValueKind != JsonValueKind.Array) throw new OperationException("IncompleteInventory", endpoint);
            foreach (var item in values.EnumerateArray())
            {
                string id = Text(item, "id") ?? throw new OperationException("IncompleteInventory", endpoint);
                if (Has(item, "deleted") || Has(item, "@removed")) { if (items.Remove(id)) changed.Add(id); }
                else
                {
                    var merged = items.TryGetValue(id, out var before) ? Merge(before, item) : item.Clone();
                    if (before.ValueKind == JsonValueKind.Undefined || !JsonElement.DeepEquals(before, merged) || endpoint.Contains("/lists/", StringComparison.Ordinal)) changed.Add(id);
                    items[id] = merged;
                }
            }
            next = Text(page, "@odata.nextLink"); final = Text(page, "@odata.deltaLink");
            if (final != null) { ValidateDeltaLink(final, endpoint); if (next != null) throw new OperationException("IncompleteInventory", endpoint); }
        }
        return new(items, changed, final ?? throw new OperationException("IncompleteInventory", endpoint));
    }
    private static JsonElement Merge(JsonElement before, JsonElement update)
    {
        var properties = before.EnumerateObject().ToDictionary(p => p.Name, p => p.Value.Clone(), StringComparer.Ordinal);
        foreach (var property in update.EnumerateObject()) properties[property.Name] = property.Value.Clone();
        return JsonSerializer.SerializeToElement(properties);
    }
    private static void ValidateDeltaLink(string value, string endpoint)
    {
        static string PathKey(Uri uri) => Uri.UnescapeDataString(uri.AbsolutePath).Replace("/items/root/delta", "/delta", StringComparison.Ordinal).Replace("/root/delta", "/delta", StringComparison.Ordinal);
        string expected = PathKey(new Uri(endpoint));
        if (!Uri.TryCreate(value, UriKind.Absolute, out var uri) || uri.Scheme != "https" || uri.Host != "graph.microsoft.com" || uri.Port != 443 || uri.UserInfo.Length != 0 ||
            (PathKey(uri) != expected && !(PathKey(uri).StartsWith(expected + "(", StringComparison.Ordinal) && PathKey(uri).EndsWith(')'))))
            throw new OperationException("InvalidDeltaLink", new Uri(endpoint).GetLeftPart(UriPartial.Path));
    }
    private static Dictionary<string, string> ItemPaths(Dictionary<string, JsonElement> items, string library)
    {
        var paths = new Dictionary<string, string>(); var visiting = new HashSet<string>();
        string Resolve(string id)
        {
            if (paths.TryGetValue(id, out var path)) return path;
            if (!visiting.Add(id) || !items.TryGetValue(id, out var row)) throw new OperationException("IncompleteInventory", library);
            if (Has(row, "root")) path = "";
            else
            {
                if (!row.TryGetProperty("parentReference", out var parent) || Text(parent, "id") is not { } parentId || Text(row, "name") is not { } name)
                    throw new OperationException("IncompleteInventory", library);
                string prefix = Resolve(parentId); path = prefix.Length == 0 ? name : prefix + "/" + name;
            }
            visiting.Remove(id); return paths[id] = path;
        }
        foreach (string id in items.Keys) Resolve(id);
        return paths;
    }
    private async Task SaveStateAsync(string directory, LibraryPass pass, RemoteCursor? previous, CancellationToken ct)
    {
        string schemaRoot = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(directory))!, "schemas");
        string schemaPath = Path.Combine(schemaRoot, pass.State.SchemaId + ".json");
        if (!File.Exists(schemaPath))
        {
            var schema = await ColumnsAsync(pass.State.Site, pass.State.Library, ct);
            await JsonFiles.WriteAsync(schemaPath, schema, ct);
        }
        string snapshotPath = Path.Combine(directory, "snapshots", pass.State.Generation + ".json");
        if (pass.Changed || !File.Exists(snapshotPath)) await JsonFiles.WriteNewAsync(snapshotPath, pass.State, ct);
        var cursor = new RemoteCursor(StateVersion, pass.State.Generation, pass.State.Account, pass.DriveLink, pass.ListLink,
            pass.Changed ? previous?.Generation : previous?.PreviousGeneration);
        if (cursor != previous) await JsonFiles.WriteAsync(Path.Combine(directory, "cursor.json"), cursor, ct);
        // Only disposable generations created by this implementation; never touch reports or backups.
        foreach (string path in Directory.EnumerateFiles(Path.Combine(directory, "snapshots"), "*.json"))
            if (Guid.TryParseExact(Path.GetFileNameWithoutExtension(path), "N", out _) && Path.GetFileNameWithoutExtension(path) != cursor.Generation && Path.GetFileNameWithoutExtension(path) != cursor.PreviousGeneration &&
                (File.GetAttributes(path) & FileAttributes.ReparsePoint) == 0) File.Delete(path);
    }
}
