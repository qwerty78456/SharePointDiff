using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;

namespace SharePointDiff.Core;

public sealed class SharePointReader : IRemoteReader, IDisposable
{
    private const string Graph = "https://graph.microsoft.com";
    private readonly ITokenProvider tokens;
    private readonly HttpClient api;
    private readonly HttpClient content;
    private readonly Dictionary<string, JsonElement> columns = [];
    public SharePointReader(ITokenProvider tokens) : this(tokens,
        new HttpClient(new HttpClientHandler { AllowAutoRedirect = false }), new HttpClient())
    { }
    public SharePointReader(ITokenProvider tokens, HttpClient api, HttpClient content)
    {
        this.tokens = tokens; this.api = api; this.content = content;
        api.Timeout = Timeout.InfiniteTimeSpan; content.Timeout = Timeout.InfiniteTimeSpan;
    }
    private async Task<HttpResponseMessage> SendAsync(string url, string resource, CancellationToken ct)
    {
        for (int attempt = 0; ; attempt++)
        {
            using var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.Authorization = new("Bearer", await tokens.GetAsync(resource, ct, attempt == 1));
            req.Headers.Accept.ParseAdd(resource == Graph ? "application/json" : "application/json;odata=nometadata");
            var response = await api.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
            if (response.StatusCode == HttpStatusCode.Unauthorized && attempt == 0) { response.Dispose(); continue; }
            int status = (int)response.StatusCode;
            if ((status == 429 || status >= 500) && attempt < 5)
            {
                TimeSpan delay = response.Headers.RetryAfter?.Delta ??
                    (response.Headers.RetryAfter?.Date is { } date ? date - DateTimeOffset.UtcNow : TimeSpan.FromSeconds(Math.Pow(2, attempt + 1)));
                response.Dispose();
                await Task.Delay(TimeSpan.FromMilliseconds(Math.Clamp(delay.TotalMilliseconds, 100, 120000)), ct); continue;
            }
            if (!response.IsSuccessStatusCode && status is not (301 or 302 or 303 or 307 or 308))
            {
                string body = await response.Content.ReadAsStringAsync(ct);
                string requestId = response.Headers.TryGetValues("request-id", out var ids) ? string.Join(',', ids) : "—";
                response.Dispose();
                throw new OperationException(status == 403 ? "AccessDenied" : status == 404 ? "RemoteMissing" : "HttpError", url,
                    new HttpRequestException($"HTTP {status}; request-id: {requestId}; response: {body[..Math.Min(body.Length, 8000)]}"));
            }
            return response;
        }
    }
    private async Task<JsonElement> GetAsync(string url, CancellationToken ct)
    {
        var uri = new Uri(url);
        string resource = uri.Host.Equals("graph.microsoft.com", StringComparison.OrdinalIgnoreCase) ? Graph : uri.GetLeftPart(UriPartial.Authority);
        if (resource != Graph && !uri.Host.EndsWith(".sharepoint.com", StringComparison.OrdinalIgnoreCase)) throw new OperationException("InvalidSite");
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct); timeout.CancelAfter(TimeSpan.FromMinutes(2));
        using var response = await SendAsync(url, resource, timeout.Token);
        await using var stream = await response.Content.ReadAsStreamAsync(timeout.Token);
        using var json = await JsonDocument.ParseAsync(stream, cancellationToken: timeout.Token);
        return json.RootElement.Clone();
    }
    private static string? Text(JsonElement e, string name) => e.TryGetProperty(name, out var p) && p.ValueKind != JsonValueKind.Null ? p.ToString() : null;
    private static bool Has(JsonElement e, string name) => e.TryGetProperty(name, out _);
    private async Task<List<JsonElement>> PagesAsync(string url, CancellationToken ct)
    {
        var items = new List<JsonElement>();
        string? next = url;
        while (next != null)
        {
            var page = await GetAsync(next, ct);
            if (!page.TryGetProperty("value", out var values)) throw new OperationException("IncompleteInventory");
            items.AddRange(values.EnumerateArray().Select(x => x.Clone()));
            next = Text(page, "@odata.nextLink") ?? Text(page, "odata.nextLink");
        }
        return items;
    }
    public async Task<IReadOnlyList<LibraryInfo>> GetLibrariesAsync(string siteUrl, IReadOnlyDictionary<string, string> mappings, CancellationToken ct)
    {
        var uri = SafePaths.SiteUri(siteUrl);
        var site = await GetAsync($"{Graph}/v1.0/sites/{uri.Host}:{uri.AbsolutePath}", ct);
        string siteId = Text(site, "id") ?? throw new OperationException("InvalidSite");
        var lists = await PagesAsync($"{uri}/_api/web/lists?$select=Id,Title,Hidden,BaseTemplate,ItemCount,RootFolder/ServerRelativeUrl&$expand=RootFolder", ct);
        var result = new List<LibraryInfo>();
        var used = new HashSet<string>();
        foreach (var list in lists.Where(x => !x.GetProperty("Hidden").GetBoolean() && x.GetProperty("BaseTemplate").GetInt32() == 101))
        {
            string id = Text(list, "Id")!, title = Text(list, "Title")!;
            var drive = await GetAsync($"{Graph}/v1.0/sites/{siteId}/lists/{id}/drive?$select=id", ct);
            string rootUrl = list.GetProperty("RootFolder").GetProperty("ServerRelativeUrl").GetString()!;
            string leaf = Uri.UnescapeDataString(rootUrl.TrimEnd('/').Split('/').Last());
            // Built-in display names vary with the account's language; keep the mirror mapping stable.
            string defaultFolder = leaf switch { "Shared Documents" => "Documents", "SiteAssets" => "Site Assets", "FormServerTemplates" => "Form Templates", "Style Library" => "Style Library", _ => title };
            string folder = mappings.TryGetValue(id, out var saved) ? saved : defaultFolder;
            SafePaths.Parts(folder);
            if (folder.Contains('/') || folder.Contains('\\') || !used.Add(SafePaths.Key(folder))) throw new OperationException("PathCollision", folder);
            result.Add(new(id, Text(drive, "id")!, title, rootUrl, folder, list.GetProperty("ItemCount").GetInt32()));
        }
        return result;
    }
    private async Task<JsonElement> ColumnsAsync(string site, string library, CancellationToken ct)
    {
        string k = site + library;
        if (columns.TryGetValue(k, out var saved)) return saved;
        var rows = await PagesAsync($"{site}/_api/web/lists(guid'{library}')/fields?$select=Id,InternalName,Title,TypeAsString,Hidden,ReadOnlyField,FromBaseType,EntityPropertyName", ct);
        var value = JsonSerializer.SerializeToElement(rows);
        columns[k] = value; return value;
    }
    private static string MetadataSelect(JsonElement schema)
    {
        var properties = schema.EnumerateArray().Where(c => !c.GetProperty("Hidden").GetBoolean() && !c.GetProperty("ReadOnlyField").GetBoolean() && !c.GetProperty("FromBaseType").GetBoolean())
            .Where(c => !(Text(c, "TypeAsString") ?? "").StartsWith("User") && !(Text(c, "TypeAsString") ?? "").StartsWith("Lookup"))
            .Select(c => Text(c, "EntityPropertyName") ?? Text(c, "InternalName")!).Distinct();
        return "*,FieldValuesAsText," + string.Join(',', properties);
    }
    private static MetadataSnapshot MakeMetadata(JsonElement item, JsonElement schema)
    {
        var fields = new Dictionary<string, JsonElement>();
        foreach (var p in item.EnumerateObject()) if (p.Name != "FieldValuesAsText" && !p.Name.StartsWith("odata.") && !p.Name.StartsWith('@')) fields[p.Name] = p.Value.Clone();
        JsonElement display = item.TryGetProperty("FieldValuesAsText", out var text) ? text.Clone() : JsonFiles.Empty;
        var rawFields = JsonSerializer.SerializeToElement(fields);
        var missing = new List<string>();
        foreach (var column in schema.EnumerateArray())
        {
            string n = Text(column, "InternalName")!;
            if (column.GetProperty("Hidden").GetBoolean() || column.GetProperty("ReadOnlyField").GetBoolean() || column.GetProperty("FromBaseType").GetBoolean()) continue;
            if (!MetadataFields.TryGet(rawFields, column, out _) && !MetadataFields.TryGet(display, column, out _)) missing.Add(n);
        }
        return new(JsonSerializer.SerializeToElement(fields), display, schema, missing.Count == 0, missing.Count == 0 ? null : string.Join(", ", missing));
    }
    private static RemoteFile ToFile(string site, string lib, string drive, string path, JsonElement item, MetadataSnapshot meta)
    {
        string? hash = null;
        if (item.TryGetProperty("file", out var f) && f.TryGetProperty("hashes", out var h)) hash = Text(h, "quickXorHash");
        string listItem = item.TryGetProperty("sharepointIds", out var ids) ? Text(ids, "listItemId") ?? "" : "";
        return new(site, lib, drive, Text(item, "id")!, listItem, path, Text(item, "name")!, item.GetProperty("size").GetInt64(), hash,
            Text(item, "eTag") ?? "", DateTimeOffset.Parse(Text(item, "createdDateTime")!), DateTimeOffset.Parse(Text(item, "lastModifiedDateTime")!),
            Text(item, "webUrl") ?? "", item.Clone(), meta);
    }
    private const string Selection = "id,name,size,file,folder,parentReference,webUrl,lastModifiedDateTime,createdDateTime,deleted,root,eTag,sharepointIds,createdBy,lastModifiedBy";
    public async Task<RemoteInventory> GetInventoryAsync(string siteUrl, IReadOnlyList<LibraryInfo> libraries, IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        columns.Clear();
        var files = new List<RemoteFile>(); var folders = new List<string>(); var issues = new List<OperationIssue>();
        int completed = 0;
        foreach (var lib in libraries)
        {
            ct.ThrowIfCancellationRequested(); progress?.Report(new("ReadLibrary", completed, libraries.Count, lib.Title));
            try
            {
                var schema = await ColumnsAsync(siteUrl, lib.Id, ct);
                var metadataRows = await PagesAsync($"{siteUrl}/_api/web/lists(guid'{lib.Id}')/items?$select={MetadataSelect(schema)}&$expand=FieldValuesAsText&$top=500", ct);
                var metadata = metadataRows.ToDictionary(x => Text(x, "Id") ?? Text(x, "ID")!, x => MakeMetadata(x, schema));
                var items = new Dictionary<string, JsonElement>();
                string? next = $"{Graph}/v1.0/drives/{lib.DriveId}/root/delta?$select={Selection}";
                bool complete = false;
                while (next != null)
                {
                    var page = await GetAsync(next, ct);
                    foreach (var item in page.GetProperty("value").EnumerateArray())
                    {
                        string id = Text(item, "id")!;
                        if (Has(item, "deleted")) items.Remove(id); else items[id] = item.Clone();
                    }
                    next = Text(page, "@odata.nextLink"); complete = Has(page, "@odata.deltaLink");
                }
                if (!complete) throw new OperationException("IncompleteInventory", lib.Title);
                var paths = new Dictionary<string, string>();
                string ItemPath(string id, HashSet<string>? seen = null)
                {
                    if (paths.TryGetValue(id, out var known)) return known;
                    seen ??= []; if (!seen.Add(id) || !items.TryGetValue(id, out var row)) throw new OperationException("IncompleteInventory", lib.Title);
                    if (Has(row, "root")) return paths[id] = "";
                    string parent = row.GetProperty("parentReference").GetProperty("id").GetString()!;
                    string prefix = ItemPath(parent, seen), name = Text(row, "name")!;
                    return paths[id] = prefix.Length == 0 ? name : prefix + "/" + name;
                }
                var libFiles = new List<RemoteFile>(); var libFolders = new List<string> { lib.LocalFolder };
                foreach (var pair in items)
                {
                    var item = pair.Value; if (Has(item, "root")) continue;
                    string path = lib.LocalFolder + "/" + ItemPath(pair.Key);
                    try { SafePaths.Parts(path); }
                    catch (OperationException ex) { issues.Add(new(ex.Code, path)); continue; }
                    if (Has(item, "folder")) { libFolders.Add(path); continue; }
                    if (!Has(item, "file")) { issues.Add(new("UnsupportedItem", path)); continue; }
                    string li = item.TryGetProperty("sharepointIds", out var sp) ? Text(sp, "listItemId") ?? "" : "";
                    var meta = metadata.TryGetValue(li, out var m) ? m : new MetadataSnapshot(JsonFiles.Empty, JsonFiles.Empty, schema, false, "Item metadata missing");
                    libFiles.Add(ToFile(siteUrl, lib.Id, lib.DriveId, path, item, meta));
                }
                if (libFiles.Count + libFolders.Count - 1 != lib.ItemCount) issues.Add(new("CountChanged", lib.Title));
                files.AddRange(libFiles); folders.AddRange(libFolders);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex) when (ex is OperationException or HttpRequestException or IOException or UnauthorizedAccessException)
            { issues.Add(new(ex is OperationException op ? op.Code : "ReadLibraryFailed", lib.Title, ex.ToString())); }
            completed++;
        }
        return new(files, folders, issues);
    }
    public async Task<RemoteFile> RefreshAsync(RemoteFile file, CancellationToken ct)
    {
        var row = await GetAsync($"{Graph}/v1.0/drives/{file.DriveId}/items/{file.ItemId}?$select={Selection}", ct);
        if (!string.Equals(SafePaths.Key(Text(row, "name")!), SafePaths.Key(file.Name), StringComparison.Ordinal)) throw new OperationException("RemoteChanged", file.RelativePath);
        if (row.TryGetProperty("parentReference", out var parent) && parent.TryGetProperty("path", out var pp) && pp.GetString() is { } raw)
        {
            int i = raw.IndexOf("root:", StringComparison.Ordinal);
            if (i >= 0)
            {
                string expected = file.RelativePath[(file.RelativePath.IndexOf('/') + 1)..];
                string actual = Uri.UnescapeDataString(raw[(i + 5)..]).Trim('/') + "/" + Text(row, "name");
                if (SafePaths.Key(actual.TrimStart('/')) != SafePaths.Key(expected)) throw new OperationException("RemoteChanged", file.RelativePath);
            }
        }
        var schema = await ColumnsAsync(file.SiteUrl, file.LibraryId, ct);
        var meta = await GetAsync($"{file.SiteUrl}/_api/web/lists(guid'{file.LibraryId}')/items({file.ListItemId})?$select={MetadataSelect(schema)}&$expand=FieldValuesAsText", ct);
        return ToFile(file.SiteUrl, file.LibraryId, file.DriveId, file.RelativePath, row, MakeMetadata(meta, schema));
    }
    public async Task<DownloadResult> DownloadAsync(RemoteFile file, string partialPath, IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(partialPath)!);
        string statePath = partialPath + ".json";
        var state = JsonFiles.Read<PartialState>(statePath);
        bool resume = state?.ETag == file.ETag && state.Size == file.Size;
        long offset = resume && File.Exists(partialPath) ? new FileInfo(partialPath).Length : 0;
        if (offset > file.Size) offset = 0;
        await JsonFiles.WriteAsync(statePath, new PartialState(file.ETag, file.Size), ct);
        using var redirect = await SendAsync($"{Graph}/v1.0/drives/{file.DriveId}/items/{file.ItemId}/content", Graph, ct);
        long received = 0;
        if (redirect.Headers.Location is { } location)
        {
            if (!location.IsAbsoluteUri || location.Scheme != "https") throw new OperationException("InvalidDownload", file.RelativePath);
            using var response = await SendContentAsync(location, offset, ct);
            if (response.StatusCode == HttpStatusCode.RequestedRangeNotSatisfiable && offset == file.Size) return new(0);
            response.EnsureSuccessStatusCode();
            if (response.StatusCode != HttpStatusCode.PartialContent) offset = 0;
            else if (response.Content.Headers.ContentRange?.From != offset) throw new OperationException("InvalidDownload", file.RelativePath);
            received = await CopyContentAsync(response, partialPath, offset, file, progress, ct);
        }
        else
        {
            // A server may return bytes directly instead of a redirect. No bearer token is forwarded.
            if (!redirect.IsSuccessStatusCode) throw new OperationException("InvalidDownload", file.RelativePath);
            received = await CopyContentAsync(redirect, partialPath, 0, file, progress, ct);
        }
        return new(received);
    }
    private async Task<HttpResponseMessage> SendContentAsync(Uri location, long offset, CancellationToken ct)
    {
        for (int attempt = 0; ; attempt++)
        {
            using var req = new HttpRequestMessage(HttpMethod.Get, location);
            if (offset > 0) req.Headers.Range = new RangeHeaderValue(offset, null);
            var response = await content.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
            if (((int)response.StatusCode == 429 || (int)response.StatusCode >= 500) && attempt < 5)
            {
                var wait = response.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(Math.Pow(2, attempt + 1));
                response.Dispose(); await Task.Delay(TimeSpan.FromMilliseconds(Math.Clamp(wait.TotalMilliseconds, 100, 120000)), ct); continue;
            }
            return response;
        }
    }
    private static async Task<long> CopyContentAsync(HttpResponseMessage response, string path, long offset, RemoteFile file, IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        await using var input = await response.Content.ReadAsStreamAsync(ct);
        await using var output = new FileStream(path, offset > 0 ? FileMode.Append : FileMode.Create, FileAccess.Write, FileShare.None, 1024 * 1024, true);
        var buffer = new byte[1024 * 1024]; long received = 0; long last = Environment.TickCount64;
        int n;
        while ((n = await input.ReadAsync(buffer, ct)) > 0)
        {
            if (offset + received + n > file.Size) throw new OperationException("HashMismatch", file.RelativePath);
            await output.WriteAsync(buffer.AsMemory(0, n), ct); received += n;
            if (Environment.TickCount64 - last > 150) { progress?.Report(new("Downloading", Path: file.RelativePath, Bytes: offset + received, TotalBytes: file.Size)); last = Environment.TickCount64; }
        }
        await output.FlushAsync(ct); return received;
    }
    public void Dispose() { api.Dispose(); content.Dispose(); }
    private sealed record PartialState(string ETag, long Size);
}
