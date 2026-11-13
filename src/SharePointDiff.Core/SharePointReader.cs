using System.Collections.Concurrent;
using System.Net;
using System.Net.Http.Headers;
using System.Runtime.CompilerServices;
using System.Text.Json;

namespace SharePointDiff.Core;

public sealed class SharePointReader : IRemoteReader, IDisposable
{
    private const string Graph = "https://graph.microsoft.com";
    private readonly ITokenProvider tokens;
    private readonly HttpClient api;
    private readonly HttpClient content;
    private readonly HttpClient contentRequest;
    private readonly ServiceCooldown cooldown = new();
    public TimeSpan RetryWaitBudget { get; set; } = TimeSpan.FromMinutes(15);
    // Shared by both transfer workers; one fetch per library even when they refresh concurrently on a cold reader.
    private readonly ConcurrentDictionary<string, Lazy<Task<JsonElement>>> columns = new();
    /// <summary>Clock for every timeout, stall timer and retry wait (tests use a fake one).</summary>
    public TimeProvider Time { get; init; } = TimeProvider.System;
    /// <summary>Limit for each attempt's response headers and for reading an API body (sign-in and Retry-After waits excluded);
    /// a timeout becomes RequestTimeout, never a cancellation.</summary>
    public TimeSpan RequestTimeout { get; init; } = TimeSpan.FromMinutes(2);
    /// <summary>A download that receives no byte for this long is a stall: the connection is dropped and the download resumes.</summary>
    public TimeSpan StallTimeout { get; init; } = TimeSpan.FromSeconds(60);
    /// <summary>Content attempts per DownloadAsync call (each resumes from the bytes already on disk).</summary>
    public int ResumeAttempts { get; init; } = 8;
    /// <summary>First wait between content attempts; doubles each time, capped at 60 s.</summary>
    public TimeSpan RetryDelay { get; init; } = TimeSpan.FromSeconds(5);
    public SharePointReader(ITokenProvider tokens) : this(tokens,
        HttpTransport.Client(true), HttpTransport.Client(false), HttpTransport.Client(false))
    { }
    public SharePointReader(ITokenProvider tokens, HttpClient api, HttpClient content, HttpClient? contentRequest = null)
    {
        this.tokens = tokens; this.api = api; this.content = content;
        this.contentRequest = contentRequest ?? api;
        this.contentRequest.Timeout = Timeout.InfiniteTimeSpan;
        api.Timeout = Timeout.InfiniteTimeSpan; content.Timeout = Timeout.InfiniteTimeSpan;
    }
    /// <summary>The caller's token plus a deadline on <see cref="Time"/>.</summary>
    private sealed class Deadline : IDisposable
    {
        private readonly CancellationTokenSource timer, linked;
        public Deadline(CancellationToken ct, TimeSpan after, TimeProvider time) { timer = new(after, time); linked = CancellationTokenSource.CreateLinkedTokenSource(ct, timer.Token); }
        public CancellationToken Token => linked.Token;
        public void Dispose() { linked.Dispose(); timer.Dispose(); }
    }
    private async Task<HttpResponseMessage> SendAsync(string url, string resource, CancellationToken ct)
    {
        bool refreshed = false, refreshNext = false; int busy = 0, transport = 0;
        while (true)
        {
            await cooldown.WaitAsync(new Uri(url), Time, RetryWaitBudget, ct);
            using var req = new HttpRequestMessage(HttpMethod.Get, url) { Version = HttpVersion.Version20, VersionPolicy = HttpVersionPolicy.RequestVersionOrLower };
            // Sign-in (possibly an interactive browser) runs on the caller's token: it has its own limits and is never a "request timeout".
            req.Headers.Authorization = new("Bearer", await tokens.GetAsync(resource, ct, refreshNext)); refreshNext = false;
            req.Headers.Accept.ParseAdd(resource == Graph ? "application/json" : "application/json;odata=nometadata");
            HttpResponseMessage response;
            using (var deadline = new Deadline(ct, RequestTimeout, Time))
            {
                try
                {
                    IoMetrics.Add(resource == Graph ? "Network.GraphRequests" : "Network.RestRequests");
                    bool bytes = new Uri(url).AbsolutePath.EndsWith("/content", StringComparison.Ordinal);
                    if (bytes) req.Headers.AcceptEncoding.ParseAdd("identity");
                    response = await (bytes ? contentRequest : api).SendAsync(req, HttpCompletionOption.ResponseHeadersRead, deadline.Token);
                }
                catch (OperationCanceledException ex) when (!ct.IsCancellationRequested) { throw new OperationException("RequestTimeout", url, ex); }
                // A dropped connection (common on a throttled line) is retried a few times before it surfaces.
                catch (HttpRequestException) when (transport < 3 && !ct.IsCancellationRequested) { await Task.Delay(TimeSpan.FromSeconds(2 << transport++), Time, ct); continue; }
            }
            if (response.StatusCode == HttpStatusCode.Unauthorized && !refreshed) { response.Dispose(); refreshed = refreshNext = true; continue; }
            int status = (int)response.StatusCode;
            if ((status == 429 || status >= 500) && busy < 5)
            {
                TimeSpan delay = response.Headers.RetryAfter?.Delta ??
                    (response.Headers.RetryAfter?.Date is { } date ? date - Time.GetUtcNow() : TimeSpan.FromSeconds(Math.Pow(2, busy + 1)));
                response.Dispose(); busy++;
                cooldown.Set(new Uri(url), Time.GetUtcNow() + (delay > TimeSpan.Zero ? delay : TimeSpan.FromMilliseconds(100)));
                continue;
            }
            if (!response.IsSuccessStatusCode && status is not (301 or 302 or 303 or 307 or 308))
            {
                string body = await response.Content.ReadAsStringAsync(ct);
                string requestId = response.Headers.TryGetValues("request-id", out var ids) ? string.Join(',', ids) : "—";
                response.Dispose();
                throw new OperationException(status == 403 ? "AccessDenied" : status == 404 ? "RemoteMissing" : status is 429 or >= 500 ? "ServiceUnavailable" : "HttpError", url,
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
        using var response = await SendAsync(url, resource, ct);
        using var deadline = new Deadline(ct, RequestTimeout, Time);
        try
        {
            await using var stream = new CountingReadStream(await response.Content.ReadAsStreamAsync(deadline.Token), "Network.JsonDecodedBytes");
            using var json = await JsonDocument.ParseAsync(stream, cancellationToken: deadline.Token);
            return json.RootElement.Clone();
        }
        catch (OperationCanceledException ex) when (!ct.IsCancellationRequested) { throw new OperationException("RequestTimeout", url, ex); }
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
        var entry = columns.GetOrAdd(k, _ => new(() => FetchColumnsAsync(site, library, ct)));
        var task = entry.Value;
        // A failed or cancelled fetch is forgotten so the next caller retries it.
        try { return await task; }
        finally { if (task.IsFaulted || task.IsCanceled) columns.TryRemove(new KeyValuePair<string, Lazy<Task<JsonElement>>>(k, entry)); }
    }
    private async Task<JsonElement> FetchColumnsAsync(string site, string library, CancellationToken ct) =>
        JsonSerializer.SerializeToElement(await PagesAsync($"{site}/_api/web/lists(guid'{library}')/fields?$select=Id,InternalName,Title,TypeAsString,Hidden,ReadOnlyField,FromBaseType,EntityPropertyName", ct));
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
            // Sign-in failures abort the whole inventory so nothing is hashed for nothing; other failures stay library-level issues.
            catch (Exception ex) when (ex is OperationException or HttpRequestException or IOException or UnauthorizedAccessException &&
                ex is not OperationException { Code: "SignInRequired" or "TokenCacheUnavailable" or "AuthenticationFailed" or "AuthNetworkFailed" })
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
    private static readonly string[] Resumable = ["DownloadStalled", "DownloadInterrupted", "RequestTimeout", "ServiceUnavailable", "AuthNetworkFailed"];
    /// <summary>Downloads into <paramref name="partialPath"/>, resuming with Range from the bytes already there. Stalls, dropped
    /// connections, expired download URLs, throttling and timeouts are retried within the call (each attempt asks for a fresh URL);
    /// the partial file and its state survive for a later session either way.</summary>
    public async Task<DownloadResult> DownloadAsync(RemoteFile file, string partialPath, IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(partialPath)!);
        string statePath = partialPath + ".json";
        var state = JsonFiles.Read<PartialState>(statePath);
        // Same bytes: same size and, when both sides know it, the same content hash (a metadata-only edit changes the ETag, not the bytes).
        bool resume = state != null && state.Size == file.Size &&
            (state.QuickXorHash is { Length: > 0 } saved && file.QuickXorHash is { Length: > 0 } current ? saved == current : state.ETag == file.ETag);
        var received = new StrongBox<long>(); int stalls = 0;
        for (int attempt = 1; ; attempt++)
        {
            long offset = (attempt > 1 || resume) && File.Exists(partialPath) ? new FileInfo(partialPath).Length : 0;
            if (offset > file.Size) offset = 0;
            await JsonFiles.WriteAsync(statePath, new PartialState(file.ETag, file.Size, file.QuickXorHash, file.LibraryId), ct);
            try
            {
                await DownloadOnceAsync(file, partialPath, offset, received, progress, ct);
                return new(received.Value, attempt, stalls);
            }
            catch (OperationException ex) when (Resumable.Contains(ex.Code) && attempt < ResumeAttempts && !ct.IsCancellationRequested)
            {
                if (ex.Code == "DownloadStalled") stalls++;
                double wait = Math.Min(RetryDelay.TotalSeconds * Math.Pow(2, attempt - 1), 60);
                await Task.Delay(TimeSpan.FromSeconds(wait), Time, ct);
            }
        }
    }
    private async Task DownloadOnceAsync(RemoteFile file, string partialPath, long offset, StrongBox<long> received, IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        // Asked again on every attempt: pre-authenticated download URLs expire.
        using var redirect = await SendAsync($"{Graph}/v1.0/drives/{file.DriveId}/items/{file.ItemId}/content", Graph, ct);
        if (redirect.Headers.Location is { } location)
        {
            if (!location.IsAbsoluteUri || location.Scheme != "https") throw new OperationException("InvalidDownload", file.RelativePath);
            using var response = await SendContentAsync(location, offset, file, ct);
            if (response.StatusCode == HttpStatusCode.RequestedRangeNotSatisfiable && offset == file.Size) return;
            int status = (int)response.StatusCode;
            if (status is 401 or 403) throw new OperationException("DownloadInterrupted", file.RelativePath); // expired URL: ask /content again
            if (status is 404 or 410) throw new OperationException("RemoteMissing", file.RelativePath);
            if (status == 429 || status >= 500) throw new OperationException("ServiceUnavailable", file.RelativePath);
            if (!response.IsSuccessStatusCode) throw new OperationException("InvalidDownload", file.RelativePath);
            if (response.StatusCode != HttpStatusCode.PartialContent) offset = 0;
            else if (response.Content.Headers.ContentRange?.From != offset) throw new OperationException("InvalidDownload", file.RelativePath);
            await CopyContentAsync(response, partialPath, offset, file, received, progress, ct);
        }
        else
        {
            // A server may return bytes directly instead of a redirect. No bearer token is forwarded.
            if (!redirect.IsSuccessStatusCode) throw new OperationException("InvalidDownload", file.RelativePath);
            await CopyContentAsync(redirect, partialPath, 0, file, received, progress, ct);
        }
    }
    private async Task<HttpResponseMessage> SendContentAsync(Uri location, long offset, RemoteFile file, CancellationToken ct)
    {
        for (int attempt = 0; ; attempt++)
        {
            await cooldown.WaitAsync(location, Time, RetryWaitBudget, ct);
            using var req = new HttpRequestMessage(HttpMethod.Get, location) { Version = HttpVersion.Version20, VersionPolicy = HttpVersionPolicy.RequestVersionOrLower };
            req.Headers.AcceptEncoding.ParseAdd("identity");
            if (offset > 0) req.Headers.Range = new RangeHeaderValue(offset, null);
            HttpResponseMessage response;
            using (var deadline = new Deadline(ct, RequestTimeout, Time))
            {
                try { IoMetrics.Add("Network.ContentRequests"); response = await content.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, deadline.Token); }
                catch (OperationCanceledException ex) when (!ct.IsCancellationRequested) { throw new OperationException("RequestTimeout", file.RelativePath, ex); }
                catch (HttpRequestException ex) when (!ct.IsCancellationRequested) { throw new OperationException("DownloadInterrupted", file.RelativePath, ex); }
            }
            if (((int)response.StatusCode == 429 || (int)response.StatusCode >= 500) && attempt < 5)
            {
                var wait = response.Headers.RetryAfter?.Delta ?? (response.Headers.RetryAfter?.Date is { } date ? date - Time.GetUtcNow() : TimeSpan.FromSeconds(Math.Pow(2, attempt + 1)));
                response.Dispose(); cooldown.Set(location, Time.GetUtcNow() + (wait > TimeSpan.Zero ? wait : TimeSpan.FromMilliseconds(100))); continue;
            }
            return response;
        }
    }
    private async Task CopyContentAsync(HttpResponseMessage response, string path, long offset, RemoteFile file, StrongBox<long> received, IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        await using var input = await response.Content.ReadAsStreamAsync(ct);
        await using var output = new FileStream(path, offset > 0 ? FileMode.Append : FileMode.Create, FileAccess.Write, FileShare.None, 1024 * 1024, true);
        var buffer = new byte[1024 * 1024]; long written = 0; long last = Environment.TickCount64;
        while (true)
        {
            int n;
            // Only the network read is classified; a failure writing the local file propagates unchanged and is never retried.
            using (var idle = new Deadline(ct, StallTimeout, Time))
            {
                try { n = await input.ReadAsync(buffer, idle.Token); }
                catch (OperationCanceledException ex) when (!ct.IsCancellationRequested) { throw new OperationException("DownloadStalled", file.RelativePath, ex); }
                catch (Exception ex) when (ex is IOException or HttpRequestException && !ct.IsCancellationRequested) { throw new OperationException("DownloadInterrupted", file.RelativePath, ex); }
            }
            if (n == 0) break;
            if (offset + written + n > file.Size) throw new OperationException("HashMismatch", file.RelativePath);
            await output.WriteAsync(buffer.AsMemory(0, n), ct); written += n; received.Value += n;
            if (Environment.TickCount64 - last > 150) { progress?.Report(new("Downloading", Path: file.RelativePath, Bytes: offset + written, TotalBytes: file.Size)); last = Environment.TickCount64; }
        }
        await output.FlushAsync(ct);
    }
    public void Dispose() { api.Dispose(); content.Dispose(); if (!ReferenceEquals(contentRequest, api)) contentRequest.Dispose(); }
    private sealed record PartialState(string ETag, long Size, string? QuickXorHash = null, string? LibraryId = null);
}
