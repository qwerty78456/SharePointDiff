using System.Text.Json;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class Fixture : IDisposable
{
    private readonly string container;
    public string Root { get; }
    public const string Site = "https://contoso.sharepoint.com/sites/review";
    public static LibraryInfo Library => new("list", "drive", "Docs", "/sites/review/Shared Documents", "Docs", 1);
    public Fixture()
    {
        string testRoot = Environment.GetEnvironmentVariable("SPDIFF_TEST_ROOT") ?? Path.Combine(Path.GetTempPath(), "SharePointDiffTests");
        container = Path.Combine(testRoot, Guid.NewGuid().ToString("N"));
        Root = Path.Combine(container, "mirror"); Directory.CreateDirectory(Root);
    }
    public async Task<LocalFile> LocalAsync(string path, byte[] bytes)
    {
        string full = SafePaths.Resolve(Root, path, true); await File.WriteAllBytesAsync(full, bytes);
        return new(path, full, await FileHashing.ComputeAsync(full), File.GetLastWriteTimeUtc(full));
    }
    public static RemoteFile Remote(string path, byte[] bytes, string? hash = "auto", MetadataSnapshot? metadata = null, string etag = "etag-1")
    {
        var q = new QuickXorHasher(); q.Append(bytes);
        var schema = JsonSerializer.SerializeToElement(new[] { new { InternalName = "Revision", Title = "Revision", TypeAsString = "Text", Hidden = false, ReadOnlyField = false, FromBaseType = false } });
        metadata ??= new(JsonSerializer.SerializeToElement(new { Revision = "B", PeopleId = new[] { 7, 8 }, Taxonomy = new { Label = "Civil", TermGuid = "term" } }), JsonSerializer.SerializeToElement(new { Revision = "B" }), schema, true);
        return new(Site, "list", "drive", path, path.GetHashCode().ToString().Replace("-", ""), path, Path.GetFileName(path), bytes.Length,
            hash == "auto" ? q.Finish() : hash, etag, DateTimeOffset.Parse("2025-01-02T03:04:05Z"), DateTimeOffset.Parse("2026-01-02T03:04:05Z"), Site + "/" + path, JsonFiles.Empty, metadata);
    }
    public static ScanResult Scan(string root, params DiffEntry[] entries) => new(Site, root, entries, ["Docs"], [], DateTimeOffset.Now);
    public static SyncPlan Plan(params WorkItem[] work) => new(work, work.Count(x => x.Entry.ReusePath != null), work.Where(x => x.Entry.ReusePath == null).Sum(x => x.Entry.Remote!.Size), work.Where(x => x.Replace).Sum(x => x.Entry.Local!.Hashes.Size));
    public void Dispose()
    {
        // Only remove this fixture's GUID directory, never a user-provided path.
        var p = new DirectoryInfo(container); if (!Guid.TryParseExact(p.Name, "N", out _)) throw new InvalidOperationException();
        SafePaths.Inside(p.Parent!.FullName, container);
        if (Directory.Exists(container))
        {
            foreach (string file in Directory.EnumerateFiles(container, "*", SearchOption.AllDirectories)) File.SetAttributes(file, FileAttributes.Normal);
            Directory.Delete(container, true);
        }
    }
}
public sealed class FakeRemote : IRemoteReader
{
    public Dictionary<string, byte[]> Content { get; } = [];
    public Dictionary<string, RemoteFile> Files { get; } = [];
    public List<OperationIssue> Issues { get; } = [];
    /// <summary>Library-level issues keyed by library title; the library is then reported as not fully read.</summary>
    public Dictionary<string, OperationIssue> IssuesByLibrary { get; } = [];
    public List<LibraryInfo> Libraries { get; } = [Fixture.Library];
    public Exception? LibrariesFault { get; set; }
    /// <summary>One-shot faults, thrown by the next GetLibrariesAsync calls in order.</summary>
    public Queue<Exception> LibraryFaults { get; } = new();
    public long Bytes { get; private set; }
    public int Downloads { get; private set; }
    public int InventoryCalls { get; private set; }
    public int RefreshCalls { get; private set; }
    public bool Corrupt { get; set; }
    public Action<RemoteFile>? OnDownload { get; set; }
    public Func<RemoteFile, CancellationToken, Task>? OnDownloadAsync { get; set; }
    public Action<RemoteFile>? OnRefresh { get; set; }
    public Func<RemoteFile, CancellationToken, Task>? OnRefreshAsync { get; set; }
    public bool ChangeAfterDownload { get; set; }
    public void Add(RemoteFile file, byte[] bytes) { Files[file.ItemId] = file; Content[file.ItemId] = bytes; }
    public void ResetCounters() { Downloads = 0; InventoryCalls = 0; RefreshCalls = 0; }
    public Task<IReadOnlyList<LibraryInfo>> GetLibrariesAsync(string site, IReadOnlyDictionary<string, string> maps, CancellationToken ct)
    {
        if (LibrariesFault != null) throw LibrariesFault;
        if (LibraryFaults.TryDequeue(out var fault)) throw fault;
        return Task.FromResult<IReadOnlyList<LibraryInfo>>(Libraries.Select(l => maps.TryGetValue(l.Id, out var folder) ? l with { LocalFolder = folder } : l).ToArray());
    }
    public Task<RemoteInventory> GetInventoryAsync(string site, IReadOnlyList<LibraryInfo> libs, IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        InventoryCalls++;
        var files = Files.Values.Where(f => libs.Any(l => SafePaths.Key(f.RelativePath).StartsWith(SafePaths.Key(l.LocalFolder) + "/", StringComparison.Ordinal))).ToArray();
        var issues = Issues.Concat(libs.Where(l => IssuesByLibrary.ContainsKey(l.Title)).Select(l => IssuesByLibrary[l.Title])).ToArray();
        string[] folders = Issues.Count == 0 ? libs.Where(l => !IssuesByLibrary.ContainsKey(l.Title)).Select(l => l.LocalFolder).ToArray() : [];
        return Task.FromResult(new RemoteInventory(files, folders, issues));
    }
    public async Task<RemoteFile> RefreshAsync(RemoteFile f, CancellationToken ct)
    {
        RefreshCalls++; ct.ThrowIfCancellationRequested(); OnRefresh?.Invoke(f);
        if (OnRefreshAsync != null) await OnRefreshAsync(f, ct);
        return Files[f.ItemId];
    }
    public async Task<DownloadResult> DownloadAsync(RemoteFile file, string partial, IProgress<OperationProgress>? p, CancellationToken ct)
    {
        Downloads++; OnDownload?.Invoke(file);
        if (OnDownloadAsync != null) await OnDownloadAsync(file, ct);
        ct.ThrowIfCancellationRequested(); var bytes = Content[file.ItemId].ToArray(); if (Corrupt && bytes.Length > 0) bytes[0] ^= 255;
        await File.WriteAllBytesAsync(partial, bytes, ct); Bytes += bytes.Length;
        if (ChangeAfterDownload) Files[file.ItemId] = file with { ETag = "changed" };
        return new(bytes.Length);
    }
}
/// <summary>Records progress synchronously (Progress&lt;T&gt; posts asynchronously and would make assertions racy).</summary>
public sealed class SyncProgress : IProgress<OperationProgress>
{
    private readonly List<OperationProgress> items = [];
    public IReadOnlyList<OperationProgress> Items { get { lock (items) return items.ToArray(); } }
    public void Report(OperationProgress value) { lock (items) items.Add(value); }
}
