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
    public static RemoteFile Remote(string path, byte[] bytes, string? hash = "auto", MetadataSnapshot? metadata = null)
    {
        var q = new QuickXorHasher(); q.Append(bytes);
        var schema = JsonSerializer.SerializeToElement(new[] { new { InternalName = "Revision", Title = "Revision", TypeAsString = "Text", Hidden = false, ReadOnlyField = false, FromBaseType = false } });
        metadata ??= new(JsonSerializer.SerializeToElement(new { Revision = "B", PeopleId = new[] { 7, 8 }, Taxonomy = new { Label = "Civil", TermGuid = "term" } }), JsonSerializer.SerializeToElement(new { Revision = "B" }), schema, true);
        return new(Site, "list", "drive", path, path.GetHashCode().ToString().Replace("-", ""), path, Path.GetFileName(path), bytes.Length,
            hash == "auto" ? q.Finish() : hash, "etag-1", DateTimeOffset.Parse("2025-01-02T03:04:05Z"), DateTimeOffset.Parse("2026-01-02T03:04:05Z"), Site + "/" + path, JsonFiles.Empty, metadata);
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
    public long Bytes { get; private set; }
    public int Downloads { get; private set; }
    public bool Corrupt { get; set; }
    public Action<RemoteFile>? OnDownload { get; set; }
    public Func<RemoteFile, CancellationToken, Task>? OnDownloadAsync { get; set; }
    public Action<RemoteFile>? OnRefresh { get; set; }
    public bool ChangeAfterDownload { get; set; }
    public void Add(RemoteFile file, byte[] bytes) { Files[file.ItemId] = file; Content[file.ItemId] = bytes; }
    public Task<IReadOnlyList<LibraryInfo>> GetLibrariesAsync(string site, IReadOnlyDictionary<string, string> maps, CancellationToken ct) => Task.FromResult<IReadOnlyList<LibraryInfo>>([Fixture.Library]);
    public Task<RemoteInventory> GetInventoryAsync(string site, IReadOnlyList<LibraryInfo> libs, IProgress<OperationProgress>? progress, CancellationToken ct) => Task.FromResult(new RemoteInventory(Files.Values.ToArray(), Issues.Count == 0 ? ["Docs"] : [], Issues));
    public Task<RemoteFile> RefreshAsync(RemoteFile f, CancellationToken ct) { ct.ThrowIfCancellationRequested(); OnRefresh?.Invoke(f); return Task.FromResult(Files[f.ItemId]); }
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
