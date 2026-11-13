using System.Text.Json;

namespace SharePointDiff.Core;

public enum DiffStatus { Match, MissingLocal, DifferentSize, DifferentHash, LocalOnly, Unverified, Error }
public enum OverwriteMode { KeepLocal, Selected, AllDifferences }
public sealed record ConnectionProfile(string Name, string TenantId, string ClientId);
public sealed record LibraryInfo(string Id, string DriveId, string Title, string RootUrl, string LocalFolder, int ItemCount);
public sealed record FileHashes(long Size, string QuickXorHash, string Sha256);
public enum HashBasis { Fresh, Saved }
public sealed record LocalFile(string RelativePath, string FullPath, FileHashes Hashes, DateTime LastWriteUtc,
    HashBasis HashBasis = HashBasis.Fresh, DateTimeOffset? LastHashedAt = null, FileIdentity? Identity = null);
public sealed record ScanOptions(bool RehashAllLocal = false, bool ReadAllRemote = false, bool Background = false);
public sealed record MetadataSnapshot(JsonElement Fields, JsonElement DisplayValues, JsonElement Columns, bool Complete, string? Error = null, DateTimeOffset? ObservedAt = null);
public sealed record RemoteFile(
    string SiteUrl, string LibraryId, string DriveId, string ItemId, string ListItemId,
    string RelativePath, string Name, long Size, string? QuickXorHash, string ETag,
    DateTimeOffset Created, DateTimeOffset Modified, string WebUrl, JsonElement Basic,
    MetadataSnapshot Metadata);
public sealed record LibraryReadInfo(string LibraryId, bool Incremental, DateTimeOffset FullReadAt);
public sealed record RemoteInventory(IReadOnlyList<RemoteFile> Files, IReadOnlyList<string> Folders, IReadOnlyList<OperationIssue> Issues,
    bool SnapshotChanged = true, IReadOnlyList<LibraryReadInfo>? Reads = null);
public sealed record InventoryReadOptions(string StateDirectory, bool ReadAll = false);
public sealed record OperationIssue(string Code, string Path, string? Detail = null);
public sealed class DiffEntry
{
    public required string RelativePath { get; init; }
    public required DiffStatus Status { get; init; }
    public LocalFile? Local { get; init; }
    public RemoteFile? Remote { get; init; }
    public string? ReusePath { get; set; }
    public string? ReuseSha256 { get; set; }
    public OperationIssue? Issue { get; init; }
    public bool IsDifference => Status is DiffStatus.DifferentHash or DiffStatus.DifferentSize;
}
public sealed record ScanResult(string SiteUrl, string LocalRoot, IReadOnlyList<DiffEntry> Entries,
    IReadOnlyList<string> RemoteFolders, IReadOnlyList<OperationIssue> Issues, DateTimeOffset ScannedAt, IReadOnlyList<LibraryInfo>? Libraries = null, IoStatistics? Statistics = null,
    IReadOnlyList<LibraryReadInfo>? RemoteReads = null);
/// <summary>Libraries seen (<see cref="Known"/>) and ticked (<see cref="Included"/>) at the last completed Compare in the main window.</summary>
public sealed record LibrarySelection(IReadOnlyList<string> Known, IReadOnlyList<string> Included, DateTimeOffset ComparedAt);
public enum Provenance { Pristine, Revert, Conflict, Unknown }
public sealed record CheckSummary(int NewFiles, long NewBytes, int Updates, long UpdateBytes, int PristineUpdates, int Conflicts, int Reverts, int Unknown,
    int NoHash, int Review, int LocalOnly, int Reused, long NetworkBytes, long BackupBytes, int NewLibraries, int FailedLibraries, bool FirstSync,
    IReadOnlyList<string> LocallyChangedSample)
{
    public int LocallyChanged => Conflicts + Reverts + Unknown;
    public bool HasWork => NewFiles + Updates > 0;
}
public sealed record CheckOutcome(ScanResult Scan, IReadOnlyList<LibraryInfo> Libraries, IReadOnlyList<LibraryInfo> NewLibraries, CheckSummary Summary, IReadOnlySet<string> LocallyChangedKeys);
public sealed record WorkItem(DiffEntry Entry, bool Replace);
public sealed record SyncPlan(IReadOnlyList<WorkItem> Items, int Reused, long NetworkBytes, long BackupBytes)
{
    public int Missing => Items.Count(x => !x.Replace);
    public int Replacements => Items.Count(x => x.Replace);
}
public sealed record OperationProgress(string Code, int Completed = 0, int Total = 0, string? Path = null, long Bytes = 0, long TotalBytes = 0);
public sealed record TransferResult(string Path, string Status, string? BackupPath, string? OldSha256,
    string? NewSha256, string? SharePointHash, long BytesDownloaded, string? Error = null, int Attempts = 0, int Stalls = 0,
    LocalFile? InstalledLocal = null, RemoteFile? InstalledRemote = null, string? PreservedPath = null);
public sealed record ApplyResult(string RunDirectory, IReadOnlyList<TransferResult> Results, bool Cancelled, string? StopReason = null, PruneResult? Pruned = null, IoStatistics? Statistics = null);
/// <summary>Stored as runs/&lt;run&gt;/backup-meta/&lt;StableId(path)&gt;.json next to every verified backup.</summary>
/// <param name="InstalledMetadata">The manifest record when the replaced bytes were a verified SharePoint version, else null.</param>
/// <param name="Sequence">Monotonic backup order (newest highest).</param>
public sealed record BackupRecord(string? RelativePath, string? BackupPath, FileHashes? Hashes, string? Source, DateTime? OriginalCreationUtc,
    DateTime? OriginalModifiedUtc, InstalledRecord? InstalledMetadata, DateTimeOffset? BackedUpAt, long Sequence);
/// <summary>A tombstone for something retention deleted (or kept/failed, with a reason).</summary>
public sealed record PrunedItem(string Kind, string Path, string? RelativePath, string? Sha256, long Bytes, string? RunId, long? Sequence,
    string? DriveId, string? ItemId, string? ETag, string? WebUrl, DateTimeOffset? Modified, DateTimeOffset? BackedUpAt, string? Reason = null);
public sealed record PruneResult(IReadOnlyList<PrunedItem> Deleted, IReadOnlyList<PrunedItem> Kept, IReadOnlyList<PrunedItem> Failed, string? CacheSkippedReason);
/// <param name="Attempts">Content attempts used (1 when nothing had to be resumed).</param><param name="Stalls">Attempts ended by a stall.</param>
public sealed record DownloadResult(long NetworkBytes, int Attempts = 1, int Stalls = 0, FileHashes? Hashes = null);

public interface IRemoteReader
{
    Task<IReadOnlyList<LibraryInfo>> GetLibrariesAsync(string siteUrl, IReadOnlyDictionary<string, string> mappings, CancellationToken ct);
    Task<RemoteInventory> GetInventoryAsync(string siteUrl, IReadOnlyList<LibraryInfo> libraries, IProgress<OperationProgress>? progress, CancellationToken ct);
    Task<RemoteInventory> GetInventoryAsync(string siteUrl, IReadOnlyList<LibraryInfo> libraries, IProgress<OperationProgress>? progress, CancellationToken ct, InventoryReadOptions options) =>
        GetInventoryAsync(siteUrl, libraries, progress, ct);
    Task<RemoteFile> RefreshAsync(RemoteFile file, CancellationToken ct);
    Task<RemoteFile> ReadContentStateAsync(RemoteFile file, CancellationToken ct) => RefreshAsync(file, ct);
    Task PrepareApplyAsync(IReadOnlyList<RemoteFile> files, CancellationToken ct) => Task.CompletedTask;
    Task<DownloadResult> DownloadAsync(RemoteFile file, string partialPath, IProgress<OperationProgress>? progress, CancellationToken ct);
}

public sealed class OperationException(string code, string? path = null, Exception? inner = null) : Exception(code, inner)
{
    public string Code { get; } = code;
    public string? ItemPath { get; } = path;
}

public static class MetadataFields
{
    public static bool TryGet(JsonElement data, JsonElement column, out JsonElement value)
    {
        string name = column.GetProperty("InternalName").GetString()!;
        if (column.TryGetProperty("EntityPropertyName", out var property) && property.GetString() is { Length: > 0 } entity && TryGet(data, entity, out value)) return true;
        return TryGet(data, name, out value);
    }
    public static bool TryGet(JsonElement data, string name, out JsonElement value)
    {
        string encoded = name.Replace("_", "_x005f_");
        foreach (string candidate in new[] { name, "OData_" + name, name + "Id", "OData_" + name + "Id", encoded, "OData_" + encoded })
            if (data.TryGetProperty(candidate, out value)) return true;
        value = default;
        return false;
    }
}
