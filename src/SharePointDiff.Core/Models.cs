using System.Text.Json;

namespace SharePointDiff.Core;

public enum DiffStatus { Match, MissingLocal, DifferentSize, DifferentHash, LocalOnly, Unverified, Error }
public enum OverwriteMode { KeepLocal, Selected, AllDifferences }
public sealed record ConnectionProfile(string Name, string TenantId, string ClientId);
public sealed record LibraryInfo(string Id, string DriveId, string Title, string RootUrl, string LocalFolder, int ItemCount);
public sealed record FileHashes(long Size, string QuickXorHash, string Sha256);
public sealed record LocalFile(string RelativePath, string FullPath, FileHashes Hashes, DateTime LastWriteUtc);
public sealed record MetadataSnapshot(JsonElement Fields, JsonElement DisplayValues, JsonElement Columns, bool Complete, string? Error = null);
public sealed record RemoteFile(
    string SiteUrl, string LibraryId, string DriveId, string ItemId, string ListItemId,
    string RelativePath, string Name, long Size, string? QuickXorHash, string ETag,
    DateTimeOffset Created, DateTimeOffset Modified, string WebUrl, JsonElement Basic,
    MetadataSnapshot Metadata);
public sealed record RemoteInventory(IReadOnlyList<RemoteFile> Files, IReadOnlyList<string> Folders, IReadOnlyList<OperationIssue> Issues);
public sealed record OperationIssue(string Code, string Path, string? Detail = null);
public sealed class DiffEntry
{
    public required string RelativePath { get; init; }
    public required DiffStatus Status { get; init; }
    public LocalFile? Local { get; init; }
    public RemoteFile? Remote { get; init; }
    public string? ReusePath { get; set; }
    public OperationIssue? Issue { get; init; }
    public bool IsDifference => Status is DiffStatus.DifferentHash or DiffStatus.DifferentSize;
}
public sealed record ScanResult(string SiteUrl, string LocalRoot, IReadOnlyList<DiffEntry> Entries,
    IReadOnlyList<string> RemoteFolders, IReadOnlyList<OperationIssue> Issues, DateTimeOffset ScannedAt);
public sealed record WorkItem(DiffEntry Entry, bool Replace);
public sealed record SyncPlan(IReadOnlyList<WorkItem> Items, int Reused, long NetworkBytes, long BackupBytes)
{
    public int Missing => Items.Count(x => !x.Replace);
    public int Replacements => Items.Count(x => x.Replace);
}
public sealed record OperationProgress(string Code, int Completed = 0, int Total = 0, string? Path = null, long Bytes = 0, long TotalBytes = 0);
public sealed record TransferResult(string Path, string Status, string? BackupPath, string? OldSha256,
    string? NewSha256, string? SharePointHash, long BytesDownloaded, string? Error = null);
public sealed record ApplyResult(string RunDirectory, IReadOnlyList<TransferResult> Results, bool Cancelled);
public sealed record DownloadResult(long NetworkBytes);

public interface IRemoteReader
{
    Task<IReadOnlyList<LibraryInfo>> GetLibrariesAsync(string siteUrl, IReadOnlyDictionary<string, string> mappings, CancellationToken ct);
    Task<RemoteInventory> GetInventoryAsync(string siteUrl, IReadOnlyList<LibraryInfo> libraries, IProgress<OperationProgress>? progress, CancellationToken ct);
    Task<RemoteFile> RefreshAsync(RemoteFile file, CancellationToken ct);
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
