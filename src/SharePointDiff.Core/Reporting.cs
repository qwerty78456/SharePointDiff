using System.Text.Json;

namespace SharePointDiff.Core;

/// <param name="Source">Audit only: Downloaded, Reused or Matched. Retention never depends on it.</param>
public sealed record InstalledRecord(string Path, string Sha256, string? SharePointHash, string ETag, DateTimeOffset ObservedAt, RemoteFile Remote, string? Source = null);
public static class ManifestStore
{
    private static readonly SemaphoreSlim Gate = new(1, 1);
    private static string PathOf(Workspace workspace) => Path.Combine(workspace.DataRoot, "installed-manifest.json");
    public static async Task<InstalledRecord?> ReadAsync(Workspace workspace, string relative, CancellationToken ct)
    {
        await Gate.WaitAsync(ct);
        try { return JsonFiles.Read<Dictionary<string, InstalledRecord>>(PathOf(workspace))?.GetValueOrDefault(SafePaths.Key(relative)); }
        finally { Gate.Release(); }
    }
    public static async Task<IReadOnlyDictionary<string, InstalledRecord>> ReadAllAsync(Workspace workspace, CancellationToken ct)
    {
        await Gate.WaitAsync(ct);
        try { return JsonFiles.Read<Dictionary<string, InstalledRecord>>(PathOf(workspace)) ?? []; }
        finally { Gate.Release(); }
    }
    public static async Task SaveAsync(Workspace workspace, RemoteFile remote, FileHashes local, CancellationToken ct, string? source = null)
    {
        await Gate.WaitAsync(ct);
        try
        {
            string path = PathOf(workspace);
            var records = JsonFiles.Read<Dictionary<string, InstalledRecord>>(path) ?? [];
            records[SafePaths.Key(remote.RelativePath)] = new(remote.RelativePath, local.Sha256, remote.QuickXorHash, remote.ETag, DateTimeOffset.Now, remote, source);
            await JsonFiles.WriteAsync(path, records, ct);
        }
        finally { Gate.Release(); }
    }
    public static async Task RefreshMatchesAsync(Workspace workspace, IEnumerable<DiffEntry> entries, CancellationToken ct)
    {
        await Gate.WaitAsync(ct);
        try
        {
            string path = PathOf(workspace);
            var records = JsonFiles.Read<Dictionary<string, InstalledRecord>>(path) ?? [];
            foreach (var e in entries)
            {
                string key = SafePaths.Key(e.RelativePath);
                // Keep how the bytes arrived while they stay the same; a fresh observation is "Matched".
                string source = records.TryGetValue(key, out var old) && old.Sha256 == e.Local!.Hashes.Sha256 && old.Source != null ? old.Source : "Matched";
                records[key] = new(e.RelativePath, e.Local!.Hashes.Sha256, e.Remote!.QuickXorHash, e.Remote.ETag, DateTimeOffset.Now, e.Remote, source);
            }
            await JsonFiles.WriteAsync(path, records, ct);
        }
        finally { Gate.Release(); }
    }
}
public static class ReportWriter
{
    public static async Task WriteScanAsync(Workspace workspace, ScanResult scan, CancellationToken ct)
    {
        string run = workspace.CreateRun();
        await JsonFiles.WriteAsync(Path.Combine(run, "scan.json"), scan, ct);
        var rows = new List<IEnumerable<object?>> { new object?[] { "Status", "Path", "LocalBytes", "SharePointBytes", "LocalQuickXor", "SharePointQuickXor", "ReusePath", "Issue" } };
        rows.AddRange(scan.Entries.Select(e => new object?[] { e.Status, e.RelativePath, e.Local?.Hashes.Size, e.Remote?.Size, e.Local?.Hashes.QuickXorHash, e.Remote?.QuickXorHash, e.ReusePath, e.Issue?.Code }));
        await JsonFiles.CsvAsync(Path.Combine(run, "diff.csv"), rows, ct);
        await WriteMetadataCsvAsync(Path.Combine(run, "metadata.csv"), scan.Entries.Where(e => e.Remote != null).Select(e => e.Remote!), ct);
    }
    public static Task WriteMetadataCsvAsync(string path, IEnumerable<RemoteFile> files, CancellationToken ct)
    {
        var rows = new List<IEnumerable<object?>> { new object?[] { "Path", "InternalName", "DisplayName", "Type", "RawValue", "DisplayValue", "Complete" } };
        foreach (var f in files)
        {
            foreach (var col in f.Metadata.Columns.EnumerateArray())
            {
                string n = col.GetProperty("InternalName").GetString()!;
                string raw = MetadataFields.TryGet(f.Metadata.Fields, col, out var v) ? v.GetRawText() : "";
                string display = MetadataFields.TryGet(f.Metadata.DisplayValues, col, out var d) ? d.ToString() : "";
                rows.Add(new object?[] { f.RelativePath, n, col.GetProperty("Title").ToString(), col.GetProperty("TypeAsString").ToString(), raw, display, f.Metadata.Complete });
            }
        }
        return JsonFiles.CsvAsync(path, rows, ct);
    }
}
