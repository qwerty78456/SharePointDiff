using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Security.Cryptography;

namespace SharePointDiff.Core;

public static class JsonFiles
{
    public static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter() }
    };
    public static JsonElement Empty => JsonSerializer.SerializeToElement(new Dictionary<string, object>());
    public static JsonElement EmptyArray => JsonSerializer.SerializeToElement(Array.Empty<object>());
    public static async Task WriteAsync<T>(string path, T value, CancellationToken ct = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        string tmp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            await using (var f = new FileStream(tmp, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                await JsonSerializer.SerializeAsync(f, value, Options, ct);
                await f.FlushAsync(ct);
                f.Flush(true);
            }
            File.Move(tmp, path, true);
        }
        finally { if (File.Exists(tmp)) File.Delete(tmp); }
    }
    public static T? Read<T>(string path) => File.Exists(path) ? JsonSerializer.Deserialize<T>(File.ReadAllText(path), Options) : default;
    public static string StableId(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant()[..16];
    public static string Cell(object? value)
    {
        string s = value?.ToString() ?? "";
        // CSV is intended for Excel as well; prevent file names/field values becoming formulas.
        if (s.Length > 0 && "=+-@\t\r".Contains(s[0])) s = "'" + s;
        return "\"" + s.Replace("\"", "\"\"") + "\"";
    }
    public static async Task CsvAsync(string path, IEnumerable<IEnumerable<object?>> rows, CancellationToken ct = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        await using var f = new StreamWriter(path, false, new UTF8Encoding(true));
        foreach (var row in rows) { ct.ThrowIfCancellationRequested(); await f.WriteLineAsync(string.Join(',', row.Select(Cell))); }
    }
}

public sealed class Workspace
{
    public string LocalRoot { get; }
    public string DataRoot { get; }
    public string CacheRoot => Path.Combine(DataRoot, "cache");
    public Workspace(string siteUrl, string localRoot)
    {
        string full = Path.GetFullPath(localRoot);
        if (new DirectoryInfo(full).Parent == null) throw new OperationException("InvalidRoot");
        LocalRoot = full.TrimEnd(Path.DirectorySeparatorChar);
        var parent = Directory.GetParent(LocalRoot) ?? throw new OperationException("InvalidRoot");
        if (!Directory.Exists(LocalRoot)) throw new OperationException("InvalidRoot");
        if ((File.GetAttributes(LocalRoot) & FileAttributes.ReparsePoint) != 0) throw new OperationException("LinkBlocked", localRoot);
        DataRoot = Path.Combine(parent.FullName, Path.GetFileName(LocalRoot) + "_SPDiff", JsonFiles.StableId(siteUrl.ToLowerInvariant()));
        if (DataRoot.StartsWith(LocalRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) throw new OperationException("OutsideRoot");
        string sibling = Path.GetDirectoryName(DataRoot)!;
        foreach (string dir in new[] { sibling, DataRoot, CacheRoot })
        {
            if (Directory.Exists(dir) && (File.GetAttributes(dir) & FileAttributes.ReparsePoint) != 0) throw new OperationException("LinkBlocked", dir);
            Directory.CreateDirectory(dir);
        }
    }
    public string CachePath(RemoteFile file) => Path.Combine(CacheRoot,
        JsonFiles.StableId(file.QuickXorHash is { Length: > 0 } ? file.QuickXorHash + ":" + file.Size : file.DriveId + file.ItemId + file.ETag) + ".bin");
    public string CreateRun() { string p = Path.Combine(DataRoot, "runs", DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N")[..6]); Directory.CreateDirectory(p); return p; }
}
