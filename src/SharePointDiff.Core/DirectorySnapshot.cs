using System.IO.Enumeration;

namespace SharePointDiff.Core;

/// <summary>A scan-only view. Apply must resolve live paths again.</summary>
internal sealed class DirectorySnapshot
{
    private sealed record Node(string Path, FileAttributes Attributes);
    private readonly Dictionary<string, List<Node>> nodes = new(StringComparer.Ordinal);
    private readonly HashSet<string> unreadable = new(StringComparer.Ordinal);
    public List<string> Files { get; } = [];
    public List<OperationIssue> Issues { get; } = [];
    public DirectorySnapshot(string root, IReadOnlySet<string> selected, CancellationToken ct)
    {
        void Walk(string directory, bool top)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                IoMetrics.Add("Directory.Enumerations");
                var enumeration = new FileSystemEnumerable<Node>(directory,
                    (ref FileSystemEntry e) => new(e.ToFullPath(), e.Attributes),
                    new EnumerationOptions { AttributesToSkip = 0, IgnoreInaccessible = false, RecurseSubdirectories = false });
                foreach (var node in enumeration)
                {
                    ct.ThrowIfCancellationRequested();
                    IoMetrics.Add("Directory.Entries");
                    string rel = Path.GetRelativePath(root, node.Path).Replace('\\', '/'), key = SafePaths.Key(rel);
                    if (!nodes.TryGetValue(key, out var hits)) nodes[key] = hits = [];
                    hits.Add(node);
                    bool folder = (node.Attributes & FileAttributes.Directory) != 0;
                    if (top && folder && !selected.Contains(key)) continue;
                    if ((node.Attributes & FileAttributes.ReparsePoint) != 0) { Issues.Add(new("LinkBlocked", rel)); continue; }
                    if (folder) Walk(node.Path, false);
                    else if (IsInstallTemp(Path.GetFileName(node.Path))) Issues.Add(new("StaleInstallTemp", rel));
                    else Files.Add(node.Path);
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                string rel = Path.GetRelativePath(root, directory).Replace('\\', '/');
                unreadable.Add(rel == "." ? "" : SafePaths.Key(rel));
                Issues.Add(new("LocalReadFailed", rel, ex.ToString()));
            }
        }
        if ((File.GetAttributes(root) & FileAttributes.ReparsePoint) != 0) throw new OperationException("LinkBlocked", root);
        Walk(root, true);
    }
    public OperationIssue? Validate(string relative)
    {
        var parts = SafePaths.Parts(relative);
        string key = "";
        if (unreadable.Contains("")) return new("LocalReadFailed", relative);
        for (int i = 0; i < parts.Length; i++)
        {
            key = key.Length == 0 ? SafePaths.Key(parts[i]) : key + "/" + SafePaths.Key(parts[i]);
            if (unreadable.Contains(key)) return new("LocalReadFailed", relative);
            if (!nodes.TryGetValue(key, out var hits)) continue;
            if (hits.Count > 1) return new("PathCollision", relative);
            var node = hits[0];
            if ((node.Attributes & FileAttributes.ReparsePoint) != 0) return new("LinkBlocked", relative);
            bool folder = (node.Attributes & FileAttributes.Directory) != 0;
            if (folder != (i < parts.Length - 1)) return new("PathCollision", relative);
            if (IsInstallTemp(Path.GetFileName(node.Path))) return new("StaleInstallTemp", relative);
        }
        return null;
    }
    private static bool IsInstallTemp(string name)
    {
        int marker = name.LastIndexOf(".spdiff-", StringComparison.OrdinalIgnoreCase);
        return marker >= 0 && name.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase) &&
            Guid.TryParseExact(name.AsSpan(marker + 8, name.Length - marker - 12), "N", out _);
    }
}
