using System.Text;

namespace SharePointDiff.Core;

public static class SafePaths
{
    public static string Key(string path) => path.Replace('\\', '/').Normalize(NormalizationForm.FormC).ToUpperInvariant();
    public static Uri SiteUri(string value)
    {
        if (!Uri.TryCreate(value.Trim(), UriKind.Absolute, out var uri) || uri.Scheme != "https" ||
            !uri.Host.EndsWith(".sharepoint.com", StringComparison.OrdinalIgnoreCase) || uri.UserInfo.Length != 0 ||
            uri.Port != 443 || uri.Query.Length != 0 || uri.Fragment.Length != 0 || uri.AbsolutePath.Contains("/_layouts/", StringComparison.OrdinalIgnoreCase))
            throw new OperationException("InvalidSite");
        return new Uri(uri.GetLeftPart(UriPartial.Path).TrimEnd('/'));
    }
    public static string[] Parts(string relative)
    {
        var parts = relative.Replace('\\', '/').Split('/');
        foreach (string p in parts)
        {
            string stem = p.Split('.')[0].ToUpperInvariant();
            if (p.Length == 0 || p is "." or ".." || p.EndsWith(' ') || p.EndsWith('.') ||
                p.Any(c => c < 32 || "<>:\"|?*".Contains(c)) ||
                stem is "CON" or "PRN" or "AUX" or "NUL" ||
                (stem.Length == 4 && (stem.StartsWith("COM") || stem.StartsWith("LPT")) && stem[3] is >= '1' and <= '9'))
                throw new OperationException("InvalidPath", relative);
        }
        return parts;
    }
    public static string Inside(string root, string path)
    {
        string r = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        string p = Path.GetFullPath(path);
        if (!p.StartsWith(r, StringComparison.OrdinalIgnoreCase)) throw new OperationException("OutsideRoot", path);
        return p;
    }
    public static string Resolve(string root, string relative, bool createParents = false)
    {
        string current = Path.GetFullPath(root);
        if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0) throw new OperationException("LinkBlocked", root);
        var parts = Parts(relative);
        for (int i = 0; i < parts.Length; i++)
        {
            string part = parts[i];
            if (Directory.Exists(current))
            {
                var hits = Directory.EnumerateFileSystemEntries(current).Where(p => Key(Path.GetFileName(p)) == Key(part)).Take(2).ToArray();
                if (hits.Length > 1) throw new OperationException("PathCollision", relative);
                if (hits.Length == 1) part = Path.GetFileName(hits[0]);
            }
            current = Inside(root, Path.Combine(current, part));
            if ((File.Exists(current) || Directory.Exists(current)) && (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                throw new OperationException("LinkBlocked", relative);
            if (createParents && i < parts.Length - 1) Directory.CreateDirectory(current);
        }
        return current;
    }
    public static void AssertDirectoryChain(string root, string fullPath)
    {
        Inside(root, fullPath);
        Resolve(root, Path.GetRelativePath(root, fullPath));
    }
}
