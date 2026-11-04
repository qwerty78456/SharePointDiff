namespace SharePointDiff.Core;

public sealed class DiffService(IRemoteReader remote)
{
    public async Task<ScanResult> ScanAsync(string site, string root, IReadOnlyList<LibraryInfo> libraries,
        IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        site = SafePaths.SiteUri(site).ToString().TrimEnd('/'); root = Path.GetFullPath(root);
        var workspace = new Workspace(site, root);
        var inventory = await remote.GetInventoryAsync(site, libraries, progress, ct);
        var issues = inventory.Issues.ToList();
        var local = new Dictionary<string, LocalFile>();
        var collisions = new HashSet<string>();
        var paths = new List<string>();
        void Walk(string directory)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                foreach (string path in Directory.EnumerateFileSystemEntries(directory))
                {
                    ct.ThrowIfCancellationRequested();
                    var attr = File.GetAttributes(path);
                    string rel = Path.GetRelativePath(root, path).Replace('\\', '/');
                    if ((attr & FileAttributes.ReparsePoint) != 0) { issues.Add(new("LinkBlocked", rel)); continue; }
                    if ((attr & FileAttributes.Directory) != 0) Walk(path); else paths.Add(path);
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            { issues.Add(new("LocalReadFailed", Path.GetRelativePath(root, directory), ex.ToString())); }
        }
        await Task.Run(() => Walk(root), ct);
        int done = 0;
        foreach (string path in paths)
        {
            ct.ThrowIfCancellationRequested();
            string rel = Path.GetRelativePath(root, path).Replace('\\', '/'), k = SafePaths.Key(rel);
            try
            {
                var h = await FileHashing.ComputeAsync(path, ct);
                if (!local.TryAdd(k, new(rel, path, h, File.GetLastWriteTimeUtc(path)))) { collisions.Add(k); issues.Add(new("PathCollision", rel)); }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex) when (ex is OperationException or IOException or UnauthorizedAccessException)
            { issues.Add(new(ex is OperationException op ? op.Code : "LocalReadFailed", rel, ex.ToString())); }
            progress?.Report(new("Hashing", ++done, paths.Count, rel));
        }
        var index = local.Values.Where(x => !collisions.Contains(SafePaths.Key(x.RelativePath)))
            .GroupBy(x => (x.Hashes.Size, x.Hashes.QuickXorHash)).ToDictionary(x => x.Key, x => x.First().FullPath);
        var entries = new List<DiffEntry>();
        var remoteKeys = new HashSet<string>();
        foreach (var file in inventory.Files)
        {
            string k = SafePaths.Key(file.RelativePath);
            if (!remoteKeys.Add(k)) throw new OperationException("PathCollision", file.RelativePath);
            local.TryGetValue(k, out var a);
            OperationIssue? problem = null;
            try { SafePaths.Resolve(root, file.RelativePath); }
            catch (OperationException ex) { problem = new(ex.Code, file.RelativePath); }
            if (collisions.Contains(k)) problem = new("PathCollision", file.RelativePath);
            if (Directory.Exists(SafePaths.Inside(root, Path.Combine(root, file.RelativePath.Replace('/', Path.DirectorySeparatorChar)))))
                problem = new("PathCollision", file.RelativePath);
            // Unreadable files/directories must never be mistaken for an absent destination.
            var readError = issues.FirstOrDefault(i => i.Code is "LocalReadFailed" or "LinkBlocked" or "FileChanged" &&
                (SafePaths.Key(i.Path) == k || k.StartsWith(SafePaths.Key(i.Path).TrimEnd('/') + "/", StringComparison.Ordinal)));
            problem ??= readError;
            if (!file.Metadata.Complete) problem ??= new("MetadataIncomplete", file.RelativePath, file.Metadata.Error);
            var status = problem != null ? DiffStatus.Error : a == null ? DiffStatus.MissingLocal :
                a.Hashes.Size != file.Size ? DiffStatus.DifferentSize :
                string.IsNullOrEmpty(file.QuickXorHash) ? DiffStatus.Unverified :
                a.Hashes.QuickXorHash == file.QuickXorHash ? DiffStatus.Match : DiffStatus.DifferentHash;
            string? reuse = null;
            if (status is DiffStatus.MissingLocal or DiffStatus.DifferentHash or DiffStatus.DifferentSize && file.QuickXorHash is { Length: > 0 })
            {
                if (index.TryGetValue((file.Size, file.QuickXorHash), out var found)) reuse = found;
                else
                {
                    string cache = workspace.CachePath(file);
                    if (File.Exists(cache))
                    {
                        var h = await FileHashing.ComputeAsync(cache, ct);
                        if (h.Size == file.Size && h.QuickXorHash == file.QuickXorHash) reuse = cache;
                    }
                }
            }
            entries.Add(new() { RelativePath = file.RelativePath, Remote = file, Local = a, Status = status, ReusePath = reuse, Issue = problem });
        }
        var prefixes = libraries.Select(l => SafePaths.Key(l.LocalFolder) + "/").ToArray();
        var failed = libraries.Where(l => !inventory.Folders.Contains(l.LocalFolder)).Select(l => SafePaths.Key(l.LocalFolder) + "/").ToArray();
        foreach (var pair in local.Where(p => !remoteKeys.Contains(p.Key)))
        {
            if (!prefixes.Any(x => pair.Key.StartsWith(x, StringComparison.Ordinal)) && pair.Value.RelativePath.Contains('/')) continue;
            bool unscanned = failed.Any(x => pair.Key.StartsWith(x, StringComparison.Ordinal));
            entries.Add(new()
            {
                RelativePath = pair.Value.RelativePath,
                Local = pair.Value,
                Status = unscanned ? DiffStatus.Error : DiffStatus.LocalOnly,
                Issue = unscanned ? new("UnscannedLibrary", pair.Value.RelativePath) : null
            });
        }
        foreach (var issue in issues.Where(i => i.Code is "LocalReadFailed" or "LinkBlocked" or "FileChanged"))
            if (!entries.Any(e => SafePaths.Key(e.RelativePath) == SafePaths.Key(issue.Path)))
                entries.Add(new() { RelativePath = issue.Path, Status = DiffStatus.Error, Issue = issue });
        var result = new ScanResult(site, root, entries.OrderBy(e => e.RelativePath, StringComparer.OrdinalIgnoreCase).ToArray(), inventory.Folders, issues, DateTimeOffset.Now);
        await JsonFiles.WriteAsync(Path.Combine(workspace.DataRoot, "remote-snapshot.json"), inventory, ct);
        await ReportWriter.WriteScanAsync(workspace, result, ct);
        await ManifestStore.RefreshMatchesAsync(workspace, result.Entries.Where(e => e.Status == DiffStatus.Match), ct);
        return result;
    }
    public static SyncPlan Plan(IEnumerable<DiffEntry> entries, ISet<string> selected, OverwriteMode mode)
    {
        var work = new List<WorkItem>();
        foreach (var e in entries)
        {
            bool picked = selected.Contains(SafePaths.Key(e.RelativePath));
            if (e.Status == DiffStatus.MissingLocal && picked) work.Add(new(e, false));
            else if (e.IsDifference && (mode == OverwriteMode.AllDifferences || (mode == OverwriteMode.Selected && picked))) work.Add(new(e, true));
        }
        return new(work, work.Count(x => x.Entry.ReusePath != null), work.Where(x => x.Entry.ReusePath == null).Sum(x => x.Entry.Remote!.Size),
            work.Where(x => x.Replace).Sum(x => x.Entry.Local!.Hashes.Size));
    }
}
