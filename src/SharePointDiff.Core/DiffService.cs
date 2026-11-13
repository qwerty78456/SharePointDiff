using System.Collections.Concurrent;

namespace SharePointDiff.Core;

public sealed partial class DiffService(IRemoteReader remote)
{
    [System.Text.RegularExpressions.GeneratedRegex(@"\.spdiff-[0-9a-f]{32}\.tmp\z", System.Text.RegularExpressions.RegexOptions.IgnoreCase)]
    private static partial System.Text.RegularExpressions.Regex InstallTemp();
    public static string NormalizeSite(string site) => SafePaths.SiteUri(site).ToString().TrimEnd('/');
    /// <param name="report">False skips the runs/&lt;ts&gt;/scan report (used by the logon check); the snapshot and manifest matches are still written.</param>
    public async Task<ScanResult> ScanAsync(string site, string root, IReadOnlyList<LibraryInfo> libraries,
        IProgress<OperationProgress>? progress, CancellationToken ct, bool report = true, ScanOptions? options = null)
    {
        site = NormalizeSite(site); root = Path.GetFullPath(root);
        using var measurements = IoMetrics.Begin("Scan");
        var workspace = new Workspace(site, root);
        using var gate = WorkspaceLock.Acquire(workspace);
        var inventory = await remote.GetInventoryAsync(site, libraries, progress, ct);
        await using var manifest = await ManifestSession.OpenAsync(workspace, ct);
        var issues = inventory.Issues.ToList();
        issues.AddRange(manifest.RecoveryIssues);
        options ??= new();
        var hashes = LocalHashCache.Load(workspace); issues.AddRange(hashes.Issues);
        var local = new ConcurrentDictionary<string, LocalFile>();
        var collisions = new ConcurrentDictionary<string, byte>();
        var hashingIssues = new ConcurrentQueue<OperationIssue>();
        var libraryFolders = libraries.Select(l => SafePaths.Key(l.LocalFolder)).ToHashSet();
        var tree = await Task.Run(() => new DirectorySnapshot(root, libraryFolders, ct), ct);
        var paths = tree.Files;
        issues.AddRange(tree.Issues);
        int done = 0;
        int parallelism = options.Background ? 1 : NativeFileInfo.ManualHashParallelism(root);
        IoMetrics.Add("Hash.Workers", parallelism);
        await Parallel.ForEachAsync(paths, new ParallelOptions { MaxDegreeOfParallelism = parallelism, CancellationToken = ct }, async (path, token) =>
        {
            token.ThrowIfCancellationRequested();
            string rel = Path.GetRelativePath(root, path).Replace('\\', '/'), k = SafePaths.Key(rel);
            try
            {
                var file = await hashes.ReadAsync(path, options.RehashAllLocal || manifest.RecoveryIssues.Any(i => SafePaths.Key(i.Path) == k), token);
                if (!local.TryAdd(k, file)) { collisions.TryAdd(k, 0); hashingIssues.Enqueue(new("PathCollision", rel)); }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex) when (ex is OperationException or IOException or UnauthorizedAccessException)
            { hashingIssues.Enqueue(new(ex is OperationException op ? op.Code : "LocalReadFailed", rel, ex.ToString())); }
            progress?.Report(new("Hashing", Interlocked.Increment(ref done), paths.Count, rel));
        });
        issues.AddRange(hashingIssues);
        var index = local.Values.Where(x => !collisions.ContainsKey(SafePaths.Key(x.RelativePath)))
            .GroupBy(x => (x.Hashes.Size, x.Hashes.QuickXorHash)).ToDictionary(x => x.Key, x => new ReuseCandidate(x.First().FullPath, x.First().Hashes.Sha256));
        foreach (var found in Retention.Records(workspace))
        {
            var record = found.Record;
            if (record?.BackupPath == null || record.Hashes == null || record.Hashes.QuickXorHash.Length == 0 || !File.Exists(record.BackupPath)) continue;
            try
            {
                SafePaths.Inside(Path.Combine(found.Run, "backups"), record.BackupPath);
                if ((File.GetAttributes(record.BackupPath) & FileAttributes.ReparsePoint) != 0) continue;
                index.TryAdd((record.Hashes.Size, record.Hashes.QuickXorHash), new(record.BackupPath, record.Hashes.Sha256));
            }
            catch (OperationException) { } // an invalid backup record is never a reuse candidate
        }
        var checkedCaches = new Dictionary<string, FileHashes>();
        var entries = new List<DiffEntry>();
        var remoteKeys = new HashSet<string>();
        foreach (var file in inventory.Files)
        {
            string k = SafePaths.Key(file.RelativePath);
            if (!remoteKeys.Add(k)) throw new OperationException("PathCollision", file.RelativePath);
            local.TryGetValue(k, out var a);
            OperationIssue? problem = null;
            try { problem = tree.Validate(file.RelativePath); }
            catch (OperationException ex) { problem = new(ex.Code, file.RelativePath); }
            if (collisions.ContainsKey(k)) problem = new("PathCollision", file.RelativePath);
            // Unreadable files/directories must never be mistaken for an absent destination.
            var readError = issues.FirstOrDefault(i => i.Code is "LocalReadFailed" or "LinkBlocked" or "FileChanged" or "StaleInstallTemp" &&
                (SafePaths.Key(i.Path) == k || k.StartsWith(SafePaths.Key(i.Path).TrimEnd('/') + "/", StringComparison.Ordinal)));
            problem ??= readError;
            if (!file.Metadata.Complete) problem ??= new("MetadataIncomplete", file.RelativePath, file.Metadata.Error);
            var status = problem != null ? DiffStatus.Error : a == null ? DiffStatus.MissingLocal :
                a.Hashes.Size != file.Size ? DiffStatus.DifferentSize :
                string.IsNullOrEmpty(file.QuickXorHash) ? DiffStatus.Unverified :
                a.Hashes.QuickXorHash == file.QuickXorHash ? DiffStatus.Match : DiffStatus.DifferentHash;
            string? reuse = null;
            string? reuseSha = null;
            if (status is DiffStatus.MissingLocal or DiffStatus.DifferentHash or DiffStatus.DifferentSize && file.QuickXorHash is { Length: > 0 })
            {
                if (index.TryGetValue((file.Size, file.QuickXorHash), out var found)) { reuse = found.Path; reuseSha = found.Sha256; }
                else
                {
                    string cache = workspace.CachePath(file);
                    if (File.Exists(cache))
                    {
                        if (!checkedCaches.TryGetValue(cache, out var h)) checkedCaches[cache] = h = await FileHashing.ComputeAsync(cache, ct);
                        if (h.Size == file.Size && h.QuickXorHash == file.QuickXorHash) { reuse = cache; reuseSha = h.Sha256; }
                    }
                }
            }
            entries.Add(new() { RelativePath = file.RelativePath, Remote = file, Local = a, Status = status, ReusePath = reuse, ReuseSha256 = reuseSha, Issue = problem });
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
        foreach (var issue in issues.Where(i => i.Code is "LocalReadFailed" or "LinkBlocked" or "FileChanged" or "StaleInstallTemp"))
            if (!entries.Any(e => SafePaths.Key(e.RelativePath) == SafePaths.Key(issue.Path)))
                entries.Add(new() { RelativePath = issue.Path, Status = DiffStatus.Error, Issue = issue });
        var result = new ScanResult(site, root, entries.OrderBy(e => e.RelativePath, StringComparer.OrdinalIgnoreCase).ToArray(), inventory.Folders, issues, DateTimeOffset.Now, libraries);
        await JsonFiles.WriteAsync(Path.Combine(workspace.DataRoot, "remote-snapshot.json"), inventory, ct);
        if (report) await ReportWriter.WriteScanAsync(workspace, result, ct);
        await manifest.ObserveMatchesAsync(result.Entries.Where(e => e.Status == DiffStatus.Match), ct);
        await manifest.FlushAsync();
        await hashes.SaveAsync(ct);
        await JsonFiles.WriteAsync(Path.Combine(workspace.DataRoot, "last-scan-io.json"), measurements.Snapshot(), ct);
        return result with { Statistics = measurements.Snapshot() };
    }
    /// <summary>Lists the site's document libraries with their stable local folder names and records the mapping.</summary>
    public async Task<IReadOnlyList<LibraryInfo>> ConnectAsync(string site, string root, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(root)) throw new OperationException("InvalidRoot");
        site = NormalizeSite(site);
        var workspace = new Workspace(site, Path.GetFullPath(root));
        string mappings = Path.Combine(workspace.DataRoot, "library-mappings.json");
        var list = await remote.GetLibrariesAsync(site, JsonFiles.Read<Dictionary<string, string>>(mappings) ?? [], ct);
        await JsonFiles.WriteAsync(mappings, list.ToDictionary(x => x.Id, x => x.LocalFolder), ct);
        return list;
    }
    /// <summary>The logon check: read-only towards the mirror. Fails before any hashing when the selection is missing or sign-in fails.</summary>
    public async Task<CheckOutcome> CheckAsync(string site, string root, IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        // A blank folder in hand-edited settings is a configuration problem, not a crash.
        if (string.IsNullOrWhiteSpace(root)) throw new OperationException("InvalidRoot");
        site = NormalizeSite(site); root = Path.GetFullPath(root);
        // Gate first, without creating any folder: only a mirror the user has compared in the main window is checked.
        var selection = LibrarySelectionStore.Read(Workspace.DataRootFor(site, root)) ?? throw new OperationException("AgentNotReady", root);
        var listed = await ConnectAsync(site, root, ct);
        var known = selection.Known.Select(LibrarySelectionStore.Id).ToHashSet();
        var included = selection.Included.Select(LibrarySelectionStore.Id).ToHashSet();
        var libraries = listed.Where(l => included.Contains(LibrarySelectionStore.Id(l.Id))).ToArray();
        var fresh = listed.Where(l => !known.Contains(LibrarySelectionStore.Id(l.Id))).ToArray();
        if (libraries.Length == 0) throw new OperationException("NoLibrariesSelected", root);
        var scan = await ScanAsync(site, root, libraries, progress, ct, report: false, options: new(Background: true));
        var workspace = new Workspace(site, root);
        var installed = await ManifestStore.ReadAllAsync(workspace, ct);
        string runs = Path.Combine(workspace.DataRoot, "runs");
        bool firstSync = !Directory.Exists(runs) || !Directory.EnumerateDirectories(runs).Any(d => File.Exists(Path.Combine(d, "results.json")));
        int failed = libraries.Count(l => !scan.RemoteFolders.Contains(l.LocalFolder));
        var provenance = Classify(scan.Entries, installed);
        var changed = provenance.Where(p => p.Value != Provenance.Pristine).Select(p => p.Key).ToHashSet();
        return new(scan, libraries, fresh, Summarize(scan, installed, fresh.Length, failed, firstSync), changed);
    }
    /// <summary>The one-click Sync: every missing file plus every difference (each replacement is backed up by the engine).</summary>
    public static SyncPlan SyncAll(IEnumerable<DiffEntry> entries)
    {
        var list = entries.ToArray();
        return Plan(list, list.Where(e => e.Status == DiffStatus.MissingLocal).Select(e => SafePaths.Key(e.RelativePath)).ToHashSet(), OverwriteMode.AllDifferences);
    }
    /// <summary>Where each difference's local bytes come from, relative to the last verified SharePoint version in the manifest.</summary>
    public static IReadOnlyDictionary<string, Provenance> Classify(IEnumerable<DiffEntry> entries, IReadOnlyDictionary<string, InstalledRecord> installed)
    {
        var result = new Dictionary<string, Provenance>();
        foreach (var e in entries.Where(e => e.IsDifference && e.Local != null && e.Remote != null))
        {
            string key = SafePaths.Key(e.RelativePath);
            result[key] = !installed.TryGetValue(key, out var record) ? Provenance.Unknown
                : record.Sha256 == e.Local!.Hashes.Sha256 ? Provenance.Pristine
                : record.SharePointHash is { Length: > 0 } h && h == e.Remote!.QuickXorHash ? Provenance.Revert
                : Provenance.Conflict;
        }
        return result;
    }
    public static CheckSummary Summarize(ScanResult scan, IReadOnlyDictionary<string, InstalledRecord> installed, int newLibraries, int failedLibraries, bool firstSync)
    {
        var plan = SyncAll(scan.Entries);
        var provenance = Classify(scan.Entries, installed);
        var replace = plan.Items.Where(x => x.Replace).ToArray(); var missing = plan.Items.Where(x => !x.Replace).ToArray();
        int Count(Provenance p) => replace.Count(x => provenance[SafePaths.Key(x.Entry.RelativePath)] == p);
        var sample = replace.Where(x => provenance[SafePaths.Key(x.Entry.RelativePath)] != Provenance.Pristine)
            .Select(x => x.Entry.RelativePath).Order(StringComparer.OrdinalIgnoreCase).Take(5).ToArray();
        return new(missing.Length, missing.Sum(x => x.Entry.Remote!.Size), replace.Length, replace.Sum(x => x.Entry.Remote!.Size),
            Count(Provenance.Pristine), Count(Provenance.Conflict), Count(Provenance.Revert), Count(Provenance.Unknown),
            replace.Count(x => string.IsNullOrEmpty(x.Entry.Remote!.QuickXorHash)),
            scan.Entries.Count(e => e.Status is DiffStatus.Error or DiffStatus.Unverified), scan.Entries.Count(e => e.Status == DiffStatus.LocalOnly),
            plan.Reused, plan.NetworkBytes, plan.BackupBytes, newLibraries, failedLibraries, firstSync, sample);
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
        var groups = work.GroupBy(x => ContentGrouping.Key(x.Entry.Remote!)).ToArray();
        return new(work, groups.Sum(g => g.Any(x => x.Entry.ReusePath != null) ? g.Count() : Math.Max(0, g.Count() - 1)),
            groups.Where(g => !g.Any(x => x.Entry.ReusePath != null)).Sum(g => g.First().Entry.Remote!.Size),
            work.Where(x => x.Replace).Sum(x => x.Entry.Local!.Hashes.Size));
    }
}
