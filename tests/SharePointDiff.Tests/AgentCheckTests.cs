using System.Security.Cryptography;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class AgentCheckTests
{
    private const string DocsId = "8F2C1A6E-3B4D-4C5E-9F60-718293A4B5C6";
    private static LibraryInfo Docs => new(DocsId, "drive", "Docs", "/sites/review/Shared Documents", "Docs", 1);
    private static LibraryInfo Other => new("other", "drive-o", "Other", "/sites/review/Other", "Other", 1);
    private static async Task SelectAsync(Fixture f, string[] known, string[] included) =>
        await LibrarySelectionStore.WriteAsync(new Workspace(Fixture.Site, f.Root), new(known, included, DateTimeOffset.Now), default);
    private static FileHashes H(string sha, string qxh = "q", long size = 1) => new(size, qxh, sha);
    private static DiffEntry Entry(string path, DiffStatus status, FileHashes? local = null, string? remoteHash = "auto") => new()
    {
        RelativePath = path, Status = status, Remote = status == DiffStatus.LocalOnly ? null : Fixture.Remote(path, [1, 2], remoteHash),
        Local = local == null ? null : new(path, path, local, DateTime.UtcNow)
    };
    private static IReadOnlyList<string> Snapshot(string root) =>
        Directory.EnumerateFileSystemEntries(root, "*", SearchOption.AllDirectories).Order(StringComparer.Ordinal).Select(p =>
        {
            var info = new FileInfo(p);
            string content = (info.Attributes & FileAttributes.Directory) != 0 ? "dir" : Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(p)));
            return $"{Path.GetRelativePath(root, p)}|{content}|{info.Attributes}|{info.CreationTimeUtc:O}|{info.LastWriteTimeUtc:O}";
        }).ToArray();

    [Fact]
    public void SyncAllPlansAllMissingAndAllDifferencesOnly()
    {
        var entries = new[]
        {
            Entry("Docs/m", DiffStatus.MissingLocal), Entry("Docs/h", DiffStatus.DifferentHash, H("a")), Entry("Docs/s", DiffStatus.DifferentSize, H("b")),
            Entry("Docs/n", DiffStatus.DifferentSize, H("c"), null), Entry("Docs/u", DiffStatus.Unverified, H("d"), null), Entry("Docs/l", DiffStatus.LocalOnly, H("e")),
            Entry("Docs/e", DiffStatus.Error), Entry("Docs/ok", DiffStatus.Match, H("f"))
        };
        var plan = DiffService.SyncAll(entries);
        Assert.Equal(["Docs/h", "Docs/m", "Docs/n", "Docs/s"], plan.Items.Select(x => x.Entry.RelativePath).Order(StringComparer.Ordinal));
        Assert.False(plan.Items.Single(x => x.Entry.RelativePath == "Docs/m").Replace);
        Assert.All(plan.Items.Where(x => x.Entry.RelativePath != "Docs/m"), x => Assert.True(x.Replace));
        var reference = DiffService.Plan(entries, new HashSet<string> { SafePaths.Key("Docs/m") }, OverwriteMode.AllDifferences);
        Assert.Equal(reference.Items.Select(x => (x.Entry.RelativePath, x.Replace)), plan.Items.Select(x => (x.Entry.RelativePath, x.Replace)));
    }
    [Fact]
    public void SummarizeSplitsPristineConflictRevertUnknownNoHash()
    {
        var entries = new List<DiffEntry>
        {
            Entry("Docs/p", DiffStatus.DifferentHash, H("P")), Entry("Docs/r", DiffStatus.DifferentHash, H("X")), Entry("Docs/c", DiffStatus.DifferentHash, H("Y")),
            Entry("Docs/u", DiffStatus.DifferentHash, H("Z")), Entry("Docs/n", DiffStatus.DifferentSize, H("N"), null), Entry("Docs/dv", DiffStatus.DifferentHash, H("W")),
            Entry("Docs/m", DiffStatus.MissingLocal), Entry("Docs/e", DiffStatus.Error), Entry("Docs/v", DiffStatus.Unverified, H("V"), null), Entry("Docs/l", DiffStatus.LocalOnly, H("L"))
        };
        InstalledRecord Record(string path, string sha, string? sharePointHash) => new(path, sha, sharePointHash, "etag", DateTimeOffset.Now, Fixture.Remote(path, [9]));
        var r = entries.Single(e => e.RelativePath == "Docs/r");
        var installed = new Dictionary<string, InstalledRecord>
        {
            [SafePaths.Key("Docs/p")] = Record("Docs/p", "P", "old"),
            [SafePaths.Key("Docs/r")] = Record("Docs/r", "OLD", r.Remote!.QuickXorHash),      // only the local copy changed
            [SafePaths.Key("Docs/c")] = Record("Docs/c", "OLD", "something-else"),            // both sides moved on
            [SafePaths.Key("Docs/n")] = Record("Docs/n", "N", null),
            [SafePaths.Key("Docs/dv")] = Record("Docs/dv", "OLD", null),                       // DownloadedUnverified provenance
        };
        var scan = Fixture.Scan("C:/unused", [.. entries]);
        var s = DiffService.Summarize(scan, installed, 2, 1, true);
        Assert.Equal((1, 6), (s.NewFiles, s.Updates));
        Assert.Equal((2, 1, 2, 1), (s.PristineUpdates, s.Reverts, s.Conflicts, s.Unknown));
        Assert.Equal(1, s.NoHash); Assert.Equal(2, s.Review); Assert.Equal(1, s.LocalOnly);
        Assert.Equal(4, s.LocallyChanged); Assert.True(s.HasWork); Assert.True(s.FirstSync); Assert.Equal((2, 1), (s.NewLibraries, s.FailedLibraries));
        Assert.Equal(["Docs/c", "Docs/dv", "Docs/r", "Docs/u"], s.LocallyChangedSample);
        Assert.Equal(Provenance.Conflict, DiffService.Classify(entries, installed)[SafePaths.Key("Docs/dv")]);
    }
    [Fact]
    public async Task CheckSignInRequiredStopsBeforeHashing()
    {
        using var f = new Fixture(); await f.LocalAsync("Docs/a", [1]); await SelectAsync(f, ["list"], ["list"]);
        var remote = new FakeRemote { LibrariesFault = new OperationException("SignInRequired", "https://graph.microsoft.com") };
        var progress = new SyncProgress();
        var ex = await Assert.ThrowsAsync<OperationException>(() => new DiffService(remote).CheckAsync(Fixture.Site, f.Root, progress, default));
        Assert.Equal("SignInRequired", ex.Code); Assert.Equal(0, remote.InventoryCalls);
        Assert.DoesNotContain(progress.Items, p => p.Code == "Hashing");
        Assert.False(Directory.Exists(Path.Combine(new Workspace(Fixture.Site, f.Root).DataRoot, "runs")));
    }
    [Fact]
    public async Task CheckUsesPersistedSelectionAndReportsNewLibraries()
    {
        using var f = new Fixture();
        var extra = new LibraryInfo("extra", "drive-x", "Extra", "/sites/review/Extra", "Extra", 1);
        var remote = new FakeRemote(); remote.Libraries.Clear(); remote.Libraries.AddRange([Docs, Other, extra]);
        foreach (string p in new[] { "Docs/a", "Other/b", "Extra/c" }) remote.Add(Fixture.Remote(p, [1]), [1]);
        await SelectAsync(f, [DocsId.ToLowerInvariant(), "OTHER"], [DocsId.ToLowerInvariant()]);
        var outcome = await new DiffService(remote).CheckAsync(Fixture.Site, f.Root, null, default);
        Assert.Equal(DocsId, Assert.Single(outcome.Libraries).Id);
        Assert.Equal("extra", Assert.Single(outcome.NewLibraries).Id);
        Assert.Equal(["Docs/a"], outcome.Scan.Entries.Select(e => e.RelativePath));
        Assert.Equal(1, outcome.Summary.NewFiles); Assert.Equal(1, outcome.Summary.NewLibraries); Assert.True(outcome.Summary.FirstSync);
    }
    [Fact]
    public async Task NewLibraryOverLocalOnlyFolderIsNotPlanned()
    {
        using var f = new Fixture(); var mine = await f.LocalAsync("Other/x", [1]);
        var remote = new FakeRemote(); remote.Libraries.Add(Other); remote.Add(Fixture.Remote("Other/x", [2]), [2]);
        await SelectAsync(f, ["list"], ["list"]);
        var outcome = await new DiffService(remote).CheckAsync(Fixture.Site, f.Root, null, default);
        Assert.Single(outcome.NewLibraries);
        Assert.DoesNotContain(DiffService.SyncAll(outcome.Scan.Entries).Items, x => x.Entry.RelativePath.StartsWith("Other/"));
        Assert.Equal(new byte[] { 1 }, await File.ReadAllBytesAsync(mine.FullPath));
    }
    [Fact]
    public async Task CheckLeavesMirrorByteIdentical()
    {
        using var f = new Fixture();
        await f.LocalAsync("Docs/a", [1]); await f.LocalAsync("Docs/b", [2]); await f.LocalAsync("Docs/sub/only-local", [3]); await f.LocalAsync("root.txt", [4]);
        var remote = new FakeRemote();
        remote.Add(Fixture.Remote("Docs/a", [1]), [1]); remote.Add(Fixture.Remote("Docs/b", [9, 9]), [9, 9]); remote.Add(Fixture.Remote("Docs/new", [5]), [5]);
        await SelectAsync(f, ["list"], ["list"]);
        var before = Snapshot(f.Root);
        var outcome = await new DiffService(remote).CheckAsync(Fixture.Site, f.Root, null, default);
        Assert.Equal(before, Snapshot(f.Root));
        Assert.Equal(DiffStatus.LocalOnly, outcome.Scan.Entries.Single(e => e.RelativePath == "Docs/sub/only-local").Status);
        Assert.Equal(0, remote.Downloads); Assert.Equal(0, remote.RefreshCalls);
        Assert.Equal((1, 1), (outcome.Summary.NewFiles, outcome.Summary.Updates));
        Assert.False(Directory.Exists(Path.Combine(new Workspace(Fixture.Site, f.Root).DataRoot, "runs")));
    }
    [Fact]
    public async Task MissingSelectionIsAgentNotReadyWithoutCreatingFolders()
    {
        using var f = new Fixture(); await f.LocalAsync("Docs/a", [1]); var remote = new FakeRemote();
        var ex = await Assert.ThrowsAsync<OperationException>(() => new DiffService(remote).CheckAsync(Fixture.Site, f.Root, null, default));
        Assert.Equal("AgentNotReady", ex.Code); Assert.False(Directory.Exists(f.Root + "_SPDiff"));
        await SelectAsync(f, ["list"], []);
        ex = await Assert.ThrowsAsync<OperationException>(() => new DiffService(remote).CheckAsync(Fixture.Site, f.Root, null, default));
        Assert.Equal("NoLibrariesSelected", ex.Code); Assert.Equal(0, remote.InventoryCalls);
    }
    [Fact]
    public async Task ExcludedAndNonLibraryRootFoldersAreNotHashed()
    {
        using var f = new Fixture();
        await f.LocalAsync("Docs/a", [1]); await f.LocalAsync("Docs/sub/x", [5]); await f.LocalAsync("Other/b", [2]); await f.LocalAsync("_backup_x/c", [3]); await f.LocalAsync("root.txt", [4]);
        var remote = new FakeRemote(); remote.Libraries.Add(Other); remote.Add(Fixture.Remote("Docs/sub/x", [5]), [5]);
        await SelectAsync(f, ["list", "other"], ["list"]);
        var progress = new SyncProgress();
        var outcome = await new DiffService(remote).CheckAsync(Fixture.Site, f.Root, progress, default);
        var hashed = progress.Items.Where(p => p.Code == "Hashing").Select(p => p.Path).ToArray();
        // Narrowing applies to root folders only: everything below a selected library is still read.
        Assert.Equal(["Docs/a", "Docs/sub/x", "root.txt"], hashed.Order(StringComparer.Ordinal));
        Assert.Equal(["Docs/a", "Docs/sub/x", "root.txt"], outcome.Scan.Entries.Select(e => e.RelativePath).Order(StringComparer.Ordinal));
        Assert.Equal(DiffStatus.Match, outcome.Scan.Entries.Single(e => e.RelativePath == "Docs/sub/x").Status);
        Assert.All(outcome.Scan.Entries.Where(e => e.RelativePath != "Docs/sub/x"), e => Assert.Equal(DiffStatus.LocalOnly, e.Status));
    }
    [Fact]
    public async Task StaleInstallTempReportedNotLocalOnlyNeverPlanned()
    {
        using var f = new Fixture(); string name = "Docs/a.pdf.spdiff-" + Guid.NewGuid().ToString("N") + ".tmp";
        var temp = await f.LocalAsync(name, [7]); var remote = new FakeRemote();
        var scan = await new DiffService(remote).ScanAsync(Fixture.Site, f.Root, [Fixture.Library], null, default);
        var entry = Assert.Single(scan.Entries);
        Assert.Equal((DiffStatus.Error, "StaleInstallTemp"), (entry.Status, entry.Issue?.Code));
        Assert.Empty(DiffService.SyncAll(scan.Entries).Items); Assert.True(File.Exists(temp.FullPath));
    }
    [Fact]
    public async Task ConnectPersistsAndHonoursLibraryMappings()
    {
        using var f = new Fixture(); var remote = new FakeRemote(); var service = new DiffService(remote);
        Assert.Equal("Docs", Assert.Single(await service.ConnectAsync(Fixture.Site, f.Root, default)).LocalFolder);
        string mappings = Path.Combine(new Workspace(Fixture.Site, f.Root).DataRoot, "library-mappings.json");
        Assert.Equal("Docs", JsonFiles.Read<Dictionary<string, string>>(mappings)!["list"]);
        await JsonFiles.WriteAsync(mappings, new Dictionary<string, string> { ["list"] = "Tài liệu" });
        Assert.Equal("Tài liệu", Assert.Single(await service.ConnectAsync(Fixture.Site, f.Root, default)).LocalFolder);
    }
    [Fact]
    public async Task SignInRequiredDuringApplyStopsRunOnce()
    {
        using var f = new Fixture();
        var files = new[] { "a", "b", "c" }.Select(n => Fixture.Remote("Docs/" + n, [1])).ToArray();
        var remote = new FakeRemote
        {
            OnRefreshAsync = async (file, ct) =>
            {
                if (file.RelativePath == "Docs/a") throw new OperationException("SignInRequired", "https://graph.microsoft.com");
                await Task.Delay(Timeout.Infinite, ct);
            }
        };
        foreach (var file in files) remote.Add(file, [1]);
        var entries = files.Select(x => new DiffEntry { RelativePath = x.RelativePath, Status = DiffStatus.MissingLocal, Remote = x }).ToArray();
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        var result = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entries), Fixture.Plan(entries.Select(e => new WorkItem(e, false)).ToArray()), null, deadline.Token);
        Assert.Equal("SignInRequired", result.StopReason); Assert.False(result.Cancelled);
        Assert.Equal("SignInRequired", result.Results.Single(r => r.Path == "Docs/a").Error);
        Assert.All(result.Results.Where(r => r.Path != "Docs/a"), r => Assert.Equal("Cancelled", r.Status));
        Assert.Single(Directory.GetFiles(Path.Combine(result.RunDirectory, "errors")));
    }
    [Fact]
    public async Task EngineTimeoutCancellationIsErrorNotCancelled()
    {
        using var f = new Fixture(); var file = Fixture.Remote("Docs/a", [1]);
        var remote = new FakeRemote { OnRefreshAsync = (_, _) => throw new OperationCanceledException("simulated timeout") }; remote.Add(file, [1]);
        var entry = new DiffEntry { RelativePath = file.RelativePath, Status = DiffStatus.MissingLocal, Remote = file };
        var result = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, false)), null, default);
        var row = Assert.Single(result.Results);
        Assert.Equal(("Error", "RequestTimeout"), (row.Status, row.Error)); Assert.False(result.Cancelled); Assert.Null(result.StopReason);
        Assert.Single(Directory.GetFiles(Path.Combine(result.RunDirectory, "errors")));
    }
    [Fact]
    public async Task ReportFalseStillRefreshesManifestMatches()
    {
        using var f = new Fixture(); await f.LocalAsync("Docs/a", [1, 2]); var remote = new FakeRemote(); remote.Add(Fixture.Remote("Docs/a", [1, 2]), [1, 2]);
        var scan = await new DiffService(remote).ScanAsync(Fixture.Site, f.Root, [Fixture.Library], null, default, report: false);
        Assert.Equal(DiffStatus.Match, Assert.Single(scan.Entries).Status);
        var workspace = new Workspace(Fixture.Site, f.Root);
        Assert.False(Directory.Exists(Path.Combine(workspace.DataRoot, "runs")));
        var record = (await ManifestStore.ReadAllAsync(workspace, default))[SafePaths.Key("Docs/a")];
        Assert.Equal("Matched", record.Source); Assert.Equal(scan.Entries[0].Local!.Hashes.Sha256, record.Sha256);
    }
    [Fact]
    public void SelectionMergeKeepsKnownAndUnlistedInclusionsAndNormalisesIds()
    {
        string guid = Guid.NewGuid().ToString("D"); string upper = guid.ToUpperInvariant();
        LibraryInfo L(string id) => new(id, "d", id, "/r/" + id, id, 0);
        var previous = new LibrarySelection([guid, "b", "c"], [guid, "c"], DateTimeOffset.MinValue);
        var merged = LibrarySelectionStore.Merge(previous, [L(upper), L("B")], [L("B")], DateTimeOffset.UnixEpoch);
        Assert.Equal(new[] { guid, "b", "c" }.Order(StringComparer.Ordinal), merged.Known);
        // The GUID library was listed and unticked: dropped. C was not listed this time: kept. B ticked: included.
        Assert.Equal(["b", "c"], merged.Included);
        Assert.Equal(DateTimeOffset.UnixEpoch, merged.ComparedAt);
        Assert.Equal(["x"], LibrarySelectionStore.Merge(null, [L("X")], [L("X")], DateTimeOffset.Now).Included);
    }
    [Fact]
    public async Task CheckReportsFailedLibrariesPreviousSyncsAndLocallyChangedKeys()
    {
        using var f = new Fixture(); var local = await f.LocalAsync("Docs/a", [1]);
        var remote = new FakeRemote(); remote.Libraries.Add(Other); remote.Add(Fixture.Remote("Docs/a", [2]), [2]);
        remote.IssuesByLibrary["Other"] = new("ReadLibraryFailed", "Other");
        await SelectAsync(f, ["list", "other"], ["list", "other"]);
        var workspace = new Workspace(Fixture.Site, f.Root);
        string run = Path.Combine(workspace.DataRoot, "runs", "20260101_000000_abcdef"); Directory.CreateDirectory(run);
        await File.WriteAllTextAsync(Path.Combine(run, "results.json"), "{}");
        await ManifestStore.SaveAsync(workspace, Fixture.Remote("Docs/a", [0]), new FileHashes(1, "old", "OLD-SHA"), default);
        var outcome = await new DiffService(remote).CheckAsync(Fixture.Site, f.Root, null, default);
        Assert.Equal(1, outcome.Summary.FailedLibraries); Assert.False(outcome.Summary.FirstSync);
        Assert.Equal([SafePaths.Key("Docs/a")], outcome.LocallyChangedKeys);
        Assert.Equal((1, 1), (outcome.Summary.Conflicts, outcome.Summary.LocallyChanged));
        Assert.Equal(outcome.Libraries.Select(l => l.Id), outcome.Scan.Libraries!.Select(l => l.Id));
        Assert.Equal(new byte[] { 1 }, await File.ReadAllBytesAsync(local.FullPath));
    }
    [Fact]
    public async Task ScanReadsAFileThatAnotherProgramHasOpenForWriting()
    {
        using var f = new Fixture(); var local = await f.LocalAsync("Docs/open.docx", [1, 2, 3]);
        var remote = new FakeRemote(); remote.Add(Fixture.Remote("Docs/open.docx", [1, 2, 3]), [1, 2, 3]);
        using (new FileStream(local.FullPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
        {
            var scan = await new DiffService(remote).ScanAsync(Fixture.Site, f.Root, [Fixture.Library], null, default);
            Assert.Equal(DiffStatus.Match, Assert.Single(scan.Entries).Status);
        }
    }
}
