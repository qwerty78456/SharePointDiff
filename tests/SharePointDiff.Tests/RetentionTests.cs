using System.Diagnostics;
using System.Text.Json;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

/// <summary>Versions are produced through the real engine: install v0, then replace with v1, v2, … (each replacement backs up the previous version).</summary>
internal sealed class Mirror : IDisposable
{
    public Fixture F { get; } = new();
    public FakeRemote Remote { get; } = new();
    public Workspace W => new(Fixture.Site, F.Root);
    public string Runs => Path.Combine(W.DataRoot, "runs");
    public static byte[] V(int k, int size = 3) => Enumerable.Repeat((byte)(k + 1), size).ToArray();
    public async Task<ApplyResult> InstallAsync(string path, byte[] bytes, string? hash = "auto", string etag = "etag-1")
    {
        var file = Fixture.Remote(path, bytes, hash, etag: etag); Remote.Add(file, bytes);
        var entry = new DiffEntry { RelativePath = path, Status = DiffStatus.MissingLocal, Remote = file };
        var result = await new SyncEngine(Remote).ApplyAsync(Fixture.Scan(F.Root, entry), Fixture.Plan(new WorkItem(entry, false)), null, default);
        Assert.All(result.Results, r => Assert.StartsWith(hash == null ? "DownloadedUnverified" : "Verified", r.Status));
        return result;
    }
    /// <summary>Replaces the local file at <paramref name="localPath"/> (default: same as the remote path) with new remote bytes.</summary>
    public async Task<ApplyResult> UpdateAsync(string path, byte[] bytes, string? localPath = null, string? hash = "auto", string etag = "etag-1", CancellationToken ct = default, bool expectVerified = true)
    {
        string local = localPath ?? path, full = SafePaths.Resolve(F.Root, local);
        var hashes = await FileHashing.ComputeAsync(full);
        var file = Fixture.Remote(path, bytes, hash, etag: etag); Remote.Add(file, bytes);
        var entry = new DiffEntry { RelativePath = path, Status = hashes.Size != bytes.Length ? DiffStatus.DifferentSize : DiffStatus.DifferentHash, Local = new(local, full, hashes, File.GetLastWriteTimeUtc(full)), Remote = file };
        var result = await new SyncEngine(Remote).ApplyAsync(Fixture.Scan(F.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, ct);
        if (expectVerified) Assert.All(result.Results, r => Assert.StartsWith(hash == null ? "DownloadedUnverified" : "Verified", r.Status));
        return result;
    }
    public async Task EditLocalAsync(string path, byte[] bytes) => await File.WriteAllBytesAsync(SafePaths.Resolve(F.Root, path), bytes);
    /// <summary>Byte contents of every backup of <paramref name="path"/> still on disk (any run).</summary>
    public string[] Backups(string path) => !Directory.Exists(Runs) ? [] : Directory.GetDirectories(Runs)
        .Select(run => Path.Combine(run, "backups", path.Replace('/', '\\'))).Where(File.Exists).Select(p => Convert.ToHexString(File.ReadAllBytes(p))).Order().ToArray();
    public static string Hex(byte[] b) => Convert.ToHexString(b);
    public string[] PrunedSha() => !Directory.Exists(Runs) ? [] : Directory.GetDirectories(Runs).Select(r => Path.Combine(r, "pruned.json")).Where(File.Exists)
        .SelectMany(p => JsonDocument.Parse(File.ReadAllText(p)).RootElement.GetProperty("Deleted").EnumerateArray()).Where(e => e.GetProperty("Kind").GetString() == "backup")
        .Select(e => e.GetProperty("Sha256").GetString()!).Order().ToArray();
    public static string Sha(byte[] b) => Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(b));
    /// <summary>The backup-meta record file for the backup of <paramref name="path"/> made in <paramref name="run"/>.</summary>
    public static string Record(ApplyResult run, string path) => Path.Combine(run.RunDirectory, "backup-meta", JsonFiles.StableId(path) + ".json");
    public static string BackupFile(ApplyResult run) => run.Results.Single().BackupPath!;
    public void Dispose() => F.Dispose();
}

public sealed class RetentionTests
{
    private const string A = "Docs/a.bin";
    [Fact]
    public async Task KeepsThreeNewestDistinctPristineVersions()
    {
        using var m = new Mirror(); var mine = await m.F.LocalAsync("Docs/mine.txt", [7, 7]);
        await m.InstallAsync(A, Mirror.V(0));
        var runs = new List<ApplyResult>();
        for (int k = 1; k <= 6; k++) runs.Add(await m.UpdateAsync(A, Mirror.V(k)));
        Assert.Equal(new[] { 3, 4, 5 }.Select(k => Mirror.Hex(Mirror.V(k))).Order(), m.Backups(A));
        Assert.Equal(new[] { 0, 1, 2 }.Select(k => Mirror.Sha(Mirror.V(k))).Order(), m.PrunedSha());
        Assert.All(runs.Where(r => File.Exists(Path.Combine(r.RunDirectory, "pruned.json"))), r => Assert.True(File.Exists(Path.Combine(r.RunDirectory, "prune-plan.json"))));
        Assert.Equal(3, runs.Count(r => r.Pruned!.Deleted.Any(d => d.Kind == "backup")));
        Assert.Equal(Mirror.V(6), await File.ReadAllBytesAsync(SafePaths.Resolve(m.F.Root, A)));
        Assert.Equal(new byte[] { 7, 7 }, await File.ReadAllBytesAsync(mine.FullPath));
        var tomb = runs.SelectMany(r => r.Pruned!.Deleted).First();
        Assert.Equal((A, "drive"), (tomb.RelativePath, tomb.DriveId)); Assert.NotNull(tomb.Sequence); Assert.NotNull(tomb.ETag);
    }
    [Fact]
    public async Task DuplicatesDoNotEvictDistinctVersions()
    {
        using var m = new Mirror(); byte[] a = Mirror.V(0), b = Mirror.V(1);
        await m.InstallAsync(A, a);
        foreach (var next in new[] { b, a, b, a }) await m.UpdateAsync(A, next);
        Assert.Equal(new[] { Mirror.Hex(a), Mirror.Hex(b) }.Order(), m.Backups(A));
    }
    [Fact]
    public async Task NeverPrunesLocalEditOrUnknownBackups()
    {
        using var m = new Mirror(); await m.InstallAsync(A, Mirror.V(0));
        for (int k = 1; k <= 5; k++) { await m.EditLocalAsync(A, [9, 9, (byte)k]); await m.UpdateAsync(A, Mirror.V(k)); }
        Assert.Equal(5, m.Backups(A).Length); Assert.Empty(m.PrunedSha());
        await m.F.LocalAsync("Docs/never-synced.bin", [5, 5, 5]); // unknown origin: no manifest record
        for (int k = 1; k <= 4; k++) await m.UpdateAsync("Docs/never-synced.bin", Mirror.V(10 + k));
        Assert.Contains("050505", m.Backups("Docs/never-synced.bin"));
    }
    [Fact]
    public async Task OnlyPristineCountTowardsLimit()
    {
        using var m = new Mirror(); await m.InstallAsync(A, Mirror.V(0));
        await m.UpdateAsync(A, Mirror.V(1)); await m.UpdateAsync(A, Mirror.V(2));            // backups v0, v1
        await m.EditLocalAsync(A, [8, 8, 8]); await m.UpdateAsync(A, Mirror.V(3));            // backup of a local edit
        await m.UpdateAsync(A, Mirror.V(4)); await m.UpdateAsync(A, Mirror.V(5));            // backups v3, v4
        Assert.Equal(new[] { Mirror.Hex([8, 8, 8]), Mirror.Hex(Mirror.V(1)), Mirror.Hex(Mirror.V(3)), Mirror.Hex(Mirror.V(4)) }.Order(), m.Backups(A));
        Assert.Equal([Mirror.Sha(Mirror.V(0))], m.PrunedSha());
    }
    [Fact]
    public async Task CurrentRunNeverPruned()
    {
        using var m = new Mirror(); await m.InstallAsync(A, Mirror.V(0));
        var oldest = await m.UpdateAsync(A, Mirror.V(1)); var middle = await m.UpdateAsync(A, Mirror.V(2)); await m.UpdateAsync(A, Mirror.V(3));
        var pruned = await Retention.PruneAsync(m.W, oldest.RunDirectory, Fixture.Scan(m.F.Root), keep: 1);
        Assert.True(File.Exists(Mirror.BackupFile(oldest))); Assert.False(File.Exists(Mirror.BackupFile(middle)));
        Assert.Equal([Mirror.BackupFile(middle)], pruned.Deleted.Select(d => d.Path));
    }
    [Fact]
    public async Task UnreadableIncompleteOrTamperedMetadataIsKept()
    {
        using var m = new Mirror();
        var cases = new (string Path, Func<ApplyResult, Task> Tamper)[]
        {
            ("Docs/empty.bin", r => File.WriteAllTextAsync(Mirror.Record(r, "Docs/empty.bin"), "{}")),
            ("Docs/array.bin", r => File.WriteAllTextAsync(Mirror.Record(r, "Docs/array.bin"), "[]")),
            ("Docs/broken.bin", r => File.WriteAllTextAsync(Mirror.Record(r, "Docs/broken.bin"), "{ not json")),
            ("Docs/nohashes.bin", r => File.WriteAllTextAsync(Mirror.Record(r, "Docs/nohashes.bin"), "{\"RelativePath\":\"Docs/nohashes.bin\",\"Sequence\":1}")),
            ("Docs/tampered.bin", r => File.WriteAllBytesAsync(Mirror.BackupFile(r), [0x55, 0x55, 0x55])),
        };
        foreach (var (path, tamper) in cases)
        {
            await m.InstallAsync(path, Mirror.V(0));
            var oldest = await m.UpdateAsync(path, Mirror.V(1)); await m.UpdateAsync(path, Mirror.V(2)); await m.UpdateAsync(path, Mirror.V(3));
            await tamper(oldest);
            var last = await m.UpdateAsync(path, Mirror.V(4)); // four candidates: the oldest would be deleted if it were trusted
            Assert.True(File.Exists(Mirror.BackupFile(oldest)), path);
            if (path == "Docs/tampered.bin") Assert.Contains(last.Pruned!.Kept, k => k.Path == Mirror.BackupFile(oldest) && k.Reason == "bytes differ from the record");
        }
        Assert.Empty(m.PrunedSha());
    }
    [Fact]
    public async Task InterruptedReplaceAndUnfinishedRunAreKept()
    {
        using var m = new Mirror();
        foreach (string path in new[] { "Docs/interrupted.bin", "Docs/unfinished.bin" })
        {
            await m.InstallAsync(path, Mirror.V(0));
            var oldest = await m.UpdateAsync(path, Mirror.V(1)); await m.UpdateAsync(path, Mirror.V(2)); await m.UpdateAsync(path, Mirror.V(3));
            if (path == "Docs/interrupted.bin") File.Delete(Path.Combine(oldest.RunDirectory, "committed", JsonFiles.StableId(path) + ".json"));
            else File.Delete(Path.Combine(oldest.RunDirectory, "results.json"));
            await m.UpdateAsync(path, Mirror.V(4));
            Assert.True(File.Exists(Mirror.BackupFile(oldest)), path);
        }
    }
    [Fact]
    public async Task LegacyAdjacentSidecarsUntouched()
    {
        using var m = new Mirror(); await m.InstallAsync(A, Mirror.V(0));
        string legacy = Path.Combine(m.Runs, "20200101_000000_legacy"); string backup = Path.Combine(legacy, "backups", "Docs", "a.bin");
        Directory.CreateDirectory(Path.GetDirectoryName(backup)!); await File.WriteAllBytesAsync(backup, Mirror.V(0));
        await File.WriteAllTextAsync(backup + ".metadata.json", "{\"RelativePath\":\"Docs/a.bin\"}"); await File.WriteAllTextAsync(Path.Combine(legacy, "results.json"), "{}");
        for (int k = 1; k <= 5; k++) await m.UpdateAsync(A, Mirror.V(k));
        Assert.True(File.Exists(backup)); Assert.True(File.Exists(backup + ".metadata.json"));
    }
    [Fact]
    public async Task DownloadedUnverifiedNeverPruned()
    {
        using var m = new Mirror(); await m.InstallAsync(A, Mirror.V(0, 2), hash: null, etag: "e0");
        for (int k = 1; k <= 4; k++) await m.UpdateAsync(A, Mirror.V(k, 2 + k), hash: null, etag: "e" + k);
        Assert.Equal(4, m.Backups(A).Length); Assert.Empty(m.PrunedSha());
    }
    [Fact]
    public async Task PruneFailureReportedSyncStillSucceeds()
    {
        using var m = new Mirror(); await m.InstallAsync(A, Mirror.V(0));
        var oldest = await m.UpdateAsync(A, Mirror.V(1)); await m.UpdateAsync(A, Mirror.V(2)); await m.UpdateAsync(A, Mirror.V(3));
        ApplyResult last;
        using (new FileStream(Mirror.BackupFile(oldest), FileMode.Open, FileAccess.Read, FileShare.Read)) last = await m.UpdateAsync(A, Mirror.V(4));
        Assert.Equal("Verified", Assert.Single(last.Results).Status);
        Assert.Contains(last.Pruned!.Failed, x => x.Path == Mirror.BackupFile(oldest));
        Assert.True(File.Exists(Mirror.BackupFile(oldest)));
    }
    [Fact]
    public async Task RunEnumerationErrorsBecomeFailedItems()
    {
        using var m = new Mirror(); await m.InstallAsync(A, Mirror.V(0));
        var first = await m.UpdateAsync(A, Mirror.V(1));
        ApplyResult next;
        using (new FileStream(Mirror.Record(first, A), FileMode.Open, FileAccess.Read, FileShare.None)) next = await m.UpdateAsync(A, Mirror.V(2));
        Assert.Equal("Verified", Assert.Single(next.Results).Status);
        Assert.Contains(next.Pruned!.Failed, x => x.Kind == "run" && x.Path == first.RunDirectory);
    }
    [Fact]
    public async Task JunctionUnderRunsNotFollowed()
    {
        using var m = new Mirror(); await m.InstallAsync(A, Mirror.V(0));
        string outside = Path.Combine(Path.GetDirectoryName(m.F.Root)!, "outside-run"); string meta = Path.Combine(outside, "backup-meta"); Directory.CreateDirectory(meta);
        string victim = Path.Combine(outside, "backups", "Docs", "a.bin"); Directory.CreateDirectory(Path.GetDirectoryName(victim)!); await File.WriteAllBytesAsync(victim, Mirror.V(0));
        await File.WriteAllTextAsync(Path.Combine(outside, "results.json"), "{}");
        var hashes = await FileHashing.ComputeAsync(victim); // a record that would be prunable if the junction were followed
        await JsonFiles.WriteAsync(Path.Combine(meta, JsonFiles.StableId(A) + ".json"), new BackupRecord(A, victim, hashes, null, null, null,
            new InstalledRecord(A, hashes.Sha256, hashes.QuickXorHash, "e", DateTimeOffset.Now, Fixture.Remote(A, Mirror.V(0))), DateTimeOffset.Now, 1));
        Directory.CreateDirectory(m.Runs);
        var link = Process.Start(new ProcessStartInfo("cmd.exe", $"/c mklink /J \"{Path.Combine(m.Runs, "00000000_000000_link")}\" \"{outside}\"") { CreateNoWindow = true, UseShellExecute = false })!;
        await link.WaitForExitAsync(); Assert.Equal(0, link.ExitCode);
        for (int k = 1; k <= 5; k++) await m.UpdateAsync(A, Mirror.V(k));
        Assert.True(File.Exists(victim));
        Directory.Delete(Path.Combine(m.Runs, "00000000_000000_link")); // removes the junction only
    }
    [Fact]
    public async Task LongPathBackupsArePruned()
    {
        using var m = new Mirror();
        string path = "Docs/" + string.Join("/", Enumerable.Range(0, 4).Select(i => new string((char)('a' + i), 60))) + "/report.bin";
        await m.InstallAsync(path, Mirror.V(0));
        var oldest = await m.UpdateAsync(path, Mirror.V(1));
        Assert.True(Mirror.BackupFile(oldest).Length > 300, Mirror.BackupFile(oldest).Length.ToString());
        for (int k = 2; k <= 4; k++) await m.UpdateAsync(path, Mirror.V(k));
        Assert.False(File.Exists(Mirror.BackupFile(oldest))); Assert.Equal([Mirror.Sha(Mirror.V(0))], m.PrunedSha());
    }
    [Fact]
    public async Task CaseAndNfcVariantsGroupTogether()
    {
        using var m = new Mirror(); string nfc = "Docs/Tài liệu.bin";
        await m.InstallAsync(nfc, Mirror.V(0));
        string[] variants = ["Docs/TÀI LIỆU.bin", nfc.Normalize(System.Text.NormalizationForm.FormD), "docs/tài liệu.BIN", nfc];
        for (int k = 0; k < variants.Length; k++) await m.UpdateAsync(variants[k], Mirror.V(k + 1), localPath: nfc);
        Assert.Equal([Mirror.Sha(Mirror.V(0))], m.PrunedSha());
    }
    [Fact]
    public async Task OrderingUsesSequenceNotClock()
    {
        using var m = new Mirror(); await m.InstallAsync(A, Mirror.V(0));
        var runs = new List<ApplyResult>(); for (int k = 1; k <= 3; k++) runs.Add(await m.UpdateAsync(A, Mirror.V(k)));
        // Reverse the wall-clock stamps: the oldest backup now claims to be the newest.
        for (int i = 0; i < runs.Count; i++)
        {
            string record = Mirror.Record(runs[i], A); var r = JsonFiles.Read<BackupRecord>(record)!;
            await JsonFiles.WriteAsync(record, r with { BackedUpAt = DateTimeOffset.Now.AddYears(10 - i * 5) });
        }
        await m.UpdateAsync(A, Mirror.V(4));
        Assert.False(File.Exists(Mirror.BackupFile(runs[0]))); Assert.Equal([Mirror.Sha(Mirror.V(0))], m.PrunedSha());
    }
    [Fact]
    public async Task FaultCancelAndStopReasonSkipAllCleanup()
    {
        using var m = new Mirror(); await m.InstallAsync(A, Mirror.V(0));
        var oldest = await m.UpdateAsync(A, Mirror.V(1)); await m.UpdateAsync(A, Mirror.V(2)); await m.UpdateAsync(A, Mirror.V(3));
        int Committed() => Directory.GetDirectories(m.Runs).Count(r => File.Exists(Path.Combine(r, "committed", JsonFiles.StableId(A) + ".json")));
        // One run replaces A (a fourth pristine backup: pruning would now delete v0) while a second file ends the run
        // after A is committed: by user cancellation, by a sign-in stop, or by a programming fault.
        async Task<ApplyResult> RunAsync(int k, Func<CancellationTokenSource, Task> afterCommit)
        {
            using var cancel = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            string full = SafePaths.Resolve(m.F.Root, A); var hashes = await FileHashing.ComputeAsync(full);
            var fa = Fixture.Remote(A, Mirror.V(k)); m.Remote.Add(fa, Mirror.V(k));
            var fb = Fixture.Remote("Docs/b" + k + ".bin", [42]); m.Remote.Add(fb, [42]);
            int before = Committed();
            m.Remote.OnRefreshAsync = async (file, _) =>
            {
                if (file.ItemId != fb.ItemId) return;
                while (Committed() == before) await Task.Delay(10);
                await afterCommit(cancel);
            };
            var ea = new DiffEntry { RelativePath = A, Status = DiffStatus.DifferentHash, Local = new(A, full, hashes, File.GetLastWriteTimeUtc(full)), Remote = fa };
            var eb = new DiffEntry { RelativePath = fb.RelativePath, Status = DiffStatus.MissingLocal, Remote = fb };
            try { return await new SyncEngine(m.Remote).ApplyAsync(Fixture.Scan(m.F.Root, ea, eb), Fixture.Plan(new WorkItem(ea, true), new WorkItem(eb, false)), null, cancel.Token); }
            finally { m.Remote.OnRefreshAsync = null; }
        }
        var cancelled = await RunAsync(4, c => { c.Cancel(); return Task.CompletedTask; });
        Assert.True(cancelled.Cancelled); Assert.Null(cancelled.Pruned); Assert.False(File.Exists(Path.Combine(cancelled.RunDirectory, "pruned.json")));
        Assert.Equal("Verified", cancelled.Results.Single(r => r.Path == A).Status);
        var stopped = await RunAsync(5, _ => throw new OperationException("SignInRequired", "graph"));
        Assert.Equal("SignInRequired", stopped.StopReason); Assert.Null(stopped.Pruned); Assert.Equal("Verified", stopped.Results.Single(r => r.Path == A).Status);
        await Assert.ThrowsAsync<InvalidOperationException>(() => RunAsync(6, _ => throw new InvalidOperationException("deliberate programming fault")));
        Assert.True(File.Exists(Mirror.BackupFile(oldest)), "no cleanup may run after a cancelled, stopped or faulted run");
        Assert.Empty(m.PrunedSha());
        // The next complete run catches up.
        await m.UpdateAsync(A, Mirror.V(7));
        Assert.False(File.Exists(Mirror.BackupFile(oldest)));
    }
    // ---- Cache ----
    private static async Task<ApplyResult> SyncAsync(Mirror m)
    {
        var scan = await new DiffService(m.Remote).ScanAsync(Fixture.Site, m.F.Root, [Fixture.Library], null, default);
        return await new SyncEngine(m.Remote).ApplyAsync(scan, DiffService.SyncAll(scan.Entries), null, default);
    }
    private static string CacheId(Mirror m, string path, byte[] bytes) => Path.GetFileName(m.W.CachePath(Fixture.Remote(path, bytes)))[..16];
    [Fact]
    public async Task SupersededCacheDeletedInstalledCurrentInvalidForeignKept()
    {
        using var m = new Mirror(); m.Remote.Add(Fixture.Remote(A, Mirror.V(0)), Mirror.V(0)); await SyncAsync(m);
        string cache = m.W.CacheRoot, old = CacheId(m, A, Mirror.V(0));
        Assert.True(File.Exists(Path.Combine(cache, old + ".bin")));
        await File.WriteAllTextAsync(Path.Combine(cache, "notes.txt"), "mine");
        await File.WriteAllBytesAsync(Path.Combine(cache, "ffffffffffffffff.bin.partial.invalid_x"), [1]);
        await File.WriteAllBytesAsync(Path.Combine(cache, "eeeeeeeeeeeeeeee.bin.partial"), [1]);
        m.Remote.Add(Fixture.Remote(A, Mirror.V(1)), Mirror.V(1));
        var result = await SyncAsync(m);
        Assert.Equal("Verified", Assert.Single(result.Results).Status); Assert.Null(result.Pruned!.CacheSkippedReason);
        string current = CacheId(m, A, Mirror.V(1));
        Assert.True(File.Exists(Path.Combine(cache, current + ".bin"))); Assert.True(File.Exists(Path.Combine(cache, current + ".bin.verified.json")));
        Assert.False(File.Exists(Path.Combine(cache, old + ".bin"))); Assert.False(File.Exists(Path.Combine(cache, old + ".bin.verified.json")));
        Assert.True(File.Exists(Path.Combine(cache, "notes.txt"))); Assert.True(File.Exists(Path.Combine(cache, "ffffffffffffffff.bin.partial.invalid_x")));
        Assert.False(File.Exists(Path.Combine(cache, "eeeeeeeeeeeeeeee.bin.partial")));
        Assert.Contains(result.Pruned.Deleted, d => d.Kind == "cache" && d.Path.EndsWith(old + ".bin"));
    }
    [Fact]
    public async Task CacheOfFailedOrExcludedLibraryLegacyOrNoManifestKept()
    {
        using var m = new Mirror(); m.Remote.Add(Fixture.Remote(A, Mirror.V(0)), Mirror.V(0)); await SyncAsync(m);
        string cache = m.W.CacheRoot, old = CacheId(m, A, Mirror.V(0));
        await File.WriteAllBytesAsync(Path.Combine(cache, "aaaaaaaaaaaaaaaa.bin"), [1]); await File.WriteAllTextAsync(Path.Combine(cache, "aaaaaaaaaaaaaaaa.bin.verified.json"), "{\"Sha256\":\"x\",\"ETag\":\"e\"}");
        await File.WriteAllBytesAsync(Path.Combine(cache, "bbbbbbbbbbbbbbbb.bin"), [1]); await File.WriteAllTextAsync(Path.Combine(cache, "bbbbbbbbbbbbbbbb.bin.verified.json"), "{\"Sha256\":\"x\",\"ETag\":\"e\",\"LibraryId\":\"other\"}");
        // The library could not be read completely: its superseded entry is kept too.
        m.Remote.IssuesByLibrary["Docs"] = new("CountChanged", "Docs");
        m.Remote.Add(Fixture.Remote(A, Mirror.V(1)), Mirror.V(1));
        var scan = await new DiffService(m.Remote).ScanAsync(Fixture.Site, m.F.Root, [Fixture.Library], null, default);
        await new SyncEngine(m.Remote).ApplyAsync(scan, DiffService.SyncAll(scan.Entries), null, default);
        foreach (string id in new[] { old, "aaaaaaaaaaaaaaaa", "bbbbbbbbbbbbbbbb" }) Assert.True(File.Exists(Path.Combine(cache, id + ".bin")), id);
        using var fresh = new Mirror(); var file = Fixture.Remote(A, [1, 2]); fresh.Remote.Add(file, [1, 2]); fresh.Remote.Corrupt = true;
        var entry = new DiffEntry { RelativePath = A, Status = DiffStatus.MissingLocal, Remote = file };
        var failed = await new SyncEngine(fresh.Remote).ApplyAsync(Fixture.Scan(fresh.F.Root, entry) with { Libraries = [Fixture.Library] }, Fixture.Plan(new WorkItem(entry, false)), null, default);
        Assert.Equal("no installed manifest yet", failed.Pruned!.CacheSkippedReason);
    }
    [Fact]
    public async Task SharedContentCacheKeptWhileReferenced()
    {
        using var m = new Mirror(); m.Remote.Add(Fixture.Remote("Docs/one.bin", Mirror.V(0)), Mirror.V(0)); m.Remote.Add(Fixture.Remote("Docs/two.bin", Mirror.V(0)), Mirror.V(0));
        await SyncAsync(m);
        string shared = CacheId(m, "Docs/one.bin", Mirror.V(0));
        m.Remote.Add(Fixture.Remote("Docs/one.bin", Mirror.V(1)), Mirror.V(1));
        await SyncAsync(m);
        Assert.True(File.Exists(Path.Combine(m.W.CacheRoot, shared + ".bin")), "still the installed version of two.bin");
    }
    [Fact]
    public async Task FileAndItsMetadataJsonNamesakeBothKeepIntactBackups()
    {
        using var f = new Fixture(); var x = await f.LocalAsync("Docs/x", [1, 1]); var namesake = await f.LocalAsync("Docs/x.metadata.json", [2, 2]);
        var remote = new FakeRemote(); var rx = Fixture.Remote("Docs/x", [3, 3]); var rn = Fixture.Remote("Docs/x.metadata.json", [4, 4]); remote.Add(rx, [3, 3]); remote.Add(rn, [4, 4]);
        var entries = new[] { new DiffEntry { RelativePath = rx.RelativePath, Status = DiffStatus.DifferentHash, Local = x, Remote = rx }, new DiffEntry { RelativePath = rn.RelativePath, Status = DiffStatus.DifferentHash, Local = namesake, Remote = rn } };
        var result = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entries), Fixture.Plan(entries.Select(e => new WorkItem(e, true)).ToArray()), null, default);
        Assert.All(result.Results, r => Assert.Equal("Verified", r.Status));
        Assert.Equal(new byte[] { 1, 1 }, await File.ReadAllBytesAsync(result.Results.Single(r => r.Path == "Docs/x").BackupPath!));
        Assert.Equal(new byte[] { 2, 2 }, await File.ReadAllBytesAsync(result.Results.Single(r => r.Path == "Docs/x.metadata.json").BackupPath!));
        Assert.Equal(2, Directory.GetFiles(Path.Combine(result.RunDirectory, "backup-meta")).Length);
    }
}
