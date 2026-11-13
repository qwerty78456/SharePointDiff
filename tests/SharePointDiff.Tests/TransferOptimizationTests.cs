using System.Security.Cryptography;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class TransferOptimizationTests
{
    [Fact]
    public async Task FreshReplacementReadsSixPayloadLengthsAndStillWritesThreeVerifiedCopies()
    {
        using var f = new Fixture(); byte[] old = new byte[65537], current = new byte[65537];
        new Random(1).NextBytes(old); new Random(2).NextBytes(current);
        await f.LocalAsync("Docs/a", old); var file = Fixture.Remote("Docs/a", current); var remote = new FakeRemote(); remote.Add(file, current);
        var scan = await new DiffService(remote).ScanAsync(Fixture.Site, f.Root, [Fixture.Library], null, default, report: false);
        var row = scan.Entries.Single();
        var result = await new SyncEngine(remote).ApplyAsync(scan, Fixture.Plan(new WorkItem(row, true)), null, default);
        Assert.Equal("Verified", result.Results.Single().Status);
        Assert.Equal(6L * current.Length, result.Statistics!.Counters["Content.ReadBytes"]);
        Assert.Equal(3L * current.Length, result.Statistics.Counters["Content.WriteBytes"]);
        Assert.Equal(current.Length, result.Statistics.Counters["Network.FilePayloadBytes"]);
        Assert.Equal(old.Length, scan.Statistics!.Counters["Content.ReadBytes"]);
    }
    [Fact]
    public async Task AReuseCandidateInAnySelectedRowFeedsTheWholeContentGroupWithoutNetwork()
    {
        using var f = new Fixture(); var source = await f.LocalAsync("source", [1, 2, 3]); var remote = new FakeRemote();
        var a = Fixture.Remote("Docs/a", [1, 2, 3]); var b = Fixture.Remote("Docs/b", [1, 2, 3]);
        remote.Add(a, [1, 2, 3]); remote.Add(b, [1, 2, 3]);
        var entries = new[]
        {
            new DiffEntry { RelativePath = a.RelativePath, Remote = a, Status = DiffStatus.MissingLocal },
            new DiffEntry { RelativePath = b.RelativePath, Remote = b, Status = DiffStatus.MissingLocal, ReusePath = source.FullPath }
        };
        var plan = DiffService.Plan(entries, entries.Select(e => SafePaths.Key(e.RelativePath)).ToHashSet(), OverwriteMode.KeepLocal);
        Assert.Equal(0, plan.NetworkBytes); Assert.Equal(2, plan.Reused);
        var result = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entries), plan, null, default);
        Assert.All(result.Results, r => { Assert.Equal("Verified", r.Status); Assert.NotNull(r.InstalledLocal); Assert.NotNull(r.InstalledRemote); });
        Assert.Equal(0, remote.Downloads); Assert.True(File.Exists(source.FullPath));
    }
    [Fact]
    public void EstimatedBytesDeduplicateKnownHashesButNeverUnknownHashesAcrossDifferentItems()
    {
        var a = Fixture.Remote("Docs/a", [1]); var b = Fixture.Remote("Docs/b", [1]);
        SyncPlan Plan(params RemoteFile[] files)
        {
            var entries = files.Select(file => new DiffEntry { RelativePath = file.RelativePath, Remote = file, Status = DiffStatus.MissingLocal }).ToArray();
            return DiffService.Plan(entries, entries.Select(e => SafePaths.Key(e.RelativePath)).ToHashSet(), OverwriteMode.KeepLocal);
        }
        Assert.Equal(1, Plan(a, b).NetworkBytes);
        Assert.Equal(2, Plan(a with { QuickXorHash = null }, b with { QuickXorHash = null }).NetworkBytes);
    }
    [Fact]
    public async Task RenamedDestinationDuringCommitIsCapturedInsteadOfBeingLost()
    {
        using var f = new Fixture(); var local = await f.LocalAsync("Docs/a", [1]); var file = Fixture.Remote("Docs/a", [2]);
        var remote = new FakeRemote(); remote.Add(file, [2]);
        var entry = new DiffEntry { RelativePath = file.RelativePath, Local = local, Remote = file, Status = DiffStatus.DifferentHash };
        string moved = local.FullPath + ".user-renamed";
        var engine = new SyncEngine(remote)
        {
            Checkpoint = (stage, path) => { if (stage == "BeforeReplace") { File.Move(path, moved); File.WriteAllBytes(path, [9]); } }
        };
        var row = Assert.Single((await engine.ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, default)).Results);
        Assert.Equal("CommitPathConflict", row.Error); Assert.NotNull(row.PreservedPath);
        Assert.Equal(new byte[] { 9 }, await File.ReadAllBytesAsync(row.PreservedPath!));
        Assert.Equal(new byte[] { 1 }, await File.ReadAllBytesAsync(row.BackupPath!));
        Assert.Equal(new byte[] { 1 }, await File.ReadAllBytesAsync(moved));
        Assert.Equal(new byte[] { 2 }, await File.ReadAllBytesAsync(local.FullPath));
    }
    [Fact]
    public async Task CopyHashesAreCheckedAgainstIndependentTargetReadback()
    {
        using var f = new Fixture(); byte[] bytes = new byte[2_100_001]; new Random(7).NextBytes(bytes);
        var local = await f.LocalAsync("source", bytes); string target = Path.Combine(f.Root, "target");
        var copied = await FileHashing.CopyVerifiedAsync(local.FullPath, target, default);
        Assert.Equal(local.Hashes, copied); Assert.Equal(Convert.ToHexString(SHA256.HashData(bytes)), copied.Sha256);
    }
}
