using System.Text.Json;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class RegressionTests
{
    [Fact]
    public void RestMetadataUsesEntityPropertyNameAndEncodedDisplayAliases()
    {
        var data = JsonSerializer.SerializeToElement(new Dictionary<string, object> { { "OData__ExtendedDescription", "note" }, { "Document_x005f_x0020_x005f_Number", "D-001" } });
        var column = JsonSerializer.SerializeToElement(new { InternalName = "_ExtendedDescription", EntityPropertyName = "OData__ExtendedDescription" });
        Assert.True(MetadataFields.TryGet(data, column, out var value)); Assert.Equal("note", value.GetString());
        Assert.True(MetadataFields.TryGet(data, "Document_x0020_Number", out value)); Assert.Equal("D-001", value.GetString());
    }
    [Fact]
    public async Task UnexpectedProgrammingFailurePropagatesWithPathInsteadOfBeingSwallowed()
    {
        using var f = new Fixture(); var local = await f.LocalAsync("Docs/a", [1]); var file = Fixture.Remote("Docs/a", [2]);
        var remote = new FakeRemote { OnDownload = _ => throw new InvalidOperationException("deliberate programming fault") }; remote.Add(file, [2]);
        var entry = new DiffEntry { RelativePath = file.RelativePath, Status = DiffStatus.DifferentHash, Local = local, Remote = file };
        var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, default));
        Assert.Equal(file.RelativePath, ex.Data["Path"]); Assert.Equal("Apply", ex.Data["Operation"]);
        Assert.Equal(new byte[] { 1 }, await File.ReadAllBytesAsync(local.FullPath));
    }
    [Fact]
    public async Task ProgrammingFaultCancelsActiveTransferAndPreventsQueuedDownloads()
    {
        using var f = new Fixture();
        var files = new[] { "active", "fault", "queued" }.Select((name, index) => Fixture.Remote("Docs/" + name, [(byte)(index + 2)])).ToArray();
        var active = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        bool activeCancelled = false;
        var remote = new FakeRemote
        {
            OnDownloadAsync = async (file, ct) =>
            {
                if (file.RelativePath == "Docs/active")
                {
                    active.SetResult();
                    try { await Task.Delay(Timeout.Infinite, ct); }
                    catch (OperationCanceledException) { activeCancelled = true; throw; }
                }
                else if (file.RelativePath == "Docs/fault")
                {
                    await active.Task.WaitAsync(ct);
                    throw new InvalidOperationException("deliberate parallel programming fault");
                }
            }
        };
        for (int index = 0; index < files.Length; index++) remote.Add(files[index], [(byte)(index + 2)]);
        var entries = files.Select(file => new DiffEntry { RelativePath = file.RelativePath, Status = DiffStatus.MissingLocal, Remote = file }).ToArray();
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => new SyncEngine(remote).ApplyAsync(
            Fixture.Scan(f.Root, entries), Fixture.Plan(entries.Select(entry => new WorkItem(entry, false)).ToArray()), null, deadline.Token));
        Assert.Equal("Docs/fault", ex.Data["Path"]);
        Assert.True(activeCancelled);
        Assert.Equal(2, remote.Downloads);
        Assert.False(File.Exists(Path.Combine(f.Root, "Docs", "queued")));
    }
    [Fact]
    public async Task AUserLockedDocumentIsNotReplaced()
    {
        using var f = new Fixture(); var local = await f.LocalAsync("Docs/a", [1]); var file = Fixture.Remote("Docs/a", [2]); var remote = new FakeRemote(); remote.Add(file, [2]);
        var entry = new DiffEntry { RelativePath = file.RelativePath, Status = DiffStatus.DifferentHash, Local = local, Remote = file };
        using (var locked = new FileStream(local.FullPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            var result = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, default);
            Assert.Equal("Error", Assert.Single(result.Results).Status);
        }
        Assert.Equal(new byte[] { 1 }, await File.ReadAllBytesAsync(local.FullPath));
    }
    [Fact]
    public async Task ReplacementWritesCurrentMetadataRatherThanTheOldScanSnapshot()
    {
        using var f = new Fixture(); var local = await f.LocalAsync("Docs/a", [1]); var scanned = Fixture.Remote("Docs/a", [2]);
        var current = scanned with { ETag = "metadata-revision-2", Metadata = scanned.Metadata with { Fields = JsonSerializer.SerializeToElement(new { Revision = "C" }) } };
        var remote = new FakeRemote(); remote.Add(current, [2]);
        var entry = new DiffEntry { RelativePath = scanned.RelativePath, Status = DiffStatus.DifferentHash, Local = local, Remote = scanned };
        var result = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, default);
        Assert.Equal("Verified", Assert.Single(result.Results).Status);
        var manifest = JsonFiles.Read<Dictionary<string, InstalledRecord>>(Path.Combine(new Workspace(Fixture.Site, f.Root).DataRoot, "installed-manifest.json"))!;
        Assert.Equal("C", manifest[SafePaths.Key(scanned.RelativePath)].Remote.Metadata.Fields.GetProperty("Revision").GetString());
    }
    [Fact]
    public void DriveRootIsNotAcceptedAsMirrorRoot()
    {
        Assert.Throws<OperationException>(() => new Workspace(Fixture.Site, Path.GetPathRoot(Path.GetTempPath())!));
    }
    [Fact]
    public async Task InsufficientSpaceFailsBeforeAnyDownloadOrReplacement()
    {
        using var f = new Fixture();
        var local = await f.LocalAsync("Docs/a", [1]);
        var file = Fixture.Remote("Docs/a", [2]) with { Size = new DriveInfo(Path.GetPathRoot(f.Root)!).TotalSize };
        var remote = new FakeRemote(); remote.Add(file, [2]);
        var entry = new DiffEntry { RelativePath = file.RelativePath, Status = DiffStatus.DifferentSize, Local = local, Remote = file };
        var ex = await Assert.ThrowsAsync<OperationException>(() => new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, default));
        Assert.Equal("DiskFull", ex.Code); Assert.Equal(0, remote.Downloads);
        Assert.Equal(new byte[] { 1 }, await File.ReadAllBytesAsync(local.FullPath));
    }
    [Fact]
    public void SecondLockFailsWithWorkspaceBusyAndWorksAfterDispose()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root);
        var first = WorkspaceLock.Acquire(workspace);
        Assert.Equal("WorkspaceBusy", Assert.Throws<OperationException>(() => WorkspaceLock.Acquire(workspace)).Code);
        first.Dispose();
        using var again = WorkspaceLock.Acquire(workspace);
    }
    [Fact]
    public async Task LockedWorkspaceRejectsScanAndApplyWithoutTouchingAnything()
    {
        using var f = new Fixture(); var local = await f.LocalAsync("Docs/a", [1]); var file = Fixture.Remote("Docs/a", [2]); var remote = new FakeRemote(); remote.Add(file, [2]);
        var workspace = new Workspace(Fixture.Site, f.Root);
        using (WorkspaceLock.Acquire(workspace))
        {
            var scan = await Assert.ThrowsAsync<OperationException>(() => new DiffService(remote).ScanAsync(Fixture.Site, f.Root, [Fixture.Library], null, default));
            Assert.Equal("WorkspaceBusy", scan.Code); Assert.Equal(0, remote.InventoryCalls);
            var entry = new DiffEntry { RelativePath = file.RelativePath, Status = DiffStatus.DifferentHash, Local = local, Remote = file };
            var apply = await Assert.ThrowsAsync<OperationException>(() => new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, default));
            Assert.Equal("WorkspaceBusy", apply.Code);
        }
        Assert.False(Directory.Exists(Path.Combine(workspace.DataRoot, "runs"))); Assert.False(File.Exists(Path.Combine(workspace.DataRoot, "installed-manifest.json")));
        Assert.Equal(0, remote.Downloads); Assert.Equal(new byte[] { 1 }, await File.ReadAllBytesAsync(local.FullPath));
    }
    [Fact]
    public async Task TwoSitesOnOneMirrorShareTheLock()
    {
        using var f = new Fixture(); var other = new Workspace("https://contoso.sharepoint.com/sites/other", f.Root);
        Assert.NotEqual(new Workspace(Fixture.Site, f.Root).DataRoot, other.DataRoot);
        using var held = WorkspaceLock.Acquire(other);
        var ex = await Assert.ThrowsAsync<OperationException>(() => new DiffService(new FakeRemote()).ScanAsync(Fixture.Site, f.Root, [Fixture.Library], null, default));
        Assert.Equal("WorkspaceBusy", ex.Code);
    }
    [Fact]
    public async Task LeftoverLockFileDoesNotBlockAndLockReleasedAfterFault()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root);
        await File.WriteAllTextAsync(workspace.LockPath, "left behind by a crash");
        WorkspaceLock.Acquire(workspace).Dispose();
        var file = Fixture.Remote("Docs/a", [2]); var remote = new FakeRemote { OnDownload = _ => throw new InvalidOperationException("deliberate programming fault") }; remote.Add(file, [2]);
        var entry = new DiffEntry { RelativePath = file.RelativePath, Status = DiffStatus.MissingLocal, Remote = file };
        await Assert.ThrowsAsync<InvalidOperationException>(() => new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, false)), null, default));
        using var again = WorkspaceLock.Acquire(workspace);
    }
}
/// <summary>CommitGate is process-wide, so these run after (not alongside) the parallel test collections.</summary>
[CollectionDefinition(nameof(CommitGateCollection), DisableParallelization = true)]
public sealed class CommitGateCollection;
[Collection(nameof(CommitGateCollection))]
public sealed class CommitGateTests
{
    [Fact]
    public void WaitIdleBlocksOnlyWhileATransferIsInsideTheGate()
    {
        CommitGate.Enter();
        try { Assert.Equal(1, CommitGate.Active); Assert.False(CommitGate.WaitIdle(TimeSpan.FromMilliseconds(50))); }
        finally { CommitGate.Exit(); }
        Assert.True(CommitGate.WaitIdle(TimeSpan.Zero));
    }
    [Fact]
    public async Task CommitGateReturnsToZeroAfterSuccessAndAfterAFaultInsideTheGate()
    {
        using var f = new Fixture();
        var local = await f.LocalAsync("Docs/a", [1]); var file = Fixture.Remote("Docs/a", [2]); var remote = new FakeRemote(); remote.Add(file, [2]);
        var entry = new DiffEntry { RelativePath = file.RelativePath, Status = DiffStatus.DifferentHash, Local = local, Remote = file };
        Assert.Equal("Verified", Assert.Single((await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, default)).Results).Status);
        Assert.Equal(0, CommitGate.Active);
        // The third refresh (just before installing) returns an unrepresentable creation time, so setting the
        // timestamp after File.Move throws a programming fault from inside the gate.
        var fresh = Fixture.Remote("Docs/b", [3]); remote.Add(fresh, [3]);
        remote.OnRefresh = r => { if (remote.RefreshCalls == 3 && r.ItemId == fresh.ItemId) remote.Files[fresh.ItemId] = fresh with { Created = DateTimeOffset.MinValue }; };
        var newEntry = new DiffEntry { RelativePath = fresh.RelativePath, Status = DiffStatus.MissingLocal, Remote = fresh };
        remote.ResetCounters();
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, newEntry), Fixture.Plan(new WorkItem(newEntry, false)), null, default));
        Assert.True(File.Exists(Path.Combine(f.Root, "Docs", "b")), "the fault must have happened after the commit");
        Assert.Equal(0, CommitGate.Active); Assert.True(CommitGate.WaitIdle(TimeSpan.Zero));
    }
}
