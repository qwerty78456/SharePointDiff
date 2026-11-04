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
}
