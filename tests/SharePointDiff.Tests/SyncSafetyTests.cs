using System.Text.Json;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class SyncSafetyTests
{
    private static DiffEntry Difference(LocalFile local, RemoteFile remote) => new() { RelativePath = remote.RelativePath, Local = local, Remote = remote, Status = DiffStatus.DifferentHash };
    [Fact]
    public async Task ReplacementBacksUpOriginalAndPreservesMetadataAndLocalOnly()
    {
        using var f = new Fixture(); var local = await f.LocalAsync("Docs/a.bin", [1, 2, 3]); var extra = await f.LocalAsync("Docs/only-local.bin", [9, 8, 7]);
        var file = Fixture.Remote("Docs/a.bin", [4, 5, 6]); var remote = new FakeRemote(); remote.Add(file, [4, 5, 6]); var entry = Difference(local, file);
        var result = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, default);
        var row = Assert.Single(result.Results);
        string detail = Directory.Exists(Path.Combine(result.RunDirectory, "errors")) ? File.ReadAllText(Directory.GetFiles(Path.Combine(result.RunDirectory, "errors")).Single()) : "";
        Assert.True(row.Status == "Verified", row.Error + detail);
        Assert.Equal(new byte[] { 1, 2, 3 }, await File.ReadAllBytesAsync(row.BackupPath!));
        Assert.Equal(new byte[] { 4, 5, 6 }, await File.ReadAllBytesAsync(local.FullPath));
        Assert.Equal(new byte[] { 9, 8, 7 }, await File.ReadAllBytesAsync(extra.FullPath));
        var workspace = new Workspace(Fixture.Site, f.Root);
        var installed = JsonFiles.Read<Dictionary<string, InstalledRecord>>(Path.Combine(workspace.DataRoot, "installed-manifest.json"))!;
        Assert.Equal("B", installed[SafePaths.Key(file.RelativePath)].Remote.Metadata.Fields.GetProperty("Revision").GetString());
        Assert.Equal("Civil", installed[SafePaths.Key(file.RelativePath)].Remote.Metadata.Fields.GetProperty("Taxonomy").GetProperty("Label").GetString());
        Assert.Equal(file.Modified.UtcDateTime, File.GetLastWriteTimeUtc(local.FullPath));
    }
    [Fact]
    public async Task LocalReuseDoesNotConsumeNetworkOrMoveSource()
    {
        using var f = new Fixture(); var source = await f.LocalAsync("Docs/original.bin", [1, 2, 3]); var file = Fixture.Remote("Docs/new.bin", [1, 2, 3]); var remote = new FakeRemote(); remote.Add(file, [1, 2, 3]);
        var entry = new DiffEntry { RelativePath = file.RelativePath, Status = DiffStatus.MissingLocal, Remote = file, ReusePath = source.FullPath };
        var result = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, false)), null, default);
        Assert.Equal("Verified", Assert.Single(result.Results).Status); Assert.Equal(0, remote.Downloads); Assert.True(File.Exists(source.FullPath));
    }
    [Fact]
    public async Task CorruptDownloadNeverOverwritesOriginal()
    {
        using var f = new Fixture(); var local = await f.LocalAsync("Docs/a", [1, 2, 3]); var file = Fixture.Remote("Docs/a", [4, 5, 6]); var remote = new FakeRemote { Corrupt = true }; remote.Add(file, [4, 5, 6]); var entry = Difference(local, file);
        var r = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, default);
        Assert.Equal("HashMismatch", Assert.Single(r.Results).Error); Assert.Equal(new byte[] { 1, 2, 3 }, await File.ReadAllBytesAsync(local.FullPath));
    }
    [Fact]
    public async Task ConcurrentLocalEditIsPreserved()
    {
        using var f = new Fixture(); var local = await f.LocalAsync("Docs/a", [1, 2, 3]); var file = Fixture.Remote("Docs/a", [4, 5, 6]);
        var remote = new FakeRemote { OnDownload = _ => File.WriteAllBytes(local.FullPath, [9, 9, 9]) }; remote.Add(file, [4, 5, 6]); var entry = Difference(local, file);
        var r = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, default);
        Assert.Equal("FileChanged", Assert.Single(r.Results).Error); Assert.Equal(new byte[] { 9, 9, 9 }, await File.ReadAllBytesAsync(local.FullPath));
    }
    [Fact]
    public async Task BackupFailureKeepsOriginal()
    {
        using var f = new Fixture(); var local = await f.LocalAsync("Docs/a", [1]); var file = Fixture.Remote("Docs/a", [2]); var workspace = new Workspace(Fixture.Site, f.Root);
        var remote = new FakeRemote
        {
            OnRefresh = _ =>
        {
            string run = Directory.GetDirectories(Path.Combine(workspace.DataRoot, "runs")).Single();
            string backup = Path.Combine(run, "backups"); if (!File.Exists(backup)) File.WriteAllText(backup, "block backup directory");
        }
        };
        remote.Add(file, [2]); var entry = Difference(local, file);
        var result = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, default);
        Assert.Equal("Error", Assert.Single(result.Results).Status); Assert.Equal(new byte[] { 1 }, await File.ReadAllBytesAsync(local.FullPath));
    }
    [Fact]
    public async Task ChangedRemoteAndIncompleteMetadataDoNotWriteLocal()
    {
        using var f = new Fixture(); var local = await f.LocalAsync("Docs/a", [1]); var file = Fixture.Remote("Docs/a", [2]); var remote = new FakeRemote { ChangeAfterDownload = true }; remote.Add(file, [2]); var entry = Difference(local, file);
        var r = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, default);
        Assert.Equal("RemoteChanged", Assert.Single(r.Results).Error); Assert.Equal(new byte[] { 1 }, await File.ReadAllBytesAsync(local.FullPath));
        remote.ChangeAfterDownload = false; remote.Files[file.ItemId] = file with { Metadata = file.Metadata with { Complete = false } };
        r = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, default);
        Assert.Equal("MetadataIncomplete", Assert.Single(r.Results).Error); Assert.Equal(new byte[] { 1 }, await File.ReadAllBytesAsync(local.FullPath));
    }
    [Fact]
    public async Task VerifiedCacheSurvivesCancelledAttemptWithoutRedownloading()
    {
        using var f = new Fixture(); var file = Fixture.Remote("Docs/new", [1, 2, 3]); var remote = new FakeRemote(); remote.Add(file, [1, 2, 3]);
        var entry = new DiffEntry { RelativePath = file.RelativePath, Status = DiffStatus.MissingLocal, Remote = file };
        var engine = new SyncEngine(remote); var r = await engine.ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, false)), null, default);
        Assert.Equal("Verified", Assert.Single(r.Results).Status); Assert.Equal(1, remote.Downloads);
        // Simulate a later missing local destination; only this fixture's generated file is removed.
        File.Delete(SafePaths.Resolve(f.Root, file.RelativePath));
        r = await engine.ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, false)), null, default);
        Assert.Equal("Verified", Assert.Single(r.Results).Status); Assert.Equal(1, remote.Downloads); Assert.Equal(0, Assert.Single(r.Results).BytesDownloaded);
    }
    [Fact]
    public async Task NoHashMissingFileIsNotFalselyReportedAsVerified()
    {
        using var f = new Fixture(); var file = Fixture.Remote("Docs/nohash", [1, 2, 3], null); var remote = new FakeRemote(); remote.Add(file, [1, 2, 3]);
        var entry = new DiffEntry { RelativePath = file.RelativePath, Status = DiffStatus.MissingLocal, Remote = file };
        var r = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, false)), null, default);
        Assert.Equal("DownloadedUnverified", Assert.Single(r.Results).Status);
    }
    [Fact]
    public async Task CancelledTransferProducesAuditAndPreservesOriginal()
    {
        using var f = new Fixture(); using var cancel = new CancellationTokenSource(); var local = await f.LocalAsync("Docs/a", [1]); var file = Fixture.Remote("Docs/a", [2]);
        var remote = new FakeRemote { OnDownload = _ => cancel.Cancel() }; remote.Add(file, [2]); var entry = Difference(local, file);
        var r = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entry), Fixture.Plan(new WorkItem(entry, true)), null, cancel.Token);
        Assert.True(r.Cancelled); Assert.Equal("Cancelled", Assert.Single(r.Results).Status); Assert.True(File.Exists(Path.Combine(r.RunDirectory, "results.json")));
        Assert.Equal(new byte[] { 1 }, await File.ReadAllBytesAsync(local.FullPath));
    }
}
