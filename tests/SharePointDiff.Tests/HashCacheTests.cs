using System.Text.Json;
using Microsoft.Extensions.Time.Testing;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class HashCacheTests
{
    [Fact]
    public async Task CachedHashNeverAuthorizesReplacingDifferentActualBytes()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root); var local = await f.LocalAsync("Docs/a", [9]);
        var cache = LocalHashCache.Load(workspace); await cache.ReadAsync(local.FullPath, false, default); await cache.SaveAsync(default);
        string cachePath = Path.Combine(workspace.DataRoot, "local-hash-cache.json");
        var doc = JsonFiles.Read<HashCacheDocument>(cachePath)!; string key = doc.Entries.Keys.Single();
        var wrong = await f.LocalAsync("wrong-hash-source", [1]);
        doc.Entries[key] = doc.Entries[key] with { Hashes = wrong.Hashes };
        await JsonFiles.WriteAsync(cachePath, doc);
        var remote = new FakeRemote(); remote.Add(Fixture.Remote("Docs/a", [2]), [2]);
        var scan = await new DiffService(remote).ScanAsync(Fixture.Site, f.Root, [Fixture.Library], null, default, report: false);
        var entry = scan.Entries.Single(e => e.RelativePath == "Docs/a");
        Assert.Equal(HashBasis.Saved, entry.Local!.HashBasis);
        var applied = await new SyncEngine(remote).ApplyAsync(scan, Fixture.Plan(new WorkItem(entry, true)), null, default);
        Assert.Equal("FileChanged", Assert.Single(applied.Results).Error);
        Assert.Equal(new byte[] { 9 }, await File.ReadAllBytesAsync(local.FullPath));
    }
    [Fact]
    public async Task WarmCompareDoesNotReadContentAndManualFullReadDoes()
    {
        using var f = new Fixture(); var remote = new FakeRemote();
        await f.LocalAsync("Docs/a", [1, 2, 3]); remote.Add(Fixture.Remote("Docs/a", [1, 2, 3]), [1, 2, 3]);
        var service = new DiffService(remote);
        var first = await service.ScanAsync(Fixture.Site, f.Root, [Fixture.Library], null, default, report: false);
        Assert.Equal(HashBasis.Fresh, Assert.Single(first.Entries).Local!.HashBasis);
        var second = await service.ScanAsync(Fixture.Site, f.Root, [Fixture.Library], null, default, report: false);
        Assert.True(NativeFileInfo.SupportsCache(f.Root));
        Assert.Equal(HashBasis.Saved, Assert.Single(second.Entries).Local!.HashBasis);
        Assert.Equal(0, second.Statistics!.Counters.GetValueOrDefault("Content.ReadBytes"));
        Assert.Equal(first.Entries[0].Local!.LastHashedAt, second.Entries[0].Local!.LastHashedAt);
        var forced = await service.ScanAsync(Fixture.Site, f.Root, [Fixture.Library], null, default, report: false, options: new(RehashAllLocal: true, Background: true));
        Assert.Equal(HashBasis.Fresh, forced.Entries[0].Local!.HashBasis);
        Assert.Equal(3, forced.Statistics!.Counters["Content.ReadBytes"]);
        Assert.Equal(1, forced.Statistics.Counters["Hash.Workers"]);
    }
    [Fact]
    public async Task TimestampPreservingEditAndAtomicReplacementInvalidateSavedHashes()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root); var local = await f.LocalAsync("Docs/a", [1, 2]);
        var cache = LocalHashCache.Load(workspace); var first = await cache.ReadAsync(local.FullPath, false, default); await cache.SaveAsync(default);
        await File.WriteAllBytesAsync(local.FullPath, [2, 1]); File.SetLastWriteTimeUtc(local.FullPath, first.LastWriteUtc);
        cache = LocalHashCache.Load(workspace); var edited = await cache.ReadAsync(local.FullPath, false, default);
        Assert.Equal(HashBasis.Fresh, edited.HashBasis); Assert.NotEqual(first.Hashes, edited.Hashes); await cache.SaveAsync(default);
        string replacement = Path.Combine(f.Root, "replacement"); await File.WriteAllBytesAsync(replacement, [2, 1]);
        File.SetLastWriteTimeUtc(replacement, edited.LastWriteUtc); File.Replace(replacement, local.FullPath, null);
        var replaced = await LocalHashCache.Load(workspace).ReadAsync(local.FullPath, false, default);
        Assert.Equal(HashBasis.Fresh, replaced.HashBasis); Assert.NotEqual(edited.Identity!.FileId, replaced.Identity!.FileId);
    }
    [Fact]
    public async Task CorruptCacheFallsBackToRealHash()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root); var local = await f.LocalAsync("Docs/a", [1]);
        await File.WriteAllTextAsync(Path.Combine(workspace.DataRoot, "local-hash-cache.json"), "{broken");
        var cache = LocalHashCache.Load(workspace);
        Assert.Equal("HashCacheInvalid", Assert.Single(cache.Issues).Code);
        Assert.Equal(local.Hashes, (await cache.ReadAsync(local.FullPath, false, default)).Hashes);
    }
    [Fact]
    public async Task IdleCheckIsBoundedAndNeverUpdatesTheFullHashDate()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root);
        byte[] bytes = new byte[LocalHashCache.BlockSize + 10]; new Random(4).NextBytes(bytes);
        var local = await f.LocalAsync("Docs/a", bytes); var cache = LocalHashCache.Load(workspace);
        var observed = await cache.ReadAsync(local.FullPath, false, default); await cache.SaveAsync(default);
        var time = new FakeTimeProvider(DateTimeOffset.UtcNow.AddDays(2));
        var result = await LocalScrubber.RunAsync(workspace, () => false, default, time);
        Assert.Equal(LocalHashCache.BlockSize, result.ReadBytes); Assert.Equal(1, result.VerifiedBlocks);
        var doc = JsonFiles.Read<HashCacheDocument>(Path.Combine(workspace.DataRoot, "local-hash-cache.json"))!;
        Assert.Equal(observed.LastHashedAt, doc.Entries.Values.Single().HashedAt);
        var second = await LocalScrubber.RunAsync(workspace, () => false, default, time);
        Assert.Equal(0, second.ReadBytes);
        time.Advance(TimeSpan.FromDays(1));
        var third = await LocalScrubber.RunAsync(workspace, () => false, default, time);
        Assert.True(third.VerifiedBlocks > 0);
        Assert.Equal(new byte[] { bytes[0], bytes[^1] }, new[] { (await File.ReadAllBytesAsync(local.FullPath))[0], (await File.ReadAllBytesAsync(local.FullPath))[^1] });
    }
    [Fact]
    public async Task IdleCheckPausesWithoutReadingAndCorruptBlockInvalidatesCache()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root); var local = await f.LocalAsync("Docs/a", [1, 2, 3]);
        var cache = LocalHashCache.Load(workspace); await cache.ReadAsync(local.FullPath, false, default); await cache.SaveAsync(default);
        var time = new FakeTimeProvider(DateTimeOffset.UtcNow.AddDays(2));
        var paused = await LocalScrubber.RunAsync(workspace, () => true, default, time);
        Assert.True(paused.Paused); Assert.Equal(0, paused.ReadBytes);
        string path = Path.Combine(workspace.DataRoot, "local-hash-cache.json");
        var doc = JsonFiles.Read<HashCacheDocument>(path)!; string key = doc.Entries.Keys.Single(); var entry = doc.Entries[key];
        doc.Entries[key] = entry with { Blocks = [entry.Blocks[0] with { Sha256 = new string('0', 64) }] };
        await JsonFiles.WriteAsync(path, doc);
        var checkedResult = await LocalScrubber.RunAsync(workspace, () => false, default, time);
        Assert.Equal("HashCacheMismatch", Assert.Single(checkedResult.Issues).Code);
        Assert.Empty(JsonFiles.Read<HashCacheDocument>(path)!.Entries);
        Assert.Equal(new byte[] { 1, 2, 3 }, await File.ReadAllBytesAsync(local.FullPath));
    }
}
