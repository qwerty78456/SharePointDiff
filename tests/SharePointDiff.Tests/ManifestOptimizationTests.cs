using System.Text.Json;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class ManifestOptimizationTests
{
    [Fact]
    public async Task SharedSchemaIsReadOnlyOnceAcrossManyJournalOrBackupRecordsInOneOperation()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root); var remote = Fixture.Remote("Docs/a", [1]);
        string record = Path.Combine(workspace.DataRoot, "sample-record.json");
        await JsonFiles.WriteAsync(record, new InstalledRecord(remote.RelativePath, "sha", remote.QuickXorHash, remote.ETag, DateTimeOffset.Now, remote));
        using var metrics = IoMetrics.Begin("SchemaReads"); using var schemas = SchemaStorage.Begin(workspace);
        for (int i = 0; i < 100; i++) Assert.NotNull(JsonFiles.Read<InstalledRecord>(record));
        Assert.Equal(101, metrics.Snapshot().Counters["Metadata.Reads"]);
    }
    [Fact]
    public async Task ApplyLoadsManifestOnceAndCheckpointsInBatchesWithOneSharedSchema()
    {
        using var f = new Fixture(); var remote = new FakeRemote();
        var files = Enumerable.Range(0, 120).Select(i => Fixture.Remote("Docs/" + i, [(byte)i])).ToArray();
        foreach (var file in files) remote.Add(file, [(byte)int.Parse(file.Name)]);
        var entries = files.Select(file => new DiffEntry { RelativePath = file.RelativePath, Remote = file, Status = DiffStatus.MissingLocal }).ToArray();
        var result = await new SyncEngine(remote).ApplyAsync(Fixture.Scan(f.Root, entries), Fixture.Plan(entries.Select(e => new WorkItem(e, false)).ToArray()), null, default);
        Assert.All(result.Results, r => Assert.Equal("Verified", r.Status));
        var workspace = new Workspace(Fixture.Site, f.Root);
        var io = JsonFiles.Read<IoStatistics>(Path.Combine(result.RunDirectory, "io.json"))!;
        Assert.Equal(1, io.Counters["Manifest.Loads"]);
        Assert.InRange(io.Counters["Manifest.Checkpoints"], 3, 5);
        Assert.Single(Directory.GetFiles(Path.Combine(workspace.DataRoot, "schemas"), "*.json"));
        var records = await ManifestStore.ReadAllAsync(workspace, default);
        Assert.Equal(120, records.Count);
        Assert.All(records.Values, r => Assert.Equal("B", r.Remote.Metadata.Fields.GetProperty("Revision").GetString()));
        string manifestText = File.ReadAllText(Path.Combine(workspace.DataRoot, "installed-manifest.json"));
        Assert.Contains("\"FormatVersion\": 2", manifestText); Assert.DoesNotContain("TypeAsString", manifestText);
    }

    [Fact]
    public async Task ReplaysEveryUnmergedRunInSequenceAndDoesNotReplayOlderJournalOverCheckpoint()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root);
        var old = Fixture.Remote("Docs/a", [1]); var latest = old with { ETag = "latest" };
        var h = new FileHashes(1, old.QuickXorHash!, "SHA");
        async Task Journal(string run, long sequence, RemoteFile remote)
        {
            var record = new InstalledRecord(remote.RelativePath, h.Sha256, remote.QuickXorHash, remote.ETag, DateTimeOffset.Now, remote);
            await JsonFiles.WriteAsync(Path.Combine(run, "committed", JsonFiles.StableId(remote.RelativePath) + ".json"),
                new CommitJournal(sequence, record, h, new(remote.RelativePath, "Verified", null, null, h.Sha256, remote.QuickXorHash, 0)));
        }
        await Journal(Path.Combine(workspace.DataRoot, "runs", "z-earlier"), 1, old);
        await Journal(Path.Combine(workspace.DataRoot, "runs", "a-later"), 2, latest);
        await using (var session = await ManifestSession.OpenAsync(workspace, default))
            Assert.Equal("latest", (await session.ReadAsync(old.RelativePath, default))!.ETag);
        await using (var session = await ManifestSession.OpenAsync(workspace, default))
            Assert.Equal("latest", (await session.ReadAsync(old.RelativePath, default))!.ETag);
        Assert.Equal(2, JsonFiles.Read<ManifestEnvelope>(Path.Combine(workspace.DataRoot, "installed-manifest.json"))!.LastSequence);
    }

    [Fact]
    public async Task LegacyManifestIsPreservedByteForByteAndTypedMetadataStillRoundTrips()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root);
        var remote = Fixture.Remote("Docs/a", [1]);
        var record = new InstalledRecord("Docs/a", "sha", remote.QuickXorHash, remote.ETag, DateTimeOffset.Now, remote);
        string path = Path.Combine(workspace.DataRoot, "installed-manifest.json");
        // A true v1 file has an inline schema, serialized outside the workspace context.
        byte[] original = JsonSerializer.SerializeToUtf8Bytes(new Dictionary<string, InstalledRecord> { [SafePaths.Key(record.Path)] = record }, JsonFiles.Options);
        await File.WriteAllBytesAsync(path, original);
        await using (await ManifestSession.OpenAsync(workspace, default)) { }
        Assert.Equal(original, await File.ReadAllBytesAsync(path + ".v1.backup"));
        var installed = (await ManifestStore.ReadAllAsync(workspace, default)).Values.Single();
        Assert.Equal(JsonValueKind.Array, installed.Remote.Metadata.Fields.GetProperty("PeopleId").ValueKind);
        Assert.Equal("term", installed.Remote.Metadata.Fields.GetProperty("Taxonomy").GetProperty("TermGuid").GetString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MissingOrCorruptSharedSchemaFailsClosedInsteadOfDroppingMetadata(bool corrupt)
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root);
        var remote = Fixture.Remote("Docs/a", [1]);
        await ManifestStore.SaveAsync(workspace, remote, new(1, remote.QuickXorHash!, "sha"), default);
        string schema = Directory.GetFiles(Path.Combine(workspace.DataRoot, "schemas")).Single();
        if (corrupt) await File.WriteAllTextAsync(schema, "{broken"); else File.Delete(schema);
        var error = await Assert.ThrowsAsync<OperationException>(() => ManifestStore.ReadAllAsync(workspace, default));
        Assert.Equal("MetadataStoreCorrupt", error.Code);
    }
    [Fact]
    public async Task InterruptedIntentOnlyObservesAndNeverRestoresOverUserEdits()
    {
        using var f = new Fixture(); var workspace = new Workspace(Fixture.Site, f.Root);
        var local = await f.LocalAsync("Docs/a", [9]); var remote = Fixture.Remote("Docs/a", [1]);
        string run = workspace.CreateRun();
        await JsonFiles.WriteAsync(Path.Combine(run, "intents", JsonFiles.StableId(local.RelativePath) + ".json"),
            new CommitIntent(local.RelativePath, null, "old", "new", remote));
        await using (var session = await ManifestSession.OpenAsync(workspace, default))
            Assert.Equal("InterruptedCommit", Assert.Single(session.RecoveryIssues).Code);
        Assert.Equal(new byte[] { 9 }, await File.ReadAllBytesAsync(local.FullPath));
        Assert.True(File.Exists(Path.Combine(run, "recovery-reviewed.json")));
    }
}
