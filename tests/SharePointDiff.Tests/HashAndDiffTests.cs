using System.Numerics;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class HashAndDiffTests
{
    private static string Reference(byte[] bytes)
    {
        BigInteger h = 0, mask = (BigInteger.One << 160) - 1;
        for (int i = 0; i < bytes.Length; i++) { var v = (BigInteger)bytes[i] << ((i * 11) % 160); h ^= (v & mask) | (v >> 160); }
        h ^= (BigInteger)bytes.LongLength << 96;
        var result = new byte[20]; h.ToByteArray(true, false).CopyTo(result, 0); return Convert.ToBase64String(result);
    }
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(19)]
    [InlineData(159)]
    [InlineData(160)]
    [InlineData(161)]
    [InlineData(321)]
    [InlineData(10001)]
    [InlineData(1048737)]
    public void StreamingHashMatchesIndependentReference(int size)
    {
        byte[] bytes = new byte[size]; new Random(700).NextBytes(bytes);
        var hash = new QuickXorHasher();
        for (int i = 0; i < size; i += 137) hash.Append(bytes.AsSpan(i, Math.Min(137, size - i)));
        Assert.Equal(Reference(bytes), hash.Finish());
    }
    [Fact]
    public async Task FileHashIsIndependentOfIoChunkSize()
    {
        using var f = new Fixture(); byte[] data = new byte[2_100_031]; new Random(7).NextBytes(data);
        var local = await f.LocalAsync("Docs/large.bin", data);
        Assert.Equal(Reference(data), local.Hashes.QuickXorHash);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(data)), local.Hashes.Sha256);
    }
    [Theory]
    [InlineData("../file")]
    [InlineData("Docs/../../file")]
    [InlineData("C:/file")]
    [InlineData("Docs/NUL.txt")]
    [InlineData("Docs/bad. ")]
    [InlineData("Docs/a:b")]
    public void UnsafeWindowsPathsAreBlocked(string path) => Assert.Throws<OperationException>(() => SafePaths.Parts(path));
    [Fact]
    public async Task UnicodeMatchingAndDedupAreReadOnly()
    {
        using var f = new Fixture(); var old = await f.LocalAsync("Docs/Tài liệu.bin", [1, 2, 3]);
        var same = Fixture.Remote("Docs/Tài liệu.bin", [1, 2, 3]);
        var missing = Fixture.Remote("Docs/new.bin", [1, 2, 3]);
        var remote = new FakeRemote(); remote.Add(same, [1, 2, 3]); remote.Add(missing, [1, 2, 3]);
        var result = await new DiffService(remote).ScanAsync(Fixture.Site, f.Root, [Fixture.Library], null, default);
        Assert.Contains(result.Entries, e => e.RelativePath == same.RelativePath && e.Status == DiffStatus.Match);
        Assert.Equal(old.FullPath, result.Entries.Single(e => e.RelativePath == missing.RelativePath).ReusePath);
        Assert.Equal(0, remote.Downloads);
        Assert.False(File.Exists(Path.Combine(f.Root, "Docs", "new.bin")));
    }
    [Fact]
    public async Task SameSizeDifferentHashAndNoHashAreDistinct()
    {
        using var f = new Fixture(); await f.LocalAsync("Docs/a", [1, 2]); await f.LocalAsync("Docs/b", [3, 4]); await f.LocalAsync("Docs/local", [5]);
        var remote = new FakeRemote(); remote.Add(Fixture.Remote("Docs/a", [2, 1]), [2, 1]); remote.Add(Fixture.Remote("Docs/b", [3, 4], null), [3, 4]);
        var scan = await new DiffService(remote).ScanAsync(Fixture.Site, f.Root, [Fixture.Library], null, default);
        Assert.Equal(DiffStatus.DifferentHash, scan.Entries.Single(e => e.RelativePath == "Docs/a").Status);
        Assert.Equal(DiffStatus.Unverified, scan.Entries.Single(e => e.RelativePath == "Docs/b").Status);
        Assert.Equal(DiffStatus.LocalOnly, scan.Entries.Single(e => e.RelativePath == "Docs/local").Status);
        var plan = DiffService.Plan(scan.Entries, new HashSet<string>(), OverwriteMode.AllDifferences);
        Assert.Single(plan.Items); Assert.Equal("Docs/a", plan.Items[0].Entry.RelativePath);
    }
    [Fact]
    public async Task FailedLibraryIsNotEmptyOrLocalOnly()
    {
        using var f = new Fixture(); await f.LocalAsync("Docs/a", [1]); var remote = new FakeRemote(); remote.Issues.Add(new("ReadLibraryFailed", "Docs"));
        var scan = await new DiffService(remote).ScanAsync(Fixture.Site, f.Root, [Fixture.Library], null, default);
        Assert.Equal(DiffStatus.Error, Assert.Single(scan.Entries).Status); Assert.Empty(DiffService.Plan(scan.Entries, new HashSet<string>(), OverwriteMode.AllDifferences).Items);
    }
    [Fact]
    public void SelectionAndAllModesHaveExplicitScopes()
    {
        DiffEntry Entry(string p, DiffStatus s) => new() { RelativePath = p, Status = s, Remote = Fixture.Remote(p, [1]), Local = new(p, p, new(1, "x", "y"), DateTime.UtcNow) };
        var entries = new[] { Entry("Docs/a", DiffStatus.DifferentHash), Entry("Docs/b", DiffStatus.DifferentSize), Entry("Docs/m", DiffStatus.MissingLocal), Entry("Docs/u", DiffStatus.Unverified), Entry("Docs/l", DiffStatus.LocalOnly) };
        var selected = new HashSet<string> { SafePaths.Key("Docs/a"), SafePaths.Key("Docs/m"), SafePaths.Key("Docs/u"), SafePaths.Key("Docs/l") };
        Assert.Single(DiffService.Plan(entries, selected, OverwriteMode.KeepLocal).Items);
        Assert.Equal(2, DiffService.Plan(entries, selected, OverwriteMode.Selected).Items.Count);
        Assert.Equal(3, DiffService.Plan(entries, selected, OverwriteMode.AllDifferences).Items.Count);
    }
    [Fact]
    public void ResourcesHaveIdenticalKeysAndFormatParameters()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory); while (dir != null && !File.Exists(Path.Combine(dir.FullName, "Directory.Build.props"))) dir = dir.Parent;
        Assert.NotNull(dir); string root = Path.Combine(dir!.FullName, "src", "SharePointDiff.App", "Resources");
        var en = XDocument.Load(Path.Combine(root, "Strings.en-US.resx")).Root!.Elements("data").ToDictionary(x => (string)x.Attribute("name")!, x => x.Element("value")!.Value);
        var vi = XDocument.Load(Path.Combine(root, "Strings.vi-VN.resx")).Root!.Elements("data").ToDictionary(x => (string)x.Attribute("name")!, x => x.Element("value")!.Value);
        Assert.Equal(en.Keys.Order(), vi.Keys.Order());
        foreach (string k in en.Keys)
        {
            string Pattern(string s) => string.Join("|", System.Text.RegularExpressions.Regex.Matches(s, @"\{\d+[^}]*\}").Select(m => m.Value));
            Assert.Equal(Pattern(en[k]), Pattern(vi[k]));
            Assert.False(string.IsNullOrWhiteSpace(en[k])); Assert.False(string.IsNullOrWhiteSpace(vi[k]));
        }
    }
}
