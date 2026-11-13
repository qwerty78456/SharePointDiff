using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class AgentRunnerTests
{
    private sealed class Probe(Func<int, bool> online) : INetworkProbe
    {
        public int Calls { get; private set; }
        public Task<bool> IsOnlineAsync(CancellationToken ct) { ct.ThrowIfCancellationRequested(); return Task.FromResult(online(++Calls)); }
    }
    /// <summary>Virtual clock: every delay completes at once and moves the clock forward by exactly its length, so waits are deterministic.</summary>
    private sealed class SteppingTime : TimeProvider
    {
        private long ticks;
        public override long TimestampFrequency => TimeSpan.TicksPerSecond;
        public override long GetTimestamp() => Interlocked.Read(ref ticks);
        public override DateTimeOffset GetUtcNow() => DateTimeOffset.UnixEpoch.AddTicks(Interlocked.Read(ref ticks));
        public override ITimer CreateTimer(TimerCallback callback, object? state, TimeSpan dueTime, TimeSpan period)
        {
            var timer = new Stepped();
            if (dueTime != Timeout.InfiniteTimeSpan) { Interlocked.Add(ref ticks, dueTime.Ticks); ThreadPool.QueueUserWorkItem(_ => { if (!timer.Disposed) callback(state); }); }
            return timer;
        }
        private sealed class Stepped : ITimer
        {
            public volatile bool Disposed;
            public bool Change(TimeSpan dueTime, TimeSpan period) => true;
            public void Dispose() => Disposed = true;
            public ValueTask DisposeAsync() { Disposed = true; return default; }
        }
    }
    private static async Task<AgentCheckResult> DriveAsync(TimeProvider _, Task<AgentCheckResult> run, TimeSpan __) => await run.WaitAsync(TimeSpan.FromSeconds(30));
    [Fact]
    public async Task RunnerStopsRetryingWhenTheNextWaitWouldExceedTheBudget()
    {
        using var f = new Fixture(); await SelectAsync(f); var time = new SteppingTime();
        var stuck = new FakeRemote { LibrariesFault = new OperationException("RequestTimeout", "graph") };
        // Waits of 1 and 2 minutes fit a 4-minute budget; the next one (5 minutes) does not.
        var r = await new AgentCheckRunner(stuck, new Probe(_ => true), time) { LogonBudget = TimeSpan.FromMinutes(4) }.RunAsync(Fixture.Site, f.Root, false, null, default).WaitAsync(TimeSpan.FromSeconds(30));
        Assert.Equal((AgentState.Failed, "RequestTimeout", 3), (r.State, r.Code, r.Attempts)); Assert.Equal(TimeSpan.FromMinutes(3), r.Duration);
    }
    [Fact]
    public async Task BlankFolderInSettingsIsNotConfiguredNotACrash()
    {
        var r = await new AgentCheckRunner(new FakeRemote(), new Probe(_ => true), new SteppingTime()).RunAsync(Fixture.Site, "  ", false, null, default).WaitAsync(TimeSpan.FromSeconds(30));
        Assert.Equal((AgentState.NotConfigured, "InvalidRoot"), (r.State, r.Code));
    }
    private static async Task SelectAsync(Fixture f) =>
        await LibrarySelectionStore.WriteAsync(new Workspace(Fixture.Site, f.Root), new(["list"], ["list"], DateTimeOffset.Now), default);

    [Fact]
    public async Task RunnerWaitsForNetworkThenReportsOffline()
    {
        using var f = new Fixture(); var time = new SteppingTime(); var probe = new Probe(_ => false);
        var r = await DriveAsync(time, new AgentCheckRunner(new FakeRemote(), probe, time).RunAsync(Fixture.Site, f.Root, false, null, default), TimeSpan.FromSeconds(15));
        Assert.Equal(AgentState.Offline, r.State); Assert.Equal(TimeSpan.FromMinutes(15), r.NetworkWait); Assert.Equal(61, probe.Calls);
        var manual = new Probe(_ => false);
        r = await DriveAsync(time, new AgentCheckRunner(new FakeRemote(), manual, time).RunAsync(Fixture.Site, f.Root, true, null, default), TimeSpan.FromSeconds(15));
        // A manual request goes straight to the actual check even if the interface hint says offline.
        Assert.Equal((AgentState.NotConfigured, "AgentNotReady"), (r.State, r.Code));
        Assert.Equal(0, manual.Calls); Assert.Equal(TimeSpan.Zero, r.NetworkWait);
        // Online on the fourth probe: the check runs (and reports that this mirror was never compared).
        var late = new Probe(n => n >= 4);
        r = await DriveAsync(time, new AgentCheckRunner(new FakeRemote(), late, time).RunAsync(Fixture.Site, f.Root, false, null, default), TimeSpan.FromSeconds(15));
        Assert.Equal((AgentState.NotConfigured, "AgentNotReady"), (r.State, r.Code)); Assert.Equal(TimeSpan.FromSeconds(45), r.NetworkWait); Assert.Equal(4, late.Calls);
    }
    [Fact]
    public async Task RunnerRetriesTransientFailuresWithinBudget()
    {
        using var f = new Fixture(); await SelectAsync(f); var time = new SteppingTime();
        var remote = new FakeRemote(); remote.LibraryFaults.Enqueue(new OperationException("ServiceUnavailable", "graph")); remote.LibraryFaults.Enqueue(new OperationException("AuthNetworkFailed", "login"));
        var r = await DriveAsync(time, new AgentCheckRunner(remote, new Probe(_ => true), time).RunAsync(Fixture.Site, f.Root, false, null, default), TimeSpan.FromSeconds(30));
        Assert.Equal((AgentState.UpToDate, 3), (r.State, r.Attempts)); Assert.Equal(TimeSpan.FromMinutes(3), r.Duration);
        // Still failing after every backoff step (1+2+5+10+15 min, inside 60 min): give up with the code.
        var stuck = new FakeRemote { LibrariesFault = new OperationException("RequestTimeout", "graph") };
        r = await DriveAsync(time, new AgentCheckRunner(stuck, new Probe(_ => true), time).RunAsync(Fixture.Site, f.Root, false, null, default), TimeSpan.FromSeconds(30));
        Assert.Equal((AgentState.Failed, "RequestTimeout", 6), (r.State, r.Code, r.Attempts));
        // A manual check never waits for a retry.
        r = await DriveAsync(time, new AgentCheckRunner(stuck, new Probe(_ => true), time).RunAsync(Fixture.Site, f.Root, true, null, default), TimeSpan.FromSeconds(30));
        Assert.Equal((AgentState.Failed, 1), (r.State, r.Attempts));
        // A library that could not be read is retried; with no work at the end the result is Incomplete.
        var partial = new FakeRemote(); partial.IssuesByLibrary["Docs"] = new("ReadLibraryFailed", "Docs");
        r = await DriveAsync(time, new AgentCheckRunner(partial, new Probe(_ => true), time).RunAsync(Fixture.Site, f.Root, false, null, default), TimeSpan.FromSeconds(30));
        Assert.Equal((AgentState.Incomplete, 6), (r.State, r.Attempts));
    }
    [Fact]
    public async Task RunnerClassifiesAllStatesAndRethrowsUserCancel()
    {
        using var f = new Fixture(); var time = new SteppingTime(); var online = new Probe(_ => true);
        async Task<AgentCheckResult> Run(FakeRemote remote) => await DriveAsync(time, new AgentCheckRunner(remote, online, time).RunAsync(Fixture.Site, f.Root, false, null, default), TimeSpan.FromSeconds(30));
        Assert.Equal(AgentState.NotConfigured, (await Run(new FakeRemote())).State);
        await SelectAsync(f);
        Assert.Equal(AgentState.SignInRequired, (await Run(new FakeRemote { LibrariesFault = new OperationException("SignInRequired", "graph") })).State);
        Assert.Equal(AgentState.UpToDate, (await Run(new FakeRemote())).State);
        var pending = new FakeRemote(); pending.Add(Fixture.Remote("Docs/new", [1]), [1]);
        var r = await Run(pending); Assert.Equal(AgentState.Pending, r.State); Assert.Equal(1, r.Outcome!.Summary.NewFiles);
        await f.LocalAsync("Docs/nohash", [1, 2]); var review = new FakeRemote(); review.Add(Fixture.Remote("Docs/nohash", [1, 2], null), [1, 2]);
        Assert.Equal(AgentState.ReviewOnly, (await Run(review)).State);
        var denied = await Run(new FakeRemote { LibrariesFault = new OperationException("AccessDenied", "graph") });
        Assert.Equal((AgentState.Failed, "AccessDenied"), (denied.State, denied.Code));
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new AgentCheckRunner(new FakeRemote(), online, time).RunAsync(Fixture.Site, f.Root, false, null, cancel.Token));
        await Assert.ThrowsAsync<InvalidOperationException>(() => new AgentCheckRunner(new FakeRemote { LibrariesFault = new InvalidOperationException("deliberate programming fault") }, online, time)
            .RunAsync(Fixture.Site, f.Root, false, null, default));
    }
}

[CollectionDefinition(nameof(ProcessEnvironmentCollection), DisableParallelization = true)]
public sealed class ProcessEnvironmentCollection;

/// <summary>SPDIFF_SETTINGS_PATH is process-wide, so these run outside the parallel collections.</summary>
[Collection(nameof(ProcessEnvironmentCollection))]
public sealed class SettingsAndStartupTests
{
    private static string TempFolder() { string p = Path.Combine(Environment.GetEnvironmentVariable("SPDIFF_TEST_ROOT") ?? Path.GetTempPath(), "settings-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(p); return p; }
    [Fact]
    public async Task OldOrNullFieldsNormalize()
    {
        string dir = TempFolder();
        try
        {
            string path = Path.Combine(dir, "settings.json");
            await File.WriteAllTextAsync(path, "{\"Language\":null,\"SiteUrl\":null,\"Profiles\":null,\"AgentDelaySeconds\":5000}");
            var s = JsonFiles.Read<UserSettings>(path)!.Normalize();
            Assert.Equal(("vi-VN", "", "", 900, false), (s.Language, s.SiteUrl, s.LocalRoot, s.AgentDelaySeconds, s.StartWithWindows)); Assert.Empty(s.Profiles);
            await File.WriteAllTextAsync(path, "{\"Language\":\"en-US\",\"SiteUrl\":\"https://contoso.sharepoint.com/sites/a\",\"Profiles\":[]}");
            Assert.Equal((60, false), (JsonFiles.Read<UserSettings>(path)!.Normalize().AgentDelaySeconds, JsonFiles.Read<UserSettings>(path)!.StartWithWindows));
            await File.WriteAllTextAsync(path, "{\"SelectedProfile\":7,\"Profiles\":[{\"Name\":\"a\",\"TenantId\":\"t\",\"ClientId\":\"c\"},{\"Name\":\"b\",\"TenantId\":\"t\",\"ClientId\":\"c\"}]}");
            Assert.Equal(1, JsonFiles.Read<UserSettings>(path)!.Normalize().SelectedProfile);
        }
        finally { Directory.Delete(dir, true); }
    }
    [Fact]
    public async Task LoadExistingNeverImportsButLoadImportsFromInjectedRoot()
    {
        string dir = TempFolder(); string? previous = Environment.GetEnvironmentVariable("SPDIFF_SETTINGS_PATH");
        try
        {
            Environment.SetEnvironmentVariable("SPDIFF_SETTINGS_PATH", Path.Combine(dir, "app", "settings.json"));
            string config = Path.Combine(dir, "tree", "sp_chuanhoa", "cauhinh"); Directory.CreateDirectory(config);
            await File.WriteAllTextAsync(Path.Combine(config, "cauhinh.json"), "{\"SiteThat\":\"https://contoso.sharepoint.com/sites/x\",\"ThuMucMau\":\"D:\\\\m\",\"Tenant\":\"contoso.onmicrosoft.com\",\"ClientId\":\"11111111-1111-1111-1111-111111111111\"}");
            string importRoot = Path.Combine(dir, "tree", "bin"); Directory.CreateDirectory(importRoot);
            Assert.True(SettingsStore.IsRedirected);
            Assert.Null(SettingsStore.LoadExisting());
            var imported = SettingsStore.Load(importRoot);
            Assert.Equal(("https://contoso.sharepoint.com/sites/x", "contoso.onmicrosoft.com"), (imported.SiteUrl, Assert.Single(imported.Profiles).TenantId));
            Assert.Null(SettingsStore.LoadExisting());
            Assert.Equal(Path.Combine(dir, "app", "msal"), SettingsStore.TokenCacheDirectory); Assert.Equal(Path.Combine(dir, "app", "logs"), SettingsStore.LogDirectory);
            await File.WriteAllTextAsync(EnsureParent(SettingsStore.PathName), "{ not json");
            Assert.Equal("SettingsRead", Assert.Throws<OperationException>(() => SettingsStore.LoadExisting()).Code);
        }
        finally { Environment.SetEnvironmentVariable("SPDIFF_SETTINGS_PATH", previous); Directory.Delete(dir, true); }
    }
    private static string EnsureParent(string path) { Directory.CreateDirectory(Path.GetDirectoryName(path)!); return path; }
    [Theory]
    [InlineData(@"C:\Program Files\SharePoint Diff\SharePointDiff.exe")]
    [InlineData(@"D:\Công cụ\Đồng bộ\SharePointDiff.exe")]
    public void StartupCommandQuotesParsesAndComparesIgnoringCase(string exe)
    {
        string value = StartupCommand.Build(exe);
        Assert.Equal("\"" + exe + "\" --agent", value);
        Assert.True(StartupCommand.TryParse(value, out string parsed, out bool agent)); Assert.Equal((exe, true), (parsed, agent));
        Assert.True(StartupCommand.IsCurrent(value, exe.ToUpperInvariant()));
        Assert.False(StartupCommand.IsCurrent("\"" + exe + "\"", exe));
        Assert.False(StartupCommand.IsCurrent(value, @"C:\Other\SharePointDiff.exe"));
        Assert.True(StartupCommand.TryParse(exe + " --agent", out parsed, out agent)); Assert.Equal((exe, true), (parsed, agent));
        Assert.False(StartupCommand.TryParse("", out _, out _)); Assert.False(StartupCommand.TryParse("\"unterminated", out _, out _));
    }
    [Fact]
    public void TemporaryLocationsAreFlagged()
    {
        Assert.True(StartupCommand.IsTemporaryLocation(@"D:\src\SharePointDiff\src\SharePointDiff.App\bin\Release\net10.0-windows\win-x64\SharePointDiff.exe"));
        Assert.True(StartupCommand.IsTemporaryLocation(@"D:\src\SharePointDiff\publish\win-x64\SharePointDiff.exe"));
        Assert.True(StartupCommand.IsTemporaryLocation(@"D:\src\SharePointDiff\tmp\e2e\bin\SharePointDiff.exe"));
        Assert.True(StartupCommand.IsTemporaryLocation(Path.Combine(Path.GetTempPath(), "x", "SharePointDiff.exe")));
        Assert.False(StartupCommand.IsTemporaryLocation(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "SharePointDiff", "SharePointDiff.exe")));
    }
}
