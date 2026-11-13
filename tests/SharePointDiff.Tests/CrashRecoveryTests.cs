using System.Diagnostics;
using System.Reflection;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class CrashRecoveryTests
{
    [Theory]
    [InlineData("AfterBackup", false, false)]
    [InlineData("AfterIntent", false, false)]
    [InlineData("AfterReplace", true, false)]
    [InlineData("AfterVerification", true, false)]
    [InlineData("AfterCommitJournal", true, true)]
    [InlineData("AfterCheckpoint", true, true)]
    public async Task ForcedProcessTerminationPreservesOriginalAndReplaysOnlyDurableCommits(string stage, bool changed, bool committed)
    {
        using var f = new Fixture();
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "SharePointDiff.sln"))) directory = directory.Parent;
        Assert.NotNull(directory);
        string configuration = typeof(CrashRecoveryTests).Assembly.GetCustomAttribute<AssemblyConfigurationAttribute>()!.Configuration;
        string probe = Environment.GetEnvironmentVariable("SPDIFF_PROBE_PATH") ?? Path.Combine(directory!.FullName, "tests", "SharePointDiff.Probe", "bin", configuration, "net10.0", "SharePointDiff.Probe.dll");
        Assert.True(File.Exists(probe), "Build the referenced crash probe first");
        var start = new ProcessStartInfo("dotnet") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardError = true, RedirectStandardOutput = true };
        start.ArgumentList.Add(probe); start.ArgumentList.Add("--crash"); start.ArgumentList.Add(f.Root); start.ArgumentList.Add(stage);
        using var process = Process.Start(start)!;
        var stderr = process.StandardError.ReadToEndAsync(); var stdout = process.StandardOutput.ReadToEndAsync();
        try { await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(30)); }
        catch (TimeoutException) { process.Kill(); throw; }
        string detail = await stderr + await stdout;
        Assert.True(process.ExitCode != 0, detail);
        Assert.DoesNotContain("requested crash boundary was not reached", detail);
        var workspace = new Workspace("https://contoso.sharepoint.com/sites/crash", f.Root);
        string backup = Directory.GetFiles(Path.Combine(workspace.DataRoot, "runs"), "a", SearchOption.AllDirectories).Single(p => p.Contains("backups"));
        Assert.Equal(new byte[] { 1, 2, 3 }, await File.ReadAllBytesAsync(backup));
        Assert.Equal(changed ? new byte[] { 4, 5, 6 } : new byte[] { 1, 2, 3 }, await File.ReadAllBytesAsync(Path.Combine(f.Root, "Docs", "a")));
        using (WorkspaceLock.Acquire(workspace))
        await using (var session = await ManifestSession.OpenAsync(workspace, default))
        {
            var restored = await session.ReadAsync("Docs/a", default);
            if (committed) { Assert.NotNull(restored); Assert.Equal((await FileHashing.ComputeAsync(Path.Combine(f.Root, "Docs", "a"))).Sha256, restored.Sha256); }
            else Assert.Null(restored);
        }
        Assert.Equal(new byte[] { 1, 2, 3 }, await File.ReadAllBytesAsync(backup));
    }
}
