namespace SharePointDiff.Core;

public enum AgentState { NotConfigured, Offline, SignInRequired, UpToDate, ReviewOnly, Pending, Incomplete, Failed }
public sealed record AgentCheckResult(AgentState State, CheckOutcome? Outcome, string? Code, string? Path, TimeSpan Duration, int Attempts, TimeSpan NetworkWait);

public interface INetworkProbe { Task<bool> IsOnlineAsync(CancellationToken ct); }

/// <summary>Online means Microsoft sign-in answers at all (any HTTP status); an anonymous GET, no token involved.</summary>
public sealed class HttpNetworkProbe(string tenant, HttpClient? client = null) : INetworkProbe
{
    private readonly HttpClient http = client ?? new HttpClient { Timeout = Timeout.InfiniteTimeSpan };
    public async Task<bool> IsOnlineAsync(CancellationToken ct)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct); timeout.CancelAfter(TimeSpan.FromSeconds(10));
        try
        {
            using var response = await http.GetAsync($"https://login.microsoftonline.com/{Uri.EscapeDataString(tenant)}/v2.0/.well-known/openid-configuration", HttpCompletionOption.ResponseHeadersRead, timeout.Token);
            return true;
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested) { return false; }
        catch (HttpRequestException) { return false; }
    }
}

/// <summary>The background check: waits for the network, runs <see cref="DiffService.CheckAsync"/>, retries transient failures
/// inside a time budget and classifies the result. It never writes to the mirror.</summary>
public sealed class AgentCheckRunner(IRemoteReader remote, INetworkProbe probe, TimeProvider time)
{
    public TimeSpan NetworkPoll { get; init; } = TimeSpan.FromSeconds(15);
    public TimeSpan LogonNetworkDeadline { get; init; } = TimeSpan.FromMinutes(15);
    public TimeSpan ManualNetworkDeadline { get; init; } = TimeSpan.FromSeconds(30);
    public TimeSpan LogonBudget { get; init; } = TimeSpan.FromMinutes(60);
    public TimeSpan ManualBudget { get; init; } = TimeSpan.FromMinutes(2);
    public IReadOnlyList<TimeSpan> Backoff { get; init; } = [TimeSpan.FromMinutes(1), TimeSpan.FromMinutes(2), TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(10), TimeSpan.FromMinutes(15)];
    private static readonly string[] NotConfiguredCodes = ["AgentNotReady", "InvalidSite", "InvalidRoot", "InvalidProfile", "NoLibrariesSelected", "LinkBlocked", "OutsideRoot"];
    private static readonly string[] TransientCodes = ["AuthNetworkFailed", "RequestTimeout", "ServiceUnavailable", "WorkspaceBusy"];
    private static readonly string[] TransientLibraryCodes = ["ReadLibraryFailed", "IncompleteInventory", "RequestTimeout", "ServiceUnavailable"];
    public async Task<AgentCheckResult> RunAsync(string site, string root, bool userInitiated, IProgress<OperationProgress>? progress, CancellationToken ct)
    {
        long started = time.GetTimestamp();
        TimeSpan budget = userInitiated ? ManualBudget : LogonBudget, networkDeadline = userInitiated ? ManualNetworkDeadline : LogonNetworkDeadline;
        AgentCheckResult Result(AgentState state, CheckOutcome? outcome, string? code, string? path, int attempts, TimeSpan wait) =>
            new(state, outcome, code, path, time.GetElapsedTime(started), attempts, wait);
        while (!await probe.IsOnlineAsync(ct))
        {
            if (time.GetElapsedTime(started) + NetworkPoll > networkDeadline) return Result(AgentState.Offline, null, null, null, 0, time.GetElapsedTime(started));
            await Task.Delay(NetworkPoll, time, ct);
        }
        TimeSpan networkWait = time.GetElapsedTime(started);
        for (int attempt = 1; ; attempt++)
        {
            bool CanRetry() => !userInitiated && attempt <= Backoff.Count && time.GetElapsedTime(started) + Backoff[attempt - 1] <= budget;
            try
            {
                var outcome = await new DiffService(remote).CheckAsync(site, root, progress, ct);
                var included = outcome.Libraries.Select(l => l.Title).ToHashSet();
                bool partial = outcome.Scan.Issues.Any(i => TransientLibraryCodes.Contains(i.Code) && included.Contains(i.Path));
                if (partial && CanRetry()) { await Task.Delay(Backoff[attempt - 1], time, ct); continue; }
                var s = outcome.Summary;
                var state = s.HasWork ? AgentState.Pending : partial ? AgentState.Incomplete : s.Review > 0 ? AgentState.ReviewOnly : AgentState.UpToDate;
                return Result(state, outcome, null, null, attempt, networkWait);
            }
            catch (OperationException ex) when (ex.Code == "SignInRequired") { return Result(AgentState.SignInRequired, null, ex.Code, ex.ItemPath, attempt, networkWait); }
            catch (OperationException ex) when (NotConfiguredCodes.Contains(ex.Code)) { return Result(AgentState.NotConfigured, null, ex.Code, ex.ItemPath, attempt, networkWait); }
            catch (Exception ex) when (ex is OperationException { Code: var code } && TransientCodes.Contains(code) || ex is HttpRequestException)
            {
                if (CanRetry()) { await Task.Delay(Backoff[attempt - 1], time, ct); continue; }
                return Result(AgentState.Failed, null, (ex as OperationException)?.Code ?? "HttpError", (ex as OperationException)?.ItemPath, attempt, networkWait);
            }
            catch (Exception ex) when (ex is OperationException or IOException or UnauthorizedAccessException)
            {
                return Result(AgentState.Failed, null, ex is OperationException op ? op.Code : "FileIoError", (ex as OperationException)?.ItemPath, attempt, networkWait);
            }
        }
    }
}

/// <summary>The HKCU Run value: <c>"&lt;exe&gt;" --agent</c>.</summary>
public static class StartupCommand
{
    public const string Flag = "--agent";
    public static string Build(string exe) => $"\"{Path.GetFullPath(exe)}\" {Flag}";
    public static bool TryParse(string? value, out string exe, out bool agent)
    {
        exe = ""; agent = false;
        if (string.IsNullOrWhiteSpace(value)) return false;
        string v = value.Trim(), rest;
        if (v.StartsWith('"')) { int end = v.IndexOf('"', 1); if (end < 0) return false; exe = v[1..end]; rest = v[(end + 1)..]; }
        else
        {
            // Unquoted: the path ends at ".exe" (it may contain spaces).
            int end = v.IndexOf(".exe", StringComparison.OrdinalIgnoreCase); if (end < 0) return false;
            exe = v[..(end + 4)]; rest = v[(end + 4)..];
        }
        agent = rest.Split(' ', StringSplitOptions.RemoveEmptyEntries).Contains(Flag, StringComparer.OrdinalIgnoreCase);
        return exe.Length > 0;
    }
    public static bool IsCurrent(string? value, string exe) =>
        TryParse(value, out string registered, out bool agent) && agent && string.Equals(Normalize(registered), Normalize(exe), StringComparison.OrdinalIgnoreCase);
    /// <summary>Build outputs and temporary folders move or vanish; a startup entry pointing there will break.</summary>
    public static bool IsTemporaryLocation(string exe)
    {
        string full = Normalize(exe), temp = Normalize(Path.GetTempPath()).TrimEnd('\\') + '\\';
        if (full.StartsWith(temp, StringComparison.OrdinalIgnoreCase)) return true;
        string[] parts = full.Split('\\');
        return parts.Take(parts.Length - 1).Any(p => p.Equals("bin", StringComparison.OrdinalIgnoreCase) || p.Equals("obj", StringComparison.OrdinalIgnoreCase) ||
            p.Equals("publish", StringComparison.OrdinalIgnoreCase) || p.Equals("tmp", StringComparison.OrdinalIgnoreCase));
    }
    private static string Normalize(string path) { try { return Path.GetFullPath(path); } catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException) { return path; } }
}
