using System.Collections.Concurrent;
using System.Text.Json;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;

namespace SharePointDiff.Core;

public interface ITokenProvider { Task<string> GetAsync(string resource, CancellationToken ct, bool refresh = false); }

/// <summary>Delegated sign-in through the system browser. With a cache directory the MSAL token cache is persisted,
/// DPAPI-encrypted for the current Windows user, unless a token turns out to carry write scopes.</summary>
public sealed class BrowserAuthentication : ITokenProvider
{
    private static readonly ConcurrentDictionary<string, BrowserAuthentication> Sessions = new();
    private readonly IPublicClientApplication app;
    private readonly SemaphoreSlim gate = new(1, 1);
    private readonly bool selectAccount;
    private readonly string? cacheDirectory;
    private MsalCacheHelper? cache;
    private bool cacheReady;
    private IAccount? account;
    // Resources whose current token was seen with read-only scopes in this process.
    private readonly HashSet<string> verified = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, TokenPermissionReport> permissions = new(StringComparer.OrdinalIgnoreCase);
    /// <summary>Permission names only. Never contains access/refresh tokens, account identifiers or credentials.</summary>
    public IReadOnlyList<TokenPermissionReport> PermissionReports => permissions.Values.OrderBy(p => p.Resource, StringComparer.OrdinalIgnoreCase).ToArray();
    private static string ResourceKey(string resource) => resource.TrimEnd('/');
    public string? AccountName => account?.Username;
    /// <summary>The encrypted cache file, or null when nothing is persisted.</summary>
    public string? CacheFilePath { get; }
    /// <summary>A token that is not provably read-only was seen: the cache file was removed and tokens stay in memory for this process.</summary>
    public bool PersistenceBlocked { get; private set; }
    /// <summary>Never opens a browser; a missing or expired sign-in becomes SignInRequired.</summary>
    public ITokenProvider Background { get; }
    /// <summary>May open the system browser; use only for requests that start from a user's click.</summary>
    public ITokenProvider Interactive => this;
    public BrowserAuthentication(ConnectionProfile profile, bool selectAccount = true, string? cacheDirectory = null, IMsalHttpClientFactory? http = null)
    {
        this.selectAccount = selectAccount;
        if (!Guid.TryParse(profile.ClientId, out _) || string.IsNullOrWhiteSpace(profile.TenantId) ||
            profile.TenantId.Any(c => !(char.IsAsciiLetterOrDigit(c) || c is '.' or '-')))
            throw new OperationException("InvalidProfile");
        var builder = PublicClientApplicationBuilder.Create(profile.ClientId)
            .WithAuthority($"https://login.microsoftonline.com/{profile.TenantId}")
            .WithRedirectUri("http://localhost");
        if (http != null) builder = builder.WithHttpClientFactory(http);
        app = builder.Build();
        this.cacheDirectory = cacheDirectory;
        if (cacheDirectory != null) CacheFilePath = Path.Combine(cacheDirectory, CacheFileName(profile));
        Background = new NonInteractive(this);
    }
    public static string CacheFileName(ConnectionProfile profile) => "msal-" + JsonFiles.StableId((profile.TenantId + "|" + profile.ClientId).ToLowerInvariant()) + ".cache";
    /// <summary>One MSAL client per profile and cache folder per process, so the main window and the agent share accounts and tokens.</summary>
    public static BrowserAuthentication Shared(ConnectionProfile profile, string? cacheDirectory) =>
        Sessions.GetOrAdd($"{profile.TenantId}|{profile.ClientId}|{cacheDirectory}".ToLowerInvariant(), _ => new(profile, true, cacheDirectory));
    public Task<string> GetAsync(string resource, CancellationToken ct, bool refresh = false) => GetCoreAsync(resource, ct, refresh, true);
    /// <summary>User-initiated sign-in for each resource the reader will need. Fails closed: unless every resource's token was seen
    /// read-only, nothing from this sign-in stays on disk (the refresh token redeems tokens for every resource of the app).</summary>
    public async Task SignInAsync(IReadOnlyList<string> resources, CancellationToken ct)
    {
        try { foreach (string resource in resources) await GetCoreAsync(resource, ct, false, true); }
        finally
        {
            if (resources.Any(r => !verified.Contains(ResourceKey(r))))
            {
                await gate.WaitAsync(CancellationToken.None);
                try { if (cache != null) DeleteCacheFile(); }
                finally { gate.Release(); }
            }
        }
    }
    /// <summary>Removes every account from memory and from the persisted cache.</summary>
    public async Task ForgetAsync(CancellationToken ct)
    {
        await gate.WaitAsync(ct);
        try
        {
            await EnsureCacheAsync();
            foreach (var a in await app.GetAccountsAsync()) await app.RemoveAsync(a);
            account = null; verified.Clear(); permissions.Clear();
            if ((await app.GetAccountsAsync()).Any()) throw new OperationException("TokenCacheUnavailable", CacheFilePath);
            // While persistence is blocked the helper is unregistered, so remove a file a failed delete may have left.
            if (PersistenceBlocked) DeleteCacheFile();
        }
        catch (MsalException ex) { throw new OperationException("TokenCacheUnavailable", CacheFilePath, ex); }
        finally { gate.Release(); }
    }
    private async Task<string> GetCoreAsync(string resource, CancellationToken ct, bool refresh, bool interactive)
    {
        await gate.WaitAsync(ct);
        try
        {
            await EnsureCacheAsync();
            string[] scopes = [resource.TrimEnd('/') + "/.default"];
            account ??= await SingleAccountAsync();
            if (account != null)
            {
                try { return Accept(resource, await app.AcquireTokenSilent(scopes, account).WithForceRefresh(refresh).ExecuteAsync(ct)); }
                catch (MsalUiRequiredException)
                {
                    // Another provider in this process may have replaced the account: re-read once.
                    var current = await SingleAccountAsync();
                    if (current != null && current.HomeAccountId?.Identifier != account.HomeAccountId?.Identifier)
                    {
                        account = current;
                        try { return Accept(resource, await app.AcquireTokenSilent(scopes, account).WithForceRefresh(refresh).ExecuteAsync(ct)); }
                        catch (MsalUiRequiredException) { }
                    }
                }
            }
            if (!interactive) throw new OperationException("SignInRequired", resource);
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
            timeout.CancelAfter(TimeSpan.FromMinutes(5));
            var request = app.AcquireTokenInteractive(scopes).WithUseEmbeddedWebView(false);
            if (account != null) request = request.WithAccount(account);
            else if (selectAccount) request = request.WithPrompt(Prompt.SelectAccount);
            AuthenticationResult result;
            try { result = await request.ExecuteAsync(timeout.Token); }
            // The browser was left unanswered: a sign-in failure, never a cancellation or a request timeout.
            catch (OperationCanceledException ex) when (!ct.IsCancellationRequested) { throw new OperationException("AuthenticationFailed", resource, ex); }
            account = result.Account;
            foreach (var other in (await app.GetAccountsAsync()).Where(a => a.HomeAccountId?.Identifier != account.HomeAccountId?.Identifier)) await app.RemoveAsync(other);
            return Accept(resource, result);
        }
        catch (MsalException ex) { throw new OperationException(IsNetworkFailure(ex) ? "AuthNetworkFailed" : "AuthenticationFailed", resource, ex); }
        // MSAL lets transport failures of a silent refresh escape unwrapped; name them so callers stop instead of retrying per library.
        catch (HttpRequestException ex) { throw new OperationException("AuthNetworkFailed", resource, ex); }
        catch (OperationCanceledException ex) when (!ct.IsCancellationRequested) { throw new OperationException("AuthNetworkFailed", resource, ex); }
        finally { gate.Release(); }
    }
    private async Task<IAccount?> SingleAccountAsync() => (await app.GetAccountsAsync()).ToArray() is [var only] ? only : null;
    private static bool IsNetworkFailure(MsalException ex) =>
        ex.IsRetryable || ex.InnerException is HttpRequestException or OperationCanceledException || ex.ErrorCode is MsalError.RequestTimeout or MsalError.ServiceNotAvailable;
    private async Task EnsureCacheAsync()
    {
        if (cacheReady || CacheFilePath == null || PersistenceBlocked) return;
        try
        {
            Directory.CreateDirectory(cacheDirectory!);
            // On Windows the builder's default storage is a DPAPI (CurrentUser) encrypted file.
            cache = await MsalCacheHelper.CreateAsync(new StorageCreationPropertiesBuilder(Path.GetFileName(CacheFilePath), cacheDirectory).Build());
            cache.VerifyPersistence();
            cache.RegisterCache(app.UserTokenCache);
            cacheReady = true;
        }
        catch (Exception ex) when (ex is MsalCachePersistenceException or IOException or UnauthorizedAccessException)
        { throw new OperationException("TokenCacheUnavailable", CacheFilePath, ex); }
    }
    /// <summary>Only provably read-only tokens may stay on disk. Anything else stops persisting for this process and deletes the file
    /// (retried on every later token while it is still there).</summary>
    private string Accept(string resource, AuthenticationResult result)
    {
        var report = TokenScopes.Inspect(ResourceKey(resource), result.AccessToken);
        permissions[ResourceKey(resource)] = report;
        if (cache != null && !PersistenceBlocked)
        {
            if (report.IsReadOnly) verified.Add(ResourceKey(resource));
            else { PersistenceBlocked = true; cache.UnregisterCache(app.UserTokenCache); }
        }
        if (PersistenceBlocked) DeleteCacheFile();
        return result.AccessToken;
    }
    private void DeleteCacheFile()
    {
        if (CacheFilePath == null) return;
        for (int attempt = 1; ; attempt++)
        {
            try { if (File.Exists(CacheFilePath)) File.Delete(CacheFilePath); return; }
            catch (Exception ex) when ((ex is IOException or UnauthorizedAccessException) && attempt < 3) { Thread.Sleep(50); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { throw new OperationException("TokenCacheUnavailable", CacheFilePath, ex); }
        }
    }
    private sealed class NonInteractive(BrowserAuthentication owner) : ITokenProvider
    {
        public Task<string> GetAsync(string resource, CancellationToken ct, bool refresh = false) => owner.GetCoreAsync(resource, ct, refresh, false);
    }
}

public sealed record TokenPermissionReport(string Resource, bool ScopesReadable, IReadOnlyList<string> GrantedScopes, IReadOnlyList<string> RejectedScopes)
{
    public bool IsReadOnly => ScopesReadable && GrantedScopes.Count > 0 && RejectedScopes.Count == 0;
}

public static class TokenScopes
{
    /// <summary>Delegated scopes that cannot change anything. Everything else counts as possibly writing, including per-site grants such as
    /// Sites.Selected, whose level the token does not show.</summary>
    private static readonly HashSet<string> ReadOnly = new(StringComparer.OrdinalIgnoreCase)
    { "Sites.Read.All", "AllSites.Read", "Files.Read", "Files.Read.All", "MyFiles.Read", "TermStore.Read.All", "Sites.Search.All", "User.Read", "User.ReadBasic.All", "openid", "profile", "email", "offline_access" };
    public static IReadOnlyList<string> AllowedScopes { get; } = Array.AsReadOnly(ReadOnly.OrderBy(s => s, StringComparer.OrdinalIgnoreCase).ToArray());
    public static TokenPermissionReport Inspect(string resource, string accessToken)
    {
        var scopes = Scopes(accessToken);
        var granted = (scopes ?? []).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s, StringComparer.OrdinalIgnoreCase).ToArray();
        return new(resource, scopes is { Count: > 0 }, Array.AsReadOnly(granted), Array.AsReadOnly(granted.Where(s => !ReadOnly.Contains(s)).ToArray()));
    }
    /// <summary>The delegated <c>scp</c> claim, read without validating the token; null when the token is not a readable JWT with scopes.</summary>
    public static IReadOnlyList<string>? Scopes(string accessToken)
    {
        string[] parts = accessToken.Split('.');
        if (parts.Length < 2) return null;
        try
        {
            string payload = parts[1].Replace('-', '+').Replace('_', '/');
            payload = payload.PadRight(payload.Length + (4 - payload.Length % 4) % 4, '=');
            using var json = JsonDocument.Parse(Convert.FromBase64String(payload));
            if (!json.RootElement.TryGetProperty("scp", out var scp) || scp.ValueKind != JsonValueKind.String) return null;
            return scp.GetString()!.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        }
        catch (Exception ex) when (ex is FormatException or JsonException) { return null; }
    }
    /// <summary>True only for a readable token whose every scope is known to be read-only (fails closed).</summary>
    public static bool IsReadOnly(string accessToken) => Scopes(accessToken) is { Count: > 0 } scopes && scopes.All(ReadOnly.Contains);
}
