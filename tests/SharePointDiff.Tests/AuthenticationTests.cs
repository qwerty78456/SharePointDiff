using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

/// <summary>All offline: MSAL's HTTP goes to a recorder, and caches are seeded with synthetic, unsigned test tokens.</summary>
public sealed class AuthenticationTests
{
    private const string Graph = "https://graph.microsoft.com";
    private sealed class Recorder(Func<HttpRequestMessage, HttpResponseMessage>? respond = null) : HttpMessageHandler, IMsalHttpClientFactory
    {
        public List<string> Requests { get; } = [];
        public HttpClient GetHttpClient() => new(this, disposeHandler: false);
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            lock (Requests) Requests.Add(request.RequestUri!.AbsolutePath);
            return Task.FromResult(respond?.Invoke(request) ?? new HttpResponseMessage(HttpStatusCode.BadRequest) { Content = new StringContent("{\"error\":\"unexpected_test_request\"}") });
        }
        public int TokenRequests { get { lock (Requests) return Requests.Count(r => r.EndsWith("/oauth2/v2.0/token")); } }
    }
    private sealed class Temp : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(Environment.GetEnvironmentVariable("SPDIFF_TEST_ROOT") ?? System.IO.Path.GetTempPath(), "auth-" + Guid.NewGuid().ToString("N"));
        public Temp() => Directory.CreateDirectory(Path);
        public void Dispose() { if (Directory.Exists(Path)) Directory.Delete(Path, true); }
    }
    private static string B64(string s) => Convert.ToBase64String(Encoding.UTF8.GetBytes(s)).TrimEnd('=').Replace('+', '-').Replace('/', '_');
    private static string Jwt(object payload) => B64("{\"alg\":\"none\",\"typ\":\"JWT\"}") + "." + B64(JsonSerializer.Serialize(payload)) + ".";
    private static ConnectionProfile Profile() => new("test", Guid.NewGuid().ToString(), Guid.NewGuid().ToString());
    /// <summary>A minimal MSAL v3 cache: one account per user id, its id token and refresh token, and a Graph access token.</summary>
    private static byte[] Cache(ConnectionProfile p, string scp = "Sites.Read.All", bool expired = false, int users = 1)
    {
        const string env = "login.microsoftonline.com"; string tid = p.TenantId, client = p.ClientId;
        long now = DateTimeOffset.UtcNow.ToUnixTimeSeconds(), expires = expired ? now - 600 : now + 3600;
        var accounts = new Dictionary<string, object>(); var ids = new Dictionary<string, object>(); var ats = new Dictionary<string, object>(); var rts = new Dictionary<string, object>();
        for (int i = 0; i < users; i++)
        {
            string uid = Guid.NewGuid().ToString(), hai = uid + "." + tid, user = $"user{i}@contoso.test";
            accounts[$"{hai}-{env}-{tid}"] = new { home_account_id = hai, environment = env, realm = tid, local_account_id = uid, username = user, authority_type = "MSSTS", name = "Test", client_info = B64(JsonSerializer.Serialize(new { uid, utid = tid })) };
            ids[$"{hai}-{env}-idtoken-{client}-{tid}-"] = new { home_account_id = hai, environment = env, credential_type = "IdToken", client_id = client, realm = tid,
                secret = Jwt(new { aud = client, iss = $"https://login.microsoftonline.com/{tid}/v2.0", iat = now, nbf = now, exp = now + 3600, name = "Test", oid = uid, preferred_username = user, sub = uid, tid, ver = "2.0" }) };
            ats[$"{hai}-{env}-accesstoken-{client}-{tid}-{Graph}/.default"] = new { home_account_id = hai, environment = env, credential_type = "AccessToken", client_id = client, realm = tid,
                target = Graph + "/.default", secret = Jwt(new { aud = Graph, tid, scp }), cached_at = now.ToString(), expires_on = expires.ToString(), extended_expires_on = expires.ToString(), token_type = "Bearer" };
            rts[$"{hai}-{env}-refreshtoken-{client}--"] = new { home_account_id = hai, environment = env, credential_type = "RefreshToken", client_id = client, secret = "TEST-ONLY-REFRESH-" + i };
        }
        return JsonSerializer.SerializeToUtf8Bytes(new Dictionary<string, object>
        {
            ["Account"] = accounts, ["IdToken"] = ids, ["AccessToken"] = ats, ["RefreshToken"] = rts,
            ["AppMetadata"] = new Dictionary<string, object> { [$"appmetadata-{env}-{client}"] = new { environment = env, client_id = client } }
        });
    }
    private static async Task<MsalCacheHelper> HelperAsync(string dir, ConnectionProfile p) =>
        await MsalCacheHelper.CreateAsync(new StorageCreationPropertiesBuilder(BrowserAuthentication.CacheFileName(p), dir).Build());
    private static async Task SeedAsync(string dir, ConnectionProfile p, byte[] cache) => (await HelperAsync(dir, p)).SaveUnencryptedTokenCache(cache);

    [Fact]
    public async Task NonInteractiveEmptyCacheThrowsSignInRequiredWithoutHttp()
    {
        using var temp = new Temp(); var http = new Recorder();
        var auth = new BrowserAuthentication(Profile(), cacheDirectory: temp.Path, http: http);
        var ex = await Assert.ThrowsAsync<OperationException>(() => auth.Background.GetAsync(Graph, default));
        Assert.Equal("SignInRequired", ex.Code); Assert.Empty(http.Requests);
    }
    [Fact]
    public async Task TokenCacheFileIsDpapiEncryptedAndRoundTrips()
    {
        using var temp = new Temp(); var p = Profile(); byte[] payload = Cache(p);
        await SeedAsync(temp.Path, p, payload);
        byte[] raw = await File.ReadAllBytesAsync(Path.Combine(temp.Path, BrowserAuthentication.CacheFileName(p)));
        Assert.DoesNotContain("TEST-ONLY-REFRESH", Encoding.UTF8.GetString(raw)); Assert.DoesNotContain("RefreshToken", Encoding.UTF8.GetString(raw));
        Assert.Equal(payload, ProtectedData.Unprotect(raw, null, DataProtectionScope.CurrentUser));
    }
    [Fact]
    public void TokenCacheIsPerProfileInsideGivenDirectory()
    {
        using var temp = new Temp(); var a = Profile(); var b = Profile();
        var auth = new BrowserAuthentication(a, cacheDirectory: temp.Path);
        Assert.Equal(Path.Combine(temp.Path, BrowserAuthentication.CacheFileName(a)), auth.CacheFilePath);
        Assert.NotEqual(BrowserAuthentication.CacheFileName(a), BrowserAuthentication.CacheFileName(b));
        Assert.Equal(BrowserAuthentication.CacheFileName(a), BrowserAuthentication.CacheFileName(a with { TenantId = a.TenantId.ToUpperInvariant(), ClientId = a.ClientId.ToUpperInvariant() }));
        Assert.Null(new BrowserAuthentication(a).CacheFilePath);
    }
    [Fact]
    public async Task SeededValidAccessTokenIsSilentAndVisibleToSecondInstance()
    {
        using var temp = new Temp(); var p = Profile(); await SeedAsync(temp.Path, p, Cache(p));
        var http = new Recorder();
        string token = await new BrowserAuthentication(p, cacheDirectory: temp.Path, http: http).Background.GetAsync(Graph, default);
        Assert.True(TokenScopes.IsReadOnly(token)); Assert.Equal(0, http.TokenRequests);
        var second = new BrowserAuthentication(p, cacheDirectory: temp.Path, http: http);
        Assert.Equal(token, await second.Background.GetAsync(Graph, default)); Assert.Equal("user0@contoso.test", second.AccountName);
        Assert.False(second.PersistenceBlocked); Assert.True(File.Exists(second.CacheFilePath));
    }
    [Fact]
    public async Task InvalidGrantMapsToSignInRequiredAndTransportFailureToAuthNetworkFailed()
    {
        using var temp = new Temp(); var p = Profile(); await SeedAsync(temp.Path, p, Cache(p, expired: true));
        var revoked = new Recorder(req => req.RequestUri!.AbsolutePath.EndsWith("/oauth2/v2.0/token")
            ? new(HttpStatusCode.BadRequest) { Content = new StringContent("{\"error\":\"invalid_grant\",\"error_description\":\"AADSTS700082: test refresh token expired\",\"error_codes\":[700082]}", Encoding.UTF8, "application/json") }
            : null!);
        var ex = await Assert.ThrowsAsync<OperationException>(() => new BrowserAuthentication(p, cacheDirectory: temp.Path, http: revoked).Background.GetAsync(Graph, default));
        Assert.Equal("SignInRequired", ex.Code); Assert.True(revoked.TokenRequests >= 1);
        // A fresh cache (MSAL drops a refresh token it was told is invalid), now with the network failing.
        var q = Profile(); await SeedAsync(temp.Path, q, Cache(q, expired: true));
        var offline = new Recorder(_ => throw new HttpRequestException("network unreachable (test)"));
        ex = await Assert.ThrowsAsync<OperationException>(() => new BrowserAuthentication(q, cacheDirectory: temp.Path, http: offline).Background.GetAsync(Graph, default));
        Assert.True(ex.Code == "AuthNetworkFailed", ex.Code + ": " + ex.InnerException); Assert.True(offline.TokenRequests >= 1);
    }
    [Fact]
    public async Task ForgetOnFreshInstanceRemovesPersistedRefreshTokenAndTwoAccountsNeedSignIn()
    {
        using var temp = new Temp(); var p = Profile(); await SeedAsync(temp.Path, p, Cache(p));
        var http = new Recorder();
        await new BrowserAuthentication(p, cacheDirectory: temp.Path, http: http).ForgetAsync(default);
        string remaining = Encoding.UTF8.GetString((await HelperAsync(temp.Path, p)).LoadUnencryptedTokenCache());
        Assert.DoesNotContain("TEST-ONLY-REFRESH", remaining);
        Assert.Equal("SignInRequired", (await Assert.ThrowsAsync<OperationException>(() => new BrowserAuthentication(p, cacheDirectory: temp.Path, http: http).Background.GetAsync(Graph, default))).Code);
        var two = Profile(); await SeedAsync(temp.Path, two, Cache(two, users: 2));
        Assert.Equal("SignInRequired", (await Assert.ThrowsAsync<OperationException>(() => new BrowserAuthentication(two, cacheDirectory: temp.Path, http: http).Background.GetAsync(Graph, default))).Code);
        Assert.Equal(0, http.TokenRequests);
    }
    [Theory]
    [InlineData("AllSites.Read Sites.Read.All User.Read", true)]
    [InlineData("Sites.Read.All offline_access openid profile email", true)]
    [InlineData("AllSites.FullControl", false)]
    [InlineData("Sites.Read.All TermStore.ReadWrite.All", false)]
    [InlineData("Sites.ReadWrite.All", false)]
    [InlineData("AllSites.Write", false)]
    [InlineData("AllSites.Manage", false)]
    [InlineData("Sites.Selected", false)]
    [InlineData("Files.SelectedOperations.Selected", false)]
    public void IsReadOnlyOnlyForKnownReadOnlyScopes(string scp, bool expected) => Assert.Equal(expected, TokenScopes.IsReadOnly(Jwt(new { scp })));
    [Fact]
    public void PermissionReportNamesGrantedAndRejectedScopesWithoutKeepingTheToken()
    {
        string token = Jwt(new { scp = "profile AllSites.FullControl Sites.Read.All AllSites.FullControl" });
        var report = TokenScopes.Inspect(Graph, token);
        Assert.True(report.ScopesReadable); Assert.False(report.IsReadOnly);
        Assert.Equal(Graph, report.Resource);
        Assert.Equal(new[] { "AllSites.FullControl", "profile", "Sites.Read.All" }, report.GrantedScopes);
        Assert.Equal(new[] { "AllSites.FullControl" }, report.RejectedScopes);
        Assert.DoesNotContain(token, JsonSerializer.Serialize(report));
        Assert.DoesNotContain("AllSites.FullControl", TokenScopes.AllowedScopes);
    }
    [Fact]
    public void MissingScopeInformationIsReportedAsUnknownRatherThanAWritePermission()
    {
        foreach (string token in new[] { "opaque-token", Jwt(new { roles = "Sites.Read.All" }), Jwt(new { scp = "" }) })
        {
            var report = TokenScopes.Inspect(Graph, token);
            Assert.False(report.ScopesReadable); Assert.False(report.IsReadOnly);
            Assert.Empty(report.GrantedScopes); Assert.Empty(report.RejectedScopes);
        }
    }
    [Fact]
    public void UnreadableTokensAreNotReadOnly()
    {
        Assert.False(TokenScopes.IsReadOnly("opaque-token")); Assert.Null(TokenScopes.Scopes("opaque-token"));
        Assert.False(TokenScopes.IsReadOnly("a.%%%.c"));
        Assert.False(TokenScopes.IsReadOnly(Jwt(new { roles = "Sites.Read.All" })));
        Assert.False(TokenScopes.IsReadOnly(Jwt(new { scp = "" })));
    }
    [Theory]
    [InlineData("AllSites.FullControl TermStore.ReadWrite.All")]
    [InlineData("Sites.Selected")]
    public async Task TokenThatIsNotReadOnlyIsUsedButNeverKeptOnDisk(string scp)
    {
        using var temp = new Temp(); var p = Profile(); await SeedAsync(temp.Path, p, Cache(p, scp: scp));
        var auth = new BrowserAuthentication(p, cacheDirectory: temp.Path, http: new Recorder());
        string token = await auth.Background.GetAsync(Graph, default);
        Assert.False(TokenScopes.IsReadOnly(token));
        Assert.True(auth.PersistenceBlocked); Assert.False(File.Exists(auth.CacheFilePath));
        var report = Assert.Single(auth.PermissionReports);
        Assert.Equal(Graph, report.Resource); Assert.NotEmpty(report.RejectedScopes);
        Assert.All(scp.Split(' '), scope => Assert.Contains(scope, report.GrantedScopes));
        Assert.Equal(token, await auth.Background.GetAsync(Graph, default)); // still usable from memory in this process
        Assert.False(File.Exists(auth.CacheFilePath));
        // A file that reappears while blocked (for example after a failed delete) is removed by the next token or by Forget.
        await File.WriteAllTextAsync(auth.CacheFilePath!, "left over");
        await auth.Background.GetAsync(Graph, default); Assert.False(File.Exists(auth.CacheFilePath));
        await File.WriteAllTextAsync(auth.CacheFilePath!, "left over");
        await auth.ForgetAsync(default); Assert.False(File.Exists(auth.CacheFilePath));
    }
    [Fact]
    public async Task SignInThatCannotCheckEveryResourceLeavesNothingOnDisk()
    {
        using var temp = new Temp(); var p = Profile(); await SeedAsync(temp.Path, p, Cache(p)); // Graph token only; SharePoint needs the refresh token
        var offline = new Recorder(_ => throw new HttpRequestException("network unreachable (test)"));
        var auth = new BrowserAuthentication(p, cacheDirectory: temp.Path, http: offline);
        var ex = await Assert.ThrowsAsync<OperationException>(() => auth.SignInAsync([Graph, "https://contoso.sharepoint.com"], default));
        Assert.Equal("AuthNetworkFailed", ex.Code);
        Assert.False(File.Exists(auth.CacheFilePath), "a sign-in that saw only the Graph token must not keep the refresh token on disk");
        Assert.False(auth.PersistenceBlocked);
    }
}
