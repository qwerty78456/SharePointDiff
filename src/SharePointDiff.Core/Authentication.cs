using Microsoft.Identity.Client;

namespace SharePointDiff.Core;

public interface ITokenProvider { Task<string> GetAsync(string resource, CancellationToken ct, bool refresh = false); }

public sealed class BrowserAuthentication : ITokenProvider
{
    private readonly IPublicClientApplication app;
    private readonly SemaphoreSlim gate = new(1, 1);
    private IAccount? account;
    private readonly bool selectAccount;
    public string? AccountName => account?.Username;
    public BrowserAuthentication(ConnectionProfile profile, bool selectAccount = true)
    {
        this.selectAccount = selectAccount;
        if (!Guid.TryParse(profile.ClientId, out _) || string.IsNullOrWhiteSpace(profile.TenantId) ||
            profile.TenantId.Any(c => !(char.IsAsciiLetterOrDigit(c) || c is '.' or '-')))
            throw new OperationException("InvalidProfile");
        app = PublicClientApplicationBuilder.Create(profile.ClientId)
            .WithAuthority($"https://login.microsoftonline.com/{profile.TenantId}")
            .WithRedirectUri("http://localhost").Build();
    }
    public async Task<string> GetAsync(string resource, CancellationToken ct, bool refresh = false)
    {
        await gate.WaitAsync(ct);
        try
        {
            string[] scopes = [resource.TrimEnd('/') + "/.default"];
            try
            {
                if (account != null)
                    return (await app.AcquireTokenSilent(scopes, account).WithForceRefresh(refresh).ExecuteAsync(ct)).AccessToken;
            }
            catch (MsalUiRequiredException) { }
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
            timeout.CancelAfter(TimeSpan.FromMinutes(5));
            var request = app.AcquireTokenInteractive(scopes).WithUseEmbeddedWebView(false);
            if (account != null) request = request.WithAccount(account);
            else if (selectAccount) request = request.WithPrompt(Prompt.SelectAccount);
            var result = await request.ExecuteAsync(timeout.Token);
            account = result.Account;
            return result.AccessToken;
        }
        catch (MsalException ex) { throw new OperationException("AuthenticationFailed", null, ex); }
        finally { gate.Release(); }
    }
}
