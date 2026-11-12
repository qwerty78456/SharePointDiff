using System.Text.Json;
using SharePointDiff.Core;

namespace SharePointDiff.App;

internal static class PermissionMessages
{
    public static string Format(LocalizationService l, IReadOnlyList<TokenPermissionReport> reports)
    {
        var lines = new List<string> { l["PermissionsCurrent"] };
        if (reports.Count == 0) lines.Add(l["PermissionsNotAvailable"]);
        foreach (var report in reports.OrderBy(r => r.Resource.Contains("graph.microsoft.com", StringComparison.OrdinalIgnoreCase) ? 0 : 1))
        {
            string resource = report.Resource.Contains("graph.microsoft.com", StringComparison.OrdinalIgnoreCase) ? "Microsoft Graph" :
                Uri.TryCreate(report.Resource, UriKind.Absolute, out var uri) && uri.Host.EndsWith(".sharepoint.com", StringComparison.OrdinalIgnoreCase) ? "SharePoint" : report.Resource;
            lines.Add("• " + resource + ": " + (report.ScopesReadable ? string.Join(", ", report.GrantedScopes) : l["PermissionsNotAvailable"]));
        }
        lines.Add(""); lines.Add(l["PermissionsRequired"]);
        var rejected = reports.SelectMany(r => r.RejectedScopes).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s, StringComparer.OrdinalIgnoreCase).ToArray();
        if (rejected.Length > 0)
        {
            lines.Add(""); lines.Add(l["PermissionsRemove"]);
            foreach (string scope in rejected)
                lines.Add("• " + scope + ": " + l[scope.Equals("AllSites.FullControl", StringComparison.OrdinalIgnoreCase) || scope.Equals("Sites.FullControl.All", StringComparison.OrdinalIgnoreCase) ? "PermissionsFullControl" :
                    scope.Contains("Write", StringComparison.OrdinalIgnoreCase) || scope.Contains("Manage", StringComparison.OrdinalIgnoreCase) ? "PermissionsWrite" : "PermissionsUnknown"]);
        }
        if (reports.Any(r => !r.ScopesReadable)) { lines.Add(""); lines.Add(l["PermissionsCannotVerify"]); }
        lines.Add(""); lines.Add(l["PermissionsSessionOnly"]);
        return string.Join(Environment.NewLine, lines);
    }

    public static object Audit(IReadOnlyList<TokenPermissionReport> reports) => new
    {
        Code = "WriteScopedApp", GrantedPermissions = reports, AllowedPermissions = TokenScopes.AllowedScopes,
        RequiredGraphPermission = "Sites.Read.All", RequiredSharePointPermission = "AllSites.Read", SignInSaved = false
    };

    public static string Save(LocalizationService l, IReadOnlyList<TokenPermissionReport> reports)
    {
        Directory.CreateDirectory(SettingsStore.LogDirectory);
        string path = Path.Combine(SettingsStore.LogDirectory, "permissions-" + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + ".txt");
        File.WriteAllText(path, Format(l, reports) + Environment.NewLine + Environment.NewLine + JsonSerializer.Serialize(Audit(reports), JsonFiles.Options));
        return path;
    }
}
