using System.Text.Json;

namespace SharePointDiff.Core;

public sealed class UserSettings
{
    public string Language { get; set; } = "vi-VN";
    public string SiteUrl { get; set; } = "";
    public string LocalRoot { get; set; } = "";
    public int SelectedProfile { get; set; }
    public List<ConnectionProfile> Profiles { get; set; } = [];
    /// <summary>Check this site and folder at Windows sign-in (the app manages its own HKCU Run entry).</summary>
    public bool StartWithWindows { get; set; }
    /// <summary>Seconds to wait after sign-in before the background check, 0–900.</summary>
    public int AgentDelaySeconds { get; set; } = 60;
    public bool IdleVerification { get; set; } = true;
    /// <summary>Repairs values that older or hand-edited files may hold (null lists and strings, out-of-range numbers).</summary>
    public UserSettings Normalize()
    {
        Profiles ??= []; Language ??= "vi-VN"; SiteUrl ??= ""; LocalRoot ??= "";
        // The window clamps the selection the same way, so the agent never disagrees with what the user sees.
        SelectedProfile = Profiles.Count == 0 ? 0 : Math.Clamp(SelectedProfile, 0, Profiles.Count - 1);
        AgentDelaySeconds = Math.Clamp(AgentDelaySeconds, 0, 900);
        return this;
    }
}

public static class SettingsStore
{
    private const string Variable = "SPDIFF_SETTINGS_PATH";
    /// <summary>True when SPDIFF_SETTINGS_PATH points somewhere else (tests, smoke runs); the startup entry is then never touched.</summary>
    public static bool IsRedirected => Environment.GetEnvironmentVariable(Variable) is { Length: > 0 };
    public static string PathName => Path.GetFullPath(Environment.GetEnvironmentVariable(Variable) is { Length: > 0 } custom ? custom
        : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SharePointDiff", "settings.json"));
    public static string Folder => Path.GetDirectoryName(PathName)!;
    /// <summary>DPAPI-encrypted MSAL caches live next to the settings, so SPDIFF_SETTINGS_PATH redirects them too.</summary>
    public static string TokenCacheDirectory => Path.Combine(Folder, "msal");
    public static string LogDirectory => Path.Combine(Folder, "logs");
    /// <summary>The main window's load: existing settings, else first-run defaults imported (read-only) from a sibling sp_chuanhoa configuration.</summary>
    public static UserSettings Load(string? importRoot = null)
    {
        try
        {
            var existing = JsonFiles.Read<UserSettings>(PathName); if (existing != null) return existing.Normalize();
            // Import only connection identifiers and folder defaults, never passwords or project content.
            var dir = new DirectoryInfo(importRoot ?? AppContext.BaseDirectory);
            for (int i = 0; i < 8 && dir != null; i++, dir = dir.Parent)
            {
                string config = Path.Combine(dir.FullName, "sp_chuanhoa", "cauhinh", "cauhinh.json");
                if (!File.Exists(config)) continue;
                using var json = JsonDocument.Parse(File.ReadAllText(config)); var r = json.RootElement;
                return new()
                {
                    SiteUrl = r.GetProperty("SiteThat").GetString() ?? "",
                    LocalRoot = r.GetProperty("ThuMucMau").GetString() ?? "",
                    Profiles = [new("IVD", r.GetProperty("Tenant").GetString()!, r.GetProperty("ClientId").GetString()!)]
                };
            }
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        { throw new OperationException("SettingsRead", PathName, ex); }
        return new();
    }
    /// <summary>The agent's load: never imports anything; null when the user has not saved settings yet.</summary>
    public static UserSettings? LoadExisting()
    {
        try { return JsonFiles.Read<UserSettings>(PathName)?.Normalize(); }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        { throw new OperationException("SettingsRead", PathName, ex); }
    }
    public static Task SaveAsync(UserSettings settings) => JsonFiles.WriteAsync(PathName, settings);
}
