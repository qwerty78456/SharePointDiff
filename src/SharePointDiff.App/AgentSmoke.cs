using System.Collections;
using System.Globalization;
using System.Resources;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using SharePointDiff.Core;

namespace SharePointDiff.App;

/// <summary>--demo --agent --ui-smoke: exercises the tray agent and its card on synthetic data. No network, registry writes or settings.</summary>
internal static class AgentSmoke
{
    private static async Task IdleAsync(Application app) => await app.Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
    private static void Require(bool condition, string what) { if (!condition) throw new InvalidOperationException(what); }
    public static async Task<bool> RunAsync(Application app, TrayAgent agent, string directory)
    {
        Directory.CreateDirectory(directory);
        var checks = new Dictionary<string, object>();
        try
        {
            Require(agent.Demo, "agent smoke requires --demo");
            string? runBefore = StartupRegistration.Read().Value; // read only
            bool settingsBefore = File.Exists(SettingsStore.PathName), cacheBefore = Directory.Exists(SettingsStore.TokenCacheDirectory);
            agent.StartAtLogon(); await IdleAsync(app);
            checks["TrayIconHasHandle"] = agent.Icon.Icon != null && agent.Icon.Icon.Handle != IntPtr.Zero; Require((bool)checks["TrayIconHasHandle"], "tray icon has no handle");
            checks["IconSize"] = agent.Icon.Icon!.Width; Require(agent.Icon.Icon.Width == TrayIcons.Size, "icon size differs from SM_CXSMICON");
            var outcome = agent.Pending ?? throw new InvalidOperationException("demo prompt missing"); var s = outcome.Summary;
            checks["Summary"] = new { s.NewFiles, s.Updates, s.LocallyChanged, s.Review };
            Require(s.NewFiles == 1000 && s.Updates == 1000 && s.LocallyChanged == 500 && s.Review == 1000, "demo summary counts");
            var plan = DiffService.SyncAll(outcome.Scan.Entries);
            checks["SyncAllItems"] = plan.Items.Count;
            Require(plan.Items.Count == 2000 && plan.Items.All(x => x.Replace ? x.Entry.IsDifference : x.Entry.Status == DiffStatus.MissingLocal), "SyncAll scope");
            var card = agent.Card ?? throw new InvalidOperationException("card not shown");
            Require(card.WarningLine != null, "locally-changed warning must show when K > 0"); Require(card.ReviewLine != null, "review line must show when Review > 0");
            Require(card.SyncButton is { IsEnabled: false }, "Sync must be disabled in demo"); Require(!agent.CheckItem.Enabled, "Check now must be disabled in demo");
            foreach (var (culture, file) in new[] { ("vi-VN", "card-vi.png"), ("en-US", "card-en.png") })
            {
                agent.L.Switch(culture); await IdleAsync(app);
                var area = SystemParameters.WorkArea;
                Require(card.Left >= area.Left && card.Top >= area.Top && card.Left + card.ActualWidth <= area.Right + 1 && card.Top + card.ActualHeight <= area.Bottom + 1, "card outside work area (" + culture + ")");
                UiSmoke.Save(card, Path.Combine(directory, file));
            }
            UiSmoke.Save(card, Path.Combine(directory, "card-en-150.png"), 1.5);
            checks["CardFitsWorkArea"] = true;
            checks["AllKeysResolved"] = MissingKeys().ToArray(); Require(MissingKeys().Count == 0, "unresolved resource keys");
            checks["TooltipsFit"] = TooltipsFit(agent.L); checks["BalloonsFit"] = BalloonsFit(agent.L);
            card.ShowProgress();
            card.UpdateProgress(new("Applying", 3, 10, "Engineering/" + string.Concat(Enumerable.Repeat("Thư mục dài tiếng Việt/", 10)) + "Báo cáo kiểm định.pdf"), agent.L);
            await IdleAsync(app);
            Require(card.Buttons.All(b => b.TranslatePoint(new Point(b.ActualWidth, 0), card).X <= ((FrameworkElement)card.Content).ActualWidth + 1), "progress buttons clipped by a long path");
            UiSmoke.Save(card, Path.Combine(directory, "card-progress-en.png"));
            agent.RenderForSmoke(new(AgentState.SignInRequired, null, "SignInRequired", null, TimeSpan.Zero, 1, TimeSpan.Zero)); await IdleAsync(app);
            Require(agent.Attention && agent.SignInItem.Available, "sign-in state must change the icon and show the menu item");
            UiSmoke.Save(agent.Card!, Path.Combine(directory, "card-signin-en.png"));
            agent.RenderForSmoke(new(AgentState.Failed, null, "RequestTimeout", null, TimeSpan.Zero, 1, TimeSpan.Zero)); await IdleAsync(app);
            UiSmoke.Save(agent.Card!, Path.Combine(directory, "card-problem-en.png"));
            var permissions = new[] { new TokenPermissionReport("https://graph.microsoft.com", true, ["AllSites.FullControl", "Sites.Read.All", "openid", "profile", "email"], ["AllSites.FullControl"]),
                new TokenPermissionReport("https://contoso.sharepoint.com", true, ["AllSites.FullControl", "Sites.Read.All"], ["AllSites.FullControl"]) };
            agent.RenderForSmoke(new(AgentState.NotConfigured, null, "WriteScopedApp", null, TimeSpan.Zero, 1, TimeSpan.Zero), permissions);
            foreach (string culture in new[] { "vi-VN", "en-US" })
            {
                agent.L.Switch(culture); await IdleAsync(app);
                string explanation = agent.ProblemText("WriteScopedApp");
                Require(explanation.Contains("AllSites.FullControl") && explanation.Contains("AllSites.Read") && explanation.Contains(agent.L["PermissionsRemove"]), "permission names or rejection reason missing");
                var permissionCard = agent.Card!; var area = SystemParameters.WorkArea;
                Require(permissionCard.Top >= area.Top && permissionCard.Top + permissionCard.ActualHeight <= area.Bottom + 1, "permission card exceeds the work area");
                UiSmoke.Save(permissionCard, Path.Combine(directory, "card-permissions-" + culture + ".png"));
            }
            checks["PermissionCardNamesScopesAndFits"] = true;
            var window = agent.OpenMain(outcome) ?? throw new InvalidOperationException("Details did not open a window"); await IdleAsync(app);
            var vm = (MainViewModel)window.DataContext;
            checks["DetailsRows"] = vm.Rows.Count; checks["DetailsMode"] = vm.ModeIndex; checks["DetailsPlan"] = vm.CurrentPlan.Items.Count;
            Require(vm.Rows.Count == 10000 && vm.ModeIndex == 0 && vm.CurrentPlan.Items.Count == 1000, "Details must open in Keep-local mode");
            // Only the agent's scan carries the locally-changed set (plain demo data has none): proves Details loaded it.
            vm.StatusFilterIndex = 6; int changed = vm.FilesView.Cast<object>().Count(); vm.StatusFilterIndex = 0;
            checks["DetailsLocallyChanged"] = changed; Require(changed == 500, "Details must show the agent's scan (locally changed rows)");
            vm.ModeIndex = 2; Require(vm.CurrentPlan.Items.Count == 2000, "Replace-all plan in Details");
            UiSmoke.Save(window, Path.Combine(directory, "details-en.png")); window.Close();
            Require(StartupRegistration.Read().Value == runBefore, "demo touched the Run key");
            Require(File.Exists(SettingsStore.PathName) == settingsBefore && Directory.Exists(SettingsStore.TokenCacheDirectory) == cacheBefore, "demo created settings or a token cache");
            checks["DemoDoesNotTouchRunKeyOrSettings"] = true;
            var icon = agent.Icon; agent.Dispose();
            Require(!icon.Visible, "tray icon still visible after dispose"); checks["IconDisposedOnExit"] = true;
            if (App.HasFatalError) throw new InvalidOperationException("Fatal UI error; agent smoke must fail");
            checks["Passed"] = true;
        }
        catch (Exception ex) { checks["Passed"] = false; checks["Error"] = ex.ToString(); }
        await JsonFiles.WriteAsync(Path.Combine(directory, "agent-checks.json"), checks);
        return checks["Passed"] is true;
    }
    private static List<string> MissingKeys()
    {
        var resources = new ResourceManager("SharePointDiff.App.Resources.Strings", typeof(LocalizationService).Assembly);
        var keys = resources.GetResourceSet(CultureInfo.InvariantCulture, true, true)!.Cast<DictionaryEntry>().Select(e => (string)e.Key).ToArray();
        return [.. from culture in new[] { "vi-VN", "en-US" } from key in keys
                   let value = resources.GetString(key, CultureInfo.GetCultureInfo(culture)) where string.IsNullOrWhiteSpace(value) select culture + ":" + key];
    }
    /// <summary>The full tooltip ("SharePoint Diff\n" + text) must fit Windows' 63-character limit without truncation, in both languages.</summary>
    private static bool TooltipsFit(LocalizationService l)
    {
        string original = l.CultureName; var big = 99_999; var time = new DateTime(2026, 12, 31, 23, 59, 0);
        try
        {
            foreach (string culture in new[] { "vi-VN", "en-US" })
            {
                l.Switch(culture);
                string[] texts = [l["TrayWaiting"], l["TrayCheckingStart"], l.Format("TrayChecking", big, big), l.Format("TrayUpToDate", time), l.Format("TrayReviewOnly", big),
                    l.Format("TrayPending", big, big), l["TraySyncing"], l.Format("TraySyncProgress", big, big), l["TrayOffline"], l["TraySignInNeeded"], l["TrayProblem"], l["TrayIdle"]];
                foreach (string text in texts) Require(("SharePoint Diff\n" + text).Length <= 63, "tooltip too long: " + text);
            }
            return true;
        }
        finally { l.Switch(original); }
    }
    private static bool BalloonsFit(LocalizationService l)
    {
        string original = l.CultureName;
        try
        {
            foreach (string culture in new[] { "vi-VN", "en-US" })
            {
                l.Switch(culture);
                foreach (string title in new[] { l["BalloonSignInTitle"], l["BalloonDoneTitle"] }) Require(title.Length <= 63, "balloon title too long: " + title);
                foreach (string text in new[] { l["BalloonSignInText"], l.Format("BalloonDone", 99_999, 99_999), l["BalloonCancelled"], l["BalloonStoppedForSignIn"] })
                    Require(text.Length <= 255, "balloon text too long: " + text);
            }
            return true;
        }
        finally { l.Switch(original); }
    }
}
