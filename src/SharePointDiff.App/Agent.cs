using System.Diagnostics;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Threading;
using Microsoft.Win32;
using SharePointDiff.Core;
using Drawing = System.Drawing;
using Forms = System.Windows.Forms;

namespace SharePointDiff.App;

/// <summary>One owner process per settings path. A later launch without --agent asks the owner to show its window; a later --agent just exits.</summary>
internal sealed class SingleInstance : IDisposable
{
    private readonly Mutex mutex;
    private readonly EventWaitHandle activate;
    private RegisteredWaitHandle? wait;
    private SingleInstance(Mutex mutex, EventWaitHandle activate) { this.mutex = mutex; this.activate = activate; }
    private static string Name => @"Local\SharePointDiff-" + JsonFiles.StableId(SettingsStore.PathName.ToLowerInvariant());
    public static SingleInstance? TryClaim(bool signalOwner)
    {
        var mutex = new Mutex(true, Name, out bool created);
        var activate = new EventWaitHandle(false, EventResetMode.AutoReset, Name + "-activate");
        if (created) return new(mutex, activate);
        if (signalOwner) { AllowSetForegroundWindow(-1); activate.Set(); }
        activate.Dispose(); mutex.Dispose();
        return null;
    }
    public void OnActivate(Dispatcher dispatcher, Action show) =>
        wait = ThreadPool.RegisterWaitForSingleObject(activate, (_, _) => dispatcher.BeginInvoke(() => { if (!App.HasFatalError && !dispatcher.HasShutdownStarted) show(); }), null, Timeout.Infinite, false);
    public void Dispose()
    {
        wait?.Unregister(null); activate.Dispose();
        try { mutex.ReleaseMutex(); } catch (ApplicationException) { } // not owned by this thread: the handle is closed anyway
        mutex.Dispose();
    }
    [DllImport("user32.dll")] private static extern bool AllowSetForegroundWindow(int processId);
}

/// <summary>The per-user HKCU Run value. Never written in demo or when SPDIFF_SETTINGS_PATH redirects the settings.</summary>
internal static class StartupRegistration
{
    private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run", ApprovedKey = @"Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run", Name = "SharePointDiff";
    public static string Exe => Environment.ProcessPath!;
    public sealed record State(string? Value, bool TargetExists, bool IsCurrent, bool DisabledByUser);
    public static State Read()
    {
        try
        {
            using var run = Registry.CurrentUser.OpenSubKey(RunKey);
            string? value = run?.GetValue(Name) as string;
            bool exists = StartupCommand.TryParse(value, out string exe, out _) && File.Exists(exe);
            // Task Manager's "Disabled" is stored here (first byte odd); read only, never written by this app.
            using var approved = Registry.CurrentUser.OpenSubKey(ApprovedKey);
            bool disabled = approved?.GetValue(Name) is byte[] { Length: > 0 } flags && (flags[0] & 1) == 1;
            return new(value, exists, StartupCommand.IsCurrent(value, Exe), disabled);
        }
        catch (Exception ex) when (ex is System.Security.SecurityException or UnauthorizedAccessException or IOException)
        { throw new OperationException("StartupRegistrationFailed", RunKey, ex); }
    }
    public static void Apply(bool enable)
    {
        if (SettingsStore.IsRedirected) throw new OperationException("StartupRegistrationFailed", "SPDIFF_SETTINGS_PATH");
        try
        {
            using var run = Registry.CurrentUser.CreateSubKey(RunKey);
            if (enable) run.SetValue(Name, StartupCommand.Build(Exe)); else run.DeleteValue(Name, false);
        }
        catch (Exception ex) when (ex is System.Security.SecurityException or UnauthorizedAccessException or IOException)
        { throw new OperationException("StartupRegistrationFailed", RunKey, ex); }
    }
    /// <summary>Removes the entry only when it points at this executable; another copy's entry is left alone.</summary>
    public static void RemoveIfCurrent() { if (!SettingsStore.IsRedirected && Read().IsCurrent) Apply(false); }
}

/// <summary>Background (low I/O and CPU priority) mode for the sign-in scan only; ended as soon as the user opens a window.</summary>
internal sealed class ProcessMode : IDisposable
{
    private bool active;
    public ProcessMode() => active = SetPriorityClass(GetCurrentProcess(), 0x00100000); // PROCESS_MODE_BACKGROUND_BEGIN; FALSE = keep normal priority
    public void End() { if (active) { SetPriorityClass(GetCurrentProcess(), 0x00200000); active = false; } }
    public void Dispose() => End();
    [DllImport("kernel32.dll")] private static extern IntPtr GetCurrentProcess();
    [DllImport("kernel32.dll")] private static extern bool SetPriorityClass(IntPtr process, uint priority);
}

/// <summary>Tray icons drawn at run time at the small-icon size of the current DPI (no binary assets).</summary>
internal static class TrayIcons
{
    public static int Size { get { int size = GetSystemMetricsForDpi(49, GetDpiForSystem()); return size > 0 ? size : 16; } } // SM_CXSMICON
    public static Drawing.Icon Create(bool attention)
    {
        int size = Size;
        using var bitmap = new Drawing.Bitmap(size, size);
        using (var g = Drawing.Graphics.FromImage(bitmap))
        {
            g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias; g.Clear(Drawing.Color.Transparent);
            using var fill = new Drawing.SolidBrush(attention ? Drawing.Color.FromArgb(0xB5, 0x47, 0x08) : Drawing.Color.FromArgb(0x12, 0x6A, 0x70));
            g.FillEllipse(fill, 0.5f, 0.5f, size - 1.5f, size - 1.5f);
            using var pen = new Drawing.Pen(Drawing.Color.White, Math.Max(1.5f, size / 9f)) { StartCap = Drawing.Drawing2D.LineCap.Round, EndCap = Drawing.Drawing2D.LineCap.Round };
            float c = size / 2f, top = size * 0.24f, bottom = size * 0.72f, wing = size * 0.2f;
            g.DrawLine(pen, c, top, c, bottom); g.DrawLine(pen, c - wing, bottom - wing, c, bottom); g.DrawLine(pen, c + wing, bottom - wing, c, bottom);
        }
        IntPtr handle = bitmap.GetHicon();
        try { using var temporary = Drawing.Icon.FromHandle(handle); return (Drawing.Icon)temporary.Clone(); }
        finally { DestroyIcon(handle); }
    }
    [DllImport("user32.dll")] private static extern int GetSystemMetricsForDpi(int index, uint dpi);
    [DllImport("user32.dll")] private static extern uint GetDpiForSystem();
    [DllImport("user32.dll")] private static extern bool DestroyIcon(IntPtr handle);
}

/// <summary>The tray agent: checks once after Windows sign-in (and on "Check now"), asks before writing anything to the mirror.
/// WinForms handlers only post to the WPF dispatcher; expected failures end on a card, anything else is a programming fault.</summary>
internal sealed class TrayAgent : IDisposable
{
    private const string Graph = "https://graph.microsoft.com";
    private readonly Application app;
    private readonly bool demo;
    private readonly Forms.NotifyIcon icon = new();
    private readonly Forms.ContextMenuStrip menu = new();
    private readonly Forms.ToolStripMenuItem open = new(), check = new(), signIn = new(), exit = new();
    private readonly Forms.ToolStripMenuItem faultForSmoke = new() { Available = false };
    /// <summary>--fatal-smoke tray: a real WinForms click whose handler throws, to prove tray faults reach the fatal handler.</summary>
    internal void ClickFaultForSmoke() { faultForSmoke.Available = true; faultForSmoke.PerformClick(); }
    private readonly HttpClient probeClient = new() { Timeout = Timeout.InfiniteTimeSpan };
    private Drawing.Icon normalIcon, attentionIcon;
    private CancellationTokenSource? delay, work;
    private UserSettings? settings;
    private BrowserAuthentication? session;
    private SharePointReader? reader;
    private ProcessMode? background;
    private Action? balloonAction;
    private bool busy, disposed, attention, syncing, logonRun;
    private long lastTooltip;
    public LocalizationService L { get; } = new();
    internal AgentCard? Card { get; private set; }
    internal CheckOutcome? Pending { get; private set; }
    internal AgentState State { get; private set; } = AgentState.UpToDate;
    internal Forms.NotifyIcon Icon => icon;
    internal bool Attention => attention;
    internal bool Demo => demo;
    internal bool IsBusy => busy;
    private IReadOnlyList<TokenPermissionReport>? demoPermissions;
    internal Forms.ToolStripMenuItem SignInItem => signIn;
    internal Forms.ToolStripMenuItem CheckItem => check;
    public TrayAgent(Application app, bool demo)
    {
        this.app = app; this.demo = demo;
        Forms.Application.SetUnhandledExceptionMode(Forms.UnhandledExceptionMode.ThrowException);
        normalIcon = TrayIcons.Create(false); attentionIcon = TrayIcons.Create(true);
        menu.Items.AddRange([open, check, signIn, new Forms.ToolStripSeparator(), exit, faultForSmoke]);
        open.Click += (_, _) => Post(() => OpenMain(Pending));
        check.Click += (_, _) => Post(() => Fire(CheckNowAsync));
        signIn.Click += (_, _) => Post(() => Fire(SignInAsync));
        exit.Click += (_, _) => Post(() => Fire(ExitAsync));
        faultForSmoke.Click += (_, _) => Post(() => Fire(() => throw new InvalidOperationException("FATAL_SMOKE: deliberate tray fault; the application must terminate with exit code 1.")));
        icon.ContextMenuStrip = menu; icon.Icon = normalIcon; icon.Text = "SharePoint Diff";
        icon.MouseClick += (_, e) => { if (e.Button == Forms.MouseButtons.Left) Post(LeftClick); };
        icon.BalloonTipClicked += (_, _) => Post(() => { var action = balloonAction; balloonAction = null; action?.Invoke(); });
        L.PropertyChanged += (_, _) => Relabel();
        SystemEvents.DisplaySettingsChanged += OnDisplayChanged;
        Relabel(); icon.Visible = true;
    }
    private void Post(Action action) { if (!disposed) app.Dispatcher.BeginInvoke(action); }
    /// <summary>Runs an async handler on the dispatcher; an exception that escapes is a programming fault and reaches the fatal handler.</summary>
    private static async void Fire(Func<Task> action) => await action();
    private void OnDisplayChanged(object? sender, EventArgs e) => Post(() =>
    {
        if (disposed) return;
        var oldNormal = normalIcon; var oldAttention = attentionIcon;
        normalIcon = TrayIcons.Create(false); attentionIcon = TrayIcons.Create(true); icon.Icon = attention ? attentionIcon : normalIcon;
        oldNormal.Dispose(); oldAttention.Dispose();
    });
    private void Relabel()
    {
        if (disposed) return;
        open.Text = L["TrayOpen"]; check.Text = L["TrayCheckNow"]; signIn.Text = L["TraySignIn"]; exit.Text = L["TrayExit"];
        signIn.Visible = State == AgentState.SignInRequired; check.Enabled = !demo; signIn.Enabled = !demo;
        Card?.Relabel();
    }
    private void SetIcon(bool needsAttention) { if (disposed) return; attention = needsAttention; icon.Icon = needsAttention ? attentionIcon : normalIcon; }
    internal static string Fit(string text, int max) => text.Length <= max ? text : text[..(max - 1)] + "…";
    private void SetTooltip(string text) { if (!disposed) icon.Text = Fit("SharePoint Diff\n" + text, 63); }
    private void Balloon(string title, string text, Forms.ToolTipIcon kind, Action? onClick)
    {
        if (disposed) return;
        balloonAction = onClick;
        icon.ShowBalloonTip(10_000, Fit(title, 63), Fit(text, 255), kind);
    }

    // ---- Startup and checks -----------------------------------------------------------------------
    public void StartAtLogon()
    {
        if (demo) { ShowDemo(); return; }
        try { settings = SettingsStore.LoadExisting(); }
        catch (OperationException ex) { TrySave(ex, "AgentStart"); State = AgentState.Failed; SetIcon(true); ShowProblem(ex.Code); return; }
        if (settings == null || !settings.StartWithWindows)
        {
            // Turned off, or never set up: remove our own startup entry and leave quietly.
            try { StartupRegistration.RemoveIfCurrent(); } catch (OperationException ex) { TrySave(ex, "AgentStart"); }
            Dispose(); app.Shutdown(0); return;
        }
        L.Switch(settings.Language); SetTooltip(L["TrayWaiting"]);
        var wait = delay = new CancellationTokenSource(); int seconds = settings.AgentDelaySeconds;
        Fire(async () =>
        {
            try { await Task.Delay(TimeSpan.FromSeconds(seconds), wait.Token); }
            catch (OperationCanceledException) { return; } // "Check now" or Exit took over
            await CheckAsync(false);
        });
    }
    /// <summary>Demo mode never reads settings, the registry or the network: it shows the synthetic prompt.</summary>
    internal void ShowDemo()
    {
        var outcome = DemoData.Outcome();
        Render(new AgentCheckResult(AgentState.Pending, outcome, null, null, TimeSpan.Zero, 1, TimeSpan.Zero));
    }
    /// <summary>"Check now" never waits behind the logon run (which may be waiting for the network or backing off): it stops it and checks at once.</summary>
    private async Task CheckNowAsync()
    {
        delay?.Cancel();
        if (busy && logonRun && !syncing) { work?.Cancel(); for (int i = 0; busy && i < 600; i++) await Task.Delay(50); }
        await CheckAsync(true);
    }
    private void EnsureSession()
    {
        settings = SettingsStore.LoadExisting() ?? throw new OperationException("AgentNotReady", SettingsStore.PathName);
        L.Switch(settings.Language);
        var profile = settings.Profiles.ElementAtOrDefault(settings.SelectedProfile) ?? throw new OperationException("InvalidProfile");
        var shared = BrowserAuthentication.Shared(profile, SettingsStore.TokenCacheDirectory);
        if (!ReferenceEquals(shared, session) || reader == null) { reader?.Dispose(); reader = new SharePointReader(shared.Background); session = shared; }
    }
    private async Task CheckAsync(bool userInitiated)
    {
        if (busy || disposed || demo) return;
        busy = true; logonRun = !userInitiated; var source = work = new CancellationTokenSource(); var token = source.Token;
        try
        {
            if ((app as App)?.IdleChecks is { } idle) await idle.PauseAsync();
            EnsureSession();
            var runner = new AgentCheckRunner(reader!, new NetworkAvailabilityProbe(), TimeProvider.System);
            var currentSettings = settings ?? throw new OperationException("AgentNotReady", SettingsStore.PathName);
            var progress = Progress(); string site = currentSettings.SiteUrl, root = currentSettings.LocalRoot;
            SetTooltip(L["TrayCheckingStart"]);
            // Background priority only while nobody is working in the main window.
            if (!userInitiated && !app.Windows.OfType<MainWindow>().Any()) background = new ProcessMode();
            AgentCheckResult result;
            try { result = await Task.Run(() => runner.RunAsync(site, root, userInitiated, progress, token), token); }
            finally { background?.Dispose(); background = null; }
            if (disposed) return;
            if (session!.PersistenceBlocked) { WriteLog(result with { State = AgentState.NotConfigured, Code = "WriteScopedApp" }); State = AgentState.NotConfigured; SetIcon(true); SetTooltip(L["TrayProblem"]); ShowProblem("WriteScopedApp"); }
            else { WriteLog(result); Render(result); }
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { SetTooltip(L["TrayIdle"]); }
        catch (Exception ex) when (ex is OperationException or IOException or UnauthorizedAccessException or HttpRequestException)
        { TrySave(ex, "AgentCheck"); if (disposed) return; State = AgentState.Failed; SetIcon(true); SetTooltip(L["TrayProblem"]); ShowProblem(ex is OperationException op ? op.Code : ex is HttpRequestException ? "HttpError" : "FileIoError"); }
        finally { busy = false; logonRun = false; Relabel(); }
    }
    internal void RenderForSmoke(AgentCheckResult result, IReadOnlyList<TokenPermissionReport>? permissions = null)
    {
        if (!demo) throw new InvalidOperationException("Synthetic permission reports require demo mode.");
        demoPermissions = permissions; Render(result);
    }
    private void Render(AgentCheckResult result)
    {
        if (disposed) return;
        State = result.State;
        switch (result.State)
        {
            case AgentState.Pending:
                Pending = result.Outcome; SetIcon(true);
                SetTooltip(L.Format("TrayPending", result.Outcome!.Summary.NewFiles, result.Outcome.Summary.Updates));
                ShowPrompt(result.Outcome);
                break;
            case AgentState.UpToDate: Pending = null; SetIcon(false); SetTooltip(L.Format("TrayUpToDate", DateTime.Now)); CloseCard(); break;
            case AgentState.ReviewOnly: Pending = null; SetIcon(false); SetTooltip(L.Format("TrayReviewOnly", result.Outcome!.Summary.Review)); CloseCard(); break;
            case AgentState.Offline: SetIcon(false); SetTooltip(L["TrayOffline"]); break;
            case AgentState.SignInRequired:
                SetIcon(true); SetTooltip(L["TraySignInNeeded"]); ShowSignIn();
                Balloon(L["BalloonSignInTitle"], L["BalloonSignInText"], Forms.ToolTipIcon.Warning, () => Fire(SignInAsync));
                break;
            default:
                SetIcon(true); SetTooltip(L["TrayProblem"]);
                ShowProblem(result.Code ?? (result.State == AgentState.Incomplete ? "IncompleteInventory" : "Unexpected"));
                break;
        }
        Relabel();
    }
    private async Task SignInAsync()
    {
        if (busy || disposed || demo) return;
        busy = true; var source = work = new CancellationTokenSource(); var token = source.Token;
        try
        {
            EnsureSession();
            string authority = new Uri(DiffService.NormalizeSite(settings!.SiteUrl)).GetLeftPart(UriPartial.Authority);
            await session!.SignInAsync([Graph, authority], token); // the only interactive path; always a user's click
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { busy = false; return; }
        catch (Exception ex) when (ex is OperationException or IOException or UnauthorizedAccessException or HttpRequestException)
        { TrySave(ex, "AgentSignIn"); busy = false; ShowProblem(ex is OperationException op ? op.Code : "AuthenticationFailed"); return; }
        busy = false;
        await CheckAsync(true);
    }
    private async Task SyncAsync()
    {
        var outcome = Pending;
        if (busy || disposed || demo || outcome == null || reader == null) return;
        busy = true; syncing = true; var source = work = new CancellationTokenSource(); var token = source.Token;
        SetThreadExecutionState(EsContinuous | EsSystemRequired); // no sleep while files are being written
        try
        {
            if ((app as App)?.IdleChecks is { } idle) await idle.PauseAsync();
            var plan = DiffService.SyncAll(outcome.Scan.Entries);
            await ReportWriter.WriteScanAsync(new Workspace(outcome.Scan.SiteUrl, outcome.Scan.LocalRoot), outcome.Scan, token);
            Card?.ShowProgress(); SetTooltip(L["TraySyncing"]);
            var progress = Progress();
            var result = await Task.Run(() => new SyncEngine(reader).ApplyAsync(outcome.Scan, plan, progress, token), token);
            int done = result.Results.Count(r => r.Status is "Verified" or "DownloadedUnverified"), failed = result.Results.Count(r => r.Status == "Error");
            await File.WriteAllTextAsync(Path.Combine(result.RunDirectory, "summary.txt"),
                L.Format("ReportSummary", done, failed, result.Results.Sum(x => x.BytesDownloaded), result.RunDirectory), CancellationToken.None);
            Pending = null;
            if (disposed) return;
            if (result.StopReason == "SignInRequired")
            {
                State = AgentState.SignInRequired; SetIcon(true); SetTooltip(L["TraySignInNeeded"]); ShowSignIn();
                Balloon(L["BalloonSignInTitle"], L["BalloonStoppedForSignIn"], Forms.ToolTipIcon.Warning, () => Fire(SignInAsync));
            }
            else
            {
                // Only a complete run without errors is "up to date"; otherwise ask for another check.
                bool complete = !result.Cancelled && failed == 0;
                State = complete ? AgentState.UpToDate : AgentState.Incomplete; SetIcon(!complete); CloseCard();
                SetTooltip(complete ? L.Format("TrayUpToDate", DateTime.Now) : L["TraySyncIncomplete"]);
                string run = result.RunDirectory;
                Balloon(L["BalloonDoneTitle"], result.Cancelled ? L["BalloonCancelled"] : L.Format("BalloonDone", done, failed),
                    failed > 0 ? Forms.ToolTipIcon.Warning : Forms.ToolTipIcon.Info, () => OpenFolder(run));
            }
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { CloseCard(); SetTooltip(L["TrayIdle"]); }
        catch (Exception ex) when (ex is OperationException or IOException or UnauthorizedAccessException or HttpRequestException)
        { TrySave(ex, "AgentSync"); if (disposed) return; SetIcon(true); SetTooltip(L["TrayProblem"]); ShowProblem(ex is OperationException op ? op.Code : "FileIoError"); }
        finally { SetThreadExecutionState(EsContinuous); busy = false; syncing = false; Relabel(); }
    }
    private IProgress<OperationProgress> Progress() => new Progress<OperationProgress>(p =>
    {
        Card?.UpdateProgress(p, L);
        long now = Environment.TickCount64; if (now - lastTooltip < 1000) return; lastTooltip = now;
        SetTooltip(p.Code == "Hashing" ? L.Format("TrayChecking", p.Completed, p.Total) : p.Code == "Applying" ? L.Format("TraySyncProgress", p.Completed, p.Total) : L["TrayCheckingStart"]);
    });

    // ---- Card and window ----------------------------------------------------------------------------
    private AgentCard EnsureCard() => Card ??= new AgentCard(this);
    private void ShowPrompt(CheckOutcome outcome) { if (!disposed) EnsureCard().ShowPrompt(outcome, demo); }
    private void ShowSignIn() { if (!disposed) EnsureCard().ShowSignIn(); }
    internal string ProblemText(string code) => code == "WriteScopedApp" ? PermissionMessages.Format(L, demo ? demoPermissions ?? [] : session?.PermissionReports ?? []) : L["Error" + code];
    private void ShowProblem(string code)
    {
        if (disposed) return;
        if (code == "WriteScopedApp" && !demo)
        {
            try { PermissionMessages.Save(L, session?.PermissionReports ?? []); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { TrySave(ex, "PermissionLog"); }
        }
        EnsureCard().ShowProblem(code);
    }
    internal void CloseCard() { Card?.Close(); Card = null; }
    internal void CardClosed(AgentCard card) { if (ReferenceEquals(card, Card)) Card = null; }
    internal void CardSync() => Fire(SyncAsync);
    internal void CardStop() => work?.Cancel();
    internal void CardSignIn() => Fire(SignInAsync);
    internal void CardRetry() => Fire(CheckNowAsync);
    private void LeftClick()
    {
        if (syncing) EnsureCard().ShowProgress(reset: false); // never the stale prompt while files are being written
        else if (State == AgentState.SignInRequired && !demo) Fire(SignInAsync);
        else if (Pending != null) ShowPrompt(Pending);
        else OpenMain(null);
    }
    /// <summary>Shows the main window, with the agent's scan in Keep-local mode when there is one and the window is idle.
    /// Returns null when the settings cannot be read (shown on the Problem card).</summary>
    public MainWindow? OpenMain(CheckOutcome? outcome)
    {
        background?.End();
        var window = app.Windows.OfType<MainWindow>().FirstOrDefault();
        try
        {
            if (window == null)
            {
                var vm = new MainViewModel(demo);
                vm.WorkspaceChanged += OnWorkspaceChanged;
                if (outcome != null) vm.LoadAgentScan(outcome, demo || session == null ? null : new SharePointReader(session.Interactive));
                window = new MainWindow(vm);
            }
            else if (outcome != null && window.DataContext is MainViewModel { Busy: false } open)
                open.LoadAgentScan(outcome, demo || session == null ? null : new SharePointReader(session.Interactive));
        }
        catch (Exception ex) when (ex is OperationException or IOException or UnauthorizedAccessException)
        { TrySave(ex, "AgentOpen"); ShowProblem(ex is OperationException op ? op.Code : "FileIoError"); return null; }
        window.Show();
        if (window.WindowState == WindowState.Minimized) window.WindowState = WindowState.Normal;
        window.Activate();
        return window;
    }
    /// <summary>A Compare or Apply in the main window makes the agent's prompt out of date: drop it rather than sync a stale scan.</summary>
    private void OnWorkspaceChanged()
    {
        if (syncing || disposed || Pending == null) return;
        Pending = null; State = AgentState.UpToDate; SetIcon(false); SetTooltip(L["TrayIdle"]);
        if (Card != null && !syncing) CloseCard();
        Relabel();
    }
    private static void OpenFolder(string path) { if (Directory.Exists(path)) Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }

    // ---- Exit, sign-out and logs ----------------------------------------------------------------------
    private void CancelEverything()
    {
        delay?.Cancel(); work?.Cancel();
        foreach (var vm in app.Windows.OfType<MainWindow>().Select(w => w.DataContext).OfType<MainViewModel>()) if (vm.Busy) vm.CancelCommand.Execute(null);
    }
    private async Task ExitAsync()
    {
        CancelEverything();
        var deadline = DateTime.UtcNow.AddSeconds(30);
        while ((busy || app.Windows.OfType<MainWindow>().Any(w => (w.DataContext as MainViewModel)?.Busy == true)) && DateTime.UtcNow < deadline) await Task.Delay(100);
        CommitGate.WaitIdle(TimeSpan.FromSeconds(4));
        Dispose(); app.Shutdown(0);
    }
    /// <summary>Windows is signing out: stop, give committed replacements a few seconds to verify and record, then remove the icon.</summary>
    public void OnSessionEnding()
    {
        CancelEverything();
        CommitGate.WaitIdle(TimeSpan.FromSeconds(4));
        Dispose();
    }
    private void WriteLog(AgentCheckResult result)
    {
        try
        {
            string dir = SettingsStore.LogDirectory; Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, "agent-" + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + ".json");
            File.WriteAllText(path, JsonSerializer.Serialize(new
            {
                result.State, result.Code, result.Path, DurationSeconds = result.Duration.TotalSeconds, NetworkWaitSeconds = result.NetworkWait.TotalSeconds, result.Attempts,
                HashedFiles = result.Outcome?.Scan.Entries.Count(e => e.Local != null), result.Outcome?.Summary,
                GrantedPermissions = session?.PermissionReports,
                AllowedPermissions = TokenScopes.AllowedScopes,
                PermissionExplanation = session?.PersistenceBlocked == true ? PermissionMessages.Format(L, session.PermissionReports) : null
            }, JsonFiles.Options));
            // This tool's own check logs, not user data or backups: keep the newest 60.
            foreach (var old in new DirectoryInfo(dir).GetFiles("agent-*.json").OrderByDescending(f => f.Name, StringComparer.Ordinal).Skip(60)) old.Delete();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }
    private static void TrySave(Exception ex, string operation)
    {
        try { ErrorDetails.Save(ex, operation, (ex as OperationException)?.ItemPath); }
        catch (Exception logError) when (logError is IOException or UnauthorizedAccessException) { }
    }
    public void Dispose()
    {
        if (disposed) return; disposed = true;
        SystemEvents.DisplaySettingsChanged -= OnDisplayChanged;
        delay?.Cancel(); work?.Cancel();
        icon.Visible = false; icon.Dispose(); menu.Dispose();
        normalIcon.Dispose(); attentionIcon.Dispose(); probeClient.Dispose(); reader?.Dispose();
        var card = Card; Card = null; card?.Close();
    }
    private const uint EsContinuous = 0x80000000, EsSystemRequired = 0x00000001;
    [DllImport("kernel32.dll")] private static extern uint SetThreadExecutionState(uint flags);
}
