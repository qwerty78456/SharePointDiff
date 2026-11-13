using System.Windows;
using SharePointDiff.Core;

namespace SharePointDiff.App;

public partial class App : Application
{
    internal static bool HasFatalError { get; private set; }
    internal TrayAgent? Agent { get; private set; }
    internal IdleVerifier? IdleChecks { get; private set; }
    private SingleInstance? instance;
    private MainWindow? FirstMainWindow => Windows.OfType<MainWindow>().FirstOrDefault();
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        bool demo = e.Args.Contains("--demo"), agent = e.Args.Contains(StartupCommand.Flag);
        int smoke = Array.IndexOf(e.Args, "--ui-smoke"), fatal = Array.IndexOf(e.Args, "--fatal-smoke");
        string? smokeDirectory = smoke >= 0 && smoke + 1 < e.Args.Length ? e.Args[smoke + 1] : null;
        bool fatalSmoke = fatal >= 0, trayFault = fatalSmoke && fatal + 1 < e.Args.Length && e.Args[fatal + 1] == "tray";
        // Last resort for exceptions outside the dispatcher: record them; the process still terminates.
        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            try { if (args.ExceptionObject is Exception ex) ErrorDetails.Save(ex, "Unhandled"); }
            catch (Exception logError) when (logError is IOException or UnauthorizedAccessException) { }
        };
        DispatcherUnhandledException += (_, args) =>
        {
            if (HasFatalError) { Environment.Exit(1); return; }
            HasFatalError = true;
            // Remove the tray icon first so no ghost icon is left behind.
            Agent?.Dispose();
            var main = FirstMainWindow; var vm = main?.DataContext as MainViewModel;
            vm?.Dispose();
            if (main != null) main.IsEnabled = false;
            var l = vm?.L ?? Agent?.L ?? new LocalizationService();
            string operation = (string?)args.Exception.Data["Operation"] ?? vm?.StatusText ?? l["StartingApplication"];
            string? path = (string?)args.Exception.Data["Path"] ?? vm?.ProgressPath;
            string? log = null;
            try { log = ErrorDetails.Save(args.Exception, operation, path); }
            catch (Exception logError) when (logError is IOException or UnauthorizedAccessException)
            { log = l["LogWriteFailed"] + ": " + logError.Message; }
            if (!fatalSmoke) new ErrorDialog(l, args.Exception, operation, path, log) { Topmost = Agent != null }.ShowDialog();
            // Mark handled only to replace the generic OS dialog with diagnostics, then terminate.
            // An unexpected exception must never return to normal application execution.
            args.Handled = true;
            Shutdown(1);
        };
        if (!demo)
        {
            instance = SingleInstance.TryClaim(signalOwner: !agent);
            if (instance == null) { Shutdown(0); return; }
            IdleChecks = new IdleVerifier(this);
        }
        if (agent)
        {
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            var tray = Agent = new TrayAgent(this, demo);
            SessionEnding += (_, _) => tray.OnSessionEnding();
            instance?.OnActivate(Dispatcher, () => tray.OpenMain(tray.Pending));
            // The agent smoke is a demo-only harness (it must never read real settings, the Run key or the network).
            if (smokeDirectory != null && !demo) { tray.Dispose(); Shutdown(3); return; }
            // async void: a fault in the harness reaches the fatal handler instead of hanging the run.
            if (smokeDirectory != null) Dispatcher.BeginInvoke(new Action(async () => { bool passed = await AgentSmoke.RunAsync(this, tray, smokeDirectory); Shutdown(passed && !HasFatalError ? 0 : 3); }));
            else if (trayFault) Dispatcher.BeginInvoke(tray.ClickFaultForSmoke);
            else tray.StartAtLogon();
            return;
        }
        var model = new MainViewModel(demo);
        var window = new MainWindow(model); MainWindow = window;
        // Signing out during an Apply: stop, and give a committed replacement a few seconds to verify and record.
        SessionEnding += (_, _) => { if (model.Busy) model.CancelCommand.Execute(null); CommitGate.WaitIdle(TimeSpan.FromSeconds(4)); };
        instance?.OnActivate(Dispatcher, () => { window.Show(); if (window.WindowState == WindowState.Minimized) window.WindowState = WindowState.Normal; window.Activate(); });
        if (smokeDirectory != null) window.Loaded += async (_, _) => { bool passed = await UiSmoke.RunAsync(window, model, smokeDirectory); Shutdown(passed && !HasFatalError ? 0 : 3); };
        window.Show();
        if (fatalSmoke) Dispatcher.BeginInvoke(new Action(() => throw new InvalidOperationException("FATAL_SMOKE: deliberate diagnostic test; application must terminate with exit code 1.")));
    }
    protected override void OnExit(ExitEventArgs e)
    {
        IdleChecks?.Dispose(); Agent?.Dispose(); instance?.Dispose();
        base.OnExit(e);
    }
}
