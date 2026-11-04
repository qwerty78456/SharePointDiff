using System.Windows;

namespace SharePointDiff.App;

public partial class App : Application
{
    internal static bool HasFatalError { get; private set; }
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += (_, args) =>
        {
            if (HasFatalError) { Environment.Exit(1); return; }
            HasFatalError = true;
            (MainWindow?.DataContext as MainViewModel)?.Dispose();
            if (MainWindow != null) MainWindow.IsEnabled = false;
            var l = (MainWindow?.DataContext as MainViewModel)?.L ?? new LocalizationService();
            string? log = null;
            try
            {
                log = ErrorDetails.Save(args.Exception, (string?)args.Exception.Data["Operation"] ?? (MainWindow?.DataContext as MainViewModel)?.StatusText ?? l["StartingApplication"],
                    (string?)args.Exception.Data["Path"] ?? (MainWindow?.DataContext as MainViewModel)?.ProgressPath);
            }
            catch (Exception logError) when (logError is IOException or UnauthorizedAccessException)
            { log = l["LogWriteFailed"] + ": " + logError.Message; }
            if (!e.Args.Contains("--fatal-smoke")) new ErrorDialog(l, args.Exception, (string?)args.Exception.Data["Operation"] ?? (MainWindow?.DataContext as MainViewModel)?.StatusText ?? l["StartingApplication"],
                (string?)args.Exception.Data["Path"] ?? (MainWindow?.DataContext as MainViewModel)?.ProgressPath, log).ShowDialog();
            // Mark handled only to replace the generic OS dialog with diagnostics, then terminate.
            // An unexpected exception must never return to normal application execution.
            args.Handled = true;
            Shutdown(1);
        };
        var vm = new MainViewModel(e.Args.Contains("--demo"));
        var window = new MainWindow(vm); MainWindow = window;
        int smoke = Array.IndexOf(e.Args, "--ui-smoke");
        if (smoke >= 0 && smoke + 1 < e.Args.Length) window.Loaded += async (_, _) => await UiSmoke.RunAsync(window, vm, e.Args[smoke + 1]);
        window.Show();
        if (e.Args.Contains("--fatal-smoke")) Dispatcher.BeginInvoke(new Action(() => throw new InvalidOperationException("FATAL_SMOKE: deliberate diagnostic test; application must terminate with exit code 1.")));
    }
}
