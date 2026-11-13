using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using SharePointDiff.Core;

namespace SharePointDiff.App;

/// <summary>One optional idle reader for the existing app process. No Windows task or additional process.</summary>
internal sealed class IdleVerifier : IDisposable
{
    private readonly Application app;
    private readonly DispatcherTimer timer;
    private CancellationTokenSource? cancellation;
    private Task<ScrubResult>? running;
    private bool stopped, disposed;
    public bool Active => running is { IsCompleted: false };
    public bool PausedByUser => stopped;
    public ScrubResult? LastResult { get; private set; }
    public OperationIssue? LastError { get; private set; }
    public event Action? Changed;
    public IdleVerifier(Application app)
    {
        this.app = app;
        timer = new DispatcherTimer(TimeSpan.FromSeconds(30), DispatcherPriority.Background, Tick, app.Dispatcher);
    }
    public void Resume() { stopped = false; LastError = null; Changed?.Invoke(); }
    public async Task PauseAsync(bool user = false)
    {
        if (user) stopped = true;
        cancellation?.Cancel();
        if (running != null)
        {
            try { await running; }
            catch (OperationCanceledException) { }
            catch (Exception ex) when (ex is OperationException or IOException or UnauthorizedAccessException) { } // Tick records and displays expected errors.
        }
        Changed?.Invoke();
    }
    private async void Tick(object? sender, EventArgs args)
    {
        if (disposed || stopped || Active || IdleTime() < TimeSpan.FromMinutes(5) || (app as App)?.Agent?.IsBusy == true ||
            app.Windows.OfType<MainWindow>().Any(w => w.DataContext is MainViewModel { Busy: true })) return;
        try
        {
            var settings = SettingsStore.LoadExisting();
            if (settings is not { IdleVerification: true } || string.IsNullOrWhiteSpace(settings.LocalRoot) || !Directory.Exists(settings.LocalRoot)) return;
            string data = Workspace.DataRootFor(settings.SiteUrl, settings.LocalRoot);
            if (!File.Exists(Path.Combine(data, "local-hash-cache.json"))) return;
            var workspace = new Workspace(settings.SiteUrl, settings.LocalRoot);
            cancellation?.Dispose(); cancellation = new(); var token = cancellation.Token;
            LastError = null;
            running = Task.Run(() => LocalScrubber.RunAsync(workspace, () => IdleTime() < TimeSpan.FromMinutes(5), token), token);
            Changed?.Invoke();
            LastResult = await running;
            if (LastResult.Issues.Count > 0) LastError = LastResult.Issues[0];
        }
        catch (OperationCanceledException) { }
        catch (OperationException ex) when (ex.Code == "WorkspaceBusy") { }
        catch (Exception ex) when (ex is OperationException or IOException or UnauthorizedAccessException)
        {
            LastError = new(ex is OperationException op ? op.Code : "LocalReadFailed", (ex as OperationException)?.ItemPath ?? "", ex.ToString());
            try { ErrorDetails.Save(ex, "IdleVerification", LastError.Path); }
            catch (Exception logError) when (logError is IOException or UnauthorizedAccessException)
            { LastError = LastError with { Detail = LastError.Detail + Environment.NewLine + logError }; }
        }
        finally { if (!disposed) Changed?.Invoke(); }
    }
    [StructLayout(LayoutKind.Sequential)] private struct LastInput { public uint Size, Ticks; }
    [DllImport("user32.dll")] private static extern bool GetLastInputInfo(ref LastInput input);
    internal static TimeSpan IdleTime()
    {
        var input = new LastInput { Size = (uint)Marshal.SizeOf<LastInput>() };
        return GetLastInputInfo(ref input) ? TimeSpan.FromMilliseconds(unchecked((uint)Environment.TickCount - input.Ticks)) : TimeSpan.Zero;
    }
    public void Dispose() { disposed = true; timer.Stop(); cancellation?.Cancel(); }
}
