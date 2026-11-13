using System.ComponentModel;
using System.Globalization;
using System.Resources;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using System.Windows;
using SharePointDiff.Core;

namespace SharePointDiff.App;

public abstract class Observable : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;
    protected bool Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false; field = value; Raise(name); return true;
    }
    public void Raise([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new(name));
}
public sealed class RelayCommand(Action execute, Func<bool>? can = null) : ICommand
{
    public event EventHandler? CanExecuteChanged;
    public bool CanExecute(object? parameter) => can?.Invoke() ?? true;
    public void Execute(object? parameter) => execute();
    public void Update() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
}
public sealed class AsyncCommand(Func<Task> execute, Func<bool>? can = null) : ICommand
{
    private bool executing;
    public event EventHandler? CanExecuteChanged;
    public bool CanExecute(object? parameter) => !executing && (can?.Invoke() ?? true);
    public async void Execute(object? parameter)
    {
        if (!CanExecute(parameter)) return; executing = true; Update();
        try { await execute(); } finally { executing = false; Update(); }
    }
    public void Update() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
}
public sealed class LocalizationService : Observable
{
    private readonly ResourceManager resources = new("SharePointDiff.App.Resources.Strings", typeof(LocalizationService).Assembly);
    public string CultureName { get; private set; } = "vi-VN";
    public CultureInfo Culture => CultureInfo.GetCultureInfo(CultureName);
    public string this[string key] => resources.GetString(key, Culture) ?? resources.GetString(key, CultureInfo.GetCultureInfo("en-US")) ?? key;
    public string Format(string key, params object?[] args) => string.Format(Culture, this[key], args);
    public void Switch(string language)
    {
        CultureName = language == "en-US" ? "en-US" : "vi-VN";
        CultureInfo.CurrentUICulture = Culture; CultureInfo.CurrentCulture = Culture;
        Raise("Item[]"); Raise(nameof(CultureName));
    }
}
public sealed class ViewModelProxy : Freezable
{
    public static readonly DependencyProperty DataProperty = DependencyProperty.Register(nameof(Data), typeof(object), typeof(ViewModelProxy));
    public object? Data { get => GetValue(DataProperty); set => SetValue(DataProperty, value); }
    protected override Freezable CreateInstanceCore() => new ViewModelProxy();
}

/// <summary>Throttle before posting to WPF rather than filling its queue with discarded callbacks.</summary>
internal sealed class ThrottledProgress(IProgress<OperationProgress> target) : IProgress<OperationProgress>
{
    private long last;
    public void Report(OperationProgress value)
    {
        long now = Environment.TickCount64;
        bool final = value.Total > 0 && value.Completed == value.Total;
        long previous = Volatile.Read(ref last);
        if (!final && now - previous < 120) return;
        if (!final && Interlocked.CompareExchange(ref last, now, previous) != previous) return;
        if (final) Interlocked.Exchange(ref last, now);
        target.Report(value);
    }
}
