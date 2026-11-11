using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using SharePointDiff.Core;

namespace SharePointDiff.App;

public static partial class ErrorDetails
{
    public static string Format(Exception exception, string operation, string? path = null, string? log = null)
    {
        var chain = new List<string>();
        for (Exception? e = exception; e != null; e = e.InnerException)
        {
            string location = e is System.Windows.Markup.XamlParseException x ? $" (XAML line {x.LineNumber}, position {x.LinePosition})" : "";
            chain.Add(e.GetType().FullName + location + ": " + e.Message);
        }
        string text = $"Operation: {operation}\nPath / resource: {path ?? "—"}\nCode: {(exception as OperationException)?.Code ?? exception.GetType().Name}\nTime: {DateTimeOffset.Now:O}\nLog: {log ?? "—"}\n\n" + string.Join("\n\nCaused by:\n", chain) + "\n\n" + exception.StackTrace;
        return Redact(text);
    }
    public static string Redact(string text) => Regex.Replace(text, @"(?i)(bearer\s+)[a-z0-9._~-]+|((?:access_token|refresh_token|client_secret|tempauth)=)[^\s&]+", "$1$2[redacted]");
    public static string Save(Exception ex, string operation, string? path = null)
    {
        string directory = SettingsStore.LogDirectory;
        Directory.CreateDirectory(directory);
        string log = Path.Combine(directory, "error-" + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + ".txt");
        // Exception reporting must not synchronously wait on an async UI-context continuation.
        File.WriteAllText(log, Format(ex, operation, path, log));
        return log;
    }
}

public sealed class ErrorDialog : Window
{
    public ErrorDialog(LocalizationService l, Exception ex, string operation, string? path = null, string? log = null)
    {
        Title = "SharePoint Diff — " + l["ErrorTitle"]; Width = 850; Height = 580; MinWidth = 600; MinHeight = 400;
        WindowStartupLocation = WindowStartupLocation.CenterScreen; Background = new SolidColorBrush(Color.FromRgb(245, 247, 250));
        FontFamily = new FontFamily("Segoe UI"); FontSize = 13;
        var panel = new Grid { Margin = new Thickness(22) };
        panel.RowDefinitions.Add(new() { Height = GridLength.Auto }); panel.RowDefinitions.Add(new() { Height = GridLength.Auto });
        panel.RowDefinitions.Add(new() { Height = new GridLength(1, GridUnitType.Star) }); panel.RowDefinitions.Add(new() { Height = GridLength.Auto });
        var title = new TextBlock { Text = l["ErrorTitle"], FontSize = 22, FontWeight = FontWeights.SemiBold, Margin = new(0, 0, 0, 10) };
        panel.Children.Add(title);
        var root = ex.GetBaseException();
        var summary = new TextBlock { Text = l["OperationLabel"] + ": " + operation + Environment.NewLine + root.Message, TextWrapping = TextWrapping.Wrap, Margin = new(0, 0, 0, 12) };
        Grid.SetRow(summary, 1); panel.Children.Add(summary);
        string detail = ErrorDetails.Format(ex, operation, path, log);
        var text = new TextBox { Text = detail, IsReadOnly = true, TextWrapping = TextWrapping.Wrap, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, FontFamily = new FontFamily("Consolas"), FontSize = 12, Padding = new(10) };
        Grid.SetRow(text, 2); panel.Children.Add(text);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new(0, 14, 0, 0) };
        var copy = new Button { Content = l["CopyError"], Padding = new(14, 7, 14, 7), Margin = new(0, 0, 8, 0) };
        copy.Click += (_, _) => { try { Clipboard.SetText(detail); } catch { text.SelectAll(); text.Focus(); } };
        var close = new Button { Content = l["Close"], Padding = new(14, 7, 14, 7), IsCancel = true }; close.Click += (_, _) => Close();
        buttons.Children.Add(copy); buttons.Children.Add(close); Grid.SetRow(buttons, 3); panel.Children.Add(buttons);
        Content = panel;
    }
}
