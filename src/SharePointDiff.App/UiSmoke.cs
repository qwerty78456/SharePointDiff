using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using SharePointDiff.Core;

namespace SharePointDiff.App;

/// <summary>Opt-in rendering/integration harness on synthetic data, not a production operation.</summary>
internal static class UiSmoke
{
    public static async Task<bool> RunAsync(MainWindow window, MainViewModel vm, string directory)
    {
        Directory.CreateDirectory(directory);
        var checks = new Dictionary<string, object>();
        try
        {
            if (!vm.Demo) throw new InvalidOperationException("UI smoke requires --demo");
            await window.Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
            if (App.HasFatalError) throw new InvalidOperationException("Fatal startup error; UI test must fail");
            var grid = window.FilesGrid;
            checks["SyntheticRows"] = vm.Rows.Count;
            var diagnostics = new List<object>();
            for (int i = 0; i < Math.Min(8, grid.Items.Count); i++)
            {
                grid.ScrollIntoView(grid.Items[i]); grid.UpdateLayout();
                var row = grid.ItemContainerGenerator.ContainerFromIndex(i) as DataGridRow;
                var pathText = row == null ? null : Descendants<TextBlock>(row).FirstOrDefault(t => t.Text == (row.Item as FileRow)?.Path);
                if (pathText == null || pathText.Foreground.ToString() == "#FFFFFFFF") throw new InvalidOperationException("File row text is white or missing");
                diagnostics.Add(new
                {
                    Index = i,
                    ItemType = row?.Item?.GetType().Name,
                    DataContextType = row?.DataContext?.GetType().Name,
                    ModelPath = (row?.Item as FileRow)?.Path,
                    Texts = row == null ? [] : Descendants<TextBlock>(row).Select(t => new { t.Text, t.ActualWidth, t.ActualHeight, Foreground = t.Foreground.ToString() }).ToArray()
                });
            }
            await JsonFiles.WriteAsync(Path.Combine(directory, "row-diagnostics.json"), diagnostics);
            var first = vm.Rows.First(x => x.CanSelect); first.Selected = true;
            vm.SelectedRow = first;
            var permissionReports = new[] { new TokenPermissionReport("https://graph.microsoft.com", true, ["AllSites.FullControl", "Sites.Read.All", "openid", "profile", "email"], ["AllSites.FullControl"]),
                new TokenPermissionReport("https://contoso.sharepoint.com", true, ["AllSites.FullControl", "Sites.Read.All"], ["AllSites.FullControl"]) };
            vm.SetLocalizedError(() => PermissionMessages.Format(vm.L, permissionReports));
            foreach (var language in new[] { 0, 1 })
            {
                vm.LanguageIndex = language;
                window.Width = Math.Min(1320, SystemParameters.WorkArea.Width - 32); window.Height = Math.Min(880, SystemParameters.WorkArea.Height - 32);
                await window.Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
                Save(window, Path.Combine(directory, language == 0 ? "vi-wide.png" : "en-wide.png"));
                if (!first.Selected) throw new InvalidOperationException("Selection lost during language change");
                if (!vm.Error.StartsWith(vm.L["PermissionsCurrent"]) || !vm.Error.Contains(vm.L["PermissionsRemove"]) || !vm.Error.Contains("AllSites.FullControl") || !vm.Error.Contains("AllSites.Read"))
                    throw new InvalidOperationException("Permission error did not list granted/required/rejected scopes in the current language");
            }
            checks["PermissionErrorRelocalizes"] = true;
            vm.SetLocalizedError(() => "");
            checks["LanguageSelectionPreserved"] = true;
            checks["LanguageLabelHasRoom"] = window.LanguageLabel.ActualWidth >= 76;
            if (window.LanguageLabel.ActualWidth < 76) throw new InvalidOperationException("Language label clipped");
            checks["VietnameseResources"] = new LocalizationService()["Compare"];
            window.Width = Math.Min(1000, SystemParameters.WorkArea.Width - 32); window.Height = Math.Min(740, SystemParameters.WorkArea.Height - 32);
            await window.Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
            Save(window, Path.Combine(directory, "en-compact.png"));
            Save(window, Path.Combine(directory, "en-compact-150.png"), 1.5);
            Save(window, Path.Combine(directory, "en-compact-200.png"), 2);
            int realized = Enumerable.Range(0, grid.Items.Count).Count(i => grid.ItemContainerGenerator.ContainerFromIndex(i) != null);
            checks["RealizedRows"] = realized;
            if (realized >= 500) throw new InvalidOperationException("Row virtualization is not active");
            vm.ModeIndex = 1;
            var watch = Stopwatch.StartNew(); vm.SelectVisibleCommand.Execute(null); watch.Stop();
            checks["SelectVisibleMilliseconds"] = watch.ElapsedMilliseconds;
            checks["SelectedRows"] = vm.Rows.Count(x => x.Selected);
            if (watch.Elapsed > TimeSpan.FromSeconds(3)) throw new InvalidOperationException("Selection blocked UI too long");
            vm.ModeIndex = 2;
            checks["AllDifferencePlan"] = vm.CurrentPlan.Replacements;
            vm.Filter = "Review_09999"; vm.FilesView.Refresh();
            checks["FilteredRows"] = vm.FilesView.Cast<object>().Count();
            checks["AllModeIncludesHiddenDifferences"] = vm.CurrentPlan.Replacements == 1000;
            if (vm.CurrentPlan.Replacements != 1000) throw new InvalidOperationException("All mode was limited by UI filter");
            vm.Filter = ""; vm.LanguageIndex = 0; window.Width = Math.Min(1320, SystemParameters.WorkArea.Width - 32); window.Height = Math.Min(880, SystemParameters.WorkArea.Height - 32);
            if (App.HasFatalError) throw new InvalidOperationException("Fatal UI error; UI test must fail");
            checks["StartWithWindowsDisabledInDemo"] = !vm.CanStartWithWindows;
            if (vm.CanStartWithWindows) throw new InvalidOperationException("Startup toggle must be disabled in demo");
            vm.ConnectionExpanded = true; await window.Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
            Save(window, Path.Combine(directory, "en-connection.png")); vm.ConnectionExpanded = false;
            checks["Passed"] = true;
        }
        catch (Exception ex) { checks["Passed"] = false; checks["Error"] = ex.ToString(); }
        await JsonFiles.WriteAsync(Path.Combine(directory, "ui-checks.json"), checks);
        return checks["Passed"] is true;
    }
    internal static void Save(Window window, string path, double scale = 1)
    {
        window.UpdateLayout();
        var visual = (FrameworkElement)window.Content;
        var bitmap = new RenderTargetBitmap((int)(visual.ActualWidth * scale), (int)(visual.ActualHeight * scale), 96 * scale, 96 * scale, PixelFormats.Pbgra32);
        var drawing = new DrawingVisual();
        using (var context = drawing.RenderOpen())
        {
            var bounds = new Rect(0, 0, visual.ActualWidth, visual.ActualHeight);
            context.DrawRectangle(window.Background, null, bounds);
            var brush = new VisualBrush(visual) { ViewboxUnits = BrushMappingMode.Absolute, Viewbox = bounds, Stretch = Stretch.Fill };
            context.DrawRectangle(brush, null, bounds);
        }
        bitmap.Render(drawing);
        var png = new PngBitmapEncoder(); png.Frames.Add(BitmapFrame.Create(bitmap));
        using var file = File.Create(path); png.Save(file);
    }
    private static IEnumerable<T> Descendants<T>(DependencyObject parent) where T : DependencyObject
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i); if (child is T t) yield return t;
            foreach (var nested in Descendants<T>(child)) yield return nested;
        }
    }
}
