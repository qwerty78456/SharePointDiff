using System.Windows;

namespace SharePointDiff.App;

public partial class MainWindow : Window
{
    public MainWindow(MainViewModel vm)
    {
        InitializeComponent(); DataContext = vm;
        Width = Math.Max(MinWidth, Math.Min(Width, SystemParameters.WorkArea.Width - 32));
        Height = Math.Max(MinHeight, Math.Min(Height, SystemParameters.WorkArea.Height - 32));
        Closing += (_, e) => { if (vm.Busy) { vm.CancelCommand.Execute(null); e.Cancel = true; } else vm.Dispose(); };
    }
}
