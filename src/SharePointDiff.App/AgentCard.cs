using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using SharePointDiff.Core;

namespace SharePointDiff.App;

/// <summary>The tray agent's single, code-built card (prompt, progress, sign-in, problem), shown bottom-right without taking focus.</summary>
internal sealed class AgentCard : Window
{
    private enum Mode { Prompt, Progress, SignIn, Problem }
    private readonly TrayAgent agent;
    private readonly StackPanel body = new();
    private readonly StackPanel buttons = new() { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new(0, 16, 0, 0) };
    private readonly ProgressBar bar = new() { Height = 6, Minimum = 0, Maximum = 1000, Margin = new(0, 12, 0, 6) };
    private readonly TextBlock progressText = new() { TextWrapping = TextWrapping.Wrap }, progressPath = new() { TextTrimming = TextTrimming.CharacterEllipsis, FontSize = 12 };
    private Mode mode; private CheckOutcome? outcome; private bool demo; private string problem = "";
    private static readonly Brush Muted = new SolidColorBrush(Color.FromRgb(0x52, 0x65, 0x7B)), Amber = new SolidColorBrush(Color.FromRgb(0xB5, 0x47, 0x08));
    internal Button? SyncButton { get; private set; }
    internal TextBlock? WarningLine { get; private set; }
    internal TextBlock? ReviewLine { get; private set; }
    internal IReadOnlyList<Button> Buttons => buttons.Children.OfType<Button>().ToArray();
    private LocalizationService L => agent.L;
    public AgentCard(TrayAgent agent)
    {
        this.agent = agent;
        Title = "SharePoint Diff"; Width = Math.Min(520, SystemParameters.WorkArea.Width - 32); MaxHeight = SystemParameters.WorkArea.Height - 32; SizeToContent = SizeToContent.Height; ResizeMode = ResizeMode.NoResize;
        Topmost = true; ShowActivated = false; ShowInTaskbar = true; WindowStartupLocation = WindowStartupLocation.Manual;
        FontFamily = new FontFamily("Segoe UI"); FontSize = 13; Background = new SolidColorBrush(Color.FromRgb(245, 247, 250));
        var root = new DockPanel { Margin = new(20) }; DockPanel.SetDock(buttons, Dock.Bottom); root.Children.Add(buttons);
        root.Children.Add(new ScrollViewer { Content = body, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, MaxHeight = Math.Max(180, SystemParameters.WorkArea.Height - 160) });
        Content = new Border { Child = root };
        SizeChanged += (_, _) => PlaceAtCorner();
        Closed += (_, _) => agent.CardClosed(this);
    }
    private void PlaceAtCorner()
    {
        var area = SystemParameters.WorkArea;
        Left = area.Right - ActualWidth - 16; Top = Math.Max(area.Top + 16, area.Bottom - ActualHeight - 16);
    }
    private void Present() { if (!IsVisible) Show(); PlaceAtCorner(); }
    private TextBlock Line(string text, Brush? brush = null, double size = 13, FontWeight? weight = null, double top = 4) =>
        new() { Text = text, TextWrapping = TextWrapping.Wrap, Foreground = brush ?? (Brush)new SolidColorBrush(Color.FromRgb(0x17, 0x2C, 0x43)), FontSize = size, FontWeight = weight ?? FontWeights.Normal, Margin = new(0, top, 0, 0) };
    private Button Action(string key, Action click, bool primary = false, bool enabled = true)
    {
        var button = new Button { Content = L[key], IsEnabled = enabled, Margin = new(8, 0, 0, 0), MinWidth = 88, Tag = key };
        if (primary && TryFindResource("PrimaryButton") is Style style) button.Style = style;
        button.Click += (_, _) => click(); buttons.Children.Add(button); return button;
    }
    public void ShowPrompt(CheckOutcome value, bool isDemo) { mode = Mode.Prompt; outcome = value; demo = isDemo; Build(); Present(); }
    /// <param name="reset">False re-shows a running sync's progress (after Hide or a tray click) without clearing it.</param>
    public void ShowProgress(bool reset = true)
    {
        if (reset || mode != Mode.Progress) { bar.Value = 0; progressText.Text = L["TraySyncing"]; progressPath.Text = ""; }
        mode = Mode.Progress; Build(); Present();
    }
    public void ShowSignIn() { mode = Mode.SignIn; Build(); Present(); }
    public void ShowProblem(string code) { mode = Mode.Problem; problem = code; Build(); Present(); }
    public void Relabel() => Build();
    public void UpdateProgress(OperationProgress p, LocalizationService l)
    {
        if (mode != Mode.Progress) return;
        if (p.TotalBytes > 0) bar.Value = Math.Clamp(1000d * p.Bytes / p.TotalBytes, 0, 1000);
        else if (p.Total > 0) bar.Value = Math.Clamp(1000d * p.Completed / p.Total, 0, 1000);
        progressText.Text = l.Format(p.Code, p.Completed, p.Total); progressPath.Text = p.Path ?? "";
    }
    private void Build()
    {
        body.Children.Clear(); buttons.Children.Clear(); SyncButton = null; WarningLine = null; ReviewLine = null;
        switch (mode)
        {
            case Mode.Prompt when outcome != null: BuildPrompt(outcome); break;
            case Mode.Progress:
                // Clear() above detached the reusable progress controls from the body, so they can be added again.
                body.Children.Add(Line(L["TraySyncing"], size: 17, weight: FontWeights.SemiBold, top: 0));
                body.Children.Add(bar); body.Children.Add(progressText); body.Children.Add(progressPath);
                Action("CardStop", agent.CardStop); Action("CardHide", Hide, primary: true);
                break;
            case Mode.SignIn:
                body.Children.Add(Line(L["BalloonSignInTitle"], size: 17, weight: FontWeights.SemiBold, top: 0));
                body.Children.Add(Line(L["CardSignInText"], top: 8));
                Action("CardClose", Close); Action("CardSignIn", agent.CardSignIn, primary: true, enabled: !demo);
                break;
            case Mode.Problem:
                body.Children.Add(Line(L["CardProblemTitle"], size: 17, weight: FontWeights.SemiBold, top: 0));
                body.Children.Add(Line(agent.ProblemText(problem), top: 8));
                Action("CardClose", Close); Action("CardOpen", () => agent.OpenMain(null)); Action("CardRetry", agent.CardRetry, primary: true, enabled: !demo);
                break;
        }
    }
    private void BuildPrompt(CheckOutcome o)
    {
        var s = o.Summary;
        body.Children.Add(Line(L["CardTitle"], size: 17, weight: FontWeights.SemiBold, top: 0));
        body.Children.Add(Line(L.Format("CardSite", o.Scan.SiteUrl), Muted, 12, top: 6));
        body.Children.Add(Line(L.Format("CardFolder", o.Scan.LocalRoot), Muted, 12, top: 2));
        body.Children.Add(Line(L.Format("CardScannedAt", o.Scan.ScannedAt.LocalDateTime), Muted, 12, top: 2));
        body.Children.Add(Line(L.Format("CardNew", s.NewFiles, FileRow.FormatBytes(s.NewBytes)), weight: FontWeights.SemiBold, top: 12));
        body.Children.Add(Line(L.Format("CardUpdates", s.Updates, FileRow.FormatBytes(s.UpdateBytes)), weight: FontWeights.SemiBold));
        if (s.LocallyChanged > 0)
        {
            WarningLine = Line(L.Format("CardLocallyChanged", s.LocallyChanged, s.Conflicts, s.Reverts, s.Unknown) + " " + L["CardWarning"], Amber, top: 8);
            body.Children.Add(WarningLine);
            foreach (string path in s.LocallyChangedSample) body.Children.Add(Line("• " + path, Amber, 12, top: 1));
            if (s.LocallyChanged > s.LocallyChangedSample.Count) body.Children.Add(Line(L.Format("CardMoreFiles", s.LocallyChanged - s.LocallyChangedSample.Count), Amber, 12, top: 1));
        }
        if (s.NoHash > 0) body.Children.Add(Line(L.Format("CardNoHash", s.NoHash), Muted, 12, top: 6));
        if (s.Review > 0) { ReviewLine = Line(L.Format("CardReview", s.Review), Muted, 12, top: 6); body.Children.Add(ReviewLine); }
        if (s.NewLibraries > 0) body.Children.Add(Line(L.Format("CardNewLibraries", s.NewLibraries), Muted, 12, top: 6));
        if (s.FailedLibraries > 0) body.Children.Add(Line(L.Format("CardFailedLibraries", s.FailedLibraries), Amber, 12, top: 6));
        if (s.FirstSync) body.Children.Add(Line(L["CardFirstSync"], Muted, 12, top: 6));
        body.Children.Add(Line(L.Format("CardEstimate", FileRow.FormatBytes(s.NetworkBytes), FileRow.FormatBytes(s.BackupBytes), s.Reused), Muted, 12, top: 10));
        body.Children.Add(Line(L["CardNothingWritten"], Muted, 12, top: 2));
        Action("CardLater", Close);
        Action("CardDetails", () => agent.OpenMain(o));
        SyncButton = Action("CardSync", agent.CardSync, primary: true, enabled: !demo);
    }
}
