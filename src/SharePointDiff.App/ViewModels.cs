using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Text.Json;
using System.Windows;
using System.Windows.Data;
using SharePointDiff.Core;

namespace SharePointDiff.App;

public sealed class LibraryRow(LibraryInfo value) : Observable
{
    private bool selected = true;
    public LibraryInfo Value { get; } = value;
    public string Title => Value.Title;
    public string Folder => Value.LocalFolder;
    public bool Selected { get => selected; set => Set(ref selected, value); }
}
public sealed class FileRow(DiffEntry entry, LocalizationService l, Func<OverwriteMode> mode, Action changed) : Observable
{
    private bool selected = entry.Status == DiffStatus.MissingLocal;
    public DiffEntry Entry { get; } = entry;
    public string Path => Entry.RelativePath;
    public bool Selected { get => selected; set { if (Set(ref selected, value)) changed(); } }
    public bool CanSelect => Entry.Status == DiffStatus.MissingLocal || (Entry.IsDifference && mode() == OverwriteMode.Selected);
    public string Status => l["Status" + Entry.Status];
    public string LocalSize => Entry.Local == null ? "—" : FormatBytes(Entry.Local.Hashes.Size);
    public string RemoteSize => Entry.Remote == null ? "—" : FormatBytes(Entry.Remote.Size);
    public string Reuse => Entry.ReusePath == null ? "" : l["ReuseYes"];
    public string StatusColor => Entry.Status switch { DiffStatus.Match => "#047857", DiffStatus.MissingLocal => "#086A72", DiffStatus.DifferentHash or DiffStatus.DifferentSize => "#995E13", DiffStatus.Error => "#B42318", DiffStatus.Unverified => "#6941C6", _ => "#52657B" };
    public string StatusBackground => Entry.Status switch { DiffStatus.Match => "#ECFDF3", DiffStatus.MissingLocal => "#E8F7F8", DiffStatus.DifferentHash or DiffStatus.DifferentSize => "#FFF5DE", DiffStatus.Error => "#FEF3F2", DiffStatus.Unverified => "#F4F0FF", _ => "#F1F5F9" };
    public string Details
    {
        get
        {
            var e = Entry; var r = e.Remote;
            return string.Join(Environment.NewLine, new[]{e.RelativePath,"",l["LocalHash"]+": "+(e.Local?.Hashes.QuickXorHash??"—"),
                l["RemoteHash"]+": "+(r?.QuickXorHash??"—"),l["Sha256"]+": "+(e.Local?.Hashes.Sha256??"—"),
                l["ReuseSource"]+": "+(e.ReusePath??"—"),e.Issue==null?"":l["Error"+e.Issue.Code]+": "+e.Issue.Detail,"",
                l["Metadata"]+":",r==null?"—":JsonSerializer.Serialize(r.Metadata,JsonFiles.Options)});
        }
    }
    public void RefreshText() { Raise(nameof(Status)); Raise(nameof(Reuse)); Raise(nameof(Details)); Raise(nameof(CanSelect)); }
    public static string FormatBytes(long bytes) => bytes >= 1_000_000_000 ? $"{bytes / 1_000_000_000d:0.00} GB" : bytes >= 1_000_000 ? $"{bytes / 1_000_000d:0.0} MB" : bytes >= 1000 ? $"{bytes / 1000d:0.0} KB" : $"{bytes} B";
}

public sealed class MainViewModel : Observable, IDisposable
{
    public LocalizationService L { get; } = new();
    public ObservableCollection<ConnectionProfile> Profiles { get; } = [];
    public ObservableCollection<LibraryRow> Libraries { get; } = [];
    public ObservableCollection<FileRow> Rows { get; } = [];
    public ICollectionView FilesView { get; }
    public string[] Languages { get; } = ["Tiếng Việt", "English"];
    public bool Demo { get; }
    private UserSettings settings;
    private SharePointReader? reader;
    private CancellationTokenSource? cancellation;
    private ScanResult? scan;
    private bool busy, connected, bulkSelecting, connectionExpanded = true;
    private string siteUrl = "", localRoot = "", filter = "", statusCode = "Ready", progressPath = "", reportDirectory = "";
    private string tenant = "", client = "", profileName = "";
    private int languageIndex, profileIndex = -1, modeIndex, statusFilterIndex;
    private int completed, total;
    private FileRow? selectedRow;
    private string error = "";
    private string errorDetail = "";
    public MainViewModel(bool demo = false)
    {
        Demo = demo; settings = demo ? new() : SettingsStore.Load();
        foreach (var p in settings.Profiles) Profiles.Add(p);
        siteUrl = settings.SiteUrl; localRoot = settings.LocalRoot; profileIndex = Profiles.Count > 0 ? Math.Clamp(settings.SelectedProfile, 0, Profiles.Count - 1) : -1;
        languageIndex = settings.Language == "en-US" ? 1 : 0; L.Switch(languageIndex == 1 ? "en-US" : "vi-VN");
        FilesView = CollectionViewSource.GetDefaultView(Rows); FilesView.Filter = FilterRow;
        ConnectCommand = new(() => RunAsync(ConnectAsync), () => !Busy && !Demo);
        ScanCommand = new(() => RunAsync(ScanAsync), () => !Busy && connected && Libraries.Any(x => x.Selected));
        ApplyCommand = new(() => RunAsync(ApplyAsync), () => !Busy && !Demo && scan != null && CurrentPlan.Items.Count > 0);
        CancelCommand = new(() => cancellation?.Cancel(), () => Busy);
        BrowseCommand = new(Browse, () => !Busy);
        AddProfileCommand = new(AddProfile, () => !Busy);
        SaveProfileCommand = new(SaveProfile, () => !Busy);
        SelectVisibleCommand = new(() => SelectVisible(true), () => !Busy);
        ClearVisibleCommand = new(() => SelectVisible(false), () => !Busy);
        OpenReportsCommand = new(OpenReports, () => Directory.Exists(reportDirectory));
        CopyErrorCommand = new(() => { try { Clipboard.SetText(ErrorDetail); } catch { } }, () => ErrorDetail.Length > 0);
        L.PropertyChanged += (_, _) => RefreshLanguage();
        if (profileIndex >= 0) LoadProfileFields();
        if (demo) LoadDemo();
    }
    public string SiteUrl { get => siteUrl; set { if (Set(ref siteUrl, value)) InvalidateConnection(); } }
    public string LocalRoot { get => localRoot; set { if (Set(ref localRoot, value)) InvalidateConnection(); } }
    public string Filter { get => filter; set { if (Set(ref filter, value)) FilesView.Refresh(); } }
    public bool Busy { get => busy; private set { Set(ref busy, value); Raise(nameof(NotBusy)); UpdateCommands(); } }
    public bool NotBusy => !Busy;
    public bool ConnectionExpanded { get => connectionExpanded; set => Set(ref connectionExpanded, value); }
    public int LanguageIndex { get => languageIndex; set { if (Set(ref languageIndex, value)) { L.Switch(value == 1 ? "en-US" : "vi-VN"); Persist(); } } }
    public int ProfileIndex { get => profileIndex; set { if (Set(ref profileIndex, value)) { LoadProfileFields(); InvalidateConnection(); } } }
    public string ProfileName { get => profileName; set => Set(ref profileName, value); }
    public string Tenant { get => tenant; set => Set(ref tenant, value); }
    public string Client { get => client; set => Set(ref client, value); }
    public int ModeIndex { get => modeIndex; set { if (Set(ref modeIndex, value)) { foreach (var r in Rows) r.RefreshText(); RefreshPlan(); } } }
    public int StatusFilterIndex { get => statusFilterIndex; set { if (Set(ref statusFilterIndex, value)) FilesView.Refresh(); } }
    public FileRow? SelectedRow { get => selectedRow; set { Set(ref selectedRow, value); Raise(nameof(Details)); Raise(nameof(SelectedPath)); Raise(nameof(MetadataItems)); } }
    public string Details => SelectedRow?.Details ?? L["DetailsHint"];
    public string SelectedPath => SelectedRow?.Path ?? L["DetailsHint"];
    public IReadOnlyList<MetadataDisplay> MetadataItems
    {
        get
        {
            if (SelectedRow?.Entry.Remote is not { } r) return [];
            var m = r.Metadata; var rows = new List<MetadataDisplay>();
            foreach (var c in m.Columns.EnumerateArray())
            {
                if (c.TryGetProperty("Hidden", out var hidden) && hidden.GetBoolean()) continue;
                string n = c.GetProperty("InternalName").GetString()!;
                string value = MetadataFields.TryGet(m.DisplayValues, c, out var d) ? d.ToString() : MetadataFields.TryGet(m.Fields, c, out var raw) ? raw.ToString() : "";
                if (value.Length > 0) rows.Add(new(c.GetProperty("Title").ToString(), value));
            }
            if (rows.Count == 0) foreach (var p in m.Fields.EnumerateObject()) rows.Add(new(p.Name, p.Value.ToString()));
            return rows;
        }
    }
    public IReadOnlyList<SummaryStat> SummaryStats => [
        new(L["StatTotal"],Rows.Count,"#172C43"),new(L["StatMatch"],Rows.Count(x=>x.Entry.Status==DiffStatus.Match),"#047857"),
        new(L["StatMissing"],Rows.Count(x=>x.Entry.Status==DiffStatus.MissingLocal),"#086A72"),new(L["StatDifferent"],Rows.Count(x=>x.Entry.IsDifference),"#995E13"),
        new(L["StatLocal"],Rows.Count(x=>x.Entry.Status==DiffStatus.LocalOnly),"#52657B"),new(L["StatReview"],Rows.Count(x=>x.Entry.Status is DiffStatus.Error or DiffStatus.Unverified),"#6941C6")];
    public string Error { get => error; private set => Set(ref error, value); }
    public string ErrorDetail { get => errorDetail; private set { Set(ref errorDetail, value); CopyErrorCommand?.Update(); } }
    public string StatusText => L.Format(statusCode, completed, total);
    public string ProgressPath { get => progressPath; private set => Set(ref progressPath, value); }
    public double ProgressPercent => total > 0 ? Math.Clamp(100d * completed / total, 0, 100) : 0;
    public bool Indeterminate => Busy && total == 0;
    public string Summary => L.Format("Summary", Rows.Count, Rows.Count(x => x.Entry.Status == DiffStatus.Match), Rows.Count(x => x.Entry.Status == DiffStatus.MissingLocal), Rows.Count(x => x.Entry.IsDifference), Rows.Count(x => x.Entry.Status == DiffStatus.LocalOnly), Rows.Count(x => x.Entry.Status is DiffStatus.Error or DiffStatus.Unverified));
    public SyncPlan CurrentPlan => DiffService.Plan(Rows.Select(x => x.Entry), Rows.Where(x => x.Selected).Select(x => SafePaths.Key(x.Path)).ToHashSet(), (OverwriteMode)ModeIndex);
    public string Estimate
    {
        get { var p = CurrentPlan; return L.Format("Estimate", p.Missing, p.Replacements, p.Reused, FileRow.FormatBytes(p.NetworkBytes), FileRow.FormatBytes(p.BackupBytes)); }
    }
    public AsyncCommand ConnectCommand { get; }
    public AsyncCommand ScanCommand { get; }
    public AsyncCommand ApplyCommand { get; }
    public RelayCommand CancelCommand { get; }
    public RelayCommand BrowseCommand { get; }
    public RelayCommand AddProfileCommand { get; }
    public RelayCommand SaveProfileCommand { get; }
    public RelayCommand SelectVisibleCommand { get; }
    public RelayCommand ClearVisibleCommand { get; }
    public RelayCommand OpenReportsCommand { get; }
    public RelayCommand CopyErrorCommand { get; }
    private bool FilterRow(object o)
    {
        if (o is not FileRow r) return false;
        if (Filter.Length > 0 && !r.Path.Contains(Filter, StringComparison.CurrentCultureIgnoreCase)) return false;
        return StatusFilterIndex switch { 1 => r.Entry.Status == DiffStatus.MissingLocal, 2 => r.Entry.IsDifference, 3 => r.Entry.Status == DiffStatus.LocalOnly, 4 => r.Entry.Status is DiffStatus.Error or DiffStatus.Unverified, 5 => r.Entry.Status == DiffStatus.Match, _ => true };
    }
    private void InvalidateConnection() { connected = false; scan = null; Libraries.Clear(); Rows.Clear(); RefreshPlan(); }
    private void LoadProfileFields() { if (ProfileIndex < 0 || ProfileIndex >= Profiles.Count) return; var p = Profiles[ProfileIndex]; ProfileName = p.Name; Tenant = p.TenantId; Client = p.ClientId; }
    private void AddProfile() { ProfileIndex = -1; ProfileName = ""; Tenant = ""; Client = ""; InvalidateConnection(); }
    private void SaveProfile()
    {
        try
        {
            var p = new ConnectionProfile(ProfileName.Trim(), Tenant.Trim(), Client.Trim());
            if (string.IsNullOrEmpty(p.Name) || !Guid.TryParse(p.ClientId, out _) || string.IsNullOrWhiteSpace(p.TenantId)) throw new OperationException("InvalidProfile");
            int i = ProfileIndex;
            if (i >= 0 && i < Profiles.Count) Profiles[i] = p; else { Profiles.Add(p); i = Profiles.Count - 1; }
            profileIndex = i; Raise(nameof(ProfileIndex)); InvalidateConnection(); Persist(); Error = "";
        }
        catch (OperationException ex) { SetError(ex); }
    }
    private void Browse()
    {
        var dialog = new Microsoft.Win32.OpenFolderDialog { Title = L["ChooseFolder"] };
        if (dialog.ShowDialog() == true) LocalRoot = dialog.FolderName;
    }
    private async Task ConnectAsync(CancellationToken ct)
    {
        if (ProfileIndex < 0 || ProfileIndex >= Profiles.Count) throw new OperationException("InvalidProfile");
        string site = SafePaths.SiteUri(SiteUrl).ToString().TrimEnd('/');
        if (!Directory.Exists(LocalRoot)) throw new OperationException("InvalidRoot");
        reader?.Dispose(); reader = new(new BrowserAuthentication(Profiles[ProfileIndex]));
        SetStatus("Connecting");
        var workspace = new Workspace(site, LocalRoot);
        var mappings = JsonFiles.Read<Dictionary<string, string>>(Path.Combine(workspace.DataRoot, "library-mappings.json")) ?? [];
        var list = await reader.GetLibrariesAsync(site, mappings, ct);
        Libraries.Clear(); foreach (var library in list) { var row = new LibraryRow(library); row.PropertyChanged += (_, _) => { scan = null; Rows.Clear(); RefreshPlan(); }; Libraries.Add(row); }
        await JsonFiles.WriteAsync(Path.Combine(workspace.DataRoot, "library-mappings.json"), list.ToDictionary(x => x.Id, x => x.LocalFolder), ct);
        connected = true; ConnectionExpanded = false; SetStatus("Connected", list.Count, list.Count); Persist(); UpdateCommands();
    }
    private async Task ScanAsync(CancellationToken ct)
    {
        if (reader == null) throw new OperationException("InvalidProfile");
        var selected = Libraries.Where(x => x.Selected).Select(x => x.Value).ToArray();
        string site = SiteUrl, root = LocalRoot; var progress = CreateProgress();
        scan = await Task.Run(() => new DiffService(reader).ScanAsync(site, root, selected, progress, ct), ct);
        Rows.Clear(); foreach (var e in scan.Entries) Rows.Add(new(e, L, () => (OverwriteMode)ModeIndex, RefreshPlan));
        modeIndex = 0; Raise(nameof(ModeIndex));
        reportDirectory = new Workspace(site, root).DataRoot;
        if (scan.Issues.Count > 0) Error = L.Format("ScanIssues", scan.Issues.Count) + Environment.NewLine + string.Join(Environment.NewLine, scan.Issues.Take(20).Select(x => L["Error" + x.Code] + ": " + x.Path));
        if (scan.Issues.Count > 0) ErrorDetail = string.Join(Environment.NewLine + Environment.NewLine, scan.Issues.Select(x => $"{x.Code}: {x.Path}\n{x.Detail}"));
        SetStatus("ScanComplete", Rows.Count, Rows.Count); RefreshPlan();
    }
    private async Task ApplyAsync(CancellationToken ct)
    {
        if (reader == null || scan == null) return;
        var plan = CurrentPlan;
        // Clicking Apply after the explicit mode/selection and estimate is the user's authorization.
        var progress = CreateProgress();
        var result = await Task.Run(() => new SyncEngine(reader).ApplyAsync(scan, plan, progress, ct), ct);
        reportDirectory = result.RunDirectory;
        var success = result.Results.Where(r => r.Status is "Verified" or "DownloadedUnverified").Select(r => SafePaths.Key(r.Path)).ToHashSet();
        for (int i = 0; i < Rows.Count; i++)
        {
            var row = Rows[i]; if (!success.Contains(SafePaths.Key(row.Path))) continue;
            string actual = SafePaths.Resolve(scan.LocalRoot, row.Path); var hash = await FileHashing.ComputeAsync(actual, CancellationToken.None);
            var e = row.Entry;
            var installed = await ManifestStore.ReadAsync(new Workspace(scan.SiteUrl, scan.LocalRoot), e.RelativePath, CancellationToken.None);
            var latest = installed?.Remote ?? e.Remote;
            Rows[i] = new(new()
            {
                RelativePath = e.RelativePath,
                Remote = latest,
                Local = new(e.RelativePath, actual, hash, File.GetLastWriteTimeUtc(actual)),
                Status = latest?.QuickXorHash is { Length: > 0 } ? DiffStatus.Match : DiffStatus.Unverified
            }, L, () => (OverwriteMode)ModeIndex, RefreshPlan);
        }
        modeIndex = 0; Raise(nameof(ModeIndex));
        var failed = result.Results.Where(x => x.Status == "Error").ToArray();
        if (failed.Length > 0)
        {
            Error = string.Join(Environment.NewLine, failed.Take(30).Select(x => L["Error" + (x.Error ?? "TransferFailed")] + ": " + x.Path));
            string errors = Path.Combine(reportDirectory, "errors");
            ErrorDetail = Directory.Exists(errors) ? string.Join(Environment.NewLine, Directory.GetFiles(errors, "*.json").Select(File.ReadAllText)) : Error;
        }
        SetStatus(result.Cancelled ? "Cancelled" : "ApplyComplete", success.Count, result.Results.Count);
        await File.WriteAllTextAsync(Path.Combine(reportDirectory, "summary.txt"), L.Format("ReportSummary", success.Count, failed.Length, result.Results.Sum(x => x.BytesDownloaded), reportDirectory), CancellationToken.None);
        RefreshPlan();
    }
    private IProgress<OperationProgress> CreateProgress()
    {
        long last = 0;
        return new Progress<OperationProgress>(p =>
        {
            long now = Environment.TickCount64; if (now - last < 120 && p.Completed != p.Total) return; last = now;
            SetStatus(p.Code, p.Completed, p.Total); ProgressPath = p.Path ?? "";
            if (p.TotalBytes > 0) { completed = (int)Math.Min(1000, 1000d * p.Bytes / p.TotalBytes); total = 1000; Raise(nameof(ProgressPercent)); }
        });
    }
    private async Task RunAsync(Func<CancellationToken, Task> action)
    {
        if (Busy) return; Busy = true; Error = ""; ErrorDetail = ""; cancellation = new(); completed = total = 0;
        try { await action(cancellation.Token); }
        catch (OperationCanceledException) { SetStatus("Cancelled"); }
        catch (Exception ex) when (ex is OperationException or IOException or UnauthorizedAccessException or System.Net.Http.HttpRequestException)
        { SetError(ex); SetStatus("Failed"); }
        finally { Busy = false; cancellation.Dispose(); cancellation = null; Raise(nameof(Indeterminate)); RefreshPlan(); }
    }
    private void SetError(Exception ex)
    {
        string code = ex is OperationException op ? op.Code : ex is IOException ? "FileIoError" : "Unexpected";
        string action = StatusText; string? path = (ex as OperationException)?.ItemPath ?? ProgressPath;
        string? log = null;
        try { log = ErrorDetails.Save(ex, action, path); }
        catch (Exception logError) when (logError is IOException or UnauthorizedAccessException) { log = L["LogWriteFailed"] + ": " + logError.Message; }
        Error = L["Error" + code] + Environment.NewLine + ex.GetBaseException().Message + (path != null ? Environment.NewLine + path : "");
        ErrorDetail = ErrorDetails.Format(ex, action, path, log);
    }
    private void SetStatus(string code, int done = 0, int count = 0) { statusCode = code; completed = done; total = count; Raise(nameof(StatusText)); Raise(nameof(ProgressPercent)); Raise(nameof(Indeterminate)); }
    private void RefreshPlan() { if (bulkSelecting) return; Raise(nameof(Estimate)); Raise(nameof(Summary)); Raise(nameof(SummaryStats)); Raise(nameof(Details)); Raise(nameof(SelectedPath)); Raise(nameof(MetadataItems)); UpdateCommands(); }
    private void RefreshLanguage() { foreach (var r in Rows) r.RefreshText(); FilesView.Refresh(); Raise(nameof(StatusText)); RefreshPlan(); }
    private void SelectVisible(bool selected) { bulkSelecting = true; try { foreach (FileRow row in FilesView.Cast<FileRow>().Where(x => x.CanSelect)) row.Selected = selected; } finally { bulkSelecting = false; RefreshPlan(); } }
    private void UpdateCommands() { ConnectCommand?.Update(); ScanCommand?.Update(); ApplyCommand?.Update(); CancelCommand?.Update(); BrowseCommand?.Update(); AddProfileCommand?.Update(); SaveProfileCommand?.Update(); SelectVisibleCommand?.Update(); ClearVisibleCommand?.Update(); OpenReportsCommand?.Update(); }
    private async void Persist()
    {
        if (Demo) return;
        try { settings.Language = LanguageIndex == 1 ? "en-US" : "vi-VN"; settings.SiteUrl = SiteUrl; settings.LocalRoot = LocalRoot; settings.SelectedProfile = ProfileIndex; settings.Profiles = Profiles.ToList(); await SettingsStore.SaveAsync(settings); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { SetError(new OperationException("SettingsWrite", SettingsStore.PathName, ex)); }
    }
    private void OpenReports() { if (Directory.Exists(reportDirectory)) Process.Start(new ProcessStartInfo(reportDirectory) { UseShellExecute = true }); }
    public void LoadDemo()
    {
        siteUrl = "https://contoso.sharepoint.com/sites/Engineering"; localRoot = @"D:\SharePointMirror";
        Profiles.Clear(); Profiles.Add(new("Contoso • demo", "contoso.onmicrosoft.com", "00000000-0000-0000-0000-000000000000")); profileIndex = 0; LoadProfileFields();
        Libraries.Add(new(new("demo", "demo", "Engineering", "/sites/Engineering/Shared Documents", "Engineering", 10000)));
        var meta = new MetadataSnapshot(JsonSerializer.SerializeToElement(new { DocumentNumber = "DOC-2026-0042", Revision = "B", Discipline = "Civil" }), JsonFiles.Empty, JsonFiles.EmptyArray, true);
        for (int i = 0; i < 10000; i++)
        {
            string p = $"Engineering/{(i % 3 == 0 ? "Thiết kế / Design" : "Tài liệu / Documents")}/Review_{i:00000}.pdf";
            var status = (i % 10) switch { 0 => DiffStatus.MissingLocal, 1 => DiffStatus.DifferentHash, 2 => DiffStatus.LocalOnly, 3 => DiffStatus.Unverified, _ => DiffStatus.Match };
            var h = new FileHashes(42000 + i * 10, "QYqzIkdUfXTy0qQy5geE9Qe54cQ=", new string('A', 64));
            var r = new RemoteFile(siteUrl, "demo", "demo", i.ToString(), i.ToString(), p, Path.GetFileName(p), h.Size, h.QuickXorHash, "demo", DateTimeOffset.Now, DateTimeOffset.Now, "", JsonFiles.Empty, meta);
            Rows.Add(new(new() { RelativePath = p, Status = status, Remote = status == DiffStatus.LocalOnly ? null : r, Local = status == DiffStatus.MissingLocal ? null : new(p, p, h, DateTime.UtcNow) }, L, () => (OverwriteMode)ModeIndex, RefreshPlan));
        }
        ConnectionExpanded = false; SetStatus("DemoMode"); RefreshPlan();
    }
    public void Dispose() { cancellation?.Cancel(); reader?.Dispose(); }
}
public sealed record SummaryStat(string Label, int Value, string Color);
public sealed record MetadataDisplay(string Label, string Value);
