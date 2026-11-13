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
    private const string Graph = "https://graph.microsoft.com";
    /// <summary>Raised after a completed Compare or Apply, so the tray agent can drop a prompt that is now out of date.</summary>
    public event Action? WorkspaceChanged;
    public LocalizationService L { get; } = new();
    public ObservableCollection<ConnectionProfile> Profiles { get; } = [];
    public ObservableCollection<LibraryRow> Libraries { get; } = [];
    public ObservableCollection<FileRow> Rows { get; } = [];
    public ICollectionView FilesView { get; }
    public string[] Languages { get; } = ["Tiếng Việt", "English"];
    public bool Demo { get; }
    private UserSettings settings;
    private SharePointReader? reader;
    private BrowserAuthentication? session;
    private CancellationTokenSource? cancellation;
    private ScanResult? scan;
    private IReadOnlySet<string> locallyChanged = new HashSet<string>();
    private bool busy, connected, bulkSelecting, connectionExpanded = true;
    private string siteUrl = "", localRoot = "", filter = "", statusCode = "Ready", progressPath = "", reportDirectory = "";
    private string tenant = "", client = "", profileName = "";
    private int languageIndex, profileIndex = -1, modeIndex, statusFilterIndex;
    private int completed, total;
    private FileRow? selectedRow;
    private string error = "";
    private Func<string>? errorFormatter;
    private string errorDetail = "";
    private bool rehashAllLocal, readAllRemote;
    private IoStatistics? statistics;
    private readonly IdleVerifier? idle;
    public MainViewModel(bool demo = false)
    {
        Demo = demo; settings = demo ? new() : SettingsStore.Load();
        idle = demo ? null : (Application.Current as App)?.IdleChecks;
        if (idle != null) idle.Changed += RefreshIdle;
        foreach (var p in settings.Profiles) Profiles.Add(p);
        siteUrl = settings.SiteUrl; localRoot = settings.LocalRoot; profileIndex = Profiles.Count > 0 ? Math.Clamp(settings.SelectedProfile, 0, Profiles.Count - 1) : -1;
        languageIndex = settings.Language == "en-US" ? 1 : 0; L.Switch(languageIndex == 1 ? "en-US" : "vi-VN");
        FilesView = CollectionViewSource.GetDefaultView(Rows); FilesView.Filter = FilterRow;
        ConnectCommand = new(() => RunAsync(ConnectAsync), () => !Busy && !Demo);
        ForgetSignInCommand = new(() => RunAsync(ForgetSignInAsync), () => !Busy && !Demo && ProfileIndex >= 0 && ProfileIndex < Profiles.Count);
        ScanCommand = new(() => RunAsync(ScanAsync), () => !Busy && connected && Libraries.Any(x => x.Selected));
        ApplyCommand = new(() => RunAsync(ApplyAsync), () => !Busy && !Demo && scan != null && CurrentPlan.Items.Count > 0);
        CancelCommand = new(() => { cancellation?.Cancel(); if (!Busy && idle != null) _ = idle.PauseAsync(user: true); }, () => Busy || idle?.Active == true);
        BrowseCommand = new(Browse, () => !Busy);
        AddProfileCommand = new(AddProfile, () => !Busy);
        SaveProfileCommand = new(SaveProfile, () => !Busy);
        SelectVisibleCommand = new(() => SelectVisible(true), () => !Busy);
        ClearVisibleCommand = new(() => SelectVisible(false), () => !Busy);
        OpenReportsCommand = new(OpenReports, () => Directory.Exists(reportDirectory));
        CopyErrorCommand = new(() => { try { Clipboard.SetText(ErrorDetail); } catch { } }, () => ErrorDetail.Length > 0);
        L.PropertyChanged += (_, _) => RefreshLanguage();
        UseThisCopyCommand = new(UseThisCopy, () => !Busy && CanStartWithWindows && settings.StartWithWindows);
        if (profileIndex >= 0) LoadProfileFields();
        if (demo) LoadDemo();
        RefreshStartup();
    }
    public string SiteUrl { get => siteUrl; set { if (Set(ref siteUrl, value)) InvalidateConnection(); } }
    public string LocalRoot { get => localRoot; set { if (Set(ref localRoot, value)) InvalidateConnection(); } }
    public string Filter { get => filter; set { if (Set(ref filter, value)) FilesView.Refresh(); } }
    public bool Busy { get => busy; private set { Set(ref busy, value); Raise(nameof(NotBusy)); UpdateCommands(); } }
    public bool NotBusy => !Busy;
    public bool RehashAllLocal { get => rehashAllLocal; set => Set(ref rehashAllLocal, value); }
    public bool ReadAllRemote { get => readAllRemote; set => Set(ref readAllRemote, value); }
    public bool IdleVerification { get => settings.IdleVerification; set { settings.IdleVerification = value; if (value) idle?.Resume(); else if (idle != null) _ = idle.PauseAsync(user: true); Persist(); Raise(); } }
    public string HashSummary => L.Format("HashSummary", Rows.Count(r => r.Entry.Local?.HashBasis == HashBasis.Saved), Rows.Count(r => r.Entry.Local?.HashBasis == HashBasis.Fresh));
    public string IoSummary => statistics == null ? "" : L.Format("IoSummary", FileRow.FormatBytes(statistics.Counters.GetValueOrDefault("Content.ReadBytes")),
        FileRow.FormatBytes(statistics.Counters.GetValueOrDefault("Metadata.WriteBytes")), statistics.Counters.GetValueOrDefault("Network.GraphRequests") + statistics.Counters.GetValueOrDefault("Network.RestRequests"));
    public string IdleInfo => idle?.LastError is { } issue ? L["Error" + issue.Code] + " " + issue.Path
        : idle?.PausedByUser == true ? L["IdlePaused"] : idle?.Active == true ? L["IdleChecking"] : idle?.LastResult is { } result ? L.Format("IdleChecked", result.VerifiedBlocks, FileRow.FormatBytes(result.ReadBytes)) : L["IdleCheckHint"];
    private void RefreshIdle()
    {
        Raise(nameof(IdleInfo)); CancelCommand.Update();
        if (idle?.LastError?.Code == "HashCacheMismatch") scan = null;
        ApplyCommand.Update();
    }
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
            var rows = new List<MetadataDisplay>();
            if (SelectedRow?.Entry.Local is { } local)
            {
                rows.Add(new(L["HashBasisLabel"], L[local.HashBasis == HashBasis.Saved ? "SavedHash" : "FreshHash"]));
                rows.Add(new(L["LastHashedAt"], local.LastHashedAt?.ToLocalTime().ToString("g", L.Culture) ?? L["UnknownTime"]));
            }
            if (SelectedRow?.Entry.Remote is not { } r) return rows;
            var m = r.Metadata;
            rows.Add(new(L["MetadataReadAt"], m.ObservedAt?.ToLocalTime().ToString("g", L.Culture) ?? L["UnknownTime"]));
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
    public string Error { get => errorFormatter?.Invoke() ?? error; private set { errorFormatter = null; Set(ref error, value); Raise(nameof(Error)); } }
    internal void SetLocalizedError(Func<string> formatter) { errorFormatter = formatter; Raise(nameof(Error)); }
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
    public AsyncCommand ForgetSignInCommand { get; }
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
        return StatusFilterIndex switch { 1 => r.Entry.Status == DiffStatus.MissingLocal, 2 => r.Entry.IsDifference, 3 => r.Entry.Status == DiffStatus.LocalOnly, 4 => r.Entry.Status is DiffStatus.Error or DiffStatus.Unverified, 5 => r.Entry.Status == DiffStatus.Match, 6 => locallyChanged.Contains(SafePaths.Key(r.Path)), _ => true };
    }
    private void InvalidateConnection() { connected = false; scan = null; Libraries.Clear(); Rows.Clear(); RefreshPlan(); RefreshStartup(); }
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
        session = BrowserAuthentication.Shared(Profiles[ProfileIndex], SettingsStore.TokenCacheDirectory);
        reader?.Dispose(); reader = new(session.Interactive);
        SetStatus("Connecting");
        // Sign in for both resources first, so the saved sign-in is checked read-only for Graph and SharePoint before anything else.
        await session.SignInAsync([Graph, new Uri(site).GetLeftPart(UriPartial.Authority)], ct);
        var list = await new DiffService(reader).ConnectAsync(site, LocalRoot, ct);
        // Libraries known from the last Compare keep their tick; new ones start ticked as before.
        LibrarySelection? previous;
        try { previous = LibrarySelectionStore.Read(Workspace.DataRootFor(site, LocalRoot)); }
        catch (OperationException) { previous = null; } // a damaged selection file is rewritten by the next Compare
        var known = (previous?.Known ?? []).Select(LibrarySelectionStore.Id).ToHashSet();
        var included = (previous?.Included ?? []).Select(LibrarySelectionStore.Id).ToHashSet();
        Libraries.Clear();
        foreach (var library in list)
        {
            string id = LibrarySelectionStore.Id(library.Id);
            var row = new LibraryRow(library) { Selected = !known.Contains(id) || included.Contains(id) };
            row.PropertyChanged += (_, _) => { scan = null; Rows.Clear(); RefreshPlan(); }; Libraries.Add(row);
        }
        connected = true; ConnectionExpanded = false; SetStatus("Connected", list.Count, list.Count); Persist(); UpdateCommands();
        // A write-capable app registration still works here, but its sign-in is never kept on disk.
        if (session.PersistenceBlocked)
        {
            var permissions = session.PermissionReports;
            SetLocalizedError(() => PermissionMessages.Format(L, permissions));
            ErrorDetail = JsonSerializer.Serialize(PermissionMessages.Audit(permissions), JsonFiles.Options);
            try { ErrorDetail += Environment.NewLine + PermissionMessages.Save(L, permissions); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { ErrorDetail += Environment.NewLine + L["LogWriteFailed"] + ": " + ex.Message; }
        }
        RefreshStartup();
    }
    private async Task ForgetSignInAsync(CancellationToken ct)
    {
        if (ProfileIndex < 0 || ProfileIndex >= Profiles.Count) throw new OperationException("InvalidProfile");
        await BrowserAuthentication.Shared(Profiles[ProfileIndex], SettingsStore.TokenCacheDirectory).ForgetAsync(ct);
        reader?.Dispose(); reader = null; InvalidateConnection(); ConnectionExpanded = true; SetStatus("SignInForgotten");
    }
    private async Task ScanAsync(CancellationToken ct)
    {
        if (reader == null) throw new OperationException("InvalidProfile");
        var selected = Libraries.Where(x => x.Selected).Select(x => x.Value).ToArray();
        string site = DiffService.NormalizeSite(SiteUrl), root = LocalRoot; var progress = CreateProgress();
        var options = new ScanOptions(RehashAllLocal, ReadAllRemote);
        scan = await Task.Run(() => new DiffService(reader).ScanAsync(site, root, selected, progress, ct, options: options), ct);
        statistics = scan.Statistics; RehashAllLocal = false; ReadAllRemote = false;
        var workspace = new Workspace(site, root);
        // A completed Compare is what the logon check reuses: which libraries exist and which are ticked.
        LibrarySelection? previous;
        try { previous = LibrarySelectionStore.Read(workspace.DataRoot); }
        catch (OperationException) { previous = null; }
        await LibrarySelectionStore.WriteAsync(workspace, LibrarySelectionStore.Merge(previous, Libraries.Select(x => x.Value), selected, DateTimeOffset.Now), ct);
        var provenance = DiffService.Classify(scan.Entries, await ManifestStore.ReadAllAsync(workspace, ct));
        locallyChanged = provenance.Where(p => p.Value != Provenance.Pristine).Select(p => p.Key).ToHashSet();
        Rows.Clear(); foreach (var e in scan.Entries) Rows.Add(new(e, L, () => (OverwriteMode)ModeIndex, RefreshPlan));
        modeIndex = 0; Raise(nameof(ModeIndex));
        reportDirectory = workspace.DataRoot;
        if (scan.Issues.Count > 0)
        {
            var issues = scan.Issues.ToArray();
            SetLocalizedError(() => L.Format("ScanIssues", issues.Length) + Environment.NewLine + string.Join(Environment.NewLine, issues.Take(20).Select(x => L["Error" + x.Code] + ": " + x.Path)));
        }
        if (scan.Issues.Count > 0) ErrorDetail = string.Join(Environment.NewLine + Environment.NewLine, scan.Issues.Select(x => $"{x.Code}: {x.Path}\n{x.Detail}"));
        WorkspaceChanged?.Invoke();
        SetStatus("ScanComplete", Rows.Count, Rows.Count); RefreshPlan(); RefreshStartup();
    }
    private async Task ApplyAsync(CancellationToken ct)
    {
        if (reader == null || scan == null) return;
        var plan = CurrentPlan;
        // Clicking Apply after the explicit mode/selection and estimate is the user's authorization.
        var progress = CreateProgress();
        var result = await Task.Run(() => new SyncEngine(reader).ApplyAsync(scan, plan, progress, ct), ct);
        statistics = result.Statistics;
        reportDirectory = result.RunDirectory;
        var completedResults = result.Results.Where(r => r.Status is "Verified" or "DownloadedUnverified").ToDictionary(r => SafePaths.Key(r.Path));
        var success = completedResults.Keys.ToHashSet();
        for (int i = 0; i < Rows.Count; i++)
        {
            var row = Rows[i]; if (!success.Contains(SafePaths.Key(row.Path))) continue;
            var e = row.Entry;
            var completed = completedResults[SafePaths.Key(row.Path)];
            var latest = completed.InstalledRemote ?? throw new InvalidOperationException("Successful transfer has no installed metadata: " + row.Path);
            Rows[i] = new(new()
            {
                RelativePath = e.RelativePath,
                Remote = latest,
                Local = completed.InstalledLocal ?? throw new InvalidOperationException("Successful transfer has no installed hashes: " + row.Path),
                Status = latest?.QuickXorHash is { Length: > 0 } ? DiffStatus.Match : DiffStatus.Unverified
            }, L, () => (OverwriteMode)ModeIndex, RefreshPlan);
        }
        // Installed files now equal SharePoint, so they are no longer "locally changed".
        locallyChanged = locallyChanged.Where(k => !success.Contains(k)).ToHashSet(); FilesView.Refresh();
        modeIndex = 0; Raise(nameof(ModeIndex));
        var failed = result.Results.Where(x => x.Status == "Error").ToArray();
        if (failed.Length > 0)
        {
            SetLocalizedError(() => string.Join(Environment.NewLine, failed.Take(30).Select(x => L["Error" + (x.Error ?? "TransferFailed")] + ": " + x.Path)));
            string errors = Path.Combine(reportDirectory, "errors");
            ErrorDetail = Directory.Exists(errors) ? string.Join(Environment.NewLine, Directory.GetFiles(errors, "*.json").Select(File.ReadAllText)) : Error;
        }
        SetStatus(result.Cancelled ? "Cancelled" : "ApplyComplete", success.Count, result.Results.Count);
        WorkspaceChanged?.Invoke();
        await File.WriteAllTextAsync(Path.Combine(reportDirectory, "summary.txt"), L.Format("ReportSummary", success.Count, failed.Length, result.Results.Sum(x => x.BytesDownloaded), reportDirectory), CancellationToken.None);
        RefreshPlan();
    }
    private IProgress<OperationProgress> CreateProgress()
    {
        return new ThrottledProgress(new Progress<OperationProgress>(p =>
        {
            SetStatus(p.Code, p.Completed, p.Total); ProgressPath = p.Path ?? "";
            if (p.TotalBytes > 0) { completed = (int)Math.Min(1000, 1000d * p.Bytes / p.TotalBytes); total = 1000; Raise(nameof(ProgressPercent)); }
        }));
    }
    private async Task RunAsync(Func<CancellationToken, Task> action)
    {
        if (Busy) return; Busy = true; Error = ""; ErrorDetail = ""; cancellation = new(); completed = total = 0;
        try { if (idle != null) await idle.PauseAsync(); await action(cancellation.Token); }
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
        SetLocalizedError(() => L["Error" + code] + Environment.NewLine + ex.GetBaseException().Message + (path != null ? Environment.NewLine + path : ""));
        ErrorDetail = ErrorDetails.Format(ex, action, path, log);
    }
    private void SetStatus(string code, int done = 0, int count = 0) { statusCode = code; completed = done; total = count; Raise(nameof(StatusText)); Raise(nameof(ProgressPercent)); Raise(nameof(Indeterminate)); }
    private void RefreshPlan() { if (bulkSelecting) return; Raise(nameof(Estimate)); Raise(nameof(Summary)); Raise(nameof(SummaryStats)); Raise(nameof(HashSummary)); Raise(nameof(IoSummary)); Raise(nameof(Details)); Raise(nameof(SelectedPath)); Raise(nameof(MetadataItems)); Raise(nameof(IdleInfo)); UpdateCommands(); }
    private void RefreshLanguage() { foreach (var r in Rows) r.RefreshText(); FilesView.Refresh(); Raise(nameof(StatusText)); Raise(nameof(Error)); RefreshPlan(); RefreshStartup(); }
    // ---- Check at Windows sign-in -------------------------------------------------------------
    public RelayCommand UseThisCopyCommand { get; }
    /// <summary>The logon check needs a completed Compare for this site and folder, real settings, and a sign-in that may be kept on disk.</summary>
    public bool CanStartWithWindows => !Demo && !SettingsStore.IsRedirected && HasSelection && !PersistenceBlocked && ProfileIndex >= 0 && ProfileIndex < Profiles.Count;
    /// <summary>Turning the option off is always possible; turning it on needs <see cref="CanStartWithWindows"/>.</summary>
    public bool CanToggleStartWithWindows => !Demo && !SettingsStore.IsRedirected && (settings.StartWithWindows || CanStartWithWindows);
    /// <summary>The shared session of the selected profile (the same one the tray agent uses) has seen a token that is not read-only.</summary>
    private bool PersistenceBlocked
    {
        get
        {
            if (ProfileIndex < 0 || ProfileIndex >= Profiles.Count) return false;
            try { return BrowserAuthentication.Shared(Profiles[ProfileIndex], SettingsStore.TokenCacheDirectory).PersistenceBlocked; }
            catch (OperationException) { return false; } // an incomplete profile is reported by Connect
        }
    }
    public bool StartWithWindows
    {
        get => settings.StartWithWindows;
        set
        {
            if (value == settings.StartWithWindows || (value && !CanStartWithWindows)) { RefreshStartup(); return; }
            try { StartupRegistration.Apply(value); settings.StartWithWindows = value; Persist(); }
            catch (OperationException ex) { SetError(ex); }
            RefreshStartup();
        }
    }
    public string StartupHint => Demo ? L["StartupDemo"] : SettingsStore.IsRedirected ? L["StartupRedirected"] : PersistenceBlocked ? PermissionMessages.Format(L, BrowserAuthentication.Shared(Profiles[ProfileIndex], SettingsStore.TokenCacheDirectory).PermissionReports)
        : !HasSelection ? L["StartupNeedsCompare"] : L["StartupHint"];
    public string StartupWarning
    {
        get
        {
            if (Demo || SettingsStore.IsRedirected || !settings.StartWithWindows) return "";
            StartupRegistration.State state;
            try { state = StartupRegistration.Read(); } catch (OperationException ex) { return L["Error" + ex.Code]; }
            if (!state.IsCurrent) return state.TargetExists && StartupCommand.TryParse(state.Value, out string other, out _) ? L.Format("StartupPathMismatch", other) : L["StartupMissingTarget"];
            if (state.DisabledByUser) return L["StartupDisabledInTaskManager"];
            return StartupCommand.IsTemporaryLocation(StartupRegistration.Exe) ? L["StartupTemporaryLocation"] : "";
        }
    }
    private bool HasSelection
    {
        get
        {
            try { return File.Exists(Path.Combine(Workspace.DataRootFor(SiteUrl, LocalRoot), LibrarySelectionStore.FileName)); }
            catch (Exception ex) when (ex is OperationException or ArgumentException or IOException) { return false; }
        }
    }
    private void UseThisCopy()
    {
        try { StartupRegistration.Apply(true); } catch (OperationException ex) { SetError(ex); }
        RefreshStartup();
    }
    private void RefreshStartup()
    {
        Raise(nameof(StartWithWindows)); Raise(nameof(CanStartWithWindows)); Raise(nameof(CanToggleStartWithWindows)); Raise(nameof(StartupHint)); Raise(nameof(StartupWarning)); UseThisCopyCommand?.Update();
    }
    private void SelectVisible(bool selected) { bulkSelecting = true; try { foreach (FileRow row in FilesView.Cast<FileRow>().Where(x => x.CanSelect)) row.Selected = selected; } finally { bulkSelecting = false; RefreshPlan(); } }
    private void UpdateCommands() { UseThisCopyCommand?.Update(); ConnectCommand?.Update(); ForgetSignInCommand?.Update(); ScanCommand?.Update(); ApplyCommand?.Update(); CancelCommand?.Update(); BrowseCommand?.Update(); AddProfileCommand?.Update(); SaveProfileCommand?.Update(); SelectVisibleCommand?.Update(); ClearVisibleCommand?.Update(); OpenReportsCommand?.Update(); }
    private async void Persist()
    {
        if (Demo) return;
        try { settings.Language = LanguageIndex == 1 ? "en-US" : "vi-VN"; settings.SiteUrl = SiteUrl; settings.LocalRoot = LocalRoot; settings.SelectedProfile = ProfileIndex; settings.Profiles = Profiles.ToList(); await SettingsStore.SaveAsync(settings); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { SetError(new OperationException("SettingsWrite", SettingsStore.PathName, ex)); }
    }
    private void OpenReports() { if (Directory.Exists(reportDirectory)) Process.Start(new ProcessStartInfo(reportDirectory) { UseShellExecute = true }); }
    public void LoadDemo()
    {
        siteUrl = DemoData.Site; localRoot = DemoData.Root;
        Profiles.Clear(); Profiles.Add(new("Contoso • demo", "contoso.onmicrosoft.com", "00000000-0000-0000-0000-000000000000")); profileIndex = 0; LoadProfileFields();
        Libraries.Add(new(DemoData.Library));
        foreach (var entry in DemoData.Scan().Entries) Rows.Add(new(entry, L, () => (OverwriteMode)ModeIndex, RefreshPlan));
        ConnectionExpanded = false; SetStatus("DemoMode"); RefreshPlan();
    }
    /// <summary>Shows a scan made by the logon check ("Details" on the card), in Keep-local mode so nothing is replaced by accident.</summary>
    public void LoadAgentScan(CheckOutcome outcome, SharePointReader? agentReader)
    {
        siteUrl = outcome.Scan.SiteUrl; localRoot = outcome.Scan.LocalRoot; Raise(nameof(SiteUrl)); Raise(nameof(LocalRoot));
        if (agentReader != null) { reader?.Dispose(); reader = agentReader; }
        Libraries.Clear();
        foreach (var library in outcome.Libraries) { var row = new LibraryRow(library); row.PropertyChanged += (_, _) => { scan = null; Rows.Clear(); RefreshPlan(); }; Libraries.Add(row); }
        scan = outcome.Scan; locallyChanged = outcome.LocallyChangedKeys;
        Rows.Clear(); foreach (var e in scan.Entries) Rows.Add(new(e, L, () => (OverwriteMode)ModeIndex, RefreshPlan));
        connected = true; ConnectionExpanded = false; modeIndex = 0; Raise(nameof(ModeIndex));
        if (!Demo) reportDirectory = Workspace.DataRootFor(scan.SiteUrl, scan.LocalRoot);
        SetStatus("ScanComplete", Rows.Count, Rows.Count); RefreshPlan(); RefreshStartup();
    }
    public void Dispose() { if (idle != null) idle.Changed -= RefreshIdle; cancellation?.Cancel(); reader?.Dispose(); }
}
public sealed record SummaryStat(string Label, int Value, string Color);
public sealed record MetadataDisplay(string Label, string Value);
/// <summary>Synthetic data shared by --demo, the UI smoke and the agent smoke (no network, no real paths).</summary>
internal static class DemoData
{
    public const string Site = "https://contoso.sharepoint.com/sites/Engineering", Root = @"D:\SharePointMirror";
    public static LibraryInfo Library => new("demo", "demo", "Engineering", "/sites/Engineering/Shared Documents", "Engineering", 10000);
    private static readonly string LocalSha = new('A', 64);
    public static ScanResult Scan()
    {
        var meta = new MetadataSnapshot(JsonSerializer.SerializeToElement(new { DocumentNumber = "DOC-2026-0042", Revision = "B", Discipline = "Civil" }), JsonFiles.Empty, JsonFiles.EmptyArray, true);
        var entries = new List<DiffEntry>();
        for (int i = 0; i < 10000; i++)
        {
            string p = $"Engineering/{(i % 3 == 0 ? "Thiết kế / Design" : "Tài liệu / Documents")}/Review_{i:00000}.pdf";
            var status = (i % 10) switch { 0 => DiffStatus.MissingLocal, 1 => DiffStatus.DifferentHash, 2 => DiffStatus.LocalOnly, 3 => DiffStatus.Unverified, _ => DiffStatus.Match };
            var h = new FileHashes(42000 + i * 10, "QYqzIkdUfXTy0qQy5geE9Qe54cQ=", LocalSha);
            var r = new RemoteFile(Site, "demo", "demo", i.ToString(), i.ToString(), p, Path.GetFileName(p), h.Size, h.QuickXorHash, "demo", DateTimeOffset.Now, DateTimeOffset.Now, "", JsonFiles.Empty, meta);
            entries.Add(new() { RelativePath = p, Status = status, Remote = status == DiffStatus.LocalOnly ? null : r, Local = status == DiffStatus.MissingLocal ? null : new(p, p, h, DateTime.UtcNow) });
        }
        return new(Site, Root, entries, ["Engineering"], [], DateTimeOffset.Now, [Library]);
    }
    /// <summary>Half of the differences were installed by an earlier sync (pristine); the other half have no record (unknown).</summary>
    public static IReadOnlyDictionary<string, InstalledRecord> Manifest(ScanResult scan) => scan.Entries
        .Where(e => e.Status == DiffStatus.DifferentHash && int.Parse(e.Remote!.ItemId) % 20 == 1)
        .ToDictionary(e => SafePaths.Key(e.RelativePath), e => new InstalledRecord(e.RelativePath, LocalSha, "old", "demo", DateTimeOffset.Now, e.Remote!, "Downloaded"));
    public static CheckOutcome Outcome()
    {
        var scan = Scan(); var manifest = Manifest(scan);
        var changed = DiffService.Classify(scan.Entries, manifest).Where(p => p.Value != Provenance.Pristine).Select(p => p.Key).ToHashSet();
        return new(scan, [Library], [], DiffService.Summarize(scan, manifest, 0, 0, false), changed);
    }
}
