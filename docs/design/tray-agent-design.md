# SharePointDiff v1.1.0 — logon tray agent: full design

> Internal design record (English). Produced on 2026-11-09 by one architect and three adversarial reviewers, then revised; the user-facing summary and the approved phase plan are in Vietnamese in the session plan. Real tenant names and paths are replaced with placeholders.
>
> **Decisions made after this document was written (they override it):**
> - Section C question 1 (reverts): keep decision 3. One-click Sync also replaces files that were only edited locally; their backups are kept forever.
> - Section C question 2 (new libraries): keep the default. Libraries added after the last Compare are excluded from the logon check and reported on the card.
> - Decision 8: SharePointDiff uses its own read-only Entra app (Graph `Sites.Read.All` + SharePoint `AllSites.Read`). `TokenScopes.IsReadOnly` checks the `scp` claim against an allowlist of read-only scopes (fails closed: write scopes, `Sites.Selected` and unreadable tokens count as not read-only). Such a token is never persisted: the cache file for that profile is removed, the startup toggle is disabled and the agent reports `WriteScopedApp`. A sign-in that cannot check both the Graph and the SharePoint token keeps nothing on disk.
> - Phase order: core check, persistent auth, agent and tray, download robustness, retention, release.

# A) Revised plan: SharePointDiff v1.1.0 logon tray agent

Verified while revising:
- **Microsoft.Identity.Client.Extensions.Msal 4.92.0.** The nuspec on nuget.org has a net8.0 group requiring Microsoft.Identity.Client 4.92.0 and System.Security.Cryptography.ProtectedData 4.5.0.
- **DPAPI call.** MSAL extensions source on the main branch (`DpApiEncryptedFileAccessor`) calls Protect/Unprotect with `optionalEntropy: null` and `CurrentUser`. Re-confirm at the 4.92.0 tag.
- **Restore folder.** The last restore used the package folder `<NUGET_PACKAGES folder>`, per `obj/project.assets.json`.
- **Existing code claims that hold.** `BackupFailureKeepsOriginal` creates a file named `backups`. `--ui-smoke` never exits (App.xaml.cs:36). `columns` is a plain `Dictionary` (SharePointReader.cs:13, :99-106). The scan hashes every file under the root (DiffService.cs:15-48). The sidecar is written with an overwrite move (SyncEngine.cs:160, JsonFiles.cs:30).

## 0. Summary
1. **How the agent runs.**
   - The same `SharePointDiff.exe` starts at Windows sign-in through `HKCU\...\Run` with `--agent`, only when the settings toggle is on.
   - There is one process per settings path, guarded by a mutex. It owns a WinForms `NotifyIcon` and WPF windows.
   - It checks once per sign-in. Transient failures are retried inside a 60-minute budget, which still counts as the same sign-in check. There is also a "Check now" menu item.
2. **Gate.** The agent runs only for a site and mirror the user has already compared in the main window, meaning `<DataRoot>/library-selection.json` exists. It uses the library selection from that last completed Compare. Libraries that have appeared since then are excluded and reported.
3. **New Core pieces:**
   - `DiffService.CheckAsync`: gate, then library preflight, then a scan without a report, then a summary that splits locally changed files by provenance.
   - `DiffService.SyncAll`.
   - A mirror-level `WorkspaceLock`.
   - Downloads that resume after a stall:
     - stall and timeout detection on `TimeProvider`;
     - retry decisions based on where the failure happened;
     - resuming a `.partial` keyed on the content hash.
   - Timeout and transient failures get their own error codes.
   - A drain of in-flight commits at sign-out.
   - `Retention`:
     - backup metadata moves into its own `backup-meta/` tree;
     - the 3 newest distinct pristine versions per file are kept;
     - cache pruning runs only when its preconditions hold.
   - `AgentCheckRunner`, `StartupCommand`, and settings moved into Core so they can be tested.
4. **Auth.**
   - One MSAL core per profile per process, with the token cache persisted encrypted with DPAPI for the current user.
   - It offers separate `Background` (never interactive) and `Interactive` token providers.
   - A silent failure becomes `SignInRequired`, shown on a card the user can click.
   - The browser opens only after a user click.
5. **Nothing reaches the mirror until Sync.** Nothing is written to the mirror until the user clicks **Sync**. Sync goes through `SyncEngine.ApplyAsync`. The order of the `TransferOneAsync` chain is unchanged; only three local edits are made (§1, SyncEngine.cs).
6. **Delivery.**
   - About 70 new tests.
   - Smoke runs that exit with a status code, run from `scripts\Smoke.ps1`.
   - An end-to-end test on the TEST site, run by the user, because sign-in is interactive.
   - Bilingual docs and v1.1.0.
   - 7 phases, about 14 working days.

## 1. Core changes (file by file)

### Models.cs
```csharp
public sealed record ScanResult(..., DateTimeOffset ScannedAt, IReadOnlyList<LibraryInfo>? Libraries = null);  // set by ScanAsync
public sealed record LibrarySelection(IReadOnlyList<string> Known, IReadOnlyList<string> Included, DateTimeOffset ComparedAt);
public sealed record CheckSummary(int NewFiles, long NewBytes, int Updates, long UpdateBytes, int PristineUpdates, int Conflicts, int Reverts, int Unknown,
    int NoHash, int Review, int LocalOnly, int Reused, long NetworkBytes, long BackupBytes, int NewLibraries, int FailedLibraries, bool FirstSync,
    IReadOnlyList<string> LocallyChangedSample)
{ public int LocallyChanged => Conflicts + Reverts + Unknown; public bool HasWork => NewFiles + Updates > 0; }
public sealed record CheckOutcome(ScanResult Scan, IReadOnlyList<LibraryInfo> Libraries, IReadOnlyList<LibraryInfo> NewLibraries, CheckSummary Summary, IReadOnlySet<string> LocallyChangedKeys);
public sealed record PrunedItem(string Kind, string Path, string? RelativePath, string? Sha256, long Bytes, string? RunId, long? Sequence,
    string? DriveId, string? ItemId, string? ETag, string? WebUrl, DateTimeOffset? Modified, DateTimeOffset? BackedUpAt, string? Reason = null);
public sealed record PruneResult(IReadOnlyList<PrunedItem> Deleted, IReadOnlyList<PrunedItem> Kept, IReadOnlyList<PrunedItem> Failed, string? CacheSkippedReason);
// changed (optional trailing params keep existing call sites and tests valid):
public sealed record TransferResult(..., string? Error = null, int Attempts = 0, int Stalls = 0);
public sealed record ApplyResult(string RunDirectory, IReadOnlyList<TransferResult> Results, bool Cancelled, string? StopReason = null, PruneResult? Pruned = null);
public sealed record DownloadResult(long NetworkBytes, int Attempts = 1, int Stalls = 0);
```
`InstalledRecord` gains a trailing `string? Source = null`, set to "Downloaded", "Reused" or "Matched". It is used for audit only; pruning does not depend on it.

### Authentication.cs
```csharp
public sealed class BrowserAuthentication : ITokenProvider          // GetAsync stays interactive-capable (unchanged behaviour)
{
    public BrowserAuthentication(ConnectionProfile p, bool selectAccount = true, string? cacheDirectory = null, IMsalHttpClientFactory? http = null);
    public static BrowserAuthentication Shared(ConnectionProfile p, string? cacheDirectory); // process registry keyed tenant|client|dir (lower-invariant)
    public ITokenProvider Background { get; }   // adapter: never interactive
    public ITokenProvider Interactive => this;
    public string? CacheFilePath { get; }
    public Task SignInAsync(IReadOnlyList<string> resources, CancellationToken ct);  // user-initiated only
    public Task ForgetAsync(CancellationToken ct);
    public static StorageCreationProperties CacheProperties(string dir, ConnectionProfile p); // "msal-" + StableId(tenant|client) + ".cache"
}
```
- `GetCoreAsync(resource, ct, refresh, bool interactive)` runs inside the existing gate:
  1. **`EnsureCacheAsync`**, run once: `MsalCacheHelper.CreateAsync`, then `VerifyPersistence`, then `RegisterCache`. `MsalCachePersistenceException`, which derives from `System.Exception`, becomes `TokenCacheUnavailable`.
  2. **Account resolution.** `account ??=` the single cached account. If there are 0 or more than 1, account stays null.
  3. **Silent call.** On `MsalUiRequiredException`, re-read the accounts once and retry silent once, so an account that another provider in the same process replaced is picked up.
  4. **Non-interactive.** If no account was found, or step 3 still fails, throw `OperationException("SignInRequired", resource)`. When no account exists, no MSAL request is made at all.
  5. **Interactive.** Otherwise run today's interactive flow. After it succeeds, `RemoveAsync` every other account.
- **MSAL error mapping:**
  - `MsalException { IsRetryable: true }`, or one whose inner exception is `HttpRequestException` or a timeout `OperationCanceledException`, maps to `AuthNetworkFailed`.
  - Every other `MsalException` maps to `AuthenticationFailed`, as today.
- **`ForgetAsync`.** Takes the gate, runs `EnsureCacheAsync`, calls `RemoveAsync` on every account, then asserts `GetAccountsAsync()` is empty and sets `account = null`.
- `.WithHttpClientFactory(http)` is set only when `http` is non-null; this is the test seam.

### SharePointReader.cs
- **New init properties:**
  - `TimeProvider Time` (default `System`);
  - `StallTimeout` 60 s;
  - `HeaderTimeout` 2 min, applied to each HTTP attempt, never across Retry-After waits;
  - `ResumeAttempts` 8 (exactly 8 content attempts);
  - `RetryDelay` 5 s, doubling, capped at 60 s.
- **API timeouts.** The 2-minute timeout in `GetAsync` now runs on `Time`. `OperationCanceledException` raised while `!ct.IsCancellationRequested` becomes `OperationException("RequestTimeout", url)`, and the same applies to the header phase of `SendContentAsync`.
- **`SendAsync`:**
  - It keeps separate counters for the 401 refresh and for transient retries.
  - It retries `HttpRequestException` up to 3 times (2, 4, 8 s).
  - When 429/5xx retries run out, it throws `ServiceUnavailable` (transient). Other 4xx responses keep AccessDenied, RemoteMissing or HttpError.
- **Thread-safe column cache.** `columns` becomes `ConcurrentDictionary<string, Lazy<Task<JsonElement>>>`, and a faulted entry is removed. This fixes concurrent cold-cache `RefreshAsync` calls from the two transfer workers.
- **Inventory catch filter (line 193).** Add `&& ex is not OperationException { Code: "SignInRequired" or "TokenCacheUnavailable" or "AuthenticationFailed" or "AuthNetworkFailed" }`. Auth failures then propagate. Other per-library failures, including RequestTimeout and ServiceUnavailable, stay library-level issues keyed by `lib.Title`.
- **Partial state.** `PartialState(string ETag, long Size, string? QuickXorHash = null, string? LibraryId = null)`. Resume rule: `Size` must equal, and then:
  - if both sides have a hash, the hashes must match;
  - otherwise the ETags must match.

  Old state files without a hash fall back to the ETag rule, so `ExpiredPartialVersionIsNeverAppended` still holds.
- **`DownloadAsync`** becomes an attempt loop around `DownloadOnceAsync`:
  - **Each attempt:**
    - the offset is the `.partial` length (the first attempt applies the resume rule);
    - the state is rewritten;
    - `/content` is requested again, so the download URL is always fresh.
  - **Body phase:**
    - an idle `ITimer` from `Time.CreateTimer` cancels an idle CTS;
    - the timer is armed before the first read and re-armed after every read;
    - reads use `CreateLinkedTokenSource(ct, idle.Token)`.
  - **Classification happens at the network read call only:**
    - an OCE from the idle CTS with `!ct` becomes `DownloadStalled`;
    - an `IOException` (which includes `HttpIOException`) with `!ct` becomes `DownloadInterrupted`;
    - exceptions from the `FileStream` write propagate unchanged and are never retried.
  - **Content-host status:**
    - 401/403 becomes `DownloadInterrupted` (an expired URL, so a fresh `/content` is requested);
    - 404/410 becomes `RemoteMissing`, with no retry;
    - 429/5xx after the inner retries becomes `ServiceUnavailable`;
    - any other status becomes `InvalidDownload`.
  - **Retry** when `attempt < ResumeAttempts && !ct.IsCancellationRequested` and the code is DownloadStalled, DownloadInterrupted, RequestTimeout, ServiceUnavailable or AuthNetworkFailed. Wait the backoff via `Task.Delay(d, Time, ct)`.
  - **Not retried:** SignInRequired, AccessDenied, RemoteMissing, InvalidDownload, HashMismatch, and local I/O errors.
  - **Result.** The result carries the sum of `NetworkBytes`, plus `Attempts` and `Stalls`. `await using` keeps a valid prefix on disk, and `.partial.json` survives for resume in a later session.

### DiffService.cs
```csharp
public async Task<ScanResult> ScanAsync(string site, string root, IReadOnlyList<LibraryInfo> libraries, IProgress<OperationProgress>? progress, CancellationToken ct, bool report = true)
public async Task<IReadOnlyList<LibraryInfo>> ConnectAsync(string site, string root, CancellationToken ct)    // moved from ViewModels.cs:183-187
public async Task<CheckOutcome> CheckAsync(string site, string root, IProgress<OperationProgress>? progress, CancellationToken ct)
public static SyncPlan SyncAll(IEnumerable<DiffEntry> entries)   // Plan(entries, MissingLocal keys, OverwriteMode.AllDifferences)
public static CheckSummary Summarize(ScanResult scan, IReadOnlyDictionary<string, InstalledRecord> installed, int newLibraries, int failedLibraries, bool firstSync)
```
- **`ScanAsync`:**
  - **Lock.** `using var gate = WorkspaceLock.Acquire(workspace);` right after `new Workspace` (line 9).
  - **Narrower walk.** At the root, descend only into directories whose `SafePaths.Key` equals a selected library's `LocalFolder` key.
    - Root-level files are still hashed and become LocalOnly, as line 93 does today.
    - Other root folders are skipped entirely, including reparse points. The real mirror's local-only `_backup_*` folder and excluded libraries are therefore no longer hashed at every sign-in.
  - **Share mode.** Local files are hashed with `FileShare.ReadWrite | FileShare.Delete`, so Office saves are not blocked. The existing size and mtime post-check still yields `FileChanged`.
  - **Stale install temps.** In `Walk`, a name matching `\.spdiff-[0-9a-f]{32}\.tmp\z` (case-insensitive) does `issues.Add(new("StaleInstallTemp", rel)); continue;`. `StaleInstallTemp` is added to the code list at line 103. Such a file becomes an Error entry that is never LocalOnly, never planned and never deleted.
  - **Report flag.** `!report` skips `ReportWriter.WriteScanAsync`. The snapshot and `RefreshMatchesAsync` (source "Matched") still run. `Libraries` is set on the result.
- **`CheckAsync`:**
  1. **Gate, with no side effects.** `Workspace.DataRootFor(site, root)` (pure) leads to `LibrarySelectionStore.Read`. If it is null, throw `AgentNotReady`. No directory is created.
  2. **`ConnectAsync`** acts as the preflight, so an auth failure is thrown before any hashing.
  3. **Library sets.**
     - included = the listed libraries whose normalised id is in `Included`;
     - new = the listed libraries not in `Known`;
     - an empty included set throws `NoLibrariesSelected`.
  4. **Scan.** `ScanAsync(report: false)`.
  5. **Inputs to the summary.** `ManifestStore.ReadAllAsync`. `firstSync` = no `runs/*/results.json` exists. `failedLibraries` = included libraries missing from `RemoteFolders`.
  6. **`Summarize`.**
- **Id normalisation.** `Guid.TryParse` gives `"D"` lower-case; anything else is lower-invariant.
- **`Summarize` uses `SyncAll`.** For each Replace item, with record = installed[key]:
  - **PristineUpdate:** `record.Sha256 == local.Sha256`.
  - **Revert:** the record exists, its Sha256 differs from the local one, and `record.SharePointHash` is non-empty and equals the remote QuickXorHash. Only the local copy changed, and Sync will put back the SharePoint version.
  - **Conflict:** the record exists, its Sha256 differs from the local one, and the remote also differs from the record.
  - **Unknown:** there is no record.
  - **NoHash:** the remote has no QuickXorHash (DifferentSize with no hash). It is still in Sync (decision 3) and is shown separately.

  Also computed:
  - Review = Error + Unverified;
  - `LocallyChangedSample` = the first 5 paths;
  - `Reused`, `NetworkBytes` and `BackupBytes` come from the plan.

### Reporting.cs
- `ManifestStore.ReadAllAsync(Workspace, ct)` uses the same `Gate`.
- `SaveAsync` and `RefreshMatchesAsync` record `Source`.

### JsonFiles.cs
- `JsonFiles.WriteNewAsync<T>`: temp file, flush, then `File.Move(tmp, path, overwrite: false)`. It throws `IOException` if the target already exists.
- `Workspace` additions:
  - `MirrorDataRoot`: the `<LocalRoot>_SPDiff` sibling;
  - `LockPath = MirrorDataRoot/mirror.lock`;
  - `static string DataRootFor(string site, string root)`, which is pure and uses the same normalisation as `ScanAsync`.

### Hashing.cs
`ComputeAsync(string path, CancellationToken ct = default, FileShare share = FileShare.Read)`. Every existing caller keeps `FileShare.Read`.

### Maintenance.cs (new)
```csharp
public sealed class WorkspaceLock : IDisposable { public static WorkspaceLock Acquire(Workspace w); }
public static class BackupSequence { public static Task<long> ReserveAsync(Workspace w, int count, CancellationToken ct); }
public static class LibrarySelectionStore { public static LibrarySelection? Read(string dataRoot); public static Task WriteAsync(Workspace w, LibrarySelection s, CancellationToken ct); }
public static class Retention { public const int PristineVersionsPerFile = 3; public static Task<PruneResult> PruneAsync(Workspace w, string currentRun, ScanResult scan, int keep = PristineVersionsPerFile); }
internal sealed record BackupRecord(string? RelativePath, string? BackupPath, FileHashes? Hashes, string? Source, DateTime? OriginalCreationUtc,
    DateTime? OriginalModifiedUtc, InstalledRecord? InstalledMetadata, DateTimeOffset? BackedUpAt, long Sequence);
```

**`WorkspaceLock`**
- It opens `mirror.lock` with `FileMode.OpenOrCreate`, `FileShare.None`, and no `DeleteOnClose`, which avoids delete-pending access-denied races. A leftover lock file is harmless.
- The lock is mirror-level, so the TEST and real sites sharing one mirror exclude each other. It also excludes in-process callers, because Windows enforces sharing per handle.
- An `IOException` with HResult 0x80070020 or 0x80070021 becomes `OperationException("WorkspaceBusy", LockPath, ex)`. Other expected exceptions propagate unchanged.

**`BackupSequence`**
- Called under the lock. It reads `<DataRoot>/backup-sequence.json {Next}` and writes `Next + count` atomically.
- If the file is missing, it seeds from 1 plus the maximum `Sequence` found in `runs/*/backup-meta`.
- Ordering then never depends on the wall clock.

**`Retention`, backups.**
- **Enumeration:**
  - only `runs/*/backup-meta/*.json`;
  - `EnumerationOptions { AttributesToSkip = ReparsePoint }`;
  - `Directory.Exists` is checked per run;
  - a reparse-point run folder is skipped;
  - an `IOException` or `UnauthorizedAccessException` during one run's enumeration becomes a Failed item, never a throw.
- **Legacy sidecars.** v1.0 adjacent `*.metadata.json` sidecars are never read, so those backups are kept forever.
- **A record is eligible (counted and prunable) only if all of these hold:**
  - it deserializes with `RelativePath`, `BackupPath` and `Hashes` non-null and `Sequence > 0`; `JsonException` or missing fields mean it is kept with a reason;
  - `SafePaths.Inside(runDir/backups, BackupPath)`, and the file exists and is not a reparse point;
  - the run has `results.json`, so the run finished; crashed or in-progress runs, including a v1.0 copy running concurrently, are protected;
  - it is not an interrupted replace, meaning `intents/<id>.json` exists without `committed/<id>.json`;
  - it is pristine: `InstalledMetadata?.Sha256 == Hashes.Sha256`, and `SharePointHash` is non-empty and equals `Hashes.QuickXorHash`. This excludes local edits, unknown provenance and DownloadedUnverified provenance.
- **Selection.**
  - Group by `SafePaths.Key(RelativePath)` and order by `Sequence` descending.
  - Walk each group with a set of retained Sha256 values. A Sha256 already retained makes the record a duplicate, and so a candidate. A new Sha256 is retained while the set holds fewer than 3. Anything after that is a candidate.
  - Records from the current run count but are never candidates.
- **Deletion:**
  1. Atomically write `runs/<current>/prune-plan.json` with full tombstones: ItemId, DriveId, ETag, WebUrl, Modified, BackedUpAt and Sequence, taken from `InstalledMetadata.Remote`.
  2. For each candidate, re-hash with `FileShare.Read`. If it equals `Hashes`, delete the backup and then its record. A mismatch or `FileChanged` keeps it.
  3. Per item, `IOException` and `UnauthorizedAccessException` go to Failed.
  4. Write `pruned.json`.
  5. Write failures of either JSON file are caught (expected types only) into `PruneResult`.

**`Retention`, cache.**
- **Preconditions.** All must hold, otherwise `CacheSkippedReason` is set and nothing in the cache is deleted:
  - `installed-manifest.json` exists;
  - `scan.Libraries != null`.
- **Complete libraries.** A library is complete when `scan.RemoteFolders` contains its `LocalFolder` and there is no scan issue whose `Path == lib.Title`.
- **Live ids:**
  - `CachePath(r.Remote)` for every manifest record;
  - `CachePath(e.Remote)` for every entry in the scan;
  - every id touched by this run.
- **Eligibility.** Only names matching `^[0-9a-f]{16}\.bin(\.verified\.json|\.partial|\.partial\.json)?$` are considered.
- **Deletion.** A non-live id is deleted (its .bin, stamp, partial and partial state) only when its stamp or partial-state `LibraryId` is a complete library of this scan.
- **Kept:**
  - entries without a `LibraryId` (legacy);
  - entries of excluded, failed or unscanned libraries.
- **Orphans.** A `.partial` without `.partial.json` can never resume, so it is always deleted (under the lock nothing is writing it).
- **Left alone:** `.invalid_*`, `*.tmp`, foreign files.
- **Mirror.** The mirror is never enumerated or touched.

### AgentCheck.cs (new, Core)
```csharp
public enum AgentState { NotConfigured, Offline, SignInRequired, UpToDate, ReviewOnly, Pending, Incomplete, Failed }
public sealed record AgentCheckResult(AgentState State, CheckOutcome? Outcome, string? Code, string? Path, TimeSpan Duration, int Attempts, TimeSpan NetworkWait);
public interface INetworkProbe { Task<bool> IsOnlineAsync(CancellationToken ct); }
public sealed class HttpNetworkProbe(string tenant) : INetworkProbe;   // GET login.microsoftonline.com/{tenant}/v2.0/.well-known/openid-configuration, 10 s; any HTTP response = online
public sealed class AgentCheckRunner(IRemoteReader remote, INetworkProbe probe, TimeProvider time)
{ public Task<AgentCheckResult> RunAsync(string site, string root, bool userInitiated, IProgress<OperationProgress>? progress, CancellationToken ct); }
public static class StartupCommand { Build(exe); TryParse(value, out exe, out agent); IsCurrent(value, exe) /* GetFullPath, OrdinalIgnoreCase */; IsTemporaryLocation(exe) /* \bin\ \obj\ \publish\ \tmp\ %TEMP% */ }
```

**Waiting and retry policy**

| | Sign-in check | Manual check |
|---|---|---|
| Network wait | Poll every 15 s, deadline 15 min | Deadline 30 s |
| Total budget | 60 min | 2 min |
| Backoff | 1, 2, 5, 10, 15 min | No retry |

**Transient failures:**
- `OperationException` with code AuthNetworkFailed, RequestTimeout or ServiceUnavailable;
- `HttpRequestException`;
- an outcome where an included library has a library-level issue: ReadLibraryFailed, IncompleteInventory, RequestTimeout or ServiceUnavailable. CountChanged is not transient.

**Classification:**

| Condition | State |
|---|---|
| SignInRequired | SignInRequired |
| AgentNotReady, InvalidSite, InvalidRoot, InvalidProfile, NoLibrariesSelected, LinkBlocked | NotConfigured |
| Transient failure that is still failing when the budget runs out | Pending (with `FailedLibraries`) if there is work, otherwise Incomplete or Failed(code) |
| Any other expected exception | Failed(code) |
| `!HasWork` and Review > 0 | ReviewOnly |
| `!HasWork` and Review = 0 | UpToDate |
| User cancellation (`ct`) | Rethrown |
| Anything else (programming fault) | Propagates |

### Settings.cs (moved from App Infrastructure.cs into Core, same JSON)
- `UserSettings` gains:
  - `bool StartWithWindows`
  - `int AgentDelaySeconds = 60`, clamped to 0–900
  - `Normalize()`: a null `Profiles` list becomes an empty list, and delay values are clamped.
- `SettingsStore` members:
  - `PathName`: `Path.GetFullPath` of the environment value, or the default path;
  - `IsRedirected`;
  - `TokenCacheDirectory = <dir>\msal`;
  - `LogDirectory`;
  - `Load(string? importRoot = AppContext.BaseDirectory)` keeps today's behaviour for the main window;
  - `LoadExisting()`: never imports, returns null when the file is missing, and throws `SettingsRead` on corrupt JSON;
  - `SaveAsync`.

### SyncEngine.cs
- **Lock first.** Line 10: `using var gate = WorkspaceLock.Acquire(workspace);` before `CreateRun()`.
- **Sequence.** After the disk-space check: `long seq = await BackupSequence.ReserveAsync(workspace, plan.Replacements, ct)`.
- **`Execute`:**
  - Add `catch (OperationCanceledException) when (!workToken.IsCancellationRequested)` before the existing OCE catch. It becomes an Error result `RequestTimeout` plus an errors JSON, never "Cancelled".
  - In the expected-failure catch: `OperationException { Code: "SignInRequired" }` sets `stopReason ??= "SignInRequired"; stopOnFault.Cancel();`. This is a run-level stop: one error file, and the rest become Cancelled.
- **`TransferOneAsync`: the order is unchanged, with exactly three edits:**
  1. **Backup record.** It moves to `runs/<run>/backup-meta/<StableId(rel)>.json`, written via `WriteNewAsync` (CreateNew semantics). The new record type is `BackupRecord` with `BackupPath` and `Sequence = Interlocked.Increment(ref seq)`. A user file named `X.metadata.json` can then never be overwritten.
  2. **Cache stamp.** `CacheStamp(Sha256, ETag, LibraryId, DriveId, ItemId)`.
  3. **Commit gate.**
     - `CommitGate.Enter()`, then `try { ct.ThrowIfCancellationRequested(); File.Replace/File.Move; … committed record } finally { CommitGate.Exit(); }`.
     - The post-commit steps keep `CancellationToken.None`.
     - `ManifestStore.SaveAsync` receives the source: Reused or Downloaded.

  `CommitGate` is an in-process counter with `bool WaitIdle(TimeSpan)`.
- **Pruning at the end.** The order is: results.json, then the CSVs, then remote-metadata.csv. Pruning then runs only `if (!ct.IsCancellationRequested && stopReason == null)`:
  ```csharp
  pruned = await Retention.PruneAsync(workspace, run, scan)
  ```
  Programming faults rethrown from `Task.WhenAll` leave before this point. The return value is `ApplyResult(run, results, cancelled, stopReason, pruned)`.
- **Scope.** This applies to the main-window Apply as well.

### New OperationException codes (each needs `Error<Code>` in all three resx files)
- **Core:**
  - **Sign-in and token cache:** SignInRequired, TokenCacheUnavailable, AuthNetworkFailed.
  - **Network:** RequestTimeout, ServiceUnavailable, DownloadStalled, DownloadInterrupted.
  - **Workspace and agent:** WorkspaceBusy, StaleInstallTemp (issue code), NoLibrariesSelected, AgentNotReady.
- **App:** StartupRegistrationFailed.

## 2. App changes (file by file)

### SharePointDiff.App.csproj and app.manifest
- Add `<UseWindowsForms>true</UseWindowsForms>`.
- Add `<ItemGroup><Using Remove="System.Windows.Forms" /><Using Remove="System.Drawing" /></ItemGroup>`. This is required: the WindowsForms props add both as global usings whenever ImplicitUsings is on.
- Set `<Version>1.1.0</Version>`.
- Add `NoWarn WFAC010` only if that analyzer fires.
- In `app.manifest`, set `assemblyIdentity version="1.1.0.0"`.

### Infrastructure.cs
- `UserSettings` and `SettingsStore` are removed; they now live in Core and are used through `using SharePointDiff.Core`.

### Agent.cs (new)
**`SingleInstance`**
- Uses `new Mutex(true, "Local\\SharePointDiff-" + StableId(PathName.ToLowerInvariant()), out createdNew)`, kept in a field, plus an AutoReset `EventWaitHandle` named `…-activate`.
- **Second process:**
  - without `--agent`: `AllowSetForegroundWindow(ASFW_ANY)`, then `Set()`, then exit 0;
  - with `--agent`: exit 0 silently.
- **Owner process:**
  - uses `ThreadPool.RegisterWaitForSingleObject` → `Dispatcher.BeginInvoke(OpenMain)`;
  - ignores the activation if `App.HasFatalError` or `Dispatcher.HasShutdownStarted`;
  - unregisters the wait in `OnExit`.
- Skipped in demo.

**`StartupRegistration`**
- Run value `SharePointDiff` = `StartupCommand.Build(Environment.ProcessPath)`.
- **API:**
  - `Read()` returns the value, whether the target exists, and whether `StartupApproved\Run` disables it (read-only; an odd first byte means disabled).
  - `Apply(bool)` uses `CreateSubKey` and `SetValue`, or `DeleteValue(name, false)`.
- **Errors.** `SecurityException`, `UnauthorizedAccessException` and `IOException` become `StartupRegistrationFailed`.
- **Never called:**
  - in demo;
  - when `SettingsStore.IsRedirected`.
- It never writes `StartupApproved`.

**`ProcessMode`**
- Wraps `SetPriorityClass(GetCurrentProcess(), PROCESS_MODE_BACKGROUND_BEGIN/END)`.
- A FALSE return is a no-op that is recorded in the agent log; it never throws.
- Used only for the sign-in scan. It ends early when the user opens a window or starts a manual check.

**`TrayIcons`**
- Drawn at `GetSystemMetricsForDpi(SM_CXSMICON, GetDpiForSystem())`.
- Teal #126A70 for normal, amber #B54708 for attention.
- Created with `Icon.FromHandle(GetHicon())`, then `Clone()`, then `DestroyIcon`.
- Rebuilt on `SystemEvents.DisplaySettingsChanged` via BeginInvoke.
- Uses the aliases `Forms` and `Drawing`.

**`TrayAgent : IDisposable`**
- **Constructor.** `Forms.Application.SetUnhandledExceptionMode(ThrowException)` before any control is created. The `NotifyIcon` gets a `ContextMenuStrip`:
  - Open;
  - Check now;
  - Sign in…, visible only in the SignInRequired state;
  - Exit.

  The labels are rebuilt whenever `L.CultureName` changes.
- **Handler rule.** Every NotifyIcon or ToolStrip handler only calls `Dispatcher.BeginInvoke(() => Fire(...))`, with `private async void Fire(Func<Task> a) => await a();`.
- **Demo.**
  - It never calls `SettingsStore`, the registry or the network.
  - Check now and Sign in are disabled.
  - It shows the card from `DemoData.Scan()` with a synthetic manifest.
- **Startup, non-demo:**
  1. Settings:
     - call `LoadExisting()`;
     - `SettingsRead` shows the Problem card;
     - a null result, or `!StartWithWindows`: if the Run value points at this exe and the settings path is not redirected, remove the value, then `Shutdown(0)`.
  2. The delayed check uses a CTS that "Check now" cancels.
  3. Then `CheckAsync(userInitiated: false)`.
- **`CheckAsync(userInitiated)`:**
  1. Busy guard.
  2. Reload settings with `LoadExisting` and call `L.Switch`.
  3. Get the session with `BrowserAuthentication.Shared(profile, SettingsStore.TokenCacheDirectory)` and build `new SharePointReader(session.Background)`.
  4. Create the `Progress<OperationProgress>` on the dispatcher before `Task.Run`. Tooltip updates are throttled to 1 s.
  5. Run the check: `using (logon ? new ProcessMode() : null) result = await Task.Run(() => runner.RunAsync(...))`.
  6. Render the result:

  | State | Shown as |
  |---|---|
  | UpToDate, ReviewOnly | Tooltip only |
  | Pending | Prompt card |
  | SignInRequired | Amber icon, Sign-in card (non-activated) and balloon |
  | Incomplete, Failed, NotConfigured | Problem card with Retry and Open |

  7. Write `logs/agent-<ts>.json` with state, network wait, scan duration, hashed-file count, attempts and codes. Keep the newest 60 of these logs; this tool's own log files are not backups.
  8. Catches:
     - `OperationCanceledException` from the user's token;
     - `OperationException`, `IOException`, `UnauthorizedAccessException`, `HttpRequestException`: `ErrorDetails.Save`, then the Problem card.

     Nothing else is caught.
- **`SignInAsync`.**
  - `session.SignInAsync([Graph, siteAuthority], ct)`, the only interactive path, then `CheckAsync(true)`.
  - Triggered only by the card button, a balloon click, a left-click in the SignInRequired state, or the menu.
- **`SyncAsync(outcome)`:**
  1. `plan = DiffService.SyncAll(scan.Entries)`.
  2. `ReportWriter.WriteScanAsync`.
  3. Switch the card to its progress view.
  4. `SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)` on the UI thread, reset in `finally`.
  5. `await Task.Run(() => new SyncEngine(reader).ApplyAsync(...))`.
  6. Write `summary.txt` exactly as `MainViewModel.ApplyAsync` does.
  7. Balloon done, cancelled or stopped; clicking it opens `RunDirectory`.
  8. `StopReason == "SignInRequired"` switches to the Sign-in card.
  9. Wrapped in the same expected-type catch as `CheckAsync`. A `WorkspaceBusy` error is therefore shown on the card and never becomes fatal.
- **Left-click on the icon:**
  - in the SignInRequired state: sign in;
  - with a pending scan: re-show the card (no re-scan; it shows "Scanned at HH:mm" and a "Check again" button);
  - otherwise: Open.
- **`OpenMain(outcome?)`:**
  - reuses `Windows.OfType<MainWindow>()`, or creates `new MainViewModel(demo)` and calls `vm.LoadAgentScan(outcome)`;
  - then Show, restore from minimized, `Activate()`;
  - ends `ProcessMode`.
- **Exit.** Cancel the agent and `vm.CancelCommand`, wait for the tasks (at most 30 s) and `CommitGate.WaitIdle(4 s)`, then `Dispose()` and `Shutdown(0)`.
- **`OnSessionEnding`.**
  - Cancel everything, then call `CommitGate.WaitIdle(4 s)` synchronously inside the handler. WPF calls `Shutdown()` when this handler returns, and the work started with `Task.Run` then dies.
  - Then dispose the icon. Never set `Cancel`.
- **Text limits.**
  - Tooltip: `"SharePoint Diff\n" + text`, at most 63 characters.
  - Balloon: title at most 63 characters, text at most 255.

### AgentCard.cs (new; one code-built window, styled like `ErrorDialog`)
- **Window.** About 480×360, bottom-right of `SystemParameters.WorkArea` with a 16 px margin. `Topmost=true`, `ShowActivated=false`, `ShowInTaskbar=true`.
- **Prompt mode:**
  - lines:
    - site and folder;
    - Scanned at;
    - New (count and size), Updates (count and size);
    - locally changed K, split into conflicts, reverts and unknown, in amber with `PromptWarning`, only when K > 0;
    - up to 5 names plus "and N more";
    - NoHash, when > 0;
    - Review, when > 0;
    - new libraries and failed libraries;
    - FirstSync, when true;
    - Estimate;
    - `PromptNothingWritten`;
  - buttons: **Sync** (disabled in demo), **Details**, **Later**. Closing the card is the same as Later.
- **Progress mode:** a progress bar, the `Applying` text, the path, **Stop** and **Hide**.
- **Sign-in mode:** a **Sign in** button and Close.
- **Problem mode:** the code text, **Retry**, **Open** and Close.

### App.xaml.cs
- **Arguments.** `demo`, `agent`, `--ui-smoke <dir>`, `--fatal-smoke [tray]`.
- **Single instance.** When not demo: `SingleInstance.TryClaim(signal: !agent)`. On failure: `Shutdown(0)`.
- **Agent mode:**
  - `ShutdownMode = OnExplicitShutdown`;
  - `Agent = new TrayAgent(this, demo)`;
  - `SessionEnding += Agent.OnSessionEnding`;
  - smoke: `AgentSmoke.RunAsync`, then `Shutdown(passed ? 0 : 3)`;
  - `--fatal-smoke tray`: a hidden test-only menu item's `PerformClick()` throws through `Fire`;
  - otherwise `Agent.StartAtLogon()`.
- **Non-agent mode.** Today's path, plus activation → `OpenMain`. `--ui-smoke` now ends with `Shutdown(passed ? 0 : 3)`.
- **Fatal handler:**
  1. `main = Windows.OfType<MainWindow>().FirstOrDefault()`.
  2. `Agent?.Dispose()` first, so no ghost tray icon is left.
  3. Fallbacks: `vm?.L ?? Agent?.L ?? new()`.
  4. Log; show the `ErrorDialog` with `Topmost` in agent mode (skipped for `--fatal-smoke`).
  5. `Shutdown(1)`. There is no auto-restart.
- **Last-resort handler.** `AppDomain.UnhandledException` does a best-effort `ErrorDetails.Save` and swallows nothing.
- **`OnExit`.** `Agent?.Dispose()`, unregister the wait handle, release the mutex.

### ViewModels.cs
- **`DemoData`.** `internal static class DemoData { ScanResult Scan(); IReadOnlyDictionary<string, InstalledRecord> Manifest(); }`. Extracted from `LoadDemo`, and the manifest matches half of the DifferentHash rows.
- **Reader.** `reader = new(BrowserAuthentication.Shared(profile, SettingsStore.TokenCacheDirectory).Interactive)`. The main window may open a browser; every such path starts from a user click.
- **Connect.**
  - `ConnectAsync` calls `DiffService.ConnectAsync`.
  - Library rows get `Selected = known ? included : true`, set before the handler is attached.
- **After a successful Compare:**
  - `LibrarySelectionStore.WriteAsync`: Known = old ∪ listed; Included = selected ∪ (old included − listed).
  - Compute `LocallyChangedKeys` via `ReadAllAsync` + `Summarize`.
- **Filter.** `StatusFilterIndex` 6 = "Locally changed".
- **`LoadAgentScan(CheckOutcome o)`:**
  1. Set the `siteUrl` and `localRoot` fields directly and raise change notifications; do not call `InvalidateConnection`.
  2. `ShowLibraries`, then build the rows.
  3. `connected = true`, `ConnectionExpanded = false`.
  4. `modeIndex = 0` (Keep local, deliberately). The card's Sync stays the one-click path; Details is for review.
  5. `reportDirectory = DataRoot`; `SetStatus("ScanComplete")`.
- **`StartWithWindows` property:**
  - enabled only when `!Demo`, `!SettingsStore.IsRedirected`, and `library-selection.json` exists for the current site and folder (checked via `DataRootFor`, with no side effects);
  - applying it calls `StartupRegistration.Apply`, then `Persist()`.
  - The window shows:
    - the registered path;
    - a mismatch or missing-target warning with a **Use this copy** button;
    - a "disabled in Task Manager" note;
    - a warning when registering from a temporary location.
- **`ForgetSignInCommand`.** `RunAsync` → `Shared(...).ForgetAsync(ct)`, dispose the reader, `connected = false`, `SetStatus("SignInForgotten")`. Can execute when `!Busy && !Demo && ProfileIndex >= 0`.
- **Persist.** `Persist` writes `StartWithWindows`.

### MainWindow.xaml
- In the right column under `ReadOnlyNotice`, add:
  - the `StartWithWindows` checkbox, its hint or reason text, and the registered-path warning;
  - a **Use this copy** button;
  - a **Forget sign-in** button.
- Add a status-filter item for Locally changed.
- Change `BackupNotice` to: 3 SharePoint versions kept per file; local edits kept permanently.

### UiSmoke.cs
- `Save` becomes internal.
- Add `AgentSmoke` (§6).
- Both smoke runs return a pass flag.

### Localization (all three resx files; identical `{n}` tokens in en-US and vi-VN)

| Group | Keys |
|---|---|
| Tray | Open, Check now, Sign in, Exit |
| Tooltips | WaitingNetwork, Checking `{0:N0}{1:N0}`, UpToDate `{0:HH:mm}`, ReviewOnly, Pending, Retrying `{0:HH:mm}`, Syncing, Offline, SignInNeeded, NotConfigured, Stopping |
| Balloons | sign-in, check failed, done `{0:N0}{1:N0}`, cancelled, stopped for sign-in |
| Card | Title, Site, Folder, ScannedAt, New, Updates, Conflicts, Reverts, Unknown, NoHash, Review, NewLibraries, FailedLibraries, FirstSync, MoreFiles, Estimate, Warning, NothingWritten; buttons Sync, Details, Later, Hide, Stop, SignIn, Retry, Close, CheckAgain |
| Main window | StartWithWindows, Hint, NeedsCompare, Redirected, RegisteredPath, PathMismatch, UseThisCopy, DisabledInTaskManager, TemporaryLocation, ForgetSignIn, SignInForgotten, FilterLocallyChanged |
| Errors | the 12 new codes |
| Changed values | BackupNotice |

- About 75 keys. Tests and smoke derive the key list from the resx files; nothing hard-codes a count.
- Existing keys reused: `ReadLibrary`, `Hashing`, `Downloading`, `Applying`, `ReportSummary`.

## 3. Package and dependency changes
- **Core.** `src/SharePointDiff.Core/SharePointDiff.Core.csproj` adds `Microsoft.Identity.Client.Extensions.Msal` 4.92.0. Its net8.0 dependencies (MSAL 4.92.0, ProtectedData 4.5.0) are verified on nuget.org. Do not pin ProtectedData.
- **Tests only.** `tests/...csproj` adds `Microsoft.Extensions.TimeProvider.Testing`, the latest stable version, pinned. It is not shipped, so it does not go in the notices.
- **Lock files.**
  - Run a normal `dotnet restore SharePointDiff.sln`. Network access to nuget.org is required. The previous restore used package folder `<NUGET_PACKAGES folder>`; either folder works because the lock files are keyed by content hash.
  - Then `dotnet restore --locked-mode` must pass.
  - Do not hand-predict the App lock diff: the Desktop runtime overrides ProtectedData (10.0.10), so .NET 10 pruning may drop it from App.
  - Commit as `build: add MSAL token cache extension`.
- **Publish check.** Run `dotnet publish ... -p:PublishSingleFile=false -o tmp\publish-check`. Expected: Extensions.Msal present and no ProtectedData DLL. If one appears, add it to the notices.
- **`THIRD-PARTY-NOTICES.md`.** Add Extensions.Msal 4.92.0 (MIT), and one sentence that WinForms, System.Drawing and ProtectedData come from the installed .NET Desktop Runtime.
- **Packaging.** The `Package.ps1` allowlist is unchanged. `docs/releases/v1.1.0.md` is required.

## 4. Documentation changes
All committed docs use placeholders (`<tenant>.sharepoint.com/sites/<site>`, `<mirror>`). Never the real host or `<mirror>`, which Verification.md:46-48 promises.

**README, VI §8 and EN "Check for updates at Windows sign-in":**
- the toggle requires one completed Compare, and takes effect at the next sign-in;
- the logon flow: delay, network wait, silent token, background scan of the selected libraries only;
- retries within 60 min; no re-check on unlock or resume;
- every card field, including the conflict/revert/unknown split, with the warning that Sync also replaces local edits after a verified backup;
- new libraries are excluded until a Compare;
- Details opens in Keep-local mode;
- Later, Check now, Exit, and the sign-in card (never an automatic browser);
- pin the tray icon (Settings > Personalization > Taskbar > Other system tray icons);
- keep the exe in a stable folder such as `%LOCALAPPDATA%\Programs\SharePointDiff`; use "Use this copy" after moving it; exit the agent before republishing; unblock a downloaded ZIP;
- how to turn it off, and how a Task Manager disable is shown;
- a sign-out cancelled by another app also ends the agent until the next sign-in.

**README §3 / EN Connect.** Replace the "token only in memory" (line 56) and "tokens are not saved" (line 170) statements with:
- the token cache is encrypted with DPAPI for the current Windows user in `%LOCALAPPDATA%\SharePointDiff\msal\`;
- **Forget sign-in** removes it;
- passwords are never stored;
- the toggle is disabled when `SPDIFF_SETTINGS_PATH` is set.

**README §6 / EN Backups** (lines 124, 130, 190, 196):
- Retention rules:
  - the 3 newest distinct SharePoint versions per file are kept;
  - duplicate copies are removed;
  - local-edit, unknown-provenance and hash-less backups are kept forever;
  - only at the end of a completed, non-cancelled Apply;
  - pruned versions may be recoverable from SharePoint version history while the item and that version still exist.
- "Pristine" relies on QuickXorHash, which is not cryptographic.
- Superseded cache entries are deleted only for fully read libraries; the cache of installed versions is kept; only downloaded files are cached.
- The layout gains `mirror.lock`, `library-selection.json`, `backup-sequence.json`, `runs/<run>/backup-meta/`, `prune-plan.json` and `pruned.json`; the sidecar location moves.
- Restore: check `pruned.json` first. Backup paths can exceed 260 characters, so use `Copy-Item -LiteralPath` in PowerShell 7 or `robocopy`.

**README §7:**
- error rows for every new code;
- stall/resume behaviour; Attempts and Stalls in `results.csv`;
- timeouts are shown as errors, not cancellations;
- auth failures now abort a Compare.

**README Build and coverage (line 223):**
- `--agent`, the agent smoke flags, smoke exit codes, `scripts\Smoke.ps1`;
- coverage areas: stall/resume, retention, locking, agent check, token cache.

**CLAUDE.md:**
- Rewrite the backups/cache invariant to the exact Retention rules.
- Pipeline additions:
  - `WorkspaceLock` (mirror-level);
  - `CommitGate`;
  - the backup-meta location;
  - the narrower scan walk;
  - the code mapping for timeouts and transient failures.
- Auth: the shared session, Background vs Interactive.
- Agent architecture:
  - WinForms only for the NotifyIcon;
  - handlers only call `BeginInvoke`; `Fire` is `async void`;
  - demo never loads settings.
- Settings now live in Core.
- New flags and codes; the test count from the last run.
- The implementing agent never runs Connect, Compare, Sync or the agent against the real mirror.

**docs/Verification.md:**
- the new date and the final N/N;
- agent smoke results;
- E2E figures;
- rewrite the limits paragraph and the live rows (lines 18-20, 38-40) in VI and EN;
- the real-mirror logon-check duration, with and without background mode.

**docs/releases/v1.1.0.md, VI and EN:**
- the final N/N;
- behaviour changes: persisted token, retention, auth failures abort the scan, the narrower walk, the backup-meta location;
- remove or replace v1.0 copies (`publish\win-x64`, any extracted v1.0) before enabling the agent, because v1.0 ignores the lock.

**docs/design/auto-sync-agent-plan.md** (untracked) contradicts the final decisions. Delete it, or replace it with a scrubbed copy of this plan.

**samples/settings.example.json:** add `StartWithWindows` and `AgentDelaySeconds`.

## 5. Automated tests (xUnit, Core only)

**Fixture and helper changes:**
- `Fixture.Remote(..., string etag = "etag-1")`.
- **`FakeRemote` additions:**
  - `Libraries`, applying `maps` to `LocalFolder`;
  - `LibrariesFault`, plus per-library `IssuesByLibrary`;
  - `InventoryCalls`, `RefreshCalls`;
  - `GetInventoryAsync` filters files by the `lib.LocalFolder + "/"` prefix, which is equivalent for the existing tests.
- `SyncProgress` collector, which is synchronous.
- **New handler helpers:** async `Handler(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>>)`, `StallingStream`, `ThrowingStream(prefix, exception)`, `FakeSharePoint`.
- **Auth helpers:** a recording `IMsalHttpClientFactory` and an MSAL v3 cache seeding helper (GUID tenant; Account, IdToken, AccessToken, RefreshToken) using `SaveUnencryptedTokenCache`.
- `FakeTimeProvider`.

**AuthenticationTests**

| # | Test | Asserts |
|---|---|---|
| A1 | NonInteractiveEmptyCacheThrowsSignInRequiredWithoutHttp | 0 recorded requests |
| A2 | TokenCacheFileIsDpapiEncryptedAndRoundTrips | No plaintext marker on disk; `ProtectedData.Unprotect(null entropy, CurrentUser)` equals the payload |
| A3 | TokenCacheIsPerProfileInsideGivenDirectory | |
| A4 | SeededValidAccessTokenIsSilentAndVisibleToSecondInstance | 0 token-endpoint requests |
| A5 | InvalidGrantMapsToSignInRequiredAndTransportFailureToAuthNetworkFailed | |
| A6 | ForgetOnFreshInstanceRemovesPersistedRefreshTokenAndTwoAccountsNeedSignIn | |

**HttpReaderTests** (all use `FakeTimeProvider`)

| # | Test | Asserts |
|---|---|---|
| H1 | StalledBodyResumesWithRangeAndFreshDownloadUrl | Range 4; new Location; NetworkBytes 10; Attempts 2; Stalls 1 |
| H2 | StallBeforeFirstByteIsDetected | |
| H3 | SlowSteadyBodyIsNeverStalled | One byte per `StallTimeout − ε` |
| H4 | RetryAfterLongerThanHeaderTimeoutIsNotAStall | |
| H5 | PersistentStallFailsWithDownloadStalledThenNextReaderResumes | Exactly `ResumeAttempts` content calls |
| H6 | SocketResetIsRetriedButLocalWriteFailureIsNot | `IOException(inner SocketException 10054)` vs a read-only partial: 1 call |
| H7 | UserCancellationIsNotTreatedAsStall | |
| H8 | PartialResumesAfterMetadataOnlyETagChange | |
| H9 | ApiTimeoutBecomesRequestTimeout | |
| H10 | AuthFailuresDuringInventoryPropagateOtherFailuresStayIssues | |
| H11 | DownloadHost404AndSignInRequiredAreNotRetried | |
| H12 | ConcurrentColdRefreshFetchesSchemaOnce | |

**OfflineEndToEndTests** (FakeSharePoint handler, real SharePointReader and engine)

| # | Test | Asserts |
|---|---|---|
| O1 | CheckThenSyncAllInstallsAndIssuesOnlyGets | |
| O2 | StalledFirstAttemptResumesWithinRun | |
| O3 | CorruptResumedPrefixIsQuarantinedAndNextRunRestartsFromZero | |
| O4 | MetadataOnlyChangeAfterDownloadGivesRemoteChangedThenZeroBytes | |

**AgentCheckTests**

| # | Test | Asserts |
|---|---|---|
| C1 | SyncAllPlansAllMissingAndAllDifferencesOnly | Includes no-hash DifferentSize; equals `Plan(…, AllDifferences)` |
| C2 | SummarizeSplitsPristineConflictRevertUnknownNoHash | Includes a DownloadedUnverified record |
| C3 | CheckSignInRequiredStopsBeforeHashing | `InventoryCalls` 0; no Hashing progress; no `runs/` |
| C4 | CheckUsesPersistedSelectionAndReportsNewLibraries | GUID case-insensitive |
| C5 | NewLibraryOverLocalOnlyFolderIsNotPlanned | |
| C6 | CheckLeavesMirrorByteIdentical | Bytes, timestamps, attributes, file list; Downloads 0; RefreshCalls 0 |
| C7 | MissingSelectionIsAgentNotReadyWithoutCreatingFolders | Also: all excluded gives NoLibrariesSelected before hashing |
| C8 | ExcludedAndNonLibraryRootFoldersAreNotHashed | |
| C9 | StaleInstallTempReportedNotLocalOnlyNeverPlanned | |
| C10 | ConnectPersistsAndHonoursLibraryMappings | |
| C11 | RunnerWaitsForNetworkThenReportsOffline | |
| C12 | RunnerRetriesTransientFailuresWithinBudget | |
| C13 | RunnerClassifiesAllStatesAndRethrowsUserCancel | |
| C14 | SignInRequiredDuringApplyStopsRunOnce | |
| C15 | EngineTimeoutCancellationIsErrorNotCancelled | |
| C16 | ReportFalseStillRefreshesManifestMatches | |

**SettingsAndStartupTests**

| # | Test |
|---|---|
| S1 | OldOrNullFieldsNormalize |
| S2 | LoadExistingNeverImportsButLoadImportsFromInjectedRoot |
| S3 | StartupCommandQuotesParsesAndComparesIgnoringCase (spaces and Unicode) |
| S4 | TemporaryLocationsAreFlagged |

**RegressionTests additions**

| # | Test | Asserts |
|---|---|---|
| R1 | SecondLockFailsWithWorkspaceBusyAndWorksAfterDispose | |
| R2 | LockedWorkspaceRejectsScanAndApplyWithoutTouchingAnything | |
| R3 | TwoSitesOnOneMirrorShareTheLock | |
| R4 | LeftoverLockFileDoesNotBlockAndLockReleasedAfterFault | |
| R5 | CommitGateReturnsToZeroAfterSuccessErrorAndCancel | |
| R6 | FileAndItsMetadataJsonNamesakeBothKeepIntactBackups | |

**RetentionTests**

| # | Test | Asserts |
|---|---|---|
| T1 | KeepsThreeNewestDistinctPristineVersions | 6 cycles; asserts across all `runs/*/pruned.json`; `prune-plan.json` exists for each deletion; mirror and local-only files unchanged |
| T2 | DuplicatesDoNotEvictDistinctVersions | |
| T3 | NeverPrunesLocalEditOrUnknownBackups | |
| T4 | OnlyPristineCountTowardsLimit | |
| T5 | CurrentRunNeverPruned | |
| T6 | UnreadableIncompleteOrTamperedMetadataIsKept | `{}`, `[]`, null Hashes, invalid JSON, overwritten bytes |
| T7 | InterruptedReplaceAndUnfinishedRunAreKept | |
| T8 | LegacyAdjacentSidecarsUntouched | |
| T9 | DownloadedUnverifiedNeverPruned | |
| T10 | PruneFailureReportedSyncStillSucceeds | |
| T11 | RunEnumerationErrorsBecomeFailedItems | Existing `BackupFailureKeepsOriginal` stays unchanged as the regression check |
| T12 | JunctionUnderRunsNotFollowed | |
| T13 | LongPathBackupsArePruned | Backup path over 300 characters |
| T14 | CaseAndNfcVariantsGroupTogether | |
| T15 | SupersededCacheDeletedInstalledCurrentInvalidForeignKept | Includes an orphan partial |
| T16 | CacheOfFailedOrExcludedLibraryLegacyOrNoManifestKept | |
| T17 | SharedContentCacheKeptWhileReferenced | |
| T18 | FaultCancelAndStopReasonSkipAllCleanup | |
| T19 | OrderingUsesSequenceNotClock | |

**HashAndDiffTests additions**

| # | Test | Asserts |
|---|---|---|
| E1 | EveryErrorCodeHasLocalizedString | Every string literal used as an `OperationException`/`OperationIssue` code in Core and App, including ternaries, `Issue = new(` and `problem = new(`, across all three resx files |
| E2 | NeutralResxHasSameKeysAndTokensAsEnUs | |
| E3 | EveryReferencedResourceKeyExists | Scans App .cs/.xaml for `L[...]`, `L.Format(`, `SetStatus(`, `{Binding L[...]}` |

All 43 existing tests stay unchanged. The suite must pass 5 consecutive runs with no flakes, gated in Verification.md.

## 6. UI smoke extensions (demo, no network, no registry write, no settings)
- **Isolation.** Every smoke runs with `SPDIFF_SETTINGS_PATH=tmp\smoke-<guid>\settings.json`.
- **`scripts\Smoke.ps1`:**
  1. Creates the folder, records a before-listing (`*_SPDiff` siblings under `D:\Intertek`; the default settings file's hash and timestamp).
  2. Runs `Start-Process -Wait -PassThru` for the main smoke, the agent smoke, `--fatal-smoke` and `--agent --fatal-smoke tray`.
  3. Asserts exit codes 0/0/1/1, `Passed = true`, and that a log appears for each fatal run.
  4. Asserts the before and after listings are identical.

**`--demo --agent --ui-smoke <dir>`** writes `agent-checks.json`, then exits 0 on pass or 3 on fail.

| Check | What it asserts |
|---|---|
| TrayIconHasHandle | |
| IconSizeMatchesSmIcon | |
| SignInStateUsesDifferentIconAndShowsSignInItem | |
| Summary | New 1000, Updates 1000, LocallyChanged 500, Review 1000, SyncAllItems 2000, no Replace on Unverified/LocalOnly |
| WarningVisibleOnlyWhenK>0 | |
| ReviewLineCollapsedWhenZero | |
| CardFitsWorkArea | |
| SyncDisabledInDemo | |
| CheckNowDisabledInDemo | |
| AllKeysResolved | Every neutral-set key resolves in vi and en |
| TooltipsFit | ≤ 63 characters |
| BalloonsFit | Title ≤ 63, text ≤ 255, using the longest error string |
| ProgressWithLongVietnamesePathDoesNotClipButtons | 240-character path |
| DetailsOpensKeepLocal | `vm.Rows.Count == 10000`, `ModeIndex == 0`, `CurrentPlan.Items.Count == 1000`; with `ModeIndex = 2` it becomes 2000 |
| DemoDoesNotTouchRunKey | |
| NoSettingsOrMsalCreated | |
| IconDisposedOnExit | |

Screenshots:
- `card-vi.png`, `card-en.png`, `card-en-150.png`;
- `card-progress-en.png`, `card-signin-en.png`, `card-problem-en.png`;
- `details-en.png`.

**Existing `--demo --ui-smoke`:**
- now exits 0 or 3;
- adds `en-connection.png` (`ConnectionExpanded=true`), the `StartWithWindowsDisabledInDemo` check and the Locally-changed filter item;
- all existing checks are unchanged.

## 7. Manual end-to-end test (TEST site only; run by the user, because Microsoft sign-in and Windows sign-out are interactive)
**Setup:**
1. Publish and copy the exe to `tmp\e2e\bin\`, then a second time to `tmp\e2e\bin with space\`.
2. Mirror: `tmp\e2e\mirror`.
3. Create `%LOCALAPPDATA%\SharePointDiff\settings.json` by hand first: the TEST site, the mirror, the user's own Tenant and ClientId, and `StartWithWindows=false`. This avoids the sp_chuanhoa import.
4. In the browser, upload to the TEST site:
   - 10 PDFs;
   - a Vietnamese-named file;
   - a file with a local path over 240 characters;
   - a file of at least 300 MB;
   - `edit.docx` and `b.docx`;
   - one extra library.
5. Before the sign-out tests, do one main-window Sync of half the files, so the mirror is populated.
6. The implementing agent may provide read-only evidence snippets (hash snapshot, `reg query`, JSON reads).

| ID | Steps | Expected |
|---|---|---|
| E1 | Toggle before any Compare; then Compare, untick one library, tick the toggle | Toggle disabled before the Compare. Afterwards: the Run value is the quoted path plus `--agent`; `library-selection.json` exists; the msal cache contains no plain `refresh_token` |
| E2 | Sign out and back in (icon left in the overflow) | After about 60 s the card appears. The mirror hash, timestamp and file list are identical before the check, after the card and after Later. The `runs/` count is unchanged |
| E3 | Later, then left-click the icon, then Sync | `remote-snapshot.json` unchanged (no re-scan). All files Verified; dates match SharePoint; excluded library not downloaded; long-path file present |
| E4 | Sign out and in with no changes | No card; tooltip "Up to date" |
| E5 | New version of a.pdf, Check now, Sync | Updates 1, K 0; the backup-meta record is pristine |
| E6 | Edit `edit.docx` locally and upload a new version | Conflicts 1, the file name is listed; after Sync, the backup is kept |
| E7 | Edit `b.docx` locally only | Reverts 1 shown; behaviour per question C1 |
| E8 | Repeat E5 four more times | Exactly 3 distinct a.pdf backups; `prune-plan.json` and `pruned.json` hold tombstones; the edit.docx backup is present |
| E9 | Count the cache; delete local a.pdf, Check now, Sync | Only installed versions are cached; `BytesDownloaded` 0 |
| E10 | Create `mine.txt` and a local-only root folder | Never in the card; untouched; not hashed (agent log) |
| E11 | Delete one file and rename another on SharePoint | Local copies untouched; LocalOnly plus Missing with reuse, 0 bytes |
| E12 | Disable the network adapter for 2 min during the 300 MB download | Attempts > 1, resumes with Range, Verified |
| E13 | Unplug the router's WAN link for more than 60 s, adapter up | Stalls ≥ 1, resumes |
| E14 | Sign out during a download; sign in; Sync | Resumes from `.partial`; every `intents/` record has a `committed/` record; timestamps correct; no `*.spdiff-*.tmp`, or it is reported as StaleInstallTemp |
| E15 | Kill the process in Task Manager during a large replace | Target is the old or new version, never partial; that backup is kept by the next prune |
| E16 | Edit a column in the browser during a download | RemoteChanged; the next Sync takes about 0 bytes |
| E17 | Forget sign-in, then Check now | Sign-in card, amber icon, no browser; clicking Sign in opens the browser, then the check runs and the card appears |
| E18 | Optional and disruptive: "Sign out everywhere", then Check now | Sign-in card; no browser |
| E19 | Adapter off at sign-in; on after 5 min | The check runs |
| E20 | Lock/unlock and sleep/resume | No new check: agent log count and snapshot unchanged |
| E21 | Double-click the exe; run a second `--agent` | Main window in front with focus; one process; the second `--agent` exits 0 |
| E22 | Click Compare during an agent Sync; then the reverse | WorkspaceBusy shown both ways; the app keeps running |
| E23 | Word holds edit.docx open; new remote version; Sync | Per-file FileIoError; the others succeed |
| E24 | Switch to English, then Check now | Menu, card and balloon in English |
| E25 | Exit from the tray during a Sync | `Cancelled=true`; no `pruned.json` |
| E26 | Move the exe folder | Warning shown; "Use this copy" works; a path with spaces works |
| E27 | Disable the item in Task Manager | Shown as disabled |
| E28 | Turn the toggle off; sign out and in | Run value removed; no icon |
| E29 | DPI 150% | Icon crisp; card fits |
| E30 | Check version history and Modified By of every TEST file | Only the user's own uploads |

**Cleanup.** Remove the Run value if it points at `tmp\e2e`, then delete `tmp\e2e\`.

**Real adoption, done by the user only.** From the final exe location: Connect, Compare, then Sync on the real mirror, then tick the toggle. Click Later at the first real sign-in check and record its duration.

## 8. Phasing (each phase ends with build at 0 warnings and all tests green)
1. **Core check and safety, about 2 days.**
   - Mirror lock, narrower walk and scan share mode, StaleInstallTemp, `report` flag.
   - `ConnectAsync`, `CheckAsync`, `SyncAll`, `Summarize`; LibrarySelection; `ReadAllAsync`.
   - Columns fix; timeout and transient codes; auth-code propagation.
   - Engine: timeout OCE mapping, `CommitGate`, SignInRequired run-stop.
   - VM: `DiffService.ConnectAsync`, selection writes, the Locally-changed filter.
   - Error strings and tests E1–E3.
   - Tests C1–C10, C14–C16, R1–R5, H9–H12.
   - **Checkpoint A** on the TEST site with a mirror under `tmp\` (the user signs in).
2. **Download robustness, about 1.5 days.** TimeProvider, stall and resume, read-site classification, partial state, Attempts and Stalls. Tests H1–H8 and O1–O4.
3. **Retention, about 2 days.**
   - backup-meta tree, `WriteNewAsync`, BackupSequence, cache stamps;
   - `Retention`, prune-plan and pruned JSON;
   - `BackupNotice`.
   - Tests R6 and T1–T19.
   - The CLAUDE.md invariant rewrite and README §6 land in the same commit.
4. **Persistent auth, about 1.5 days.**
   - Package, lock files, notices.
   - `BrowserAuthentication` core, registry, Background provider, Forget.
   - VM ForgetSignIn.
   - Tests A1–A6.
   - **Checkpoint B** on the TEST site (Forget, then the silent path).
5. **Agent core, about 1.5 days.** Settings moved to Core, `AgentCheckRunner`, `HttpNetworkProbe`, `StartupCommand`. Tests C11–C13 and S1–S4.
6. **Tray app, about 4 days.**
   - csproj, Agent.cs, AgentCard.cs, App.xaml.cs, ViewModels, MainWindow;
   - resx; both smokes; `Smoke.ps1`.
7. **Release, about 1.5 days plus the user's E2E.** Docs, v1.1.0, E2E, Verification.md with measured numbers.

Commits: `feat(core): …`, `feat(app): …`, `build: …`, `docs: …`.

## 9. Risks and open questions
- **Sign-in scan cost.**
  - Only selected libraries and root files are hashed. The real mirror is about 17.5 GB minus excluded or local-only folders, roughly 35–90 s on the SATA SSD, longer in background mode.
  - Measure with and without background mode; add a hash index only if it is too slow.
  - The mirror was hand-synced this morning, so the first Compare should be mostly Match. That seeds the manifest, and most later backups will be pristine.
- **Tray and notifications.**
  - Windows 11 hides new tray icons; toasts may be suppressed or not clickable. The cards are the primary signal and the README asks the user to pin the icon.
  - `ContextMenuStrip` under a WPF message pump may have focus quirks; E2E checks it.
- **Sign-out.**
  - The commit drain is limited to 4 s inside the `SessionEnding` handler. A replace whose post-commit re-hash takes longer keeps intents without committed; its backup is kept and the timestamps are fixed on the next Sync.
  - Work before the commit resumes later from `.partial`.
- **Token at rest.** Any process running as this Windows user can decrypt the DPAPI cache (accepted). Forget sign-in is the control. Conditional Access may force SignInRequired often.
- **MSAL lock-file contention.** In-process contention is avoided by one shared core per profile. If an external process holds the lock past MSAL's retry budget, MSAL throws `InvalidOperationException`, which the app treats as fatal. Accepted rather than broadening catches.
- **Permanent deletion.** Pruned pristine versions are recoverable only from SharePoint version history while it lasts (accepted by the user). "Pristine" is based on QuickXorHash, which is not cryptographic. It is documented.
- **Growth.** Each failing attempt on a locked file creates a new non-pristine backup, kept forever. `.invalid_*` files and `runs/` folders are not pruned. Agent logs are capped at 60.
- **Long paths.** 492 of 1,860 real files would get backup paths over 260 characters (longest about 369). The app is longPathAware; Office and Explorer may not open them, so restore with PowerShell 7 or robocopy.
- **First real sync.** There is no manifest yet for files that differ, so they are "unknown" and their backups are kept forever. The card flags FirstSync.
- **Exe location.** The exe under `publish\win-x64` is locked while the agent runs, and v1.0 copies ignore the lock. Install to a stable folder and remove v1.0 copies.
- **Stale scan.** Later followed by Sync hours afterwards is still safe (every file is re-verified). The card shows "Scanned at" and offers "Check again".
- **MSAL instance discovery.** A4 assumes no network call for `login.microsoftonline.com` (a known cloud). If discovery happens, the recording factory serves it and the test counts token-endpoint requests only.
- **DPAPI entropy.** Null entropy is confirmed on the main branch; re-check at the 4.92.0 tag before A2.

### Critical Files for Implementation
- D:\Intertek\_tools\SharePointDiff\src\SharePointDiff.Core\SyncEngine.cs
- D:\Intertek\_tools\SharePointDiff\src\SharePointDiff.Core\SharePointReader.cs
- D:\Intertek\_tools\SharePointDiff\src\SharePointDiff.Core\DiffService.cs
- D:\Intertek\_tools\SharePointDiff\src\SharePointDiff.Core\Authentication.cs
- D:\Intertek\_tools\SharePointDiff\src\SharePointDiff.App\App.xaml.cs

# B) Review findings

| # | Lens | Sev | Issue | Decision | Reason |
|---|---|---|---|---|---|
| S1 | safety | major | Cache prune deletes entries absent from a partial scan or with a missing manifest | accepted | `LibraryId` added to stamps and partial state. Deletes only in libraries read completely. Skipped when there is no manifest or no `scan.Libraries`. Legacy entries kept. |
| S2 | safety | major | Sidecar `X.metadata.json` can overwrite a backup | accepted | Verified (SyncEngine.cs:160, JsonFiles.cs:30). Records moved to `backup-meta/<id>.json` via `WriteNewAsync`; legacy sidecars never read. Test R6. |
| S3 | safety | major | Exclusion list auto-includes unknown libraries | accepted | Selection {Known, Included} stored per workspace after a completed Compare. Unknown libraries excluded and reported. GUIDs normalised. Test C5. |
| S4 | safety | major | K conflates conflicts with revert of local-only edits | partially | Split into conflict, revert and unknown, with names and a Locally-changed filter. The "open backup copy" helper is rejected; a documented restore is used instead. Whether to keep reverting is question C1. |
| S5 | safety | major | `SyncAsync` has no expected-exception handling | accepted | Same expected-type catch as `CheckAsync`. WorkspaceBusy goes to the card; E22 tests both directions. |
| S6 | safety | minor | Backups without commit evidence count toward 3 and evict versions | partially | Run must have `results.json`; interrupted replaces kept; dedupe by Sha256. A committed record is not required, because pristine copies are recoverable from SharePoint. |
| S7 | safety | minor | "Pristine" means QuickXor-equal only | partially | `Source` recorded for audit, and the QuickXor basis documented. Restricting pruning to downloaded provenance is rejected: the first Compare's Match records would make one backup per file permanent. |
| S8 | safety | minor | Enumeration `IOException` breaks `BackupFailureKeepsOriginal` | accepted | Verified. Per-run guards; Failed items; never throws. T11 plus the existing test. |
| S9 | safety | minor | No reparse or containment guards on deletes | accepted | `AttributesToSkip` ReparsePoint, `Inside()` and attribute checks. T12. |
| S10 | safety | minor | Ordering by wall-clock `BackedUpAt` | accepted | Monotonic `BackupSequence` under the lock. T19. |
| S11 | safety | minor | Audit trail lost on kill; thin tombstones | accepted | `prune-plan.json` written before deletes; full tombstones; README restore wording. |
| S12 | safety | minor | Lock is per site, so two sites share a mirror unprotected; v1.0 takes no lock | accepted | A single mirror-level `mirror.lock`, which is simpler than two locks. Release-note warning about v1.0. |
| S13 | safety | minor | No-hash DifferentSize replaced without hash verification | partially | Kept in Sync, since decision 3 covers DifferentSize. Shown as a separate NoHash count. Test C1. |
| S14 | safety | minor | Error-code test misses ternary and other codes | accepted | Literal extraction across Core and App. E1. |
| S15 | safety | minor | MSAL network failures become terminal AuthenticationFailed | accepted | `AuthNetworkFailed` mapping, retried in downloads and by the runner. |
| S16 | safety | minor | `FileShare.Read` during the long sign-in hash blocks Office saves | partially | ReadWrite and Delete sharing for scan hashing only; the backup and install verification keep `FileShare.Read`. Background mode kept but measured. |
| S17 | safety | minor | Agent could do the first Connect, Scan and manifest seeding unattended | accepted | Gate on `library-selection.json`, written only after a main-window Compare. The real-mirror first run is the user's own action. |
| R1 | runtime | major | Sign-out kills post-commit steps | accepted | `CommitGate` with a bounded 4 s wait inside `SessionEnding`. `ShutdownBlockReasonCreate` is rejected (needs an hwnd; the wait is short). |
| R2 | runtime | major | Timeouts look like cancellation; no transient retry at sign-in | accepted | RequestTimeout and ServiceUnavailable codes; runner backoff inside a 60-minute budget; HTTPS probe. The `NetworkAvailabilityChanged` re-trigger is rejected; the single budget covers it. |
| R3 | runtime | major | Retry chosen by exception type misses socket resets | accepted | Classified at the read call; header timeout per attempt; test H6. |
| R4 | runtime | major | Forget on a new instance does not touch the persisted token; multiple PCAs per process | accepted | `EnsureCache` in Forget; shared core per profile; Background and Interactive adapters; accounts re-read once. |
| R5 | runtime | major | Cross-process activation lacks foreground rights | accepted | `AllowSetForegroundWindow(ASFW_ANY)`; restore and Activate. E21. |
| R6 | runtime | major | Sign-in flow depends on a hidden tray icon and balloon clicks | accepted | Non-activated Sign-in and Problem cards; README pin step. |
| R7 | runtime | minor | Run-key self-heal can never run | accepted | Checked in the main window with a "Use this copy" action; path compare ignores case; temporary locations flagged. |
| R8 | runtime | minor | `SPDIFF_SETTINGS_PATH` not carried into the Run value | accepted | Toggle disabled when redirected; the path normalised with `GetFullPath`. |
| R9 | runtime | minor | StartupApproved disabled state not shown | accepted | Read-only detection, shown in the UI. |
| R10 | runtime | minor | `columns` Dictionary race on a cold reader | accepted | Verified (:13, :99-106). `ConcurrentDictionary` + `Lazy`. H12. |
| R11 | runtime | minor | Agent startup imports real settings; `SettingsRead` becomes fatal | accepted | `LoadExisting` never imports; `SettingsRead` is expected; Topmost fatal dialog. |
| R12 | runtime | minor | Background mode is process-wide; FALSE return unhandled | accepted | Ends when a window opens; FALSE is a no-op; measured. |
| R13 | runtime | minor | 32 px icon blurry; balloon limits unchecked | accepted | Icon sized from `SM_CXSMICON` for the current DPI; BalloonsFit smoke check. |
| R14 | runtime | minor | Mutex and activation edge cases | accepted | Owned mutex kept in a field, AutoReset event, ignore during fatal or shutdown, unregister on exit. |
| R15 | runtime | minor | App lock-file diff hand-predicted | accepted | Rely on `--locked-mode` plus the publish check. Nuspec verified. |
| R16 | runtime | minor | SignInRequired mid-Sync fails every remaining file | accepted | Run-level stop with `StopReason`. C14. |
| R17 | runtime | minor | Delayed check clobbers a manual one; Progress thread; menu language | accepted | Cancellable delay; Progress created on the dispatcher; menu relabelled. |
| T1 | tests | blocker | Demo or smoke can import real settings and touch the real mirror | accepted | Demo never loads settings; injectable import root; isolated `Smoke.ps1` with before/after listing. |
| T2 | tests | major | Agent decisions live in untestable App code | accepted | `AgentCheckRunner`, `StartupCommand` and settings moved to Core. |
| T3 | tests | major | Fake exception types hide real reset behaviour | accepted | Same as R3. H6. |
| T4 | tests | major | Timeouts reported as cancellation | accepted | RequestTimeout in the reader and the engine. H9, C15. |
| T5 | tests | major | No time seam; false-positive stall cases untested | accepted | `TimeProvider` plus `FakeTimeProvider`. H2–H4. |
| T6 | tests | major | Auth tests cover only the empty cache | partially | Seeded-cache tests A4–A6 added. The lock-contention mapping test (f) is rejected: it needs a broad `InvalidOperationException` catch, against the error model. |
| T7 | tests | major | No assertion that a check leaves the mirror untouched | accepted | C6 snapshot test; E2 on a populated mirror. |
| T8 | tests | major | No offline real-reader-plus-engine integration | accepted | O1–O4 with the FakeSharePoint handler. |
| T9 | tests | major | Retention gaps (nulls, collisions, unscanned partials, long paths) | accepted | T6–T16 and R6. |
| T10 | tests | major | Smoke misses tray faults, sign-in state, Details mode; key count wrong | accepted | Tray fatal smoke, state checks, keys derived from resx. Details stays Keep-local by design and is asserted. |
| T11 | tests | major | E2E gaps | accepted | E10–E30 added. The global sign-out step is optional because it is disruptive. |
| T12 | tests | major | Tenant and real paths in committed docs; docs to update | accepted | Placeholders only; limits paragraph rewritten; final N/N quoted. |
| T13 | tests | minor | Code and key tests incomplete; neutral resx ignored | accepted | E1–E3. |
| T14 | tests | minor | Smoke never exits; retries invisible | accepted | Verified (App.xaml.cs:36). Exit codes, `Smoke.ps1`, Attempts/Stalls columns, agent log. |
| T15 | tests | minor | Fixture details make tests vacuous | accepted | FakeRemote applies mappings and gains counters; synchronous progress; assertions across runs; exact attempt counts. |
| T16 | tests | minor | Excluded libraries still hashed | accepted | Verified (:15-48). Narrower walk; C8. |
| T17 | tests | minor | Main-window changes unexercised until phase 5 | accepted | Checkpoints A and B on the TEST site. The user does the sign-in. |
| T18 | tests | minor | K vs pristine inconsistent for DownloadedUnverified | accepted | Pinned in C2 and T9; documented. |
| M1 | runtime (missing items) | – | Sleep during long Syncs | accepted | `SetThreadExecutionState` for the duration of a Sync only. |
| M2 | tests (missing items) | – | Verify the 4.92.0 nuspec | accepted | Verified: net8.0 requires MSAL 4.92.0 and ProtectedData 4.5.0. |
| M3 | tests (missing items) | – | Measurable acceptance: flake runs, timing, browser launches | partially | 5 clean consecutive runs; timing recorded; E17 checks no browser opens. A formal traceability matrix is rejected as overhead; §5 and §7 IDs map to the decisions inline. |

# C) Questions for the user

1. **Local edit with no SharePoint change.** You edited the local copy and SharePoint has not changed since the last sync. Decision 3 says Sync puts the SharePoint version back, and your edit goes to a backup that is kept forever. The card will show these as "Reverts" with their names.
   - Default: keep decision 3.
   - Alternative: leave these files out of the one-click Sync. They stay visible under the "Locally changed" filter in Details.
2. **New libraries on the site.** By default, a library added after your last main-window Compare is left out of the sign-in check until you run Compare with it ticked. The card says "N new libraries".
   - Default: keep this.
   - Alternative: include new libraries automatically, accepting that a new library whose name matches a local-only folder could replace files in it, with backups.