# Verification / Báo cáo kiểm chứng

Ngày kiểm tra / Verification date: **2026-11-09** (v1.1.0), Windows x64, .NET SDK **10.0.302**.

## Kết quả / Results

| Hạng mục / Check | Kết quả / Result |
|---|---|
| Release build | 0 warnings, 0 errors |
| Automated core and regression tests | **136/136 passed**, 5 consecutive runs (offline: simulated HTTP, fake clock, isolated temp folders) |
| Smoke runs (`scripts\Smoke.ps1`) | Main window `--demo --ui-smoke`: exit 0. Tray agent `--demo --agent --ui-smoke`: exit 0. `--demo --fatal-smoke`: exit 1 with log. `--demo --agent --fatal-smoke tray`: exit 1 with the deliberate-fault log. No unexpected error logs; no settings or token cache created; real app data, `*_SPDiff` folders (recursively), the demo root and the HKCU Run value unchanged |
| Tray agent checks (demo) | Summary 1,000 new / 1,000 updates / 500 locally changed / 1,000 to review; Sync plan 2,000 items, no replacement outside differences; Details opens the agent's scan in Keep-local mode (10,000 rows, 500 locally changed, plan 1,000; Replace all 2,000); every resource key resolves in VI and EN; tooltips ≤ 63 and balloons ≤ 255 characters; icon removed on exit |
| Published package | Windows x64, single executable, framework-dependent .NET 10 Desktop Runtime; adds `Microsoft.Identity.Client.Extensions.Msal` 4.92.0 (ProtectedData, Windows Forms and System.Drawing come from the Desktop Runtime) |
| Synthetic WPF dataset | **10,000 files**, VI/EN language switching preserves selection |
| Row virtualization | **9** realized rows during the measured run |
| Select visible | **5 ms** during the measured run; 2,000 eligible rows selected |
| Replace-all scope | 1,000 differences remain in the plan when the filter displays one row |
| Live SharePoint (v1.0.0, 2026-11-04) | Read-only inventory of **12 libraries**; document library **134 files / 79 folders**; complete metadata; **0 content downloads, 0 remote writes** |
| Live end-to-end on the TEST site (v1.1.0) | **Not run yet**: it needs an interactive Microsoft sign-in and Windows sign-out/sign-in |

Ảnh do smoke chụp trong phiên tự động của lần kiểm chứng này ra trắng (giới hạn chụp VisualBrush khi không có desktop tương tác; bản v1.0 cũng vậy trong cùng phiên). Các kiểm tra bố cục bằng JSON vẫn chạy và đạt. Cần tự xem thẻ thông báo bằng `SharePointDiff.exe --demo --agent`.

Smoke screenshots taken in this automated session came out blank (a VisualBrush capture limit without an interactive desktop; v1.0 behaves the same in that session). The layout checks in JSON ran and passed. Look at the card yourself with `SharePointDiff.exe --demo --agent`.

## Phạm vi kiểm thử / Test coverage

- QuickXorHash against an independent reference, empty files, block boundaries and multiple blocks.
- Vietnamese/Chinese Unicode paths, case-insensitive matching, path traversal/collisions and reuse of matching local bytes.
- Different size/hash, missing remote hash, failed library reads, metadata field aliases and typed values.
- Keep-local, individually selected and replace-all plans; local-only files remain untouched.
- Verified backups, current metadata manifests, locked files, insufficient space, corrupt downloads, source/local changes and unexpected programming faults. A programming fault cancels the other active transfer and prevents queued downloads from starting.
- Cancellation, verified cache reuse, partial downloads, server Range behavior, retry throttling and authorization-header isolation.
- **v1.1.0:**
  - **Mirror lock:** second scan or apply, two sites on one folder, leftover lock file, release after a fault; commit gate.
  - **Logon check:**
    - persisted library selection and new-library exclusion;
    - auth failure before any hashing;
    - the mirror stays byte-identical (bytes, timestamps, attributes);
    - narrowed walk; stale install temps;
    - Pristine/Revert/Conflict/Unknown/NoHash split;
    - network wait, retry budget and state classification (virtual clock).
  - **Download resilience:**
    - stall mid-body and before the first byte;
    - slow but steady bodies are not stalls; long Retry-After;
    - 8-attempt limit, then resume in a later run;
    - socket reset retried, local write failure not retried;
    - user cancellation; metadata-only ETag change;
    - 404/sign-in not retried;
    - offline end-to-end through the real reader and engine with only GET requests, including a corrupt resumed prefix and a remote metadata change.
  - **Retention:**
    - 3 newest distinct pristine versions; duplicates;
    - local edits, unknown or unverified origin kept;
    - current run, unfinished runs and interrupted replacements kept;
    - unreadable or tampered records, legacy sidecars and junctions under `runs/`;
    - long paths over 300 characters; case/NFC grouping; sequence ordering;
    - prune failures;
    - cache liveness, shared content, failed/excluded libraries, no manifest;
    - no cleanup after a cancel, stop or fault.
  - **Sign-in:**
    - DPAPI-encrypted cache round trip; per-profile file;
    - silent token shared across instances;
    - `invalid_grant`, offline and Forget; two accounts;
    - read-only allowlist for `scp` (write scopes, `Sites.Selected` and unreadable tokens fail closed); such tokens are never kept on disk and a left-over file is removed;
    - a sign-in that cannot check both Graph and SharePoint keeps nothing on disk.
  - **Settings and startup:** settings normalisation (including the selected profile), no import in agent mode, Run-command quoting and parsing, temporary locations; a blank folder is NotConfigured, not a crash; the retry budget stops waiting when the next wait would exceed it.
  - **Resource keys:** every error code and every referenced key exists in all three resource files.
- Resource-key/format parity between Vietnamese and English; synthetic WPF rendering and selection checks.

## Giới hạn kiểm chứng / Verification limits

Kiểm thử ghi đè, xoay vòng và phục hồi dùng thư mục thử biệt lập cùng HTTP giả lập. Lần tích hợp site thật (v1.0.0) chỉ đọc danh sách và metadata. Kiểm thử E2E thủ công của v1.1.0 trên site TEST (đăng nhập, đăng xuất/đăng nhập Windows, mạng chập chờn, Task Manager) chưa được thực hiện; thời gian kiểm tra khi đăng nhập trên thư mục thật chưa được đo. Hãy chạy kịch bản E2E và thử trên bản sao thư mục trước khi bật cho một bộ tài liệu mới.

Replacement, retention and recovery tests use isolated temporary folders and simulated HTTP responses. The live integration (v1.0.0) only read inventory and metadata. The manual v1.1.0 end-to-end script on the TEST site has not been run yet: interactive sign-in, Windows sign-out/sign-in, unreliable network, Task Manager. The logon-check duration on a real mirror has not been measured. Run the end-to-end script and try a copy of the folder before enabling the agent for a new document set.

Ảnh thử được dựng ở hệ số 100%, 150% và 200%; đây không thay thế việc kiểm tra chuyển cửa sổ giữa nhiều màn hình có DPI khác nhau. Thời gian đo phụ thuộc phần cứng và tải hệ thống.

Test images were rendered at 100%, 150% and 200% scale. This does not replace physical mixed-DPI/multiple-monitor testing. Timings depend on hardware and system load.

Không đưa URL site thực, token, danh sách hồ sơ hoặc nội dung metadata riêng tư vào kho mã nguồn.

No actual tenant URL, tokens, project inventory or private metadata are included in this repository.
