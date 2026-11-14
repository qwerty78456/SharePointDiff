# Verification / Báo cáo kiểm chứng

## Bản tối ưu 1.2.0 — 13/11/2026

Mã nguồn mới được kiểm tra trên Windows x64, .NET SDK 10.0.302. Kết quả dưới đây tách riêng thử tự động, phép đo trên dữ liệu lưu sẵn và thử site thật.

| Nội dung | Kết quả / giới hạn |
|---|---|
| Kiểm thử tự động | **179/179 đạt**, gồm các kiểm thử cũ và ca mới cho nén HTTP, delta hai nguồn, bộ nhớ hash, ghi đè và phục hồi |
| Build Release | **0 cảnh báo, 0 lỗi** |
| Thư mục phẳng 3.000 tệp | 2 lượt duyệt: thư mục gốc và thư mục tài liệu; không duyệt lại thư mục theo từng tệp SharePoint |
| Apply 120 tệp | Đọc manifest 1 lần; ghi checkpoint theo đợt; một sơ đồ cột dùng chung |
| So sánh lại không đổi | Tệp có nhận dạng NTFS hợp lệ dùng hash đã lưu; đọc nội dung **0 byte** trong ca thử, không tính kiểm tra sâu riêng |
| Bắt buộc tính lại | Đọc lại đúng nội dung; cập nhật ngày tính hash thật; kiểm tra nền dùng 1 worker |
| Tệp vừa bị sửa, giữ mtime / thay nguyên tử | Không dùng nhầm hash cũ; tính lại từ nội dung |
| Hash lưu sai | Dù bước quan sát dùng hash lưu, kiểm tra nội dung trước ghi đè vẫn chặn thay tệp gốc |
| Tải mới và thay tệp 65.537 byte | Apply đọc **6S**, ghi **3S**; Scan đọc **S**. Tổng **10S**. V1.1.1 khoảng **15S** theo đếm luồng mã, không phải phép đo thiết bị |
| Dùng lại theo nhóm | Bản trên máy ở bất kỳ hàng được chọn nào có thể cấp dữ liệu cho cả nhóm; dự toán và thực hiện đều không tải nhóm đó khi bản nguồn hợp lệ |
| Đổi tên đích đúng lúc thay | Bản thực tế bị thay được giữ, báo `CommitPathConflict`; không tự khôi phục đè lên tệp có thể vừa được sửa tiếp |
| Kết thúc tiến trình cưỡng bức | **6/6 mốc đạt**: sau sao lưu, intent, thay tệp, xác minh, nhật ký hoàn tất và checkpoint. Giữ bản gốc, mở lại chỉ nhập bản ghi đã ghi bền |
| JSON gzip, Brotli, deflate | Kiểm tra qua TCP cục bộ bằng handler thật của .NET; giải nén đúng. Luồng nội dung riêng không giải nén |
| Delta | Không đổi: không gọi `/items`, không ghi lại snapshot lớn; metadata-only bắt qua list delta; đổi tên thư mục dựng lại đường dẫn con |
| Lỗi delta | Token hết hạn, thiếu cha, trang rỗng, URL lạ, cột/metadata thiếu; tối đa một lượt đọc đầy đủ lại, không nhập trạng thái lỗi |
| API list delta không hỗ trợ | Nhớ khả năng theo tài khoản/thư viện, tránh thử rồi đọc đầy đủ hai lần ở mọi lượt quét; lựa chọn đọc lại đầy đủ thử khả năng mới |
| Kiểm tra khi máy rảnh | Đọc theo khối, giới hạn ngân sách, tạm dừng và vô hiệu hash khi không khớp; không đổi ngày hash toàn tệp |
| WPF | Bốn lượt main/tray/fatal-main/fatal-tray đạt; 10.000 hàng có ảo hoá; đổi Việt–Anh và giữ lựa chọn; lọc/đổi ngôn ngữ/dừng được khi chạy công việc nền giả lập |
| Hình ảnh giao diện | Đã xem ảnh hai ngôn ngữ, cửa sổ nhỏ và ảnh kết xuất ở tỷ lệ 100–200%. Chưa nghiệm thu chuyển qua lại nhiều màn hình có DPI vật lý khác nhau |
| Cấu hình và dữ liệu thật | Bộ smoke không tạo settings/token cache và không thay mục khởi động hay thư mục dữ liệu thật |
| Site thật với 1.2.0 | **Đã thử thành công**: đối chiếu đầy đủ/delta trên thư viện nhỏ có tài liệu; tải tệp 40 byte, thay có sao lưu, dùng lại cache/local. Chi tiết và giới hạn bên dưới |

### Đo chuẩn hoá dữ liệu lưu sẵn

Dùng một bản lưu kiểm kê riêng **2.472 tệp** đã có, không tải lại dữ liệu từ SharePoint. JSON cũ **161.088.906 byte**, JSON mới **24.533.570 byte**, cộng **86.382 byte** cho hai sơ đồ cột. Tổng mới **24.619.952 byte**, giảm khoảng **84,7% dung lượng JSON lưu trên đĩa**. Đọc lại đã đối chiếu từng trường gốc, giá trị hiển thị, sơ đồ, dữ liệu Graph, đường dẫn, dung lượng và hash; không đổi kiểu dữ liệu.

Đây không phải phép đo giảm 84,7% data internet hoặc độ hao SSD. Bộ đếm ghi lượng đọc/ghi do ứng dụng yêu cầu; Windows có bộ đệm. Nén JSON không giảm tương ứng dung lượng tài liệu. Lượt Compare thủ công vẫn xuất báo cáo đầy đủ JSON/CSV; kiểm tra nền bỏ bộ báo cáo đó. Hash đã lưu và delta có thể dùng kết quả quan sát trước; trước ghi đè vẫn kiểm tra dữ liệu thật. Không dùng hard link hoặc bỏ đọc kiểm tra bản sao lưu, tệp tạm, tệp đã cài.

### Thử site thật với lõi xử lý 1.2.0

Ngày 13/11/2026, đã đăng nhập Microsoft bằng trình duyệt và thử trên site người dùng chỉ định. Dùng lõi xử lý tại commit `9f9dcda` của bản ZIP RC1 qua chương trình kiểm thử riêng; executable WPF trong ZIP được thử giao diện cách ly, không dùng nó để tự động thao tác đăng nhập. Tất cả cập nhật local nằm trong thư mục thử có tên ngẫu nhiên dưới `tmp/`. Không ghi dữ liệu lên SharePoint.

| Phép thử thực tế | Kết quả |
|---|---|
| Kết nối | Liệt kê được 6 thư viện |
| Chọn thư viện để kiểm tra delta | Thư viện nhỏ nhất có 2 thư mục và 0 tệp; không lấy lượt đó làm bằng chứng delta tài liệu. Đã thử tiếp thư viện nhỏ nhất có tài liệu: **1 tệp, 3 thư mục** |
| Đọc đầy đủ → delta → đọc đầy đủ đối chứng | Cả ba lượt không có issue; khớp đường dẫn, thư mục, dung lượng, hash, eTag, JSON Graph, giá trị metadata gốc/hiển thị và sơ đồ cột, giữ nguyên kiểu JSON |
| Delta khi không đổi | **4 yêu cầu API**: 1 drive delta, 1 list delta, 1 sơ đồ cột, 1 đếm mục; **0 `/items`**, không ghi lại snapshot lớn. Lượt đọc đầy đủ đối chứng dùng 7 yêu cầu. Không tính đăng nhập và bước kết nối chung |
| Hash local khi không đổi | Dùng hash đã lưu cho 2 tệp thử trên máy; **0 byte nội dung được đọc** trong lượt delta |
| Tải tệp được chọn | Một tệp thử **40 byte**, trạng thái `Verified`, dung lượng và QuickXorHash đúng; tệp cài được đọc lại để kiểm tra SHA-256 |
| Ghi đè bản local đã sửa cùng dung lượng | Sao lưu đúng toàn bộ bản local đã sửa; bản cài xác minh thành công, dùng cache nên tải thêm **0 byte** |
| Dùng lại local khác đường dẫn | Tách cache của tệp thử ra ngoài để buộc dùng bản local; chép về đúng chỗ, giữ nguồn và tệp local-only; tải thêm **0 byte** |
| Lượng tải nội dung | Tổng **40 byte** cho lượt tải, ghi đè và dùng lại. Lượt đọc đầy đủ/delta đối chứng không tải nội dung. Con số này không bao gồm đăng nhập, JSON hay HTTP |
| Đăng nhập | Hồ sơ có `AllSites.FullControl`; dùng phiên trong bộ nhớ theo lựa chọn của người dùng, không có đường dẫn cache đăng nhập và không lưu token |

Lượt delta đọc 34.621 byte JSON **sau giải nén**; đây không phải số byte internet. Bản gói cuối giữ nguyên mã ứng dụng/lõi đã thử, chỉ cập nhật chương trình kiểm thử và báo cáo. Các tình huống làm thay đổi dữ liệu SharePoint vẫn thử bằng giả lập, không tạo thay đổi trên site thật.

### Chưa kiểm chứng

- Delta trên site thật mới kiểm tra thư viện nhỏ đang không đổi. Thêm/sửa/xoá/đổi cha/metadata-only và token hết hạn được kiểm thử giả lập; không chủ động sửa dữ liệu SharePoint để tạo các ca đó.
- Chưa thử cắt điện thật, lỗi bộ điều khiển SSD, hoặc rút mạng vật lý. Cưỡng bức kết thúc tiến trình không chứng minh mọi tình huống mất điện.
- Chưa thử lại đăng xuất/đăng nhập Windows hoặc lưu đăng nhập bằng app Entra chỉ đọc trên tenant thật.
- Chưa thử ổ FAT/SMB thật hoặc DPI vật lý nhiều màn hình; khi không đủ thông tin nhận dạng, mã đi theo nhánh tính lại hash.

**English:** 1.2.0 passes 179 automated tests and the four isolated WPF/tray/fatal smoke runs. Six actual forced-process termination boundaries preserve originals and recover only durable commits. A private saved 2,472-file inventory shrank from 161,088,906 bytes to 24,619,952 bytes including schemas, with typed roundtrip validation. The replacement fixture measures 10 payload lengths of logical Scan+Apply I/O versus approximately 15 from the prior code-path count. These are not physical SSD or internet-bandwidth measurements. **New live tests passed** using the RC1 core at `9f9dcda` through an isolated test program: full/delta/full inventories match on a library with one file and three folders; the unchanged pass makes four API requests and no `/items` request, reuses two local hashes without content reads, and downloads no documents. The transfer tests download 40 bytes total, verify a replacement backup, reuse cache/local bytes and preserve the local-only file. Sign-in stays in memory with the user's FullControl app. WPF acceptance uses the packaged executable in isolated demo mode. Remote mutations, physical power/network loss, actual Windows sign-in, real FAT/SMB and physical multi-monitor DPI transitions remain unverified on the live environment; simulated cases are separate from these observations.

## Thử trên site thật ngày 12–13/11/2026

Đã chạy đúng executable trong ZIP phát hành **v1.1.0**, đối chiếu SHA-256 với tệp trên GitHub. Dùng site người dùng chỉ định và thư mục thử riêng; không thay tài liệu dự án hoặc dữ liệu SharePoint.

| Nội dung | Kết quả |
|---|---|
| Đăng nhập Microsoft trong ứng dụng | Thành công; kết nối **6 thư viện** |
| Lượt so sánh đầu tiên | **2.452 tệp**, tất cả có QuickXorHash và thông tin tệp đầy đủ; chưa có tệp nội dung nào trong thư mục hoặc bộ nhớ đệm |
| Tải tệp được chọn | Một tệp thử **40 byte**, kết quả `Verified`; SHA-256, dung lượng và QuickXorHash đúng |
| Sửa local nhưng giữ nguyên dung lượng | Phát hiện `DifferentHash`; chế độ không ghi đè giữ nguyên bản đã sửa |
| Ghi đè từng tệp | Chỉ thay một tệp đã chọn; bản sao lưu khớp SHA-256 của bản local đã sửa; bản cài đặt khớp bản SharePoint |
| Dùng lại bản đã tải | Lần ghi đè dùng bộ nhớ đệm, tải thêm **0 byte** |
| Dùng lại bản local ở đường dẫn khác | Đã chuyển bản trong thư mục thử sang chỗ khác và giữ bộ nhớ đệm ra ngoài để buộc dùng local; chép về đúng chỗ, giữ bản nguồn, tải thêm **0 byte** |
| Tệp chỉ có trên máy | Giữ nguyên cả nội dung, ngày tạo/sửa và thuộc tính sau lượt kiểm tra nền |
| Tải tiếp qua HTTP thật | Lõi xử lý của tag v1.1.0; chủ động ngắt luồng đọc sau **16 byte**, lượt sau gửi `Range: bytes=16-`, máy chủ trả **206**, tệp ghép xong đúng hash sau **2 lượt** |
| Agent thật sau khi đăng nhập bằng tay | Hoàn tất kiểm tra trong **20,9 giây**, băm 3 tệp local, nhận biết **2.464 tệp mới**; không tải nội dung và không thay tệp local |
| Chạy hai lượt kiểm tra cùng thư mục | Lượt thứ hai nhận `WorkspaceBusy`; không thay tệp local |
| Hồ sơ Entra đang dùng | Có `AllSites.FullControl`; người dùng cho phép thử trong bộ nhớ. Đã xác nhận không lưu thông tin đăng nhập; chưa nghiệm thu đăng nhập tự động sau khi mở lại Windows |
| Cấu hình thật | Tệp cấu hình mặc định và mục tự khởi động của Windows giữ nguyên |

Ứng dụng đã tải **40 byte nội dung**; lượt thử tải tiếp riêng tải thêm **40 byte nội dung**. Con số này không bao gồm dữ liệu đăng nhập, danh sách tệp, thông tin tệp hoặc HTTP. Không đọc nội dung để giải thích khác biệt. Không có thao tác ghi lên SharePoint.

Site tiếp tục nhận tệp mới trong lúc thử, nên các lượt quét sau có số tệp cao hơn và xuất hiện `CountChanged`. Đây không được ghi thành một thư viện rỗng hoặc một lượt so sánh hoàn toàn sạch.

Phần ngắt luồng đọc là lỗi được tạo có chủ đích trên kết nối SharePoint thật, **không phải rút mạng hoặc đổi cấu hình mạng Windows**. Không đăng xuất Windows trong phiên làm việc của người dùng. Chưa kiểm chứng bằng site thật việc tự đăng nhập sau khi khởi động lại hoặc xoá các phiên bản sao lưu cũ; phần xoay vòng vẫn được kiểm thử bằng dữ liệu thử riêng.

### Sửa sau lượt thử: v1.1.1

- Thông báo cũ về quyền quá chung chung: đã thay bằng danh sách quyền thực tế, quyền cần dùng, quyền bị từ chối và lý do; có trong cửa sổ chính, thẻ thông báo và nhật ký.
- Thông báo lỗi đã hiện không đổi ngôn ngữ: đã sửa để đổi cùng giao diện, giữ nguyên lựa chọn tệp.
- Viết lại các nhãn và thông báo tiếng Việt thường gặp; phần thông báo dài có thể cuộn để đọc hết.
- **138/138 kiểm thử tự động đạt**; bốn lượt kiểm tra giao diện/lỗi có chủ đích đạt. Kiểm tra thêm thông báo quyền bằng cả Việt–Anh, nội dung không bị cắt và không tạo cấu hình hoặc thông tin đăng nhập thật trong chế độ thử.

Live v1.1.0 testing used the actual release executable, an isolated mirror and real Microsoft sign-in. Inventory, selected download, same-size local edits, verified backups, local/cache reuse, read-only background checks and an intentional read interruption with a real HTTP 206 response were exercised. The current registration has FullControl; persistence remained disabled. Actual Windows sign-out/sign-in and live version-retention pruning were not exercised. v1.1.1 fixes the permission explanation and switching the language of an existing error; the automated suite now has 138 passing tests.

Private evidence remains outside tracked source. No customer host, file inventory, Entra IDs or token values are included here.

## Báo cáo tự động trước đó của v1.1.0

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
