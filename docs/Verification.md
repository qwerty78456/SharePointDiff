# Verification / Báo cáo kiểm chứng

Ngày kiểm tra / Verification date: **2026-11-04**, Windows x64, .NET SDK **10.0.302**.

## Kết quả / Results

| Hạng mục / Check | Kết quả / Result |
|---|---|
| Release build | 0 warnings, 0 errors |
| Automated core and regression tests | **43/43 passed** |
| Published package | Windows x64, single executable, framework-dependent .NET 10 Desktop Runtime |
| Synthetic WPF dataset | **10,000 files**, VI/EN language switching preserves selection |
| Language label and file text | Label has sufficient measured width; realized row text exists and is readable on the light theme |
| Row virtualization | **9** realized rows during the measured run |
| Select visible | **5 ms** during the measured run; 2,000 eligible rows selected |
| Replace-all scope | 1,000 differences remain in the plan when the filter displays one row |
| Unexpected exception | Deliberate fatal smoke logged diagnostics and exited with code **1** |
| Live SharePoint read-only inventory | **12 libraries**; document library: **134 files / 79 folders** |
| Live metadata | Complete readable metadata for the tested library; no column or library issues reported |
| Live file content and writes | **0 content downloads, 0 remote writes** |

Đã quan sát cửa sổ Windows thực tế sau khi sửa bố cục, nhãn ngôn ngữ và màu chữ của bảng. Dữ liệu trong ảnh giao diện thử là dữ liệu giả, không chứa hồ sơ dự án.

The native Windows window was visually inspected after the layout, language label and table text fixes. The UI dataset is synthetic and contains no project records.

## Phạm vi kiểm thử / Test coverage

- QuickXorHash against an independent reference, empty files, block boundaries and multiple blocks.
- Vietnamese/Chinese Unicode paths, case-insensitive matching, path traversal/collisions and reuse of matching local bytes.
- Different size/hash, missing remote hash, failed library reads, metadata field aliases and typed values.
- Keep-local, individually selected and replace-all plans; local-only files remain untouched.
- Verified backups, current metadata manifests, locked files, insufficient space, corrupt downloads, source/local changes and unexpected programming faults. A programming fault cancels the other active transfer and prevents queued downloads from starting.
- Cancellation, verified cache reuse, partial downloads, server Range behavior, retry throttling and authorization-header isolation.
- Resource-key/format parity between Vietnamese and English; synthetic WPF rendering and selection checks.

## Giới hạn kiểm chứng / Verification limits

Kiểm thử ghi đè và phục hồi dùng thư mục thử biệt lập cùng HTTP giả lập. Lần tích hợp site thật chỉ đọc danh sách và metadata; không dùng để nghiệm thu ghi đè tài liệu thật. Thử quy trình trên bản sao thư mục trước khi triển khai cho một bộ tài liệu mới.

Replacement and recovery tests use isolated temporary folders and simulated HTTP responses. Live integration only reads inventory and metadata; it does not validate production overwrites. Exercise the workflow on a separate test copy before adopting a new document set.

Ảnh thử được dựng ở hệ số 100%, 150% và 200%; đây không thay thế việc kiểm tra chuyển cửa sổ giữa nhiều màn hình có DPI khác nhau. Thời gian đo phụ thuộc phần cứng và tải hệ thống.

Test images were rendered at 100%, 150% and 200% scale. This does not replace physical mixed-DPI/multiple-monitor testing. Timings depend on hardware and system load.

Không đưa URL site thực, token, danh sách hồ sơ hoặc nội dung metadata riêng tư vào kho mã nguồn.

No actual tenant URL, tokens, project inventory or private metadata are included in this repository.
