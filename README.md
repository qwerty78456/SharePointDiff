# SharePoint Diff

**Tác giả / Authors:** Vũ Mạnh Quân (taladinhcao@gmail.com) & Trong-Nghia Be (benghia1st@gmail.com)

**Giấy phép / License:** [MIT](LICENSE)

Ứng dụng Windows **C# / WPF / .NET 10**, giao diện **Tiếng Việt và English**, để đối chiếu thư mục local với SharePoint Online bằng hash, tải tệp thiếu và thay các bản khác biệt theo lựa chọn của người dùng.

**Chỉ đọc SharePoint. Không upload, sửa metadata trên site, xoá hoặc di chuyển tệp chỉ có ở local. Mọi lần ghi đè local đều có bản sao lưu đã kiểm tra SHA-256.**

- [Hướng dẫn tiếng Việt](#hướng-dẫn-tiếng-việt)
- [English guide](#english-guide)
- [Build, kiểm thử và Git](#build-kiểm-thử-và-git)
- [Kiến trúc và nguồn tham khảo](#kiến-trúc-và-nguồn-tham-khảo)

## Hướng dẫn tiếng Việt

### 1. Chạy ứng dụng

Yêu cầu Windows x64 và **.NET 10 Windows Desktop Runtime**. Visual Studio không cần thiết khi chạy bản phát hành. Máy đã có .NET 10 Desktop Runtime thì không phải cài thêm Python, PowerShell, PnP.PowerShell hoặc trình biên dịch.

Mở `publish\win-x64\SharePointDiff.exe`. Có thể tạo shortcut tới tệp này theo cách thông thường của Windows.

Ứng dụng mặc định tiếng Việt. Bộ chọn **Tiếng Việt / English** ở góc trên phải đổi ngôn ngữ ngay, không mất lựa chọn tệp hoặc dừng tác vụ. Tên tệp, hash, tên cột và giá trị gốc SharePoint không bị dịch.

### 2. Chuẩn bị đăng nhập

Mỗi tenant cần một ứng dụng Entra ID đã được quản trị viên cấu hình:

| Thiết lập | Giá trị |
|---|---|
| Loại ứng dụng | Public client / Mobile and desktop application |
| Redirect URI | `http://localhost` |
| Microsoft Graph, delegated | `Sites.Read.All` |
| SharePoint, delegated | `AllSites.Read` |
| Consent | Được quản trị viên cấp theo chính sách tổ chức |
| Secret / certificate | Không cần cho đăng nhập tương tác này |

Ứng dụng dùng MSAL.NET và trình duyệt hệ thống, hỗ trợ MFA. Không dùng device-code login, không yêu cầu tắt Security Defaults và không tự đăng ký ứng dụng hay cấp quyền. Một Client ID chỉ dùng được tại tenant nơi nó được đăng ký hoặc chấp thuận phù hợp.

Nếu chạy trong cùng cây công cụ với `sp_chuanhoa\cauhinh\cauhinh.json`, lần đầu ứng dụng có thể đọc Tenant, Client ID, site và thư mục mặc định từ cấu hình có sẵn. Tệp cấu hình gốc không bị sửa. Nếu không có cấu hình đó, tạo hồ sơ mới trong giao diện.

### 3. Kết nối một site

1. Mở **Kết nối & thư mục**.
2. Dán URL của **site**, ví dụ `https://contoso.sharepoint.com/sites/Engineering`. Không dùng link chia sẻ của một tệp, trang `Forms/AllItems.aspx` hoặc link có query string.
3. Chọn thư mục local tương ứng. Thư mục phải tồn tại; chọn thư mục thông thường thay vì gốc ổ đĩa hay symbolic link/junction.
4. Chọn hồ sơ có sẵn; hoặc bấm **Thêm**, mở **Chỉnh hồ sơ đăng nhập**, nhập tên hồ sơ, Tenant ID/tên miền và Entra Client ID, rồi **Lưu hồ sơ**.
5. Bấm **Kết nối** và hoàn tất đăng nhập Microsoft. Tài khoản phải có quyền vào site; có thể cần hai lượt cấp token cho Graph và SharePoint, dùng cùng phiên đăng nhập.
6. Chọn các thư viện cần đối chiếu. Thư viện tài liệu không ẩn được chọn mặc định. Sau kết nối, phần cấu hình thu gọn để nhường chỗ cho bảng.

Hồ sơ và ngôn ngữ được lưu tại `%LOCALAPPDATA%\SharePointDiff\settings.json`. Tệp này chỉ chứa thông tin cấu hình, không chứa mật khẩu hay token. Token chỉ nằm trong bộ nhớ phiên chạy.

Cấu hình mẫu có hai tenant ở [`samples/settings.example.json`](samples/settings.example.json). Các Client ID trong mẫu là giá trị giả: thay bằng ID được cấp cho từng tenant, URL site và thư mục có thật. Có thể tạo hồ sơ bằng giao diện; nếu chỉnh JSON thủ công, đóng ứng dụng và sao lưu cấu hình hiện tại trước khi chỉnh. `SelectedProfile` là chỉ số bắt đầu từ `0`; `Language` nhận `vi-VN` hoặc `en-US`.

### 4. So sánh

Bấm **So sánh**. Bước này đọc danh sách và metadata SharePoint, tính lại QuickXorHash và SHA-256 của local; **không tải nội dung tệp SharePoint**.

Ứng dụng đọc local trong thư mục đã chọn để tìm bản có thể tái dùng. Phạm vi hiển thị/cập nhật là các thư viện đã chọn, cùng các tệp rời ở gốc local. Tệp trong thư viện không chọn không bị cập nhật.

| Trạng thái | Ý nghĩa |
|---|---|
| Khớp hash | Cùng đường dẫn, dung lượng và QuickXorHash |
| Thiếu ở máy | Chưa có tệp tại đường dẫn SharePoint tương ứng |
| Khác dung lượng | Tệp cùng đường dẫn nhưng số byte khác nhau |
| Khác hash | Dung lượng bằng nhau nhưng QuickXorHash khác |
| Chỉ có ở local | Có đường dẫn local nhưng không có trong thư viện đã đọc thành công |
| Chưa xác minh hash | Cùng dung lượng nhưng SharePoint không cung cấp hash |
| Lỗi / xung đột | Không đọc được, metadata chưa đủ, tên không hợp lệ hoặc xung đột đường dẫn |

QuickXorHash là hash đối chiếu của Microsoft, **không phải hash mật mã**. SHA-256 dùng để kiểm tra sao lưu và tệp vừa ghi. Hash không cho biết đoạn văn nào thay đổi; ứng dụng không mở Word/Excel/PDF để phân tích nội dung.

Lọc tên/đường dẫn và trạng thái ở thanh công cụ. Chọn một hàng để xem metadata dễ đọc; mở **Hash & bản ghi gốc** để xem hash và JSON chi tiết. **Chọn đang hiển thị** chỉ thay đổi những hàng đang lọc và đủ điều kiện chọn.

### 5. Tải tệp thiếu và ghi đè

Các tệp thiếu được chọn sẵn. Bỏ chọn video hoặc tệp lớn nếu không muốn tải chúng. Cột **Tái dùng** báo có bản cùng dung lượng/hash trong local hoặc cache; bản nguồn được giữ nguyên.

| Chế độ | Kết quả khi bấm Thực hiện |
|---|---|
| Không ghi đè — mặc định | Chỉ bổ sung các tệp thiếu đã chọn |
| Chọn từng tệp | Bổ sung tệp thiếu đã chọn và thay các tệp khác biệt được tích chọn |
| Ghi đè tất cả tệp khác biệt | Thay tất cả tệp khác dung lượng/hash trong phạm vi quét, kể cả hàng đang bị bộ lọc che đi; vẫn chỉ bổ sung các tệp thiếu đã chọn |

**Chưa xác minh hash**, **chỉ local** và **lỗi** không tự đưa vào “ghi đè tất cả”. Nếu thiếu tệp và API không có hash, vẫn có thể tải tệp đó; báo cáo ghi rõ chưa đối chiếu được hash SharePoint.

Trước khi bấm **Thực hiện**, xem số tệp tải mới, số tệp ghi đè, số bản tái dùng, MB mạng dự kiến và dung lượng sao lưu. Ước tính mạng là mức dự kiến; cache được kiểm tra lại với bản SharePoint hiện hành trước khi dùng.

Quy trình mỗi tệp:

1. Kiểm tra lại bản SharePoint và metadata.
2. Dùng lại bản hợp lệ đã có hoặc tải vào vùng tạm; tối đa hai lượt tải đồng thời.
3. Kiểm tra dung lượng/hash và xác nhận nguồn không đổi trong lúc lấy dữ liệu.
4. Nếu ghi đè: xác nhận local chưa bị sửa sau lúc quét, sao lưu và kiểm tra SHA-256 của sao lưu.
5. Kiểm tra local lần cuối, thay tệp nguyên tử trên Windows, kiểm tra SHA-256 sau ghi và lưu metadata gắn với đúng bản đã cài đặt.

Ứng dụng không nhúng metadata vào Word/Excel, vì việc đó làm thay đổi bytes/hash. Ngày tạo và ngày sửa local được đặt theo SharePoint sau khi cài đặt tệp thành công.

### 6. Metadata, báo cáo và khôi phục

Với thư mục local `D:\Mirror`, dữ liệu đi kèm mặc định nằm ở thư mục anh em `D:\Mirror_SPDiff\<site-id-hash>`. Không tạo các tệp báo cáo lẫn vào thư mục đang đối chiếu.

```text
Mirror_SPDiff/<site-id-hash>/
  library-mappings.json         # ID thư viện -> tên thư mục local ổn định
  remote-snapshot.json          # Metadata SharePoint ở lần quét
  installed-manifest.json       # Metadata gắn với đúng hash bản local đã xác nhận
  cache/                        # Bản tải hợp lệ và tệp .partial để chạy tiếp
  runs/<ngày-giờ-id>/
    scan.json, diff.csv, metadata.csv
    plan.json
    backups/<cây đường dẫn cũ>   # Bản local trước khi ghi đè
    intents/, committed/         # Dấu vết trước/sau thao tác ghi
    results.json, results.csv
    remote-metadata.csv, summary.txt
    errors/                     # Lỗi cụ thể theo tệp
```

Các tệp trong `runs/` xuất hiện tuỳ loại tác vụ: quét tạo bộ `scan`, cập nhật tạo bộ `plan/results/backups`. Không tự dọn cache hoặc xoá sao lưu.

JSON giữ kiểu dữ liệu và các giá trị gốc; CSV metadata có tên nội bộ, tên hiển thị, loại, giá trị gốc và giá trị hiển thị của cột. Bao gồm lookup/người dùng, taxonomy và đa giá trị mà tài khoản đọc được. Cột lỗi hoặc thiếu không bị âm thầm bỏ qua.

Nếu giữ một bản local khác biệt, metadata SharePoint mới chỉ là **snapshot remote**, không bị gán nhầm là metadata của bản local đó.

Để khôi phục một tệp, đóng tài liệu đang mở, tìm hàng tương ứng trong `results.csv`, lấy bản ở `BackupPath` và sao chép về đường dẫn cũ sau khi tự kiểm tra. Tệp `.metadata.json` kèm sao lưu ghi hash và thông tin cần đối chiếu.

### 7. Dừng, thử lại và đọc lỗi

- **Dừng** huỷ công việc chưa hoàn tất. Các tệp đã cài đặt hợp lệ vẫn được giữ và có báo cáo; không hoàn tác âm thầm các kết quả đã hoàn tất.
- Chạy so sánh/thực hiện lại để thử lại tệp lỗi. Bản tải hoàn chỉnh đã xác minh được dùng lại. Tệp dở chỉ tiếp tục khi phiên bản nguồn không đổi và máy chủ hỗ trợ Range; nếu không, lượt tải đó bắt đầu lại.
- Khi bị giới hạn truy cập, ứng dụng chờ theo `Retry-After`. Không tạo thêm lượt tải cho tệp chưa chọn chỉ để kiểm tra hash.
- Mở **Chi tiết kỹ thuật & nguyên nhân** và **Sao chép lỗi** để lấy thao tác, tệp/tài nguyên, nguyên nhân, mã HTTP/request ID khi có, đường dẫn log và stack trace.
- Lỗi vận hành dự kiến như tệp bị khoá, hết dung lượng hoặc mất quyền được ghi theo từng tệp. Bản gốc chưa được phép thay vẫn được giữ; nếu đã có sao lưu thì đường dẫn sao lưu được ghi trong báo cáo.
- **Lỗi lập trình ngoài dự kiến không bị nuốt để chạy tiếp**: ứng dụng dừng tác vụ, ghi log, hiện lỗi chi tiết, rồi thoát với mã `1` sau khi đóng hộp lỗi. Log UI nằm tại `%LOCALAPPDATA%\SharePointDiff\logs`.

| Lỗi | Cách xử lý |
|---|---|
| AuthenticationFailed / AADSTS | Kiểm tra tenant, Client ID, redirect URI, tài khoản, MFA và quyền được cấp; không tắt chính sách bảo mật |
| AccessDenied | Kiểm tra cả quyền tài khoản và delegated permissions của ứng dụng |
| MetadataIncomplete | Xem cột nào chưa đọc được; chưa cập nhật tệp đó |
| FileChanged | Local đã đổi hoặc tệp mới xuất hiện sau lúc quét; so sánh lại |
| RemoteChanged | Bản SharePoint đổi trong khi xử lý; so sánh lại |
| HashMismatch | Không dùng bản tải sai; bản sai được giữ tách biệt khỏi cache hợp lệ |
| FileIoError | Đóng tài liệu đang mở; kiểm tra quyền ghi và dung lượng |
| PathCollision / InvalidPath | Kiểm tra tên Unicode, tệp/thư mục trùng tên và tên không hợp lệ trên Windows; công cụ không tự đổi tên |

Phạm vi v1: SharePoint Online thương mại và bản tệp hiện hành. Không đồng bộ ngược, không tải lịch sử phiên bản/quyền/chú thích, không diễn giải nội dung tài liệu. Trước khi dùng “ghi đè tất cả” trên một thư mục mới, thử quy trình trên thư mục thử riêng.

## English guide

### Requirements and launch

Windows x64 with **.NET 10 Windows Desktop Runtime**. Open `publish\win-x64\SharePointDiff.exe`; no Python, PowerShell, PnP module or Visual Studio is required at runtime. The language selector switches between Vietnamese and English immediately and remembers the choice.

### Connect

1. Enter an HTTPS SharePoint **site URL**, not a file-sharing or library-view URL.
2. Choose an existing local folder; drive roots, symlinks and junctions are not accepted as a mirror root.
3. Select or save a profile with its tenant ID/domain and Entra application/client ID.
4. The registered public desktop application needs `http://localhost` redirect, delegated Graph `Sites.Read.All` and SharePoint `AllSites.Read`, with consent according to your organization. No client secret is used.
5. Click **Connect**, complete Microsoft browser sign-in/MFA, then select document libraries.

Settings are stored in `%LOCALAPPDATA%\SharePointDiff\settings.json`; passwords and tokens are not saved. A compatible existing `sp_chuanhoa` configuration can be imported without modifying it. Each tenant uses its own connection profile. No security policy or application permission is changed by this program.

See [`samples/settings.example.json`](samples/settings.example.json) for two tenant profiles. Replace all example IDs, the site URL and local path with your own values. Prefer editing profiles in the UI; close the application and back up existing settings before editing JSON manually. `SelectedProfile` is zero-based; `Language` accepts `vi-VN` or `en-US`.

### Compare and select actions

**Compare** reads inventory/metadata and hashes local files. It never downloads remote file content. The statuses are Hash match, Missing locally, Different size/hash, Local only, Hash unverified and Error/conflict.

QuickXorHash plus size is used for SharePoint comparison; SHA-256 verifies backups and installed bytes. QuickXorHash is non-cryptographic. Equal size alone does not prove equal content. The tool does not open documents or explain textual changes.

Missing files are selected by default. Uncheck large files before applying. Matching local or verified cached bytes are reused without moving the source. The filter affects **Select visible**, but **Replace all differences** covers the entire scanned library scope, including rows hidden by filters.

- **Keep local:** add selected missing files only.
- **Select files:** add selected missing files and replace selected known differences.
- **Replace all differences:** replace every known size/hash difference and add selected missing files.

Unverified equal-size files, local-only files and errors are excluded from automatic replacement. **Local-only files are never moved or deleted.** Review estimated network and backup sizes, then click **Apply**. Remote writes are never performed.

### Backups, metadata and reports

For `D:\Mirror`, application data is placed in the sibling `D:\Mirror_SPDiff\<site-hash>`. Each run has JSON/CSV reports. Every replacement requires a verified backup under `runs/<run>/backups/<original-path>`. Backups are never automatically removed.

Downloads are staged and checked against current remote metadata/hash. Local changes since the scan prevent replacement. The final Windows replacement is atomic and the installed bytes are verified. File creation/modification dates follow SharePoint.

Full readable column values, schema, lookups, taxonomy and multivalues are saved separately as typed JSON and readable CSV. File bytes are not modified to embed metadata. `remote-snapshot.json` describes remote observations; `installed-manifest.json` associates metadata with verified local content, so skipped local differences are not labelled with the wrong remote version.

Use **Reports** to find `results.csv`, backup locations, SHA-256 values, downloaded byte counts and per-file diagnostics. To restore, close the affected document and manually copy the verified backup to its original path after reviewing the report.

### Cancellation and failures

**Stop** cancels outstanding work but preserves completed installations and their audit records. Verified downloads are reused on retry. Partial content resumes only when the source version is unchanged and the server supports Range. Up to two transfers run concurrently; throttling honors `Retry-After`.

Expected operational errors are reported per file. Expand **Technical details & cause** and use **Copy error** for the exact operation, path, root cause, HTTP status/request ID, log and stack trace. Unexpected programming exceptions stop processing and terminate the application with exit code `1` after the diagnostic dialog is closed; they are not swallowed to continue running.

Version 1 supports commercial SharePoint Online and current file versions only. It does not upload, synchronize permissions or download version history/comments. Validate replacement on a separate test folder before using a new production mirror.

## Build, kiểm thử và Git

Mở `SharePointDiff.sln` trong Visual Studio 2026 hoặc dùng .NET SDK 10:

```powershell
dotnet restore SharePointDiff.sln
dotnet build SharePointDiff.sln -c Release --no-restore
dotnet test tests\SharePointDiff.Tests -c Release --no-build
dotnet publish src\SharePointDiff.App -c Release -r win-x64 --self-contained false -o publish\win-x64
```

Các gói NuGet được ghim phiên bản. Bản publish gom ứng dụng thành một `.exe` và dùng .NET 10 Desktop Runtime đã cài. Không bật trimming cho WPF.

Có thể chạy `scripts\Publish.ps1` bằng PowerShell 7 để restore theo lockfile, build, kiểm thử và publish trong một lượt. Xem [báo cáo kiểm chứng / verification report](docs/Verification.md) để biết kết quả và phạm vi đã thử.

Kiểm thử lõi dùng tệp giả lập và HTTP giả lập: hash độc lập, Unicode, chọn từng tệp/tất cả, không tải khi quét, giữ local-only, backup, nguồn/local thay đổi, tải sai, metadata, huỷ/thử lại, HTTP Range và throttling. Không ghi lên SharePoint hoặc hồ sơ dự án thật.

Chạy giao diện với 10.000 tệp giả, không kết nối hay cập nhật dữ liệu:

```powershell
.\publish\win-x64\SharePointDiff.exe --demo
```

Chế độ kiểm tra giao diện dành cho phát triển xuất ảnh và JSON vào thư mục thử đã chỉ định:

```powershell
.\publish\win-x64\SharePointDiff.exe --demo --ui-smoke D:\Temp\SharePointDiff-UI
```

`--demo --fatal-smoke` chủ động tạo lỗi để kiểm tra log và mã thoát `1`, không hiện hộp lỗi tương tác. Không dùng các cờ kiểm thử này cho quy trình sử dụng thông thường. Biến môi trường `SPDIFF_SETTINGS_PATH` có thể chuyển cấu hình/log sang thư mục thử riêng.

### Quy ước xuống dòng

- `.gitattributes` chỉ định **CRLF trong working tree** cho C#, XAML, RESX, project/solution, Markdown, JSON và PowerShell.
- Git chuẩn hoá nội dung text trong object database về LF; đó là hành vi bình thường. `git ls-files --eol` thường hiện `i/lf w/crlf`.
- `.editorconfig` quy định UTF-8, CRLF, khoảng trắng và newline cuối tệp. Ảnh/icon được đánh dấu binary.
- Không phụ thuộc cấu hình `core.autocrlf` toàn máy. Repo dùng cấu hình local và attributes rõ ràng; không đổi Git config toàn cục.
- `bin`, `obj`, `.vs`, `publish`, test results và dữ liệu người dùng không được đưa vào commit. Repo chỉ chứa mã ứng dụng, kiểm thử, tài liệu và tài nguyên giao diện.

## Kiến trúc và nguồn tham khảo

| Thành phần | Trách nhiệm |
|---|---|
| `SharePointDiff.App` | WPF/MVVM, Fluent light theme, tài nguyên VI/EN, hồ sơ và lỗi chi tiết |
| `SharePointDiff.Core` | MSAL, Graph/REST chỉ đọc, kiểm kê, hash, kế hoạch cập nhật, cache/backup/manifest |
| `SharePointDiff.Tests` | Kiểm thử dữ liệu thử và HTTP giả lập |

Nguồn kỹ thuật:

- [WPF trên .NET 10](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/whats-new/net100)
- [MSAL system browser và localhost redirect](https://learn.microsoft.com/en-us/entra/msal/dotnet/acquiring-tokens/using-web-browsers)
- [Microsoft QuickXorHash algorithm](https://learn.microsoft.com/en-us/onedrive/developer/code-snippets/quickxorhash?view=odsp-graph-online)
- [Graph hashes và giới hạn hash được cung cấp](https://learn.microsoft.com/en-us/graph/api/resources/hashes?view=graph-rest-1.0)
- [Graph delta và phân trang](https://learn.microsoft.com/en-us/graph/api/driveitem-delta?view=graph-rest-1.0)
- [Graph content download và Range](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0)
- [Graph throttling](https://learn.microsoft.com/en-us/graph/throttling)
