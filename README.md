# SharePoint Diff

**ENGLISH BELOW**

**Tác giả / Authors:** Vũ Mạnh Quân (taladinhcao@gmail.com) & Trong-Nghia Be (benghia1st@gmail.com)

**Giấy phép / License:** [MIT](LICENSE)

Thư viện đi kèm / Bundled dependencies: [Third-party notices](THIRD-PARTY-NOTICES.md).

Ứng dụng Windows **C# / WPF / .NET 10**, giao diện **Tiếng Việt và English**, để đối chiếu thư mục local với SharePoint Online bằng hash, tải tệp thiếu và thay các bản khác biệt theo lựa chọn của người dùng.

**Chỉ đọc SharePoint. Không upload, sửa metadata trên site, xoá hoặc di chuyển tệp chỉ có ở local. Mọi lần ghi đè local đều có bản sao lưu đã kiểm tra SHA-256.**

- [Hướng dẫn tiếng Việt](#hướng-dẫn-tiếng-việt)
- [Kiểm tra khi đăng nhập Windows](#8-kiểm-tra-khi-đăng-nhập-windows)
- [English guide](#english-guide)
- [Build, kiểm thử và Git](#build-kiểm-thử-và-git)
- [Kiến trúc và nguồn tham khảo](#kiến-trúc-và-nguồn-tham-khảo)

## Hướng dẫn tiếng Việt

### 1. Chạy ứng dụng

Yêu cầu Windows x64 và **.NET 10 Windows Desktop Runtime**. Visual Studio không cần thiết khi chạy bản phát hành. Máy đã có .NET 10 Desktop Runtime thì không phải cài thêm Python, PowerShell, PnP.PowerShell hoặc trình biên dịch.

Tải ZIP Windows x64 từ [GitHub Releases](https://github.com/qwerty78456/SharePointDiff/releases), giải nén toàn bộ rồi mở `SharePointDiff.exe` trong thư mục vừa giải nén. Nếu build từ mã nguồn, tệp nằm ở `publish\win-x64\SharePointDiff.exe`. Có thể tạo shortcut tới tệp này theo cách thông thường của Windows.

Tệp `SHA256SUMS.txt` cạnh bản tải ghi SHA-256 của ZIP. Có thể đối chiếu bằng `Get-FileHash -Algorithm SHA256 <đường-dẫn-ZIP>`; thao tác kiểm tra này là tuỳ chọn và không cần thiết để chạy ứng dụng.

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

Để ứng dụng nhớ đăng nhập và tự kiểm tra khi bạn đăng nhập Windows, hãy dùng một app Entra riêng với hai quyền đọc trong bảng trên. Ứng dụng kiểm tra quyền Microsoft thực sự cấp cho lần đăng nhập này, cho cả Microsoft Graph và SharePoint.

Nếu app đang có `AllSites.FullControl` (toàn quyền, có thể thêm/sửa/xoá dữ liệu), bạn vẫn có thể tự kết nối để xem và tải tệp. SharePoint Diff không sửa dữ liệu trên SharePoint, **không lưu đăng nhập** của app đó và không bật tự kiểm tra khi đăng nhập Windows. Thông báo sẽ liệt kê ngay quyền đang có, quyền đọc cần dùng và quyền phải bỏ. Các quyền chưa xác định được mức truy cập, như `Sites.Selected`, cũng chưa được chấp nhận để lưu đăng nhập. Nhật ký chỉ ghi tên quyền; không ghi nội dung dùng để đăng nhập.

Ứng dụng dùng MSAL.NET và trình duyệt hệ thống, hỗ trợ MFA. Không dùng device-code login, không yêu cầu tắt Security Defaults và không tự đăng ký ứng dụng hay cấp quyền. Một Client ID chỉ dùng được tại tenant nơi nó được đăng ký hoặc chấp thuận phù hợp.

Nếu chạy trong cùng cây công cụ với `sp_chuanhoa\cauhinh\cauhinh.json`, lần đầu ứng dụng có thể đọc Tenant, Client ID, site và thư mục mặc định từ cấu hình có sẵn. Tệp cấu hình gốc không bị sửa. Nếu không có cấu hình đó, tạo hồ sơ mới trong giao diện.

### 3. Kết nối một site

1. Mở **Kết nối & thư mục**.
2. Dán URL của **site**, ví dụ `https://contoso.sharepoint.com/sites/Engineering`. Không dùng link chia sẻ của một tệp, trang `Forms/AllItems.aspx` hoặc link có query string.
3. Chọn thư mục local tương ứng. Thư mục phải tồn tại; chọn thư mục thông thường thay vì gốc ổ đĩa hay symbolic link/junction.
4. Chọn hồ sơ có sẵn; hoặc bấm **Thêm**, mở **Chỉnh hồ sơ đăng nhập**, nhập tên hồ sơ, Tenant ID/tên miền và Entra Client ID, rồi **Lưu hồ sơ**.
5. Bấm **Kết nối** và hoàn tất đăng nhập Microsoft. Tài khoản phải có quyền vào site; có thể cần hai lượt cấp token cho Graph và SharePoint, dùng cùng phiên đăng nhập.
6. Chọn các thư viện cần đối chiếu. Thư viện tài liệu không ẩn được chọn mặc định. Sau kết nối, phần cấu hình thu gọn để nhường chỗ cho bảng.

Hồ sơ và ngôn ngữ được lưu tại `%LOCALAPPDATA%\SharePointDiff\settings.json`. Tệp này chỉ chứa thông tin cấu hình, không chứa mật khẩu hay token.

Phiên đăng nhập Microsoft (token cache của MSAL) được lưu **mã hoá DPAPI theo tài khoản Windows hiện tại** trong `%LOCALAPPDATA%\SharePointDiff\msal\`, để lần sau và khi kiểm tra nền không phải đăng nhập lại. Tài khoản Windows khác không giải mã được; chương trình nào chạy dưới chính tài khoản Windows của bạn thì giải mã được. Mật khẩu không bao giờ được lưu. Bấm **Quên đăng nhập** để xoá phiên đã lưu của hồ sơ đang chọn. Token có quyền ghi không bao giờ được giữ trên đĩa.

Cấu hình mẫu có hai tenant ở [`samples/settings.example.json`](samples/settings.example.json). Các Client ID trong mẫu là giá trị giả: thay bằng ID được cấp cho từng tenant, URL site và thư mục có thật. Có thể tạo hồ sơ bằng giao diện; nếu chỉnh JSON thủ công, đóng ứng dụng và sao lưu cấu hình hiện tại trước khi chỉnh. `SelectedProfile` là chỉ số bắt đầu từ `0`; `Language` nhận `vi-VN` hoặc `en-US`.

### 4. So sánh

Bấm **So sánh**. Bước này đọc danh sách và metadata SharePoint, tính lại QuickXorHash và SHA-256 của local; **không tải nội dung tệp SharePoint**.

Ứng dụng chỉ đọc các thư mục local ứng với thư viện đã chọn cùng các tệp rời ở gốc; thư mục khác ở gốc (ví dụ thư mục sao lưu riêng của bạn) không bị đọc hay băm. Bản cùng hash trong phạm vi đó được tái dùng. Tệp trong thư viện không chọn không bị cập nhật. Tệp tạm `*.spdiff-<guid>.tmp` còn sót sau một lần bị ngắt được báo lỗi `StaleInstallTemp`, không bị so sánh hay tự xoá.

Nếu đăng nhập hết hạn giữa chừng, lượt So sánh dừng ngay thay vì băm toàn bộ thư mục rồi báo lỗi từng thư viện. Bộ lọc **Đã sửa ở local** hiện các tệp khác biệt mà bản local khác bản đã đồng bộ lần trước.

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
  library-selection.json        # Thư viện đã thấy/đã chọn ở lần So sánh gần nhất (dùng cho kiểm tra nền)
  backup-sequence.json          # Bộ đếm thứ tự sao lưu
  remote-snapshot.json          # Metadata SharePoint ở lần quét
  installed-manifest.json       # Metadata gắn với đúng hash bản local đã xác nhận
  cache/                        # Bản tải hợp lệ và tệp .partial để chạy tiếp
  runs/<ngày-giờ-id>/
    scan.json, diff.csv, metadata.csv
    plan.json
    backups/<cây đường dẫn cũ>   # Bản local trước khi ghi đè
    backup-meta/<id>.json        # Hash, nguồn gốc và số thứ tự của từng bản sao lưu
    prune-plan.json, pruned.json # Những gì quy tắc xoay vòng sắp xoá / đã xoá
    intents/, committed/         # Dấu vết trước/sau thao tác ghi
    results.json, results.csv
    remote-metadata.csv, summary.txt
    errors/                     # Lỗi cụ thể theo tệp
```

Các tệp trong `runs/` xuất hiện tuỳ loại tác vụ: quét tạo bộ `scan`, cập nhật tạo bộ `plan/results/backups`. Tệp `Mirror_SPDiff/mirror.lock` ngăn hai lượt quét/đồng bộ chạy cùng lúc trên một thư mục (lỗi `WorkspaceBusy`).

Quy tắc dọn (chỉ chạy cuối một lượt Thực hiện/Đồng bộ hoàn tất, không bị huỷ hay dừng):

- **Sao lưu là phiên bản SharePoint cũ** (bản local trước khi thay đúng bằng bản đã đồng bộ và xác minh trước đó): mỗi tệp giữ **3 phiên bản khác nhau mới nhất**, bản cũ hơn và bản trùng bị xoá vĩnh viễn. Phiên bản đó vẫn có thể lấy lại từ version history của SharePoint khi còn tồn tại. Việc so khớp dựa trên QuickXorHash (không phải hash mật mã).
- **Sao lưu chứa chỉnh sửa local, không rõ nguồn gốc, hoặc không có hash SharePoint được giữ vĩnh viễn.** Sao lưu của lượt chưa hoàn tất hoặc lượt thay bị ngắt cũng được giữ.
- Trước khi xoá, `prune-plan.json` ghi đầy đủ dấu vết (đường dẫn, hash, ItemId, ETag, WebUrl, ngày sửa); mỗi bản được băm lại và chỉ bị xoá khi khớp đúng bản ghi. `pruned.json` ghi kết quả.
- **Cache:** giữ bản của các phiên bản đang cài (khôi phục được khi không có mạng nếu lỡ xoá tệp local); xoá bản của phiên bản đã bị thay, chỉ với thư viện được đọc đầy đủ ở lần quét đó. Tệp `.invalid_*` và tệp lạ không bị đụng tới.

JSON giữ kiểu dữ liệu và các giá trị gốc; CSV metadata có tên nội bộ, tên hiển thị, loại, giá trị gốc và giá trị hiển thị của cột. Bao gồm lookup/người dùng, taxonomy và đa giá trị mà tài khoản đọc được. Cột lỗi hoặc thiếu không bị âm thầm bỏ qua.

Nếu giữ một bản local khác biệt, metadata SharePoint mới chỉ là **snapshot remote**, không bị gán nhầm là metadata của bản local đó.

Để khôi phục một tệp, đóng tài liệu đang mở, tìm hàng tương ứng trong `results.csv`, lấy bản ở `BackupPath` và sao chép về đường dẫn cũ sau khi tự kiểm tra. Bản ghi `runs/<run>/backup-meta/<id>.json` chứa hash và thông tin cần đối chiếu (bản v1.0 để `.metadata.json` ngay cạnh sao lưu). Nếu không thấy bản cần tìm, xem `pruned.json` rồi lấy lại từ version history của SharePoint. Đường dẫn sao lưu có thể dài hơn 260 ký tự: dùng PowerShell 7 `Copy-Item -LiteralPath` hoặc `robocopy` thay vì Explorer.

### 7. Dừng, thử lại và đọc lỗi

- **Dừng** huỷ công việc chưa hoàn tất. Các tệp đã cài đặt hợp lệ vẫn được giữ và có báo cáo; không hoàn tác âm thầm các kết quả đã hoàn tất.
- Chạy so sánh/thực hiện lại để thử lại tệp lỗi. Bản tải hoàn chỉnh đã xác minh được dùng lại. Tệp dở chỉ tiếp tục khi nội dung nguồn không đổi (cùng dung lượng và hash; sửa cột metadata không làm mất phần đã tải) và máy chủ hỗ trợ Range; nếu không, lượt tải đó bắt đầu lại.
- Khi mạng chập chờn: lượt tải không nhận được byte nào trong 60 giây được coi là treo; kết nối bị ngắt, URL tải hết hạn hay máy chủ bận cũng được thử lại. Mỗi lần thử lấy URL tải mới và tải tiếp bằng Range từ phần đã có, tối đa 8 lần (chờ 5 s, nhân đôi, tối đa 60 s). `results.csv` có cột `Attempts` và `Stalls`. Lỗi ghi đĩa local không bao giờ bị thử lại. Hết lượt thử thì phần đã tải vẫn được giữ để lần sau tải tiếp.
- Hết thời gian chờ được báo là lỗi (`RequestTimeout`), không phải huỷ.
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
| SignInRequired | Phiên đăng nhập đã lưu không còn hoặc hết hạn; bấm Đăng nhập/Kết nối |
| WriteScopedApp | Client ID không chỉ có quyền đọc; dùng app chỉ đọc riêng (mục 2) |
| WorkspaceBusy | Một lượt quét/đồng bộ khác đang dùng thư mục; chờ rồi thử lại |
| RequestTimeout / ServiceUnavailable | Mạng chậm hoặc SharePoint bận; thử lại sau |
| DownloadStalled / DownloadInterrupted | Mạng treo hoặc đứt khi tải; phần đã tải được giữ, chạy lại để tải tiếp |
| StaleInstallTemp | Tệp tạm còn sót sau lần bị ngắt; kiểm tra rồi tự xoá |
| AgentNotReady / NoLibrariesSelected | Bấm So sánh một lần cho site và thư mục này, với ít nhất một thư viện được chọn |

Phạm vi v1: SharePoint Online thương mại và bản tệp hiện hành. Không đồng bộ ngược, không tải lịch sử phiên bản/quyền/chú thích, không diễn giải nội dung tài liệu. Trước khi dùng “ghi đè tất cả” trên một thư mục mới, thử quy trình trên thư mục thử riêng.

### 8. Kiểm tra khi đăng nhập Windows

Tuỳ chọn **Kiểm tra thay đổi khi đăng nhập Windows** trong phần Kết nối giúp có sẵn bản local khi M365 online chậm. Mỗi lần đăng nhập Windows, biểu tượng SharePoint Diff ở khay hệ thống sẽ kiểm tra nền rồi **hỏi trước khi ghi bất cứ thứ gì** vào thư mục.

- **Điều kiện bật:** đã bấm So sánh ít nhất một lần cho đúng site và thư mục đang chọn; dùng app Entra chỉ đọc; không đặt `SPDIFF_SETTINGS_PATH`; không chạy ở chế độ demo. Tuỳ chọn có hiệu lực từ lần đăng nhập Windows sau.
- **Cách chạy:**
  1. Ứng dụng tự đăng ký mục `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` (`"<exe>" --agent`); không cần quyền admin, không dùng Windows Service.
  2. Sau khoảng 60 giây, agent chờ mạng (tối đa 15 phút) và lấy token đã lưu.
  3. Agent quét **chỉ các thư viện đã chọn ở lần So sánh gần nhất**, ưu tiên thấp. Thư viện mới xuất hiện sau đó chưa được kiểm tra cho tới khi bạn So sánh lại.
  4. Lỗi tạm thời được thử lại trong 60 phút. Agent không kiểm tra lại khi mở khoá hay thức dậy từ Sleep; dùng **Kiểm tra ngay** trong menu khay.
- **Thẻ thông báo** (góc dưới phải) hiện:
  - số tệp mới và tệp cập nhật kèm dung lượng;
  - **các tệp khác bản đồng bộ lần trước**: sửa ở cả hai phía, chỉ sửa ở máy này, hoặc chưa từng đồng bộ. Có tên tối đa 5 tệp;
  - mục cần xem, thư viện mới hoặc lỗi, ước tính tải và sao lưu.

  Nút **Đồng bộ** tải mọi tệp thiếu và **thay mọi tệp khác biệt, kể cả tệp bạn đã sửa ở local**, luôn sao lưu và kiểm tra trước. **Chi tiết** mở cửa sổ chính ở chế độ Không ghi đè để bạn tự chọn. **Để sau** không làm gì. Nếu không có thay đổi, chỉ tooltip cập nhật.
- **Đăng nhập:** khi phiên đã lưu hết hạn, biểu tượng chuyển màu cam và thẻ "Cần đăng nhập" hiện ra. Trình duyệt **chỉ mở khi bạn bấm** Đăng nhập.
- **Nên làm:**
  - Ghim biểu tượng (Settings > Personalization > Taskbar > Other system tray icons).
  - Đặt `SharePointDiff.exe` ở thư mục cố định như `%LOCALAPPDATA%\Programs\SharePointDiff`. Sau khi chuyển thư mục, mở bản mới và bấm **Dùng bản này**.
  - Thoát agent trước khi publish lại, và gỡ các bản v1.0 (chúng không lấy khoá thư mục).
  - Nếu tải ZIP từ Internet, chọn Unblock trước khi giải nén.
- **Tắt:** bỏ chọn tuỳ chọn (luôn bỏ chọn được, kể cả khi điều kiện bật không còn đúng; mục Run bị gỡ), hoặc chọn Thoát trong menu khay. Nếu tắt trong Task Manager, cửa sổ chính sẽ báo đang bị tắt.
- **Kiểm tra ngay** dừng lượt kiểm tra khi đăng nhập đang chờ mạng hoặc chờ thử lại, rồi kiểm tra ngay. So sánh hoặc Thực hiện trong cửa sổ chính sẽ huỷ thẻ thông báo cũ của agent. Một lượt Đồng bộ bị dừng hoặc còn lỗi được báo "Đồng bộ chưa xong", không báo "Đã khớp".

## English guide

### Requirements and launch

Windows x64 with **.NET 10 Windows Desktop Runtime**. Download the Windows x64 ZIP from [GitHub Releases](https://github.com/qwerty78456/SharePointDiff/releases), extract all files and open `SharePointDiff.exe`. Source builds place it in `publish\win-x64`. No Python, PowerShell, PnP module or Visual Studio is required at runtime. The language selector switches between Vietnamese and English immediately and remembers the choice.

`SHA256SUMS.txt` contains the ZIP checksum. Optionally verify it with `Get-FileHash -Algorithm SHA256 <ZIP-path>`; this is not required to run the app.

### Connect

1. Enter an HTTPS SharePoint **site URL**, not a file-sharing or library-view URL.
2. Choose an existing local folder; drive roots, symlinks and junctions are not accepted as a mirror root.
3. Select or save a profile with its tenant ID/domain and Entra application/client ID.
4. The registered public desktop application needs `http://localhost` redirect, delegated Graph `Sites.Read.All` and SharePoint `AllSites.Read`, with consent according to your organization. No client secret is used.
5. Click **Connect**, complete Microsoft browser sign-in/MFA, then select document libraries.

Settings are stored in `%LOCALAPPDATA%\SharePointDiff\settings.json` (no passwords or tokens). The Microsoft sign-in (MSAL token cache) is saved **encrypted with DPAPI for the current Windows user** in `%LOCALAPPDATA%\SharePointDiff\msal\`, so later runs and the background check need no new sign-in; other Windows accounts cannot decrypt it, programs running as your own account can. **Forget sign-in** removes it. Use a dedicated **read-only** Entra app. A sign-in is saved only when **every** scope in the token is a known read-only scope (such as `Sites.Read.All`, `AllSites.Read`, `User.Read`). Anything else means the sign-in is never kept on disk, checking at Windows sign-in is unavailable and the app shows `WriteScopedApp`. Examples are write scopes (a PnP app with `AllSites.FullControl`) and per-site grants such as `Sites.Selected` whose level the token does not show. Connect checks both the Graph and the SharePoint token first and keeps nothing on disk unless both pass. A compatible existing `sp_chuanhoa` configuration can be imported without modifying it. Each tenant uses its own connection profile. No security policy or application permission is changed by this program.

See [`samples/settings.example.json`](samples/settings.example.json) for two tenant profiles. Replace all example IDs, the site URL and local path with your own values. Prefer editing profiles in the UI; close the application and back up existing settings before editing JSON manually. `SelectedProfile` is zero-based; `Language` accepts `vi-VN` or `en-US`.

Permission warnings list the actual granted scopes for Graph and SharePoint, the required read permissions and every rejected scope. Opaque tokens are reported as unverified, not incorrectly labelled as write permissions. The audit records permission names only, never access or refresh tokens. Changing the UI language also updates an error that is already displayed.

### Compare and select actions

**Compare** reads inventory/metadata and hashes local files in the selected library folders plus loose root files (other root folders are not read). It never downloads remote file content. A sign-in failure stops the scan at once. Leftover `*.spdiff-<guid>.tmp` files from an interrupted run are reported as `StaleInstallTemp`, never compared or deleted. The **Locally changed** filter lists differences whose local copy is not the last synced version. The statuses are Hash match, Missing locally, Different size/hash, Local only, Hash unverified and Error/conflict.

QuickXorHash plus size is used for SharePoint comparison; SHA-256 verifies backups and installed bytes. QuickXorHash is non-cryptographic. Equal size alone does not prove equal content. The tool does not open documents or explain textual changes.

Missing files are selected by default. Uncheck large files before applying. Matching local or verified cached bytes are reused without moving the source. The filter affects **Select visible**, but **Replace all differences** covers the entire scanned library scope, including rows hidden by filters.

- **Keep local:** add selected missing files only.
- **Select files:** add selected missing files and replace selected known differences.
- **Replace all differences:** replace every known size/hash difference and add selected missing files.

Unverified equal-size files, local-only files and errors are excluded from automatic replacement. **Local-only files are never moved or deleted.** Review estimated network and backup sizes, then click **Apply**. Remote writes are never performed.

### Backups, metadata and reports

For `D:\Mirror`, application data is placed in the sibling `D:\Mirror_SPDiff\<site-hash>`; `D:\Mirror_SPDiff\mirror.lock` keeps two scans/syncs off the same folder (`WorkspaceBusy`). Each run has JSON/CSV reports. Every replacement requires a verified backup under `runs/<run>/backups/<original-path>`, described by `runs/<run>/backup-meta/<id>.json`.

Retention runs only at the end of a complete, uncancelled Apply or Sync:

- **Backups that are earlier SharePoint versions** (the replaced bytes equal a previously synced and verified version): the **3 newest distinct versions per file** are kept; older and duplicate copies are deleted permanently. They may still be recovered from SharePoint version history while it exists. Matching relies on QuickXorHash, which is not cryptographic.
- **Backups of local edits, of unknown origin or without a SharePoint hash are kept forever**, as are backups from unfinished runs or interrupted replacements.
- `prune-plan.json` records full tombstones (path, hashes, ItemId, ETag, WebUrl, modified date) before anything is deleted; each backup is re-hashed and deleted only if it matches its record; `pruned.json` records the outcome.
- **Cache:** entries of installed versions are kept (offline restore); entries of superseded versions are deleted, only for libraries read completely in that scan. `.invalid_*` and foreign files are left alone.

Downloads are staged and checked against current remote metadata/hash. Local changes since the scan prevent replacement. The final Windows replacement is atomic and the installed bytes are verified. File creation/modification dates follow SharePoint.

Full readable column values, schema, lookups, taxonomy and multivalues are saved separately as typed JSON and readable CSV. File bytes are not modified to embed metadata. `remote-snapshot.json` describes remote observations; `installed-manifest.json` associates metadata with verified local content, so skipped local differences are not labelled with the wrong remote version.

Use **Reports** to find `results.csv`, backup locations, SHA-256 values, downloaded byte counts and per-file diagnostics. To restore, close the affected document and manually copy the verified backup to its original path after reviewing the report; check `pruned.json` if a version is missing. Backup paths can exceed 260 characters: use PowerShell 7 `Copy-Item -LiteralPath` or `robocopy` rather than Explorer.

### Cancellation and failures

**Stop** cancels outstanding work but preserves completed installations and their audit records. Verified downloads are reused on retry. Partial content resumes when the source bytes are unchanged (same size and hash; a metadata-only edit keeps the partial) and the server supports Range. Up to two transfers run concurrently; throttling honors `Retry-After`.

On an unreliable line a download that receives no byte for 60 s is treated as stalled; dropped connections, expired download URLs and busy servers are retried too. Every attempt asks for a fresh download URL and resumes with Range, up to 8 attempts (5 s, doubling, at most 60 s apart). `results.csv` reports `Attempts` and `Stalls`. Local write failures are never retried. A partial file left after the last attempt resumes next time. Timeouts are reported as errors (`RequestTimeout`), not as cancellations.

Expected operational errors are reported per file. Expand **Technical details & cause** and use **Copy error** for the exact operation, path, root cause, HTTP status/request ID, log and stack trace. Unexpected programming exceptions stop processing and terminate the application with exit code `1` after the diagnostic dialog is closed; they are not swallowed to continue running.

New error codes: `SignInRequired`, `WriteScopedApp`, `WorkspaceBusy`, `RequestTimeout`, `ServiceUnavailable`, `DownloadStalled`, `DownloadInterrupted`, `StaleInstallTemp`, `AgentNotReady`, `NoLibrariesSelected` (see the Vietnamese table for the action to take).

### Check at Windows sign-in

The **Check for changes when I sign in to Windows** option keeps a local copy ready for when M365 online is slow. It needs a completed Compare for the current site and folder, a read-only app, no `SPDIFF_SETTINGS_PATH` and non-demo mode. It takes effect at the next Windows sign-in.

1. The app registers its own `HKCU\...\CurrentVersion\Run` entry (`"<exe>" --agent`): no admin rights, no Windows service.
2. About 60 s after sign-in the tray agent waits for the network (up to 15 min) and uses the saved sign-in.
3. It scans, at low priority, only the libraries ticked at the last Compare; libraries added later are left out until the next Compare. Transient failures are retried for up to 60 min.
4. It does not re-check on unlock or resume from Sleep; use **Check now** in the tray menu.

The card (bottom-right) shows:
- new and updated files with sizes;
- **files that differ from the last synced copy** (changed on both sides, changed only on this computer, never synced), naming up to 5;
- items to review, new or unreadable libraries, and the download and backup estimate.

**Sync** downloads every missing file and **replaces every differing file, including your local edits**, always after a verified backup. **Details** opens the main window in Keep-local mode. **Later** does nothing.

Nothing is written to the folder before you click Sync. When the saved sign-in has expired the icon turns amber and a Sign-in card appears; the browser opens **only when you click**.

Keep the exe in a stable folder such as `%LOCALAPPDATA%\Programs\SharePointDiff`, and use **Use this copy** after moving it. Before using the option:
- pin the tray icon;
- exit the agent before republishing;
- remove v1.0 copies, which ignore the folder lock;
- unblock a downloaded ZIP.

Untick the option (always possible, even when the conditions for turning it on no longer hold), or choose Exit in the tray menu, to stop it. **Check now** stops a logon check that is waiting for the network or a retry and checks at once. A Compare or Apply in the main window drops the agent's out-of-date prompt. A stopped or partly failed Sync is reported as incomplete, never as up to date.

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

Chạy `scripts\Package.ps1` bằng PowerShell 7 để build, chạy kiểm thử và tạo ZIP cùng SHA-256 trong `dist\v<version>`. Script chỉ đóng gói danh sách tệp phát hành xác định sẵn, kiểm tra số mục và hash executable trong ZIP; không tự đăng lên GitHub. `dist` không được đưa vào Git.

Kiểm thử lõi dùng tệp giả lập, HTTP giả lập và đồng hồ giả: hash độc lập, Unicode, chọn từng tệp/tất cả, không tải khi quét, giữ local-only, backup, nguồn/local thay đổi, tải sai, metadata, huỷ/thử lại, HTTP Range và throttling. Phần mới gồm: tải tiếp khi treo hoặc đứt mạng, xoay vòng sao lưu và dọn cache, khoá thư mục, kiểm tra nền, và cache đăng nhập DPAPI cùng chặn token có quyền ghi. Không ghi lên SharePoint hoặc hồ sơ dự án thật.

Chạy giao diện với 10.000 tệp giả, không kết nối hay cập nhật dữ liệu:

```powershell
.\publish\win-x64\SharePointDiff.exe --demo
```

Chế độ kiểm tra giao diện dành cho phát triển xuất ảnh và JSON vào thư mục thử đã chỉ định:

```powershell
.\publish\win-x64\SharePointDiff.exe --demo --ui-smoke D:\Temp\SharePointDiff-UI
```

Các cờ khác:
- `--demo --agent` chạy agent khay với dữ liệu giả (không đọc cấu hình, registry hay mạng).
- `--demo --agent --ui-smoke <thư-mục>` kiểm tra agent và thẻ thông báo.
- Hai lệnh smoke thoát `0` khi đạt, `3` khi không đạt.
- `--demo --agent --fatal-smoke tray` kiểm tra lỗi phát sinh từ menu khay.
- `scripts\Smoke.ps1 -Exe <đường-dẫn-exe>` (PowerShell 7) chạy cả bốn lượt với cấu hình cách ly và xác nhận không đụng gì ngoài `tmp\`.

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
| `SharePointDiff.App` | WPF/MVVM, Fluent light theme, tài nguyên VI/EN, hồ sơ và lỗi chi tiết; agent khay (`--agent`) |
| `SharePointDiff.Core` | MSAL (cache DPAPI), Graph/REST chỉ đọc, kiểm kê, hash, kế hoạch cập nhật, cache/backup/manifest, xoay vòng, kiểm tra nền |
| `SharePointDiff.Tests` | Kiểm thử dữ liệu thử và HTTP giả lập |

Nguồn kỹ thuật:

- [WPF trên .NET 10](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/whats-new/net100)
- [MSAL system browser và localhost redirect](https://learn.microsoft.com/en-us/entra/msal/dotnet/acquiring-tokens/using-web-browsers)
- [Microsoft QuickXorHash algorithm](https://learn.microsoft.com/en-us/onedrive/developer/code-snippets/quickxorhash?view=odsp-graph-online)
- [Graph hashes và giới hạn hash được cung cấp](https://learn.microsoft.com/en-us/graph/api/resources/hashes?view=graph-rest-1.0)
- [Graph delta và phân trang](https://learn.microsoft.com/en-us/graph/api/driveitem-delta?view=graph-rest-1.0)
- [Graph content download và Range](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0)
- [Graph throttling](https://learn.microsoft.com/en-us/graph/throttling)
