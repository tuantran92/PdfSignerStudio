# PdfSignerStudio

Ứng dụng desktop nội bộ để mở DOCX/PDF, quét thẻ ký từ DOCX và đặt trường chữ ký lên PDF.  
Giao diện xem/chỉnh PDF hiển thị qua WebView2; phần xem PDF sử dụng pdf.js.

## Giấy phép & Nghĩa vụ (AGPLv3)

Ứng dụng này liên kết với **iText 7** theo mô hình **AGPLv3**. Vì vậy:
- Mã nguồn đầy đủ của **toàn bộ ứng dụng** (bao gồm mọi thay đổi, patch và script build)
  phải được cung cấp cho **tất cả người dùng nhận binary**.
- Chúng tôi cung cấp mã nguồn tại: **<điền URL repo>**.
- Trong UI (Help → About) và gói cài đặt có đường link tới repo và tệp `LICENSE-AGPL`.

> Nếu công ty muốn giữ đóng nguồn, hãy thay iText7 bằng thư viện giấy phép phù hợp khác **hoặc**
> mua license thương mại của iText. (Tệp này phản ánh tuân thủ khi dùng bản AGPL.)

## Tính năng chính (theo mã nguồn)

- **Mở DOCX/PDF**; nếu là DOCX, quét cặp thẻ `<<<Name>>>` để sinh trường chữ ký ở PDF đầu ra.  
- **Đặt/đổi kích thước trường chữ ký** trên trang PDF, đảm bảo không vượt quá crop box.  
- **UI hiển thị PDF** dựa vào WebView2 + pdf.js, kèm sidebar template/field.

## Cách build từ source

### Yêu cầu
- Windows 10/11  
- .NET SDK: **<điền phiên bản, ví dụ .NET 8.0.x>**  
- (Khuyến nghị) Visual Studio **<điền phiên bản>**  
- Microsoft Office (nếu dùng tính năng chuyển DOCX → PDF qua Interop)  
- WebView2 Runtime (máy người dùng): dùng Evergreen bootstrapper hoặc đóng gói Fixed Runtime.

### Quy trình build
```bash
dotnet restore
dotnet build -c Release
dotnet publish -c Release -r win-x64 /p:PublishSingleFile=true
```

### Đóng gói phát hành nội bộ
Tạo thư mục `dist/<version>` chứa:
- Binary đã publish
- `NOTICE.txt`
- `LICENSE-AGPL` (toàn văn AGPLv3)
- `THIRD-PARTY-NOTICES.md`
- `README-AGPL.md`

(Nếu dùng WebView2 Fixed Runtime, kèm theo bộ cài Fixed theo hướng dẫn của Microsoft.)

## Hướng dẫn sử dụng nhanh

1. **Mở file** (Ctrl+O) — chọn `.docx` hoặc `.pdf`.  
2. Với `.docx`, dùng thẻ `<<<SignatoryName>>>` để đánh dấu vị trí ký; ứng dụng sẽ chuyển DOCX → PDF và
   sinh các trường chữ ký tương ứng.  
3. **Di chuyển/đổi kích thước** các trường; dùng nút Export để xuất ra PDF có trường chữ ký.  

## Thành phần bên thứ ba

- **iText 7 & Bouncy Castle Adapter** — thao tác trường chữ ký PDF.  
- **Microsoft Office Interop (Word)** — chuyển DOCX → PDF & quét thẻ.  
- **Microsoft WebView2 (WinForms)** — nhúng UI web trong WinForms.  
- **pdf.js (CDN)** — render/preview PDF trong UI nhúng.  
- (Các thư viện .NET/Microsoft khác đi kèm dự án.)

> Các giấy phép tương ứng được liệt kê trong `THIRD-PARTY-NOTICES.md` và/hoặc thư mục `licenses/`.

## Ghi chú tuân thủ bổ sung

- **Link tới mã nguồn** phải xuất hiện trong About/Help của ứng dụng và trong `NOTICE.txt`.  
- **Version sync**: commit/tag của repo cần khớp bản binary phát hành.  
- **Lưu bằng chứng cung cấp source** (ví dụ tệp “PHÁT-HÀNH-NỘI-BỘ.md” ghi lại version, URL repo, phạm vi cấp quyền).  
- **Office Interop** chỉ chạy hợp lệ khi máy có bản quyền Office; không dùng automation Office trên server.

## Hỗ trợ nội bộ

- Nhóm phụ trách: <tên nhóm/email>  
- Kênh liên lạc: <Teams/Slack/Email>
