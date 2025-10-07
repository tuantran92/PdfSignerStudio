# Third-Party Notices

Tài liệu này liệt kê các thành phần bên thứ ba được sử dụng bởi **PdfSignerStudio** và giấy phép tương ứng.
Nếu không có ghi chú khác, toàn văn giấy phép gốc được đặt trong thư mục `licenses/` của gói phân phối.

> Lưu ý: Danh sách này được tổng hợp dựa trên mã nguồn dự án hiện tại. Nếu bạn thêm/bớt gói,
> vui lòng cập nhật tài liệu này cho phù hợp.

---

## iText 7 / iText Bouncy Castle Adapter
- Phiên bản: (ví dụ) 9.3.0
- Giấy phép: **GNU Affero General Public License v3.0 (AGPL-3.0)**
- Lưu ý: yêu cầu cung cấp **Corresponding Source** của toàn bộ ứng dụng cho tất cả người dùng nhận binary.

## Bouncy Castle (được iText sử dụng)
- Giấy phép: **Bouncy Castle License** (kiểu MIT/ISC)
- Ghi chú: thường đi kèm qua adapter của iText; nếu dự án tham chiếu trực tiếp, cần kèm giấy phép tương ứng.

## Microsoft WebView2 (SDK cho WinForms/WPF)
- Giấy phép: **MIT** (cho SDK).
- Ghi chú: **WebView2 Runtime** được phép phân phối lại theo điều khoản của Microsoft; nếu đóng gói kiểu Fixed, cần kèm runtime.

## Microsoft .NET Libraries & Analyzers
Bao gồm (nhưng không giới hạn):
- `System.Text.Json` (kể cả SourceGeneration)
- Roslyn/Windows Forms analyzers và các gói analyzer khác của Microsoft
- Giấy phép: **MIT**

## Microsoft Office Interop (Word/Office)
- Giấy phép/EULA: theo **Microsoft Office** (sản phẩm thương mại).
- Yêu cầu: thiết bị chạy automation phải có **bản quyền Office hợp lệ**; không dùng cho xử lý server-side.

## pdf.js (được tải qua CDN trong UI web nhúng)
- Tác giả: Mozilla
- Giấy phép: **Apache License 2.0**
- Ghi chú: nếu bạn phân phối kèm bản đã chỉnh sửa hoặc bundle lại, hãy kèm toàn văn Apache-2.0.

---

### Gợi ý bố trí thư mục giấy phép
```
/licenses/
  AGPL-3.0.txt                (toàn văn AGPLv3 cho iText 7)
  MIT.txt                     (bản MIT dùng chung cho các gói Microsoft)
  Apache-2.0.txt              (toàn văn Apache 2.0 cho pdf.js)
  BouncyCastle-License.txt    (toàn văn giấy phép Bouncy Castle)
```
Hãy đảm bảo các tệp giấy phép ở trên được phân phối cùng binary, và cập nhật lại danh sách nếu thêm gói mới.
