# Web app tạo Hồ sơ CTV VNG

Chạy 100% trong trình duyệt (không server, không upload dữ liệu đi đâu). Upload `data.xlsx`, sinh ra:

- **Hợp đồng** (PDF) — vẽ đè lên `assets/Template_CTV_new.pdf` (mẫu mới do Legal cung cấp, 09/2026 — thay cho mẫu `template_co_chuky.pdf` cũ).
- **BBNT** (docx) — Biên bản nghiệm thu, tự chọn mẫu `BBNT_1..4_item.docx` theo số hạng mục công việc.
- **Mẫu 02/TNDN** (docx) — bảng kê thu mua không hóa đơn, liệt kê tất cả CTV trong 1 file.

Tất cả đóng gói thành 1 file `HoSo_CTV.zip` để tải về.

## Cách chạy

**Chỉ cần double-click mở `index.html`** — không cần cài gì, không cần server. Toàn bộ PDF/font/docx template đã được nhúng sẵn dạng base64 ngay trong `js/embeddedAssets.js` (xem `build_embedded_assets.py`), nên không còn `fetch()` file cục bộ nào bị trình duyệt chặn khi mở qua `file://` nữa.

(Vẫn có thể chạy qua static server bình thường nếu muốn, ví dụ `python -m http.server 8080` trong thư mục `webapp` — không bắt buộc.)

## Canh chỉnh tọa độ trên PDF (khi thấy trường bị lệch)

Mở `tools/coord-editor.html` (double-click, không cần server) — mini-tool kéo-thả:

- Nền là ảnh chụp thật của `Template_CTV_new.pdf`, mỗi trường (Họ tên, CCCD, ngân hàng...) hiện thành 1 nhãn đỏ tại đúng tọa độ hiện tại trong `pdfContract.js`.
- Kéo nhãn cho đúng chỗ, hoặc bấm chọn 1 nhãn rồi dùng phím mũi tên (giữ Shift = nhảy 5pt) hoặc gõ số X/Y trực tiếp ở khung bên trái để tinh chỉnh.
- Có thể kéo thả `data.xlsx` vào để xem preview với dữ liệu thật (chọn từng người trong danh sách) thay vì dữ liệu mẫu.
- Bấm **"Xem thử PDF thật"** để render bằng đúng engine pdf-lib với tọa độ đang chỉnh — xem trực tiếp trong khung bên dưới, không cần đoán qua nhãn kéo-thả.
- Xong thì bấm **"Tạo code để copy"** → **"Sao chép"**, rồi dán đè vào khối `const COORDINATES = {...}` trong `js/pdfContract.js`.
- Nếu Legal đổi template PDF lần nữa: thay file `assets/Template_CTV_new.pdf`, chạy lại `python tools/build_tool_assets.py` để tool lấy ảnh nền mới, rồi mở lại tool để canh lại tọa độ.

## Cấu trúc

- `js/embeddedAssets.js` — dữ liệu base64 của PDF/font/docx template (tự sinh bởi `build_embedded_assets.py`, không sửa tay).
- `js/embeddedAssetsHelper.js` — giải mã base64 → ArrayBuffer.
- `js/xlsxParser.js` — đọc `data.xlsx` bằng SheetJS.
- `js/pdfContract.js` — điền PDF hợp đồng bằng pdf-lib theo dict `COORDINATES` (tọa độ đo trực tiếp trên `Template_CTV_new.pdf` bằng PyMuPDF, không còn liên quan tọa độ mẫu cũ). Load template PDF gốc 1 lần rồi dùng `copyPages()` để nhân bản cho từng người (nhanh và ổn định hơn gọi `PDFDocument.load()` lặp lại). Trường "Thời hạn thực hiện" gộp 2 cột Excel `Time`/`Time2` thành 1 chuỗi "từ - đến" vì mẫu mới chỉ có 1 dòng trống duy nhất cho mục này.
- `js/vnNumberWords.js` — đọc số tiền ra chữ tiếng Việt (dùng riêng cho tổng tiền Mẫu 02, vì đó là số tự cộng dồn nhiều người, không có sẵn trong Excel như cột `Bang_chu` của từng người).
- `js/docxFill.js` — điền BBNT/Mẫu 02 bằng docxtemplater, có bảng `FIELD_MAP` ánh xạ tên tag trong template (đặt theo Word field code cũ, có gạch dưới) sang đúng tên cột trong `data.xlsx` (có dấu cách). Mẫu 02 có thêm `MAU02_TOP_TAGS` (Ngày/Tháng/Năm lập bảng kê, lấy từ dòng đầu tiên) và tự tính `Tong_M2`/`Tong_M2_chu` (tổng tiền + tổng tiền bằng chữ) — 2 tag này nằm NGOÀI vòng lặp `{#rows}` nên phải truyền riêng ở cấp cao nhất khi render, không đặt chung với dữ liệu từng người.
- `js/main.js` — điều phối upload → preview → generate → zip → download.
- `vendor/` — thư viện JS vendor sẵn (không phụ thuộc CDN, chạy được cả mạng nội bộ/offline).
- `source_docx/` — 4 file `BBNT_1..4_item.docx` gốc (chưa qua xử lý), phục hồi từ `HD_VNG_11.04.2026.exe`.
- `templates/` — bản đã tiền xử lý của `Mau_02.docx` và `BBNT_1..4_item.docx`. `preprocess_templates.py` làm 4 việc: (1) field code Word gốc → tag `{...}` cho docxtemplater, (2) bọc loop `{#rows}` quanh dòng người bán của Mẫu 02, (3) xóa 9 dòng bảng thừa còn sót lại từ thiết kế cũ (tối đa 10 người/bảng, đã thay bằng vòng lặp tự động), (4) chèn tag `{Tong_M2}`/`{Tong_M2_chu}` vào dòng "Tổng giá trị hàng hóa..." — và **xóa `<w:mailMerge>` khỏi `word/settings.xml` của cả 5 file** (nguyên nhân hộp thoại "Opening this document will run the following SQL command" khi mở bằng Word, do template gốc từng nối thật với `data.xlsx` qua Mail Merge). Chạy lại script này nếu sửa template gốc, rồi chạy `python build_embedded_assets.py` để cập nhật `js/embeddedAssets.js`.
- `tools/coord-editor.html` — mini-tool canh chỉnh tọa độ bằng kéo-thả (xem mục riêng bên trên). `tools/build_tool_assets.py` sinh ảnh nền + font cho tool này.

## Lưu ý khi sửa code (browser cache file JS)

Trình duyệt có thể cache file `.js` cũ và không tự nhận bản mới sau khi cập nhật tool (đã gặp thực tế lúc test). Mỗi file trong `js/` được nạp qua `index.html`/`tools/coord-editor.html` với query `?v=N` ở cuối — **mỗi khi sửa bất kỳ file `.js` nào, phải tăng số `N` này trong cả 2 file HTML**, nếu không trình duyệt có thể vẫn chạy code cũ dù đã sửa xong.

## Đã kiểm thử

Đã chạy thử với `data.xlsx` thật (17 CTV): 17 hợp đồng PDF (theo `Template_CTV_new.pdf`), 17 BBNT, 1 Mẫu 02 liệt kê đủ 17 người, tổng tiền + tổng tiền bằng chữ tự tính đúng (khớp mẫu tham chiếu "48,000,000 - Bốn mươi tám triệu đồng"), ngày lập bảng kê ở đầu trang hiện đúng, không còn dòng thừa/chữ "undefined", không còn `<w:mailMerge>` trong file xuất ra. Đã kiểm tra thêm bằng dữ liệu giả cho cả 4 biến thể BBNT (1-4 hạng mục).

## Giả định / giới hạn hiện tại

- Cột "Tổng tiền bằng chữ" của TỪNG NGƯỜI (BBNT) lấy trực tiếp từ Excel; riêng tổng tiền BẰNG CHỮ của Mẫu 02 (tổng nhiều người cộng lại) được tự tính bằng `vnNumberWords.js`.
- Nếu 1 dòng không có hạng mục công việc nào (`Items 1..4` đều trống), dòng đó bị bỏ qua khi tạo BBNT (vẫn tạo Hợp đồng và có mặt trong Mẫu 02 nếu có Họ tên).
- BBNT chọn 1 trong 4 mẫu có sẵn (`BBNT_1..4_item.docx`) theo số hạng mục, không tự thêm dòng động — nhưng vì Excel cũng chỉ hỗ trợ tối đa 4 hạng mục (`Items 1..4`) nên không có trường hợp vượt quá.
