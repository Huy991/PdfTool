# PdfTool

Bộ công cụ PDF cho Windows, giao diện tiếng Việt, tương tự PDF24.

**Soạn PDF**: duyệt thư mục bên trái, kéo thả file PDF/ảnh sang bên phải để xem trang dạng ảnh thu nhỏ; kéo trang để sắp xếp hoặc chuyển giữa các tài liệu, xoay, xóa, tách, gộp rồi lưu.

**Hộp công cụ**: gộp, tách, trích xuất/sắp xếp trang, xóa trang, xoay, nén, ảnh sang PDF, PDF sang ảnh, nhận dạng chữ (OCR tiếng Việt/Anh), đánh số trang, đóng dấu chữ, đặt/gỡ mật khẩu.

## Biên dịch

Chạy `build.bat`. Dùng `csc.exe` có sẵn của .NET Framework 4 trên Windows, không cần cài Visual Studio.

## Công cụ đi kèm

PdfTool gọi các chương trình mã nguồn mở [qpdf](https://github.com/qpdf/qpdf), [Ghostscript](https://ghostscript.com) và [Tesseract](https://github.com/tesseract-ocr/tesseract). Chúng được tìm theo thứ tự:

1. Thư mục `tools\` cạnh `PdfTool.exe` (`tools\qpdf\bin\qpdf.exe`, `tools\gs\bin\gswin64c.exe`, `tools\tesseract\tesseract.exe`)
2. Thư mục cài PDF24 (`C:\Program Files\PDF24`)
3. Biến môi trường `PATH`

## Đóng gói gửi máy khác

Trên máy đã cài PDF24, chạy `build.bat` rồi `package.bat`. Kết quả là `dist\PdfTool.zip`, đã kèm sẵn các công cụ trên và dữ liệu OCR, máy nhận chỉ cần giải nén rồi chạy `PdfTool.exe`.

## Giấy phép các thành phần

qpdf và Tesseract: Apache 2.0. Ghostscript: AGPL 3.0 (muốn bán phần mềm kèm Ghostscript cần giấy phép thương mại của Artifex).
