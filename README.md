# Phần mềm Bóc tách & Gộp báo cáo

Ứng dụng Windows chạy cục bộ, dùng để nhận diện các phần/mục trong nhiều báo cáo có cùng cấu trúc và nối nguyên văn nội dung của các mục tương ứng.

**Phiên bản hiện tại: v0.2**

## Chức năng

- Nhập nhiều file `.docx` và `.pdf` có lớp text; không xử lý PDF scan/OCR.
- Kéo thả file vào cửa sổ hoặc chọn bằng hộp thoại.
- Tự động nhận diện tiêu đề theo số La Mã, số Ả Rập, chữ cái, định dạng in đậm và vị trí trong cây mục.
- Gộp tự động theo tên/số thứ tự/vị trí cấu trúc; cho phép đổi tên, gộp, tách, chọn bỏ và sắp xếp thủ công.
- Xem trước nội dung nguyên văn theo từng báo cáo nguồn.
- Lưu/nạp bảng quy đổi tên mục để dùng lại.
- Xuất Word và Excel.
- Trang “Ủng hộ tác giả” với mã QR ngân hàng và các kênh liên hệ/góp ý.
- Không thay đổi file gốc và không đưa tài liệu lên Internet.

## Cách dùng nhanh

1. Mở `BoctachvaGopBaoCao.exe` hoặc nhấp đúp `Chay_phan_mem.bat`.
2. Thêm/kéo thả các báo cáo.
3. Đổi tên nguồn thành tên xã/đơn vị nếu cần.
4. Bấm **Phân tích cấu trúc**.
5. Kiểm tra cây mục; dùng **Gộp mục**, **Tách theo nguồn**, **Đổi tên** và phím mũi tên để hiệu chỉnh.
6. Xuất Word hoặc Excel.

## Phạm vi “nguyên văn”

Ứng dụng giữ nguyên chữ, số và thứ tự các đoạn nội dung. Các bảng Word được sao chép vào file đầu ra. Kiểu trình bày của file tổng hợp được chuẩn hóa thành A4, Times New Roman; một số đối tượng nhúng đặc biệt (SmartArt, biểu đồ, công thức hoặc ảnh nổi) không thuộc phạm vi bản 1.0.

## Tối ưu hóa chất lượng gộp báo cáo

Các bạn nên sửa lại báo cáo nguồn cho đồng bộ cấu trúc Phần Chuong Mục Điều Khoản ... cho thống nhất cùng một đề cương để phần mềm nhận diện chính xác hơn.

Một số báo cáo sẽ viết liền cả tiêu đề vào nội dung như kiểu sau:

b) Cơ sở vật chất: Cơ sở vật chất cơ bản đáp ứng....

Thì ta nên sửa như sau để phần mềm nhận diện đúng mục b.

b) Cơ sở vật chất
Cơ sở vật chất cơ bản đáp ứng....

## Lịch sử phiên bản

### v0.2

- Sửa lỗi **“Không thể phân tích”** với một số file `.docx` không có thành phần `word/numbering.xml` hoặc quan hệ cấu hình đánh số.
- Cấu hình đánh số trong DOCX giờ được xem là thành phần tùy chọn. Phần mềm vẫn nhận diện các tiêu đề được gõ thủ công và các kiểu Heading khi thành phần này không tồn tại.
- Bổ sung kiểm thử hồi quy cho tài liệu không có cấu hình đánh số.
- Đã kiểm tra trực tiếp với file từng gây lỗi `CN Lac Son.docx`: phân tích thành công 76 mục, không có cảnh báo.

### v0.1

- Phiên bản phát hành đầu tiên.
