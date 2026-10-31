# Đọc ảnh BI → đổ số AutoBI (bản 1.3)

Trang Apps Script đọc ảnh chụp màn hình BI ra số liệu, xuất đúng file mẫu đổ tay `AutoBI_mau_do_tay.xlsx` (hoặc chép bảng để dán). Đây là tầng **Thu thập** mới của AutoBI; phần tính toán và báo cáo trong file Excel giữ nguyên.

## Cài đặt (1 lần, trên máy tính)
1. Vào script.google.com → **Dự án mới** → đặt tên `Đọc ảnh BI`.
2. Tạo 3 tệp, dán nội dung tương ứng:
   - `Code.gs` (tệp có sẵn, xóa hết code cũ rồi dán)
   - `+` → HTML → tên `index` → dán `index.html`
   - `+` → HTML → tên `logic` → dán `logic.html`
3. **Cài đặt dự án** (bánh răng) → **Thuộc tính tập lệnh** → Thêm:
   - `AI_KEY` = khóa AI. Dùng được khóa **Gemini** (aistudio.google.com → Get API key) hoặc khóa **Claude** (`sk-ant-…`, console.anthropic.com); tool tự nhận theo dạng khóa.
   - `MODEL` = không bắt buộc (mặc định `gemini-3.1-flash-lite` — miễn phí, nhanh — hoặc `claude-sonnet-5-5`)
   - `AUTOBI_KEY` = một chuỗi bí mật tự đặt (nên có). AutoBI phải gửi đúng chuỗi này mới lấy được số.
4. **Triển khai → Tác vụ triển khai mới → Ứng dụng web**
   Thực thi với tư cách: **Tôi** · Ai có quyền truy cập: **Bất kỳ ai có tài khoản Google** → Triển khai → cấp quyền → chép URL `/exec`.
5. Lần sau sửa code: **Quản lý các lần triển khai → bút chì → Phiên bản mới** để giữ nguyên URL.

## Dùng hằng ngày (điện thoại)
1. Trên BI, chụp:
   - **Doanh thu hợp nhất**, chế độ DT quy đổi + Lũy kế, chụp **từng siêu thị** (mỗi siêu thị 1 ảnh có thẻ số).
   - Tab **Nhân viên** của Doanh thu hợp nhất (chọn cả 2 siêu thị).
   - **Thi đua → Siêu thị** · Lũy kế · Danh sách: cuộn hết, mỗi màn 1 ảnh.
   - **Thi đua → Nhân viên** · Lũy kế · Danh sách: cuộn hết, mỗi màn 1 ảnh.
2. Mở link → **Chọn ảnh** (cả lô) → **Đọc ảnh**.
3. Xem ô vàng cảnh báo. Bảng nào "thiếu x" thì chụp bù phần cuối bảng đó, chọn thêm ảnh rồi bấm Đọc ảnh lần nữa.
4. **Chọn file của shop** (file đổ tay đã khai báo) → **Tải file Excel**. Tool chỉ thay số ở các sheet Dán + Nhập số, Khai báo giữ nguyên.
   Không có file thì để trống, tool tự lấy mẫu trên GitHub và điền sẵn siêu thị, nhóm, NV. Khi đó cần tự điền cột "Siêu thị (1/2)" của NV ở sheet Khai báo.

## Nối với AutoBI (bản 1.1 + AutoBI V18.57)
1. Đọc ảnh xong, ở khung **Gửi sang AutoBI**: nhập **Mã nhận số** (vd 38967) → bấm **Gửi sang AutoBI**.
2. Mở AutoBI → **Chạy báo cáo** → chọn **Nạp số từ ảnh** → **Nạp số và xem**.
   Lần đầu AutoBI hỏi 3 thứ: link trang này (dòng "Link dán vào AutoBI" hiện ở khung Gửi), mã nhận số, khóa `AUTOBI_KEY`.
3. AutoBI hiện LK NEW, Chi tiết NV, Auto Comment, ảnh LINE, Xuất Excel như khi lấy số từ BI.
- Bản 1.1 lưu số vào Drive của chủ dự án (tệp `DocAnhBI_<mã>.json`). Lần triển khai đầu sau khi cập nhật, Google sẽ hỏi thêm quyền Drive.
- Mỗi mã nhận chỉ giữ gói số gửi gần nhất. Gửi lại sẽ ghi đè.

## Mẹo chụp cho đúng số
- Ảnh cuộn nên chồng lên nhau một chút (tool tự bỏ dòng trùng). Dòng bị cắt nửa ở mép ảnh sẽ được bỏ qua.
- Mỗi lần cuộn để thấy được ô tên chương trình ở đầu bảng.
- Cột % bị cắt "184...." không sao: tool chỉ lấy Doanh thu và Target, còn % thì file Excel tự tính.
- Tool kiểm tra tổng các dòng có bằng dòng TỔNG không. Bảng nào khớp có dấu ✓ đủ.

## Cấu trúc
| Tệp | Việc |
|---|---|
| Code.gs | `cauHinhAI_`/`goiAI_` (Gemini hoặc Claude), `kiemTraKhoa` (chạy tay để thử khóa), `doGet` (mở trang, hoặc `?action=lay&ma=` trả gói số cho AutoBI), `thongTin`, `docAnh(b64)` gọi AI đọc 1 ảnh → JSON, `luuKetQua(ma, json)` lưu gói số vào Drive |
| logic.html | `DAB.gop` (gộp nhiều ảnh, bỏ trùng, kiểm TỔNG), `tatCaBang` (dựng bảng giống "Chép bảng" của Core `fillDoTay`), `dienExcel` (điền ExcelJS), `goiAutoBI` (gói số gửi AutoBI, dạng v:1) |
| index.html | Giao diện: khai báo siêu thị (lưu trên máy), chọn ảnh (thu nhỏ, cắt ảnh dài), đọc 3 ảnh song song, xem kết quả, chép bảng, xuất Excel |

Quy ước: tên chương trình bỏ tiền tố tháng `T10 - ` giống AutoBI Core. Dòng bị coi là trùng khi cùng mã NV (hoặc tên siêu thị) và cùng số. NV làm chéo shop có 2 dòng với số khác nhau nên giữ cả 2, ma trận trong file sẽ cộng dồn.
