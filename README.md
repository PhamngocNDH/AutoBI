# AutoBI

Kho lưu AutoBI của 38967 - Mr Phạm.

Bản hiện tại: **18.75**

Link cài cố định: [AutoBI.user.js](https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI.user.js)

## Thay đổi 18.65 — Nạp số từ ảnh

- Chạy báo cáo → chế độ mới **Nạp số từ ảnh**: lấy số mà trang **Đọc ảnh BI** (thư mục `DocAnhBI/`, Apps Script) đã đọc từ ảnh chụp BI, ghi vào bộ nhớ số như API (`link2`, `link4_smart`, `link6`) rồi mở báo cáo bằng đường "Xem lại" — **không gọi BI**.
- Link trang Đọc ảnh đã gắn sẵn; mã nhận mặc định = mã NV đang dùng AutoBI. Đổi bằng nút "⚙ Link / mã nhận".
- Đợt 1: Lũy kế (DT, thi đua nhóm × siêu thị, DT + thi đua NV). Chưa có từ ảnh: Realtime, Ngành hàng BI, Giờ công, CRM, trả chậm từng NV.

## Thay đổi 18.66 — sửa đơn vị Nạp số từ ảnh

- Doanh thu siêu thị và DT QĐ nhân viên nạp từ ảnh bị nhân thêm 1.000.000 (LK NEW ra Thực hiện 5,095,000,000, tỉ lệ 109381709%). Nay giữ đúng đơn vị triệu như AutoBI vẫn dùng; thi đua nhóm hàng không đổi.

## Thay đổi 18.67 — chuyển sang portal, tắt trên BI

- AutoBI **không chạy trên baocao.dienmayxanh.com nữa** (không gọi API BI). Bật trên **portal.thegioididong.com**; CRM giữ nguyên.
- Nhận nhân viên trên portal từ thông tin đăng nhập của portal (mã NV + họ tên) → kiểm quyền như cũ.
- Cụm nút nổi (▶ Chạy · ⋯ Menu) chuyển sang **góc dưới bên trái**. Nút "Trang chủ" / Dừng quay về portal.
- Khung Chạy báo cáo trên portal chỉ còn **Nạp số từ ảnh** + **Xem lại** (các chế độ cần BI đã ẩn).

## Thay đổi 18.68 — Realtime + trả chậm từ ảnh, sửa Lưu Sheet

- Nạp số từ ảnh thêm **Realtime**: Doanh thu hợp nhất chế độ Realtime từng siêu thị (DT QĐ, target ngày, trả chậm, TLPVTC hôm nay) → `link1`; Thi đua Realtime → `link3_smart`; DT nhân viên hôm nay → `link6rt`. Realtime V1 có số.
- Thêm **trả chậm từng NV** (Báo cáo trả chậm → Nhân viên) → `link7`; khớp theo họ tên trong Khai báo. Cần thấy cột DT siêu thị (hoặc %) mới tính được tỷ lệ.
- Sửa lỗi **"Không thể lưu dữ liệu!"**: lưu Sheet bị chặn vì nhật ký phiên chạy BI cũ còn lỗi (403); số nạp từ ảnh nay lưu được. Cloud trả lỗi khác thì hiện nguyên lý do.
- Trang Đọc ảnh BI 1.5: ẩn dòng phiên bản/model AI, ẩn khung Xuất Excel đổ tay, đọc thêm màn Realtime + Báo cáo trả chậm.

## Thay đổi 18.69 — gộp số ảnh theo phần trong ngày

- Trang Đọc ảnh BI 1.6: mỗi lần **Gửi sang AutoBI** không còn ghi đè cả gói. Phần nào gói mới có (DT siêu thị, thi đua siêu thị, thi đua NV, DT NV, trả chậm NV — riêng lũy kế và realtime) thì thay; phần không có thì giữ số đã gửi trước **trong cùng ngày**. Sáng đổ lũy kế, trưa đổ realtime, ca chiều đổ cả hai đều được. Sang ngày mới bỏ hết số cũ.
- AutoBI khi Nạp số từ ảnh hiện giờ của từng phần, vd "Lũy kế 08:15 · Realtime 12:30".

## Thay đổi 18.71–18.75 — chọn ảnh ngay trong AutoBI + kiểm quyền (Đọc ảnh BI 1.7)
- Nạp số từ ảnh mở khung **Chọn ảnh** ngay trên portal: ảnh thu nhỏ trên máy → Cloud Đọc ảnh đọc (3 ảnh song song) → báo bảng thiếu/lệch TỔNG, màn chưa chụp → **Nạp số và xem**. Không cần mở trang Đọc ảnh, không cần bấm Gửi.
- Kết quả đọc giữ trong ngày: chụp bù chỉ đọc ảnh mới. Nút "Lấy số đã gửi từ trang Đọc ảnh" giữ đường cũ.
- Cloud Đọc ảnh 1.7: thêm doPost (doc / luu / quyen); đọc ảnh, gửi, lấy số đều kiểm mã NV theo danh sách quyền AutoBI (tab Auth). Tắt tạm: thuộc tính AUTH_CHECK = off.
- Trang Đọc ảnh: đọc ảnh phải có mã NV có quyền.
- 18.72 + Cloud 1.8: đọc 5 ảnh song song; Cloud nhớ model AI vừa đọc được và tạm bỏ qua model đang nghẽn (bớt ảnh chờ 3–4 phút).
- 18.73: khung Nạp số từ ảnh hiện "đọc bằng Gemini x · Mistral y"; giữ chuột lên ảnh thấy model đã đọc.
- 18.74 (lõi gộp 1.10): bảng Thi đua siêu thị gộp dòng trùng theo tên siêu thị (ghi khác kiểu vẫn nhận); chụp 2 lần đọc ra 2 số khác thì tự chọn số khớp TỔNG; bỏ dòng tổng vùng đọc lẫn; NV trùng mã mà dư đúng 1 dòng thì bỏ dòng đó. Chỉ có DT siêu thị realtime thì không báo thiếu bộ realtime.
- 18.75: bỏ các ô quy đổi/thực ở bảng đầu Realtime (giao diện Excel + file Xuất Excel), thay bằng Thời gian qua và Dự kiến (hết hôm qua); số nạp từ ảnh tính ngày đã qua / số ngày tháng theo ảnh (mùng 1 nạp số tháng trước không còn ra Dự kiến 1955%).
- Cloud 1.9: Mistral làm dự phòng (thuộc tính MISTRAL_KEY) — Gemini nghẽn thì đọc bằng Mistral, không chờ model đang nghẽn.

## Thay đổi 18.70 — trả chậm NV + thưởng NV khi Nạp số từ ảnh

- Trả chậm từng NV hiện số: ảnh "Báo cáo trả chậm" không thấy cột DT siêu thị thì tỷ lệ = DT trả góp ÷ DT QĐ của NV (cùng cách thẻ Tỷ trọng trả chậm của BI).
- Bấm **Nạp số và xem** lấy luôn **Điểm thưởng NV** (newinsite): kỳ đầu tháng → hết hôm qua, 3 NV/lượt, 1 lần/ngày (dùng lại), chờ tối đa 25s; chưa đăng nhập newinsite thì báo trong thông báo, báo cáo vẫn mở.
