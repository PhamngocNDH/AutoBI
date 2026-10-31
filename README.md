# AutoBI

Kho lưu AutoBI của 38967 - Mr Phạm.

Bản hiện tại: **18.67**

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
