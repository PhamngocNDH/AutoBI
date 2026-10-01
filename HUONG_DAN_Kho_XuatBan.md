# AutoBI - Kho & Xuất Bán

Đổ **tồn kho** (BI 4286) và **xuất bán ngành Điện thoại** (BI 77) theo cụm siêu thị, xem theo hãng / nhân viên / sản phẩm / IMEI và tải Excel. Dùng trên **máy tính** (Chrome hoặc Edge).

## Cài đặt (1 lần)

1. Cài tiện ích **Tampermonkey** cho Chrome / Edge.
2. Bấm link cài:
   **https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI_Kho_XuatBan.user.js**
   → Tampermonkey mở trang cài → bấm **Cài đặt** (Install).
3. Chrome / Edge mới: vào trang quản lý tiện ích, bật **Chế độ nhà phát triển** (Developer mode) nếu Tampermonkey báo cần bật.
4. Báo anh Ngọc (38967) thêm **mã nhân viên** của bạn vào danh sách được dùng. Chưa có tên thì script báo "không có trong sheet Auth" và không chạy.

Script tự cập nhật khi có bản mới (Tampermonkey kiểm tra định kỳ). Muốn cập nhật ngay: Tampermonkey → Bảng điều khiển → **Kiểm tra cập nhật**.

## Khai báo lần đầu

1. Đăng nhập BI (`report.mwgroup.vn`) như bình thường.
2. Bấm nút **AutoBI · Kho & Xuất Bán** góc dưới phải.
3. Mở **⚙️ Cài đặt siêu thị** → nhập **Mã ST** và **Tên hiển thị** từng siêu thị của bạn (ví dụ `5243` – `Hải Anh`) → **Lưu cài đặt**.
   Tài khoản BI của bạn phải có quyền xem các siêu thị đó.

## Dùng hằng ngày

**🛒 Xuất bán**
- Chọn kỳ (Hôm nay, Hôm qua, 7 ngày, Tháng này, Tháng trước hoặc tự chọn ngày) → **Đổ xuất bán**.
- BI chỉ cho 5 lần lấy mỗi phút nên script tự giãn nhịp: khoảng 12 giây mỗi ngày. Lần đầu đổ cả tháng mất 6–7 phút; các lần sau chỉ lấy ngày còn thiếu và 7 ngày gần nhất (để bắt đơn khách nhập trả).
- Trong lúc đang đổ, **không tự bấm báo cáo 77** trên BI.
- Xem: Theo hãng · Nhân viên × hãng · Nhân viên × sản phẩm (có IMEI) · Sản phẩm · Theo ngày · Đơn treo · Nhập trả · Chi tiết (tìm theo IMEI, mã đơn, nhân viên).
- **Xem số đã lưu**: xem lại ngay, không cần gọi BI.

**📦 Tồn kho**
- Chọn siêu thị → **Đổ tồn kho** (khoảng 3 giây mỗi siêu thị).
- Lọc Ngành / Nhóm hàng / Hãng / Trạng thái, gõ IMEI hoặc tên sản phẩm để tìm.

**⬇ Tải Excel**: xuất đúng tab và bộ lọc đang xem.

## Cách tính số bán

- Báo cáo 77, Hình thức xuất = Xuất bán hàng tại siêu thị, Tìm theo = Kho tạo, ngành 13 - Điện thoại.
- Chỉ tính dòng **Đã xuất – Đã giao – Chưa hủy**. Đơn có hàng **khách nhập trả** thì bỏ cả đơn.
- Doanh thu = Giá bán × SL (đã gồm VAT). Nhân viên = người tạo đơn.
- Số được lưu trên máy tính của từng người; không chia sẻ giữa các máy.
- Không lưu tên, số điện thoại, địa chỉ khách hàng.

## Lỗi thường gặp

| Thông báo | Cách xử lý |
|---|---|
| Không có trong sheet Auth / chưa ACTIVE | Báo anh Ngọc thêm mã nhân viên |
| BI không trả dữ liệu, đăng nhập lại | Phiên BI hết hạn: tải lại trang, đăng nhập, bấm lại |
| Mã kho tạo nằm ngoài cấu hình | Kiểm tra lại danh sách siêu thị trong Cài đặt |
| Có ngày lỗi | Bấm **Đổ xuất bán** lần nữa, script tự lấy ngày còn thiếu |
| Cần hỗ trợ | Mở **📋 Nhật ký** → **Sao chép** → gửi cho anh Ngọc |
