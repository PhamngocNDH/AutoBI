# AutoBI

Kho lưu AutoBI của 38967 - Mr Phạm.

Bản hiện tại: **18.54**

Link cài cố định: [AutoBI.user.js](https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI.user.js)

## Thay đổi 18.54 — sửa lỗi đổ số không khớp

- Xuất Excel đổ tay không còn chèn dòng trống sau mỗi khối nhóm hàng: trước đây dòng trống làm đứt công thức cột K ở mẫu, cột M trống, toàn bộ nhóm hàng về 0.
- Mẫu `AutoBI_mau_do_tay.xlsx`: cột K giữ tên chương trình qua dòng trống (phòng khi dán tay từ BI có dòng trống), ô vàng `Xem` (B2) ở Realtime / Lũy kế / Chi tiết NV đã có danh sách chọn theo đúng tên siêu thị trong Khai báo.
- Khi đổ số, sheet Dán được xóa hết 1500 dòng trước khi ghi, tránh số cũ còn sót khi lượt sau ít nhóm / nhân viên hơn.
