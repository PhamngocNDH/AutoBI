# AutoBI

Kho lưu AutoBI của 38967 - Mr Phạm.

Bản hiện tại: **18.56**

Link cài cố định: [AutoBI.user.js](https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI.user.js)

## Thay đổi 18.56 — đổi tiêu đề cột TH thành HT ở Chi tiết NV

- Bảng nhóm hàng trong thẻ nhân viên ở Chi tiết NV đổi tiêu đề cột từ `TH` thành `HT` (trên màn hình báo cáo: `TH / Target` thành `HT / Target`; trong file Excel mẫu và đường xuất Excel cũ cũng đổi theo để khớp nhau).
- Chỉ đổi nhãn tiêu đề, số liệu và công thức không đổi.

## Thay đổi 18.55 — NV làm chéo shop

- Nhân viên sang shop kia làm (ví dụ Thuý ở Trực Cường sang Hải Anh bán iPhone 44.9) trước đây chỉ hiện ở shop nhà trong sheet "Dán TĐ nhân viên", còn file đổ tay thì ma trận chỉ lấy dòng đầu tiên gặp nên cũng thiếu một bên. Nay tool đổ NV theo từng shop (bán ở shop nào thì có dòng ở shop đó) và ma trận trong mẫu cộng dồn tất cả các dòng cùng nhóm + mã NV (SUMIF thay cho MATCH lấy dòng đầu).
- Đã kiểm chứng bằng số 26-10: iPhone của Thuý ra đúng 245.32 (44.9 Hải Anh + 200.42 Trực Cường), Bảo hiểm HomeCredit 130.16, Panasonic của Khánh 23.55, của Hưởng 29.70; NV không làm chéo shop không đổi.
- Tổng DT lũy kế của NV (sheet Dán DT) không đổi ở bản này — số đó báo cáo đã gộp 2 shop, vẫn thấp hơn bảng "Doanh thu hợp nhất" trên BI khoảng 1–2% nên muốn khớp hệt BI thì dán đè 2 sheet Dán DT từ BI như trước.

## Thay đổi 18.54 — sửa lỗi đổ số không khớp

- Xuất Excel đổ tay không còn chèn dòng trống sau mỗi khối nhóm hàng: trước đây dòng trống làm đứt công thức cột K ở mẫu, cột M trống, toàn bộ nhóm hàng về 0.
- Mẫu `AutoBI_mau_do_tay.xlsx`: cột K giữ tên chương trình qua dòng trống (phòng khi dán tay từ BI có dòng trống), ô vàng `Xem` (B2) ở Realtime / Lũy kế / Chi tiết NV đã có danh sách chọn theo đúng tên siêu thị trong Khai báo.
- Khi đổ số, sheet Dán được xóa hết 1500 dòng trước khi ghi, tránh số cũ còn sót khi lượt sau ít nhóm / nhân viên hơn.
