# AutoBI

Kho lưu AutoBI của 38967 - Mr Phạm.

Bản hiện tại: **18.64**

Link cài cố định: [AutoBI.user.js](https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI.user.js)

## Thay đổi 18.64 — ảnh mờ / nhạt màu

- Trả độ nét ảnh về tỉ lệ 2 cho mọi ảnh (tỉ lệ 1.5 ở 18.63 làm ảnh mờ).
- Ảnh "Cả trang" bị nhạt màu: khung chụp mang theo hiệu ứng hiện dần / trượt vào của báo cáo (report-slide…), chụp ngay lúc hiệu ứng mới chạy nên ảnh nhạt (có lúc lệch). Nay khung chụp bỏ lớp hiệu ứng và tắt mọi hiệu ứng trong bản sao. Thử giả lập: 18.63 ra ảnh trắng/nhạt, 18.64 ra đúng màu tiêu đề #1f3864.
- Vẫn giữ: không chờ 0,45 giây, chỉ sao chép khung cần chụp.

## Thay đổi 18.63 — chép ảnh nhanh hơn nữa

- Bỏ thời gian chờ 0,45 giây trước khi chụp "Cả trang" (khung chụp riêng không cần chờ hộp Tùy chọn đóng).
- Ảnh lớn (rộng > 800px: Cả trang, bảng Realtime, Tổng hợp NV) chụp tỉ lệ 1.5 thay 2 → ít điểm ảnh hơn ~44%, mã hóa PNG + chép vào bộ nhớ nhanh hơn; ảnh vẫn rộng ~1.400–2.000px, đủ nét gửi LINE. Bảng nhỏ giữ tỉ lệ 2.
- Đo trên Chrome thật (Cả trang Chi tiết NV): khoảng 1,7 giây → 0,9 giây.

## Thay đổi 18.62 — chép ảnh xong máy đứng vài giây

- Thư viện chụp ảnh trước đây sao chép CẢ trang BI + toàn bộ báo cáo rồi mới chụp 1 mục → nặng, máy đứng vài giây. Nay chỉ sao chép khung cần chụp (+ CSS). Đo trên Chrome thật: 1.2–1.5 giây → 0.3 giây mỗi ảnh.
- Bước tự cắt lề chỉ đọc 4 dải mép ảnh (không đọc cả ảnh): 185 ms → 90 ms. Bỏ lệnh khóa cỡ chữ toàn trang trong bản sao (thừa).
- Ảnh ra như 18.61 (lề đều 4 phía, không cụt).

## Thay đổi 18.61 — chép ảnh: chụp lề rộng rồi tự cắt

- Trên máy thật, thư viện chụp ảnh có lúc vẽ lệch nội dung ~25px (dư lề trái, cụt bên phải — vd Cả trang ở LK NEW / Chi tiết NV), máy thử không tái hiện được. Nay mọi ảnh "Chép ảnh / Tải" (từng mục và Cả trang ở theme Bảng Excel) chụp với lề rộng 48px rồi tự cắt khoảng trống thừa, chừa lề đều 14px 4 phía → lệch trái/phải bao nhiêu cũng không mất chữ. Đã thử giả lập lệch +26px / −26px: lề vẫn đều 2 bên.

## Thay đổi 18.60 — chụp "Cả trang" (Chi tiết NV) cụt bên phải · thứ tự NV Realtime V1

- Theme Bảng Excel: mục "Cả trang đang xem" (vd Chi tiết NV) không chụp tại chỗ nữa mà dựng bản sao riêng rộng đúng theo nội dung → hết lệch phải / cụt ~25px. Kiểm trên trang BI thật: lề trái 52px, lề phải 54px (cân 2 bên). Tab khác giữ cách chụp cũ.
- Realtime V1 (Tổng cụm): bảng "Doanh thu NV hôm nay" xếp cùng thứ tự với "Lũy kế tháng" (Dự kiến cao → thấp) → mỗi NV nằm cùng 1 hàng ở 2 bên. Xuất Excel cũng vậy.

## Thay đổi 18.59 — chép ảnh cụt bên phải (nguyên nhân thật)

- Trang BI ép mọi khung `box-sizing: border-box !important`, nên lề trắng 14px của khung chụp ăn vào bề ngang bảng → ảnh mất ~28px bên phải (cột Dự kiến). Kiểm trên trang BI thật: bản 18.58 cụt 28px, bản 18.59 cụt 0px. Khung chụp nay giữ `content-box` bắt buộc.

## Thay đổi 18.58 — chép ảnh vẫn cụt bên phải

- Ở bản sao dùng để chụp, trình duyệt (nhất là Safari iPhone) có thể tự phóng chữ to hơn lúc đo nên bảng rộng hơn khung → ảnh cụt cột cuối (vd "Dự kiến" ở Báo cáo lũy kế). Nay khóa cỡ chữ trong bản sao và đo lại bề ngang ngay trong bản sao, bảng rộng bao nhiêu thì khung ảnh nới ra bấy nhiêu.

## Thay đổi 18.57 — Realtime V1 (theme Bảng Excel) + chép ảnh

- Số Doanh thu ở phần đầu Realtime V1 và dòng Doanh thu / Trả chậm ở Lũy kế nhỏ lại (số 20px → 16px, nhãn 15px → 13px); ô "Thực hiện QĐ" không còn bị cắt.
- Bảng nhân viên dưới Realtime V1 (Tổng cụm): 2 nửa xếp riêng — bên trái theo Hôm nay, bên phải "Lũy kế tháng" có thêm cột tên NV và xếp theo Dự kiến cao → thấp. Xuất Excel ra cùng thứ tự.
- Chép ảnh / Tải ảnh từng mục không còn bị cắt 2 bên khi khung báo cáo đang thu nhỏ cho vừa màn hình (trước đây đo bề ngang theo khung đã thu nhỏ); "Cả trang đang xem" cũng chụp đủ.

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
