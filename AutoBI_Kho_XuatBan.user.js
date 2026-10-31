// ==UserScript==
// @name         AutoBI - Kho & Xuất Bán
// @namespace    https://github.com/PhamngocNDH/AutoBI/kho-xuatban-test
// @version      3.0.0
// @description  Đã ngừng sử dụng.
// @author       AutoBI / 38967 - Mr Phạm
// @homepageURL  https://github.com/PhamngocNDH/AutoBI
// @updateURL    https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI_Kho_XuatBan.user.js
// @downloadURL  https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI_Kho_XuatBan.user.js
// @match        https://report.mwgroup.vn/*
// @run-at       document-idle
// @noframes
// @grant        GM_listValues
// @grant        GM_deleteValue
// ==/UserScript==

/* 31/10/2026 — Tool Kho & Xuất Bán đã NGỪNG theo quyết định của 38967 - Mr Phạm.
   Bản này thay cho mọi bản cũ (Tampermonkey tự cập nhật): không lấy số, không gọi BI, không gửi gì đi;
   xóa cấu hình/số đã lưu trên máy và hiện 1 dòng thông báo. Có thể gỡ hẳn script trong Tampermonkey. */
(function () {
  'use strict';
  try { (GM_listValues() || []).forEach(k => { try { GM_deleteValue(k); } catch (_) { } }); } catch (_) { }
  try {
    const n = document.createElement('div');
    n.textContent = 'Tool "Kho & Xuất Bán" đã ngừng sử dụng. Vui lòng gỡ script trong Tampermonkey.';
    n.style.cssText = 'position:fixed;left:12px;bottom:12px;z-index:2147483647;max-width:360px;padding:10px 14px;border-radius:10px;background:#1f2328;color:#fff;font:600 13px/1.4 -apple-system,Segoe UI,Roboto,Arial,sans-serif;box-shadow:0 6px 20px rgba(0,0,0,.25);cursor:pointer';
    n.title = 'Bấm để ẩn';
    n.onclick = () => n.remove();
    (document.body || document.documentElement).appendChild(n);
    setTimeout(() => n.remove(), 15000);
  } catch (_) { }
})();
