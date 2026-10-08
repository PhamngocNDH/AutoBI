// ==UserScript==
// @name         AutoBI - Kho & Xuất Bán
// @namespace    https://github.com/PhamngocNDH/AutoBI/kho-xuatban-test
// @version      2.5.1
// @description  Đổ tồn kho (BI 4286) và xuất bán (BI 77) theo cụm siêu thị cho máy tính: lấy thẳng dữ liệu BI có điều tốc, sổ ngày, bộ chọn tồn kho, Excel.
// @author       AutoBI / 38967 - Mr Phạm
// @homepageURL  https://github.com/PhamngocNDH/AutoBI
// @updateURL    https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI_Kho_XuatBan.user.js
// @downloadURL  https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI_Kho_XuatBan.user.js
// @match        https://report.mwgroup.vn/*
// @run-at       document-start
// @noframes
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_listValues
// @grant        GM_deleteValue
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      docs.google.com
// @connect      googleusercontent.com
// @connect      *.googleusercontent.com
// @connect      raw.githubusercontent.com
// @connect      crm.thegioididong.com
// ==/UserScript==

/*
 * NGUỒN DỮ LIỆU (đã kiểm chứng trên BI thật ngày 01/10/2026)
 * - Tồn kho: POST /Home/FilterDataDynamicReport, báo cáo 4286, mỗi lần 1 siêu thị (~3 giây/siêu thị).
 * - Xuất bán: cùng endpoint, báo cáo 77. BI chỉ cho 1 ngày mỗi lần ("VUI LÒNG CHỌN 1 NGÀY").
 *   Bộ lọc: Hình thức xuất = 3, Tìm theo = 2 (Kho tạo), Ngành = 13. Đối chiếu khớp 41/41 dòng với file
 *   Excel xuất tay ngày 29–30/09/2026.
 * - Tham số báo cáo đọc từ /Home/GetDynamicReport nên tự theo khi BI đổi định nghĩa.
 * Quy tắc tính bán: Trạng thái xuất = Đã xuất, Trạng thái giao = Đã giao, Trả hàng = Chưa trả, Hủy = Chưa hủy.
 * Doanh thu = Giá bán × SL (đã gồm VAT). Không dùng Phải thu (tổng cả đơn, lặp ở mọi dòng).
 * Không lưu tên, số điện thoại, địa chỉ, email khách hàng.
 * Giữ nguyên nguyên tắc bản TEST: dữ liệu sai shop / sai ngày / thiếu dòng → dừng, không công bố kết quả thiếu.
 */
(function () {
    'use strict';
    const VERSION = '2.5.1';
    const UPDATE_URL = 'https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI_Kho_XuatBan.user.js';
    const SALES_SCHEMA = 4;                             // 4 = tất cả ngành + Loại hàng + Kho xuất (MASIEUTHIXUAT); ngày lưu bằng bản cũ sẽ được lấy lại
    const PREFIX = 'autobi_kxb_test_v1_';               // giữ khóa cũ để không mất khai báo shop
    const AUTH_SHEET = Object.freeze({ id: '17PxnghjkKIP36fWoSd656wo3DhlOmMiTiyjlf1g23UU', gid: '1237161146' });
    const REPORT = Object.freeze({ inventory: 4286, sales: 77 });
    const SALES_FILTER = Object.freeze({ exportType: '3', warehouseMode: '2', category: '' });   // '' = tất cả ngành hàng
    const RECHECK_DAYS = 60;                            // ngày còn treo: xem lại tối đa 60 ngày

    /* ================= TIỆN ÍCH CHUNG (thuần, test được bằng node) ================= */
    const clean = v => String(v ?? '').replace(/ /g, ' ').trim();
    const norm = v => clean(v).normalize('NFD').replace(/[̀-ͯ]/g, '')
        .replace(/đ/gi, 'd').toLowerCase().replace(/\s+/g, ' ');
    const same = (a, b) => norm(a) === norm(b);
    const fail = (message, code = 'GUARD') => { const e = new Error(message); e.code = code; throw e; };
    const invariant = (ok, message) => { if (!ok) fail(message); };
    const keyCode = value => clean(value).replace(/^0+(?=\d)/, '');
    const hasCode = (label, code) => keyCode(clean(label).match(/^(\d+)\s*(?:[-–—]|$)/)?.[1] || '') === keyCode(code);
    const clone = value => JSON.parse(JSON.stringify(value));
    const pad = n => String(n).padStart(2, '0');
    const uuid = () => (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') ? crypto.randomUUID()
        : 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, ch => { const r = Math.random() * 16 | 0; return (ch === 'x' ? r : (r & 3 | 8)).toString(16); });

    // Ngày: chuẩn nội bộ 'yyyy-mm-dd'; BI dùng 'dd/mm/yyyy'
    function day(value) {
        if (value instanceof Date) return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}`;
        const s = clean(value);
        let m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$/), y, mo, d;
        if (m) [, y, mo, d] = m;
        else { m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s.*)?$/); invariant(m, `Ngày không hợp lệ: ${s}`); [, d, mo, y] = m; }
        const dt = new Date(Number(y), Number(mo) - 1, Number(d));
        invariant(dt.getFullYear() === Number(y) && dt.getMonth() === Number(mo) - 1 && dt.getDate() === Number(d), `Ngày không tồn tại: ${s}`);
        return `${y}-${pad(mo)}-${pad(d)}`;
    }
    const toBI = iso => `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}`;
    const stamp = d => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
    const isoDate = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
    const addDays = (iso, n) => { const [y, m, d] = iso.split('-').map(Number); return isoDate(new Date(y, m - 1, d + n)); };
    function validateRange(from, to) { from = day(from); to = day(to); invariant(from <= to, 'Từ ngày phải trước hoặc bằng Đến ngày'); return { from, to }; }
    function daysIn(range) { const out = []; for (let d = range.from; d <= range.to; d = addDays(d, 1)) out.push(d); return out; }

    // Số từ API BI: luôn dạng en-US "1,234,567.89"
    function apiNumber(value) {
        if (typeof value === 'number') return value;
        const s = clean(value);
        if (s === '') return 0;
        invariant(/^-?(\d{1,3}(,\d{3})*|\d+)(\.\d+)?$/.test(s), `BI trả số lạ: ${s}`);
        return Number(s.replace(/,/g, ''));
    }


    /* ================= ĐỌC CSV (sheet Auth) ================= */
    function parseDelimited(text, delimiter = ',') {
        text = text.replace(/^﻿/, '');
        const rows = []; let row = [], cell = '', quoted = false;
        for (let i = 0; i < text.length; i++) {
            const ch = text[i];
            if (quoted) {
                if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; }
                else if (ch === '"') quoted = false;
                else cell += ch;
            } else if (ch === '"' && cell === '') quoted = true;
            else if (ch === delimiter) { row.push(cell); cell = ''; }
            else if (ch === '\n' || ch === '\r') {
                if (ch === '\r' && text[i + 1] === '\n') i++;
                row.push(cell); rows.push(row); row = []; cell = '';
            } else cell += ch;
        }
        invariant(!quoted, 'CSV bị thiếu dấu đóng ngoặc kép');
        if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
        return rows;
    }

    /* ================= QUY TẮC XUẤT BÁN & TỔNG HỢP ================= */
    // Một dòng bán chuẩn hóa — dùng chung cho dữ liệu API và file nhập. Không có dữ liệu khách hàng.
    // Loại hàng (Mới / Đã sử dụng / Trưng bày…) — BI 77 cột "Trạng thái hồ sơ". Tên cột API chưa cố định nên dò theo tên rồi theo giá trị.
    const CONDITION_KEYS = ['TRANGTHAI', 'TRANGTHAIHOSO', 'TRANGTHAISANPHAM', 'TINHTRANGSANPHAM', 'INVENTORYSTATUSNAME', 'TRANGTHAIHANG', 'LOAIHANG', 'PRODUCTSTATUSNAME'];
    const CONDITION_VALUE = /^\s*\d+\s*-\s*(mới|đã sử dụng|trưng bày|lỗi|cũ|hàng )/i;
    function conditionKey(r) {
        if (!r || typeof r !== 'object') return '';
        const keys = Object.keys(r), up = new Map(keys.map(k => [k.toUpperCase(), k]));
        for (const k of CONDITION_KEYS) if (up.has(k) && clean(r[up.get(k)])) return up.get(k);
        return keys.find(k => !/NGANH|NHOM|HINHTHUC|LOAIYEU|NGUOI|TEN|MA|IMEI/i.test(k) && CONDITION_VALUE.test(clean(r[k]))) || '';
    }
    // Kho xuất hàng (BI 77 cột "Mã siêu thị xuất hàng"): đơn tạo ở shop nhưng xuất từ kho khác thì không trừ tồn shop
    const OUTSTORE_KEYS = ['MASIEUTHIXUAT', 'OUTPUTSTOREID', 'MASIEUTHIXUATHANG', 'MASTXUATHANG', 'STOREIDXUAT', 'XUATTAISIEUTHI', 'OUTPUTSTORE'];
    function outStoreKey(r) {
        if (!r || typeof r !== 'object') return '';
        const keys = Object.keys(r), up = new Map(keys.map(k => [k.toUpperCase(), k]));
        for (const k of OUTSTORE_KEYS) if (up.has(k)) return up.get(k);
        return keys.find(k => /XUAT|OUTPUT/i.test(k) && /STORE|SIEUTHI|KHO/i.test(k) && !/ORIGINATE|TEN|NAME|^SIEUTHIXUAT$/i.test(k) && /^\s*[\d.,]+\s*$/.test(String(r[k] ?? ''))) || '';
    }
    const conditionText = c => clean(c).replace(/^\d+\s*-\s*/, '') || 'Chưa rõ';
    function lineFromApi(r, condKey) {
        return {
            created: day(String(r.NGAYTAO).slice(0, 10)), time: clean(r.NGAYTAO).slice(11, 16),
            shipped: r.NGAYXUATHANG ? day(String(r.NGAYXUATHANG).slice(0, 10)) : '',
            shop: keyCode(r.ORIGINATESTOREID), order: clean(r.MAYEUCAUXUAT), orderType: clean(r.LOAIYEUCAUXUAT),
            exportType: clean(r.HINHTHUCXUAT), creator: clean(r.NGUOITAO), brand: clean(r.NHASANXUAT),
            product: clean(r.MASANPHAM), productName: clean(r.TENSANPHAM), category: clean(r.NGANHHANG), group: clean(r.NHOMHANG),
            imei: /^x+$/i.test(clean(r.IMEI)) ? '' : clean(r.IMEI),
            qty: apiNumber(r.SOLUONG), price: apiNumber(r.GIABAN), priceNet: apiNumber(r.SALEPRICE),
            exported: clean(r.TRANGTHAIXUAT), delivered: clean(r.TRANGTHAIGIAO), cancelled: clean(r.TRANGTHAIHUY), returned: clean(r.TRAHANG),
            condition: clean(r[condKey === undefined ? conditionKey(r) : condKey] ?? ''),
            outShop: (k => k && clean(r[k]) ? keyCode(String(r[k]).replace(/\D/g, '')) : '')(outStoreKey(r)),   // BI trả dạng "2,928"
            pay: clean(r.HINHTHUCTHANHTOAN), ship: clean(r.HINHTHUCGIAOHANG)   // V2.0: hình thức thanh toán / giao hàng (không phải thông tin khách)
        };
    }
    const RULE = [['exported', 'Đã xuất', 'Chưa xuất'], ['delivered', 'Đã giao', 'Chưa giao'], ['cancelled', 'Chưa hủy', 'Đã hủy'], ['returned', 'Chưa trả', 'Đã trả hàng']];
    function reasons(line) {
        const out = RULE.filter(([k, ok]) => !same(line[k], ok)).map(([, , label]) => label);
        if (SALES_FILTER.category && !hasCode(line.category, SALES_FILTER.category)) out.push('Ngành khác ' + SALES_FILTER.category);
        if (!/xuat ban hang tai sieu thi$/.test(norm(line.exportType))) out.push('Hình thức xuất khác 3');
        return out;
    }
    // Đơn còn treo = chưa hủy nhưng chưa xuất hoặc chưa giao → ngày tạo phải lấy lại sau
    const pending = line => same(line.cancelled, 'Chưa hủy') && !(same(line.exported, 'Đã xuất') && same(line.delivered, 'Đã giao'));
    function validateSalesLines(lines, shops, dayOrRange) {
        const codes = new Set(shops.map(s => keyCode(s.code)));
        const from = typeof dayOrRange === 'string' ? dayOrRange : dayOrRange.from, to = typeof dayOrRange === 'string' ? dayOrRange : dayOrRange.to;
        lines.forEach((l, i) => {
            invariant(codes.has(l.shop), `Dòng ${i + 1}: Mã kho tạo ${l.shop} nằm ngoài cấu hình; đã dừng`);
            invariant(l.created >= from && l.created <= to, `Dòng ${i + 1}: Ngày tạo ${l.created} ngoài kỳ đã chọn; đã dừng`);
            invariant(Number.isFinite(l.qty) && l.qty >= 0 && Number.isFinite(l.price), `Dòng ${i + 1}: SL/Giá không hợp lệ`);
        });
    }
    // Tách dòng bán / dòng loại. Chỉ dòng bán được lưu; dòng loại chỉ đếm lý do.
    // Đơn khách nhập trả: chỉ cần 1 dòng "Đã trả" là bỏ qua CẢ ĐƠN
    const isReturned = l => same(l.returned, 'Đã trả') && !same(l.cancelled, 'Đã hủy');
    function splitSales(lines) {
        const valid = [], excluded = {}, pendingOrders = new Set(), pendingLines = [], returnedLines = [];
        const returnedOrders = new Set(lines.filter(isReturned).map(l => l.order));
        const brief = l => ({ created: l.created, time: l.time, shipped: l.shipped, shop: l.shop, order: l.order, orderType: l.orderType, creator: l.creator,
            brand: l.brand, product: l.product, productName: l.productName, imei: l.imei, condition: l.condition, category: l.category, group: l.group, qty: l.qty, price: l.price, exported: l.exported, delivered: l.delivered, returned: l.returned });
        for (const l of lines) {
            if (returnedOrders.has(l.order)) {
                excluded['Đơn khách nhập trả'] = (excluded['Đơn khách nhập trả'] || 0) + 1;
                if ((!SALES_FILTER.category || hasCode(l.category, SALES_FILTER.category)) && l.price > 0) returnedLines.push(brief(l));
                continue;
            }
            if (pending(l)) { pendingOrders.add(l.order); pendingLines.push(brief(l)); }
            const why = reasons(l);
            if (why.length) { why.forEach(w => excluded[w] = (excluded[w] || 0) + 1); continue; }
            invariant(l.creator && l.brand && l.product, `Dòng bán ${l.order}: thiếu Người tạo/Nhà sản xuất/Mã sản phẩm`);
            valid.push(l);
        }
        return { valid, excluded, pending: pendingOrders.size, pendingLines, returned: returnedOrders.size, returnedLines };
    }
    function shares(map, totalQ, totalR) {
        return [...map.values()].map(x => ({ ...x, pctQty: totalQ > 0 ? x.quantity / totalQ * 100 : 0, pctRev: totalR > 0 ? x.revenue / totalR * 100 : 0 }))
            .sort((a, b) => b.quantity - a.quantity || b.revenue - a.revenue);
    }
    function add(map, key, base, line) {
        const e = map.get(key) || { ...base, quantity: 0, revenue: 0 };
        e.quantity += line.qty; e.revenue += line.qty * line.price; map.set(key, e); return e;
    }
    // Tỷ trọng theo SL (như bản TEST) và theo doanh thu
    function summarizeSales(lines, basis = 'created') {
        const brands = new Map(), brandShop = new Map(), shops = new Map(), staff = new Map(), products = new Map(), days = new Map();
        let quantity = 0, revenue = 0;
        for (const l of lines) {
            quantity += l.qty; revenue += l.qty * l.price;
            add(brands, norm(l.brand), { label: l.brand }, l);
            add(brandShop, norm(l.brand) + '|' + l.shop, { brand: l.brand, shop: l.shop }, l);
            add(shops, l.shop, { code: l.shop }, l);
            add(products, l.product, { product: l.product, name: l.productName, brand: l.brand, group: l.group }, l);
            add(days, basis === 'shipped' && l.shipped ? l.shipped : l.created, { day: basis === 'shipped' && l.shipped ? l.shipped : l.created }, l);
            const employee = l.creator.match(/^(\d+)\s*[-–]/)?.[1] || norm(l.creator);
            const key = JSON.stringify([l.shop, employee]);
            const s = staff.get(key) || { shop: l.shop, employee, label: l.creator, quantity: 0, revenue: 0, brands: new Map(), products: new Map() };
            s.quantity += l.qty; s.revenue += l.qty * l.price;
            add(s.brands, norm(l.brand), { label: l.brand }, l);
            add(s.products, l.product, { label: `${l.productName}`, product: l.product, brand: l.brand }, l);
            staff.set(key, s);
        }
        const shopList = [...shops.values()];
        return {
            quantity, revenue, lineCount: lines.length,
            brands: shares(brands, quantity, revenue), brandShop: Object.fromEntries(brandShop),
            shops: shopList, products: shares(products, quantity, revenue),
            days: [...days.values()].sort((a, b) => a.day.localeCompare(b.day)),
            staff: [...staff.values()].map(s => {
                const sh = shops.get(s.shop);
                return { ...s, pctShop: sh.quantity > 0 ? s.quantity / sh.quantity * 100 : 0,
                    brands: shares(s.brands, s.quantity, s.revenue), products: shares(s.products, s.quantity, s.revenue) };
            }).sort((a, b) => b.quantity - a.quantity || b.revenue - a.revenue)
        };
    }
    function inPeriod(lines, range, basis) {
        return lines.filter(l => { const d = basis === 'shipped' && l.shipped ? l.shipped : l.created; return d >= range.from && d <= range.to; });
    }


    /* ================= ĐIỀU TỐC (BI: tối đa 5 lần/60 giây cho mỗi báo cáo) ================= */
    // Đọc thông báo "Bạn thao tác quá nhanh ... (tối đa 5 lần/60 giây). Vui lòng thử lại sau 5 giây !"
    function parseRateLimit(message) {
        const m = norm(message);
        if (!/thao tac qua nhanh/.test(m)) return null;
        const lim = m.match(/toi da\s*(\d+)\s*lan\s*\/\s*(\d+)\s*giay/), wait = m.match(/sau\s*(\d+)\s*giay/);
        return { max: lim ? Number(lim[1]) : 5, windowSec: lim ? Number(lim[2]) : 60, waitSec: wait ? Number(wait[1]) : 10 };
    }
    // Số mili-giây cần chờ để lần gọi tiếp theo vẫn trong giới hạn (cửa sổ trượt)
    function waitBeforeCall(times, now, max, windowMs) {
        const recent = times.filter(t => now - t < windowMs).sort((a, b) => a - b);
        return recent.length < max ? 0 : recent[recent.length - max] + windowMs - now;
    }

    /* ================= SỔ NGÀY ================= */
    // book = { 'yyyy-mm-dd': { shops, status, pending, rows, valid, quantity, revenue, at } }
    // Ngày chỉ "Đã chốt" khi không còn đơn treo VÀ đã qua số ngày kiểm tra lại nhập trả (khách có thể trả hàng vài ngày sau)
    function dayStatus(iso, pendingCount, today, returnDays = 7) {
        if (pendingCount > 0) return 'Còn treo';
        return iso <= addDays(today, -Math.max(2, returnDays)) ? 'Đã chốt' : 'Chưa chốt';
    }
    // Ngày vừa lấy chưa quá freshHours giờ (cùng schema, cùng bộ siêu thị) thì không lấy lại — trừ hôm nay (số còn chạy)
    function isFresh(r, d, today, shopsKey, schema, nowMs, freshHours) {
        return !!(freshHours > 0 && nowMs && r && r.atMs && d !== today && r.shops === shopsKey && (!schema || r.schema === schema)
            && nowMs - r.atMs >= 0 && nowMs - r.atMs < freshHours * 3600 * 1000);
    }
    function daysToFetch(book, range, shopsKey, today, refetchAll, returnDays = 7, schema = 0, nowMs = 0, freshHours = 0) {
        const need = new Set();
        const end = range.to < today ? range.to : today;
        const recheckFrom = addDays(today, -Math.max(2, returnDays) + 1);   // các ngày gần đây: lấy lại để bắt đơn khách nhập trả
        for (let d = range.from; d <= end; d = addDays(d, 1)) {
            const r = book[d];
            if (refetchAll || !r || (schema && r.schema !== schema) || r.shops !== shopsKey) { need.add(d); continue; }
            if ((r.status !== 'Đã chốt' || d >= recheckFrom) && !isFresh(r, d, today, shopsKey, schema, nowMs, freshHours)) need.add(d);
        }
        if (!refetchAll) for (const [d, r] of Object.entries(book))
            if (r.status === 'Còn treo' && d >= addDays(today, -RECHECK_DAYS) && d <= today && !isFresh(r, d, today, shopsKey, schema, nowMs, freshHours)) need.add(d);
        return [...need].sort();
    }

    /* ================= TỒN KHO ================= */
    function inventoryRecordFromApi(r) {
        return { shop: keyCode(r.storeid), shopName: clean(r.storename), category: clean(r.maingroupname), group: clean(r.subgroupname),
            brand: clean(r.productbrandname), product: clean(r.productid), productName: clean(r.productname), serial: clean(r.imei),
            condition: clean(r.inventorystatusname), qty: apiNumber(r.quantity), cost: apiNumber(r.totalcost), input: clean(r.inputdate) };
    }
    function summarizeInventory(records, shop) {
        const categories = new Set(), groups = new Map(), products = new Set(), serials = new Set(), hierarchy = new Map();
        let quantity = 0, cost = 0;
        records.forEach((v, i) => {
            invariant(keyCode(v.shop) === keyCode(shop.code), `Dòng ${i + 1}: sai mã siêu thị (${v.shop})`);
            invariant(v.product, `Dòng ${i + 1}: thiếu Mã sản phẩm`);
            invariant(Number.isFinite(v.qty), 'Số lượng tồn không hợp lệ');
            quantity += v.qty; cost += v.cost || 0;
            categories.add(v.category); products.add(v.product); if (v.serial) serials.add(v.serial);
            const gk = JSON.stringify([v.category, v.group]);
            const g = groups.get(gk) || { category: v.category, group: v.group, quantity: 0, fresh: 0, cost: 0 };
            g.quantity += v.qty; g.cost += v.cost || 0; if (/^1\b/.test(v.condition)) g.fresh += v.qty; groups.set(gk, g);
            const key = JSON.stringify([v.category, v.group, v.product, v.condition]);
            const x = hierarchy.get(key) || { category: v.category, group: v.group, brand: v.brand, product: v.product, name: v.productName, condition: v.condition, quantity: 0, cost: 0, rows: 0 };
            x.quantity += v.qty; x.cost += v.cost || 0; x.rows++; hierarchy.set(key, x);
        });
        const r2 = v => Math.round(v * 100) / 100;
        return { rows: records.length, quantity: r2(quantity), cost: Math.round(cost), categories: categories.size, groups: groups.size, products: products.size, serials: serials.size,
            byGroup: [...groups.values()].map(g => ({ ...g, quantity: r2(g.quantity), fresh: r2(g.fresh), cost: Math.round(g.cost) })).sort((a, b) => b.cost - a.cost),
            hierarchy: [...hierarchy.values()].map(h => ({ ...h, quantity: r2(h.quantity), cost: Math.round(h.cost) })) };
    }

    // Bộ chọn tồn kho: siêu thị, ngành, nhóm hàng, hãng, trạng thái, tìm IMEI / mã / tên
    function filterInventory(records, f) {
        const shops = new Set(f.shops || []), q = norm(f.q || '');
        return records.filter(r => (!shops.size || shops.has(r.shop)) && (!f.category || r.category === f.category) && (!f.group || r.group === f.group)
            && (!f.brand || r.brand === f.brand) && (!f.conditions?.length || f.conditions.includes(r.condition))
            && (!q || norm(r.serial).includes(q) || norm(r.product).includes(q) || norm(r.productName).includes(q)));
    }
    function inventoryOptions(base, f) {
        const uniq = a => [...new Set(a.filter(Boolean))].sort((x, y) => x.localeCompare(y, 'vi', { numeric: true }));
        const byCat = base.filter(r => !f.category || r.category === f.category);
        const byGroup = byCat.filter(r => !f.group || r.group === f.group);
        return { categories: uniq(base.map(r => r.category)), groups: uniq(byCat.map(r => r.group)), brands: uniq(byGroup.map(r => r.brand)), conditions: uniq(base.map(r => r.condition)) };
    }
    function inventoryViews(recs) {
        const r2 = v => Math.round(v * 100) / 100;
        const groups = new Map(), products = new Map(), serials = new Set(), conds = new Set();
        let quantity = 0, cost = 0;
        for (const r of recs) {
            quantity += r.qty; cost += r.cost || 0; if (r.serial) serials.add(r.serial); conds.add(r.condition);
            for (const [map, key, base] of [[groups, r.category + '|' + r.group, { category: r.category, group: r.group }], [products, r.product, { product: r.product, name: r.productName, brand: r.brand, group: r.group }]]) {
                const e = map.get(key) || { ...base, quantity: 0, cost: 0, byShop: {}, byCond: {} };
                e.quantity += r.qty; e.cost += r.cost || 0; e.byShop[r.shop] = (e.byShop[r.shop] || 0) + r.qty; e.byCond[r.condition] = (e.byCond[r.condition] || 0) + r.qty;
                map.set(key, e);
            }
        }
        const fix = e => ({ ...e, quantity: r2(e.quantity), byShop: Object.fromEntries(Object.entries(e.byShop).map(([k, v]) => [k, r2(v)])), byCond: Object.fromEntries(Object.entries(e.byCond).map(([k, v]) => [k, r2(v)])) });
        return { quantity: r2(quantity), cost, serials: serials.size, conditionList: [...conds].sort((a, b) => a.localeCompare(b, 'vi', { numeric: true })),
            groups: [...groups.values()].map(fix).sort((a, b) => a.category.localeCompare(b.category, 'vi', { numeric: true }) || b.quantity - a.quantity),
            products: [...products.values()].map(fix).sort((a, b) => b.quantity - a.quantity || a.name.localeCompare(b.name)) };
    }

    /* ================= SHOP & AUTH ================= */
    function validateShops(shops) {
        invariant(Array.isArray(shops) && shops.length, 'Hãy khai báo ít nhất một siêu thị');
        const seen = new Set();
        for (const shop of shops) {
            invariant(/^\d{1,10}$/.test(clean(shop.code)) && clean(shop.name), 'Mỗi shop cần mã số và tên hiển thị');
            invariant(!seen.has(keyCode(shop.code)), `Trùng mã shop: ${shop.code}`); seen.add(keyCode(shop.code));
        }
        return shops;
    }
    function authorizeSheetRows(rows, user) {
        invariant(/^\d{4,10}$/.test(user), 'Mã nhân viên không hợp lệ');
        const headers = rows.map((row, i) =>
            same(row[0], 'UserID') && same(row[2], 'Trạng thái') && same(row[1], 'Họ tên') ? i : -1).filter(i => i >= 0);
        invariant(headers.length === 1, 'CSV Auth thiếu hoặc trùng dòng tiêu đề UserID/Họ tên/Trạng thái');
        const matches = rows.slice(headers[0] + 1).filter(row => /^\d{4,10}$/.test(clean(row[0])) && keyCode(row[0]) === keyCode(user));
        invariant(matches.length <= 1, `Sheet Auth có nhiều dòng cho mã ${user}; đã dừng`);
        invariant(matches.length === 1, `Mã ${user} không có trong sheet Auth`);
        invariant(same(matches[0][2], 'ACTIVE'), `Mã ${user} chưa ở trạng thái ACTIVE`);
        return { user, name: clean(matches[0][1]), checkedAt: Date.now(), source: 'Auth!A:D', status: 'ACTIVE' };
    }


    /* ---------- Phiếu kiểm tồn kho (in A4 đứng) ---------- */
    function escHtml(v) { return String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
    // Gom theo Siêu thị → Nhóm hàng, sắp theo Hãng · Tên SP · IMEI để nhân viên đi theo kệ cho dễ
    function inventoryChecklist(records, shopOrder) {
        const cmp = (a, b) => String(a || '').localeCompare(String(b || ''), 'vi', { numeric: true });
        const order = new Map((shopOrder || []).map((c, i) => [keyCode(c), i]));
        const shops = new Map();
        for (const r of records) {
            const s = keyCode(r.shop);
            if (!shops.has(s)) shops.set(s, new Map());
            const g = shops.get(s), k = r.group || '(Không rõ nhóm)';
            if (!g.has(k)) g.set(k, []);
            g.get(k).push(r);
        }
        return [...shops.entries()].sort((a, b) => (order.get(a[0]) ?? 999) - (order.get(b[0]) ?? 999) || cmp(a[0], b[0])).map(([shop, groups]) => {
            const gs = [...groups.entries()].sort((a, b) => cmp(a[0], b[0])).map(([group, rows]) => {
                rows = rows.slice().sort((a, b) => cmp(a.brand, b.brand) || cmp(a.productName, b.productName) || cmp(a.product, b.product) || cmp(a.serial, b.serial));
                return { group, rows, lines: rows.length, quantity: rows.reduce((t, r) => t + (Number(r.qty) || 0), 0) };
            });
            return { shop, groups: gs, lines: gs.reduce((t, g) => t + g.lines, 0), quantity: gs.reduce((t, g) => t + g.quantity, 0) };
        });
    }
    function inventoryPrintHtml(records, meta) {
        const m = meta || {}, nameOf = m.nameOf || (c => c), num = v => new Intl.NumberFormat('vi-VN').format(v);
        const list = inventoryChecklist(records, m.shopOrder);
        invariant(list.length, 'Không có dòng tồn nào để in (kiểm tra lại siêu thị / bộ lọc)');
        const box = '<span class="box"></span>';
        const shopHtml = list.map((s, si) => {
            let stt = 0;
            const body = s.groups.map(g => `<tbody class="grp"><tr class="gh"><td colspan="7">${escHtml(g.group)} <span>— ${num(g.lines)} dòng · SL ${num(g.quantity)}</span></td></tr>`
                + g.rows.map(r => `<tr><td class="c">${++stt}</td><td class="code">${escHtml(r.product)}</td><td>${escHtml(r.productName)}</td><td class="code imei">${escHtml(r.serial) || '<i>—</i>'}</td><td class="st">${escHtml(r.condition)}</td><td class="c b">${num(Number(r.qty) || 0)}</td><td class="c">${box}</td></tr>`).join('')
                + '</tbody>').join('');
            return `<section class="${si ? 'brk' : ''}">
<div class="head"><div><h1>PHIẾU KIỂM TỒN KHO</h1><div class="shop">${escHtml(s.shop)} · ${escHtml(nameOf(s.shop))}</div></div>
<div class="meta"><div>Tồn lúc: <b>${escHtml((m.times || {})[s.shop] || m.capturedAt || '')}</b></div><div>Lọc: ${escHtml(m.filter || 'Không lọc')}</div><div>Tổng: <b>${num(s.lines)}</b> dòng · SL <b>${num(s.quantity)}</b> · ${num(s.groups.length)} nhóm hàng</div></div></div>
<table><colgroup><col style="width:10mm"><col style="width:27mm"><col><col style="width:37mm"><col style="width:20mm"><col style="width:10mm"><col style="width:13mm"></colgroup>
<thead><tr><th>STT</th><th>Mã SP</th><th>Tên sản phẩm</th><th>IMEI / Serial</th><th>Trạng thái</th><th>SL</th><th>KIỂM</th></tr></thead>${body}</table>
<div class="sign"><div>Người kiểm<br><span>(ký, ghi rõ họ tên)</span></div><div>Ngày kiểm: ....../....../........<br>Số dòng lệch: ............</div><div>Quản lý siêu thị<br><span>(ký, ghi rõ họ tên)</span></div></div>
<div class="note">Ghi chú chênh lệch: ..........................................................................................................................................................................................<br>................................................................................................................................................................................................................................</div>
</section>`;
        }).join('');
        return `<!doctype html><html lang="vi"><head><meta charset="utf-8"><title>Phiếu kiểm tồn kho</title><style>
@page{size:A4 portrait;margin:10mm 9mm 12mm 9mm;@bottom-right{content:"Trang " counter(page) "/" counter(pages);font:8pt Arial,sans-serif;color:#555}@bottom-left{content:${JSON.stringify('AutoBI V' + (m.version || '') + ' · In lúc ' + (m.printedAt || ''))};font:8pt Arial,sans-serif;color:#555}}
*{box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact}
body{margin:0;font:9.5pt/1.3 Arial,"Segoe UI",sans-serif;color:#000}
.head{display:flex;justify-content:space-between;gap:8mm;align-items:flex-end;border-bottom:1.5pt solid #000;padding-bottom:2mm;margin-bottom:2mm}
h1{font-size:15pt;margin:0;letter-spacing:.5pt}.shop{font-size:12pt;font-weight:bold;margin-top:1mm}
.meta{text-align:right;font-size:8.5pt;line-height:1.45}
table{width:100%;border-collapse:collapse;table-layout:fixed}
th,td{border:.6pt solid #444;padding:1mm 1.2mm;vertical-align:middle;overflow-wrap:anywhere;word-break:break-word}
thead{display:table-header-group}th{background:#e8e8e8;font-size:8.5pt;text-align:center}
tr{break-inside:avoid;page-break-inside:avoid}
.gh td{background:#f1f1f1;font-weight:bold;font-size:9.5pt;break-after:avoid;page-break-after:avoid}.gh span{font-weight:normal;font-size:8.5pt}
.c{text-align:center;white-space:nowrap;overflow-wrap:normal;word-break:normal}.b{font-weight:bold}.code{font-family:Consolas,"Courier New",monospace;font-size:8pt;word-break:break-all}.imei{font-size:9pt}
.st{font-size:8.5pt}.sub{font-size:7.5pt;color:#555}
.box{display:inline-block;width:5mm;height:5mm;border:1pt solid #000;border-radius:.8mm;vertical-align:middle}
.sign{display:flex;justify-content:space-between;margin-top:6mm;text-align:center;font-weight:bold;break-inside:avoid}.sign>div{width:32%;min-height:24mm}.sign span{font-weight:normal;font-style:italic;font-size:8pt}
.note{font-size:8.5pt;margin-top:2mm;line-height:2;break-inside:avoid}
.brk{break-before:page;page-break-before:always}
@media screen{body{background:#888;padding:10px}section{background:#fff;width:210mm;margin:0 auto 10px;padding:10mm 9mm}}
</style></head><body>${shopHtml}</body></html>`;
    }

    /* ================= CÂN HÀNG ================= */
    // Tồn (hàng Mới, không gồm đang chuyển kho) so với tốc độ bán N ngày → số ngày đủ bán, SL nên xin bổ sung.
    // Tồn cần có = TB bán/ngày × số ngày giữ đủ (làm tròn lên) + 1 máy dự phòng. Thiếu → Sắp hết, xin phần thiếu.
    // Dòng bán xuất từ kho khác (outShop khác shop tạo) không trừ tồn shop nên không tính.
    function balanceRows(stock, sales, o) {
        const days = Math.max(1, o.days || 10), target = o.target || 14, spare = o.spare ?? 1;
        // o.byModel: gộp các mã khác màu cùng model (theo tên đã bỏ màu)
        const m = new Map(), idOf = (p, name) => o.byModel ? modelOf(name) || clean(p) : clean(p), key = (shop, p, name) => keyCode(shop) + '|' + idOf(p, name);
        const row = (shop, p, name, brand) => {
            const k = key(shop, p, name);
            if (!m.has(k)) m.set(k, { shop: keyCode(shop), product: o.byModel ? idOf(p, name) : clean(p), name: o.byModel ? idOf(p, name) : clean(name), brand: clean(brand), stock: 0, sold: 0, transit: 0, codes: new Set() });
            const x = m.get(k); if (!x.name && name) x.name = clean(name); if (!x.brand && brand) x.brand = clean(brand); return x;
        };
        for (const r of stock) { const x = row(r.shop, r.product, r.productName, r.brand); x.stock += Number(r.qty) || 0; x.codes.add(clean(r.product)); }
        for (const r of (o.transit || [])) row(r.shop, r.product, r.productName, r.brand).transit += Number(r.qty) || 0;
        for (const l of sales) if (!l.outShop || keyCode(l.outShop) === keyCode(l.shop)) { const x = row(l.shop, l.product, l.productName, l.brand); x.sold += Number(l.qty) || 0; x.codes.add(clean(l.product)); }
        return [...m.values()].map(x => {
            const perDay = x.sold / days, cover = perDay > 0 ? x.stock / perDay : (x.stock > 0 ? Infinity : 0);
            const want = perDay > 0 ? Math.ceil(perDay * target - 1e-9) + spare : 0;
            // V2.1: hàng đang chuyển kho nằm ở siêu thị NHẬN (đã kiểm trên BI 08/10/2026) → coi là hàng đang về, trừ vào số thiếu
            const need = Math.max(0, want - x.stock - (o.useIncoming ? x.transit : 0));
            const status = x.stock <= 0 && x.sold > 0 ? 'Hết hàng' : x.sold <= 0 ? (x.stock > 0 ? 'Không bán' : '') : need > 0 ? 'Sắp hết' : cover > target * 2 ? 'Tồn nhiều' : 'Đủ';
            const { codes, ...rest } = x;
            return { ...rest, codeCount: codes.size, codeList: [...codes], want, perDay: Math.round(perDay * 100) / 100, cover: Number.isFinite(cover) ? Math.round(cover * 10) / 10 : null, need, status };
        }).filter(x => x.status).sort((a, b) => STATUS_ORDER.indexOf(a.status) - STATUS_ORDER.indexOf(b.status) || (a.cover ?? 1e9) - (b.cover ?? 1e9) || b.sold - a.sold);
    }
    const STATUS_ORDER = ['Hết hàng', 'Sắp hết', 'Đủ', 'Tồn nhiều', 'Không bán'];

    // So sánh phiên bản dạng 1.5.0 > 1.4.10
    function newerVersion(a, b) {
        const x = String(a || '').split('.').map(Number), y = String(b || '').split('.').map(Number);
        for (let i = 0; i < Math.max(x.length, y.length); i++) { const d = (x[i] || 0) - (y[i] || 0); if (d) return d > 0; }
        return false;
    }
    function parseRemoteScript(text) {
        const v = /@version\s+([\d.]+)/.exec(text || '');
        return v ? { version: v[1] } : null;
    }

    /* ================= SO CÙNG KỲ · BÁN KÈM · ĐƠN TREO · DỰ KIẾN ================= */
    const lastDayOf = (y, m) => new Date(y, m, 0).getDate();            // m: 1–12
    // Cùng kỳ tháng trước: 01–07/10 → 01–07/09; 31/10 → 30/09
    function shiftMonth(iso, k) {
        const [y, m, d] = iso.split('-').map(Number), t = new Date(y, m - 1 + k, 1);
        return isoDate(new Date(t.getFullYear(), t.getMonth(), Math.min(d, lastDayOf(t.getFullYear(), t.getMonth() + 1))));
    }
    const prevMonthRange = range => ({ from: shiftMonth(range.from, -1), to: shiftMonth(range.to, -1) });
    // Bán kèm: đơn có dòng điện thoại; "có kèm" = cùng mã đơn có dòng thuộc nhóm tính kèm. Nhân viên = người tạo dòng điện thoại.
    function attachStats(lines, isPhone, isAttach) {
        const orders = new Map();
        for (const l of lines) {
            let o = orders.get(l.order);
            if (!o) orders.set(l.order, o = { shop: l.shop, creator: '', phoneQty: 0, phoneRev: 0, attachQty: 0, attachRev: 0, groups: new Map() });
            if (isPhone(l)) { o.phoneQty += l.qty; o.phoneRev += l.qty * l.price; if (!o.creator) o.creator = l.creator; }
            else if (isAttach(l)) { o.attachQty += l.qty; o.attachRev += l.qty * l.price; const g = o.groups.get(l.group) || { qty: 0, rev: 0 }; g.qty += l.qty; g.rev += l.qty * l.price; o.groups.set(l.group, g); }
        }
        const staff = new Map(), groups = new Map(), total = { phoneOrders: 0, attached: 0, phoneQty: 0, attachQty: 0, attachRev: 0 };
        for (const o of orders.values()) {
            if (!o.phoneQty) continue;
            const key = o.shop + '|' + o.creator;
            const s = staff.get(key) || { shop: o.shop, label: o.creator, phoneOrders: 0, attached: 0, phoneQty: 0, phoneRev: 0, attachQty: 0, attachRev: 0 };
            const has = o.attachQty > 0 ? 1 : 0;
            for (const x of [s, total]) { x.phoneOrders++; x.attached += has; x.phoneQty += o.phoneQty; x.attachQty += o.attachQty; x.attachRev += o.attachRev; }
            s.phoneRev += o.phoneRev; staff.set(key, s);
            for (const [g, v] of o.groups) { const e = groups.get(g) || { group: g, orders: 0, qty: 0, rev: 0 }; e.orders++; e.qty += v.qty; e.rev += v.rev; groups.set(g, e); }
        }
        const rate = x => ({ ...x, rate: x.phoneOrders ? x.attached / x.phoneOrders * 100 : 0 });
        return { total: rate(total), staff: [...staff.values()].map(rate).sort((a, b) => b.phoneOrders - a.phoneOrders || b.rate - a.rate),
            groups: [...groups.values()].sort((a, b) => b.orders - a.orders || b.rev - a.rev) };
    }
    const daysBetween = (a, b) => { const [y1, m1, d1] = a.split('-').map(Number), [y2, m2, d2] = b.split('-').map(Number); return Math.round((Date.UTC(y2, m2 - 1, d2) - Date.UTC(y1, m1 - 1, d1)) / 864e5); };
    // Đơn treo gom theo mã đơn, tính số ngày treo (từ ngày tạo đến hôm nay), cũ nhất lên đầu
    function pendingOrders(lines, today) {
        const m = new Map();
        for (const l of lines) {
            const o = m.get(l.order) || { order: l.order, created: l.created, time: l.time, shop: l.shop, creator: l.creator, items: [], qty: 0, value: 0, notExported: false };
            if (l.created < o.created) o.created = l.created;
            o.items.push(l.productName); o.qty += l.qty; o.value += l.qty * l.price;
            if (!same(l.exported, 'Đã xuất')) o.notExported = true;
            m.set(l.order, o);
        }
        return [...m.values()].map(o => ({ ...o, age: Math.max(0, daysBetween(o.created, today)), state: o.notExported ? 'Chưa xuất' : 'Chưa giao' }))
            .sort((a, b) => b.age - a.age || a.shop.localeCompare(b.shop) || a.order.localeCompare(b.order));
    }
    // Dự kiến cuối tháng = DT các ngày đã trọn (1 → hôm qua) ÷ số ngày đó × số ngày của tháng
    function projectMonth(revToYesterday, range, today) {
        const first = range.from.slice(0, 8) + '01', y = addDays(today, -1);
        if (range.from !== first || range.to.slice(0, 7) !== first.slice(0, 7) || range.to < y || y.slice(0, 7) !== first.slice(0, 7)) return null;
        const elapsed = daysBetween(first, y) + 1, [yy, mm] = first.split('-').map(Number), dim = lastDayOf(yy, mm);
        return { elapsed, dim, until: y, value: revToYesterday / elapsed * dim };
    }


    /* ================= V2.0: MODEL · CHUYỂN NỘI CỤM · TUỔI TỒN · GIỜ BÁN · ĐANG CHUYỂN ================= */
    // Tên model = tên sản phẩm bỏ màu ở cuối ("iPhone 18 Pro Max 512GB Silver" → "iPhone 18 Pro Max 512GB")
    const COLOR_WORDS = new Set(('den trang xanh do vang tim hong xam bac nau cam be ghi kem reu ngoc duong la mint tim than titan sa mac thien nhien '
        + 'black white blue red green yellow purple pink gray grey silver gold rose orange natural desert titanium midnight starlight graphite cream lavender teal navy sky sage coral '
        + 'dam nhat nhạt dương lá ngọc nhat duong ngoc la than mac thien nhien den nham sang toi pastel').split(/\s+/));
    function modelOf(name) {
        const words = clean(name).replace(/^điện thoại\s+/i, '').split(/\s+/);
        while (words.length > 2 && COLOR_WORDS.has(norm(words[words.length - 1]))) words.pop();
        return words.join(' ');
    }
    // Gợi ý chuyển nội cụm: cùng mã, siêu thị thừa (tồn vượt mức cần giữ) → siêu thị thiếu, trước khi xin kho
    // Siêu thị cho: giữ lại đủ mức "cần có" (TB/ngày × số ngày giữ đủ + dự phòng); mã không bán ở shop đó thì giữ 1 máy trưng bày.
    function transferPlan(rows) {
        const byProd = new Map();
        for (const r of rows) { const a = byProd.get(r.product) || []; a.push(r); byProd.set(r.product, a); }
        const out = [];
        for (const [product, list] of byProd) {
            if (list.length < 2) continue;
            const takers = list.filter(r => r.need > 0).map(r => ({ r, left: r.need })).sort((a, b) => (a.r.cover ?? 0) - (b.r.cover ?? 0) || b.left - a.left);
            const givers = list.map(r => ({ r, left: Math.max(0, r.stock - (r.perDay > 0 ? r.want : 1)) })).filter(g => g.left > 0).sort((a, b) => b.left - a.left);
            for (const t of takers) for (const g of givers) {
                if (!t.left) break;
                if (!g.left || g.r.shop === t.r.shop) continue;
                const q = Math.min(t.left, g.left); t.left -= q; g.left -= q;
                out.push({ product, name: t.r.name, brand: t.r.brand, from: g.r.shop, to: t.r.shop, qty: q, fromStock: g.r.stock, fromCover: g.r.cover, toStock: t.r.stock, toCover: t.r.cover });
            }
        }
        return out.sort((a, b) => b.qty - a.qty || a.name.localeCompare(b.name, 'vi'));
    }
    // Tuổi tồn: số ngày từ Ngày nhập ("26/05/2026 12:24") đến hôm nay
    const AGE_BUCKETS = [[0, 30, 'Dưới 30 ngày'], [30, 60, '30–60 ngày'], [60, 90, '60–90 ngày'], [90, 180, '90–180 ngày'], [180, 1e9, 'Trên 180 ngày']];
    function ageDays(input, today) {
        const m = clean(input).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (!m) return null;
        return Math.max(0, daysBetween(`${m[3]}-${pad(m[2])}-${pad(m[1])}`, today));
    }
    const ageBucket = d => d == null ? 'Không rõ ngày nhập' : AGE_BUCKETS.find(([a, b]) => d >= a && d < b)[2];
    // Hàng đang chuyển kho = dòng có trong lần đổ "tính hàng đang chuyển" mà không có trong lần đổ thường (so theo mã · IMEI · trạng thái · SL)
    function transitDiff(withTransit, without) {
        const k = r => [keyCode(r.shop), clean(r.product), clean(r.serial), clean(r.condition)].join('|');
        const have = new Map(); for (const r of without) have.set(k(r), (have.get(k(r)) || 0) + (Number(r.qty) || 0));
        const out = [];
        for (const r of withTransit) {
            const q = Number(r.qty) || 0, h = have.get(k(r)) || 0;
            if (h >= q) { have.set(k(r), h - q); continue; }
            have.set(k(r), 0); out.push({ ...r, qty: q - h });
        }
        return out;
    }
    // Giờ × thứ: số đơn và doanh thu theo giờ tạo đơn (thứ 2 → CN)
    const WEEKDAYS = ['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN'];
    function weekdayOf(iso) { const [y, m, d] = iso.split('-').map(Number); return (new Date(y, m - 1, d).getDay() + 6) % 7; }
    function hourMatrix(lines) {
        const cells = new Map(), orders = new Map(), days = new Map();
        for (const l of lines) {
            const h = Number(String(l.time || '').slice(0, 2));
            if (!Number.isFinite(h) || !/^\d{2}:/.test(l.time || '')) continue;
            const w = weekdayOf(l.created), k = w + '|' + h;
            const c = cells.get(k) || { w, h, orders: new Set(), qty: 0, rev: 0 };
            c.orders.add(l.order); c.qty += l.qty; c.rev += l.qty * l.price; cells.set(k, c);
            orders.set(l.order, 1);
            const ds = days.get(w) || new Set(); ds.add(l.created); days.set(w, ds);
        }
        const out = [...cells.values()].map(c => ({ w: c.w, h: c.h, orders: c.orders.size, qty: c.qty, rev: c.rev }));
        return { cells: out, dayCount: WEEKDAYS.map((_, i) => (days.get(i) || new Set()).size), hours: [...new Set(out.map(c => c.h))].sort((a, b) => a - b) };
    }
    // Kỳ so sánh: cùng kỳ tháng trước hoặc 7 ngày trước
    function compareRange(range, mode) { return mode === 'week' ? { from: addDays(range.from, -7), to: addDays(range.to, -7) } : prevMonthRange(range); }
    // Đọc số từ chữ đã định dạng vi-VN ("1.234,5 tr", "12,3%", "▲ 4") để sắp xếp cột
    function cellNumber(v) {
        if (typeof v === 'number') return v;
        const s = clean(v);
        if (!s || s === '—') return null;
        if (/^\d{2}\/\d{2}\/\d{4}/.test(s)) { const [d, m, y] = s.slice(0, 10).split('/'); return Number(y + m + d) + (Number(s.slice(11, 13) + s.slice(14, 16)) || 0) / 1e4; }
        const m = s.replace(/^[▲▼]\s*/, '').match(/^-?[\d.]+(,\d+)?/);
        if (!m || !/^[▲▼]?\s*-?[\d.]+(,\d+)?\s*(tr|%|ngày|đ)?$/.test(s)) return null;
        const n = Number(m[0].replace(/\./g, '').replace(',', '.'));
        return s.startsWith('▼') ? -n : n;
    }
    const toTsv = (headers, rows) => [headers, ...rows].map(r => r.map(c => String(c ?? '').replace(/[\t\n]/g, ' ')).join('\t')).join('\n');
    // Tin xin hàng gửi kho / ngành hàng: gom theo hãng → siêu thị
    function requestMessage(rows, nameOf, meta) {
        const list = rows.filter(r => r.need > 0);
        invariant(list.length, 'Không có mã nào cần xin');
        const byBrand = new Map(); list.forEach(r => { const k = r.brand || '(không rõ)'; (byBrand.get(k) || byBrand.set(k, []).get(k)).push(r); });
        const lines = [`📦 ĐỀ XUẤT XIN HÀNG ${String(meta?.cat || 'Điện thoại').toUpperCase()}${meta?.date ? ' · ' + meta.date : ''}`, meta?.basis ? `(${meta.basis})` : ''].filter(Boolean);
        for (const [brand, rs] of [...byBrand.entries()].sort((a, b) => b[1].reduce((t, r) => t + r.need, 0) - a[1].reduce((t, r) => t + r.need, 0))) {
            lines.push(`\n${brand.toUpperCase()} — ${rs.reduce((t, r) => t + r.need, 0)} máy`);
            const byShop = new Map(); rs.forEach(r => (byShop.get(r.shop) || byShop.set(r.shop, []).get(r.shop)).push(r));
            for (const [shop, xs] of byShop) {
                lines.push(`  ${nameOf(shop)}:`);
                xs.sort((a, b) => b.need - a.need).forEach(r => lines.push(`  • ${r.name}${r.product && r.product !== r.name ? ` (${r.product})` : ''} · tồn ${r.stock} · bán ${r.perDay}/ngày · xin ${r.need}${meta?.sourceOf?.(r) ? ` · ${meta.sourceLabel || 'có tại'}: ${meta.sourceOf(r)}` : ''}`));
            }
        }
        return lines.join('\n');
    }

    /* ================= V2.2: TRA TỒN CRM (crm.thegioididong.com/kiem-tra-ton-kho) =================
     * POST /SaleOrder/CheckInventory_View (form): lstStore = "mã,mã,...," (BẮT BUỘC có dấu phẩy cuối, thiếu → HTTP 500),
     * key = mã SP (1 mã) hoặc tên, inventoryStatusID 1 = Mới, lstStore rỗng → chỉ ra cột "Toàn hệ thống".
     * Trả về HTML bảng: hàng đầu = tiêu đề "Tổng tồn kho" + "mã - tên siêu thị"; mỗi dòng SP: ô đầu có mã 13 số.
     * Ô "2 (1)": 2 = TỒN BÁN ĐƯỢC (đã trừ mọi khóa), (1) = SL đang khóa (đối chiếu GetLockDetail 08/10/2026: QUANTITY 3 − khóa 1 = 2).
     * Ô "-" hoặc 0 = không xin được. Không dùng kho chi nhánh (anh Ngọc chốt 08/10/2026).
     * Đã kiểm 08/10/2026: ~0,2 giây/lần, số thời gian thực. */
    const CRM_DEFAULT_STORES = '1149,709,736,92,140,362,426,446,1999,2123,634,505,1709,690,3951,3947,1188,1263,1094,1202,1063,1246,1182,1247,1654,1537,1306,1310,1425,4525,5185,5108,5266,5243,5263,4849,6351,6908,7308,8027,8276,8280,8306,8061,8023,8265,8266,8359,8554,8874,9003,8598,8935,9356,9130,9172,9244,9907,9771,9340,10175,10294,10522,10157,11480,11999,11960,12260,12836,14064,14066,14073,85,1299,1127,5410,9291,28645,29901,29902,29968,30757,34879';   // V2.2.1: đã bỏ kho chi nhánh
    // Tỉnh/thành trên CRM (ddlProvince, 34 tỉnh sau sáp nhập 2025). Ninh Bình 1014 = Ninh Bình + Nam Định + Hà Nam cũ.
    const CRM_PROVINCES = [['1010', 'TP Bắc Ninh'], ['1032', 'TP Cần Thơ'], ['1020', 'TP Đà Nẵng'], ['1026', 'TP Đồng Nai'], ['1000', 'TP Hà Nội'], ['1012', 'TP Hải Phòng'], ['1027', 'TP Hồ Chí Minh'], ['1019', 'TP Huế'], ['1009', 'TP Quảng Ninh'],
        ['1031', 'An Giang'], ['1033', 'Cà Mau'], ['1001', 'Cao Bằng'], ['1024', 'Đắk Lắk'], ['1003', 'Điện Biên'], ['1029', 'Đồng Tháp'], ['1022', 'Gia Lai'], ['1017', 'Hà Tĩnh'], ['1013', 'Hưng Yên'], ['1023', 'Khánh Hòa'], ['1004', 'Lai Châu'],
        ['1008', 'Lạng Sơn'], ['1006', 'Lào Cai'], ['1025', 'Lâm Đồng'], ['1016', 'Nghệ An'], ['1014', 'Ninh Bình (gồm Nam Định, Hà Nam)'], ['1011', 'Phú Thọ'], ['1021', 'Quảng Ngãi'], ['1018', 'Quảng Trị'], ['1005', 'Sơn La'], ['1028', 'Tây Ninh'],
        ['1007', 'Thái Nguyên'], ['1015', 'Thanh Hóa'], ['1002', 'Tuyên Quang'], ['1030', 'Vĩnh Long']];
    // GetAllERPPMStores_string trả {html: "<option value='1149'>1149 - ĐMM_NBI_NQU - Nho Quan</option>…"}
    function parseStoreOptions(html) {
        const out = [], re = /<option[^>]*value=['"]?(\d+)['"]?[^>]*>([^<]*)<\/option>/gi; let m;
        while ((m = re.exec(String(html || '')))) { const full = clean(m[2]).replace(/^\d+\s*-\s*/, ''); out.push({ code: keyCode(m[1]), full, name: full.replace(/^[^-]*-\s*/, '').trim() || full, kho: isWarehouse(full) }); }
        return out;
    }
    // Khu vực từ mã siêu thị CRM: "ĐMM_NDI_HHA - Hải Anh" → chuỗi ĐMM, tỉnh cũ NDI, huyện HHA
    const AREA_NAMES = { HHA: 'Hải Hậu', TNI: 'Trực Ninh', NDI: 'TP Nam Định', VBA: 'Vụ Bản', NHU: 'Nghĩa Hưng', GTH: 'Giao Thủy', XTR: 'Xuân Trường', YYE: 'Ý Yên', MLO: 'Mỹ Lộc', NTR: 'Nam Trực',
        NBI: 'TP Ninh Bình', NQU: 'Nho Quan', KSO: 'Kim Sơn', TDI: 'Tam Điệp', YMO: 'Yên Mô', YKH: 'Yên Khánh', HLU: 'Hoa Lư', GVI: 'Gia Viễn', PLY: 'Phủ Lý', LNH: 'Lý Nhân', DTI: 'Duy Tiên', KBA: 'Kim Bảng', BLU: 'Bình Lục', TLI: 'Thanh Liêm' };
    const PROV_NAMES = { NDI: 'Nam Định', NBI: 'Ninh Bình', HNA: 'Hà Nam' };
    function storeArea(full) {
        const t = clean(full).split(/\s+-\s+/)[0].split('_');
        return t.length >= 3 ? { chain: t[0], prov: t[1], dist: t[2] } : { chain: t[0] || '', prov: t[1] || '', dist: '' };
    }
    const areaText = a => a && a.dist ? `${AREA_NAMES[a.dist] || a.dist}${PROV_NAMES[a.prov] ? ' (' + PROV_NAMES[a.prov] + ')' : ''}` : '';
    // Xếp nơi xin hàng cho 1 siêu thị: ⭐ ưu tiên → cùng huyện → cùng tỉnh cũ → còn lại; trong nhóm: tồn bán được nhiều trước
    const TIERS = ['⭐ Ưu tiên', 'Cùng huyện', 'Cùng tỉnh cũ', 'Xa hơn'];
    function rankSources(prods, stores, me, favs, exclude) {
        const info = new Map([].concat(stores || []).map(s => [s.code, s])), my = storeArea(info.get(keyCode(me))?.full || me?.full || ''), fav = new Set((favs || []).map(keyCode));
        return crmSources(prods, stores, [me, ...(exclude || [])]).map(x => {
            const a = storeArea(info.get(x.code)?.full);
            const tier = fav.has(x.code) ? 0 : a.dist && a.dist === my.dist && a.prov === my.prov ? 1 : a.prov && a.prov === my.prov ? 2 : 3;
            return { ...x, area: a, tier, tierName: TIERS[tier] };
        }).sort((a, b) => a.tier - b.tier || b.free - a.free || a.name.localeCompare(b.name, 'vi'));
    }
    const crmStoreList = text => [...new Set(String(text || '').split(/[^\d]+/).map(keyCode).filter(Boolean))];
    function crmCell(text) {
        const t = clean(text); if (!t) return null;
        const m = t.match(/^(-?[\d.,]+)\s*(?:\(\s*(-?[\d.,]+)\s*\))?/); if (!m) return null;
        const n = v => Number(String(v || '0').replace(/[.,]/g, ''));
        return { qty: n(m[1]), lock: m[2] ? n(m[2]) : 0 };   // qty = tồn bán được, lock = đang khóa
    }
    // doc = Document đã parse từ HTML CRM → { stores: [{code, name}], products: [{ code, name, status, total, byStore: {code: {qty, lock}} }] }
    const nodeText = node => { if (!node) return ''; const t = node.cloneNode(true); t.querySelectorAll?.('br').forEach(b => b.replaceWith(' ')); return clean(t.textContent).replace(/\s+/g, ' '); };
    function parseCrmDoc(doc) {
        const rows = [...doc.querySelectorAll('tr')];
        const head = rows.find(tr => [...tr.children].some(c => /^\s*(Tổng tồn kho|Toàn hệ thống)\s*$/i.test(c.textContent)));
        invariant(head, 'CRM không trả bảng tồn kho');
        const cols = [...head.children].map(c => clean(c.textContent));
        const stores = cols.map((t, i) => { const m = t.match(/^(\d+)\s*-\s*(.+)$/); return m ? { i, code: keyCode(m[1]), full: m[2], name: m[2].replace(/^[^-]*-\s*/, '').trim() || m[2] } : null; }).filter(Boolean);
        const totalIdx = cols.findIndex(t => /^(Tổng tồn kho|Toàn hệ thống)$/i.test(t));
        const products = [];
        for (const tr of rows) {
            if (tr === head) continue;
            const cells = [...tr.children]; if (!cells.length) continue;
            const first = nodeText(cells[0]), code = (first.match(/\b(\d{13})\b/) || [])[1];
            if (!code) continue;
            const byStore = {};
            stores.forEach(s => { const c = crmCell(nodeText(cells[s.i])); if (c) byStore[s.code] = c; });
            const before = first.split(code)[0];
            const name = clean(before.replace(/\s(KD|Ngừng|Hết|Sắp|Đứt|Tạm|Chờ)\s.*$/i, '')) || clean(before);
            products.push({ code, name, status: clean(before.slice(name.length)), demo: isNotNewName(name), notSale: isServiceName(name) || !/^\d{13}$/.test(code),
                total: totalIdx >= 0 ? crmCell(nodeText(cells[totalIdx])) : null, byStore, service: Object.values(byStore).some(c => c.qty >= 9999) });
        }
        return { stores: stores.map(({ code, name, full }) => ({ code, name, full })), products };
    }
    // V2.3.1: chỉ hàng Mới để bán — bỏ mã DEMO / trưng bày / đã kích hoạt / cũ và mã dịch vụ, bảo hành, PMH
    const isNotNewName = n => /\bDEMO\b|trưng bày|trung bay|đã kích hoạt|da kich hoat|like ?new|(^|[\s(])cũ($|[\s)])|refurb/i.test(String(n || ''));
    const isServiceName = n => /^(Dịch vụ|Bảo hành|Gói|PMH|Phí|Phiếu|Voucher|Thẻ cào|Sim số)/i.test(clean(n));
    const crmUsable = p => p && !p.service && !p.demo && !p.notSale;
    const isWarehouse = full => /^VHN_|Kho /i.test(full || '');
    // Nguồn xin được = siêu thị có TỒN BÁN ĐƯỢC > 0 (không tính siêu thị trong cụm, không tính kho chi nhánh), nhiều nhất lên đầu
    // prods: 1 sản phẩm hoặc nhiều mã cùng model (cộng dồn theo siêu thị)
    function crmSources(prods, stores, exclude) {
        const ex = new Set((exclude || []).map(keyCode)), info = new Map([].concat(stores || []).map(s => [s.code, s])), sum = new Map();
        for (const p of [].concat(prods).filter(Boolean)) for (const [code, c] of Object.entries(p.byStore)) {
            if (c.qty >= 9999) continue;
            const x = sum.get(code) || { code, name: info.get(code)?.name || code, kho: isWarehouse(info.get(code)?.full), free: 0, lock: 0 };
            x.free += Math.max(0, c.qty); x.lock += c.lock; sum.set(code, x);
        }
        return [...sum.values()].filter(x => x.free > 0 && !x.kho && !ex.has(x.code)).sort((a, b) => b.free - a.free || a.name.localeCompare(b.name, 'vi'));
    }

    /* ---------- V2.4: tỉnh lân cận + hàng thay thế (khi cả tỉnh hết) ---------- */
    // Tỉnh giáp ranh (34 tỉnh sau sáp nhập 2025), gần / hay xin trước. Đổi được trong ⚙️ Cài đặt → Tra tồn CRM → Tỉnh lân cận.
    const NEAR_PROVS = { '1014': ['1015', '1013', '1000', '1011'], '1015': ['1014', '1016', '1011', '1005'], '1013': ['1014', '1000', '1012', '1010'],
        '1000': ['1010', '1013', '1014', '1011', '1007'], '1012': ['1009', '1013', '1010'], '1010': ['1000', '1007', '1009', '1012', '1013', '1008'],
        '1011': ['1000', '1014', '1015', '1005', '1002', '1006', '1007'], '1009': ['1012', '1010', '1008'], '1007': ['1000', '1010', '1002', '1001', '1008'],
        '1008': ['1001', '1007', '1010', '1009'], '1001': ['1002', '1007', '1008'], '1002': ['1006', '1011', '1007', '1001'], '1006': ['1004', '1005', '1011', '1002'],
        '1005': ['1003', '1004', '1006', '1011', '1015'], '1004': ['1003', '1006', '1005'], '1003': ['1004', '1005'], '1016': ['1015', '1017'], '1017': ['1016', '1018'],
        '1018': ['1017', '1019'], '1019': ['1018', '1020'], '1020': ['1019', '1021'], '1021': ['1020', '1022'], '1022': ['1021', '1023', '1024'],
        '1023': ['1022', '1024', '1025'], '1024': ['1022', '1023', '1025'], '1025': ['1023', '1024', '1026'], '1026': ['1027', '1025', '1028'],
        '1027': ['1026', '1028', '1030', '1029'], '1028': ['1027', '1026', '1029'], '1029': ['1028', '1030', '1031', '1032'], '1030': ['1027', '1029', '1032'],
        '1032': ['1030', '1031', '1029', '1033'], '1031': ['1029', '1032', '1033'], '1033': ['1032', '1031'] };
    // Tỉnh lân cận để tìm thêm: theo cài đặt riêng (nếu có), không thì lấy tỉnh giáp ranh của các tỉnh đang chọn; bỏ tỉnh đã chọn
    function nearProvsOf(selected, override) {
        const sel = new Set([].concat(selected || []).map(String));
        const src = Array.isArray(override) && override.length ? override.map(String) : [].concat(...[...sel].map(id => NEAR_PROVS[id] || []));
        return [...new Set(src)].filter(id => !sel.has(id)).slice(0, 8);
    }
    const provShort = id => (CRM_PROVINCES.find(p => p[0] === String(id))?.[1] || String(id)).replace(/\s*\(.*\)$/, '').replace(/^TP\s+/, '');
    // Gốc model để tìm màu / dung lượng khác: bỏ "Điện thoại", cắt từ dung lượng (256GB, (8+128GB), 1TB); không có dung lượng thì bỏ màu
    function variantBase(name) {
        const s = clean(name).replace(/^(điện thoại|máy tính bảng|đồng hồ thông minh)\s+/i, '');
        const m = s.match(/^(.*?)\s*\(?\b\d+(?:\s*\+\s*\d+)?\s*(?:GB|TB)\b/i);
        return clean(m && m[1].length >= 3 ? m[1] : modelOf(s)).replace(/[\s(/,-]+$/, '');
    }
    // Hàng thay thế: các mã cùng gốc model (khác màu / dung lượng), chỉ hàng Mới, còn tồn bán được trong tỉnh; gần nhiều trước
    function altOptions(p, prods, stores, me, favs) {
        const base = norm(variantBase(p.name));
        return [].concat(prods || []).filter(x => crmUsable(x) && x.code !== p.code && norm(variantBase(x.name)) === base).map(x => {
            const list = rankSources(x, stores, me, favs), near = list.filter(s => s.tier <= 2);
            return { p: x, list, near, nearQty: near.reduce((a, s) => a + s.free, 0), allQty: list.reduce((a, s) => a + s.free, 0) };
        }).filter(a => a.list.length).sort((a, b) => b.nearQty - a.nearQty || b.allQty - a.allQty || a.p.name.localeCompare(b.p.name, 'vi', { numeric: true }));
    }
    // Nguồn ở tỉnh lân cận. nb = [{ prov, name, res }] theo thứ tự gần trước; trong 1 tỉnh: tồn nhiều trước
    function nearbySources(p, nb, me) {
        const out = [];
        [].concat(nb || []).forEach((d, i) => {
            const x = d.res?.products?.find(q => q.code === p.code); if (!x) return;
            const info = new Map(d.res.stores.map(t => [t.code, t]));
            crmSources(x, d.res.stores, [me]).forEach(s => out.push({ ...s, provId: d.prov, provName: d.name, order: i, area: storeArea(info.get(s.code)?.full), tier: 4, tierName: 'Tỉnh lân cận' }));
        });
        return out.sort((a, b) => a.order - b.order || b.free - a.free || a.name.localeCompare(b.name, 'vi'));
    }
    // Gộp 2 kết quả CRM (tra theo từng nhóm siêu thị) thành 1
    function mergeCrm(a, b) {
        if (!a) return b; if (!b) return a;
        const map = new Map(a.products.map(p => [p.code, { ...p, byStore: { ...p.byStore } }]));
        for (const p of b.products) { const x = map.get(p.code); if (x) Object.assign(x.byStore, p.byStore); else map.set(p.code, { ...p, byStore: { ...p.byStore } }); }
        const seen = new Set(a.stores.map(s => s.code));
        return { stores: a.stores.concat(b.stores.filter(s => !seen.has(s.code))), products: [...map.values()] };
    }

    // Khôi phục sao lưu: gộp 1 tháng — ngày nào bản nào lấy mới hơn (atMs) thì dùng bản đó (cả sổ ngày lẫn dòng bán)
    function mergeMonth(cur, inc) {
        cur = cur && typeof cur === 'object' ? cur : { book: {}, lines: [] };
        const book = { ...(cur.book || {}) }, take = new Set();
        for (const [d, rec] of Object.entries(inc?.book || {})) if (!book[d] || (rec?.atMs || 0) > (book[d]?.atMs || 0)) { book[d] = rec; take.add(d); }
        const lines = (cur.lines || []).filter(l => !take.has(l.created)).concat((inc?.lines || []).filter(l => take.has(l.created)));
        lines.sort((a, b) => a.created.localeCompare(b.created) || String(a.shop).localeCompare(String(b.shop)) || String(a.time).localeCompare(String(b.time)));
        return { book, lines, taken: take.size };
    }

    /* ---------- V2.5: chia số cần xin cho từng nơi cho (theo tồn bán được trên CRM) ----------
     * sources đã xếp: ⭐ → cùng huyện → cùng tỉnh cũ → xa hơn → tỉnh lân cận (tier 0..4).
     * Chọn nơi tier nhỏ nhất, cùng tier thì nơi còn dư nhiều nhất (nơi tồn nhiều cho nhiều), lấy tối đa ở nơi đó rồi mới sang nơi khác.
     * Lượt 1: nơi cho giữ lại `keep` máy. Lượt 2 (vẫn thiếu): lấy cả máy giữ lại, đánh dấu last ("máy cuối").
     * used (Map) dùng chung giữa nhiều dòng cùng sản phẩm để 2 siêu thị trong cụm không xin trùng 1 máy. */
    function allocateAsk(need, sources, keep = 1, used, pkey = '', noLast = false) {
        keep = Math.max(0, Math.round(Number(keep) || 0));
        const k = c => c + '|' + pkey;
        const s = [].concat(sources || []).filter(x => x && x.free > 0).map((x, i) => {
            const free = Math.max(0, x.free - (used?.get(k(x.code)) || 0));
            return { x, i, free, cap: Math.max(0, free - keep), extra: Math.min(free, keep), take: 0, last: 0 };
        });
        let left = Math.max(0, Math.round(Number(need) || 0));
        const pick = key => { let b = null; for (const y of s) if (y[key] > 0 && (!b || y.x.tier < b.x.tier || (y.x.tier === b.x.tier && (y[key] > b[key] || (y[key] === b[key] && y.i < b.i))))) b = y; return b; };
        // Lấy gọn: mỗi lượt chọn nơi tốt nhất rồi lấy tối đa có thể ở nơi đó → ít lượt chuyển nhất
        for (const key of noLast ? ['cap'] : ['cap', 'extra']) while (left > 0) { const y = pick(key); if (!y) break; const q = Math.min(y[key], left); y[key] -= q; y.take += q; if (key === 'extra') y.last += q; left -= q; }
        const plan = s.filter(y => y.take).sort((a, b) => a.x.tier - b.x.tier || b.take - a.take || a.i - b.i).map(y => {
            if (used) used.set(k(y.x.code), (used.get(k(y.x.code)) || 0) + y.take);
            return { code: y.x.code, name: y.x.name, tier: y.x.tier, tierName: y.x.tierName, area: y.x.area, provName: y.x.provName, free: y.free, qty: y.take, last: y.last, left: y.free - y.take };
        });
        return { need: Math.max(0, Math.round(Number(need) || 0)), plan, got: plan.reduce((a, p) => a + p.qty, 0), short: left };
    }
    // Nhiều dòng (nhiều siêu thị trong cụm) xin cùng lúc: chia theo vòng — vòng 1 chỉ ⭐ + cùng huyện, vòng 2 thêm cùng tỉnh cũ,
    // vòng 3 mọi nơi (vẫn giữ lại máy), vòng cuối mới lấy máy giữ lại → siêu thị nào cũng được nơi gần mình trước.
    // items: [{ key, need, sources, pkey }] → Map key → { need, plan, got, short }
    function allocateMany(items, keep = 1) {
        const used = new Map(), acc = new Map(items.map(it => [it.key, { need: Math.max(0, Math.round(Number(it.need) || 0)), plan: [], got: 0, short: 0 }]));
        const merge = (a, b) => { for (const p of b.plan) { const x = a.plan.find(q => q.code === p.code); if (x) { x.qty += p.qty; x.last += p.last; x.left = p.left; } else a.plan.push({ ...p }); } a.got += b.got; };
        for (const [lim, noLast] of [[1, true], [2, true], [9, true], [9, false]])
            for (const it of items) {
                const a = acc.get(it.key), left = a.need - a.got; if (left <= 0) continue;
                merge(a, allocateAsk(left, [].concat(it.sources || []).filter(x => x.tier <= lim), keep, used, it.pkey || '', noLast));
            }
        for (const a of acc.values()) { a.short = a.need - a.got; a.plan.sort((x, y) => x.tier - y.tier || y.qty - x.qty); }
        return acc;
    }
    const planText = a => !a ? '' : (a.plan.map(p => `${p.tier === 0 ? '⭐' : ''}${p.name}${p.provName ? ` (${p.provName})` : ''} ×${p.qty}${p.last ? ' (máy cuối)' : ''}`).join(' · ') || 'không nơi nào cho được') + (a.short && a.plan.length ? ` · thiếu ${a.short}` : '');
    // Gom các đề xuất theo nơi cho → "phiếu" gửi từng siêu thị. items: [{ to, product, name, alloc }]
    function groupBySource(items) {
        const m = new Map();
        for (const it of items) for (const p of (it.alloc?.plan || [])) {
            const g = m.get(p.code) || { code: p.code, name: p.name, tier: p.tier, area: p.area, provName: p.provName, qty: 0, lines: [] };
            g.qty += p.qty; g.lines.push({ to: it.to, product: it.product, name: it.name, qty: p.qty, left: p.left, last: p.last }); m.set(p.code, g);
        }
        return [...m.values()].sort((a, b) => a.tier - b.tier || b.qty - a.qty || String(a.name).localeCompare(String(b.name), 'vi'));
    }

    // Hàm thuần cho kiểm thử offline; không cài global trên website thật.
    if (typeof module === 'object' && module.exports) {
        module.exports = { clean, norm, hasCode, day, toBI, addDays, validateRange, daysIn, apiNumber, parseDelimited,
            lineFromApi, reasons, pending, validateSalesLines, splitSales, summarizeSales, inPeriod, dayStatus, daysToFetch,
            parseRateLimit, waitBeforeCall, inventoryRecordFromApi, summarizeInventory, filterInventory, inventoryOptions, inventoryViews, validateShops, authorizeSheetRows, escHtml, inventoryChecklist, inventoryPrintHtml, conditionKey, conditionText, newerVersion, parseRemoteScript, outStoreKey, balanceRows, isFresh, shiftMonth, prevMonthRange, attachStats, pendingOrders, projectMonth, daysBetween,
            modelOf, transferPlan, ageDays, ageBucket, AGE_BUCKETS, transitDiff, hourMatrix, weekdayOf, compareRange, cellNumber, toTsv, requestMessage, mergeMonth, crmCell, parseCrmDoc, crmSources, crmStoreList, CRM_DEFAULT_STORES, CRM_PROVINCES, parseStoreOptions, storeArea, rankSources, areaText, isNotNewName, isServiceName, crmUsable,
            NEAR_PROVS, nearProvsOf, provShort, variantBase, altOptions, nearbySources, mergeCrm, allocateAsk, allocateMany, planText, groupBySource };
        return;
    }


    /* ================= TRÌNH DUYỆT ================= */
    // V2.5.1: mỗi bản đang cài ghi phiên bản của mình lên trang → phát hiện máy cài trùng 2 bản (cùng tên, 2 dòng trong Tampermonkey)
    try { const de = document.documentElement; de.dataset.kxbVersions = [de.dataset.kxbVersions, VERSION].filter(Boolean).join(','); } catch { /* bỏ qua */ }
    const W = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
    const XL = () => (typeof XLSX !== 'undefined' ? XLSX : W.XLSX);
    const defaults = { shops: [], selectors: {}, basis: 'created', returnDays: 7 };
    const load = (key, fallback) => { try { return GM_getValue(PREFIX + key, fallback); } catch { return fallback; } };
    const save = (key, value) => GM_setValue(PREFIX + key, value);
    let config = { ...defaults, ...load('config', {}) };
    let ui, running = null, auth = null;
    const view = { tab: 'sales', sales: null, salesTab: 'category', salesShops: null, salesConds: new Set(), salesCat: '', salesBrands: new Set(), openDrop: '', salesQuery: '', balDays: 10, balStatus: '', balShops: null, balBrands: new Set(), openStaff: new Set(), inv: null, invTab: 'group', invShops: null, invFilter: { category: '', group: '', brand: '', conditions: [], q: '' } };
    let journal = load('journal', []).slice(-300);
    if (load('journalVersion', '') !== VERSION) { journal = []; try { save('journal', journal); save('journalVersion', VERSION); } catch { /* bỏ qua */ } }
    function log(message, kind = 'info') {
        journal.push({ time: new Date().toISOString(), run: running?.id || null, kind, message: clean(message) });
        if (journal.length > 300) journal.shift();
        try { save('journal', journal); } catch { /* giữ trong bộ nhớ */ }
        if (ui) ui.querySelector('[data-log]').textContent = journal.slice(-150).reverse().map(x => `${new Date(x.time).toLocaleTimeString('vi-VN')} ${x.kind === 'error' ? '✖' : '•'} ${x.message}`).join('\n');
    }
    const own = e => !!e?.closest?.('[data-kxb-ui]');
    const visible = e => !!e?.isConnected && !own(e) && e.getClientRects().length > 0 && getComputedStyle(e).visibility !== 'hidden' && getComputedStyle(e).display !== 'none';
    const all = (s, root = document) => [...root.querySelectorAll(s)].filter(visible);
    const txt = e => clean(e?.innerText || e?.textContent);
    const configured = (key, root = document) => {
        if (!config.selectors?.[key]) return null;
        const found = [...root.querySelectorAll(config.selectors[key])].filter(e => !own(e));
        invariant(found.length === 1, `Selector ${key} phải xác định đúng 1 phần tử, nhận ${found.length}`);
        return found[0];
    };
    function check(session) {
        if (!session || session !== running || session.controller.signal.aborted || session.status !== 'running') fail('Phiên đã dừng hoặc thay đổi', 'CANCELLED');
    }

    /* ---------- Auth (giữ nguyên bản TEST) ---------- */
    function detectUser() {
        const override = configured('authUser');
        const elements = override ? [override] : all('nav a,header a,.navbar a,.navbar span,.user-info,.user-name,.profile a');
        const ids = new Set(elements.map(e => txt(e).match(/^(\d{4,10})\s*[-–]\s*\S/)?.[1]).filter(Boolean));
        invariant(ids.size === 1, 'Chưa xác định duy nhất mã nhân viên ở vùng tài khoản BI.');
        return [...ids][0];
    }
    function authCheck(session) {
        const user = detectUser();
        const url = `https://docs.google.com/spreadsheets/d/${AUTH_SHEET.id}/export?format=csv&gid=${AUTH_SHEET.gid}&_t=${Date.now()}`;
        return new Promise((resolve, reject) => {
            const stop = () => { request?.abort(); reject(Object.assign(new Error('Đã dừng Auth'), { code: 'CANCELLED' })); };
            const finish = (error, response) => {
                session.controller.signal.removeEventListener('abort', stop);
                try {
                    check(session); if (error) throw error;
                    invariant(response.status === 200, `Không đọc được sheet Auth (HTTP ${response.status})`);
                    const csv = String(response.responseText || '');
                    invariant(csv.length > 0 && csv.length < 2 * 1024 * 1024, 'Sheet Auth trống hoặc phản hồi quá lớn');
                    auth = authorizeSheetRows(parseDelimited(csv), user);
                    invariant(detectUser() === user, 'Tài khoản đã thay đổi'); resolve(auth);
                }
                catch (e) { reject(e); }
            };
            let request;
            session.controller.signal.addEventListener('abort', stop, { once: true });
            request = GM_xmlhttpRequest({ method: 'GET', url, timeout: 15000,
                onload: r => finish(null, r), onerror: () => finish(new Error('Lỗi mạng khi đọc sheet Auth; không sử dụng quyền cũ')),
                ontimeout: () => finish(new Error('Đọc sheet Auth quá thời gian')) });
        });
    }

    /* ---------- Gọi thẳng BI bằng phiên đăng nhập sẵn có ---------- */
    async function postBI(path, body, session, timeoutMs = 180000) {
        check(session);
        const ctl = new AbortController();
        const onStop = () => ctl.abort();
        session.controller.signal.addEventListener('abort', onStop, { once: true });
        const timer = setTimeout(() => ctl.abort(), timeoutMs);
        try {
            const r = await W.fetch(path, { method: 'POST', credentials: 'include', signal: ctl.signal,
                headers: { 'Content-Type': 'application/json;charset=UTF-8' }, body: JSON.stringify(body) });
            const text = await r.text();
            invariant(r.ok, `BI trả HTTP ${r.status}`);
            try { return JSON.parse(text); } catch { fail('BI không trả dữ liệu — phiên đăng nhập có thể đã hết; đăng nhập lại BI rồi chạy lại'); }
        } catch (e) {
            if (e.name === 'AbortError') { check(session); fail(`BI phản hồi quá ${Math.round(timeoutMs / 1000)} giây; không lưu dữ liệu thiếu`, 'TIMEOUT'); }
            throw e;
        } finally { clearTimeout(timer); session.controller.signal.removeEventListener('abort', onStop); }
    }
    /* ---------- V2.2: gọi CRM bằng phiên đăng nhập CRM sẵn có trên Chrome (GM_xmlhttpRequest gửi kèm cookie CRM) ---------- */
    const CRM_URL = 'https://crm.thegioididong.com/SaleOrder/CheckInventory_View';
    function crmFetch(key, stores, session) {
        return new Promise((resolve, reject) => {
            if (session) check(session);
            const body = new URLSearchParams({ lstStore: stores.length ? stores.join(',') + ',' : '', lstMainGroup: '', key, typeSort: '1', inventoryStatusID: '1', ProductIDRef: '0', rangeType: '-1', SiteID: '', IsCheckInventory: '0' }).toString();
            let req;
            const stop = () => { req?.abort?.(); reject(Object.assign(new Error('Người dùng đã dừng'), { code: 'CANCELLED' })); };
            session?.controller.signal.addEventListener('abort', stop, { once: true });
            const done = fn => (...a) => { session?.controller.signal.removeEventListener('abort', stop); fn(...a); };
            req = GM_xmlhttpRequest({ method: 'POST', url: CRM_URL, data: body, timeout: 30000,
                headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8', 'X-Requested-With': 'XMLHttpRequest' },
                onload: done(r => {
                    const html = String(r.responseText || '');
                    if (r.status !== 200) return reject(new Error(r.status === 500 ? 'CRM báo lỗi (HTTP 500)' : `CRM trả HTTP ${r.status}`));
                    if (/type=["']?password|Đăng nhập|dang-nhap|login/i.test(html) && !/Tổng tồn kho|Toàn hệ thống/.test(html)) return reject(Object.assign(new Error('Chưa đăng nhập CRM — mở crm.thegioididong.com trên Chrome, đăng nhập rồi bấm lại'), { code: 'CRM_LOGIN' }));
                    if (!/Tổng tồn kho|Toàn hệ thống/.test(html)) return resolve({ stores: [], products: [] });   // không có sản phẩm khớp
                    try { resolve(parseCrmDoc(new DOMParser().parseFromString(html, 'text/html'))); } catch (e) { reject(e); }
                }),
                onerror: done(() => reject(new Error('Lỗi mạng khi gọi CRM'))), ontimeout: done(() => reject(new Error('CRM phản hồi quá 30 giây'))) });
        });
    }
    const crmProvs = () => (Array.isArray(config.crmProvinces) && config.crmProvinces.length ? config.crmProvinces : ['1014']).map(String).sort();
    const crmStoreCacheOk = () => { const c = config.crmStoreCache; return !!(c && c.key === crmProvs().join(',') && Array.isArray(c.stores) && c.stores.length); };
    // Siêu thị tra tồn = siêu thị của các tỉnh đã chọn (bỏ kho chi nhánh). Chưa lấy được thì dùng danh sách mặc định (tỉnh Ninh Bình).
    const crmStores = () => crmStoreCacheOk() ? config.crmStoreCache.stores.filter(x => !x.kho).map(x => x.code) : crmStoreList(CRM_DEFAULT_STORES);
    function crmGetStores(provs) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({ method: 'POST', url: 'https://crm.thegioididong.com/AjaxApi/GetAllERPPMStores_string', data: new URLSearchParams({ lstProvince: provs.join(',') + ',', type: 'true' }).toString(), timeout: 30000,
                headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8', 'X-Requested-With': 'XMLHttpRequest' },
                onload: r => {
                    const t = String(r.responseText || ''); let j = null; try { j = JSON.parse(t); } catch { /* không phải JSON */ }
                    if (!j || typeof j.html !== 'string') return reject(Object.assign(new Error(/password|Đăng nhập|login/i.test(t) ? 'Chưa đăng nhập CRM — mở crm.thegioididong.com trên Chrome, đăng nhập rồi bấm lại' : `CRM không trả danh sách siêu thị (HTTP ${r.status})`), { code: 'CRM_LOGIN' }));
                    resolve(parseStoreOptions(j.html));
                },
                onerror: () => reject(new Error('Lỗi mạng khi gọi CRM')), ontimeout: () => reject(new Error('CRM phản hồi quá 30 giây')) });
        });
    }
    // Lấy danh sách siêu thị theo tỉnh (lưu 7 ngày; đổi tỉnh là lấy lại)
    async function ensureCrmStores(force) {
        const key = crmProvs().join(',');
        if (!force && crmStoreCacheOk() && Date.now() - (config.crmStoreCache.at || 0) < 7 * 864e5) return;
        const list = await crmGetStores(crmProvs());
        invariant(list.length, 'CRM không trả siêu thị nào cho tỉnh đã chọn');
        config.crmStoreCache = { key, at: Date.now(), stores: list }; save('config', config);
        log(`Danh sách siêu thị tra tồn CRM: ${crmProvs().map(id => CRM_PROVINCES.find(p => p[0] === id)?.[1] || id).join(', ')} — ${list.filter(x => !x.kho).length} siêu thị (bỏ ${list.filter(x => x.kho).length} kho chi nhánh)`);
        renderCrmSetting();
    }
    const definitions = {};
    async function definition(id, session) {
        if (definitions[id]) return definitions[id];
        const r = await postBI('/Home/GetDynamicReport', { dynamicReportId: id }, session, 60000);
        invariant(r && r.success && r.data?.DynamicReportCondition?.length, `Không mở được định nghĩa báo cáo ${id}; tài khoản có quyền xem không?`);
        return (definitions[id] = r.data.DynamicReportCondition);
    }
    // Điều tốc: BI chỉ cho tối đa 5 lần / 60 giây trên mỗi báo cáo (tính cả lần bạn tự bấm trên web)
    const rate = { calls: {}, limit: {} };
    async function waitTurn(id, session, onWait) {
        const lim = rate.limit[id] || { max: 5, windowSec: 60 };
        const times = rate.calls[id] || (rate.calls[id] = []);
        for (;;) {
            const ms = waitBeforeCall(times, Date.now(), lim.max, lim.windowSec * 1000 + 1500);
            if (ms <= 0) break;
            onWait?.(Math.ceil(ms / 1000));
            await pause(Math.min(ms, 1000), session);
        }
        times.push(Date.now());
        while (times.length > 50) times.shift();
    }
    async function runReport(id, values, session, timeoutMs, onWait) {
        for (let attempt = 1; ; attempt++) {
            await waitTurn(id, session, onWait);
            try { return await runReportOnce(id, values, session, timeoutMs); }
            catch (e) {
                const rl = e.rateLimit;
                if (!rl || attempt >= 6) throw e;
                rate.limit[id] = { max: rl.max, windowSec: rl.windowSec };
                // BI vừa từ chối: coi như cửa sổ đã đầy, chờ đúng số giây BI yêu cầu rồi thử lại
                const until = Date.now() + (rl.waitSec + 2) * 1000;
                log(`BI báo thao tác quá nhanh — chờ ${rl.waitSec + 2} giây rồi thử lại (lần ${attempt + 1})`);
                while (Date.now() < until) { onWait?.(Math.ceil((until - Date.now()) / 1000)); await pause(1000, session); }
            }
        }
    }
    function pause(ms, session) {
        return new Promise((resolve, reject) => {
            check(session);
            const stop = () => { clearTimeout(timer); reject(Object.assign(new Error('Người dùng đã dừng'), { code: 'CANCELLED' })); };
            const timer = setTimeout(() => { session.controller.signal.removeEventListener('abort', stop); resolve(); }, ms);
            session.controller.signal.addEventListener('abort', stop, { once: true });
        });
    }
    async function runReportOnce(id, values, session, timeoutMs) {
        const c = await definition(id, session);
        for (const k of Object.keys(values)) invariant(c.some(x => x.PARAMNAME === k), `Báo cáo ${id} không còn tham số ${k}; BI đã đổi định nghĩa, cần cập nhật script`);
        const listParam = c.map(x => ({
            CONTROLTYPE: x.CONTROLTYPE, OBJECTVALUE: values[x.PARAMNAME] != null ? String(values[x.PARAMNAME]) : (x.DEFAULTVALUE || ''),
            PARAMNAME: x.PARAMNAME, PARAMTYPE: x.PARAMTYPE, DATATYPE: x.DATATYPE || '', ISSPLITVALUE: String(!!x.ISSPLITVALUE),
            ISHIDDEN: String(!!x.ISHIDDEN), ROOTCONTROLDATATYPE: x.ROOTCONTROLDATATYPE || ''
        }));
        const h = c[0];
        const body = { dynamicReportId: String(id), dynamicReportName: h.DYNAMICREPORTNAME, listParam, storeName: h.SQLCOMMAND,
            timeOut: String(h.TIMEOUT || 600), connectionString: h.CONNECTIONSTRING, isExportOnly: 'false', isGridView: 'true',
            isDataLake: String(!!h.ISDATALAKE), sourceType: h.SOURCETYPE || 'DATABASE', pageIndex: '1', pageSize: '200000',
            reportGuid: uuid(), note: '', isShowSchedule: '', take: '200000', skip: '0', page: '1' };
        const r = await postBI('/Home/FilterDataDynamicReport', body, session, timeoutMs);
        if (!r || r.Success !== true) {
            const msg = clean(r?.Message).split('\n')[0] || 'không rõ';
            const e = new Error(`BI báo lỗi: ${msg}`); e.code = 'BI'; e.rateLimit = parseRateLimit(msg); throw e;
        }
        const rows = r.Data?.listResult || [];
        const total = Number(r.Data?.totalRow ?? rows[0]?.totalrow ?? rows.length);
        invariant(rows.length === total, `BI báo ${total} dòng nhưng chỉ nhận ${rows.length}; không lưu dữ liệu thiếu`);
        return rows;
    }

    /* ---------- Lưu trữ trên máy ---------- */
    const monthOf = iso => iso.slice(0, 7);
    const shopsKey = () => config.shops.map(s => keyCode(s.code)).join(',');
    /* ---------- Kho số xuất bán theo tháng — V2.1: IndexedDB (nhanh, không giới hạn ~40 MB), dự phòng Tampermonkey ----------
     * Đọc hết vào bộ nhớ khi mở (mstore.cache) → các hàm đọc vẫn chạy đồng bộ như cũ; ghi IndexedDB chạy nền.
     * Tháng cũ trong Tampermonkey được chép sang IndexedDB và GIỮ LẠI làm bản dự phòng (không cập nhật nữa) cho tới khi bấm "Xóa bản cũ".
     * Nếu IndexedDB lỗi / bị xóa: tự quay về Tampermonkey và lấy lại các tháng từ bản dự phòng. */
    const DB_NAME = 'AutoBI_KhoXuatBan', DB_STORE = 'months';
    const mstore = { mode: 'gm', cache: new Map(), db: null, error: '', pending: 0 };
    let storeReady = Promise.resolve();
    const gmMonthKeys = () => (typeof GM_listValues === 'function' ? GM_listValues() : []).filter(k => k.startsWith(PREFIX + 'month_')).sort();
    const idbReq = req => new Promise((res, rej) => { req.onsuccess = () => res(req.result); req.onerror = () => rej(req.error || new Error('IndexedDB lỗi')); });
    function idbTx(mode, fn) {
        return new Promise((res, rej) => {
            const tx = mstore.db.transaction(DB_STORE, mode), out = fn(tx.objectStore(DB_STORE));
            tx.oncomplete = () => res(out); tx.onerror = () => rej(tx.error || new Error('IndexedDB lỗi')); tx.onabort = () => rej(tx.error || new Error('IndexedDB bị hủy (đầy bộ nhớ?)'));
        });
    }
    async function initStore() {
        try {
            const idb = (typeof indexedDB !== 'undefined' && indexedDB) || W.indexedDB;
            invariant(idb, 'Trình duyệt không có IndexedDB');
            const req = idb.open(DB_NAME, 1);
            req.onupgradeneeded = () => { if (!req.result.objectStoreNames.contains(DB_STORE)) req.result.createObjectStore(DB_STORE); };
            mstore.db = await idbReq(req);
            const got = await idbTx('readonly', os => ({ keys: os.getAllKeys(), vals: os.getAll() }));
            got.keys.result.forEach((k, i) => mstore.cache.set(String(k), got.vals.result[i]));
            let moved = 0;
            for (const k of gmMonthKeys()) {
                const m = k.slice((PREFIX + 'month_').length);
                if (mstore.cache.has(m)) continue;                      // đã có trong IndexedDB → bản Tampermonkey chỉ để dự phòng
                const v = load('month_' + m, null); if (!v || typeof v !== 'object') continue;
                await idbTx('readwrite', os => os.put(v, m)); mstore.cache.set(m, v); moved++;
            }
            mstore.mode = 'idb';
            if (moved) log(`Đã chép ${moved} tháng số xuất bán sang bộ nhớ IndexedDB (bản cũ trong Tampermonkey giữ lại làm dự phòng)`);
            try { navigator.storage?.persist?.(); } catch { /* bỏ qua */ }
        } catch (e) {
            mstore.mode = 'gm'; mstore.error = e.message || String(e); mstore.cache.clear();
            log(`Không dùng được IndexedDB (${mstore.error}) — lưu số trong Tampermonkey như bản cũ`, 'error');
        }
    }
    function loadMonth(month) {
        if (mstore.mode === 'idb') { const x = mstore.cache.get(month); return x && typeof x === 'object' ? x : { book: {}, lines: [] }; }
        const x = load('month_' + month, null); return x && typeof x === 'object' ? x : { book: {}, lines: [] };
    }
    function saveMonth(month, data) {
        if (mstore.mode !== 'idb') { save('month_' + month, data); return; }
        mstore.cache.set(month, data); mstore.pending++;
        idbTx('readwrite', os => os.put(data, month)).catch(e => {
            log(`Ghi IndexedDB tháng ${month} lỗi (${e.message}) — ghi tạm vào Tampermonkey`, 'error');
            try { save('month_' + month, data); } catch { /* bỏ qua */ }
        }).finally(() => { mstore.pending--; });
    }
    function deleteMonth(month) {
        if (mstore.mode === 'idb') { mstore.cache.delete(month); idbTx('readwrite', os => os.delete(month)).catch(e => log('Xóa tháng trong IndexedDB lỗi: ' + e.message, 'error')); }
        try { GM_deleteValue(PREFIX + 'month_' + month); } catch { /* bỏ qua */ }
    }
    const listMonths = () => mstore.mode === 'idb' ? [...mstore.cache.keys()].sort() : gmMonthKeys().map(k => k.slice((PREFIX + 'month_').length));
    function loadBook(range) {
        const months = new Set(daysIn({ from: addDays(range.from, -RECHECK_DAYS), to: range.to }).map(monthOf));
        const book = {}, lines = [];
        for (const m of months) { const d = loadMonth(m); Object.assign(book, d.book); lines.push(...d.lines); }
        return { book, lines };
    }
    function storeDay(r) {
        const month = monthOf(r.day), d = loadMonth(month);
        // Dòng bán đã lưu luôn là Đã xuất – Đã giao – Chưa hủy – Chưa trả → bỏ các trường trạng thái cho nhẹ bộ nhớ
        d.lines = d.lines.filter(l => l.created !== r.day).concat(r.lines.map(({ exportType, exported, delivered, cancelled, returned, ...x }) => x));
        d.book[r.day] = r.record;
        d.lines.sort((a, b) => a.created.localeCompare(b.created) || a.shop.localeCompare(b.shop) || a.time.localeCompare(b.time));
        saveMonth(month, d);
    }

    /* ---------- Hai việc chính ---------- */
    async function inventoryShop(shop, session, withTransit) {
        const today = new Date();
        const rows = await runReport(REPORT.inventory, {
            p_todate: `${today.getDate()}/${today.getMonth() + 1}/${today.getFullYear()} 23:59`, p_storeidlist: keyCode(shop.code),
            p_ischeckrealinput: 'false',   // KHÔNG tính hàng đang chuyển kho (bỏ tick ô trên web)
            p_instockstatus: '-1', p_storetype: '-1', p_languageid: '2'
        }, session, 240000, sec => status(`Tồn kho ${shop.code}: chờ lượt BI ${sec} giây`));
        check(session);
        const records = rows.map(inventoryRecordFromApi);
        const summary = summarizeInventory(records, shop);     // kiểm tra TỪNG dòng đúng mã shop
        log(`Tồn ${shop.code} — ${shop.name}: ${rows.length} dòng, ${fmt(summary.quantity)} SL`);
        let transit = null;
        if (withTransit) {
            // Lần đổ thứ 2: có tính hàng đang chuyển kho → phần chênh lệch = hàng đang về siêu thị này (đã xuất từ nơi gửi, chưa nhận)
            const rows2 = await runReport(REPORT.inventory, {
                p_todate: `${today.getDate()}/${today.getMonth() + 1}/${today.getFullYear()} 23:59`, p_storeidlist: keyCode(shop.code),
                p_ischeckrealinput: 'true', p_instockstatus: '-1', p_storetype: '-1', p_languageid: '2'
            }, session, 240000, sec => status(`Hàng đang chuyển ${shop.code}: chờ lượt BI ${sec} giây`));
            check(session);
            const all2 = rows2.map(inventoryRecordFromApi);
            summarizeInventory(all2, shop);
            transit = transitDiff(all2, records);
            log(`Đang về ${shop.code}: ${transit.length} dòng, SL ${fmt(transit.reduce((a, r) => a + r.qty, 0))}`);
        }
        return { records, transit };
    }
    let warnedCondition = false, warnedOutStore = false;
    async function salesDay(iso, session, onWait) {
        const rows = await runReport(REPORT.sales, {
            V_FROMDATE: toBI(iso), V_TODATE: toBI(iso), V_OUTPUTTYPEIDLIST: SALES_FILTER.exportType, V_STORESEARCHTYPE: SALES_FILTER.warehouseMode,
            V_STOREIDLIST: shopsKey(), V_MAINGROUPIDLIST: SALES_FILTER.category,   // '' = tất cả ngành
            V_SALEORDERTYPEIDLIST: '', V_PRODUCTIDLIST: '',
            V_DELIVERYTYPEIDLIST: '', V_COMPANYID: '1', V_LANGUAGEID: '2', V_USERNAME: ''
        }, session, 180000, onWait);
        check(session);
        let condKey = '';
        for (const r of rows) { condKey = conditionKey(r); if (condKey) break; }
        if (rows.length && !condKey && !warnedCondition) {
            warnedCondition = true;
            const cols = Object.keys(rows[0]).filter(k => !/KHACHHANG|CUSTOMER|DIENTHOAI|PHONE|DIACHI|ADDRESS|EMAIL/i.test(k));
            log(`Không thấy cột Loại hàng (Mới / Đã sử dụng) trong BI 77 — sao chép nhật ký gửi anh Ngọc. Các cột BI trả về: ${cols.join(', ')}`, 'error');
        }
        if (rows.length && !warnedOutStore && !rows.some(r => outStoreKey(r))) {
            warnedOutStore = true;
            log('Không thấy cột Kho xuất trong BI 77 — Cân hàng tính mọi dòng bán là trừ tồn shop. Sao chép nhật ký gửi anh Ngọc để chỉnh.', 'error');
        }
        const lines = rows.map(r => lineFromApi(r, condKey));
        validateSalesLines(lines, config.shops, iso);
        const s = splitSales(lines), now = new Date();
        const record = { schema: SALES_SCHEMA, shops: shopsKey(), status: dayStatus(iso, s.pending, isoDate(now), config.returnDays), pending: s.pending, pendingLines: s.pendingLines, returned: s.returned, returnedLines: s.returnedLines, rows: rows.length, valid: s.valid.length,
            quantity: s.valid.reduce((a, l) => a + l.qty, 0), revenue: s.valid.reduce((a, l) => a + l.qty * l.price, 0), excluded: s.excluded, at: stamp(now), atMs: now.getTime() };
        return { day: iso, record, lines: s.valid };
    }

    /* ---------- Phiên chạy (cơ chế bản TEST) ---------- */
    function clockText(start, end = Date.now()) { const s = Math.max(0, Math.floor((end - start) / 1000)); return [Math.floor(s / 3600), Math.floor(s / 60) % 60, s % 60].map(v => pad(v)).join(':'); }
    function status(message, kind = '') { if (!ui) return; const e = ui.querySelector('[data-status]'); e.textContent = message; e.className = 'kxb-status ' + kind; }
    function progress(done, total) {
        const p = total ? Math.round(done / total * 100) : 0;
        if (ui) ui.querySelector('[data-bar]').style.width = p + '%';
        if (running) setTitle(`⏳ ${p}%`);
    }
    // Tiêu đề tab trình duyệt cho biết tiến độ khi anh chuyển sang tab khác
    let baseTitle = null;
    function setTitle(prefix) {
        if (baseTitle === null) baseTitle = document.title.replace(/^(⏳ \d+%|✅ Xong|⚠️ Lỗi) · /, '');
        document.title = prefix ? `${prefix} · ${baseTitle}` : baseTitle;
    }
    function askNotify() { try { if (config.notify !== false && 'Notification' in window && Notification.permission === 'default') Notification.requestPermission(); } catch { /* bỏ qua */ } }
    function beep(ok) {
        try {
            const A = window.AudioContext || window.webkitAudioContext; if (!A) return;
            const ctx = new A(), notes = ok ? [880, 1175] : [440, 330];
            notes.forEach((f, i) => { const o = ctx.createOscillator(), g = ctx.createGain(); o.frequency.value = f; o.connect(g); g.connect(ctx.destination);
                const t0 = ctx.currentTime + i * 0.18; g.gain.setValueAtTime(0.0001, t0); g.gain.exponentialRampToValueAtTime(0.25, t0 + 0.02); g.gain.exponentialRampToValueAtTime(0.0001, t0 + 0.16); o.start(t0); o.stop(t0 + 0.17); });
            setTimeout(() => ctx.close(), 800);
        } catch { /* bỏ qua */ }
    }
    // Báo xong: chỉ khi phiên chạy trên 20 giây (đổ nhiều ngày) — tiếng "ting", thông báo góc màn hình, tiêu đề tab
    function doneSignal(ok, msg, ms) {
        setTitle(ok ? '✅ Xong' : '⚠️ Lỗi');
        const clear = () => { if (!running) setTitle(''); };
        if (document.hasFocus()) setTimeout(clear, 4000); else window.addEventListener('focus', () => setTimeout(clear, 1500), { once: true });
        if (ms < 20000 || config.notify === false) return;
        beep(ok);
        try { if (!document.hasFocus() && 'Notification' in window && Notification.permission === 'granted') new Notification(ok ? 'AutoBI · Đổ số xong' : 'AutoBI · Có lỗi', { body: msg, tag: 'autobi-kxb' }); } catch { /* bỏ qua */ }
    }
    function stop() {
        if (!running) return;
        running.status = 'cancelled'; running.controller.abort();
        log('Đã dừng. Ngày đã lấy xong vẫn được giữ; phần đang lấy dở bị bỏ.'); status('Đã dừng', 'warn');
    }
    async function withSession(mode, job) {
        invariant(!running, 'Có phiên đang chạy');
        askNotify();                                   // hỏi quyền thông báo ngay lúc bấm nút (Chrome cần thao tác của người dùng)
        await storeReady; invariant(!running, 'Có phiên đang chạy');
        const session = { id: uuid(), mode, started: Date.now(), status: 'running', controller: new AbortController() };
        running = session; ui.classList.add('busy');
        const timer = setInterval(() => { ui.querySelector('[data-timer]').textContent = clockText(session.started); }, 250);
        log(`Bắt đầu ${({ sales: 'đổ xuất bán', balance: 'đổ cân hàng', prev: 'đổ xuất bán kỳ so sánh', crm: 'tra tồn CRM', crmBal: 'tìm nguồn hàng CRM', crmStores: 'lấy danh sách siêu thị CRM', crmReq: 'check xin hàng', crmNear: 'tìm tỉnh lân cận', crmAlt: 'tìm màu / dung lượng khác' })[mode] || 'đổ tồn kho'}`); status('Đang kiểm tra quyền…'); progress(0, 1);
        try {
            await authCheck(session); check(session); log(`Quyền hợp lệ: ${auth.user} — ${auth.name}`);
            const msg = await job(session); check(session);
            session.status = 'completed'; progress(1, 1); status(msg || 'Hoàn tất', /lỗi|chưa/.test(msg || '') ? 'warn' : 'ok'); log(msg || 'Hoàn tất');
            doneSignal(true, msg || 'Hoàn tất', Date.now() - session.started);
        } catch (e) {
            if (session.status !== 'cancelled') { session.status = 'error'; status(e.message, 'err'); log(e.message, 'error'); doneSignal(false, e.message, Date.now() - session.started); }
        } finally {
            clearInterval(timer); ui.querySelector('[data-timer]').textContent = clockText(session.started);
            running = null; ui.classList.remove('busy');
            if (session.status === 'cancelled') setTitle('');
        }
        return session.status;
    }
    function selectedRange() { return validateRange(ui.querySelector('[data-from]').value, ui.querySelector('[data-to]').value); }

    // fetchRange: kỳ cần lấy (mặc định kỳ đang chọn); màn hình vẫn hiện kỳ đang chọn
    async function runSales(session, refetchAll, fetchRange) {
        const range = fetchRange || selectedRange(), today = isoDate(new Date());
        const shown = () => { try { return selectedRange(); } catch { return range; } };
        const fh = config.freshHours ?? 2;
        // Tự đổ hôm qua / Đổ cùng kỳ: CHỈ lấy đúng các ngày trong kỳ đó, không kéo theo ngày còn treo ở kỳ khác
        const need = daysToFetch(loadBook(range).book, range, shopsKey(), today, refetchAll, config.returnDays, SALES_SCHEMA, Date.now(), fh)
            .filter(d => !fetchRange || (d >= range.from && d <= range.to));
        const skipped = daysIn(range).filter(d => d <= today && !need.includes(d)).length;
        log(`Kỳ ${toBI(range.from)}–${toBI(range.to)}: cần lấy ${need.length} ngày, bỏ qua ${skipped} ngày (đã chốt${fh ? ` hoặc vừa lấy dưới ${fh} giờ` : ''})`);
        const done = [], failed = [];
        for (let i = 0; i < need.length; i++) {
            check(session); invariant(detectUser() === auth.user, 'Tài khoản đã thay đổi');
            const left = need.length - i, eta = Math.ceil(left * 13 / 60);
            const head = `Xuất bán ${i + 1}/${need.length}: ngày ${toBI(need[i])} · còn khoảng ${eta} phút (BI cho tối đa 5 lần/phút)`;
            status(head); progress(i, need.length);
            try {
                const r = await salesDay(need[i], session, sec => status(`${head} · chờ lượt BI ${sec} giây`));
                storeDay(r); done.push(r);
                log(`${toBI(r.day)}: ${r.record.rows} dòng BI, tính ${r.record.quantity} SL${r.record.pending ? `, ${r.record.pending} đơn treo` : ''}${r.record.returned ? `, bỏ ${r.record.returned} đơn khách nhập trả` : ''}`);
            } catch (e) {
                if (e.code === 'CANCELLED') throw e;
                failed.push(need[i]); log(`${toBI(need[i])}: ${e.message} — lần sau tự lấy lại`, 'error');
                if (e.code !== 'TIMEOUT' && /ngoài cấu hình|ngoài kỳ|định nghĩa|tham số/.test(e.message)) throw e;
            }
            view.sales = salesView(shown()); renderSales(); renderBalance();
        }
        view.sales = salesView(shown()); renderSales(); renderBalance();
        return failed.length ? `Xong ${done.length} ngày, ${failed.length} ngày lỗi: ${failed.map(toBI).join(', ')} — bấm "Đổ xuất bán" lần nữa sẽ tự lấy lại` : `Hoàn tất: lấy ${done.length} ngày, bỏ qua ${skipped} ngày đã chốt / vừa lấy`;
    }
    async function runInventory(session) {
        const shops = config.shops.filter(s => view.invShops.has(keyCode(s.code)));
        invariant(shops.length, 'Chọn ít nhất 1 siêu thị để đổ tồn kho');
        const records = [], transit = [], withTransit = !!config.invTransit;
        for (let i = 0; i < shops.length; i++) {
            check(session); invariant(detectUser() === auth.user, 'Tài khoản đã thay đổi');
            status(`Tồn kho ${i + 1}/${shops.length}: ${shops[i].code} — ${shops[i].name}${withTransit ? ' (kèm hàng đang chuyển)' : ''}`); progress(i, shops.length);
            const x = await inventoryShop(shops[i], session, withTransit);
            records.push(...x.records); if (x.transit) transit.push(...x.transit);
        }
        // chỉ công bố khi TẤT CẢ shop đã chọn xong
        // Gộp: chỉ thay số của siêu thị vừa đổ, giữ nguyên số các siêu thị khác (kèm giờ đổ riêng từng siêu thị)
        const now = new Date().toISOString(), fetched = new Set(shops.map(s => keyCode(s.code)));
        const prev = view.inv || { records: [], shops: [], shopTimes: {} };
        const shopTimes = { ...(prev.shopTimes || Object.fromEntries((prev.shops || []).map(c => [c, prev.capturedAt]))) };
        fetched.forEach(c => shopTimes[c] = now);
        const prevTransitShops = (prev.transitShops || []).filter(c => !fetched.has(c));
        view.inv = { noTransit: true, capturedAt: now, shops: [...new Set([...(prev.shops || []), ...fetched])], shopTimes,
            records: prev.records.filter(r => !fetched.has(r.shop)).concat(records),
            transit: (prev.transit || []).filter(r => !fetched.has(r.shop)).concat(transit),
            transitShops: withTransit ? [...prevTransitShops, ...fetched] : prevTransitShops };
        try { save('lastInventory', view.inv); } catch { log('Không lưu được tồn kho vào bộ nhớ (quá lớn); vẫn xem được tới khi tải lại trang', 'error'); }
        renderInventoryFilters(); renderInventory(); renderBalance();
        return `Hoàn tất: ${shops.length} siêu thị, ${fmt(records.length)} dòng tồn`;
    }
    function salesView(range) {
        const { book, lines } = loadBook(range);
        const period = inPeriod(lines, range, config.basis);
        const today = isoDate(new Date());
        const days = daysIn(range).filter(d => d <= today).map(d => ({ day: d, ...(book[d] || { status: 'Chưa lấy' }) }));
        const pend = days.flatMap(d => (d.pendingLines || [])), ret = days.flatMap(d => (d.returnedLines || []));
        const oldSchema = days.filter(d => d.status !== 'Chưa lấy' && d.schema !== SALES_SCHEMA).length;
        // Cùng kỳ tháng trước (chỉ đọc số đã lưu, không gọi BI)
        const prevRange = compareRange(range, config.compareMode), pb = loadBook(prevRange);
        const prevMissing = daysIn(prevRange).filter(d => d <= today && (!pb.book[d] || pb.book[d].schema !== SALES_SCHEMA)).length;
        return { range, compareMode: config.compareMode === 'week' ? 'week' : 'month', basis: config.basis, days, missing: days.filter(d => d.status === 'Chưa lấy').length, oldSchema, allLines: period, pendingLines: pend, returnedLines: ret,
            prevRange, prevMissing, prevAllLines: inPeriod(pb.lines, prevRange, config.basis) };
    }

    /* ---------- Định dạng & bảng ---------- */
    const fmt = v => new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 2 }).format(v);
    const pct = v => fmt(Math.round(v * 10) / 10) + '%';
    const mil = v => fmt(Math.round(v / 1e4) / 100) + ' tr';
    const shopName = code => config.shops.find(s => keyCode(s.code) === keyCode(code))?.name || code;
    function el(tag, text, parent, cls) { const e = document.createElement(tag); if (text !== undefined && text !== null) e.textContent = text; if (cls) e.className = cls; if (parent) parent.append(e); return e; }
    function table(parent, headers, rows, { pageSize = 200, num = [], total = false, rowClass = null, title = '', onRow = null, rowTitle = '' } = {}) {
        const wrap = el('div', undefined, parent, 'kxb-twrap');
        const bar = el('div', undefined, wrap, 'kxb-tbar');
        const info = el('span', title, bar, 'kxb-muted');
        const totalRow = total && rows.length ? rows[rows.length - 1] : null, base = totalRow ? rows.slice(0, -1) : rows;
        let order = base.map((_, i) => i), sortCol = -1, sortDir = 0, page = 0;
        const copyBtn = el('button', '📋 Chép bảng', bar, 'mini'); copyBtn.type = 'button'; copyBtn.title = 'Chép bảng (đúng thứ tự đang xem) để dán vào Excel / Zalo';
        copyBtn.onclick = safely(() => copyText(toTsv(headers, [...order.map(i => base[i]), ...(totalRow ? [totalRow] : [])]), `Đã chép ${fmt(base.length)} dòng — dán vào Excel bằng Ctrl+V`));
        const box = el('div', undefined, wrap, 'kxb-table');
        const t = el('table', undefined, box), tr = el('tr', undefined, el('thead', undefined, t));
        const ths = headers.map((h, i) => {
            const th = el('th', h, tr, (num.includes(i) ? 'n ' : '') + 'sort'); th.title = 'Bấm để sắp xếp';
            th.onclick = () => {
                const first = num.includes(i) ? -1 : 1;
                if (sortCol !== i) { sortCol = i; sortDir = first; } else if (sortDir === first) sortDir = -first; else { sortCol = -1; sortDir = 0; }
                const key = j => cellNumber(base[j][i]);
                order = base.map((_, j) => j);
                if (sortCol >= 0) {
                    const isNum = base.some(r => cellNumber(r[i]) !== null);
                    order.sort((a, b) => {
                        if (isNum) { const x = key(a), y = key(b); if (x === null && y === null) return a - b; if (x === null) return 1; if (y === null) return -1; return (x - y) * sortDir || a - b; }
                        return String(base[a][i] ?? '').localeCompare(String(base[b][i] ?? ''), 'vi', { numeric: true }) * sortDir || a - b;
                    });
                }
                ths.forEach((x, k) => x.dataset.dir = k === sortCol ? (sortDir > 0 ? '▲' : '▼') : '');
                page = 0; draw();
            };
            return th;
        });
        const body = el('tbody', undefined, t);
        const pager = base.length > pageSize ? el('div', undefined, wrap, 'kxb-pager') : null;
        if (!title) info.textContent = `${fmt(base.length)} dòng`;
        const draw = () => {
            body.replaceChildren();
            const slice = order.slice(page * pageSize, (page + 1) * pageSize);
            slice.forEach(oi => {
                const line = el('tr', undefined, body, rowClass ? rowClass(oi) || '' : ''); base[oi].forEach((c, i) => el('td', String(c ?? ''), line, num.includes(i) ? 'n' : ''));
                if (onRow) { line.classList.add('click'); line.title = rowTitle || 'Bấm để mở thẻ nhân viên'; line.onclick = () => onRow(oi); }
            });
            if (totalRow && (page + 1) * pageSize >= base.length) { const line = el('tr', undefined, body, 'tot'); totalRow.forEach((c, i) => el('td', String(c ?? ''), line, num.includes(i) ? 'n' : '')); }
            if (pager) {
                pager.replaceChildren();
                const prev = el('button', '← Trước', pager); prev.disabled = page === 0; prev.onclick = () => { page--; draw(); };
                el('span', `Trang ${page + 1}/${Math.ceil(base.length / pageSize)} · ${fmt(base.length)} dòng`, pager);
                const next = el('button', 'Sau →', pager); next.disabled = (page + 1) * pageSize >= base.length; next.onclick = () => { page++; draw(); };
            }
        };
        draw();
        if (!rows.length) { el('div', 'Không có dòng nào', box, 'kxb-empty'); copyBtn.disabled = true; }
        return wrap;
    }
    // Biểu đồ doanh thu theo ngày (cột) + kỳ so sánh (chấm), SVG không cần thư viện
    function dailyChart(parent, r) {
        const days = daysIn(r.range).filter(d => d <= isoDate(new Date()));
        if (days.length < 2) return;
        const by = new Map(), pv = new Map(), dayOf = l => (r.basis === 'shipped' && l.shipped ? l.shipped : l.created);
        r.lines.forEach(l => by.set(dayOf(l), (by.get(dayOf(l)) || 0) + l.qty * l.price));
        const prevDays = daysIn(r.prevRange);
        if (!r.prevMissing) (r.prevLines || []).forEach(l => pv.set(dayOf(l), (pv.get(dayOf(l)) || 0) + l.qty * l.price));
        const vals = days.map(d => by.get(d) || 0), pvals = r.prevMissing ? [] : days.map((_, i) => prevDays[i] ? pv.get(prevDays[i]) || 0 : null);
        const max = Math.max(1, ...vals, ...pvals.filter(v => v != null)), W = 1000, H = 150, P = 22, bw = (W - P * 2) / days.length;
        const NS = 'http://www.w3.org/2000/svg', svg = document.createElementNS(NS, 'svg');
        svg.setAttribute('viewBox', `0 0 ${W} ${H + 20}`); svg.setAttribute('class', 'kxb-chart'); svg.setAttribute('preserveAspectRatio', 'none');
        const add = (tag, attrs, text) => { const e = document.createElementNS(NS, tag); Object.entries(attrs).forEach(([k, v]) => e.setAttribute(k, v)); if (text != null) e.textContent = text; svg.append(e); return e; };
        const avg = vals.reduce((a, v) => a + v, 0) / vals.length, ya = H - avg / max * (H - 14);
        days.forEach((d, i) => {
            const h = vals[i] / max * (H - 14), x = P + i * bw + bw * 0.15, wk = weekdayOf(d) >= 5;
            const rect = add('rect', { x, y: H - h, width: bw * 0.7, height: Math.max(h, vals[i] ? 1 : 0), rx: 2, class: wk ? 'b wk' : 'b' });
            const t = document.createElementNS(NS, 'title'); t.textContent = `${toBI(d)} (${WEEKDAYS[weekdayOf(d)]}): ${mil(vals[i])}${pvals[i] != null ? ` · kỳ so sánh ${mil(pvals[i])}` : ''}`; rect.append(t);
            if (pvals[i] != null) add('circle', { cx: P + i * bw + bw / 2, cy: H - pvals[i] / max * (H - 14), r: 3, class: 'pv' });
            if (days.length <= 31 && (days.length <= 16 || i % 2 === 0)) add('text', { x: P + i * bw + bw / 2, y: H + 14, 'text-anchor': 'middle', class: 'lb' }, d.slice(8));
        });
        add('line', { x1: P, x2: W - P, y1: ya, y2: ya, class: 'avg' });
        add('text', { x: W - P, y: ya - 4, 'text-anchor': 'end', class: 'lb' }, `TB ${mil(avg)}/ngày`);
        const box = el('div', undefined, parent, 'kxb-chartbox');
        el('div', `Doanh thu theo ngày · cột cam = T7/CN${r.prevMissing ? '' : ' · chấm = ' + (r.compareMode === 'week' ? '7 ngày trước' : 'cùng kỳ tháng trước')} · rê chuột vào cột để xem số`, box, 'kxb-muted');
        box.append(svg);
    }
    function kpis(parent, items) { const box = el('div', undefined, parent, 'kxb-kpis'); items.forEach(([l, v, s, d], i) => { const k = el('div', undefined, box, i ? '' : 'main'); el('small', l, k); el('b', v, k); if (s) el('span', s, k); if (d && d.text) el('span', d.text, k, 'dl ' + (d.cls || '')); }); }
    // ▲▼ so cùng kỳ: chỉ hiện khi số cùng kỳ đã đủ ngày
    function delta(cur, prev, r) {
        if (!r || r.prevMissing) return null;
        if (!prev) return cur ? { text: `mới so ${toBI(r.prevRange.from).slice(0, 5)}–${toBI(r.prevRange.to).slice(0, 5)}`, cls: 'up' } : null;
        const p = (cur - prev) / prev * 100;
        return { text: `${p >= 0 ? '▲' : '▼'} ${fmt(Math.abs(Math.round(p)))}% so ${toBI(r.prevRange.from).slice(0, 5)}–${toBI(r.prevRange.to).slice(0, 5)}`, cls: p >= 0 ? 'up' : 'down' };
    }
    const isPhoneLine = l => categoryText(l.category) === BAL_CATEGORY;
    const attachKey = l => categoryText(l.category) + ' · ' + (l.group || '(không rõ nhóm)');
    function attachData(lines) {
        const set = new Set(config.attachGroups || []);
        return attachStats(lines, isPhoneLine, l => !set.size || set.has(attachKey(l)));
    }
    function subtabs(parent, list, current, onPick) {
        const bar = el('div', undefined, parent, 'kxb-subtabs');
        list.forEach(([k, l]) => { const b = el('button', l, bar); b.type = 'button'; b.classList.toggle('on', k === current); b.onclick = () => onPick(k); });
    }

    // Ô thả xuống chọn nhiều (danh sách có ô tích), giữ mở khi chọn
    function multiDrop(box, id, items, set, allLabel, onChange) {
        box.replaceChildren();
        const dd = el('details', undefined, box, 'kxb-drop'); dd.open = view.openDrop === id;
        dd.ontoggle = () => { if (dd.open) view.openDrop = id; else if (view.openDrop === id) view.openDrop = ''; };
        const picked = items.filter(i => set.has(i.key));
        el('summary', picked.length ? picked.map(i => i.name).join(', ') : allLabel, dd).classList.toggle('on', picked.length > 0);
        const list = el('div', undefined, dd, 'list');
        const row = (label, checked, onClick) => { const l = el('label', undefined, list); const c = el('input', undefined, l); c.type = 'checkbox'; c.checked = checked; c.onchange = onClick; el('span', label, l); };
        row(allLabel, !set.size, () => { set.clear(); onChange(); });
        items.forEach(i => row(i.label, set.has(i.key), () => { set.has(i.key) ? set.delete(i.key) : set.add(i.key); onChange(); }));
        if (!items.length) el('div', 'Chưa có dữ liệu', list, 'kxb-muted');
    }
    const brandKey = b => norm(b) || '(không rõ)';

    /* ---------- Tab Xuất bán ---------- */
    const categoryText = c => clean(c).replace(/^\d+\s*-\s*/, '') || 'Chưa rõ';
    function salesByCategory(lines, codes) {
        const m = new Map(), byShop = {}; let qty = 0, rev = 0;
        for (const l of lines) {
            const k = l.category + '|' + l.group, x = m.get(k) || { category: categoryText(l.category), group: l.group, byShop: {}, qty: 0, rev: 0 };
            x.byShop[l.shop] = (x.byShop[l.shop] || 0) + l.qty; x.qty += l.qty; x.rev += l.qty * l.price; m.set(k, x);
            byShop[l.shop] = (byShop[l.shop] || 0) + l.qty; qty += l.qty; rev += l.qty * l.price;
        }
        const catRev = new Map(); for (const x of m.values()) catRev.set(x.category, (catRev.get(x.category) || 0) + x.rev);
        const rows = [...m.values()].sort((a, b) => catRev.get(b.category) - catRev.get(a.category) || a.category.localeCompare(b.category, 'vi') || b.rev - a.rev);   // ngành doanh thu cao lên trên
        return { rows, byShop, qty, rev };
    }
    // Áp bộ chọn siêu thị + tính lại tổng hợp
    function salesData() {
        const r = view.sales; if (!r) return null;
        const shops = view.salesShops, conds = view.salesConds, cat = view.salesCat;
        const keep = l => shops.has(l.shop) && (!conds.size || conds.has(conditionText(l.condition))) && (!cat || categoryText(l.category) === cat) && (!view.salesBrands.size || view.salesBrands.has(brandKey(l.brand)));
        const lines = r.allLines.filter(keep), prevLines = (r.prevAllLines || []).filter(keep);
        return { ...r, lines, summary: summarizeSales(lines, r.basis), prevLines, prevSummary: summarizeSales(prevLines, r.basis), pendingLines: r.pendingLines.filter(keep), returnedLines: (r.returnedLines || []).filter(keep) };
    }
    function staffMatrix(s) {
        const brands = s.brands.map(b => b.label);
        return { brands, rows: s.staff.map(st => {
            const m = Object.fromEntries(st.brands.map(b => [b.label, b.quantity]));
            return { staff: st, cells: brands.map(b => m[b] || 0) };
        }) };
    }
    /* ---------- Thẻ nhân viên (V2.1): bấm tên nhân viên ở bảng Nhân viên × hãng / Bán kèm ---------- */
    const staffId = creator => clean(creator).match(/^(\d+)\s*[-–]/)?.[1] || norm(creator);
    function openStaffCard(shop, creator) {
        const r = view.sales; if (!r) return;
        const id = staffId(creator), mine = l => l.shop === shop && staffId(l.creator) === id;
        const all = r.allLines.filter(mine), prev = (r.prevAllLines || []).filter(mine);
        const shopAll = r.allLines.filter(l => l.shop === shop);
        const lines = all.filter(l => (!view.salesCat || categoryText(l.category) === view.salesCat) && (!view.salesConds.size || view.salesConds.has(conditionText(l.condition))) && (!view.salesBrands.size || view.salesBrands.has(brandKey(l.brand))));
        const prevLines = prev.filter(l => (!view.salesCat || categoryText(l.category) === view.salesCat) && (!view.salesConds.size || view.salesConds.has(conditionText(l.condition))) && (!view.salesBrands.size || view.salesBrands.has(brandKey(l.brand))));
        const s = summarizeSales(lines), ps = summarizeSales(prevLines), sumAll = summarizeSales(all);
        const shopRev = shopAll.reduce((a, l) => a + l.qty * l.price, 0);
        const orders = new Set(all.map(l => l.order)).size, att = attachData(all).total, shopAtt = attachData(shopAll).total;
        // xếp hạng trong siêu thị theo doanh thu (cùng bộ lọc)
        const shopSum = summarizeSales(shopAll.filter(l => (!view.salesCat || categoryText(l.category) === view.salesCat) && (!view.salesConds.size || view.salesConds.has(conditionText(l.condition))) && (!view.salesBrands.size || view.salesBrands.has(brandKey(l.brand)))));
        const rank = shopSum.staff.slice().sort((a, b) => b.revenue - a.revenue).findIndex(x => x.employee === id) + 1;
        document.getElementById('kxb-card')?.remove();
        const back = el('div', undefined, ui, 'kxb-cardback'); back.id = 'kxb-card';
        const box = el('div', undefined, back, 'kxb-card');
        const close = () => back.remove();
        back.addEventListener('mousedown', e => { if (e.target === back) close(); });
        const head = el('div', undefined, box, 'top');
        el('strong', `👤 ${creator}`, head); el('span', `${shopName(shop)} · kỳ ${toBI(r.range.from)}–${toBI(r.range.to)}${view.salesCat ? ' · ngành ' + view.salesCat : ''}`, head, 'kxb-muted');
        el('span', undefined, head, 'sp');
        const cp = el('button', '📋 Chép nhận xét', head); cp.type = 'button';
        const x = el('button', '✕', head); x.type = 'button'; x.onclick = close;
        const body = el('div', undefined, box, 'body');
        const dl = (c, p) => r.prevMissing ? null : delta(c, p, r);
        kpis(body, [['Doanh thu', mil(s.revenue), `SL ${fmt(s.quantity)}`, dl(s.revenue, ps.revenue)],
            ['Xếp hạng tại siêu thị', rank ? `#${rank}/${shopSum.staff.length}` : '—', `${pct(shopRev ? sumAll.revenue / shopRev * 100 : 0)} doanh thu siêu thị`],
            ['Số đơn', fmt(orders), `giá trị TB ${mil(orders ? sumAll.revenue / orders : 0)}/đơn`],
            ['Bán kèm ĐT', pct(att.rate), `${fmt(att.attached)}/${fmt(att.phoneOrders)} đơn · TB siêu thị ${pct(shopAtt.rate)}`, att.phoneOrders >= 3 && Math.abs(att.rate - shopAtt.rate) >= 0.5 ? { text: att.rate > shopAtt.rate ? 'cao hơn TB siêu thị' : 'thấp hơn TB siêu thị', cls: att.rate > shopAtt.rate ? 'up' : 'down' } : null],
            ['Kỳ so sánh', r.prevMissing ? 'chưa đủ số' : mil(ps.revenue), r.compareMode === 'week' ? '7 ngày trước' : 'cùng kỳ tháng trước']]);
        dailyChart(body, { ...r, lines, prevLines });
        const grid = el('div', undefined, body, 'kxb-grid2');
        const left = el('div', undefined, grid), right = el('div', undefined, grid);
        const cats = new Map(); all.forEach(l => { const k = categoryText(l.category), e = cats.get(k) || { q: 0, r: 0 }; e.q += l.qty; e.r += l.qty * l.price; cats.set(k, e); });
        table(left, ['Ngành hàng', 'SL', 'Doanh thu', '% DT'], [...cats.entries()].sort((a, b) => b[1].r - a[1].r).map(([k, v]) => [k, fmt(v.q), mil(v.r), pct(sumAll.revenue ? v.r / sumAll.revenue * 100 : 0)]), { num: [1, 2, 3], title: 'Theo ngành (mọi ngành)' });
        table(right, ['Hãng', 'SL', '% SL', 'Doanh thu'], s.brands.map(b => [b.label, fmt(b.quantity), pct(b.pctQty), mil(b.revenue)]), { num: [1, 2, 3], title: 'Theo hãng' });
        table(body, ['Sản phẩm', 'Hãng', 'SL', 'Doanh thu', 'IMEI'], s.products.slice(0, 30).map(p => [p.name, p.brand, fmt(p.quantity), mil(p.revenue), lines.filter(l => l.product === p.product && l.imei).map(l => l.imei).join(', ')]), { num: [2, 3], title: `Sản phẩm bán (top ${Math.min(30, s.products.length)}/${s.products.length})` });
        const hm = hourMatrix(all), hrs = new Map(); hm.cells.forEach(c => hrs.set(c.h, (hrs.get(c.h) || 0) + c.orders));
        const peak = [...hrs.entries()].sort((a, b) => b[1] - a[1]).slice(0, 3).map(([h, n]) => `${pad(h)}h (${n} đơn)`);
        el('div', `Giờ bán nhiều nhất: ${peak.join(' · ') || '—'}`, body, 'kxb-muted');
        cp.onclick = safely(() => copyText([`👤 ${creator.replace(/^\d+\s*[-–]\s*/, '')} — ${shopName(shop)} · ${toBI(r.range.from)}–${toBI(r.range.to)}`,
            `• Doanh thu ${mil(s.revenue)} (SL ${fmt(s.quantity)})${r.prevMissing ? '' : `, ${ps.revenue ? (s.revenue >= ps.revenue ? '▲ ' : '▼ ') + fmt(Math.abs(Math.round((s.revenue - ps.revenue) / ps.revenue * 100))) + '% so kỳ trước' : 'kỳ trước chưa bán'}`}`,
            `• Xếp hạng #${rank || '—'}/${shopSum.staff.length} tại siêu thị · ${fmt(orders)} đơn, TB ${mil(orders ? sumAll.revenue / orders : 0)}/đơn`,
            `• Bán kèm điện thoại ${pct(att.rate)} (${att.attached}/${att.phoneOrders} đơn) — TB siêu thị ${pct(shopAtt.rate)}`,
            `• Hãng bán nhiều: ${s.brands.slice(0, 3).map(b => `${b.label} ${fmt(b.quantity)}`).join(', ') || '—'}`,
            `• Giờ bán nhiều: ${peak.join(', ') || '—'}`].join('\n'), 'Đã chép nhận xét nhân viên'));
    }
    function renderSales() {
        const area = ui.querySelector('[data-sales-result]'); area.replaceChildren();
        const chipBox = ui.querySelector('[data-sales-shops]'); chipBox.replaceChildren();
        config.shops.forEach(sh => {
            const c = keyCode(sh.code), on = view.salesShops.has(c), b = el('button', `${sh.code} · ${sh.name}`, chipBox, 'chip'); b.type = 'button';
            b.classList.toggle('on', on); b.setAttribute('aria-pressed', on);
            b.onclick = () => { if (on && view.salesShops.size === 1) return; on ? view.salesShops.delete(c) : view.salesShops.add(c); renderSales(); };
        });
        const shopLines = (view.sales ? view.sales.allLines : []).filter(l => view.salesShops.has(l.shop));
        const chipRow = (box, set, keyOf, sortFirst, empty, src) => {
            box.replaceChildren();
            const by = new Map(); src.forEach(l => { const k = keyOf(l); by.set(k, (by.get(k) || 0) + l.qty); });
            [...set].forEach(k => { if (!by.has(k)) by.set(k, 0); });
            const list = [...by.keys()].sort((a, b) => (a === sortFirst ? -1 : b === sortFirst ? 1 : a.localeCompare(b, 'vi', { numeric: true })));
            const all = el('button', 'Tất cả', box, 'chip'); all.type = 'button'; all.classList.toggle('on', !set.size);
            all.onclick = () => { set.clear(); renderSales(); };
            list.forEach(k => {
                const on = set.has(k), b = el('button', `${k} · SL ${fmt(by.get(k))}`, box, 'chip'); b.type = 'button';
                b.classList.toggle('on', on); b.setAttribute('aria-pressed', on);
                b.onclick = () => { on ? set.delete(k) : set.add(k); renderSales(); };
            });
            if (!list.length) el('span', empty, box, 'kxb-muted');
        };
        // Ngành hàng: danh sách thả xuống, ngành doanh thu cao lên trên
        const catSel = ui.querySelector('[data-sales-cat]'), byCat = new Map();
        shopLines.filter(l => !view.salesConds.size || view.salesConds.has(conditionText(l.condition))).forEach(l => { const k = categoryText(l.category), x = byCat.get(k) || { q: 0, r: 0 }; x.q += l.qty; x.r += l.qty * l.price; byCat.set(k, x); });
        if (view.salesCat && !byCat.has(view.salesCat)) byCat.set(view.salesCat, { q: 0, r: 0 });
        const tq = [...byCat.values()].reduce((a, v) => a + v.q, 0);
        catSel.replaceChildren();
        const o0 = el('option', byCat.size ? `Tất cả ngành · SL ${fmt(tq)}` : 'Tất cả ngành', catSel); o0.value = '';
        [...byCat.entries()].sort((a, b) => b[1].r - a[1].r || b[1].q - a[1].q || a[0].localeCompare(b[0], 'vi')).forEach(([k, v]) => { const o = el('option', `${k} · SL ${fmt(v.q)}`, catSel); o.value = k; });
        catSel.value = view.salesCat;
        catSel.classList.toggle('on', !!view.salesCat);
        const setCat = v => { view.salesCat = v; try { save('salesCat', v); } catch { /* bỏ qua */ } renderSales(); };
        catSel.onchange = () => setCat(catSel.value);
        const phoneBtn = ui.querySelector('[data-phone-only]'), phoneOn = view.salesCat === BAL_CATEGORY;
        phoneBtn.classList.toggle('on', phoneOn); phoneBtn.setAttribute('aria-pressed', phoneOn);
        phoneBtn.onclick = () => setCat(phoneOn ? '' : BAL_CATEGORY);
        // Hãng: chọn nhiều, hãng doanh thu cao lên đầu (theo ngành + loại hàng đang chọn)
        const byBrand = new Map();
        shopLines.filter(l => (!view.salesCat || categoryText(l.category) === view.salesCat) && (!view.salesConds.size || view.salesConds.has(conditionText(l.condition))))
            .forEach(l => { const k = brandKey(l.brand), x = byBrand.get(k) || { key: k, name: l.brand || '(không rõ)', q: 0, r: 0 }; x.q += l.qty; x.r += l.qty * l.price; byBrand.set(k, x); });
        [...view.salesBrands].forEach(k => { if (!byBrand.has(k)) view.salesBrands.delete(k); });
        multiDrop(ui.querySelector('[data-sales-brands]'), 'sales-brand', [...byBrand.values()].sort((a, b) => b.r - a.r || b.q - a.q).map(b => ({ key: b.key, name: b.name, label: `${b.name} · SL ${fmt(b.q)}` })),
            view.salesBrands, `Tất cả hãng · SL ${fmt([...byBrand.values()].reduce((a, b) => a + b.q, 0))}`, renderSales);
        chipRow(ui.querySelector('[data-sales-conds]'), view.salesConds, l => conditionText(l.condition), 'Mới', 'Đổ xuất bán để có danh sách loại hàng',
            shopLines.filter(l => !view.salesCat || categoryText(l.category) === view.salesCat));   // SL loại hàng theo ngành đang chọn
        const r = salesData();
        if (!r) { el('div', 'Chọn kỳ rồi bấm "Đổ xuất bán". Ngày đã lấy được lưu lại, lần sau chỉ lấy ngày còn thiếu. Bấm "Xem số đã lưu" để xem ngay không cần gọi BI.', area, 'kxb-empty'); return; }
        const s = r.summary, codes = config.shops.map(x => keyCode(x.code)).filter(c => view.salesShops.has(c));
        if (r.missing) el('div', `Còn ${r.missing} ngày chưa lấy trong kỳ ${toBI(r.range.from)}–${toBI(r.range.to)} — số chưa đủ. Bấm "Đổ xuất bán" để lấy tiếp.`, area, 'kxb-warn');
        if (r.oldSchema) el('div', `${r.oldSchema} ngày lưu bằng bản cũ (chỉ ngành Điện thoại). Bấm "Đổ xuất bán" để lấy lại các ngày này với tất cả ngành hàng.`, area, 'kxb-warn');
        if (view.salesCat) el('div', `Đang lọc ngành: ${view.salesCat} — chọn "Tất cả ngành" để bỏ lọc.`, area, 'kxb-warn');
        if (view.salesBrands.size) el('div', `Đang lọc hãng: ${[...byBrand.values()].filter(b => view.salesBrands.has(b.key)).map(b => b.name).join(', ')} — tích "Tất cả hãng" để bỏ lọc.`, area, 'kxb-warn');
        if (view.salesConds.size) el('div', `Đang lọc loại hàng: ${[...view.salesConds].join(', ')} — bấm "Tất cả" để bỏ lọc.`, area, 'kxb-warn');
        const today = isoDate(new Date()), pOrders = pendingOrders(r.pendingLines, today), pendOrders = pOrders.length, oldPend = pOrders.filter(o => o.age >= 2).length;
        const retOrders = new Set(r.returnedLines.map(l => l.order)).size, ps = r.prevSummary;
        if (r.prevMissing) {
            const w = el('div', `Kỳ so sánh (${r.compareMode === 'week' ? '7 ngày trước' : 'cùng kỳ tháng trước'} ${toBI(r.prevRange.from)}–${toBI(r.prevRange.to)}) còn ${r.prevMissing} ngày chưa lấy — chưa so được ▲▼. `, area, 'kxb-warn');
            const b = el('button', 'Đổ kỳ so sánh', w, 'idle-only'); b.type = 'button';
            b.onclick = safely(() => { validateShops(config.shops); return withSession('prev', ss => runSales(ss, false, r.prevRange)); });
        }
        // Bán kèm & dự kiến: tính trên các siêu thị đang chọn (không theo bộ lọc ngành/hãng/loại hàng)
        const shopAll = r.allLines.filter(l => view.salesShops.has(l.shop)), att = attachData(shopAll);
        const attPrev = r.prevMissing ? null : attachData((r.prevAllLines || []).filter(l => view.salesShops.has(l.shop)));
        const proj = projectMonth(r.lines.filter(l => (r.basis === 'shipped' && l.shipped ? l.shipped : l.created) < today).reduce((a, l) => a + l.qty * l.price, 0), r.range, today);
        const projItem = proj ? ['Dự kiến cuối tháng', mil(proj.value), `theo TB ${proj.elapsed} ngày đã qua × ${proj.dim} ngày`] : null;
        kpis(area, [['Đã bán', `SL ${fmt(s.quantity)}`, mil(s.revenue), delta(s.revenue, ps.revenue, r)],
            ...codes.map(c => { const x = s.shops.find(z => z.code === c) || { quantity: 0, revenue: 0 }, y = ps.shops.find(z => z.code === c) || { revenue: 0 }; return [shopName(c), `SL ${fmt(x.quantity)}`, mil(x.revenue), delta(x.revenue, y.revenue, r)]; }),
            ['Tỷ lệ bán kèm ĐT', pct(att.total.rate), `${fmt(att.total.attached)}/${fmt(att.total.phoneOrders)} đơn điện thoại`,
                attPrev && attPrev.total.phoneOrders ? { text: `${att.total.rate >= attPrev.total.rate ? '▲' : '▼'} ${fmt(Math.abs(Math.round(att.total.rate - attPrev.total.rate)))} điểm so cùng kỳ`, cls: att.total.rate >= attPrev.total.rate ? 'up' : 'down' } : null],
            ...(projItem ? [projItem] : []),
            ['Đơn treo', fmt(pendOrders), 'chưa xuất / chưa giao', oldPend ? { text: `${fmt(oldPend)} đơn treo từ 2 ngày`, cls: 'down' } : null],
            ['Khách nhập trả', fmt(retOrders), `đã bỏ SL ${fmt(r.returnedLines.reduce((a, l) => a + l.qty, 0))} · ${mil(r.returnedLines.reduce((a, l) => a + l.qty * l.price, 0))}`]]);
        el('div', `Kỳ ${toBI(r.range.from)}–${toBI(r.range.to)} · Kho tạo · ngành: ${view.salesCat || 'tất cả'} · tính theo ${r.basis === 'shipped' ? 'ngày xuất' : 'ngày tạo'} · dòng Đã xuất – Đã giao – Chưa hủy · loại hàng: ${view.salesConds.size ? [...view.salesConds].join(', ') : 'tất cả'} · bỏ cả đơn khách nhập trả (kiểm tra lại ${config.returnDays} ngày gần nhất) · doanh thu = Giá bán × SL (gồm VAT)`, area, 'kxb-muted');
        subtabs(area, [['category', 'Theo ngành hàng'], ['brand', 'Theo hãng'], ['staff', 'Nhân viên × hãng'], ['attach', 'Bán kèm'], ['staffProduct', 'Nhân viên × sản phẩm'], ['product', 'Sản phẩm'], ['daily', 'Theo ngày'], ['hour', 'Theo giờ'], ['channel', 'Thanh toán & giao'], ['pending', `Đơn treo (${pendOrders})`], ['returned', `Nhập trả (${retOrders})`], ['detail', 'Chi tiết']], view.salesTab, k => { view.salesTab = k; try { save('salesTab', k); } catch { /* bỏ qua */ } renderSales(); });
        const pane = el('div', undefined, area);
        const n = (from, count) => Array.from({ length: count }, (_, i) => from + i);
        if (view.salesTab === 'hour') {
            // Giờ × thứ: tô đậm ô bán nhiều để xếp ca / giờ công
            const hm = hourMatrix(r.lines), metric = view.hourMetric || 'orders';
            const bar = el('div', undefined, pane, 'group');
            [['orders', 'Số đơn'], ['rev', 'Doanh thu'], ['avg', 'Đơn TB / ngày']].forEach(([k, l]) => { const b = el('button', l, bar, 'chip'); b.type = 'button'; b.classList.toggle('on', metric === k); b.onclick = () => { view.hourMetric = k; renderSales(); }; });
            el('div', `Theo giờ tạo đơn · ${hm.dayCount.map((c, i) => `${WEEKDAYS[i]}: ${c} ngày`).join(' · ')} · "Đơn TB / ngày" = số đơn ÷ số ngày thứ đó trong kỳ (so công bằng giữa các thứ)`, pane, 'kxb-muted');
            if (!hm.cells.length) el('div', 'Không có đơn nào trong kỳ / bộ lọc đang chọn', pane, 'kxb-empty');
            else {
                const val = c => !c ? 0 : metric === 'rev' ? c.rev : metric === 'avg' ? c.orders / Math.max(1, hm.dayCount[c.w]) : c.orders;
                const show = v => metric === 'rev' ? mil(v) : fmt(Math.round(v * 10) / 10);
                const get = (w, h) => hm.cells.find(c => c.w === w && c.h === h);
                const max = Math.max(...hm.cells.map(val), 1);
                const box = el('div', undefined, pane, 'kxb-table'), t = el('table', undefined, box, 'heat'), hr = el('tr', undefined, el('thead', undefined, t));
                el('th', 'Giờ', hr); WEEKDAYS.forEach(w => el('th', w, hr, 'n')); el('th', 'Cả tuần', hr, 'n');
                const tb = el('tbody', undefined, t);
                hm.hours.forEach(h => {
                    const line = el('tr', undefined, tb); el('td', `${pad(h)}:00–${pad(h)}:59`, line);
                    let sum = 0;
                    WEEKDAYS.forEach((_, w) => { const v = val(get(w, h)); sum += v; const td = el('td', v ? show(v) : '', line, 'n'); if (v) td.style.background = `rgba(8,127,140,${(0.08 + 0.72 * v / max).toFixed(2)})`; if (v / max > 0.55) td.style.color = '#fff'; });
                    el('td', show(sum), line, 'n b');
                });
                const top = hm.cells.slice().sort((a, b) => val(b) - val(a)).slice(0, 3).map(c => `${WEEKDAYS[c.w]} ${pad(c.h)}h (${show(val(c))})`);
                el('div', `Cao điểm: ${top.join(' · ')}`, pane, 'kxb-muted').style.marginTop = '8px';
            }
        }
        if (view.salesTab === 'channel') {
            // Hình thức thanh toán & giao hàng (lưu từ V2.0; ngày lấy bằng bản cũ hiện "Chưa có")
            const groupBy = keyOf => {
                const m = new Map(); let tot = 0;
                for (const l of r.lines) { const k = keyOf(l) || 'Chưa có (ngày lấy bằng bản cũ)'; const x = m.get(k) || { k, orders: new Set(), qty: 0, rev: 0 }; x.orders.add(l.order); x.qty += l.qty; x.rev += l.qty * l.price; m.set(k, x); tot += l.qty * l.price; }
                return { rows: [...m.values()].sort((a, b) => b.rev - a.rev), tot };
            };
            const cols = ['Số đơn', 'SL', 'Doanh thu', '% DT', 'Giá trị đơn TB'];
            for (const [title, keyOf] of [['Hình thức thanh toán', l => l.pay], ['Hình thức giao hàng', l => l.ship], ['Loại yêu cầu xuất', l => l.orderType]]) {
                const g = groupBy(keyOf);
                table(pane, [title, ...cols], g.rows.map(x => [x.k, fmt(x.orders.size), fmt(x.qty), mil(x.rev), pct(g.tot ? x.rev / g.tot * 100 : 0), mil(x.orders.size ? x.rev / x.orders.size : 0)]), { num: [1, 2, 3, 4, 5], title });
            }
            if (r.lines.some(l => !l.pay)) el('div', 'Các ngày lấy bằng bản trước V2.0 chưa có thông tin thanh toán / giao hàng — tích "Lấy lại cả ngày đã chốt" rồi "Đổ xuất bán" nếu cần đủ.', pane, 'kxb-muted');
        }
        if (view.salesTab === 'category') {
            const v = salesByCategory(r.lines, codes);
            table(pane, ['Ngành hàng', 'Nhóm hàng', ...codes.map(shopName), 'SL', 'Doanh thu', '% DT'],
                [...v.rows.map(x => [x.category, x.group, ...codes.map(c => x.byShop[c] ? fmt(x.byShop[c]) : ''), fmt(x.qty), mil(x.rev), pct(v.rev ? x.rev / v.rev * 100 : 0)]),
                 ['Tổng', '', ...codes.map(c => fmt(v.byShop[c] || 0)), fmt(v.qty), mil(v.rev), '100%']], { num: n(2, codes.length + 3), total: true });
        }
        if (view.salesTab === 'brand') table(pane, ['Hãng', ...codes.map(shopName), 'SL', '% SL', 'Doanh thu', '% DT', 'Giá TB'],
            [...s.brands.map(b => [b.label, ...codes.map(c => fmt(s.brandShop[norm(b.label) + '|' + c]?.quantity || 0)), fmt(b.quantity), pct(b.pctQty), mil(b.revenue), pct(b.pctRev), mil(b.quantity ? b.revenue / b.quantity : 0)]),
             ['Tổng', ...codes.map(c => fmt(s.shops.find(z => z.code === c)?.quantity || 0)), fmt(s.quantity), '100%', mil(s.revenue), '100%', mil(s.quantity ? s.revenue / s.quantity : 0)]], { num: n(1, codes.length + 5), total: true });
        if (view.salesTab === 'staff') {
            const m = staffMatrix(s);
            table(pane, ['Siêu thị', 'Nhân viên (người tạo)', ...m.brands, 'Tổng SL', '% shop', 'Doanh thu'],
                [...m.rows.map(x => [shopName(x.staff.shop), x.staff.label, ...x.cells.map(v => v ? fmt(v) : ''), fmt(x.staff.quantity), pct(x.staff.pctShop), mil(x.staff.revenue)]),
                 ['Tổng', '', ...s.brands.map(b => fmt(b.quantity)), fmt(s.quantity), '', mil(s.revenue)]], { num: n(2, m.brands.length + 3), total: true, title: '👤 Bấm vào một nhân viên để mở thẻ nhân viên', onRow: i => openStaffCard(m.rows[i].staff.shop, m.rows[i].staff.label) });
        }
        if (view.salesTab === 'attach') {
            // Chọn nhóm hàng tính là "kèm" (để trống = mọi dòng không phải điện thoại trong cùng đơn)
            const gq = new Map(); shopAll.filter(l => !isPhoneLine(l)).forEach(l => gq.set(attachKey(l), (gq.get(attachKey(l)) || 0) + l.qty));
            const aset = new Set(config.attachGroups || []);
            const pickBox = el('div', undefined, pane, 'group'); el('b', 'Tính là kèm', pickBox); const dropBox = el('span', undefined, pickBox);
            multiDrop(dropBox, 'attach-groups', [...new Set([...gq.keys(), ...aset])].sort((a, b) => (gq.get(b) || 0) - (gq.get(a) || 0) || a.localeCompare(b, 'vi')).map(k => ({ key: k, name: k, label: `${k} · SL ${fmt(gq.get(k) || 0)}` })),
                aset, 'Mọi nhóm không phải điện thoại', () => { config.attachGroups = [...aset]; save('config', config); renderSales(); });
            el('div', `Đơn có điện thoại · "có kèm" = cùng mã đơn có thêm ${aset.size ? 'nhóm đã chọn' : 'hàng ngoài ngành điện thoại'} · nhân viên = người tạo dòng điện thoại · theo các siêu thị đang chọn, không theo bộ lọc ngành / hãng / loại hàng · lựa chọn được lưu.`, pane, 'kxb-muted');
            const prevMap = new Map((attPrev?.staff || []).map(x => [x.shop + '|' + x.label, x]));
            const dd = d => d > 0 ? `▲ ${fmt(d)}` : d < 0 ? `▼ ${fmt(-d)}` : '=';
            const dq = (cur, x) => attPrev ? dd(cur - (prevMap.get(x.shop + '|' + x.label)?.phoneQty || 0)) : '';
            table(pane, ['Siêu thị', 'Nhân viên', 'Đơn ĐT', 'SL ĐT', 'Có kèm', 'Tỷ lệ kèm', 'SL kèm', 'DT kèm', 'SL ĐT so cùng kỳ'],
                [...att.staff.map(x => [shopName(x.shop), x.label, fmt(x.phoneOrders), fmt(x.phoneQty), fmt(x.attached), pct(x.rate), fmt(x.attachQty), mil(x.attachRev), dq(x.phoneQty, x)]),
                 ['Tổng', '', fmt(att.total.phoneOrders), fmt(att.total.phoneQty), fmt(att.total.attached), pct(att.total.rate), fmt(att.total.attachQty), mil(att.total.attachRev), attPrev ? dd(att.total.phoneQty - attPrev.total.phoneQty) : '']],
                { num: [2, 3, 4, 5, 6, 7, 8], total: true, rowClass: i => i < att.staff.length && att.staff[i].phoneOrders >= 3 && att.staff[i].rate < att.total.rate / 2 ? 'low' : '', onRow: i => openStaffCard(att.staff[i].shop, att.staff[i].label) });
            if (!attPrev) el('div', 'Cột "so cùng kỳ" trống vì kỳ so sánh chưa đủ ngày — bấm "Đổ kỳ so sánh" ở trên.', pane, 'kxb-muted');
            el('div', 'Nhóm hàng đi kèm (số đơn điện thoại có nhóm này)', pane, 'kxb-muted').style.marginTop = '10px';
            table(pane, ['Nhóm hàng', 'Số đơn', '% đơn ĐT', 'SL', 'Doanh thu'], att.groups.map(g => [g.group || '(không rõ)', fmt(g.orders), pct(att.total.phoneOrders ? g.orders / att.total.phoneOrders * 100 : 0), fmt(g.qty), mil(g.rev)]), { num: [1, 2, 3, 4] });
        }
        if (view.salesTab === 'staffProduct') {
            el('div', 'Bấm vào tên nhân viên để mở / đóng danh sách sản phẩm.', pane, 'kxb-muted');
            const imeis = new Map(); r.lines.forEach(l => { if (!l.imei) return; const k = l.creator + '|' + l.product; (imeis.get(k) || imeis.set(k, []).get(k)).push(l.imei); });
            for (const st of s.staff) {
                const box = el('details', undefined, pane, 'kxb-staff'); box.open = view.openStaff.has(st.label);
                box.ontoggle = () => { box.open ? view.openStaff.add(st.label) : view.openStaff.delete(st.label); };
                el('summary', `${shopName(st.shop)} · ${st.label} — SL ${fmt(st.quantity)} · ${mil(st.revenue)} · ${st.brands.map(b => `${b.label} ${fmt(b.quantity)}`).join(', ')}`, box);
                table(box, ['Sản phẩm', 'Hãng', 'Mã SP', 'SL', '% của NV', 'Doanh thu', 'IMEI'], st.products.map(v => [v.label, v.brand, v.product, fmt(v.quantity), pct(v.pctQty), mil(v.revenue), (imeis.get(st.label + '|' + v.product) || []).join(', ')]), { num: [3, 4, 5] });
            }
        }
        if (view.salesTab === 'product') {
            const ps2 = new Map(); r.lines.forEach(l => { const k = l.product + '|' + l.shop; ps2.set(k, (ps2.get(k) || 0) + l.qty); });
            table(pane, ['Sản phẩm', 'Hãng', 'Nhóm', 'Mã SP', ...codes.map(shopName), 'SL', '% SL', 'Doanh thu'],
                s.products.map(p => [p.name, p.brand, p.group, p.product, ...codes.map(c => fmt(ps2.get(p.product + '|' + c) || 0)), fmt(p.quantity), pct(p.pctQty), mil(p.revenue)]), { num: n(4, codes.length + 3) });
        }
        if (view.salesTab === 'daily') {
            const by = new Map(); r.lines.forEach(l => { const d = r.basis === 'shipped' && l.shipped ? l.shipped : l.created; const e = by.get(d) || {}; e[l.shop] = (e[l.shop] || 0) + l.qty; e.q = (e.q || 0) + l.qty; e.r = (e.r || 0) + l.qty * l.price; by.set(d, e); });
            table(pane, ['Ngày', 'Sổ', ...codes.map(shopName), 'SL', 'Doanh thu', 'Đơn treo', 'Dòng BI / tính', 'Dòng bị loại', 'Lấy lúc'],
                r.days.map(d => { const e = by.get(d.day) || {}; return [toBI(d.day), d.status, ...codes.map(c => e[c] ? fmt(e[c]) : ''), e.q ? fmt(e.q) : '', e.r ? mil(e.r) : '', d.pending || '', d.rows != null ? `${d.rows} / ${d.valid}` : '', d.excluded ? Object.entries(d.excluded).map(([k, v]) => `${k} ${v}`).join(', ') : '', d.at || '']; }),
                { num: n(2, codes.length + 3) });
        }
        if (view.salesTab === 'pending') {
            const bar = el('div', undefined, pane, 'bar');
            el('span', 'Đơn chưa hủy nhưng chưa xuất hoặc chưa giao, gom theo mã đơn, treo lâu nhất lên đầu (đỏ: từ 3 ngày · vàng: 2 ngày). Ngày có đơn treo được lấy lại ở lần đổ sau.', bar, 'kxb-muted');
            const pi = el('button', '🖼 Ảnh đơn treo', bar); pi.type = 'button'; pi.onclick = safely(() => pendingImage(pOrders));
            const cp = el('button', '📋 Sao chép để nhắc nhóm', bar); cp.type = 'button';
            cp.onclick = safely(() => copyText(pendingMessage(pOrders, today), `Đã sao chép tin nhắc ${fmt(pOrders.filter(o => o.age >= 2).length || pOrders.length)} đơn treo`));
            table(pane, ['Treo', 'Ngày tạo', 'Siêu thị', 'Mã đơn', 'Nhân viên', 'Sản phẩm', 'SL', 'Giá trị', 'Trạng thái'],
                pOrders.map(o => [o.age ? `${o.age} ngày` : 'hôm nay', toBI(o.created) + ' ' + (o.time || ''), shopName(o.shop), o.order, o.creator, o.items.join(' + '), fmt(o.qty), mil(o.value), o.state]),
                { num: [6, 7], rowClass: i => pOrders[i].age >= 3 ? 'old' : pOrders[i].age === 2 ? 'mid' : '' });
            el('div', 'Chi tiết từng dòng', pane, 'kxb-muted').style.marginTop = '10px';
            table(pane, ['Ngày tạo', 'Giờ', 'Siêu thị', 'Mã đơn', 'Loại YCX', 'Nhân viên', 'Hãng', 'Sản phẩm', 'IMEI', 'Loại hàng', 'SL', 'Giá bán', 'Xuất', 'Giao'],
                r.pendingLines.map(l => [toBI(l.created), l.time, shopName(l.shop), l.order, l.orderType, l.creator, l.brand, l.productName, l.imei || '', conditionText(l.condition), l.qty, fmt(l.price), l.exported, l.delivered]), { num: [10, 11] });
        }
        if (view.salesTab === 'returned') {
            el('div', `Đơn có sản phẩm khách nhập trả — không tính vào số bán (bỏ cả đơn). Ngày trong ${config.returnDays} ngày gần nhất được lấy lại mỗi lần đổ để bắt đơn trả sau.`, pane, 'kxb-muted');
            table(pane, ['Ngày tạo', 'Ngày xuất', 'Siêu thị', 'Mã đơn', 'Nhân viên', 'Hãng', 'Sản phẩm', 'IMEI', 'Loại hàng', 'SL', 'Giá bán', 'Trả hàng'],
                r.returnedLines.map(l => [toBI(l.created), l.shipped ? toBI(l.shipped) : '', shopName(l.shop), l.order, l.creator, l.brand, l.productName, l.imei || '', conditionText(l.condition), l.qty, fmt(l.price), l.returned]), { num: [9, 10] });
        }
        if (view.salesTab === 'detail') {
            const bar = el('div', undefined, pane, 'bar');
            const q = el('input', undefined, bar); q.type = 'search'; q.placeholder = 'Tìm IMEI, mã đơn, nhân viên, hãng, sản phẩm'; q.value = view.salesQuery; q.style.minWidth = '320px';
            const holder = el('div', undefined, pane);
            const draw = () => {
                holder.replaceChildren(); const k = norm(view.salesQuery);
                const rows = r.lines.filter(l => !k || [l.order, l.creator, l.brand, l.productName, l.product, l.imei].some(v => norm(v).includes(k)));
                table(holder, ['Ngày tạo', 'Giờ', 'Ngày xuất', 'Siêu thị', 'Mã đơn', 'Loại YCX', 'Nhân viên', 'Hãng', 'Sản phẩm', 'IMEI', 'Loại hàng', 'SL', 'Giá bán'], rows.map(l => [toBI(l.created), l.time, l.shipped ? toBI(l.shipped) : '', shopName(l.shop), l.order, l.orderType, l.creator, l.brand, l.productName, l.imei || '', conditionText(l.condition), l.qty, fmt(l.price)]), { num: [11, 12] });
            };
            let t; q.oninput = () => { clearTimeout(t); t = setTimeout(() => { view.salesQuery = q.value; draw(); }, 200); };
            draw();
        }
    }

    /* ---------- Tab Tồn kho: bộ chọn ---------- */
    function invRecords() { return view.inv ? filterInventory(view.inv.records, { ...view.invFilter, shops: [...view.invShops] }) : []; }
    function fillSelect(sel, values, current, label) {
        sel.replaceChildren(); const o = el('option', `${label}: tất cả`, sel); o.value = '';
        values.forEach(v => { const x = el('option', v, sel); x.value = v; });
        sel.value = values.includes(current) ? current : '';
    }
    function renderInventoryFilters() {
        const box = ui.querySelector('[data-inv-filters]');
        // Siêu thị: dùng cho cả việc chọn shop để đổ và lọc khi xem
        const shopBox = box.querySelector('[data-inv-shops]'); shopBox.replaceChildren();
        config.shops.forEach(sh => {
            const c = keyCode(sh.code), b = el('button', `${sh.code} · ${sh.name}`, shopBox, 'chip'); b.type = 'button';
            b.classList.toggle('on', view.invShops.has(c)); b.setAttribute('aria-pressed', view.invShops.has(c));
            b.onclick = () => { view.invShops.has(c) ? view.invShops.delete(c) : view.invShops.add(c); save('invShops', [...view.invShops]); renderInventoryFilters(); renderInventory(); };
        });
        const base = view.inv ? view.inv.records.filter(r => view.invShops.has(r.shop)) : [];
        const opts = inventoryOptions(base, view.invFilter);
        fillSelect(box.querySelector('[data-f="category"]'), opts.categories, view.invFilter.category, 'Ngành');
        fillSelect(box.querySelector('[data-f="group"]'), opts.groups, view.invFilter.group, 'Nhóm hàng');
        fillSelect(box.querySelector('[data-f="brand"]'), opts.brands, view.invFilter.brand, 'Hãng');
        for (const k of ['category', 'group', 'brand']) view.invFilter[k] = box.querySelector(`[data-f="${k}"]`).value;
        const condBox = box.querySelector('[data-inv-conds]'); condBox.replaceChildren();
        opts.conditions.forEach(c => {
            const b = el('button', c, condBox, 'chip'); b.type = 'button';
            const on = view.invFilter.conditions.includes(c); b.classList.toggle('on', on); b.setAttribute('aria-pressed', on);
            b.onclick = () => { view.invFilter.conditions = on ? view.invFilter.conditions.filter(x => x !== c) : [...view.invFilter.conditions, c]; renderInventoryFilters(); renderInventory(); };
        });
        if (!opts.conditions.length) el('span', 'Đổ tồn kho để có danh sách trạng thái', condBox, 'kxb-muted');
    }
    /* ---------- V2.2: Tra tồn quanh đây (CRM) — dùng được cả khi chưa đổ tồn BI ---------- */
    function renderCrmLookup(pane) {
        const L = view.crmLookup || (view.crmLookup = { key: '', res: null, total: null, at: 0, onlyFree: true });
        const bar = el('div', undefined, pane, 'bar'); bar.style.alignItems = 'center';
        const qw = el('span', undefined, bar); qw.style.cssText = 'position:relative;display:inline-flex;min-width:420px';
        const q = el('input', undefined, qw); q.type = 'search'; q.placeholder = 'Gõ tên (vd: iPhone 18) rồi chọn trong gợi ý, hoặc mã 13 số'; q.value = L.key; q.style.width = '100%';
        const go = el('button', '🔎 Tra tồn', bar, 'primary idle-only'); go.type = 'button';
        const lab = el('label', undefined, bar, 'chk'); const cb = el('input', undefined, lab); cb.type = 'checkbox'; cb.checked = L.onlyFree; el('span', 'Chỉ nơi có tồn bán được', lab);
        cb.onchange = () => { L.onlyFree = cb.checked; renderInventory(); };
        el('span', `Nguồn: CRM (thời gian thực) · hàng Mới · tỉnh ${crmProvs().map(id => CRM_PROVINCES.find(p => p[0] === id)?.[1] || id).join(', ')} (đổi trong ⚙️ Cài đặt) · không tính kho chi nhánh`, bar, 'kxb-muted');
        const run = safely(() => { L.key = clean(q.value); invariant(L.key.length >= 3, 'Gõ mã sản phẩm hoặc tên (ít nhất 3 ký tự)'); return withSession('crm', async session => {
            status('Đang lấy danh sách siêu thị theo tỉnh…'); await ensureCrmStores(); check(session);
            status(`Đang tra CRM: ${L.key}…`);
            const [res, total] = [await crmFetch(L.key, crmStores(), session), await crmFetch(L.key, [], session).catch(() => null)];
            Object.assign(L, { res, total, at: Date.now() }); renderInventory();
            return res.products.length ? `CRM: ${res.products.length} sản phẩm khớp "${L.key}"` : `CRM: không có sản phẩm khớp "${L.key}"`;
        }); });
        go.onclick = run; q.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); run(); } };
        q.addEventListener('input', () => { L.key = q.value; });
        attachSuggest(q, (code, p, all) => { q.value = all ? clean(L.key) : code; L.key = q.value; if (!all) run(); });
        if (!L.res) { el('div', 'Gõ mã hoặc tên sản phẩm rồi bấm "Tra tồn" — xem siêu thị / kho nào quanh đây còn hàng để tư vấn khách hoặc mượn hàng. Cần Chrome đang đăng nhập CRM.', pane, 'kxb-empty'); return; }
        const own = new Set(config.shops.map(x => keyCode(x.code))), info = new Map(L.res.stores.map(x => [x.code, x]));
        const tot = new Map((L.total?.products || []).map(p => [p.code, p.total]));
        el('div', `Tra lúc ${stamp(new Date(L.at))} · ${L.res.products.length} sản phẩm`, pane, 'kxb-muted');
        const rows = [];
        for (const p of L.res.products.filter(crmUsable)) {
            const list = Object.entries(p.byStore).map(([code, c]) => ({ code, c, s: info.get(code) })).filter(x => x.c.qty < 9999 && !isWarehouse(x.s?.full) && (L.onlyFree ? x.c.qty > 0 : x.c.qty + x.c.lock > 0))
                .sort((a, b) => (own.has(b.code) - own.has(a.code)) || b.c.qty - a.c.qty);
            const sys = tot.get(p.code);
            if (!list.length) rows.push([p.name, p.code, '— không nơi nào trong danh sách còn hàng', '', '', '', sys ? fmt(sys.qty) : '']);
            list.forEach((x, i) => rows.push([i ? '' : p.name, i ? '' : p.code, `${own.has(x.code) ? '🏠 ' : ''}${x.code} · ${x.s?.name || x.code}`, fmt(x.c.qty), x.c.lock ? fmt(x.c.lock) : '', fmt(x.c.qty + x.c.lock), i ? '' : (sys ? fmt(sys.qty) : '')]));
        }
        table(pane, ['Sản phẩm', 'Mã SP', 'Siêu thị', 'Tồn bán được', 'Đang khóa', 'Tổng tồn', 'Toàn hệ thống (bán được)'], rows, { num: [3, 4, 5, 6], title: '🏠 = siêu thị trong cụm · chỉ siêu thị có "Tồn bán được" > 0 mới xin / mượn được' });
    }
    function renderInventory() {
        const area = ui.querySelector('[data-inv-result]'); area.replaceChildren();
        if (view.invTab === 'crm' || !view.inv) {
            subtabs(area, [['group', 'Theo nhóm hàng'], ['product', 'Theo sản phẩm'], ['imei', 'Danh sách IMEI'], ['age', 'Tuổi tồn'], ['transit', 'Đang về'], ['crm', '🔎 Tra tồn quanh đây']], view.invTab, k => { view.invTab = k; try { save('invTab', k); } catch { /* bỏ qua */ } renderInventory(); });
            const pane = el('div', undefined, area);
            if (view.invTab === 'crm') return renderCrmLookup(pane);
            el('div', 'Chọn siêu thị rồi bấm "Đổ tồn kho". Muốn xem nhanh siêu thị khác còn hàng không thì dùng tab "🔎 Tra tồn quanh đây" (không cần đổ tồn).', pane, 'kxb-empty'); return;
        }
        const notIn = [...view.invShops].filter(c => !view.inv.shops.includes(c));
        if (notIn.length) el('div', `Siêu thị ${notIn.map(shopName).join(', ')} chưa có số tồn — bấm "Đổ tồn kho" để lấy.`, area, 'kxb-warn');
        const f = view.invFilter, active = [f.category, f.group, f.brand, ...f.conditions, f.q].filter(Boolean);
        if (active.length) el('div', `Đang lọc: ${active.join(' · ')} — bấm "Bỏ lọc" để xem tất cả.`, area, 'kxb-warn');
        const recs = invRecords();
        const v = inventoryViews(recs), today = isoDate(new Date());
        const times = view.inv.shops.filter(c => view.invShops.has(c)).map(c => `${shopName(c)} lúc ${stamp(new Date(view.inv.shopTimes?.[c] || view.inv.capturedAt))}`).join(' · ');
        el('div', `Tồn: ${times || '—'} · đang xem ${fmt(recs.length)}/${fmt(view.inv.records.filter(r => view.invShops.has(r.shop)).length)} dòng · không gồm hàng đang chuyển kho`, area, 'kxb-muted');
        const ageLimit = config.ageAlert || 60;
        const aged = recs.map(r => ({ r, age: ageDays(r.input, today) }));
        const old = aged.filter(x => x.age != null && x.age >= ageLimit);
        const transit = (view.inv.transit || []).filter(r => view.invShops.has(r.shop));
        kpis(area, [['SL tồn', fmt(v.quantity), mil(v.cost)], ['Nhóm hàng', fmt(v.groups.length)], ['Sản phẩm', fmt(v.products.length)], ['IMEI / Serial', fmt(v.serials)],
            [`Tồn từ ${ageLimit} ngày`, `SL ${fmt(old.reduce((a, x) => a + x.r.qty, 0))}`, mil(old.reduce((a, x) => a + (x.r.cost || 0), 0)), old.length ? { text: `${pct(v.cost ? old.reduce((a, x) => a + (x.r.cost || 0), 0) / v.cost * 100 : 0)} giá trị tồn`, cls: 'down' } : null],
            ...(view.inv.transitShops?.length ? [['Đang về (chưa nhận)', `SL ${fmt(transit.reduce((a, r) => a + r.qty, 0))}`, mil(transit.reduce((a, r) => a + (r.cost || 0), 0))]] : [])]);
        subtabs(area, [['group', 'Theo nhóm hàng'], ['product', 'Theo sản phẩm'], ['imei', 'Danh sách IMEI'], ['age', `Tuổi tồn (${fmt(old.length)})`], ['transit', `Đang về${view.inv.transitShops?.length ? ` (${fmt(transit.length)})` : ''}`], ['crm', '🔎 Tra tồn quanh đây']], view.invTab, k => { view.invTab = k; try { save('invTab', k); } catch { /* bỏ qua */ } renderInventory(); });
        const pane = el('div', undefined, area), codes = view.inv.shops.filter(c => view.invShops.has(c));
        const cond = v.conditionList;
        if (view.invTab === 'group') table(pane, ['Ngành', 'Nhóm hàng', ...codes.map(shopName), 'SL tồn', ...cond, 'Giá trị'],
            v.groups.map(g => [g.category, g.group, ...codes.map(c => fmt(g.byShop[c] || 0)), fmt(g.quantity), ...cond.map(c => fmt(g.byCond[c] || 0)), mil(g.cost)]),
            { num: [...codes.map((_, i) => i + 2), codes.length + 2, ...cond.map((_, i) => codes.length + 3 + i), codes.length + 3 + cond.length] });
        if (view.invTab === 'product') table(pane, ['Nhóm hàng', 'Hãng', 'Mã SP', 'Tên sản phẩm', ...codes.map(shopName), 'SL tồn', ...cond, 'Giá trị'],
            v.products.map(p => [p.group, p.brand, p.product, p.name, ...codes.map(c => fmt(p.byShop[c] || 0)), fmt(p.quantity), ...cond.map(c => fmt(p.byCond[c] || 0)), mil(p.cost)]),
            { num: [...codes.map((_, i) => i + 4), codes.length + 4, ...cond.map((_, i) => codes.length + 5 + i), codes.length + 5 + cond.length] });
        if (view.invTab === 'imei') table(pane, ['Siêu thị', 'Nhóm hàng', 'Hãng', 'Mã SP', 'Tên sản phẩm', 'IMEI / Serial', 'Trạng thái', 'SL', 'Giá vốn', 'Ngày nhập', 'Tuổi (ngày)'],
            aged.map(({ r, age }) => [shopName(r.shop), r.group, r.brand, r.product, r.productName, r.serial, r.condition, fmt(r.qty), fmt(Math.round(r.cost)), r.input, age == null ? '' : fmt(age)]), { num: [7, 8, 10] });
        if (view.invTab === 'age') {
            // Tuổi tồn = số ngày từ Ngày nhập đến hôm nay; cột "Bán 30 ngày" lấy từ số xuất bán đã lưu (không gọi BI)
            const sold = new Map(), r30 = { from: addDays(today, -30), to: addDays(today, -1) };
            inPeriod(loadBook(r30).lines, r30, 'created').forEach(l => sold.set(l.shop + '|' + l.product, (sold.get(l.shop + '|' + l.product) || 0) + l.qty));
            const buckets = [...AGE_BUCKETS.map(b => b[2]), 'Không rõ ngày nhập'].map(b => {
                const xs = aged.filter(x => ageBucket(x.age) === b);
                return { b, lines: xs.length, qty: xs.reduce((a, x) => a + x.r.qty, 0), cost: xs.reduce((a, x) => a + (x.r.cost || 0), 0) };
            }).filter(x => x.lines);
            el('div', `Tuổi tồn tính từ Ngày nhập kho của từng máy / lô đến hôm nay · mốc cảnh báo ${ageLimit} ngày (đổi trong ⚙️ Cài đặt) · "Bán 30 ngày" = SL đã bán của mã đó tại siêu thị trong 30 ngày qua (theo số xuất bán đã lưu).`, pane, 'kxb-muted');
            table(pane, ['Tuổi tồn', 'Số dòng', 'SL', 'Giá vốn', '% giá trị'], [...buckets.map(x => [x.b, fmt(x.lines), fmt(x.qty), mil(x.cost), pct(v.cost ? x.cost / v.cost * 100 : 0)]),
                ['Tổng', fmt(recs.length), fmt(v.quantity), mil(v.cost), '100%']], { num: [1, 2, 3, 4], total: true, title: 'Theo nhóm tuổi' });
            const list = old.sort((a, b) => b.age - a.age || (b.r.cost || 0) - (a.r.cost || 0));
            const deadSlow = list.filter(x => !(sold.get(x.r.shop + '|' + x.r.product))).length;
            table(pane, ['Siêu thị', 'Nhóm hàng', 'Hãng', 'Mã SP', 'Tên sản phẩm', 'IMEI / Serial', 'Trạng thái', 'SL', 'Giá vốn', 'Ngày nhập', 'Tuổi (ngày)', 'Bán 30 ngày'],
                list.map(({ r, age }) => [shopName(r.shop), r.group, r.brand, r.product, r.productName, r.serial, r.condition, fmt(r.qty), mil(r.cost || 0), r.input, fmt(age), fmt(sold.get(r.shop + '|' + r.product) || 0)]),
                { num: [7, 8, 10, 11], title: `Hàng tồn từ ${ageLimit} ngày · ${fmt(list.length)} dòng · ${fmt(deadSlow)} dòng không bán được cái nào trong 30 ngày (tô đỏ)`, rowClass: i => (list[i] && !sold.get(list[i].r.shop + '|' + list[i].r.product)) ? 'low' : '' });
        }
        if (view.invTab === 'transit') {
            if (!view.inv.transitShops?.length) { el('div', 'Chưa lấy hàng đang về. Tích ô "Kèm hàng đang về" cạnh nút Đổ tồn kho rồi bấm "Đổ tồn kho" (mỗi siêu thị gọi BI thêm 1 lần).', pane, 'kxb-empty'); return; }
            el('div', `Hàng đang về ${view.inv.transitShops.map(shopName).join(', ')} = hàng nơi khác đã xuất chuyển đến, siêu thị chưa nhận (phần BI có thêm khi tick "tính hàng đang chuyển"). Không cộng vào số tồn ở các bảng khác; Cân hàng trừ phần này vào số xin.`, pane, 'kxb-muted');
            table(pane, ['Siêu thị', 'Ngành', 'Nhóm hàng', 'Hãng', 'Mã SP', 'Tên sản phẩm', 'IMEI / Serial', 'Trạng thái', 'SL', 'Giá vốn', 'Ngày nhập'],
                transit.map(r => [shopName(r.shop), categoryText(r.category), r.group, r.brand, r.product, r.productName, r.serial, r.condition, fmt(r.qty), mil(r.cost || 0), r.input]), { num: [8, 9] });
        }
    }

    /* ---------- Tab Cân hàng: tồn (Mới) vs tốc độ bán Điện thoại ---------- */
    const BAL_CATEGORY = 'Điện thoại', BAL_CONDITION = 'Mới';
    function balanceWindow() { const to = addDays(isoDate(new Date()), -1); return { from: addDays(to, -view.balDays + 1), to }; }   // N ngày trọn, đến hôm qua
    function balanceData() {
        const range = balanceWindow(), today = isoDate(new Date());
        const { book, lines } = loadBook(range);
        const missing = daysIn(range).filter(d => d <= today && (!book[d] || book[d].schema !== SALES_SCHEMA)).length;
        const shops = view.balShops;
        const cat = view.balCat || BAL_CATEGORY;
        const isPhone = c => categoryText(c) === cat, isNew = c => conditionText(c) === BAL_CONDITION;
        const sales = inPeriod(lines, range, 'created').filter(l => shops.has(l.shop) && isPhone(l.category) && isNew(l.condition));
        const stock = view.inv ? view.inv.records.filter(r => shops.has(r.shop) && isPhone(r.category) && isNew(r.condition)) : [];
        const transit = view.inv ? (view.inv.transit || []).filter(r => shops.has(r.shop) && isPhone(r.category) && isNew(r.condition)) : [];
        const target = config.balTarget || 14, spare = config.balSpare ?? 1;
        const useIncoming = config.balIncoming !== false;
        const allRows = balanceRows(stock, sales, { days: view.balDays, target, spare, byModel: view.balModel, transit, useIncoming });
        // Chuyển nội cụm: tính trên tất cả hãng (trước khi lọc hãng)
        const moves = shops.size > 1 ? transferPlan(allRows) : [];
        const inMap = new Map(), outMap = new Map();
        moves.forEach(m => { inMap.set(m.to + '|' + m.product, (inMap.get(m.to + '|' + m.product) || 0) + m.qty); outMap.set(m.from + '|' + m.product, (outMap.get(m.from + '|' + m.product) || 0) + m.qty); });
        allRows.forEach(r => { r.moveIn = inMap.get(r.shop + '|' + r.product) || 0; r.moveOut = outMap.get(r.shop + '|' + r.product) || 0; r.ask = Math.max(0, r.need - r.moveIn); });
        // Hãng: gộp theo tên không phân biệt hoa thường (BI tồn ghi "Oppo", BI bán ghi "OPPO")
        const brands = new Map(), rev = new Map();
        sales.forEach(l => { if (!l.outShop || keyCode(l.outShop) === keyCode(l.shop)) rev.set(brandKey(l.brand), (rev.get(brandKey(l.brand)) || 0) + l.qty * l.price); });
        allRows.forEach(r => { const k = norm(r.brand) || '(không rõ)'; const b = brands.get(k) || { key: k, label: r.brand || '(không rõ)', need: 0, ask: 0, codes: 0 }; b.need += r.need; b.ask += r.ask; if (r.need > 0) b.codes++; brands.set(k, b); });
        [...view.balBrands].forEach(k => { if (!brands.has(k)) view.balBrands.delete(k); });
        const inBrand = r => !view.balBrands.size || view.balBrands.has(norm(r.brand) || '(không rõ)');
        const rows = allRows.filter(inBrand);
        const noInv = [...shops].filter(c => !view.inv || !view.inv.shops.includes(c));
        const hasTransit = !!view.inv?.transitShops?.some(c => shops.has(c));
        // Danh sách ngành để chọn: theo tồn kho + số bán đã lưu, ngành có doanh số cao lên trên
        const catRev = new Map();
        inPeriod(lines, range, 'created').filter(l => shops.has(l.shop)).forEach(l => catRev.set(categoryText(l.category), (catRev.get(categoryText(l.category)) || 0) + l.qty * l.price));
        (view.inv?.records || []).forEach(r => { const k = categoryText(r.category); if (!catRev.has(k)) catRev.set(k, 0); });
        const cats = [...catRev.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'vi')).map(([k]) => k);
        if (!cats.includes(cat)) cats.unshift(cat);
        return { range, missing, rows, target, spare, noInv, hasTransit, cat, cats, useIncoming, moves: moves.filter(inBrand), brands: [...brands.values()].map(b => ({ ...b, rev: rev.get(b.key) || 0 })).sort((a, b) => b.rev - a.rev || b.need - a.need || a.label.localeCompare(b.label, 'vi')) };   // hãng bán nhiều tiền nhất lên đầu
    }
    // Đổ cân hàng = đổ tồn kho các siêu thị đang chọn + xuất bán N ngày còn thiếu, trong một lần bấm
    // V2.0: không đổi kỳ đang xem ở tab Xuất bán
    async function runBalance(session) {
        const keep = view.invShops;
        view.invShops = new Set(view.balShops);
        try { log('Cân hàng: đổ tồn kho'); await runInventory(session); }
        finally { view.invShops = keep; }
        const w = balanceWindow();
        log(`Cân hàng: đổ xuất bán ${toBI(w.from)}–${toBI(w.to)}`);
        const msg = await runSales(session, false, w);
        renderBalance();
        return `Cân hàng xong: tồn kho mới + xuất bán ${view.balDays} ngày` + (/lỗi|chưa/i.test(msg || '') ? ` · ${msg}` : '');
    }
    // V2.2: tìm nguồn hàng trên CRM cho các mã cần xin (mỗi mã 1 lần gọi, ~0,2 giây)
    async function runCrmBalance(session, codes) {
        const out = {}, failed = [];
        status('Đang lấy danh sách siêu thị theo tỉnh…'); await ensureCrmStores(); check(session);
        for (let i = 0; i < codes.length; i++) {
            check(session); status(`Tìm nguồn hàng CRM ${i + 1}/${codes.length}: ${codes[i]}`); progress(i, codes.length);
            try {
                const res = await crmFetch(codes[i], crmStores(), session);
                const p = res.products.find(x => x.code === codes[i] && !x.service);
                out[codes[i]] = p ? { p, stores: res.stores } : null;
            } catch (e) { if (e.code === 'CANCELLED' || e.code === 'CRM_LOGIN') throw e; failed.push(codes[i]); log(`CRM ${codes[i]}: ${e.message}`, 'error'); }
            await pause(250, session);
        }
        view.crmBal = { at: Date.now(), map: { ...(view.crmBal?.map || {}), ...out } };
        renderBalance();
        return `Tìm nguồn CRM xong: ${codes.length - failed.length}/${codes.length} mã${failed.length ? ` · lỗi ${failed.length} mã, bấm lại để thử lại` : ''}`;
    }
    function crmSrc(r) {
        const codes = view.balModel ? (r.codeList || []) : [r.product], got = codes.map(c => view.crmBal?.map?.[c]).filter(x => x !== undefined);
        if (!got.length) return null;
        const hits = got.filter(Boolean);
        // V2.3: xếp gần trước (⭐ ưu tiên → cùng huyện → cùng tỉnh cũ → xa hơn) theo siêu thị cần hàng
        return rankSources(hits.map(x => x.p), [...(config.crmStoreCache?.stores || []), ...hits.flatMap(x => x.stores)], r.shop, reqFavs(r.shop), config.shops.map(s => s.code));
    }
    const srcText = (list, n = 3) => list && list.length ? list.slice(0, n).map(x => `${x.tier === 0 ? '⭐' : ''}${x.name} ${fmt(x.free)}${x.tier === 3 ? ' (xa)' : ''}`).join(' · ') + (list.length > n ? ` · +${list.length - n} nơi` : '') : (list ? 'không nơi nào xin được' : '');
    // V2.5: đề xuất xin từ đâu (cần đã "Tìm nguồn hàng (CRM)"). Chia theo thứ tự dòng (hết hàng / sắp cạn trước), 2 siêu thị trong cụm không xin trùng 1 máy.
    const srcKeep = () => config.srcKeep ?? 1;
    function balancePlans(rows) {
        if (!view.crmBal) return new Map();
        const items = rows.filter(r => r.ask > 0).map(r => ({ key: r.shop + '|' + r.product, need: r.ask, sources: crmSrc(r), pkey: r.product })).filter(it => it.sources);
        return allocateMany(items, srcKeep());
    }
    const planOf = (plans, r) => plans?.get(r.shop + '|' + r.product) || null;
    const planItems = (rows, plans) => rows.filter(r => planOf(plans, r)).map(r => ({ to: r.shop, product: view.balModel ? '' : r.product, name: r.name, alloc: planOf(plans, r) }));
    const srcPlace = g => areaText(storeArea((config.crmStoreCache?.stores || []).find(x => x.code === g.code)?.full)) || g.provName || areaText(g.area) || '';
    // Tin gửi từng nơi cho: "🏬 Trực Thuận 1 — 3 máy: • SP · 2 máy → Trực Cường"
    function sourceMessage(items, title) {
        const groups = groupBySource(items);
        invariant(groups.length, 'Chưa có đề xuất nơi cho — bấm "🔎 Tìm nguồn hàng (CRM)" trước (Chrome phải đang đăng nhập CRM)');
        const lines = [title];
        for (const g of groups) {
            const place = srcPlace(g);
            lines.push(`\n🏬 ${g.code} · ${g.name}${place ? ' · ' + place : ''} — ${g.qty} máy:`);
            g.lines.forEach(l => lines.push(`  • ${l.name}${l.product ? ` (${l.product})` : ''} · ${l.qty} máy → ${shopName(l.to)}${l.last ? ' (máy cuối của nơi cho)' : ''}`));
        }
        const short = items.filter(i => i.alloc.short);
        if (short.length) { lines.push('\n⚠️ Còn thiếu, chưa có nơi cho (xin kho tổng / tỉnh lân cận):'); short.forEach(i => lines.push(`  • ${shopName(i.to)}: ${i.name}${i.product ? ` (${i.product})` : ''} · thiếu ${i.alloc.short}`)); }
        lines.push('\nNhờ anh chị hỗ trợ chuyển giúp ạ. Cảm ơn!');
        return lines.join('\n');
    }
    function balStatusRows(d) { return d.rows.filter(r => !view.balStatus || (view.balStatus === 'need' ? r.need > 0 : r.status === view.balStatus)); }
    function renderBalance() {
        const area = ui.querySelector('[data-bal-result]'); area.replaceChildren();
        const dBox = ui.querySelector('[data-bal-days]'); dBox.replaceChildren();
        [7, 10, 14, 30].forEach(n => { const b = el('button', `${n} ngày`, dBox, 'chip'); b.type = 'button'; b.classList.toggle('on', view.balDays === n); b.onclick = () => { view.balDays = n; save('balDays', n); renderBalance(); }; });
        const sBox = ui.querySelector('[data-bal-shops]'); sBox.replaceChildren();
        config.shops.forEach(sh => {
            const c = keyCode(sh.code), on = view.balShops.has(c), b = el('button', `${sh.code} · ${sh.name}`, sBox, 'chip'); b.type = 'button'; b.classList.toggle('on', on);
            b.onclick = () => { if (on && view.balShops.size === 1) return; on ? view.balShops.delete(c) : view.balShops.add(c); renderBalance(); };
        });
        const mBtn = ui.querySelector('[data-bal-model]');
        if (mBtn) { mBtn.classList.toggle('on', !!view.balModel); mBtn.setAttribute('aria-pressed', !!view.balModel); mBtn.onclick = () => { view.balModel = !view.balModel; try { save('balModel', view.balModel); } catch { /* bỏ qua */ } renderBalance(); }; }
        const d = balanceData();
        const catSel = ui.querySelector('[data-bal-cat]');
        if (catSel) { catSel.replaceChildren(); d.cats.forEach(k => { const o = el('option', k, catSel); o.value = k; }); catSel.value = d.cat;
            catSel.onchange = () => { view.balCat = catSel.value; view.balBrands.clear(); try { save('balCat', view.balCat); } catch { /* bỏ qua */ } renderBalance(); }; }
        multiDrop(ui.querySelector('[data-bal-brands]'), 'bal-brand', d.brands.map(b => ({ key: b.key, name: b.label, label: b.need ? `${b.label} · xin ${fmt(b.ask)}` : b.label })),
            view.balBrands, `Tất cả hãng · xin ${fmt(d.brands.reduce((a, b) => a + b.ask, 0))}`, renderBalance);
        if (view.balBrands.size) el('div', `Đang lọc hãng: ${d.brands.filter(b => view.balBrands.has(b.key)).map(b => b.label).join(', ')} — tích "Tất cả hãng" để bỏ lọc.`, area, 'kxb-warn');
        el('div', `${d.cat} · hàng Mới · ${view.balModel ? 'gộp theo model (bỏ màu) · ' : ''}tồn = hàng đã nhận${d.hasTransit && d.useIncoming ? ' · Thiếu đã trừ hàng đang về' : ''} · tốc độ bán = SL bán ${toBI(d.range.from)}–${toBI(d.range.to)} (${view.balDays} ngày) ÷ ${view.balDays} · Cần có = TB/ngày × ${d.target} ngày + ${d.spare} máy dự phòng · Sắp hết = tồn dưới mức cần có → Nên xin phần thiếu · Tồn nhiều = đủ bán trên ${d.target * 2} ngày · đơn xuất từ kho khác không tính`, area, 'kxb-muted');
        if (!view.inv) { el('div', 'Chưa có số tồn kho — bấm "Đổ cân hàng" để lấy tồn kho và xuất bán.', area, 'kxb-warn'); }
        else {
            if (d.noInv.length) el('div', `Siêu thị ${d.noInv.map(shopName).join(', ')} chưa có số tồn — bấm "Đổ cân hàng".`, area, 'kxb-warn');
            el('div', `Tồn lấy lúc ${[...view.balShops].filter(c => view.inv.shops.includes(c)).map(c => `${shopName(c)} ${stamp(new Date(view.inv.shopTimes?.[c] || view.inv.capturedAt))}`).join(' · ')}.`, area, 'kxb-muted');
        }
        if (d.missing) {
            el('div', `Còn ${d.missing} ngày trong ${view.balDays} ngày gần nhất chưa đổ xuất bán (hoặc lưu bằng bản cũ) — tốc độ bán đang thấp hơn thực tế. Bấm "Đổ cân hàng".`, area, 'kxb-warn');
        }
        const counts = {}; d.rows.forEach(r => { counts[r.status] = (counts[r.status] || 0) + 1; });
        const need = d.rows.reduce((a, r) => a + r.need, 0), ask = d.rows.reduce((a, r) => a + r.ask, 0), moved = d.moves.reduce((a, m) => a + m.qty, 0);
        const plans = balancePlans(d.rows), pl = [...plans.values()], got = pl.reduce((a, x) => a + x.got, 0), noSrc = d.rows.filter(r => r.ask > 0 && !planOf(plans, r)).reduce((a, r) => a + r.ask, 0);
        const shortAll = pl.reduce((a, x) => a + x.short, 0) + noSrc;
        kpis(area, [['Xin kho', `SL ${fmt(ask)}`, `${fmt(d.rows.filter(r => r.ask > 0).length)} mã`, moved ? { text: `thiếu ${fmt(need)} − chuyển nội cụm ${fmt(moved)}`, cls: 'up' } : null],
            ...(view.crmBal ? [['Đã có nơi cho', `SL ${fmt(got)}`, `${fmt(groupBySource(planItems(d.rows, plans)).length)} siêu thị cho`], ['Còn thiếu', `SL ${fmt(shortAll)}`, 'xin kho tổng / tỉnh lân cận', shortAll ? { text: 'chưa có nơi cho', cls: 'down' } : null]] : []),
            ...(view.balShops.size > 1 ? [['Chuyển nội cụm', `SL ${fmt(moved)}`, `${fmt(d.moves.length)} lượt chuyển`]] : []),
            ...['Hết hàng', 'Sắp hết', 'Đủ', 'Tồn nhiều', 'Không bán'].map(k => [k, fmt(counts[k] || 0), 'mã'])]);
        const tools = el('div', undefined, area, 'bar');
        const bi = el('button', '🖼 Ảnh cân hàng', tools); bi.type = 'button'; bi.title = 'Ảnh danh sách cần bổ sung + đề xuất xin từ siêu thị nào, bao nhiêu máy (tự tra CRM nếu chưa tra) + chuyển nội cụm — chép sẵn, dán Zalo bằng Ctrl+V'; bi.onclick = safely(balanceImageAuto);
        const bc = el('button', '🔎 Tìm nguồn hàng (CRM)', tools, 'idle-only'); bc.type = 'button'; bc.title = 'Tra CRM: siêu thị quanh đây nào có tồn bán được các mã cần xin (Chrome phải đang đăng nhập CRM)';
        bc.onclick = safely(() => {
            const codes = [...new Set(d.rows.filter(r => r.ask > 0 || r.need > 0).flatMap(r => view.balModel ? r.codeList || [] : [r.product]))].filter(c => /^\d{6,}$/.test(c));
            invariant(codes.length, 'Không có mã nào cần xin');
            invariant(codes.length <= 120, `Có ${codes.length} mã cần xin — lọc bớt theo hãng (tối đa 120 mã mỗi lần)`);
            return withSession('crmBal', ss => runCrmBalance(ss, codes));
        });
        const b1 = el('button', '📋 Tin xin hàng', tools); b1.type = 'button'; b1.title = 'Chép danh sách xin hàng (đã trừ phần chuyển nội cụm), gom theo hãng → siêu thị';
        b1.onclick = safely(() => copyText(requestMessage(d.rows.map(r => ({ ...r, need: r.ask })), shopName, { cat: d.cat, sourceLabel: view.crmBal ? 'xin từ' : 'có tại', sourceOf: r => view.crmBal ? planText(planOf(plans, r)) : srcText(crmSrc(r), 2), date: toBI(isoDate(new Date())), basis: `bán ${view.balDays} ngày, giữ đủ ${d.target} ngày` }), 'Đã chép tin xin hàng — dán vào Zalo / LINE'));
        const b3 = el('button', '📋 Tin gửi nơi cho', tools); b3.type = 'button'; b3.title = 'Chép tin gom theo từng siêu thị cho: xin bao nhiêu máy, mã nào, chuyển về đâu (theo đề xuất)';
        b3.onclick = safely(() => copyText(sourceMessage(planItems(d.rows, plans), `🙏 XIN HỖ TRỢ HÀNG ${d.cat.toUpperCase()} · ${toBI(isoDate(new Date()))}`), 'Đã chép tin gửi nơi cho — dán vào Zalo / LINE'));
        if (d.moves.length) {
            const b2 = el('button', '📋 Tin chuyển nội cụm', tools); b2.type = 'button';
            b2.onclick = safely(() => copyText(['🔁 ĐỀ XUẤT CHUYỂN HÀNG NỘI CỤM · ' + toBI(isoDate(new Date())), ...d.moves.map(m => `• ${m.name} (${m.product}): ${shopName(m.from)} → ${shopName(m.to)} · ${m.qty} máy`)].join('\n'), 'Đã chép tin chuyển nội cụm'));
        }
        if (view.crmBal) el('div', `Nguồn CRM tra lúc ${stamp(new Date(view.crmBal.at))} · "Có tại (CRM)" = siêu thị có tồn bán được > 0 (không tính kho chi nhánh, siêu thị trong cụm) · "Xin từ (đề xuất)" = chia số Xin kho cho nơi gần trước, cùng mức gần thì nơi tồn nhiều cho nhiều, nơi cho giữ lại ${srcKeep()} máy (đổi trong ⚙️ Cài đặt).`, area, 'kxb-muted');
        else if (d.rows.some(r => r.ask > 0)) el('div', 'Bấm "🔎 Tìm nguồn hàng (CRM)" (hoặc "🖼 Ảnh cân hàng" — tự tra) để có đề xuất xin từ siêu thị nào, bao nhiêu máy.', area, 'kxb-muted');
        subtabs(area, [['list', 'Bảng cân hàng'], ['transfer', `Chuyển nội cụm (${d.moves.length})`], ['source', 'Nguồn hàng (CRM)'], ['brand', 'Theo hãng']], view.balTab || 'list', k => { view.balTab = k; renderBalance(); });
        const pane = el('div', undefined, area);
        if ((view.balTab || 'list') === 'transfer') {
            if (view.balShops.size < 2) { el('div', 'Chọn từ 2 siêu thị trở lên để xem gợi ý chuyển nội cụm.', pane, 'kxb-empty'); return; }
            el('div', `Cùng mã: siêu thị thừa (tồn vượt mức cần có; mã không bán ở đó thì giữ lại 1 máy) chuyển cho siêu thị đang thiếu — ưu tiên nơi sắp cạn nhất. Làm trước khi xin kho để giảm hàng tồn.`, pane, 'kxb-muted');
            table(pane, ['Mã SP', 'Tên sản phẩm', 'Hãng', 'Từ siêu thị', 'Tồn nơi cho', 'Đủ bán nơi cho', 'Đến siêu thị', 'Tồn nơi nhận', 'Đủ bán nơi nhận', 'SL chuyển'],
                d.moves.map(m => [m.product, m.name, m.brand, shopName(m.from), fmt(m.fromStock), m.fromCover == null ? '∞' : fmt(m.fromCover), shopName(m.to), fmt(m.toStock), m.toCover == null ? '—' : fmt(m.toCover), fmt(m.qty)]), { num: [4, 5, 7, 8, 9] });
            return;
        }
        if (view.balTab === 'source') {
            if (!view.crmBal) { el('div', 'Bấm "🔎 Tìm nguồn hàng (CRM)" để tra các mã cần xin đang có tồn bán được ở siêu thị nào quanh đây.', pane, 'kxb-empty'); return; }
            const list = d.rows.filter(r => r.ask > 0 || r.need > 0);
            el('div', 'Chỉ tính siêu thị có TỒN BÁN ĐƯỢC > 0 trên CRM (ô "-" hoặc 0 là không xin được) · không tính kho chi nhánh và siêu thị trong cụm · dòng đỏ = không nơi nào xin được.', pane, 'kxb-muted');
            table(pane, ['Siêu thị cần', 'Sản phẩm', view.balModel ? 'Model' : 'Mã SP', 'Xin', 'Xin từ (đề xuất)', 'Gần nhất (⭐ / cùng huyện / cùng tỉnh cũ)', 'Số nơi gần', 'Bán được gần', 'Nơi xa'],
                list.map(r => { const src = crmSrc(r); if (!src) return [shopName(r.shop), r.name, view.balModel ? `${r.codeCount} mã` : r.product, fmt(r.ask), '', 'chưa tra', '', '', ''];
                    const near = src.filter(x => x.tier <= 2), far = src.filter(x => x.tier > 2);
                    return [shopName(r.shop), r.name, view.balModel ? `${r.codeCount} mã` : r.product, fmt(r.ask), r.ask ? planText(planOf(plans, r)) : '', near.length ? srcText(near, 5) : (far.length ? 'gần không còn' : 'không nơi nào xin được'), fmt(near.length), fmt(near.reduce((a, x) => a + x.free, 0)), far.length ? `${fmt(far.length)} nơi · ${fmt(far.reduce((a, x) => a + x.free, 0))} máy` : '']; }),
                { num: [3, 6, 7], rowClass: i => { const src = crmSrc(list[i]); return src && !src.length ? 'old' : src && !src.some(x => x.tier <= 2) ? 'mid' : ''; } });
            const groups = groupBySource(planItems(d.rows, plans));
            if (groups.length) {
                el('div', 'Phiếu xin theo nơi cho (gửi từng siêu thị)', pane, 'kxb-reqsub');
                table(pane, ['Nơi cho', 'Khu vực', 'Sản phẩm', 'Mã SP', 'Xin', 'Nơi cho còn lại', 'Về siêu thị'],
                    groups.flatMap(g => g.lines.map((l, i) => [i ? '' : `${g.code} · ${g.name} (${g.qty} máy)`, i ? '' : srcPlace(g), l.name, l.product, fmt(l.qty), fmt(l.left) + (l.last ? ' (máy cuối)' : ''), shopName(l.to)])), { num: [4, 5] });
            }
            el('div', 'Muốn xem đủ danh sách / đánh dấu ⭐ siêu thị hay xin: dùng tab "🔁 Check xin hàng".', pane, 'kxb-muted');
            return;
        }
        if (view.balTab === 'brand') {
            const by = new Map(); d.rows.forEach(r => { const k = r.brand || '(không rõ)'; const x = by.get(k) || { k, codes: 0, stock: 0, sold: 0, need: 0, ask: 0, out: 0 }; x.codes++; x.stock += r.stock; x.sold += r.sold; x.need += r.need; x.ask += r.ask; if (r.status === 'Hết hàng') x.out++; by.set(k, x); });
            const list = [...by.values()].sort((a, b) => b.sold - a.sold);
            table(pane, ['Hãng', 'Số mã', 'Tồn', `Bán ${view.balDays} ngày`, 'Đủ bán (ngày)', 'Mã hết hàng', 'Thiếu', 'Xin kho'],
                list.map(x => [x.k, fmt(x.codes), fmt(x.stock), fmt(x.sold), x.sold ? fmt(Math.round(x.stock / (x.sold / view.balDays) * 10) / 10) : '—', fmt(x.out), fmt(x.need), fmt(x.ask)]), { num: [1, 2, 3, 4, 5, 6, 7] });
            return;
        }
        const fBox = el('div', undefined, pane, 'group');
        const nNeed = d.rows.filter(r => r.need > 0).length;
        [['', `Tất cả · ${fmt(d.rows.length)}`], ['need', `Cần xin · ${fmt(nNeed)}`], ...STATUS_ORDER.map(s => [s, `${s} · ${fmt(counts[s] || 0)}`])].forEach(([k, l]) => { const b = el('button', l, fBox, 'chip'); b.type = 'button'; b.classList.toggle('on', view.balStatus === k); b.onclick = () => { view.balStatus = k; renderBalance(); }; });
        const rows = balStatusRows(d);
        if (!rows.length && d.rows.length) { el('div', `Không có mã nào ở mục "${view.balStatus === 'need' ? 'Cần xin' : view.balStatus}" với tốc độ bán ${view.balDays} ngày.`, pane, 'kxb-empty'); return; }
        const multi = view.balShops.size > 1, tr = d.hasTransit;
        const stCls = { 'Hết hàng': 'old', 'Sắp hết': 'mid' };
        table(pane, ['Siêu thị', 'Hãng', view.balModel ? 'Model' : 'Mã SP', 'Tên sản phẩm', 'Tồn', ...(tr ? ['Đang về'] : []), `Bán ${view.balDays} ngày`, 'TB / ngày', 'Đủ bán (ngày)', 'Cần có', 'Thiếu', ...(multi ? ['Nhận nội cụm', 'Cho nội cụm'] : []), 'Xin kho', 'Trạng thái', ...(view.crmBal ? ['Xin từ (đề xuất)', 'Có tại (CRM)'] : [])],
            rows.map(r => [shopName(r.shop), r.brand, view.balModel ? `${r.codeCount} mã` : r.product, r.name, fmt(r.stock), ...(tr ? [r.transit ? fmt(r.transit) : ''] : []), fmt(r.sold), fmt(r.perDay), r.cover == null ? '—' : fmt(r.cover), r.want ? fmt(r.want) : '', r.need ? fmt(r.need) : '',
                ...(multi ? [r.moveIn ? fmt(r.moveIn) : '', r.moveOut ? fmt(r.moveOut) : ''] : []), r.ask ? fmt(r.ask) : '', r.status, ...(view.crmBal ? [r.ask > 0 ? planText(planOf(plans, r)) : '', r.need > 0 || r.ask > 0 ? srcText(crmSrc(r)) : ''] : [])]),
            { num: Array.from({ length: (tr ? 7 : 6) + (multi ? 2 : 0) + 1 }, (_, i) => i + 4), rowClass: i => stCls[rows[i]?.status] || '' });
        if (tr) el('div', d.useIncoming ? '"Đang về" = hàng nơi khác đã xuất chuyển đến, siêu thị chưa nhận — đã trừ vào cột Thiếu / Xin kho (tắt trong ⚙️ Cài đặt nếu muốn).' : '"Đang về" chỉ để tham khảo, chưa trừ vào số xin (bật lại trong ⚙️ Cài đặt).', pane, 'kxb-muted');
        if (!tr && view.inv) el('div', 'Mẹo: tích "Kèm hàng đang về" ở tab Tồn kho rồi đổ lại để trừ hàng đang về vào số xin, tránh xin trùng.', pane, 'kxb-muted');
    }

    /* ---------- V2.3.1: gợi ý tên + mã sản phẩm khi gõ (CRM, toàn hệ thống, chỉ hàng Mới — bỏ DEMO, dịch vụ, PMH) ----------
     * CRM tìm theo tên trả tối đa 30 dòng và lẫn dịch vụ → tìm 2 lượt: "<gõ>" và "Điện thoại <gõ>", gộp, lọc, xếp tồn nhiều trước. */
    const sugCache = new Map();
    async function crmSuggest(text) {
        const key = clean(text); if (sugCache.has(key)) return sugCache.get(key);
        const keys = [key]; if (!/^(điện thoại|dien thoai)/i.test(key) && !/^\d+$/.test(key)) keys.push('Điện thoại ' + key);
        const res = await Promise.all(keys.map(k => crmFetch(k, []).catch(e => { if (e.code === 'CRM_LOGIN') throw e; return { products: [] }; })));
        const map = new Map();
        res.forEach(r => r.products.filter(crmUsable).forEach(p => { if (!map.has(p.code)) map.set(p.code, p); }));
        const list = [...map.values()].sort((a, b) => (b.total?.qty || 0) - (a.total?.qty || 0) || a.name.localeCompare(b.name, 'vi', { numeric: true })).slice(0, 25);
        sugCache.set(key, list); return list;
    }
    // Gắn ô gợi ý cho 1 ô nhập (input hoặc textarea; textarea: gợi ý theo dòng đang gõ). onPick(code, product, all)
    function attachSuggest(field, onPick) {
        const box = document.createElement('div'); box.className = 'kxb-suggest'; box.hidden = true; field.parentElement.style.position = 'relative'; field.insertAdjacentElement('afterend', box);
        let timer, seq = 0;
        const curLine = () => field.tagName === 'TEXTAREA' ? field.value.slice(0, field.selectionStart).split('\n').pop() : field.value;
        const hide = () => { box.hidden = true; };
        const show = async () => {
            const t = clean(curLine());
            if (t.length < 4 || /^\d{13}$/.test(t)) return hide();
            const my = ++seq; box.hidden = false; box.replaceChildren(); el('div', `Đang tìm "${t}" trên CRM…`, box, 'kxb-muted');
            try {
                const list = await crmSuggest(t); if (my !== seq) return;
                box.replaceChildren();
                if (!list.length) { el('div', 'Không thấy sản phẩm Mới nào khớp (đã bỏ DEMO, dịch vụ, PMH)', box, 'kxb-muted'); return; }
                const head = el('div', undefined, box, 'hd'); el('span', `${list.length} sản phẩm · bấm để chọn`, head);
                const allBtn = el('button', `➕ Chọn tất cả ${list.length} mã`, head, 'mini'); allBtn.type = 'button';
                allBtn.onmousedown = e => { e.preventDefault(); onPick(null, null, list); hide(); };
                list.forEach(p => {
                    const it = el('div', undefined, box, 'it');
                    el('span', p.name + (p.status && !/^KD/i.test(p.status) ? ` · ${p.status}` : ''), it, 'nm'); el('span', p.code, it, 'cd');
                    el('span', p.total ? `HT ${fmt(p.total.qty)}` : '', it, 'qt');
                    it.onmousedown = e => { e.preventDefault(); onPick(p.code, p); hide(); };
                });
            } catch (e) { if (my === seq) { box.replaceChildren(); el('div', e.message, box, 'kxb-warn'); } }
        };
        field.addEventListener('input', () => { clearTimeout(timer); timer = setTimeout(show, 450); });
        field.addEventListener('blur', () => setTimeout(hide, 150));
        field.addEventListener('keydown', e => { if (e.key === 'Escape' && !box.hidden) { e.stopPropagation(); hide(); } });
        return { hide };
    }
    // Thay dòng đang gõ trong textarea bằng mã (hoặc nhiều mã)
    function replaceCurLine(field, text) {
        const pos = field.selectionStart, before = field.value.slice(0, pos), after = field.value.slice(pos);
        const start = before.lastIndexOf('\n') + 1, endRel = after.indexOf('\n'), end = pos + (endRel < 0 ? after.length : endRel);
        field.value = field.value.slice(0, start) + text + field.value.slice(end);
        const caret = start + text.length; field.setSelectionRange(caret, caret); field.dispatchEvent(new Event('input'));
    }

    /* ---------- V2.3: Tab "🔁 Check xin hàng" — hết hàng đột xuất: tra nhanh siêu thị nào trong tỉnh còn tồn bán được, gần mình trước ---------- */
    const reqState = () => view.req || (view.req = { q: '', shop: '', results: [], at: 0, far: false, nb: null, alt: {}, qty: 1 });
    const reqFavs = shop => (config.reqFavs && config.reqFavs[shop]) || [];
    function reqKeys(text) {
        const out = [];
        for (const part of String(text || '').split(/[\n;]+/).map(clean).filter(Boolean)) {
            const bits = part.split(',').map(clean).filter(Boolean);
            if (bits.length > 1 && bits.every(b => /^\d{6,}$/.test(b))) out.push(...bits); else out.push(part);
        }
        return [...new Set(out)].slice(0, 20);
    }
    async function runRequest(session) {
        const R = reqState(), keys = reqKeys(R.q);
        invariant(keys.length, 'Gõ mã sản phẩm (13 số) hoặc tên sản phẩm cần xin');
        status('Đang lấy danh sách siêu thị theo tỉnh…'); await ensureCrmStores(); check(session);
        const results = [];
        for (let i = 0; i < keys.length; i++) {
            status(`Check xin hàng ${i + 1}/${keys.length}: ${keys[i]}`); progress(i, keys.length);
            const res = await crmFetch(keys[i], crmStores(), session);
            const total = await crmFetch(keys[i], [], session).catch(() => null);
            results.push({ key: keys[i], res, total });
            if (i < keys.length - 1) await pause(200, session);
        }
        Object.assign(R, { results, at: Date.now(), nb: null, alt: {} });
        // V2.4: mã nào gần không còn → tự tìm màu / dung lượng khác cùng model còn trong tỉnh
        const bases = altNeeded(R);
        if (bases.length) await fetchAlts(R, bases, session);
        renderRequest();
        const n = results.reduce((a, r) => a + r.res.products.filter(crmUsable).length, 0);
        const empty = reqBlocks(R, R.shop).filter(b => !b.list.length).length;
        return n ? `Check xin hàng xong: ${n} sản phẩm${empty ? ` — ${empty} mã cả tỉnh hết, bấm "🔎 Tỉnh lân cận" để tìm thêm` : ''}` : 'CRM không có sản phẩm nào khớp — kiểm tra lại mã / tên';
    }
    // Gốc model cần tìm hàng thay thế: mã gần (cùng tỉnh cũ) không còn, chưa tìm lần nào; tối đa 5 model / lần
    function altNeeded(R) {
        const out = [];
        for (const b of reqBlocks(R, R.shop)) {
            if (b.list.some(x => x.tier <= 2)) continue;
            const base = variantBase(b.p.name), k = norm(base);
            if (base && !R.alt[k] && !out.some(o => o.k === k)) out.push({ k, base, phone: /^điện thoại/i.test(b.p.name) });
        }
        return out.slice(0, 5);
    }
    async function fetchAlts(R, bases, session) {
        for (const b of bases) {
            status(`Tìm màu / dung lượng khác: ${b.base}`);
            const res = await crmFetch((b.phone ? 'Điện thoại ' : '') + b.base, crmStores(), session).catch(e => { if (e.code === 'CRM_LOGIN' || e.code === 'CANCELLED') throw e; return { stores: [], products: [] }; });
            R.alt[b.k] = res; await pause(150, session);
        }
    }
    async function runAlts(session) {
        const R = reqState(), bases = altNeeded(R); invariant(bases.length, 'Không còn mã nào cần tìm hàng thay thế');
        await ensureCrmStores(); await fetchAlts(R, bases, session); renderRequest();
        return `Đã tìm màu / dung lượng khác cho ${bases.length} model`;
    }
    // Siêu thị (không kho) của 1 tỉnh lân cận — lưu 7 ngày
    async function nearStores(id) {
        const c = (config.crmNearCache || {})[id];
        if (c && Array.isArray(c.stores) && c.stores.length && Date.now() - (c.at || 0) < 7 * 864e5) return c.stores;
        const list = (await crmGetStores([id])).filter(x => !x.kho).map(x => x.code);
        (config.crmNearCache || (config.crmNearCache = {}))[id] = { at: Date.now(), stores: list }; save('config', config);
        return list;
    }
    const nearIds = () => nearProvsOf(crmProvs(), config.crmNearProvinces);
    async function runNearby(session) {
        const R = reqState(); invariant(R.results.length, 'Bấm "🔎 Check" trước rồi mới tìm tỉnh lân cận');
        const ids = nearIds(); invariant(ids.length, 'Chưa có tỉnh lân cận — chọn trong ⚙️ Cài đặt → Tra tồn CRM → Tỉnh lân cận');
        const provs = [];
        for (const id of ids) { status(`Lấy danh sách siêu thị ${provShort(id)}…`); provs.push({ id, name: provShort(id), stores: await nearStores(id) }); check(session); }
        const keys = R.results.map(r => r.key), data = {}, total = keys.length * provs.length; let n = 0;
        for (const k of keys) {
            data[k] = {};
            for (const pv of provs) {
                status(`Tỉnh lân cận ${++n}/${total}: ${k} · ${pv.name}`); progress(n - 1, total);
                let res = null;
                for (let i = 0; i < pv.stores.length; i += 200) { res = mergeCrm(res, await crmFetch(k, pv.stores.slice(i, i + 200), session)); await pause(150, session); }
                data[k][pv.id] = res || { stores: [], products: [] };
            }
        }
        R.nb = { provs: provs.map(({ id, name, stores }) => ({ id, name, count: stores.length })), data, at: Date.now() };
        renderRequest();
        const found = reqBlocks(R, R.shop).filter(b => b.nb?.length).length;
        return `Tỉnh lân cận (${provs.map(p => p.name).join(', ')}): ${found} sản phẩm có nơi xin được`;
    }
    function reqMessage(R, shop) {
        const lines = [`🙏 ${shopName(shop)} xin hỗ trợ hàng (${stamp(new Date(R.at))}):`];
        for (const blk of reqBlocks(R, shop)) {
            const near = blk.list.filter(x => x.tier <= 2), pick = (near.length ? near : blk.list).slice(0, 5);
            let t = blk.alloc.plan.length ? `xin ${blk.alloc.need} máy: ${planText(blk.alloc)}` : pick.length ? pick.map(x => `${x.name} ${x.free}`).join(' · ') : blk.nb?.length ? 'trong tỉnh hết · tỉnh lân cận: ' + blk.nb.slice(0, 5).map(x => `${x.name} (${x.provName}) ${x.free}`).join(' · ') : 'trong tỉnh không nơi nào còn';
            if (!near.length && blk.alts?.length) t += ` · có thể thay: ${blk.alts.slice(0, 2).map(a => `${a.p.name} (${a.p.code}) — ${(a.near.length ? a.near : a.list).slice(0, 2).map(x => `${x.name} ${x.free}`).join(', ')}`).join(' / ')}`;
            lines.push(`• ${blk.p.name} (${blk.p.code}): ${t}`);
        }
        invariant(lines.length > 1, 'Chưa có kết quả — bấm "Check" trước');
        return lines.join('\n');
    }
    function reqBlocks(R, shop) {
        const cache = config.crmStoreCache?.stores || [], blocks = [];
        for (const r of R.results) for (const p of r.res.products.filter(crmUsable).slice(0, 15)) {
            const stores = [...cache, ...r.res.stores], alt = R.alt?.[norm(variantBase(p.name))];
            const list = rankSources(p, stores, shop, reqFavs(shop));
            const nb = R.nb ? nearbySources(p, R.nb.provs.map(pv => ({ prov: pv.id, name: pv.name, res: R.nb.data[r.key]?.[pv.id] })), shop) : null;
            blocks.push({ key: r.key, p, list, alloc: allocateAsk(R.qty || 1, [...list, ...(nb || [])], srcKeep()), sys: (r.total?.products || []).find(x => x.code === p.code)?.total || null, mine: p.byStore[keyCode(shop)] || null,
                nb,
                alts: alt ? altOptions(p, alt.products, [...cache, ...alt.stores], shop, reqFavs(shop)) : null });
        }
        return blocks;
    }
    function renderRequest() {
        if (!ui) return;
        const R = reqState(), area = ui.querySelector('[data-req-result]'); area.replaceChildren();
        // chọn siêu thị cần hàng
        const sel = ui.querySelector('[data-req-shop]');
        if (!R.shop || !config.shops.some(x => keyCode(x.code) === R.shop)) R.shop = keyCode(config.shops[0]?.code || '');
        sel.replaceChildren(); config.shops.forEach(sh => { const o = el('option', `${sh.code} · ${sh.name}`, sel); o.value = keyCode(sh.code); }); sel.value = R.shop;
        sel.onchange = () => { R.shop = sel.value; renderRequest(); };
        const q = ui.querySelector('[data-req-q]'); if (document.activeElement !== q) q.value = R.q;
        const far = ui.querySelector('[data-req-far]'); far.checked = R.far; far.onchange = () => { R.far = far.checked; renderRequest(); };
        const qn = ui.querySelector('[data-req-qty]'); if (document.activeElement !== qn) qn.value = R.qty || 1;
        qn.onchange = () => { R.qty = Math.min(50, Math.max(1, Math.round(Number(qn.value) || 1))); qn.value = R.qty; renderRequest(); };
        // gợi ý nhanh: mã Hết hàng / Sắp hết của siêu thị đang chọn (theo số Cân hàng đã đổ)
        const qb = ui.querySelector('[data-req-quick]'); qb.replaceChildren();
        if (view.inv) {
            try {
                const d = balanceData(), hot = d.rows.filter(r => r.shop === R.shop && (r.status === 'Hết hàng' || r.status === 'Sắp hết')).slice(0, 12);
                hot.forEach(r => { const b = el('button', `${r.status === 'Hết hàng' ? '🔴' : '🟡'} ${r.name}`, qb, 'chip'); b.type = 'button'; b.title = 'Check xin hàng mã này';
                    b.onclick = safely(() => { R.q = view.balModel ? r.name : r.product; q.value = R.q; return withSession('crmReq', runRequest); }); });
                if (!hot.length) el('span', 'Không có mã hết / sắp hết (theo Cân hàng)', qb, 'kxb-muted');
            } catch { el('span', 'Đổ cân hàng để có gợi ý mã hết hàng', qb, 'kxb-muted'); }
        } else el('span', 'Đổ cân hàng để có gợi ý mã hết hàng', qb, 'kxb-muted');
        if (!R.results.length) {
            el('div', 'Gõ mã sản phẩm (13 số) hoặc tên rồi bấm "🔎 Check" — tool tra CRM, chỉ lấy siêu thị có TỒN BÁN ĐƯỢC > 0 trong tỉnh đã chọn (⚙️ Cài đặt), xếp: ⭐ ưu tiên → cùng huyện → cùng tỉnh cũ → xa hơn. Nhiều mã: mỗi mã một dòng. Không tính kho chi nhánh.', area, 'kxb-empty');
            return;
        }
        const my = config.crmStoreCache?.stores?.find(x => x.code === R.shop);
        el('div', `Check lúc ${stamp(new Date(R.at))} · xin cho ${shopName(R.shop)}${my ? ' — ' + areaText(storeArea(my.full)) : ''} · số thời gian thực trên CRM · bấm vào một dòng để ⭐ đánh dấu / bỏ siêu thị ưu tiên (nhớ cho lần sau)`, area, 'kxb-muted');
        const bar = el('div', undefined, area, 'bar');
        const cp = el('button', '📋 Tin xin hàng', bar); cp.type = 'button'; cp.title = 'Chép tin nhắn xin hàng: mỗi sản phẩm 5 nơi gần nhất có tồn bán được';
        cp.onclick = safely(() => copyText(reqMessage(R, R.shop), 'Đã chép tin xin hàng — dán vào Zalo / LINE'));
        const nbNames = nearIds().map(provShort).join(', ');
        const nbBtn = el('button', R.nb ? '🔄 Tra lại tỉnh lân cận' : '🔎 Tỉnh lân cận', bar, 'idle-only'); nbBtn.type = 'button';
        nbBtn.title = nbNames ? `Tìm thêm ở ${nbNames} (đổi trong ⚙️ Cài đặt → Tra tồn CRM → Tỉnh lân cận)` : 'Chưa có tỉnh lân cận — chọn trong ⚙️ Cài đặt';
        nbBtn.onclick = safely(() => withSession('crmNear', runNearby));
        if (R.nb) el('span', ` Tỉnh lân cận: ${R.nb.provs.map(p => `${p.name} (${p.count} ST)`).join(', ')} · tra lúc ${stamp(new Date(R.nb.at))}`, bar, 'kxb-muted');
        const blocks = reqBlocks(R, R.shop);
        if (!blocks.length) { el('div', `CRM không có sản phẩm khớp "${R.results.map(r => r.key).join(', ')}".`, area, 'kxb-empty'); return; }
        for (const blk of blocks) {
            const list = R.far ? blk.list : blk.list.filter(x => x.tier <= 2);
            const near = blk.list.filter(x => x.tier <= 2);
            const head = el('div', undefined, area, 'kxb-reqhead');
            el('b', `${blk.p.name} · ${blk.p.code}`, head);
            el('span', ` — tại ${shopName(R.shop)}: ${blk.mine ? fmt(blk.mine.qty) : 0} bán được · gần (cùng tỉnh cũ): ${fmt(near.length)} nơi, ${fmt(near.reduce((a, x) => a + x.free, 0))} máy · cả tỉnh: ${fmt(blk.list.length)} nơi${blk.sys ? ` · toàn hệ thống ${fmt(blk.sys.qty)}` : ''}`, head, 'kxb-muted');
            const al = blk.alloc, mine = new Map(al.plan.map(x => [x.code, x]));
            const pl = el('div', al.plan.length ? `👉 Đề xuất xin ${fmt(al.need)} máy: ${planText(al)}` : `👉 Cần ${fmt(al.need)} máy — chưa có nơi nào cho được${R.nb ? '' : ' (thử "🔎 Tỉnh lân cận")'}`, area, 'kxb-reqsub');
            pl.style.color = al.short ? '#b91c1c' : '#0b6b4f';
            if (!list.length) {
                const w = el('div', blk.list.length ? `Gần không còn — có ${blk.list.length} nơi xa hơn, tích "Hiện cả nơi xa" để xem.` : 'Trong tỉnh không siêu thị nào còn tồn bán được.', area, 'kxb-warn');
                if (!blk.list.length && !R.nb && nbNames) { const b = el('button', `🔎 Tìm thêm ở ${nbNames}`, w, 'mini idle-only'); b.type = 'button'; b.style.marginLeft = '8px'; b.onclick = safely(() => withSession('crmNear', runNearby)); }
            } else {
                const fav = new Set(reqFavs(R.shop));
                table(area, ['', 'Ưu tiên', 'Siêu thị', 'Khu vực', 'Tồn bán được', 'Đang khóa', 'Đề xuất xin'],
                    list.map(x => [fav.has(x.code) ? '⭐' : '☆', x.tierName, `${x.code} · ${x.name}`, areaText(x.area), fmt(x.free), x.lock ? fmt(x.lock) : '', mine.get(x.code) ? fmt(mine.get(x.code).qty) + (mine.get(x.code).last ? ' (máy cuối)' : '') : '']),
                    { num: [4, 5, 6], rowTitle: 'Bấm để ⭐ đánh dấu / bỏ ưu tiên siêu thị này', rowClass: i => list[i].tier === 0 ? 'fav' : list[i].tier === 1 ? 'near' : '', onRow: i => {
                        const all = config.reqFavs || (config.reqFavs = {}), cur = new Set(all[R.shop] || []), c = list[i].code;
                        cur.has(c) ? cur.delete(c) : cur.add(c); all[R.shop] = [...cur]; save('config', config); renderRequest();
                    } });
            }
            // V2.4: tỉnh lân cận
            if (blk.nb) {
                if (blk.nb.length) {
                    el('div', `📍 Tỉnh lân cận: ${fmt(blk.nb.length)} nơi, ${fmt(blk.nb.reduce((a, x) => a + x.free, 0))} máy bán được`, area, 'kxb-reqsub');
                    table(area, ['Tỉnh', 'Siêu thị', 'Khu vực (mã)', 'Tồn bán được', 'Đang khóa', 'Đề xuất xin'],
                        blk.nb.map(x => [x.provName, `${x.code} · ${x.name}`, [x.area?.prov, x.area?.dist].filter(Boolean).join('_'), fmt(x.free), x.lock ? fmt(x.lock) : '', mine.get(x.code) ? fmt(mine.get(x.code).qty) : '']), { num: [3, 4, 5] });
                } else el('div', `📍 Tỉnh lân cận (${R.nb.provs.map(p => p.name).join(', ')}) cũng không nơi nào còn tồn bán được.`, area, 'kxb-muted');
            }
            // V2.4: hàng thay thế (khác màu / dung lượng, cùng model) khi gần không còn
            if (!near.length) {
                const base = variantBase(blk.p.name);
                if (blk.alts?.length) {
                    el('div', `🔁 Có thể thay bằng (cùng ${base}, khác màu / dung lượng) — bấm một dòng để check mã đó`, area, 'kxb-reqsub');
                    const al = blk.alts;
                    table(area, ['Sản phẩm', 'Mã SP', 'Gần (cùng tỉnh cũ)', 'Cả tỉnh', 'Nơi gần nhất'],
                        al.map(a => [a.p.name + (a.p.status && !/^KD/i.test(a.p.status) ? ` · ${a.p.status}` : ''), a.p.code, `${fmt(a.near.length)} nơi · ${fmt(a.nearQty)} máy`, `${fmt(a.list.length)} nơi · ${fmt(a.allQty)} máy`, a.list.slice(0, 3).map(x => `${x.name} ${x.free}`).join(', ')]),
                        { rowTitle: 'Bấm để check xin hàng mã này', onRow: i => { R.q = al[i].p.code; ui.querySelector('[data-req-q]').value = R.q; safely(() => withSession('crmReq', runRequest))(); } });
                } else if (blk.alts) el('div', `🔁 Không có màu / dung lượng khác của ${base} còn trong tỉnh.`, area, 'kxb-muted');
                else if (base) { const b = el('button', `🔁 Tìm màu / dung lượng khác của ${base}`, area, 'mini idle-only'); b.type = 'button'; b.onclick = safely(() => withSession('crmAlt', runAlts)); }
            }
        }
    }

    /* ---------- In phiếu kiểm tồn (A4 đứng) ---------- */
    function printInventory() {
        invariant(!running, 'Đang đổ số, chờ xong rồi in');
        invariant(view.inv, 'Chưa có số tồn kho — bấm "Đổ tồn kho" trước');
        const recs = invRecords(), f = view.invFilter;
        const filter = [f.category && 'Ngành ' + f.category, f.group && 'Nhóm ' + f.group, f.brand && 'Hãng ' + f.brand, f.conditions.length && 'Trạng thái ' + f.conditions.join(', '), f.q && 'Tìm "' + f.q + '"'].filter(Boolean).join(' · ');
        const times = {}; view.inv.shops.forEach(c => { times[c] = stamp(new Date(view.inv.shopTimes?.[c] || view.inv.capturedAt)); });
        const html = inventoryPrintHtml(recs, { nameOf: shopName, shopOrder: config.shops.map(s => s.code), filter, times, capturedAt: stamp(new Date(view.inv.capturedAt)), printedAt: stamp(new Date()), version: VERSION });
        document.getElementById('kxb-print')?.remove();
        const fr = document.createElement('iframe'); fr.id = 'kxb-print'; fr.dataset.kxbUi = '';
        fr.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden';
        document.body.append(fr);
        const d = fr.contentWindow.document; d.open(); d.write(html); d.close();
        setTimeout(() => { try { fr.contentWindow.focus(); fr.contentWindow.print(); } catch (e) { status('Không mở được hộp thoại in: ' + e.message, 'err'); } }, 300);
        log(`In phiếu kiểm tồn: ${recs.length} dòng${filter ? ' · ' + filter : ''}`);
        status(`Đã mở hộp thoại in phiếu kiểm (${fmt(recs.length)} dòng). Chọn khổ A4, hướng dọc.`, 'ok');
    }

    /* ---------- Ảnh báo cáo gửi Zalo (vẽ canvas) — V2.1: khuôn chung cho Xuất bán / Cân hàng / Đơn treo ---------- */
    // spec = { title, sub, cards: [{ t, v, s, d, main }], sections: [{ title, cols: [[nhãn, độ rộng tương đối, 'right'?]], rows, rowColor(i) }], footer, warn, name }
    async function reportImage(spec) {
        const IW = spec.width || 1000, P = 28, DPR = 2, rowH = 30, perRow = 6;
        const cards = spec.cards || [], cardRows = Math.ceil(cards.length / perRow);
        const secs = (spec.sections || []).filter(x => x && x.rows && x.rows.length);
        const H = 104 + cardRows * 106 + secs.reduce((a, x) => a + 44 + rowH * (x.rows.length + 1) + 14, 0) + 50;
        const cv = document.createElement('canvas'); cv.width = IW * DPR; cv.height = H * DPR;
        const g = cv.getContext('2d'); invariant(g, 'Trình duyệt không vẽ được ảnh'); g.scale(DPR, DPR);
        const F = (wt, px) => `${wt} ${px}px Arial, "Segoe UI", sans-serif`;
        const text = (t, x, y, { font = F(400, 15), color = '#172a3a', align = 'left', max = 0 } = {}) => {
            g.font = font; g.fillStyle = color; g.textAlign = align; t = String(t ?? '');
            if (max) while (t.length > 1 && g.measureText(t).width > max) t = t.slice(0, -2) + '…';
            g.fillText(t, x, y);
        };
        const box = (x, y, w, h, fill, rad = 10) => { g.fillStyle = fill; g.beginPath(); g.roundRect ? g.roundRect(x, y, w, h, rad) : g.rect(x, y, w, h); g.fill(); };
        g.fillStyle = '#fff'; g.fillRect(0, 0, IW, H);
        box(0, 0, IW, 90, '#087f8c', 0);
        text(spec.title, P, 40, { font: F(700, 25), color: '#fff', max: IW - 2 * P });
        text(spec.sub || '', P, 70, { font: F(400, 14), color: '#e0f2f3', max: IW - 2 * P });
        let y = 104;
        for (let row = 0; row < cardRows; row++) {
            const list = cards.slice(row * perRow, (row + 1) * perRow), cw = (IW - 2 * P - 10 * (list.length - 1)) / list.length;
            list.forEach((c, i) => {
                const x = P + i * (cw + 10);
                box(x, y, cw, 96, c.main ? '#087f8c' : '#f2f6f7');
                const col = c.main ? '#fff' : '#172a3a';
                text(c.t, x + 12, y + 22, { font: F(400, 13), color: c.main ? '#e0f2f3' : '#4a5a66', max: cw - 20 });
                text(c.v, x + 12, y + 50, { font: F(700, 22), color: col, max: cw - 20 });
                text(c.s || '', x + 12, y + 70, { font: F(400, 13), color: col, max: cw - 20 });
                if (c.d) text(c.d, x + 12, y + 88, { font: F(700, 12), color: c.main ? '#fff' : /^▼|thiếu|treo/.test(c.d) ? '#b91c1c' : '#15803d', max: cw - 20 });
            });
            y += 106;
        }
        for (const sec of secs) {
            y += 24; text(sec.title, P, y, { font: F(700, 17) }); y += 12;
            const total = sec.cols.reduce((a, c) => a + c[1], 0), inner = IW - 2 * P - 20; let cx = P + 10;
            const pos = sec.cols.map(([label, wt, al]) => { const w = inner * wt / total, p = { x: cx, w: w - 8, al }; cx += w; return p; });
            box(P, y, IW - 2 * P, rowH, '#edf5f7', 6);
            const draw = (cells, yy, bold, color, ri) => pos.forEach((p, i) => text(cells[i], p.al === 'right' ? p.x + p.w : p.x, yy + 20, { font: F(bold ? 700 : 400, 14), align: p.al === 'right' ? 'right' : 'left', max: p.span && cells.slice(i + 1).every(c => c === '') ? IW - 2 * P - 20 - (p.x - P - 10) : p.w, color: (ri != null && sec.cellColor?.(ri, i)) || color || '#172a3a' }));
            if (sec.spanFirst) pos[0].span = true;
            draw(sec.cols.map(c => c[0]), y, true);
            sec.rows.forEach((r, i) => {
                const yy = y + rowH * (i + 1), bg = sec.rowColor?.(i);
                if (bg) box(P, yy, IW - 2 * P, rowH, bg, 0); else if (i % 2) box(P, yy, IW - 2 * P, rowH, '#f8fafb', 0);
                draw(r, yy, (sec.boldLast && i === sec.rows.length - 1) || !!sec.rowBold?.(i), null, i);
            });
            y += rowH * (sec.rows.length + 1) + 14;
        }
        text(spec.footer || '', P, y + 26, { font: F(400, 12), color: spec.warn ? '#b91c1c' : '#5b6b76', max: IW - 2 * P });
        const blob = await new Promise(res => cv.toBlob(res, 'image/png'));
        invariant(blob, 'Không tạo được ảnh');
        const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = spec.name; document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000);
        let copied = false;
        try { if (window.ClipboardItem) { await navigator.clipboard.write([new (window.ClipboardItem)({ 'image/png': blob })]); copied = true; } } catch { /* trình duyệt chặn → chỉ tải file */ }
        status(copied ? `Đã chép ảnh — mở Zalo bấm Ctrl+V để dán (cũng đã tải ${spec.name})` : `Đã tải ảnh ${spec.name}`, 'ok');
        log(`Ảnh báo cáo: ${spec.title}`);
    }
    const imgOpt = () => ({ cats: false, brands: true, staff: true, top: 12, ...(config.img || {}) });
    async function salesImage() {
        const r = salesData(); invariant(r && r.lines.length, 'Chưa có số xuất bán trong kỳ — bấm "Đổ xuất bán" hoặc "Xem số đã lưu"');
        const o = imgOpt(), s = r.summary, ps = r.prevSummary, codes = config.shops.map(x => keyCode(x.code)).filter(c => view.salesShops.has(c));
        const att = attachData(r.allLines.filter(l => view.salesShops.has(l.shop))), attMap = new Map(att.staff.map(x => [x.shop + '|' + x.label, x]));
        const pOrders = pendingOrders(r.pendingLines, isoDate(new Date()));
        const d = (c, p) => { const x = delta(c, p, r); return x ? x.text : ''; };
        const cards = [{ t: 'Cụm', v: `SL ${fmt(s.quantity)}`, s: mil(s.revenue), d: d(s.revenue, ps.revenue), main: true },
            ...codes.map(c => { const x = s.shops.find(z => z.code === c) || { quantity: 0, revenue: 0 }, y = ps.shops.find(z => z.code === c) || { revenue: 0 }; return { t: shopName(c), v: `SL ${fmt(x.quantity)}`, s: mil(x.revenue), d: d(x.revenue, y.revenue) }; }),
            { t: 'Bán kèm ĐT', v: pct(att.total.rate), s: `${fmt(att.total.attached)}/${fmt(att.total.phoneOrders)} đơn` },
            { t: 'Đơn treo', v: fmt(pOrders.length), s: `${fmt(pOrders.filter(x => x.age >= 2).length)} đơn từ 2 ngày`, d: pOrders.some(x => x.age >= 2) ? 'cần xử lý đơn treo' : '' }];
        const vc = salesByCategory(r.lines, codes), catTot = new Map();
        vc.rows.forEach(x => { const e = catTot.get(x.category) || { q: 0, r: 0 }; e.q += x.qty; e.r += x.rev; catTot.set(x.category, e); });
        const brands = s.brands.slice(0, o.top), staff = s.staff.slice(0, o.top);
        await reportImage({
            title: `XUẤT BÁN · ${toBI(r.range.from)} – ${toBI(r.range.to)}`,
            sub: [`Ngành: ${view.salesCat || 'tất cả'}`, view.salesBrands.size ? 'Hãng: ' + [...new Set(r.lines.map(l => l.brand))].join(', ') : '', view.salesConds.size ? 'Loại hàng: ' + [...view.salesConds].join(', ') : '', 'Kho tạo · Đã xuất – Đã giao · bỏ đơn nhập trả'].filter(Boolean).join(' · '),
            cards,
            sections: [
                o.cats && !view.salesCat ? { title: 'Theo ngành hàng', cols: [['Ngành', 5], ['SL', 1.5, 'right'], ['Doanh thu', 2, 'right'], ['% DT', 1.5, 'right']], rows: [...catTot.entries()].sort((a, b) => b[1].r - a[1].r).slice(0, o.top).map(([k, v]) => [k, fmt(v.q), mil(v.r), pct(vc.rev ? v.r / vc.rev * 100 : 0)]) } : null,
                o.brands ? { title: `Theo hãng${s.brands.length > brands.length ? ` (top ${brands.length}/${s.brands.length})` : ''}`, cols: [['Hãng', 3], ['SL', 1.2, 'right'], ['% SL', 1.2, 'right'], ['Doanh thu', 2, 'right'], ['% DT', 1.2, 'right'], ['Giá TB', 2, 'right']],
                    rows: brands.map(b => [b.label, fmt(b.quantity), pct(b.pctQty), mil(b.revenue), pct(b.pctRev), mil(b.quantity ? b.revenue / b.quantity : 0)]) } : null,
                o.staff ? { title: `Nhân viên${s.staff.length > staff.length ? ` (top ${staff.length}/${s.staff.length})` : ''}`, cols: [['Nhân viên', 4], ['Siêu thị', 2], ['SL', 1, 'right'], ['Doanh thu', 1.8, 'right'], ['Kèm ĐT', 2.2, 'right']],
                    rows: staff.map(st => { const a = attMap.get(st.shop + '|' + st.label); return [st.label, shopName(st.shop), fmt(st.quantity), mil(st.revenue), a ? `${fmt(a.attached)}/${fmt(a.phoneOrders)} · ${pct(a.rate)}` : '—']; }) } : null],
            footer: `AutoBI Kho & Xuất Bán V${VERSION} · lấy lúc ${stamp(new Date())}${r.missing ? ` · CHƯA ĐỦ: còn ${r.missing} ngày chưa lấy` : ''}`, warn: !!r.missing,
            name: `AutoBI_XuatBan_${r.range.from.replace(/-/g, '')}-${r.range.to.replace(/-/g, '')}.png`
        });
    }
    // V2.5: bấm Ảnh cân hàng → nếu chưa tra CRM (hoặc đã quá 30 phút / thiếu mã) thì tự tra trước để có đề xuất xin từ đâu.
    // Tra CRM lỗi (chưa đăng nhập…) vẫn ra ảnh, chỉ không có phần đề xuất.
    async function balanceImageAuto() {
        invariant(view.inv, 'Chưa có số tồn kho — bấm "Đổ cân hàng" trước');
        const d = balanceData();
        const codes = [...new Set(d.rows.filter(r => r.ask > 0).flatMap(r => view.balModel ? r.codeList || [] : [r.product]))].filter(c => /^\d{6,}$/.test(c));
        const stale = !view.crmBal || Date.now() - view.crmBal.at > 30 * 60 * 1000 || codes.some(c => !(c in (view.crmBal.map || {})));
        if (codes.length && codes.length <= 120 && stale && !running) {
            const st = await withSession('crmBal', ss => runCrmBalance(ss, codes));
            if (st !== 'completed') log('Ảnh cân hàng: chưa tra được CRM nên ảnh chưa có phần "Xin từ" — đăng nhập CRM trên Chrome rồi bấm lại', 'error');
        }
        await balanceImage();
    }
    async function balanceImage() {
        invariant(view.inv, 'Chưa có số tồn kho — bấm "Đổ cân hàng" trước');
        const d = balanceData(), o = imgOpt(), multi = view.balShops.size > 1;
        const plans = balancePlans(d.rows), hasPlan = plans.size > 0;
        const groups = hasPlan ? groupBySource(planItems(d.rows, plans)) : [];
        const got = [...plans.values()].reduce((a, x) => a + x.got, 0), shortAll = [...plans.values()].reduce((a, x) => a + x.short, 0) + d.rows.filter(r => r.ask > 0 && !planOf(plans, r)).reduce((a, r) => a + r.ask, 0);
        const slipRows = [], slipKind = [];
        groups.slice(0, 15).forEach(g => {
            const place = srcPlace(g);
            slipRows.push([`${g.code} · ${g.name}${place ? ' · ' + place : ''} — ${g.qty} máy`, '', '', '']); slipKind.push('h');
            g.lines.forEach(l => { slipRows.push([`   → ${shopName(l.to)}`, l.name + (l.product ? ` (${l.product})` : ''), fmt(l.qty), fmt(l.left) + (l.last ? ' (máy cuối)' : '')]); slipKind.push(''); });
        });
        const list = d.rows.filter(r => r.ask > 0 || r.moveIn > 0).sort((a, b) => a.shop.localeCompare(b.shop) || b.ask - a.ask || (a.cover ?? 0) - (b.cover ?? 0));
        const counts = {}; d.rows.forEach(r => { counts[r.status] = (counts[r.status] || 0) + 1; });
        const shown = list.slice(0, Math.max(o.top, 25));
        await reportImage({
            title: `CÂN HÀNG ${d.cat.toUpperCase()} · ${toBI(isoDate(new Date()))}`,
            sub: `Hàng Mới · bán ${toBI(d.range.from)}–${toBI(d.range.to)} (${view.balDays} ngày) · giữ đủ ${d.target} ngày + ${d.spare} máy dự phòng${view.balBrands.size ? ' · hãng ' + d.brands.filter(b => view.balBrands.has(b.key)).map(b => b.label).join(', ') : ''}`,
            cards: [{ t: 'Xin kho', v: `SL ${fmt(d.rows.reduce((a, r) => a + r.ask, 0))}`, s: `${fmt(d.rows.filter(r => r.ask > 0).length)} mã`, main: true },
                ...(hasPlan ? [{ t: 'Đã có nơi cho', v: `SL ${fmt(got)}`, s: `${fmt(groups.length)} siêu thị cho` }, { t: 'Còn thiếu', v: `SL ${fmt(shortAll)}`, s: 'xin kho tổng', d: shortAll ? 'thiếu nơi cho' : '' }] : []),
                ...(multi ? [{ t: 'Chuyển nội cụm', v: `SL ${fmt(d.moves.reduce((a, m) => a + m.qty, 0))}`, s: `${fmt(d.moves.length)} lượt` }] : []),
                { t: 'Hết hàng', v: fmt(counts['Hết hàng'] || 0), s: 'mã', d: counts['Hết hàng'] ? 'thiếu hàng bán' : '' }, { t: 'Sắp hết', v: fmt(counts['Sắp hết'] || 0), s: 'mã' }, { t: 'Tồn nhiều', v: fmt(counts['Tồn nhiều'] || 0), s: 'mã' }],
            sections: [
                { title: `Cần bổ sung${list.length > shown.length ? ` (${shown.length}/${list.length} dòng)` : ''}`, cols: [['Siêu thị', 1.6], ['Sản phẩm', 5], ['Tồn', 0.8, 'right'], ...(d.hasTransit ? [['Đang về', 1, 'right']] : []), ['TB/ngày', 1, 'right'], ['Đủ bán', 1, 'right'], ...(multi ? [['Nhận nội cụm', 1.3, 'right']] : []), ['Xin kho', 1, 'right'], ...(hasPlan ? [['Xin từ (đề xuất)', 5.5]] : [])],
                    rows: shown.map(r => [shopName(r.shop), r.name, fmt(r.stock), ...(d.hasTransit ? [r.transit ? fmt(r.transit) : ''] : []), fmt(r.perDay), r.cover == null ? '—' : fmt(r.cover) + ' ngày', ...(multi ? [r.moveIn ? fmt(r.moveIn) : ''] : []), r.ask ? fmt(r.ask) : '',
                        ...(hasPlan ? [r.ask ? (planOf(plans, r) ? planText(planOf(plans, r)) : 'chưa tra') : ''] : [])]),
                    cellColor: (i, j) => hasPlan && j === 0 + (d.hasTransit ? 7 : 6) + (multi ? 1 : 0) && planOf(plans, shown[i])?.short ? '#b91c1c' : '',
                    rowColor: i => shown[i].status === 'Hết hàng' ? '#fdecea' : '' },
                hasPlan && slipRows.length ? { title: `Phiếu xin theo nơi cho (gửi từng siêu thị)${groups.length > 15 ? ` · 15/${groups.length} nơi` : ''}`, cols: [['Nơi cho → về siêu thị', 3.2], ['Sản phẩm', 5.2], ['Xin', 0.8, 'right'], ['Nơi cho còn', 1.6, 'right']],
                    rows: slipRows, spanFirst: true, rowBold: i => slipKind[i] === 'h', rowColor: i => slipKind[i] === 'h' ? '#e6f4ea' : '' } : null,
                multi && d.moves.length ? { title: 'Chuyển nội cụm', cols: [['Sản phẩm', 5], ['Từ', 1.8], ['Đến', 1.8], ['SL', 0.8, 'right']], rows: d.moves.slice(0, 20).map(m => [m.name, shopName(m.from), shopName(m.to), fmt(m.qty)]) } : null],
            footer: `AutoBI Kho & Xuất Bán V${VERSION} · tồn lúc ${stamp(new Date(view.inv.capturedAt))}${hasPlan ? ` · nguồn CRM lúc ${stamp(new Date(view.crmBal.at))}, nơi cho giữ lại ${srcKeep()} máy` : ''}${d.missing ? ` · CHƯA ĐỦ: còn ${d.missing} ngày chưa đổ xuất bán` : ''}`, warn: !!d.missing,
            width: hasPlan ? 1240 : 1000,
            name: `AutoBI_CanHang_${isoDate(new Date()).replace(/-/g, '')}.png`
        });
    }
    async function pendingImage(orders) {
        invariant(orders.length, 'Không có đơn treo trong kỳ đang xem');
        const r = view.sales, list = orders.slice(0, 30);
        await reportImage({
            title: `ĐƠN TREO · tính đến ${toBI(isoDate(new Date()))}`,
            sub: `Kỳ ${toBI(r.range.from)}–${toBI(r.range.to)} · đơn chưa hủy nhưng chưa xuất hoặc chưa giao · đỏ: treo từ 3 ngày · vàng: 2 ngày`,
            cards: [{ t: 'Đơn treo', v: fmt(orders.length), s: mil(orders.reduce((a, o) => a + o.value, 0)), main: true },
                { t: 'Treo từ 2 ngày', v: fmt(orders.filter(o => o.age >= 2).length), s: 'đơn', d: orders.some(o => o.age >= 2) ? 'cần xử lý' : '' },
                ...config.shops.filter(sh => view.salesShops.has(keyCode(sh.code))).map(sh => ({ t: sh.name, v: fmt(orders.filter(o => o.shop === keyCode(sh.code)).length), s: 'đơn' }))],
            sections: [{ title: `Danh sách${orders.length > list.length ? ` (${list.length}/${orders.length} đơn treo lâu nhất)` : ''}`, cols: [['Treo', 0.9], ['Siêu thị', 1.4], ['Nhân viên', 2.4], ['Sản phẩm', 4], ['Giá trị', 1.2, 'right'], ['Trạng thái', 1.3]],
                rows: list.map(o => [o.age ? `${o.age} ngày` : 'hôm nay', shopName(o.shop), o.creator.replace(/^\d+\s*[-–]\s*/, ''), o.items.join(' + '), mil(o.value), o.state]),
                rowColor: i => list[i].age >= 3 ? '#fdecea' : list[i].age === 2 ? '#fff7e0' : '' }],
            footer: `AutoBI Kho & Xuất Bán V${VERSION} · lấy lúc ${stamp(new Date())}`,
            name: `AutoBI_DonTreo_${isoDate(new Date()).replace(/-/g, '')}.png`
        });
    }

    /* ---------- Excel ---------- */
    function exportExcel() {
        const X = XL(); invariant(X, 'Chưa tải được thư viện Excel (SheetJS)');
        const wb = X.utils.book_new();
        // V2.0: số có dấu phân cách, cột % hiện đúng %, dòng tiêu đề có nút lọc (AutoFilter)
        const sheet = (name, aoa, widths) => {
            const ws = X.utils.aoa_to_sheet(aoa); if (widths) ws['!cols'] = widths.map(w => ({ wch: w }));
            const h = aoa.findIndex(r => r.length >= 3);
            if (h >= 0) {
                const heads = aoa[h].map(x => String(x ?? ''));
                for (let r = h + 1; r < aoa.length; r++) aoa[r].forEach((v, c) => {
                    if (typeof v !== 'number') return;
                    const cell = ws[X.utils.encode_cell({ r, c })]; if (!cell) return;
                    cell.z = /%|tỷ lệ/i.test(heads[c]) ? '0.0%' : Number.isInteger(v) ? '#,##0' : '#,##0.0#';
                });
                if (aoa.length > h + 1) ws['!autofilter'] = { ref: X.utils.encode_range({ s: { r: h, c: 0 }, e: { r: aoa.length - 1, c: heads.length - 1 } }) };
            }
            X.utils.book_append_sheet(wb, ws, name);
        };
        if (view.tab === 'sales') {
            const r = salesData(); invariant(r, 'Chưa có số xuất bán'); const s = r.summary, codes = config.shops.map(x => keyCode(x.code)).filter(c => view.salesShops.has(c));
            const title = `Kỳ ${toBI(r.range.from)}–${toBI(r.range.to)} · ${codes.map(shopName).join(', ')} · Kho tạo · ngành ${view.salesCat || 'tất cả'}${view.salesBrands.size ? ' · hãng ' + [...new Set(r.lines.map(l => l.brand))].join(', ') : ''} · theo ${r.basis === 'shipped' ? 'ngày xuất' : 'ngày tạo'} · Đã xuất – Đã giao – Chưa hủy · loại hàng ${view.salesConds.size ? [...view.salesConds].join(', ') : 'tất cả'} · bỏ đơn khách nhập trả` + (r.missing ? ` · CHƯA ĐỦ: còn ${r.missing} ngày chưa lấy` : '');
            sheet('TheoHang', [['Bán theo hãng · ' + title], [], ['Hãng', ...codes.flatMap(c => [shopName(c) + ' SL', shopName(c) + ' DT (đ)']), 'Cụm SL', '% SL', 'Cụm DT (đ)', '% DT'],
                ...s.brands.map(b => [b.label, ...codes.flatMap(c => { const x = s.brandShop[norm(b.label) + '|' + c]; return [x ? x.quantity : 0, x ? x.revenue : 0]; }), b.quantity, b.pctQty / 100, b.revenue, b.pctRev / 100]),
                ['Tổng', ...codes.flatMap(c => { const x = s.shops.find(z => z.code === c); return [x ? x.quantity : 0, x ? x.revenue : 0]; }), s.quantity, 1, s.revenue, 1]], [16, ...codes.flatMap(() => [12, 15]), 9, 8, 15, 8]);
            const vc = salesByCategory(r.lines, codes);
            sheet('TheoNganh', [['Bán theo ngành hàng · ' + title], [], ['Ngành hàng', 'Nhóm hàng', ...codes.map(c => shopName(c) + ' SL'), 'SL', 'Doanh thu (đ)', '% DT'],
                ...vc.rows.map(x => [x.category, x.group, ...codes.map(c => x.byShop[c] || 0), x.qty, x.rev, vc.rev ? x.rev / vc.rev : 0]),
                ['Tổng', '', ...codes.map(c => vc.byShop[c] || 0), vc.qty, vc.rev, 1]], [24, 30, ...codes.map(() => 10), 8, 15, 8]);
            const m = staffMatrix(s);
            sheet('NhanVien_x_Hang', [['Nhân viên × hãng (SL) · ' + title], [], ['Siêu thị', 'Nhân viên (người tạo)', ...m.brands, 'Tổng SL', '% shop', 'Doanh thu (đ)'],
                ...m.rows.map(x => [shopName(x.staff.shop), x.staff.label, ...x.cells, x.staff.quantity, x.staff.pctShop / 100, x.staff.revenue]),
                ['Tổng', '', ...s.brands.map(b => b.quantity), s.quantity, '', s.revenue]], [14, 28, ...m.brands.map(() => 9), 8, 8, 15]);
            const att = attachData(r.allLines.filter(l => view.salesShops.has(l.shop)));
            sheet('BanKem', [['Bán kèm theo nhân viên (đơn có điện thoại; kèm = ' + ((config.attachGroups || []).length ? config.attachGroups.join(', ') : 'mọi nhóm không phải điện thoại') + ') · ' + title], [], ['Siêu thị', 'Nhân viên', 'Đơn ĐT', 'SL ĐT', 'Có kèm', 'Tỷ lệ kèm', 'SL kèm', 'DT kèm (đ)'],
                ...att.staff.map(x => [shopName(x.shop), x.label, x.phoneOrders, x.phoneQty, x.attached, x.rate / 100, x.attachQty, x.attachRev]),
                ['Tổng', '', att.total.phoneOrders, att.total.phoneQty, att.total.attached, att.total.rate / 100, att.total.attachQty, att.total.attachRev]], [14, 28, 8, 8, 8, 9, 8, 15]);
            sheet('NhanVien_SP', [['Nhân viên bán sản phẩm gì · ' + title], [], ['Siêu thị', 'Nhân viên', 'Hãng', 'Mã SP', 'Tên sản phẩm', 'SL', 'Doanh thu (đ)'],
                ...s.staff.flatMap(st => st.products.map(p => [shopName(st.shop), st.label, p.brand, p.product, p.label, p.quantity, p.revenue]))], [14, 28, 12, 16, 44, 6, 15]);
            sheet('SanPham', [['Sản phẩm · ' + title], [], ['Mã SP', 'Tên sản phẩm', 'Hãng', 'Nhóm', 'SL', '% SL', 'Doanh thu (đ)'], ...s.products.map(p => [p.product, p.name, p.brand, p.group, p.quantity, p.pctQty / 100, p.revenue])], [16, 44, 12, 22, 6, 7, 15]);
            sheet('TheoNgay', [['Ngày', 'Sổ', 'SL', 'Doanh thu (đ)', 'Đơn treo', 'Dòng BI', 'Dòng tính', 'Dòng bị loại', 'Lấy lúc'], ...r.days.map(d => {
                const L = r.lines.filter(l => (r.basis === 'shipped' && l.shipped ? l.shipped : l.created) === d.day);
                return [toBI(d.day), d.status, L.reduce((a, l) => a + l.qty, 0), L.reduce((a, l) => a + l.qty * l.price, 0), d.pending ?? '', d.rows ?? '', d.valid ?? '', d.excluded ? Object.entries(d.excluded).map(([k, v]) => `${k} ${v}`).join(', ') : '', d.at ?? ''];
            })], [11, 12, 6, 15, 8, 8, 9, 30, 17]);
            sheet('DonTreo', [['Ngày tạo', 'Giờ', 'Siêu thị', 'Mã đơn', 'Loại YCX', 'Nhân viên', 'Hãng', 'Mã SP', 'Sản phẩm', 'IMEI', 'Loại hàng', 'SL', 'Giá bán', 'Xuất', 'Giao'],
                ...r.pendingLines.map(l => [toBI(l.created), l.time, shopName(l.shop), l.order, l.orderType, l.creator, l.brand, l.product, l.productName, l.imei || '', conditionText(l.condition), l.qty, l.price, l.exported, l.delivered])], [11, 6, 14, 20, 24, 26, 10, 16, 40, 18, 12, 5, 12, 10, 10]);
            sheet('NhapTra', [['Đơn khách nhập trả — không tính vào số bán'], [], ['Ngày tạo', 'Ngày xuất', 'Siêu thị', 'Mã đơn', 'Nhân viên', 'Hãng', 'Mã SP', 'Sản phẩm', 'IMEI', 'Loại hàng', 'SL', 'Giá bán', 'Trả hàng'],
                ...r.returnedLines.map(l => [toBI(l.created), l.shipped ? toBI(l.shipped) : '', shopName(l.shop), l.order, l.creator, l.brand, l.product, l.productName, l.imei || '', conditionText(l.condition), l.qty, l.price, l.returned])], [11, 11, 14, 20, 26, 10, 16, 40, 18, 12, 5, 12, 10]);
            const hm = hourMatrix(r.lines);
            sheet('TheoGio', [['Số đơn theo giờ tạo đơn × thứ · ' + title], [], ['Giờ', ...WEEKDAYS, 'Cả tuần'],
                ...hm.hours.map(h => { const v = WEEKDAYS.map((_, w) => hm.cells.find(c => c.w === w && c.h === h)?.orders || 0); return [`${pad(h)}:00`, ...v, v.reduce((a, x) => a + x, 0)]; })], [8, ...WEEKDAYS.map(() => 7), 9]);
            sheet('Ban_ChiTiet', [['Ngày tạo', 'Giờ', 'Ngày xuất', 'Kho tạo', 'Siêu thị', 'Mã đơn', 'Loại YCX', 'Nhân viên', 'Hãng', 'Mã SP', 'Tên SP', 'IMEI', 'Loại hàng', 'Ngành', 'Nhóm', 'SL', 'Giá bán', 'Giá trước VAT', 'Doanh thu', 'Thanh toán', 'Giao hàng'],
                ...r.lines.map(l => [toBI(l.created), l.time, l.shipped ? toBI(l.shipped) : '', l.shop, shopName(l.shop), l.order, l.orderType, l.creator, l.brand, l.product, l.productName, l.imei || '', conditionText(l.condition), l.category, l.group, l.qty, l.price, l.priceNet, l.qty * l.price, l.pay || '', l.ship || ''])], [11, 6, 11, 8, 14, 20, 24, 26, 10, 16, 40, 18, 12, 22, 20, 5, 12, 12, 13, 18, 20]);
            X.writeFile(wb, `AutoBI_XuatBan_${r.range.from.replace(/-/g, '')}-${r.range.to.replace(/-/g, '')}.xlsx`);
        } else if (view.tab === 'request') {
            const R = reqState(); invariant(R.results.length, 'Chưa có kết quả Check xin hàng');
            sheet('CheckXinHang', [[`Check xin hàng cho ${shopName(R.shop)} · ${stamp(new Date(R.at))} · tồn bán được trên CRM, không tính kho chi nhánh`], [], ['Mã SP', 'Sản phẩm', 'Ưu tiên', 'Mã ST', 'Siêu thị', 'Khu vực', 'Tồn bán được', 'Đang khóa', 'Đề xuất xin'],
                ...reqBlocks(R, R.shop).flatMap(b => { const q = c => b.alloc.plan.find(x => x.code === c)?.qty || ''; return b.list.map(x => [b.p.code, b.p.name, x.tierName, x.code, x.name, areaText(x.area), x.free, x.lock, q(x.code)])
                    .concat((b.nb || []).map(x => [b.p.code, b.p.name, 'Tỉnh lân cận', x.code, x.name, x.provName, x.free, x.lock, q(x.code)])); })], [16, 40, 14, 8, 30, 22, 12, 10, 11]);
            const alts = reqBlocks(R, R.shop).filter(b => b.alts?.length && !b.list.some(x => x.tier <= 2));
            if (alts.length) sheet('HangThayThe', [['Hàng thay thế (cùng model, khác màu / dung lượng) còn tồn bán được trong tỉnh'], [], ['Mã SP hết', 'Sản phẩm hết', 'Mã thay thế', 'Sản phẩm thay thế', 'Gần: số nơi', 'Gần: số máy', 'Cả tỉnh: số nơi', 'Cả tỉnh: số máy', 'Nơi gần nhất'],
                ...alts.flatMap(b => b.alts.map(a => [b.p.code, b.p.name, a.p.code, a.p.name, a.near.length, a.nearQty, a.list.length, a.allQty, a.list.slice(0, 3).map(x => `${x.name} ${x.free}`).join(', ')]))], [16, 36, 16, 36, 10, 10, 12, 12, 40]);
            X.writeFile(wb, `AutoBI_CheckXinHang_${isoDate(new Date()).replace(/-/g, '')}.xlsx`);
        } else if (view.tab === 'balance') {
            const d = balanceData(); invariant(view.inv, 'Chưa có số tồn kho');
            const xPlans = balancePlans(d.rows);
            sheet('CanHang', [[`Cân hàng ${d.cat} (hàng Mới)${view.balBrands.size ? ' · hãng ' + d.brands.filter(b => view.balBrands.has(b.key)).map(b => b.label).join(', ') : ''}${view.balStatus ? ' · ' + (view.balStatus === 'need' ? 'Cần xin' : view.balStatus) : ''} · bán ${toBI(d.range.from)}–${toBI(d.range.to)} (${view.balDays} ngày) · giữ đủ ${d.target} ngày · tồn lúc ${stamp(new Date(view.inv.capturedAt))}` + (d.missing ? ` · CHƯA ĐỦ: còn ${d.missing} ngày chưa đổ xuất bán` : '')], [],
                ['Siêu thị', 'Hãng', view.balModel ? 'Model' : 'Mã SP', 'Tên sản phẩm', 'Tồn', 'Đang về', `Bán ${view.balDays} ngày`, 'TB / ngày', 'Đủ bán (ngày)', 'Cần có', 'Thiếu', 'Nhận nội cụm', 'Cho nội cụm', 'Xin kho', 'Trạng thái', 'Xin từ (đề xuất)', 'Có tại (CRM)'],
                ...balStatusRows(d).map(r => [shopName(r.shop), r.brand, view.balModel ? r.codeCount + ' mã' : r.product, r.name, r.stock, r.transit || 0, r.sold, r.perDay, r.cover ?? '', r.want, r.need, r.moveIn, r.moveOut, r.ask, r.status, r.ask ? planText(planOf(xPlans, r)) : '', srcText(crmSrc(r), 5)])], [14, 12, 16, 44, 7, 10, 9, 9, 11, 8, 8, 11, 11, 8, 11, 44, 40]);
            const xGroups = groupBySource(planItems(d.rows, xPlans));
            if (xGroups.length) sheet('XinTheoNoiCho', [[`Đề xuất xin theo nơi cho · nguồn CRM lúc ${stamp(new Date(view.crmBal.at))} · nơi cho giữ lại ${srcKeep()} máy`], [],
                ['Mã nơi cho', 'Nơi cho', 'Khu vực', 'Mã SP', 'Sản phẩm', 'SL xin', 'Nơi cho còn lại', 'Máy cuối', 'Về siêu thị'],
                ...xGroups.flatMap(g => g.lines.map(l => [g.code, g.name, srcPlace(g), l.product, l.name, l.qty, l.left, l.last ? 'có' : '', shopName(l.to)]))], [9, 26, 22, 16, 40, 7, 12, 9, 14]);
            if (d.moves.length) sheet('ChuyenNoiCum', [['Gợi ý chuyển hàng nội cụm (cùng mã, nơi thừa → nơi thiếu)'], [],
                ['Mã SP', 'Tên sản phẩm', 'Hãng', 'Từ siêu thị', 'Tồn nơi cho', 'Đủ bán nơi cho (ngày)', 'Đến siêu thị', 'Tồn nơi nhận', 'Đủ bán nơi nhận (ngày)', 'SL chuyển'],
                ...d.moves.map(m => [m.product, m.name, m.brand, shopName(m.from), m.fromStock, m.fromCover ?? '', shopName(m.to), m.toStock, m.toCover ?? '', m.qty])], [16, 44, 12, 14, 10, 12, 14, 10, 12, 9]);
            X.writeFile(wb, `AutoBI_CanHang_${isoDate(new Date()).replace(/-/g, '')}.xlsx`);
        } else {
            invariant(view.inv, 'Chưa có số tồn kho');
            const recs = invRecords(), v = inventoryViews(recs), codes = view.inv.shops.filter(c => view.invShops.has(c)), cond = v.conditionList;
            const f = view.invFilter, loc = [f.category && 'Ngành ' + f.category, f.group && 'Nhóm ' + f.group, f.brand && 'Hãng ' + f.brand, f.conditions.length && 'Trạng thái ' + f.conditions.join(', '), f.q && 'Tìm "' + f.q + '"'].filter(Boolean).join(' · ') || 'Không lọc';
            const title = `Tồn lúc ${stamp(new Date(view.inv.capturedAt))} · ${codes.map(shopName).join(', ')} · ${loc}`;
            sheet('TonKho_NhomHang', [[title], [], ['Ngành', 'Nhóm hàng', ...codes.map(shopName), 'SL tồn', ...cond, 'Giá trị (đ)'],
                ...v.groups.map(g => [g.category, g.group, ...codes.map(c => g.byShop[c] || 0), g.quantity, ...cond.map(c => g.byCond[c] || 0), Math.round(g.cost)])], [24, 30, ...codes.map(() => 10), 8, ...cond.map(() => 10), 16]);
            sheet('TonKho_SanPham', [[title], [], ['Nhóm hàng', 'Hãng', 'Mã SP', 'Tên sản phẩm', ...codes.map(shopName), 'SL tồn', ...cond, 'Giá trị (đ)'],
                ...v.products.map(p => [p.group, p.brand, p.product, p.name, ...codes.map(c => p.byShop[c] || 0), p.quantity, ...cond.map(c => p.byCond[c] || 0), Math.round(p.cost)])], [24, 12, 16, 44, ...codes.map(() => 10), 8, ...cond.map(() => 10), 16]);
            sheet('TonKho_IMEI', [['Mã ST', 'Siêu thị', 'Ngành', 'Nhóm hàng', 'Hãng', 'Mã SP', 'Tên sản phẩm', 'IMEI / Serial', 'Trạng thái', 'SL', 'Giá vốn (đ)', 'Ngày nhập', 'Tuổi tồn (ngày)'],
                ...recs.map(r => [r.shop, shopName(r.shop), r.category, r.group, r.brand, r.product, r.productName, r.serial, r.condition, r.qty, Math.round(r.cost), r.input, ageDays(r.input, isoDate(new Date())) ?? ''])], [7, 14, 22, 26, 12, 16, 40, 20, 16, 6, 13, 16, 9]);
            const tdy = isoDate(new Date()), aged = recs.map(r => ({ r, age: ageDays(r.input, tdy) })).filter(x => x.age != null && x.age >= (config.ageAlert || 60)).sort((a, b) => b.age - a.age);
            sheet('TuoiTon', [[`Hàng tồn từ ${config.ageAlert || 60} ngày · ${title}`], [], ['Siêu thị', 'Nhóm hàng', 'Hãng', 'Mã SP', 'Tên sản phẩm', 'IMEI / Serial', 'Trạng thái', 'SL', 'Giá vốn (đ)', 'Ngày nhập', 'Tuổi (ngày)'],
                ...aged.map(({ r, age }) => [shopName(r.shop), r.group, r.brand, r.product, r.productName, r.serial, r.condition, r.qty, Math.round(r.cost), r.input, age])], [14, 26, 12, 16, 40, 20, 16, 6, 13, 16, 9]);
            const tr = (view.inv.transit || []).filter(r => view.invShops.has(r.shop));
            if (tr.length) sheet('HangDangVe', [['Mã ST', 'Siêu thị', 'Ngành', 'Nhóm hàng', 'Hãng', 'Mã SP', 'Tên sản phẩm', 'IMEI / Serial', 'Trạng thái', 'SL', 'Giá vốn (đ)', 'Ngày nhập'],
                ...tr.map(r => [r.shop, shopName(r.shop), r.category, r.group, r.brand, r.product, r.productName, r.serial, r.condition, r.qty, Math.round(r.cost), r.input])], [7, 14, 22, 26, 12, 16, 40, 20, 16, 6, 13, 16]);
            X.writeFile(wb, `AutoBI_TonKho_${isoDate(new Date(view.inv.capturedAt)).replace(/-/g, '')}.xlsx`);
        }
    }

    /* ---------- Cài đặt ---------- */
    async function copyText(text, done) {
        try { await navigator.clipboard.writeText(text); }
        catch { const t = el('textarea', undefined, document.body); t.value = text; t.dataset.kxbUi = ''; t.select(); document.execCommand('copy'); t.remove(); }
        status(done, 'ok');
    }
    async function copyLog() {
        await copyText(journal.slice(-300).reverse().map(x => `${new Date(x.time).toLocaleString('vi-VN')} ${x.kind === 'error' ? '✖' : '•'} ${x.message}`).join('\n'), 'Đã sao chép nhật ký');
    }
    // Tin nhắc đơn treo gửi nhóm: ưu tiên đơn treo từ 2 ngày, gom theo siêu thị
    function pendingMessage(orders, today) {
        invariant(orders.length, 'Không có đơn treo trong kỳ đang xem');
        const list = orders.some(o => o.age >= 2) ? orders.filter(o => o.age >= 2) : orders;
        const head = orders.some(o => o.age >= 2) ? `⚠️ ĐƠN TREO TỪ 2 NGÀY (tính đến ${toBI(today)})` : `⏳ ĐƠN CHƯA XUẤT / CHƯA GIAO (${toBI(today)})`;
        const byShop = new Map(); list.forEach(o => { (byShop.get(o.shop) || byShop.set(o.shop, []).get(o.shop)).push(o); });
        // Gom theo siêu thị → nhân viên: ai có đơn treo thấy ngay tên mình
        const who = o => o.creator.replace(/^\d+\s*[-–]\s*/, '') || o.creator;
        return [head, ...[...byShop.entries()].flatMap(([shop, os]) => {
            const byStaff = new Map(); os.forEach(o => { (byStaff.get(who(o)) || byStaff.set(who(o), []).get(who(o))).push(o); });
            return [`\n🏬 ${shopName(shop)} — ${os.length} đơn:`, ...[...byStaff.entries()].sort((x, y) => y[1].length - x[1].length).flatMap(([name, xs]) => [
                `👤 ${name} (${xs.length} đơn)`, ...xs.map(o => `   • ${o.items.join(' + ')} · ${mil(o.value)} · treo ${o.age} ngày · ${o.state.toLowerCase()} · đơn ${o.order.trim()}`)])];
        }), '\nAnh em kiểm tra, xuất / giao hoặc hủy đơn giúp nhé.'].join('\n');
    }
    function renderStorage() {
        const box = ui.querySelector('[data-storage]'); if (!box) return;
        try {
            const keys = listMonths();
            let bytes = 0, lines = 0; const per = [];
            keys.forEach(m => { const v = loadMonth(m); const b = JSON.stringify(v).length; bytes += b; lines += (v.lines || []).length; per.push([m, b, (v.lines || []).length, Object.keys(v.book || {}).length]); });
            const mb = bytes / 1048576, inv = JSON.stringify(load('lastInventory', null) || '').length / 1048576, idb = mstore.mode === 'idb';
            const legacy = idb ? gmMonthKeys() : [];
            box.replaceChildren();
            el('div', `Bộ nhớ trên máy (${idb ? 'IndexedDB' : 'Tampermonkey' + (mstore.error ? ' — IndexedDB lỗi: ' + mstore.error : '')}): xuất bán ${keys.length} tháng · ${fmt(lines)} dòng · ${fmt(Math.round(mb * 10) / 10)} MB · tồn kho ${fmt(Math.round(inv * 10) / 10)} MB`
                + (!idb && mb + inv > 40 ? ' — khá lớn, Tampermonkey có thể chậm khi lưu: nên sao lưu rồi xóa tháng cũ.' : ''), box).style.color = !idb && mb + inv > 40 ? '#8a1c1c' : '';
            if (legacy.length) {
                const lb = el('div', `Bản dự phòng cũ trong Tampermonkey: ${legacy.length} tháng (không còn cập nhật, chỉ dùng khi IndexedDB bị xóa). `, box);
                const bx = el('button', 'Xóa bản dự phòng cũ', lb, 'mini'); bx.type = 'button'; bx.title = 'Giải phóng Tampermonkey — nên Sao lưu dữ liệu trước';
                bx.onclick = safely(() => { invariant(!running, 'Đang đổ số, chờ xong'); if (!window.confirm('Xóa bản dự phòng số xuất bán cũ trong Tampermonkey? Số trong IndexedDB vẫn giữ nguyên.')) return; legacy.forEach(k => GM_deleteValue(k)); log(`Đã xóa ${legacy.length} tháng dự phòng cũ trong Tampermonkey`); renderStorage(); });
            }
            const list = el('div', undefined, box, 'group'); list.style.marginTop = '6px';
            per.forEach(([m, b, n, dcount]) => {
                const chip = el('span', `${m.split('-').reverse().join('/')} · ${dcount} ngày · ${fmt(n)} dòng · ${fmt(Math.round(b / 1048576 * 10) / 10)} MB `, list, 'kxb-month');
                const x = el('button', '✕', chip, 'mini'); x.type = 'button'; x.title = 'Xóa số đã lưu của tháng này trên máy';
                x.onclick = safely(() => {
                    invariant(!running, 'Đang đổ số, chờ xong rồi xóa');
                    if (!window.confirm(`Xóa số xuất bán đã lưu tháng ${m.split('-').reverse().join('/')} trên máy này?\nLần sau xem kỳ này sẽ phải đổ lại từ BI. Nên "Sao lưu" trước.`)) return;
                    deleteMonth(m); log(`Đã xóa số đã lưu tháng ${m}`); status(`Đã xóa tháng ${m.split('-').reverse().join('/')}`, 'ok');
                    renderStorage(); if (view.sales) { view.sales = salesView(view.sales.range); renderSales(); }
                });
            });
        } catch (e) { box.textContent = 'Không đọc được dung lượng: ' + e.message; }
    }
    // Sao lưu toàn bộ số đã lưu (xuất bán từng tháng, tồn kho, cài đặt) ra 1 file để chuyển sang máy khác
    function backupData() {
        const keys = (typeof GM_listValues === 'function' ? GM_listValues() : []).filter(k => k.startsWith(PREFIX) && !/journal|update$/.test(k) && !k.startsWith(PREFIX + 'month_'));
        const data = {}; keys.forEach(k => { data[k.slice(PREFIX.length)] = load(k.slice(PREFIX.length), null); });
        listMonths().forEach(m => { data['month_' + m] = loadMonth(m); keys.push('month_' + m); });
        const blob = new Blob([JSON.stringify({ app: 'AutoBI_Kho_XuatBan', version: VERSION, exportedAt: new Date().toISOString(), data })], { type: 'application/json' });
        const name = `AutoBI_KhoXuatBan_SaoLuu_${isoDate(new Date()).replace(/-/g, '')}.json`;
        const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000);
        log(`Sao lưu ${keys.length} mục (${fmt(Math.round(blob.size / 1048576 * 10) / 10)} MB) → ${name}`); status(`Đã tải file sao lưu ${name}`, 'ok');
    }
    async function restoreData(file) {
        invariant(!running, 'Đang đổ số, chờ xong rồi khôi phục');
        invariant(file, 'Chưa chọn file');
        let pack; try { pack = JSON.parse(await file.text()); } catch { fail('File không đọc được (không phải file sao lưu AutoBI)'); }
        invariant(pack && pack.app === 'AutoBI_Kho_XuatBan' && pack.data && typeof pack.data === 'object', 'Không phải file sao lưu AutoBI Kho & Xuất Bán');
        let months = 0, days = 0;
        for (const [k, v] of Object.entries(pack.data)) {
            if (/^month_\d{4}-\d{2}$/.test(k)) { const mo = k.slice(6), m = mergeMonth(loadMonth(mo), v); days += m.taken; months++; saveMonth(mo, { book: m.book, lines: m.lines }); }
            else if (k === 'lastInventory' && v && Array.isArray(v.records)) { const cur = load('lastInventory', null); if (!cur || String(v.capturedAt) > String(cur.capturedAt)) { save(k, v); view.inv = v; } }
            else if (k === 'config' && v && Array.isArray(v.shops) && !config.shops.length) { config = { ...defaults, ...v }; save('config', config); }
        }
        log(`Khôi phục từ ${file.name}: ${months} tháng, nhận ${days} ngày mới hơn số trên máy`);
        status(`Đã khôi phục: ${months} tháng, ${days} ngày được cập nhật${config.shops.length ? '' : ''}. Tải lại trang BI (F5) nếu cài đặt siêu thị vừa được nhận.`, 'ok');
        renderStorage(); if (config.shops.length) { view.sales = salesView(selectedRange()); renderAll(); }
    }
    function shopEditor(shop = { code: '', name: '' }) {
        const row = el('div', undefined, ui.querySelector('[data-shops]'), 'kxb-shop');
        for (const [key, label] of [['code', 'Mã ST'], ['name', 'Tên hiển thị']]) {
            const l = el('label', label, row), input = el('input', undefined, l); input.dataset.shopField = key; input.value = shop[key] || '';
        }
        const del = el('button', 'Xóa', row); del.type = 'button'; del.onclick = () => { invariant(!running, 'Không đổi cấu hình khi đang chạy'); row.remove(); };
    }
    function saveConfig() {
        invariant(!running, 'Không đổi cấu hình khi đang chạy');
        const shops = [...ui.querySelectorAll('.kxb-shop')].map(row => Object.fromEntries([...row.querySelectorAll('input')].map(input => [input.dataset.shopField, clean(input.value)])));
        validateShops(shops);
        const v = k => clean(ui.querySelector(`[data-cfg="${k}"]`).value);
        const fh = Math.round(Number(v('freshHours')));
        invariant(fh >= 0 && fh <= 24, 'Số giờ "vừa lấy" phải từ 0 đến 24');
        const rd = Math.round(Number(v('returnDays')));
        invariant(rd >= 2 && rd <= 35, 'Số ngày kiểm tra lại nhập trả phải từ 2 đến 35');
        const bt = Math.round(Number(v('balTarget')));
        invariant(bt >= 3 && bt <= 60, 'Số ngày giữ đủ bán (Cân hàng) phải từ 3 đến 60');
        const sp = Math.round(Number(v('balSpare')));
        invariant(sp >= 0 && sp <= 5, 'Số máy dự phòng (Cân hàng) phải từ 0 đến 5');
        const sk = Math.round(Number(v('srcKeep')));
        invariant(sk >= 0 && sk <= 5, 'Số máy nơi cho giữ lại phải từ 0 đến 5');
        const ag = Math.round(Number(v('ageAlert')));
        invariant(ag >= 15 && ag <= 365, 'Mốc cảnh báo tuổi tồn phải từ 15 đến 365 ngày');
        const top = Math.round(Number(ui.querySelector('[data-img="top"]').value));
        invariant(top >= 5 && top <= 30, 'Số dòng top trong ảnh phải từ 5 đến 30');
        const img = { cats: ui.querySelector('[data-img="cats"]').checked, brands: ui.querySelector('[data-img="brands"]').checked, staff: ui.querySelector('[data-img="staff"]').checked, top };
        config = { ...config, img, shops, selectors: config.selectors || {}, basis: v('basis') === 'shipped' ? 'shipped' : 'created', returnDays: rd, balTarget: bt, balSpare: sp, ageAlert: ag, srcKeep: sk,
            notify: ui.querySelector('[data-cfg="notify"]').checked, balIncoming: ui.querySelector('[data-cfg="balIncoming"]').checked, freshHours: fh };
        save('config', config);
        const codes = new Set(shops.map(s => keyCode(s.code)));
        view.invShops = new Set([...view.invShops].filter(c => codes.has(c))); if (!view.invShops.size) view.invShops = new Set(codes);
        view.salesShops = new Set(codes); view.balShops = new Set(codes); if (view.sales) view.sales = salesView(view.sales.range);
        log('Đã lưu cài đặt'); status('Đã lưu cài đặt', 'ok');
        ui.querySelector('[data-settings]').open = false; renderAll();
    }
    // Ô chọn tỉnh cho Tra tồn CRM (lưu ngay khi chọn; đổi tỉnh → lần tra sau tự lấy lại danh sách siêu thị)
    function renderCrmSetting() {
        const box = ui?.querySelector('[data-crm-prov]'); if (!box) return;
        const set = new Set(crmProvs());
        multiDrop(box, 'crm-prov', CRM_PROVINCES.map(([id, n]) => ({ key: id, name: n, label: n })), set, 'Chọn tỉnh', () => {
            if (!set.size) set.add('1014');
            config.crmProvinces = [...set]; save('config', config); renderCrmSetting();
        });
        const nbox = ui.querySelector('[data-crm-near]');
        if (nbox) {
            const nset = new Set(nearIds());
            multiDrop(nbox, 'crm-near', CRM_PROVINCES.filter(([id]) => !set.has(id)).map(([id, n]) => ({ key: id, name: provShort(id), label: n })), nset, 'Chưa chọn', () => {
                config.crmNearProvinces = [...nset]; if (!nset.size) delete config.crmNearProvinces; save('config', config); renderCrmSetting();
            });
        }
        const info = ui.querySelector('[data-crm-info]');
        info.textContent = crmStoreCacheOk() ? `Đang dùng ${crmStores().length} siêu thị (lấy lúc ${stamp(new Date(config.crmStoreCache.at))}, không tính kho chi nhánh)` : 'Chưa lấy danh sách theo tỉnh — lần tra đầu sẽ tự lấy (cần đăng nhập CRM)';
    }
    // V2.2.1: báo lỗi rõ — cuộn tới dòng trạng thái để không bị khuất khi đang cuộn xuống dưới
    function safely(fn) { return async (...args) => { try { await fn(...args); } catch (e) { status(e.message, 'err'); log(e.message, 'error'); try { ui.querySelector('[data-status]').scrollIntoView({ block: 'nearest', behavior: 'smooth' }); } catch { /* bỏ qua */ } } }; }
    function renderAll() {
        ui.querySelectorAll('[data-pane]').forEach(p => p.hidden = p.dataset.pane !== view.tab);
        ui.querySelectorAll('[data-tab]').forEach(b => { b.classList.toggle('on', b.dataset.tab === view.tab); b.setAttribute('aria-selected', b.dataset.tab === view.tab); });
        renderSales(); renderInventoryFilters(); renderInventory(); renderBalance(); renderRequest();
    }

    /* ---------- Thông báo bản mới ---------- */
    // V2.1.1: giống AutoBI Core — tự kiểm GitHub khi mở trang (sau 4 giây) và 30 phút/lần; có bản mới thì hiện dải vàng trên đầu trang
    const UPDATE_EVERY = 30 * 60 * 1000;
    let updateClosedFor = '';
    function fetchRemoteVersion() {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({ method: 'GET', url: UPDATE_URL + '?t=' + Date.now(), timeout: 15000, headers: { 'Cache-Control': 'no-cache' },
                onload: r => { const info = r.status === 200 ? parseRemoteScript(String(r.responseText || '')) : null; info ? resolve(info) : reject(new Error(`GitHub trả HTTP ${r.status}`)); },
                onerror: () => reject(new Error('Lỗi mạng khi đọc GitHub')), ontimeout: () => reject(new Error('GitHub phản hồi quá lâu')) });
        });
    }
    async function checkUpdate(force) {
        const cached = load('update', null);
        if (!force && cached && Date.now() - cached.at < UPDATE_EVERY) { renderUpdate(); return cached; }
        try {
            const info = { ...(await fetchRemoteVersion()), at: Date.now() };
            save('update', info); renderUpdate();
            if (force) status(newerVersion(info.version, VERSION) ? `Có bản mới V${info.version}` : `Đang dùng bản mới nhất (V${VERSION})`, newerVersion(info.version, VERSION) ? 'warn' : 'ok');
            if (newerVersion(info.version, VERSION)) log(`Có bản mới V${info.version} trên GitHub (đang dùng V${VERSION})`);
            return info;
        } catch (e) { if (force) status('Không kiểm tra được bản mới: ' + e.message, 'err'); return null; }
    }
    function renderUpdate() {
        if (!ui) return;
        const box = ui.querySelector('[data-update]'); box.replaceChildren();
        const info = load('update', null), launch = document.getElementById('kxb-launch');
        const hasNew = info && newerVersion(info.version, VERSION);
        if (launch) { launch.textContent = 'AutoBI · Kho & Xuất Bán' + (hasNew ? ' 🔔' : ''); launch.classList.toggle('new', !!hasNew); launch.title = hasNew ? `Có bản mới V${info.version}` : ''; }
        if (hasNew) {
            const b = el('div', undefined, box, 'kxb-update');
            el('b', `🔔 Có bản mới V${info.version} (đang dùng V${VERSION})`, b);
            const row = el('div', undefined, b, 'bar');
            const a = el('a', '⬇ Cập nhật ngay', row, 'btn primary'); a.href = UPDATE_URL + '?v=' + encodeURIComponent(info.version); a.target = '_blank'; a.rel = 'noopener';
            el('span', 'Tampermonkey mở trang cài → bấm "Cập nhật" → quay lại đây tải lại trang BI (F5).', row, 'kxb-muted');
        }
        updateBar(hasNew ? info.version : '');
    }
    // Dải vàng trên đầu trang BI (cùng kiểu với AutoBI Core): hiện cả khi chưa mở khung; ✕ thì ẩn tới lần tải trang sau
    function updateBar(remote) {
        const old = document.getElementById('kxb-update-bar');
        const hide = () => { document.getElementById('kxb-update-bar')?.remove(); if (document.body) document.body.style.paddingTop = ''; };
        if (!remote || updateClosedFor === remote) { if (old) hide(); return; }
        if (old && old.dataset.v === remote) return;
        if (old) hide();
        if (!document.body) return;
        const bar = document.createElement('div');
        bar.id = 'kxb-update-bar'; bar.dataset.v = remote; bar.dataset.kxbUi = ''; bar.setAttribute('data-html2canvas-ignore', 'true');
        bar.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:2147483646;background:#fdf3d7;border-bottom:1px solid #e9c46a;color:#7a4b00;font:14px/1.35 -apple-system,BlinkMacSystemFont,"Segoe UI",system-ui,sans-serif;display:flex;align-items:center;gap:12px;padding:8px 14px;flex-wrap:wrap;box-sizing:border-box';
        const txt = el('div', undefined, bar); txt.style.cssText = 'flex:1;min-width:180px';
        el('div', `Có bản AutoBI Kho & Xuất Bán mới: V${remote}`, txt).style.fontWeight = '800';
        el('div', `Máy đang dùng V${VERSION} · bấm Cập nhật ngay → Tampermonkey hỏi thì bấm Cập nhật → tải lại trang`, txt).style.fontSize = '12px';
        const go = el('button', 'Cập nhật ngay', bar); go.type = 'button';
        go.style.cssText = 'min-height:36px;padding:0 16px;border-radius:8px;border:1px solid #b7791f;background:#fff;color:#7a4b00;font-weight:800;cursor:pointer';
        go.onclick = () => window.open(UPDATE_URL + '?v=' + encodeURIComponent(remote), '_blank');
        const x = el('button', '✕', bar); x.type = 'button'; x.setAttribute('aria-label', 'Đóng');
        x.style.cssText = 'min-height:36px;min-width:36px;border-radius:8px;border:1px solid #e9c46a;background:transparent;color:#7a4b00;font-weight:800;cursor:pointer';
        x.onclick = () => { updateClosedFor = remote; hide(); };
        document.body.appendChild(bar);
        document.body.style.paddingTop = bar.offsetHeight + 'px';
    }

    /* ---------- Giao diện: khung lớn giữa màn hình (máy tính) ---------- */
    /* ---------- V2.5.1: cảnh báo cài trùng ----------
     * Lỗi đã gặp 08/10/2026: 2 dòng "AutoBI - Kho & Xuất Bán" trong Tampermonkey (1 dòng cài từ link, 1 dòng tạo bằng "+" rồi dán code).
     * Bản nào chạy trước chiếm nút; "Cập nhật ngay" chỉ đè vào dòng cài từ link → dải vàng báo bản mới mãi không hết. */
    function dupWarning(versions, running) {
        const list = versions.filter(Boolean), old = list.filter(v => v !== running);
        if (list.length < 2) return;
        const text = `⚠️ Máy đang cài ${list.length} bản AutoBI Kho & Xuất Bán (${list.map(v => 'V' + v).join(' và ')}) — đang chạy V${running}.`;
        const how = `Cách sửa: mở nút đang chạy → ⚙️ Cài đặt → 💾 Sao lưu dữ liệu · bấm biểu tượng Tampermonkey → Bảng điều khiển → bấm 🗑 ở dòng AutoBI - Kho & Xuất Bán có phiên bản ${old.length ? 'cũ (V' + [...new Set(old)].join(', V') + ')' : 'trùng'}, giữ 1 dòng bản mới nhất → tải lại trang (F5). Cài đặt thiếu thì 📂 Khôi phục từ file.`;
        log(text + ' ' + how, 'error');
        if (document.getElementById('kxb-dup-bar') || !document.body) return;
        const bar = document.createElement('div'); bar.id = 'kxb-dup-bar'; bar.dataset.kxbUi = ''; bar.setAttribute('data-html2canvas-ignore', 'true');
        bar.style.cssText = 'position:fixed;left:16px;bottom:16px;max-width:min(720px,calc(100vw - 32px));z-index:2147483646;background:#fdecea;border:1px solid #e08a80;color:#8a1c1c;font:13px/1.45 Arial,"Segoe UI",sans-serif;padding:10px 40px 10px 14px;border-radius:10px;box-shadow:0 6px 20px #0003';
        const b = document.createElement('div'); b.style.fontWeight = '700'; b.textContent = text; bar.append(b);
        const h = document.createElement('div'); h.textContent = how; bar.append(h);
        const x = document.createElement('button'); x.type = 'button'; x.textContent = '✕'; x.setAttribute('aria-label', 'Đóng');
        x.style.cssText = 'position:absolute;right:6px;top:6px;border:0;background:none;color:#8a1c1c;font-weight:700;cursor:pointer;font-size:14px';
        x.onclick = () => bar.remove(); bar.append(x);
        document.body.append(bar);
    }
    function mount() {
        if (document.querySelector('[data-kxb-ui]')) {
            // Một bản khác đã chiếm nút → bản này không mở giao diện, chỉ báo trùng (đọc phiên bản đang chạy từ khung của nó)
            const other = (clean(document.querySelector('#kxb-panel .top .kxb-muted')?.textContent).match(/V?([\d.]+)/) || [])[1] || '?';
            setTimeout(() => dupWarning([other, VERSION], other), 1500);
            return;
        }
        const style = el('style', `
        #kxb-launch{position:fixed;right:20px;bottom:20px;z-index:2147483645;background:#087f8c;color:#fff;padding:12px 20px;border:0;border-radius:24px;cursor:pointer;font:600 14px Arial,sans-serif;box-shadow:0 4px 14px #0003}
        #kxb-launch[hidden],#kxb-back[hidden]{display:none}
        #kxb-back{position:fixed;inset:0;z-index:2147483643;background:#0f172a66;display:flex;align-items:center;justify-content:center}
        #kxb-panel{box-sizing:border-box;margin:0;width:min(1480px,96vw);max-width:96vw;height:92vh;display:flex;flex-direction:column;background:#fff;color:#172a3a;border-radius:14px;box-shadow:0 20px 60px #0005;font:14px/1.45 Arial,sans-serif;overflow:clip;position:relative}
        #kxb-panel *{box-sizing:border-box}
        #kxb-panel .top{display:flex;align-items:center;gap:12px;padding:14px 22px;border-bottom:1px solid #e3e9ed}
        #kxb-panel .top strong{font-size:17px}#kxb-panel .top .sp{flex:1}
        #kxb-panel .tabs{display:flex;gap:4px;margin-left:16px}
        #kxb-panel .tabs button{border:0;background:none;padding:10px 18px;font:600 15px Arial,sans-serif;color:#4a5a66;border-bottom:3px solid transparent;cursor:pointer;border-radius:0}
        #kxb-panel .tabs button.on{color:#087f8c;border-bottom-color:#087f8c}
        #kxb-panel .body{flex:1;overflow-y:auto;overflow-x:hidden;padding:16px 22px}#kxb-panel .top{flex:none}
        #kxb-panel button{min-height:36px;padding:7px 14px;border:1px solid #bac9d1;border-radius:8px;background:#f2f7f8;color:#173047;cursor:pointer;font:inherit}
        #kxb-panel button:disabled{opacity:.45;cursor:default}
        #kxb-panel button.primary{background:#087f8c;color:#fff;border-color:#087f8c;font-weight:600}
        #kxb-panel button.danger{background:#b91c1c;color:#fff;border-color:#b91c1c}
        #kxb-panel .busy-only{display:none}#kxb-panel.busy .busy-only{display:inline-block}#kxb-panel.busy .idle-only{display:none}
        #kxb-panel input,#kxb-panel select{min-height:36px;padding:6px 8px;border:1px solid #b8c9ce;border-radius:8px;color:#172a3a;background:#fff;font:inherit}
        #kxb-panel .kxb-drop{position:relative;display:inline-block}
        #kxb-panel .kxb-drop>summary{list-style:none;cursor:pointer;min-width:320px;max-width:520px;min-height:36px;padding:7px 30px 7px 10px;border:1px solid #b8c9ce;border-radius:8px;background:#fff;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;position:relative}
        #kxb-panel .kxb-drop>summary::-webkit-details-marker{display:none}
        #kxb-panel .kxb-drop>summary::after{content:'▾';position:absolute;right:10px;top:7px}
        #kxb-panel .kxb-drop>summary.on{border-color:#087f8c;background:#e6f4f5;font-weight:600}
        #kxb-panel .kxb-drop .list{position:absolute;z-index:5;left:0;top:calc(100% + 4px);min-width:100%;max-height:360px;overflow:auto;background:#fff;border:1px solid #b8c9ce;border-radius:8px;box-shadow:0 8px 24px #0002;padding:4px 0}
        #kxb-panel .kxb-drop .list label{display:flex;flex-direction:row;align-items:center;justify-content:flex-start;gap:8px;margin:0;color:#172a3a;font-weight:400;padding:6px 12px;cursor:pointer;white-space:nowrap;font-size:14px}
        #kxb-panel .kxb-drop .list label:hover{background:#eef6f7}
        #kxb-panel .kxb-drop .list input{min-height:auto;width:16px;height:16px;padding:0}
        #kxb-panel select.on{border-color:#087f8c;background:#e6f4f5;font-weight:600}
        #kxb-panel label{display:inline-flex;flex-direction:column;gap:3px;font-size:12px;color:#4a5a66}
        #kxb-panel .bar{display:flex;flex-wrap:wrap;align-items:flex-end;gap:10px;margin-bottom:12px}
        #kxb-panel .kxb-status{padding:8px 12px;border-radius:8px;background:#f2f7f8;margin-bottom:6px}
        #kxb-panel .kxb-status.ok{background:#e6f4ea;color:#14532d}#kxb-panel .kxb-status.warn{background:#fff4cf;color:#5c3800}#kxb-panel .kxb-status.err{background:#fdecea;color:#8a1c1c}
        #kxb-panel .prog{height:4px;background:#e3e9ed;border-radius:2px;overflow:hidden;margin-bottom:14px}#kxb-panel [data-bar]{height:4px;width:0;background:#087f8c;transition:width .3s}
        #kxb-panel .chip{min-height:32px;padding:5px 12px;border-radius:16px;background:#fff}
        #kxb-panel .chip.on{background:#087f8c;color:#fff;border-color:#087f8c}
        #kxb-panel .group{display:flex;flex-wrap:wrap;align-items:center;gap:6px}
        #kxb-panel .group>b{font-size:12px;color:#4a5a66;font-weight:600;min-width:72px}
        #kxb-panel .filters{border:1px solid #e3e9ed;border-radius:10px;padding:12px;display:flex;flex-direction:column;gap:10px;margin-bottom:12px}
        #kxb-panel .kxb-kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin:10px 0}
        #kxb-panel .kxb-kpis>div{background:#f2f6f7;border-radius:10px;padding:10px 14px;display:flex;flex-direction:column}
        #kxb-panel .kxb-kpis>div.main{background:#087f8c;color:#fff}#kxb-panel .kxb-kpis b{font-size:22px}#kxb-panel .kxb-kpis small,#kxb-panel .kxb-kpis span{font-size:12px}
        #kxb-panel .kxb-kpis span.dl{font-weight:600;margin-top:2px}#kxb-panel .kxb-kpis span.dl.up{color:#15803d}#kxb-panel .kxb-kpis span.dl.down{color:#b91c1c}
        #kxb-panel .kxb-kpis>div.main span.dl.up,#kxb-panel .kxb-kpis>div.main span.dl.down{color:#fff}
        #kxb-panel .kxb-warn button{margin-left:8px;min-height:30px;padding:4px 12px}
        #kxb-panel tr.old td{background:#fdecea}#kxb-panel tr.old td:first-child{color:#b91c1c;font-weight:700}
        #kxb-panel tr.mid td{background:#fff7e0}#kxb-panel tr.mid td:first-child{color:#8a5a00;font-weight:700}
        #kxb-panel tr.low td{color:#b91c1c}
        #kxb-panel .kxb-subtabs{display:flex;gap:2px;border-bottom:1px solid #dae3e8;margin:6px 0 10px}
        #kxb-panel .kxb-subtabs button{border:0;border-bottom:3px solid transparent;border-radius:0;background:none;color:#4a5a66}
        #kxb-panel .kxb-subtabs button.on{border-bottom-color:#087f8c;color:#087f8c;font-weight:700}
        #kxb-panel .kxb-table{overflow:auto;max-height:58vh;border:1px solid #dae3e8;border-radius:8px}
        #kxb-panel table{border-collapse:collapse;width:100%;font-size:13px}
        #kxb-panel th,#kxb-panel td{padding:7px 10px;border-bottom:1px solid #edf1f3;white-space:nowrap;text-align:left}
        #kxb-panel th{background:#edf5f7;position:sticky;top:0;z-index:1;font-weight:600}#kxb-panel tbody tr:hover{background:#f7fafb}
        #kxb-panel th.n,#kxb-panel td.n{text-align:right;font-variant-numeric:tabular-nums}#kxb-panel tr.tot td{font-weight:700;background:#edf5f7}
        #kxb-panel .kxb-staff{border:1px solid #e3e9ed;border-radius:8px;padding:2px 10px;margin:6px 0}#kxb-panel .logbar{display:flex;align-items:center;gap:10px}
        #kxb-panel .kxb-pager{display:flex;align-items:center;gap:10px;margin-top:8px}
        #kxb-panel details{margin-top:6px}#kxb-panel summary{cursor:pointer;font-weight:600;padding:6px 0}
        #kxb-launch.new{background:#c2410c}
        #kxb-panel .kxb-update{position:relative;background:#fff7e6;border:1px solid #f5b041;color:#5c3b00;padding:10px 14px;border-radius:10px;margin:0 0 10px}
        #kxb-panel .kxb-update.done{background:#e8f6f3;border-color:#7fc8bd;color:#0b4f47}
        #kxb-panel .kxb-update ul{margin:6px 0 6px 18px;padding:0}#kxb-panel .kxb-update .bar{margin:6px 0 0;align-items:center}
        #kxb-panel .kxb-update .x{position:absolute;right:8px;top:6px;min-height:auto;padding:2px 8px}
        #kxb-panel a.btn{display:inline-block;text-decoration:none;padding:8px 14px;border-radius:8px;background:#087f8c;color:#fff;font-weight:600}
        #kxb-panel .kxb-warn{background:#fdecea;color:#8a1c1c;padding:8px 12px;border-radius:8px;margin:6px 0}
        #kxb-panel .kxb-muted{color:#5b6b76;font-size:12px}#kxb-panel .kxb-empty{color:#5b6b76;padding:24px;text-align:center}
        #kxb-panel .settings{border:1px solid #e3e9ed;border-radius:10px;padding:4px 12px;margin-bottom:12px}
        #kxb-panel .kxb-shop{display:flex;gap:8px;align-items:flex-end;padding:4px 0}
        #kxb-panel pre{white-space:pre-wrap;max-height:200px;overflow:auto;font-size:12px;background:#f6f8f9;padding:8px;border-radius:6px;margin:0}
        #kxb-panel .kxb-twrap{margin:4px 0 12px}
        #kxb-panel .kxb-tbar{display:flex;align-items:center;justify-content:space-between;gap:8px;margin:2px 0 4px}
        #kxb-panel button.mini{min-height:26px;padding:2px 10px;font-size:12px;border-radius:6px}
        #kxb-panel th.sort{cursor:pointer;user-select:none}#kxb-panel th.sort:hover{background:#dcecef}
        #kxb-panel th.sort[data-dir]:not([data-dir=""])::after{content:' ' attr(data-dir);color:#087f8c}
        #kxb-panel td.b{font-weight:700}
        #kxb-panel .kxb-chartbox{border:1px solid #e3e9ed;border-radius:10px;padding:8px 12px;margin:4px 0 10px}
        #kxb-panel svg.kxb-chart{width:100%;height:150px;display:block}
        #kxb-panel svg.kxb-chart .b{fill:#087f8c}#kxb-panel svg.kxb-chart .b.wk{fill:#e07b22}#kxb-panel svg.kxb-chart .b:hover{opacity:.75}
        #kxb-panel svg.kxb-chart .pv{fill:#fff;stroke:#334155;stroke-width:1.5}#kxb-panel svg.kxb-chart .avg{stroke:#94a3b8;stroke-dasharray:4 4}
        #kxb-panel svg.kxb-chart .lb{font:11px Arial,sans-serif;fill:#5b6b76}
        #kxb-panel table.heat td{text-align:center}#kxb-panel table.heat td:first-child{text-align:left;color:#4a5a66}
        #kxb-panel .kxb-month{display:inline-flex;align-items:center;gap:4px;background:#f2f6f7;border-radius:14px;padding:2px 4px 2px 10px;font-size:12px}
        #kxb-panel .kxb-month button.mini{min-height:20px;padding:0 6px;border-radius:10px}
        #kxb-panel .chk{flex-direction:row!important;align-items:center;gap:6px;font-size:13px!important}#kxb-panel .chk input{min-height:auto}
        #kxb-panel tr.click{cursor:pointer}#kxb-panel tr.click:hover td{background:#e6f4f5}
        #kxb-panel .kxb-cardback{position:absolute;inset:0;z-index:20;background:#0f172a55;display:flex;align-items:center;justify-content:center}
        #kxb-panel .kxb-card{width:min(1100px,94%);max-height:92%;display:flex;flex-direction:column;background:#fff;border-radius:12px;box-shadow:0 16px 48px #0006;overflow:hidden}
        #kxb-panel .kxb-card .top{display:flex;align-items:center;gap:10px;padding:12px 18px;border-bottom:1px solid #e3e9ed}#kxb-panel .kxb-card .top .sp{flex:1}
        #kxb-panel .kxb-card .body{overflow:auto;padding:12px 18px}
        #kxb-panel .kxb-grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
        #kxb-panel .kxb-card .kxb-table{max-height:300px}
                #kxb-panel .kxb-suggest{position:absolute;z-index:30;left:0;right:0;top:100%;margin-top:4px;max-height:360px;overflow:auto;background:#fff;border:1px solid #b8c9ce;border-radius:8px;box-shadow:0 10px 28px #0003;padding:4px 0;font-size:13px}
        #kxb-panel .kxb-suggest .hd{display:flex;justify-content:space-between;align-items:center;padding:4px 10px;color:#5b6b76;font-size:12px;border-bottom:1px solid #edf1f3}
        #kxb-panel .kxb-suggest .it{display:flex;gap:10px;align-items:center;padding:6px 10px;cursor:pointer}#kxb-panel .kxb-suggest .it:hover{background:#e6f4f5}
        #kxb-panel .kxb-suggest .nm{flex:1}#kxb-panel .kxb-suggest .cd{font-family:Consolas,monospace;color:#087f8c}#kxb-panel .kxb-suggest .qt{color:#5b6b76;min-width:70px;text-align:right}
        #kxb-panel .kxb-suggest>div:not(.hd):not(.it){padding:8px 10px}
        #kxb-panel .kxb-reqhead{margin:14px 0 2px;font-size:15px}#kxb-panel .kxb-reqsub{margin:8px 0 2px;font-weight:700;color:#4a5a66}#kxb-panel .kxb-reqhead b{color:#087f8c}
        #kxb-panel tr.fav td{background:#fff7e0}#kxb-panel tr.near td{background:#eef8f0}
        @media (max-width:1400px){#kxb-panel{font-size:13px;height:94vh;width:98vw;max-width:98vw}#kxb-panel .body{padding:12px 14px}#kxb-panel .top{padding:10px 14px}
          #kxb-panel .tabs button{padding:8px 12px;font-size:14px}#kxb-panel .kxb-kpis{grid-template-columns:repeat(auto-fit,minmax(150px,1fr))}#kxb-panel .kxb-kpis b{font-size:19px}
          #kxb-panel th,#kxb-panel td{padding:6px 8px;font-size:12.5px}#kxb-panel .kxb-drop>summary{min-width:240px}}
        `, document.head || document.documentElement); style.dataset.kxbUi = '';
        const launch = el('button', 'AutoBI · Kho & Xuất Bán', document.body); launch.id = 'kxb-launch'; launch.dataset.kxbUi = ''; launch.type = 'button';
        const back = el('div', undefined, document.body); back.id = 'kxb-back'; back.dataset.kxbUi = ''; back.hidden = true;
        ui = el('div', undefined, back); ui.id = 'kxb-panel'; ui.setAttribute('role', 'dialog'); ui.setAttribute('aria-label', 'AutoBI Kho và Xuất bán');
        ui.innerHTML = `
        <div class="top"><strong>AutoBI · Kho &amp; Xuất Bán</strong><span class="kxb-muted">V${VERSION}</span>
          <div class="tabs" role="tablist"><button type="button" role="tab" data-tab="sales">🛒 Xuất bán</button><button type="button" role="tab" data-tab="inventory">📦 Tồn kho</button><button type="button" role="tab" data-tab="balance">⚖️ Cân hàng</button><button type="button" role="tab" data-tab="request">🔁 Check xin hàng</button></div>
          <span class="sp"></span><span data-timer class="kxb-muted">00:00:00</span>
          <button type="button" class="danger busy-only" data-stop>⛔ Dừng</button>
          <button type="button" class="idle-only" data-excel>⬇ Tải Excel</button>
          <button type="button" data-close aria-label="Đóng">✕</button></div>
        <div class="body">
          <div data-update></div>
          <details class="settings" data-settings><summary>⚙️ Cài đặt siêu thị</summary>
            <div data-shops></div><button type="button" data-add>+ Thêm siêu thị</button>
            <div class="bar" style="margin-top:10px"><label>Tính doanh số theo<select data-cfg="basis"><option value="created">Ngày tạo đơn</option><option value="shipped">Ngày xuất hàng</option></select></label>
              <label>Kiểm tra lại nhập trả (ngày gần nhất)<input type="number" min="2" max="35" data-cfg="returnDays" style="width:120px"></label>
              <label>Cân hàng: giữ đủ bán (ngày)<input type="number" min="3" max="60" data-cfg="balTarget" style="width:120px"></label>
              <label>Không lấy lại ngày vừa lấy dưới (giờ)<input type="number" min="0" max="24" data-cfg="freshHours" style="width:120px"></label>
              <label>Cân hàng: máy dự phòng<input type="number" min="0" max="5" data-cfg="balSpare" style="width:120px"></label>
              <label>Cảnh báo tồn lâu từ (ngày)<input type="number" min="15" max="365" data-cfg="ageAlert" style="width:120px"></label>
              <label title="Khi chia số cần xin cho các siêu thị khác: mỗi nơi cho giữ lại bấy nhiêu máy; chỉ khi thiếu mới lấy cả máy giữ lại (ghi “máy cuối”)">Xin hàng: nơi cho giữ lại (máy)<input type="number" min="0" max="5" data-cfg="srcKeep" style="width:120px"></label>
              <span class="group" style="gap:10px"><b style="min-width:0">Ảnh xuất bán gồm</b><label class="chk"><input type="checkbox" data-img="cats">Theo ngành</label><label class="chk"><input type="checkbox" data-img="brands">Theo hãng</label><label class="chk"><input type="checkbox" data-img="staff">Nhân viên</label>
                <label class="chk">Top<input type="number" min="5" max="30" data-img="top" style="width:64px"></label></span>
              <label class="chk"><input type="checkbox" data-cfg="balIncoming">Cân hàng: trừ hàng đang về vào số xin</label>
              <label class="chk"><input type="checkbox" data-cfg="notify">Kêu "ting" + thông báo khi đổ xong (phiên trên 20 giây)</label></div>
            <div class="group" style="margin:6px 0 10px"><b>Tra tồn CRM</b><span>Tỉnh</span><span data-crm-prov></span><button type="button" data-crm-reload title="Lấy lại danh sách siêu thị của các tỉnh đã chọn từ CRM (cần đăng nhập CRM)">🔄 Lấy lại danh sách siêu thị</button><span class="kxb-muted" data-crm-info></span></div>
            <div class="group" style="margin:-4px 0 10px"><b></b><span>Tỉnh lân cận</span><span data-crm-near></span><span class="kxb-muted">Check xin hàng: khi cả tỉnh hết, bấm "🔎 Tỉnh lân cận" để tìm thêm ở các tỉnh này. Bỏ chọn hết = dùng tỉnh giáp ranh mặc định.</span></div>
            <div class="kxb-muted" data-storage style="margin-bottom:8px"></div>
            <div class="bar" style="align-items:center"><button type="button" class="primary" data-save>Lưu cài đặt</button> <button type="button" data-check-update>🔄 Kiểm tra bản mới</button>
              <span style="flex:1"></span><button type="button" data-backup title="Tải 1 file chứa toàn bộ số đã lưu (xuất bán, tồn kho, cài đặt) để mang sang máy khác">💾 Sao lưu dữ liệu</button>
              <button type="button" data-restore title="Nhận file sao lưu: ngày nào bản nào mới hơn thì dùng bản đó, không mất số đang có">📂 Khôi phục từ file</button><input type="file" accept=".json,application/json" data-restore-file hidden></div></details>
          <div data-status class="kxb-status">Sẵn sàng.</div><div class="prog"><div data-bar></div></div>
          <div data-pane="sales">
            <div class="filters">
              <div class="group"><b>Kỳ</b><span class="group" data-presets></span><span style="flex:1"></span><b style="min-width:0">So với</b><span class="group" data-compare></span></div>
              <div class="bar" style="margin:0"><label>Từ ngày<input type="date" data-from></label><label>Đến ngày<input type="date" data-to></label>
                <label style="flex-direction:row;align-items:center;gap:6px;font-size:13px"><input type="checkbox" data-refetch style="min-height:auto">Lấy lại cả ngày đã chốt</label>
                <span style="flex:1"></span><button type="button" data-img title="Tạo ảnh báo cáo theo bộ lọc đang chọn: chép vào bộ nhớ để dán Zalo, đồng thời tải file PNG">🖼 Ảnh báo cáo</button><button type="button" class="idle-only" data-view>Xem số đã lưu</button><button type="button" class="primary idle-only" data-run-sales>Đổ xuất bán</button></div>
              <div class="group"><b>Siêu thị</b><span class="group" data-sales-shops></span></div>
              <div class="group"><b>Ngành hàng</b><select data-sales-cat style="min-width:320px"></select><button type="button" class="chip" data-phone-only>📱 Chỉ điện thoại</button><b style="margin-left:12px">Hãng</b><span data-sales-brands></span></div>
              <div class="group"><b>Loại hàng</b><span class="group" data-sales-conds></span></div></div>
            <div data-sales-result></div></div>
          <div data-pane="inventory" hidden>
            <div class="filters" data-inv-filters>
              <div class="group"><b>Siêu thị</b><span class="group" data-inv-shops></span><span class="sp" style="flex:1"></span><button type="button" class="idle-only" data-print-inv title="In danh sách đang lọc ra giấy A4 dọc, có ô KIỂM để tích">🖨 In phiếu kiểm</button><label class="chk idle-only" title="Gọi BI thêm 1 lần mỗi siêu thị để lấy hàng đang chuyển về (đã xuất từ nơi khác, chưa nhận) — xem ở tab con Đang về; Cân hàng trừ vào số xin"><input type="checkbox" data-inv-transit>Kèm hàng đang về</label><button type="button" class="primary idle-only" data-run-inv>Đổ tồn kho</button></div>
              <div class="group"><b>Lọc</b><select data-f="category"></select><select data-f="group"></select><select data-f="brand"></select>
                <input type="search" data-f="q" placeholder="Tìm IMEI, mã hoặc tên sản phẩm" style="min-width:280px"><button type="button" data-clear>Bỏ lọc</button></div>
              <div class="group"><b>Trạng thái</b><span class="group" data-inv-conds></span></div></div>
            <div data-inv-result></div></div>
          <div data-pane="balance" hidden>
            <div class="filters">
              <div class="group"><b>Tốc độ bán</b><span class="group" data-bal-days></span><span class="sp" style="flex:1"></span><button type="button" class="primary idle-only" data-run-bal title="Đổ tồn kho mới + xuất bán các ngày còn thiếu">Đổ cân hàng</button></div>
              <div class="group"><b>Siêu thị</b><span class="group" data-bal-shops></span></div>
              <div class="group"><b>Ngành hàng</b><select data-bal-cat style="min-width:220px"></select><b style="margin-left:12px;min-width:0">Hãng</b><span data-bal-brands></span><button type="button" class="chip" data-bal-model title="Gộp các mã khác màu của cùng một model thành 1 dòng">🧩 Gộp theo model</button></div></div>
            <div data-bal-result></div></div>
          <div data-pane="request" hidden>
            <div class="filters">
              <div class="group"><b>Xin cho</b><select data-req-shop style="min-width:220px"></select><span class="kxb-muted">siêu thị đang cần hàng — quyết định "gần" là cùng huyện / cùng tỉnh cũ với siêu thị này</span></div>
              <div class="group" style="align-items:flex-start"><b>Sản phẩm</b><span style="flex:1;min-width:360px;position:relative;display:flex"><textarea data-req-q rows="2" placeholder="Gõ tên (vd: iPhone 18) → chọn trong danh sách gợi ý, hoặc dán mã 13 số · nhiều mã: mỗi mã một dòng (Shift+Enter) · Enter = Check" style="width:100%;font:inherit;border:1px solid #b8c9ce;border-radius:8px;padding:6px 8px"></textarea></span>
                <label class="chk" title="Số máy cần xin cho mỗi sản phẩm — tool chia cho từng nơi cho (gần trước, nơi tồn nhiều cho nhiều)">SL cần<input type="number" min="1" max="50" step="1" data-req-qty style="width:64px"></label><button type="button" class="primary idle-only" data-run-req>🔎 Check</button><label class="chk"><input type="checkbox" data-req-far>Hiện cả nơi xa</label></div>
              <div class="group"><b>Đang hết</b><span class="group" data-req-quick></span></div></div>
            <div data-req-result></div></div>
          <details data-logbox><summary class="logbar">📋 Nhật ký <button type="button" data-copy-log>Sao chép</button><button type="button" data-clear-log>Xóa nhật ký</button></summary><pre data-log></pre></details>
        </div>`;
        const today = isoDate(new Date());
        ui.querySelector('[data-from]').value = today.slice(0, 8) + '01'; ui.querySelector('[data-to]').value = today;
        ui.querySelector('[data-cfg="basis"]').value = config.basis || 'created'; ui.querySelector('[data-cfg="returnDays"]').value = config.returnDays || 7; ui.querySelector('[data-cfg="balTarget"]').value = config.balTarget || 14;
        ui.querySelector('[data-cfg="freshHours"]').value = config.freshHours ?? 2;
        ui.querySelector('[data-cfg="balSpare"]').value = config.balSpare ?? 1; ui.querySelector('[data-cfg="ageAlert"]').value = config.ageAlert || 60; ui.querySelector('[data-cfg="srcKeep"]').value = config.srcKeep ?? 1;
        ui.querySelector('[data-cfg="notify"]').checked = config.notify !== false;
        ui.querySelector('[data-cfg="balIncoming"]').checked = config.balIncoming !== false;
        renderCrmSetting();
        ui.querySelector('[data-crm-reload]').onclick = safely(() => withSession('crmStores', async () => { await ensureCrmStores(true); return `Đã lấy danh sách siêu thị tra tồn: ${crmStores().length} siêu thị`; }));
        { const o = imgOpt(); ['cats', 'brands', 'staff'].forEach(k => { ui.querySelector(`[data-img="${k}"]`).checked = !!o[k]; }); ui.querySelector('[data-img="top"]').value = o.top; }
        ui.querySelector('[data-settings]').addEventListener('toggle', e => { if (e.target.open) renderStorage(); });
        (config.shops.length ? config.shops : [{ code: '', name: '' }]).forEach(shopEditor);
        if (!config.shops.length) ui.querySelector('[data-settings]').open = true;
        const savedShops = load('invShops', null);
        const codes = config.shops.map(s => keyCode(s.code));
        view.salesShops = new Set(codes); view.balShops = new Set(codes);
        view.salesCat = String(load('salesCat', BAL_CATEGORY) ?? '');         // lần đầu: chọn sẵn ngành Điện thoại
        view.balDays = [7, 10, 14, 30].includes(load('balDays', 10)) ? load('balDays', 10) : 10;
        view.invShops = new Set(Array.isArray(savedShops) && savedShops.some(c => codes.includes(c)) ? savedShops.filter(c => codes.includes(c)) : codes);
        // V2.0: nhớ tab con, chế độ gộp model, tab chính lần trước
        const okTab = (k, list, d) => list.includes(k) ? k : d;
        view.salesTab = okTab(load('salesTab', 'category'), ['category', 'brand', 'staff', 'attach', 'staffProduct', 'product', 'daily', 'hour', 'channel', 'pending', 'returned', 'detail'], 'category');
        view.invTab = okTab(load('invTab', 'group'), ['group', 'product', 'imei', 'age', 'transit', 'crm'], 'group');
        view.tab = okTab(load('mainTab', 'sales'), ['sales', 'inventory', 'balance', 'request'], 'sales');
        view.balModel = !!load('balModel', false);
        view.balCat = String(load('balCat', BAL_CATEGORY) || BAL_CATEGORY);
        const trBox = ui.querySelector('[data-inv-transit]'); trBox.checked = !!config.invTransit;
        trBox.onchange = () => { config.invTransit = trBox.checked; save('config', config); };
        // So với: cùng kỳ tháng trước / 7 ngày trước (chỉ đọc số đã lưu)
        const cmpBox = ui.querySelector('[data-compare]');
        const drawCompare = () => { cmpBox.replaceChildren(); [['month', 'Tháng trước'], ['week', '7 ngày trước']].forEach(([k, l]) => { const b = el('button', l, cmpBox, 'chip'); b.type = 'button'; b.classList.toggle('on', (config.compareMode === 'week' ? 'week' : 'month') === k);
            b.onclick = safely(() => { config.compareMode = k; save('config', config); drawCompare(); if (config.shops.length) { view.sales = salesView(selectedRange()); renderSales(); } }); }); };
        drawCompare();
        // Nút chọn nhanh kỳ
        const presets = ui.querySelector('[data-presets]');
        const setRange = (f, t) => { ui.querySelector('[data-from]').value = f; ui.querySelector('[data-to]').value = t; };
        const firstOf = (d, k = 0) => isoDate(new Date(d.getFullYear(), d.getMonth() + k, 1));
        const lastOf = (d, k = 0) => isoDate(new Date(d.getFullYear(), d.getMonth() + k + 1, 0));
        const now = new Date(), td = isoDate(now);
        [['Hôm nay', td, td], ['Hôm qua', addDays(td, -1), addDays(td, -1)], ['7 ngày', addDays(td, -6), td], ['Tháng này', firstOf(now), td], ['Tháng trước', firstOf(now, -1), lastOf(now, -1)]].forEach(([l, f, t]) => {
            const b = el('button', l, presets, 'chip'); b.type = 'button';
            b.onclick = safely(() => { setRange(f, t); presets.querySelectorAll('.chip').forEach(x => x.classList.toggle('on', x === b)); validateShops(config.shops); view.sales = salesView(selectedRange()); renderSales(); });
        });
        const last = load('lastInventory', null);
        if (last && Array.isArray(last.records) && last.noTransit) view.inv = last;   // số tồn cũ có hàng đang chuyển kho → bỏ, đổ lại
        const open = on => { back.hidden = !on; launch.hidden = on; if (on) { storeReady.then(() => { renderAll(); renderUpdate(); }); checkUpdate(false); } };
        launch.onclick = () => open(true);
        ui.querySelector('[data-close]').onclick = () => open(false);
        back.addEventListener('mousedown', e => { if (e.target === back && !running) open(false); });
        document.addEventListener('keydown', e => { if (e.key !== 'Escape' || back.hidden) return; const c = document.getElementById('kxb-card'); if (c) { c.remove(); return; } if (!running) open(false); });
        ui.querySelectorAll('[data-tab]').forEach(b => b.onclick = () => { view.tab = b.dataset.tab; try { save('mainTab', view.tab); } catch { /* bỏ qua */ } renderAll(); });
        // Đổi ngày bằng tay: tự hiện số đã lưu của kỳ mới (không gọi BI)
        let td2; const onDate = () => { clearTimeout(td2); td2 = setTimeout(() => { try { if (!config.shops.length) return; presets.querySelectorAll('.chip').forEach(x => x.classList.remove('on')); view.sales = salesView(selectedRange()); renderSales(); } catch (e) { status(e.message, 'err'); } }, 300); };
        ui.querySelector('[data-from]').addEventListener('change', onDate); ui.querySelector('[data-to]').addEventListener('change', onDate);
        ui.querySelector('[data-backup]').onclick = safely(backupData);
        const rf = ui.querySelector('[data-restore-file]');
        ui.querySelector('[data-restore]').onclick = () => rf.click();
        rf.onchange = safely(async () => { const f = rf.files[0]; rf.value = ''; await restoreData(f); });
        // Phím tắt: Alt+1/2/3 chuyển tab, Esc đóng
        document.addEventListener('keydown', e => { if (back.hidden || !e.altKey || !['1', '2', '3', '4'].includes(e.key)) return; e.preventDefault(); view.tab = ['sales', 'inventory', 'balance', 'request'][Number(e.key) - 1]; try { save('mainTab', view.tab); } catch { /* bỏ qua */ } renderAll(); });
        ui.querySelector('[data-stop]').onclick = stop;
        ui.querySelector('[data-add]').onclick = safely(() => { invariant(!running, 'Không sửa cài đặt khi đang chạy'); shopEditor(); });
        ui.querySelector('[data-save]').onclick = safely(saveConfig);
        ui.querySelector('[data-excel]').onclick = safely(exportExcel);
        ui.querySelector('[data-copy-log]').onclick = e => { e.preventDefault(); e.stopPropagation(); copyLog(); };
        ui.querySelector('[data-clear-log]').onclick = e => { e.preventDefault(); e.stopPropagation(); journal.length = 0; try { save('journal', journal); } catch { /* bỏ qua */ } log('Đã xóa nhật ký cũ'); };
        storeReady = initStore().then(() => { if (config.shops.length) { view.sales = salesView(selectedRange()); if (!back.hidden) renderAll(); } });
        ui.querySelector('[data-run-sales]').onclick = safely(() => { validateShops(config.shops); selectedRange(); return withSession('sales', s => runSales(s, ui.querySelector('[data-refetch]').checked)); });
        ui.querySelector('[data-img]').onclick = safely(salesImage);
        ui.querySelector('[data-view]').onclick = safely(() => { validateShops(config.shops); view.sales = salesView(selectedRange()); renderSales(); status(view.sales.missing ? `Còn ${view.sales.missing} ngày chưa lấy trong kỳ` : 'Số đã lưu trên máy này', view.sales.missing ? 'warn' : 'ok'); });
        ui.querySelector('[data-print-inv]').onclick = safely(printInventory);
        ui.addEventListener('mousedown', e => { if (!e.target.closest('.kxb-drop')) { view.openDrop = ''; ui.querySelectorAll('.kxb-drop[open]').forEach(d => d.open = false); } });
        ui.querySelector('[data-run-req]').onclick = safely(() => { reqState().q = ui.querySelector('[data-req-q]').value; return withSession('crmReq', runRequest); });
        ui.querySelector('[data-req-q]').addEventListener('keydown', e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); ui.querySelector('[data-run-req]').click(); } });
        ui.querySelector('[data-req-q]').addEventListener('input', e => { reqState().q = e.target.value; });
        { const f = ui.querySelector('[data-req-q]'); const sg = attachSuggest(f, (code, p, all) => { replaceCurLine(f, all ? all.map(x => x.code).join('\n') : code); });
          // Enter: nếu đang mở gợi ý thì không chạy Check ngay
          f.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.shiftKey) sg.hide(); }, true); }
        ui.querySelector('[data-run-bal]').onclick = safely(() => { validateShops(config.shops); return withSession('balance', runBalance); });
        ui.querySelector('[data-run-inv]').onclick = safely(() => { validateShops(config.shops); return withSession('inventory', runInventory); });
        for (const k of ['category', 'group', 'brand']) ui.querySelector(`[data-f="${k}"]`).onchange = e => { view.invFilter[k] = e.target.value; if (k === 'category') view.invFilter.group = ''; renderInventoryFilters(); renderInventory(); };
        let tq; ui.querySelector('[data-f="q"]').oninput = e => { clearTimeout(tq); tq = setTimeout(() => { view.invFilter.q = e.target.value; renderInventory(); }, 250); };
        ui.querySelector('[data-clear]').onclick = () => { view.invFilter = { category: '', group: '', brand: '', conditions: [], q: '' }; ui.querySelector('[data-f="q"]').value = ''; renderInventoryFilters(); renderInventory(); };
        ui.querySelector('[data-check-update]').onclick = safely(() => checkUpdate(true));
        renderUpdate();
        setTimeout(() => checkUpdate(false), 4000);
        setInterval(() => checkUpdate(false), UPDATE_EVERY);
        log(`Mở AutoBI Kho & Xuất Bán V${VERSION}.`);
        // V2.5.1: bản này chiếm nút, bản khác (từ 2.5.1) cũng đang cài → báo trùng
        setTimeout(() => { try { const v = String(document.documentElement.dataset.kxbVersions || '').split(','); if (v.filter(Boolean).length > 1) dupWarning(v, VERSION); } catch { /* bỏ qua */ } }, 2000);
    }
    window.addEventListener('pagehide', () => { if (running) stop(); });
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mount, { once: true }); else mount();
})();
