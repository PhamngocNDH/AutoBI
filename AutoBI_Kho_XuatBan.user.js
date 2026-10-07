// ==UserScript==
// @name         AutoBI - Kho & Xuất Bán
// @namespace    https://github.com/PhamngocNDH/AutoBI/kho-xuatban-test
// @version      1.7.0
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
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      docs.google.com
// @connect      googleusercontent.com
// @connect      *.googleusercontent.com
// @connect      raw.githubusercontent.com
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
    const VERSION = '1.7.0';
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
    const CONDITION_KEYS = ['TRANGTHAIHOSO', 'TRANGTHAISANPHAM', 'TINHTRANGSANPHAM', 'INVENTORYSTATUSNAME', 'TRANGTHAIHANG', 'LOAIHANG', 'PRODUCTSTATUSNAME'];
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
            outShop: (k => k && clean(r[k]) ? keyCode(String(r[k]).replace(/\D/g, '')) : '')(outStoreKey(r))   // BI trả dạng "2,928"
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
        const m = new Map(), key = (shop, p) => keyCode(shop) + '|' + clean(p);
        const row = (shop, p, name, brand) => {
            const k = key(shop, p);
            if (!m.has(k)) m.set(k, { shop: keyCode(shop), product: clean(p), name: clean(name), brand: clean(brand), stock: 0, sold: 0 });
            const x = m.get(k); if (!x.name && name) x.name = clean(name); if (!x.brand && brand) x.brand = clean(brand); return x;
        };
        for (const r of stock) row(r.shop, r.product, r.productName, r.brand).stock += Number(r.qty) || 0;
        for (const l of sales) if (!l.outShop || keyCode(l.outShop) === keyCode(l.shop)) row(l.shop, l.product, l.productName, l.brand).sold += Number(l.qty) || 0;
        return [...m.values()].map(x => {
            const perDay = x.sold / days, cover = perDay > 0 ? x.stock / perDay : (x.stock > 0 ? Infinity : 0);
            const want = perDay > 0 ? Math.ceil(perDay * target - 1e-9) + spare : 0;
            const need = Math.max(0, want - x.stock);
            const status = x.stock <= 0 && x.sold > 0 ? 'Hết hàng' : x.sold <= 0 ? (x.stock > 0 ? 'Không bán' : '') : need > 0 ? 'Sắp hết' : cover > target * 2 ? 'Tồn nhiều' : 'Đủ';
            return { ...x, perDay: Math.round(perDay * 100) / 100, cover: Number.isFinite(cover) ? Math.round(cover * 10) / 10 : null, need, status };
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

    // Hàm thuần cho kiểm thử offline; không cài global trên website thật.
    if (typeof module === 'object' && module.exports) {
        module.exports = { clean, norm, hasCode, day, toBI, addDays, validateRange, daysIn, apiNumber, parseDelimited,
            lineFromApi, reasons, pending, validateSalesLines, splitSales, summarizeSales, inPeriod, dayStatus, daysToFetch,
            parseRateLimit, waitBeforeCall, inventoryRecordFromApi, summarizeInventory, filterInventory, inventoryOptions, inventoryViews, validateShops, authorizeSheetRows, escHtml, inventoryChecklist, inventoryPrintHtml, conditionKey, conditionText, newerVersion, parseRemoteScript, outStoreKey, balanceRows, isFresh, shiftMonth, prevMonthRange, attachStats, pendingOrders, projectMonth, daysBetween };
        return;
    }


    /* ================= TRÌNH DUYỆT ================= */
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
    function loadMonth(month) { const x = load('month_' + month, null); return x && typeof x === 'object' ? x : { book: {}, lines: [] }; }
    function saveMonth(month, data) { save('month_' + month, data); }
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
    async function inventoryShop(shop, session) {
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
        return records;
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
    function progress(done, total) { if (ui) ui.querySelector('[data-bar]').style.width = (total ? Math.round(done / total * 100) : 0) + '%'; }
    function stop() {
        if (!running) return;
        running.status = 'cancelled'; running.controller.abort();
        log('Đã dừng. Ngày đã lấy xong vẫn được giữ; phần đang lấy dở bị bỏ.'); status('Đã dừng', 'warn');
    }
    async function withSession(mode, job) {
        invariant(!running, 'Có phiên đang chạy');
        const session = { id: uuid(), mode, started: Date.now(), status: 'running', controller: new AbortController() };
        running = session; ui.classList.add('busy');
        const timer = setInterval(() => { ui.querySelector('[data-timer]').textContent = clockText(session.started); }, 250);
        log(`Bắt đầu ${({ sales: 'đổ xuất bán', balance: 'đổ cân hàng', auto: 'tự đổ xuất bán hôm qua', prev: 'đổ xuất bán cùng kỳ tháng trước' })[mode] || 'đổ tồn kho'}`); status('Đang kiểm tra quyền…'); progress(0, 1);
        try {
            await authCheck(session); check(session); log(`Quyền hợp lệ: ${auth.user} — ${auth.name}`);
            const msg = await job(session); check(session);
            session.status = 'completed'; progress(1, 1); status(msg || 'Hoàn tất', /lỗi|chưa/.test(msg || '') ? 'warn' : 'ok'); log(msg || 'Hoàn tất');
        } catch (e) {
            if (session.status !== 'cancelled') { session.status = 'error'; status(e.message, 'err'); log(e.message, 'error'); }
        } finally {
            clearInterval(timer); ui.querySelector('[data-timer]').textContent = clockText(session.started);
            running = null; ui.classList.remove('busy');
        }
        return session.status;
    }
    // Lần mở BI đầu tiên trong ngày: tự lấy xuất bán hôm qua (chạy nền, 1 lượt gọi BI)
    async function autoYesterday() {
        if (config.autoDaily === false || !config.shops.length || running) return;
        const today = isoDate(new Date()), y = addDays(today, -1);
        if (load('autoDay', '') === today) return;
        try { detectUser(); } catch { return; }                // chưa đăng nhập BI → để lần sau
        const range = { from: y, to: y };
        if (!daysToFetch(loadBook(range).book, range, shopsKey(), today, false, config.returnDays, SALES_SCHEMA, Date.now(), config.freshHours ?? 2).includes(y)) { save('autoDay', today); return; }
        const launch = document.getElementById('kxb-launch'), label = launch?.textContent;
        if (launch) launch.textContent = '⏳ AutoBI đang lấy xuất bán hôm qua…';
        const t0 = Date.now(), st = await withSession('auto', s => runSales(s, false, range));
        if (launch) launch.textContent = label;
        if (st === 'completed' && (loadBook(range).book[y]?.atMs || 0) >= t0) save('autoDay', today);
        renderUpdate();
    }
    function selectedRange() { return validateRange(ui.querySelector('[data-from]').value, ui.querySelector('[data-to]').value); }

    // fetchRange: kỳ cần lấy (mặc định kỳ đang chọn); màn hình vẫn hiện kỳ đang chọn
    async function runSales(session, refetchAll, fetchRange) {
        const range = fetchRange || selectedRange(), today = isoDate(new Date());
        const shown = () => { try { return selectedRange(); } catch { return range; } };
        const fh = config.freshHours ?? 2;
        const need = daysToFetch(loadBook(range).book, range, shopsKey(), today, refetchAll, config.returnDays, SALES_SCHEMA, Date.now(), fh);
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
        const records = [];
        for (let i = 0; i < shops.length; i++) {
            check(session); invariant(detectUser() === auth.user, 'Tài khoản đã thay đổi');
            status(`Tồn kho ${i + 1}/${shops.length}: ${shops[i].code} — ${shops[i].name}`); progress(i, shops.length);
            records.push(...await inventoryShop(shops[i], session));
        }
        // chỉ công bố khi TẤT CẢ shop đã chọn xong
        // Gộp: chỉ thay số của siêu thị vừa đổ, giữ nguyên số các siêu thị khác (kèm giờ đổ riêng từng siêu thị)
        const now = new Date().toISOString(), fetched = new Set(shops.map(s => keyCode(s.code)));
        const prev = view.inv || { records: [], shops: [], shopTimes: {} };
        const shopTimes = { ...(prev.shopTimes || Object.fromEntries((prev.shops || []).map(c => [c, prev.capturedAt]))) };
        fetched.forEach(c => shopTimes[c] = now);
        view.inv = { noTransit: true, capturedAt: now, shops: [...new Set([...(prev.shops || []), ...fetched])], shopTimes,
            records: prev.records.filter(r => !fetched.has(r.shop)).concat(records) };
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
        const prevRange = prevMonthRange(range), pb = loadBook(prevRange);
        const prevMissing = daysIn(prevRange).filter(d => d <= today && (!pb.book[d] || pb.book[d].schema !== SALES_SCHEMA)).length;
        return { range, basis: config.basis, days, missing: days.filter(d => d.status === 'Chưa lấy').length, oldSchema, allLines: period, pendingLines: pend, returnedLines: ret,
            prevRange, prevMissing, prevAllLines: inPeriod(pb.lines, prevRange, config.basis) };
    }

    /* ---------- Định dạng & bảng ---------- */
    const fmt = v => new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 2 }).format(v);
    const pct = v => fmt(Math.round(v * 10) / 10) + '%';
    const mil = v => fmt(Math.round(v / 1e4) / 100) + ' tr';
    const shopName = code => config.shops.find(s => keyCode(s.code) === keyCode(code))?.name || code;
    function el(tag, text, parent, cls) { const e = document.createElement(tag); if (text !== undefined && text !== null) e.textContent = text; if (cls) e.className = cls; if (parent) parent.append(e); return e; }
    function table(parent, headers, rows, { pageSize = 200, num = [], total = false, rowClass = null } = {}) {
        const box = el('div', undefined, parent, 'kxb-table');
        const t = el('table', undefined, box), tr = el('tr', undefined, el('thead', undefined, t));
        headers.forEach((h, i) => el('th', h, tr, num.includes(i) ? 'n' : ''));
        const body = el('tbody', undefined, t); let page = 0;
        const pager = rows.length > pageSize ? el('div', undefined, parent, 'kxb-pager') : null;
        const draw = () => {
            body.replaceChildren();
            rows.slice(page * pageSize, (page + 1) * pageSize).forEach((row, ri, arr) => { const line = el('tr', undefined, body, total && ri === arr.length - 1 && (page + 1) * pageSize >= rows.length ? 'tot' : (rowClass ? rowClass(page * pageSize + ri) || '' : '')); row.forEach((c, i) => el('td', String(c ?? ''), line, num.includes(i) ? 'n' : '')); });
            if (pager) {
                pager.replaceChildren();
                const prev = el('button', '← Trước', pager); prev.disabled = page === 0; prev.onclick = () => { page--; draw(); };
                el('span', `Trang ${page + 1}/${Math.ceil(rows.length / pageSize)} · ${fmt(rows.length)} dòng`, pager);
                const next = el('button', 'Sau →', pager); next.disabled = (page + 1) * pageSize >= rows.length; next.onclick = () => { page++; draw(); };
            }
        };
        draw();
        if (!rows.length) el('div', 'Không có dòng nào', box, 'kxb-empty');
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
            const w = el('div', `Cùng kỳ tháng trước (${toBI(r.prevRange.from)}–${toBI(r.prevRange.to)}) còn ${r.prevMissing} ngày chưa lấy — chưa so được ▲▼. `, area, 'kxb-warn');
            const b = el('button', 'Đổ cùng kỳ tháng trước', w, 'idle-only'); b.type = 'button';
            b.onclick = safely(() => { validateShops(config.shops); return withSession('prev', ss => runSales(ss, false, r.prevRange)); });
        }
        // Bán kèm & dự kiến: tính trên các siêu thị đang chọn (không theo bộ lọc ngành/hãng/loại hàng)
        const shopAll = r.allLines.filter(l => view.salesShops.has(l.shop)), att = attachData(shopAll);
        const attPrev = r.prevMissing ? null : attachData((r.prevAllLines || []).filter(l => view.salesShops.has(l.shop)));
        const proj = projectMonth(r.lines.filter(l => (r.basis === 'shipped' && l.shipped ? l.shipped : l.created) < today).reduce((a, l) => a + l.qty * l.price, 0), r.range, today);
        const tCat = config.targetCat ?? BAL_CATEGORY, tShops = config.shops.filter(x => view.salesShops.has(keyCode(x.code)));
        const target = tShops.reduce((a, x) => a + (Number(x.target) || 0), 0) * 1e6, tMatch = view.salesCat === tCat && !view.salesBrands.size && !view.salesConds.size;
        const projItem = proj ? ['Dự kiến cuối tháng', mil(proj.value), !target ? 'chưa đặt mục tiêu (Cài đặt)' : tMatch ? `mục tiêu ${mil(target)} · ${pct(proj.value / target * 100)}` : `chọn ngành ${tCat || 'tất cả'} để so mục tiêu`,
            target && tMatch ? { text: proj.value >= target ? 'đạt mục tiêu' : `thiếu ${mil(target - proj.value)}`, cls: proj.value >= target ? 'up' : 'down' } : null] : null;
        kpis(area, [['Đã bán', `SL ${fmt(s.quantity)}`, mil(s.revenue), delta(s.revenue, ps.revenue, r)],
            ...codes.map(c => { const x = s.shops.find(z => z.code === c) || { quantity: 0, revenue: 0 }, y = ps.shops.find(z => z.code === c) || { revenue: 0 }; return [shopName(c), `SL ${fmt(x.quantity)}`, mil(x.revenue), delta(x.revenue, y.revenue, r)]; }),
            ['Tỷ lệ bán kèm ĐT', pct(att.total.rate), `${fmt(att.total.attached)}/${fmt(att.total.phoneOrders)} đơn điện thoại`,
                attPrev && attPrev.total.phoneOrders ? { text: `${att.total.rate >= attPrev.total.rate ? '▲' : '▼'} ${fmt(Math.abs(Math.round(att.total.rate - attPrev.total.rate)))} điểm so cùng kỳ`, cls: att.total.rate >= attPrev.total.rate ? 'up' : 'down' } : null],
            ...(projItem ? [projItem] : []),
            ['Đơn treo', fmt(pendOrders), 'chưa xuất / chưa giao', oldPend ? { text: `${fmt(oldPend)} đơn treo từ 2 ngày`, cls: 'down' } : null],
            ['Khách nhập trả', fmt(retOrders), `đã bỏ SL ${fmt(r.returnedLines.reduce((a, l) => a + l.qty, 0))} · ${mil(r.returnedLines.reduce((a, l) => a + l.qty * l.price, 0))}`]]);
        el('div', `Kỳ ${toBI(r.range.from)}–${toBI(r.range.to)} · Kho tạo · ngành: ${view.salesCat || 'tất cả'} · tính theo ${r.basis === 'shipped' ? 'ngày xuất' : 'ngày tạo'} · dòng Đã xuất – Đã giao – Chưa hủy · loại hàng: ${view.salesConds.size ? [...view.salesConds].join(', ') : 'tất cả'} · bỏ cả đơn khách nhập trả (kiểm tra lại ${config.returnDays} ngày gần nhất) · doanh thu = Giá bán × SL (gồm VAT)`, area, 'kxb-muted');
        subtabs(area, [['category', 'Theo ngành hàng'], ['brand', 'Theo hãng'], ['staff', 'Nhân viên × hãng'], ['attach', 'Bán kèm'], ['staffProduct', 'Nhân viên × sản phẩm'], ['product', 'Sản phẩm'], ['daily', 'Theo ngày'], ['pending', `Đơn treo (${pendOrders})`], ['returned', `Nhập trả (${retOrders})`], ['detail', 'Chi tiết']], view.salesTab, k => { view.salesTab = k; renderSales(); });
        const pane = el('div', undefined, area);
        const n = (from, count) => Array.from({ length: count }, (_, i) => from + i);
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
                 ['Tổng', '', ...s.brands.map(b => fmt(b.quantity)), fmt(s.quantity), '', mil(s.revenue)]], { num: n(2, m.brands.length + 3), total: true });
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
                { num: [2, 3, 4, 5, 6, 7, 8], total: true, rowClass: i => i < att.staff.length && att.staff[i].phoneOrders >= 3 && att.staff[i].rate < att.total.rate / 2 ? 'low' : '' });
            if (!attPrev) el('div', 'Cột "so cùng kỳ" trống vì cùng kỳ tháng trước chưa đủ ngày — bấm "Đổ cùng kỳ tháng trước" ở trên.', pane, 'kxb-muted');
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
    function renderInventory() {
        const area = ui.querySelector('[data-inv-result]'); area.replaceChildren();
        if (!view.inv) { el('div', 'Chọn siêu thị rồi bấm "Đổ tồn kho".', area, 'kxb-empty'); return; }
        const notIn = [...view.invShops].filter(c => !view.inv.shops.includes(c));
        if (notIn.length) el('div', `Siêu thị ${notIn.map(shopName).join(', ')} chưa có số tồn — bấm "Đổ tồn kho" để lấy.`, area, 'kxb-warn');
        const f = view.invFilter, active = [f.category, f.group, f.brand, ...f.conditions, f.q].filter(Boolean);
        if (active.length) el('div', `Đang lọc: ${active.join(' · ')} — bấm "Bỏ lọc" để xem tất cả.`, area, 'kxb-warn');
        const recs = invRecords();
        const v = inventoryViews(recs);
        const times = view.inv.shops.filter(c => view.invShops.has(c)).map(c => `${shopName(c)} lúc ${stamp(new Date(view.inv.shopTimes?.[c] || view.inv.capturedAt))}`).join(' · ');
        el('div', `Tồn: ${times || '—'} · đang xem ${fmt(recs.length)}/${fmt(view.inv.records.filter(r => view.invShops.has(r.shop)).length)} dòng`, area, 'kxb-muted');
        kpis(area, [['SL tồn', fmt(v.quantity), mil(v.cost)], ['Nhóm hàng', fmt(v.groups.length)], ['Sản phẩm', fmt(v.products.length)], ['IMEI / Serial', fmt(v.serials)]]);
        subtabs(area, [['group', 'Theo nhóm hàng'], ['product', 'Theo sản phẩm'], ['imei', 'Danh sách IMEI']], view.invTab, k => { view.invTab = k; renderInventory(); });
        const pane = el('div', undefined, area), codes = view.inv.shops.filter(c => view.invShops.has(c));
        const cond = v.conditionList;
        if (view.invTab === 'group') table(pane, ['Ngành', 'Nhóm hàng', ...codes.map(shopName), 'SL tồn', ...cond, 'Giá trị'],
            v.groups.map(g => [g.category, g.group, ...codes.map(c => fmt(g.byShop[c] || 0)), fmt(g.quantity), ...cond.map(c => fmt(g.byCond[c] || 0)), mil(g.cost)]),
            { num: [...codes.map((_, i) => i + 2), codes.length + 2, ...cond.map((_, i) => codes.length + 3 + i), codes.length + 3 + cond.length] });
        if (view.invTab === 'product') table(pane, ['Nhóm hàng', 'Hãng', 'Mã SP', 'Tên sản phẩm', ...codes.map(shopName), 'SL tồn', ...cond, 'Giá trị'],
            v.products.map(p => [p.group, p.brand, p.product, p.name, ...codes.map(c => fmt(p.byShop[c] || 0)), fmt(p.quantity), ...cond.map(c => fmt(p.byCond[c] || 0)), mil(p.cost)]),
            { num: [...codes.map((_, i) => i + 4), codes.length + 4, ...cond.map((_, i) => codes.length + 5 + i), codes.length + 5 + cond.length] });
        if (view.invTab === 'imei') table(pane, ['Siêu thị', 'Nhóm hàng', 'Hãng', 'Mã SP', 'Tên sản phẩm', 'IMEI / Serial', 'Trạng thái', 'SL', 'Giá vốn', 'Ngày nhập'],
            recs.map(r => [shopName(r.shop), r.group, r.brand, r.product, r.productName, r.serial, r.condition, fmt(r.qty), fmt(Math.round(r.cost)), r.input]), { num: [7, 8] });
    }

    /* ---------- Tab Cân hàng: tồn (Mới) vs tốc độ bán Điện thoại ---------- */
    const BAL_CATEGORY = 'Điện thoại', BAL_CONDITION = 'Mới';
    function balanceWindow() { const to = addDays(isoDate(new Date()), -1); return { from: addDays(to, -view.balDays + 1), to }; }   // N ngày trọn, đến hôm qua
    function balanceData() {
        const range = balanceWindow(), today = isoDate(new Date());
        const { book, lines } = loadBook(range);
        const missing = daysIn(range).filter(d => d <= today && (!book[d] || book[d].schema !== SALES_SCHEMA)).length;
        const shops = view.balShops;
        const isPhone = c => categoryText(c) === BAL_CATEGORY, isNew = c => conditionText(c) === BAL_CONDITION;
        const sales = inPeriod(lines, range, 'created').filter(l => shops.has(l.shop) && isPhone(l.category) && isNew(l.condition));
        const stock = view.inv ? view.inv.records.filter(r => shops.has(r.shop) && isPhone(r.category) && isNew(r.condition)) : [];
        const target = config.balTarget || 14;
        const allRows = balanceRows(stock, sales, { days: view.balDays, target, spare: 1 });
        // Hãng: gộp theo tên không phân biệt hoa thường (BI tồn ghi "Oppo", BI bán ghi "OPPO")
        const brands = new Map(), rev = new Map();
        sales.forEach(l => { if (!l.outShop || keyCode(l.outShop) === keyCode(l.shop)) rev.set(brandKey(l.brand), (rev.get(brandKey(l.brand)) || 0) + l.qty * l.price); });
        allRows.forEach(r => { const k = norm(r.brand) || '(không rõ)'; const b = brands.get(k) || { key: k, label: r.brand || '(không rõ)', need: 0, codes: 0 }; b.need += r.need; if (r.need > 0) b.codes++; brands.set(k, b); });
        [...view.balBrands].forEach(k => { if (!brands.has(k)) view.balBrands.delete(k); });
        const rows = view.balBrands.size ? allRows.filter(r => view.balBrands.has(norm(r.brand) || '(không rõ)')) : allRows;
        const noInv = [...shops].filter(c => !view.inv || !view.inv.shops.includes(c));
        return { range, missing, rows, target, noInv, brands: [...brands.values()].map(b => ({ ...b, rev: rev.get(b.key) || 0 })).sort((a, b) => b.rev - a.rev || b.need - a.need || a.label.localeCompare(b.label, 'vi')) };   // hãng bán nhiều tiền nhất lên đầu
    }
    // Đổ cân hàng = đổ tồn kho các siêu thị đang chọn + xuất bán N ngày còn thiếu, trong một lần bấm
    async function runBalance(session) {
        const keep = view.invShops;
        view.invShops = new Set(view.balShops);
        try { log('Cân hàng: đổ tồn kho'); await runInventory(session); }
        finally { view.invShops = keep; }
        const w = balanceWindow();
        ui.querySelector('[data-from]').value = w.from; ui.querySelector('[data-to]').value = w.to;
        log(`Cân hàng: đổ xuất bán ${toBI(w.from)}–${toBI(w.to)}`);
        const msg = await runSales(session, false);
        renderBalance();
        return `Cân hàng xong: tồn kho mới + xuất bán ${view.balDays} ngày` + (/lỗi|chưa/i.test(msg || '') ? ` · ${msg}` : '');
    }
    function renderBalance() {
        const area = ui.querySelector('[data-bal-result]'); area.replaceChildren();
        const dBox = ui.querySelector('[data-bal-days]'); dBox.replaceChildren();
        [7, 10, 14, 30].forEach(n => { const b = el('button', `${n} ngày`, dBox, 'chip'); b.type = 'button'; b.classList.toggle('on', view.balDays === n); b.onclick = () => { view.balDays = n; save('balDays', n); renderBalance(); }; });
        const sBox = ui.querySelector('[data-bal-shops]'); sBox.replaceChildren();
        config.shops.forEach(sh => {
            const c = keyCode(sh.code), on = view.balShops.has(c), b = el('button', `${sh.code} · ${sh.name}`, sBox, 'chip'); b.type = 'button'; b.classList.toggle('on', on);
            b.onclick = () => { if (on && view.balShops.size === 1) return; on ? view.balShops.delete(c) : view.balShops.add(c); renderBalance(); };
        });
        const d = balanceData();
        multiDrop(ui.querySelector('[data-bal-brands]'), 'bal-brand', d.brands.map(b => ({ key: b.key, name: b.label, label: b.need ? `${b.label} · xin ${fmt(b.need)}` : b.label })),
            view.balBrands, `Tất cả hãng · xin ${fmt(d.brands.reduce((a, b) => a + b.need, 0))}`, renderBalance);
        if (view.balBrands.size) el('div', `Đang lọc hãng: ${d.brands.filter(b => view.balBrands.has(b.key)).map(b => b.label).join(', ')} — tích "Tất cả hãng" để bỏ lọc.`, area, 'kxb-warn');
        el('div', `Điện thoại · hàng Mới · tồn không gồm hàng đang chuyển kho · tốc độ bán = SL bán ${toBI(d.range.from)}–${toBI(d.range.to)} (${view.balDays} ngày) ÷ ${view.balDays} · Sắp hết = tồn chưa đủ bán ${d.target} ngày + 1 máy dự phòng → Nên xin phần thiếu · Tồn nhiều = đủ bán trên ${d.target * 2} ngày · đơn xuất từ kho khác không tính`, area, 'kxb-muted');
        if (!view.inv) { el('div', 'Chưa có số tồn kho — bấm "Đổ cân hàng" để lấy tồn kho và xuất bán.', area, 'kxb-warn'); }
        else {
            if (d.noInv.length) el('div', `Siêu thị ${d.noInv.map(shopName).join(', ')} chưa có số tồn — bấm "Đổ cân hàng".`, area, 'kxb-warn');
            el('div', `Tồn lấy lúc ${stamp(new Date(view.inv.capturedAt))}.`, area, 'kxb-muted');
        }
        if (d.missing) {
            el('div', `Còn ${d.missing} ngày trong ${view.balDays} ngày gần nhất chưa đổ xuất bán (hoặc lưu bằng bản cũ) — tốc độ bán đang thấp hơn thực tế. Bấm "Đổ cân hàng".`, area, 'kxb-warn');
        }
        const counts = {}; d.rows.forEach(r => { counts[r.status] = (counts[r.status] || 0) + 1; });
        const need = d.rows.reduce((a, r) => a + r.need, 0);
        kpis(area, [['Nên xin bổ sung', `SL ${fmt(need)}`, `${fmt(d.rows.filter(r => r.need > 0).length)} mã`], ...['Hết hàng', 'Sắp hết', 'Đủ', 'Tồn nhiều', 'Không bán'].map(k => [k, fmt(counts[k] || 0), 'mã'])]);
        const fBox = el('div', undefined, area, 'group');
        const nNeed = d.rows.filter(r => r.need > 0).length;
        [['', `Tất cả · ${fmt(d.rows.length)}`], ['need', `Cần xin · ${fmt(nNeed)}`], ...STATUS_ORDER.map(s => [s, `${s} · ${fmt(counts[s] || 0)}`])].forEach(([k, l]) => { const b = el('button', l, fBox, 'chip'); b.type = 'button'; b.classList.toggle('on', view.balStatus === k); b.onclick = () => { view.balStatus = k; renderBalance(); }; });
        const rows = d.rows.filter(r => !view.balStatus || (view.balStatus === 'need' ? r.need > 0 : r.status === view.balStatus));
        if (!rows.length && d.rows.length) { el('div', `Không có mã nào ở mục "${view.balStatus === 'need' ? 'Cần xin' : view.balStatus}" với tốc độ bán ${view.balDays} ngày.`, area, 'kxb-empty'); return; }
        table(area, ['Siêu thị', 'Hãng', 'Mã SP', 'Tên sản phẩm', 'Tồn', `Bán ${view.balDays} ngày`, 'TB / ngày', 'Đủ bán (ngày)', 'Nên xin', 'Trạng thái'],
            rows.map(r => [shopName(r.shop), r.brand, r.product, r.name, fmt(r.stock), fmt(r.sold), fmt(r.perDay), r.cover == null ? '—' : fmt(r.cover), r.need ? fmt(r.need) : '', r.status]), { num: [4, 5, 6, 7, 8] });
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

    /* ---------- Ảnh báo cáo gửi Zalo (vẽ canvas, theo bộ lọc đang chọn) ---------- */
    async function salesImage() {
        const r = salesData(); invariant(r && r.lines.length, 'Chưa có số xuất bán trong kỳ — bấm "Đổ xuất bán" hoặc "Xem số đã lưu"');
        const s = r.summary, ps = r.prevSummary, codes = config.shops.map(x => keyCode(x.code)).filter(c => view.salesShops.has(c));
        const att = attachData(r.allLines.filter(l => view.salesShops.has(l.shop))), attMap = new Map(att.staff.map(x => [x.shop + '|' + x.label, x]));
        const pOrders = pendingOrders(r.pendingLines, isoDate(new Date()));
        const brands = s.brands.slice(0, 8), staff = s.staff.slice(0, 12);
        const IW = 1000, P = 28, DPR = 2, rowH = 30;
        const H = 120 + 112 + 20 + 36 + rowH * (brands.length + 1) + 30 + 36 + rowH * (staff.length + 1) + 50;
        const cv = document.createElement('canvas'); cv.width = IW * DPR; cv.height = H * DPR;
        const g = cv.getContext('2d'); g.scale(DPR, DPR);
        const F = (w, px) => `${w} ${px}px Arial, "Segoe UI", sans-serif`;
        const text = (t, x, y, { font = F(400, 15), color = '#172a3a', align = 'left', max = 0 } = {}) => {
            g.font = font; g.fillStyle = color; g.textAlign = align; t = String(t ?? '');
            if (max) while (t.length > 1 && g.measureText(t).width > max) t = t.slice(0, -2) + '…';
            g.fillText(t, x, y);
        };
        const box = (x, y, w, h, fill, rad = 10) => { g.fillStyle = fill; g.beginPath(); g.roundRect ? g.roundRect(x, y, w, h, rad) : g.rect(x, y, w, h); g.fill(); };
        g.fillStyle = '#fff'; g.fillRect(0, 0, IW, H);
        box(0, 0, IW, 96, '#087f8c', 0);
        text(`XUẤT BÁN · ${toBI(r.range.from)} – ${toBI(r.range.to)}`, P, 42, { font: F(700, 26), color: '#fff' });
        const filt = [`Ngành: ${view.salesCat || 'tất cả'}`, view.salesBrands.size ? 'Hãng: ' + [...new Set(r.lines.map(l => l.brand))].join(', ') : '', view.salesConds.size ? 'Loại hàng: ' + [...view.salesConds].join(', ') : '', 'Kho tạo · Đã xuất – Đã giao · bỏ đơn nhập trả'].filter(Boolean).join(' · ');
        text(filt, P, 72, { font: F(400, 14), color: '#e0f2f3', max: IW - 2 * P });
        // Thẻ chỉ số
        const d = (c, p) => { const x = delta(c, p, r); return x ? x.text : ''; };
        const cards = [['Cụm', s.quantity, s.revenue, d(s.revenue, ps.revenue), true],
            ...codes.map(c => { const x = s.shops.find(z => z.code === c) || { quantity: 0, revenue: 0 }, y = ps.shops.find(z => z.code === c) || { revenue: 0 }; return [shopName(c), x.quantity, x.revenue, d(x.revenue, y.revenue)]; })];
        const extra = [['Bán kèm ĐT', pct(att.total.rate), `${fmt(att.total.attached)}/${fmt(att.total.phoneOrders)} đơn`], ['Đơn treo', fmt(pOrders.length), `${fmt(pOrders.filter(o => o.age >= 2).length)} đơn từ 2 ngày`]];
        const all = [...cards.map(c => ({ t: c[0], v: `SL ${fmt(c[1])}`, s: mil(c[2]), d: c[3], main: c[4] })), ...extra.map(c => ({ t: c[0], v: c[1], s: c[2] }))];
        const cw = (IW - 2 * P - 10 * (all.length - 1)) / all.length;
        all.forEach((c, i) => {
            const x = P + i * (cw + 10), y = 116;
            box(x, y, cw, 96, c.main ? '#087f8c' : '#f2f6f7');
            const col = c.main ? '#fff' : '#172a3a';
            text(c.t, x + 12, y + 22, { font: F(400, 13), color: c.main ? '#e0f2f3' : '#4a5a66', max: cw - 20 });
            text(c.v, x + 12, y + 50, { font: F(700, 22), color: col, max: cw - 20 });
            text(c.s, x + 12, y + 70, { font: F(400, 13), color: col, max: cw - 20 });
            if (c.d) text(c.d, x + 12, y + 88, { font: F(700, 12), color: c.main ? '#fff' : c.d.startsWith('▼') ? '#b91c1c' : '#15803d', max: cw - 20 });
        });
        // Bảng
        const tableAt = (y, title, cols, rows) => {
            text(title, P, y, { font: F(700, 17) });
            y += 14; box(P, y, IW - 2 * P, rowH, '#edf5f7', 6);
            const draw = (cells, yy, bold) => cols.forEach(([, x, w, al], i) => text(cells[i], al === 'right' ? x + w : x, yy + 20, { font: F(bold ? 700 : 400, 14), align: al || 'left', max: w }));
            draw(cols.map(c => c[0]), y, true);
            rows.forEach((row, i) => { const yy = y + rowH * (i + 1); if (i % 2) box(P, yy, IW - 2 * P, rowH, '#f8fafb', 0); draw(row, yy, false); });
            return y + rowH * (rows.length + 1);
        };
        let y = tableAt(252, 'Theo hãng', [['Hãng', P + 10, 280], ['SL', P + 300, 90, 'right'], ['% SL', P + 400, 90, 'right'], ['Doanh thu', P + 500, 150, 'right'], ['% DT', P + 660, 90, 'right'], ['Giá TB', P + 760, 170, 'right']],
            brands.map(b => [b.label, fmt(b.quantity), pct(b.pctQty), mil(b.revenue), pct(b.pctRev), mil(b.quantity ? b.revenue / b.quantity : 0)]));
        y = tableAt(y + 36, `Nhân viên${s.staff.length > staff.length ? ` (top ${staff.length}/${s.staff.length})` : ''}`, [['Nhân viên', P + 10, 330], ['Siêu thị', P + 350, 150], ['SL', P + 510, 80, 'right'], ['Doanh thu', P + 600, 140, 'right'], ['Kèm ĐT', P + 750, 180, 'right']],
            staff.map(st => { const a = attMap.get(st.shop + '|' + st.label); return [st.label, shopName(st.shop), fmt(st.quantity), mil(st.revenue), a ? `${fmt(a.attached)}/${fmt(a.phoneOrders)} · ${pct(a.rate)}` : '—']; }));
        text(`AutoBI Kho & Xuất Bán V${VERSION} · lấy lúc ${stamp(new Date())}${r.missing ? ` · CHƯA ĐỦ: còn ${r.missing} ngày chưa lấy` : ''}`, P, y + 30, { font: F(400, 12), color: r.missing ? '#b91c1c' : '#5b6b76' });
        const blob = await new Promise(res => cv.toBlob(res, 'image/png'));
        invariant(blob, 'Không tạo được ảnh');
        const name = `AutoBI_XuatBan_${r.range.from.replace(/-/g, '')}-${r.range.to.replace(/-/g, '')}.png`;
        const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000);
        let copied = false;
        try { if (window.ClipboardItem) { await navigator.clipboard.write([new (window.ClipboardItem)({ 'image/png': blob })]); copied = true; } } catch { /* trình duyệt chặn → chỉ tải file */ }
        status(copied ? `Đã chép ảnh báo cáo — mở Zalo bấm Ctrl+V để dán (cũng đã tải ${name})` : `Đã tải ảnh ${name}`, 'ok');
        log(`Ảnh báo cáo xuất bán: ${toBI(r.range.from)}–${toBI(r.range.to)}`);
    }

    /* ---------- Excel ---------- */
    function exportExcel() {
        const X = XL(); invariant(X, 'Chưa tải được thư viện Excel (SheetJS)');
        const wb = X.utils.book_new();
        const sheet = (name, aoa, widths) => { const ws = X.utils.aoa_to_sheet(aoa); if (widths) ws['!cols'] = widths.map(w => ({ wch: w })); X.utils.book_append_sheet(wb, ws, name); };
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
            sheet('Ban_ChiTiet', [['Ngày tạo', 'Giờ', 'Ngày xuất', 'Kho tạo', 'Siêu thị', 'Mã đơn', 'Loại YCX', 'Nhân viên', 'Hãng', 'Mã SP', 'Tên SP', 'IMEI', 'Loại hàng', 'Ngành', 'Nhóm', 'SL', 'Giá bán', 'Giá trước VAT', 'Doanh thu'],
                ...r.lines.map(l => [toBI(l.created), l.time, l.shipped ? toBI(l.shipped) : '', l.shop, shopName(l.shop), l.order, l.orderType, l.creator, l.brand, l.product, l.productName, l.imei || '', conditionText(l.condition), l.category, l.group, l.qty, l.price, l.priceNet, l.qty * l.price])], [11, 6, 11, 8, 14, 20, 24, 26, 10, 16, 40, 18, 12, 22, 20, 5, 12, 12, 13]);
            X.writeFile(wb, `AutoBI_XuatBan_${r.range.from.replace(/-/g, '')}-${r.range.to.replace(/-/g, '')}.xlsx`);
        } else if (view.tab === 'balance') {
            const d = balanceData(); invariant(view.inv, 'Chưa có số tồn kho');
            sheet('CanHang', [[`Cân hàng Điện thoại (hàng Mới)${view.balBrands.size ? ' · hãng ' + d.brands.filter(b => view.balBrands.has(b.key)).map(b => b.label).join(', ') : ''}${view.balStatus ? ' · ' + (view.balStatus === 'need' ? 'Cần xin' : view.balStatus) : ''} · bán ${toBI(d.range.from)}–${toBI(d.range.to)} (${view.balDays} ngày) · giữ đủ ${d.target} ngày · tồn lúc ${stamp(new Date(view.inv.capturedAt))}` + (d.missing ? ` · CHƯA ĐỦ: còn ${d.missing} ngày chưa đổ xuất bán` : '')], [],
                ['Siêu thị', 'Hãng', 'Mã SP', 'Tên sản phẩm', 'Tồn', `Bán ${view.balDays} ngày`, 'TB / ngày', 'Đủ bán (ngày)', 'Nên xin', 'Trạng thái'],
                ...d.rows.filter(r => !view.balStatus || (view.balStatus === 'need' ? r.need > 0 : r.status === view.balStatus)).map(r => [shopName(r.shop), r.brand, r.product, r.name, r.stock, r.sold, r.perDay, r.cover ?? '', r.need, r.status])], [14, 12, 16, 44, 7, 9, 9, 11, 8, 11]);
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
            sheet('TonKho_IMEI', [['Mã ST', 'Siêu thị', 'Ngành', 'Nhóm hàng', 'Hãng', 'Mã SP', 'Tên sản phẩm', 'IMEI / Serial', 'Trạng thái', 'SL', 'Giá vốn (đ)', 'Ngày nhập'],
                ...recs.map(r => [r.shop, shopName(r.shop), r.category, r.group, r.brand, r.product, r.productName, r.serial, r.condition, r.qty, Math.round(r.cost), r.input])], [7, 14, 22, 26, 12, 16, 40, 20, 16, 6, 13, 16]);
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
        return [head, ...[...byShop.entries()].flatMap(([shop, os]) => [`\n${shopName(shop)} — ${os.length} đơn:`,
            ...os.map(o => `• ${o.creator.replace(/^\d+\s*[-–]\s*/, '')} · ${o.items.join(' + ')} · ${o.age} ngày · ${o.state.toLowerCase()} · đơn ${o.order}`)]),
            '\nAnh em kiểm tra, xuất / giao hoặc hủy đơn giúp nhé.'].join('\n');
    }
    function renderStorage() {
        const box = ui.querySelector('[data-storage]'); if (!box) return;
        try {
            const keys = (typeof GM_listValues === 'function' ? GM_listValues() : []).filter(k => k.startsWith(PREFIX + 'month_')).sort();
            let bytes = 0, lines = 0; keys.forEach(k => { const v = load(k.slice(PREFIX.length), null); if (v) { bytes += JSON.stringify(v).length; lines += (v.lines || []).length; } });
            const mb = bytes / 1048576, inv = JSON.stringify(load('lastInventory', null) || '').length / 1048576;
            box.textContent = `Bộ nhớ trên máy: xuất bán ${keys.length} tháng (${keys.map(k => k.slice(-7).split('-').reverse().join('/')).join(', ') || '—'}) · ${fmt(lines)} dòng · ${fmt(Math.round(mb * 10) / 10)} MB · tồn kho ${fmt(Math.round(inv * 10) / 10)} MB`
                + (mb + inv > 40 ? ' — khá lớn, Tampermonkey có thể chậm khi lưu.' : '');
            box.style.color = mb + inv > 40 ? '#8a1c1c' : '';
        } catch (e) { box.textContent = 'Không đọc được dung lượng: ' + e.message; }
    }
    function shopEditor(shop = { code: '', name: '' }) {
        const row = el('div', undefined, ui.querySelector('[data-shops]'), 'kxb-shop');
        for (const [key, label] of [['code', 'Mã ST'], ['name', 'Tên hiển thị'], ['target', 'Mục tiêu DT tháng (triệu)']]) {
            const l = el('label', label, row), input = el('input', undefined, l); input.dataset.shopField = key; input.value = shop[key] || '';
            if (key === 'target') { input.inputMode = 'numeric'; input.placeholder = 'vd 1700'; input.style.width = '150px'; }
        }
        const del = el('button', 'Xóa', row); del.type = 'button'; del.onclick = () => { invariant(!running, 'Không đổi cấu hình khi đang chạy'); row.remove(); };
    }
    function saveConfig() {
        invariant(!running, 'Không đổi cấu hình khi đang chạy');
        const shops = [...ui.querySelectorAll('.kxb-shop')].map(row => Object.fromEntries([...row.querySelectorAll('input')].map(input => [input.dataset.shopField, clean(input.value)])));
        validateShops(shops);
        shops.forEach(sh => {
            const t = clean(sh.target).replace(/[.,\s]/g, '');
            invariant(!t || /^\d{1,7}$/.test(t), `Mục tiêu của ${sh.name} phải là số triệu đồng (vd 1700)`);
            if (t) sh.target = Number(t); else delete sh.target;
        });
        const v = k => clean(ui.querySelector(`[data-cfg="${k}"]`).value);
        const fh = Math.round(Number(v('freshHours')));
        invariant(fh >= 0 && fh <= 24, 'Số giờ "vừa lấy" phải từ 0 đến 24');
        const rd = Math.round(Number(v('returnDays')));
        invariant(rd >= 2 && rd <= 35, 'Số ngày kiểm tra lại nhập trả phải từ 2 đến 35');
        const bt = Math.round(Number(v('balTarget')));
        invariant(bt >= 3 && bt <= 60, 'Số ngày giữ đủ bán (Cân hàng) phải từ 3 đến 60');
        config = { ...config, shops, selectors: config.selectors || {}, basis: v('basis') === 'shipped' ? 'shipped' : 'created', returnDays: rd, balTarget: bt,
            freshHours: fh, autoDaily: ui.querySelector('[data-cfg="autoDaily"]').checked, targetCat: v('targetCat') === '*' ? '' : v('targetCat') };
        save('config', config);
        const codes = new Set(shops.map(s => keyCode(s.code)));
        view.invShops = new Set([...view.invShops].filter(c => codes.has(c))); if (!view.invShops.size) view.invShops = new Set(codes);
        view.salesShops = new Set(codes); view.balShops = new Set(codes); if (view.sales) view.sales = salesView(view.sales.range);
        log('Đã lưu cài đặt'); status('Đã lưu cài đặt', 'ok');
        ui.querySelector('[data-settings]').open = false; renderAll();
    }
    function safely(fn) { return async (...args) => { try { await fn(...args); } catch (e) { status(e.message, 'err'); log(e.message, 'error'); } }; }
    function renderAll() {
        ui.querySelectorAll('[data-pane]').forEach(p => p.hidden = p.dataset.pane !== view.tab);
        ui.querySelectorAll('[data-tab]').forEach(b => { b.classList.toggle('on', b.dataset.tab === view.tab); b.setAttribute('aria-selected', b.dataset.tab === view.tab); });
        renderSales(); renderInventoryFilters(); renderInventory(); renderBalance();
    }

    /* ---------- Thông báo bản mới ---------- */
    const UPDATE_EVERY = 3 * 3600 * 1000;               // tự kiểm tra GitHub tối đa 3 giờ/lần
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
            const a = el('a', '⬇ Cập nhật ngay', row, 'btn primary'); a.href = UPDATE_URL; a.target = '_blank'; a.rel = 'noopener';
            el('span', 'Tampermonkey mở trang cài → bấm "Cập nhật" → quay lại đây tải lại trang BI (F5).', row, 'kxb-muted');
            return;
        }
    }

    /* ---------- Giao diện: khung lớn giữa màn hình (máy tính) ---------- */
    function mount() {
        if (document.querySelector('[data-kxb-ui]')) return;
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
        `, document.head || document.documentElement); style.dataset.kxbUi = '';
        const launch = el('button', 'AutoBI · Kho & Xuất Bán', document.body); launch.id = 'kxb-launch'; launch.dataset.kxbUi = ''; launch.type = 'button';
        const back = el('div', undefined, document.body); back.id = 'kxb-back'; back.dataset.kxbUi = ''; back.hidden = true;
        ui = el('div', undefined, back); ui.id = 'kxb-panel'; ui.setAttribute('role', 'dialog'); ui.setAttribute('aria-label', 'AutoBI Kho và Xuất bán');
        ui.innerHTML = `
        <div class="top"><strong>AutoBI · Kho &amp; Xuất Bán</strong><span class="kxb-muted">V${VERSION}</span>
          <div class="tabs" role="tablist"><button type="button" role="tab" data-tab="sales">🛒 Xuất bán</button><button type="button" role="tab" data-tab="inventory">📦 Tồn kho</button><button type="button" role="tab" data-tab="balance">⚖️ Cân hàng</button></div>
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
              <label>Mục tiêu DT áp cho ngành<select data-cfg="targetCat"><option value="Điện thoại">Điện thoại</option><option value="*">Tất cả ngành</option></select></label>
              <label style="flex-direction:row;align-items:center;gap:6px;font-size:13px"><input type="checkbox" data-cfg="autoDaily" style="min-height:auto">Mở BI lần đầu trong ngày: tự lấy xuất bán hôm qua</label></div>
            <div class="kxb-muted" data-storage style="margin-bottom:8px"></div>
            <button type="button" class="primary" data-save>Lưu cài đặt</button> <button type="button" data-check-update>🔄 Kiểm tra bản mới</button></details>
          <div data-status class="kxb-status">Sẵn sàng.</div><div class="prog"><div data-bar></div></div>
          <div data-pane="sales">
            <div class="filters">
              <div class="group"><b>Kỳ</b><span class="group" data-presets></span></div>
              <div class="bar" style="margin:0"><label>Từ ngày<input type="date" data-from></label><label>Đến ngày<input type="date" data-to></label>
                <label style="flex-direction:row;align-items:center;gap:6px;font-size:13px"><input type="checkbox" data-refetch style="min-height:auto">Lấy lại cả ngày đã chốt</label>
                <span style="flex:1"></span><button type="button" data-img title="Tạo ảnh báo cáo theo bộ lọc đang chọn: chép vào bộ nhớ để dán Zalo, đồng thời tải file PNG">🖼 Ảnh báo cáo</button><button type="button" class="idle-only" data-view>Xem số đã lưu</button><button type="button" class="primary idle-only" data-run-sales>Đổ xuất bán</button></div>
              <div class="group"><b>Siêu thị</b><span class="group" data-sales-shops></span></div>
              <div class="group"><b>Ngành hàng</b><select data-sales-cat style="min-width:320px"></select><button type="button" class="chip" data-phone-only>📱 Chỉ điện thoại</button><b style="margin-left:12px">Hãng</b><span data-sales-brands></span></div>
              <div class="group"><b>Loại hàng</b><span class="group" data-sales-conds></span></div></div>
            <div data-sales-result></div></div>
          <div data-pane="inventory" hidden>
            <div class="filters" data-inv-filters>
              <div class="group"><b>Siêu thị</b><span class="group" data-inv-shops></span><span class="sp" style="flex:1"></span><button type="button" class="idle-only" data-print-inv title="In danh sách đang lọc ra giấy A4 dọc, có ô KIỂM để tích">🖨 In phiếu kiểm</button><button type="button" class="primary idle-only" data-run-inv>Đổ tồn kho</button></div>
              <div class="group"><b>Lọc</b><select data-f="category"></select><select data-f="group"></select><select data-f="brand"></select>
                <input type="search" data-f="q" placeholder="Tìm IMEI, mã hoặc tên sản phẩm" style="min-width:280px"><button type="button" data-clear>Bỏ lọc</button></div>
              <div class="group"><b>Trạng thái</b><span class="group" data-inv-conds></span></div></div>
            <div data-inv-result></div></div>
          <div data-pane="balance" hidden>
            <div class="filters">
              <div class="group"><b>Tốc độ bán</b><span class="group" data-bal-days></span><span class="sp" style="flex:1"></span><button type="button" class="primary idle-only" data-run-bal title="Đổ tồn kho mới + xuất bán các ngày còn thiếu">Đổ cân hàng</button></div>
              <div class="group"><b>Siêu thị</b><span class="group" data-bal-shops></span></div>
              <div class="group"><b>Hãng</b><span data-bal-brands></span></div></div>
            <div data-bal-result></div></div>
          <details data-logbox><summary class="logbar">📋 Nhật ký <button type="button" data-copy-log>Sao chép</button><button type="button" data-clear-log>Xóa nhật ký</button></summary><pre data-log></pre></details>
        </div>`;
        const today = isoDate(new Date());
        ui.querySelector('[data-from]').value = today.slice(0, 8) + '01'; ui.querySelector('[data-to]').value = today;
        ui.querySelector('[data-cfg="basis"]').value = config.basis || 'created'; ui.querySelector('[data-cfg="returnDays"]').value = config.returnDays || 7; ui.querySelector('[data-cfg="balTarget"]').value = config.balTarget || 14;
        ui.querySelector('[data-cfg="freshHours"]').value = config.freshHours ?? 2; ui.querySelector('[data-cfg="autoDaily"]').checked = config.autoDaily !== false;
        ui.querySelector('[data-cfg="targetCat"]').value = (config.targetCat ?? 'Điện thoại') || '*';
        ui.querySelector('[data-settings]').addEventListener('toggle', e => { if (e.target.open) renderStorage(); });
        (config.shops.length ? config.shops : [{ code: '', name: '' }]).forEach(shopEditor);
        if (!config.shops.length) ui.querySelector('[data-settings]').open = true;
        const savedShops = load('invShops', null);
        const codes = config.shops.map(s => keyCode(s.code));
        view.salesShops = new Set(codes); view.balShops = new Set(codes);
        view.salesCat = String(load('salesCat', BAL_CATEGORY) ?? '');         // lần đầu: chọn sẵn ngành Điện thoại
        view.balDays = [7, 10, 14, 30].includes(load('balDays', 10)) ? load('balDays', 10) : 10;
        view.invShops = new Set(Array.isArray(savedShops) && savedShops.some(c => codes.includes(c)) ? savedShops.filter(c => codes.includes(c)) : codes);
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
        const open = on => { back.hidden = !on; launch.hidden = on; if (on) { renderAll(); renderUpdate(); checkUpdate(false); } };
        launch.onclick = () => open(true);
        ui.querySelector('[data-close]').onclick = () => open(false);
        back.addEventListener('mousedown', e => { if (e.target === back && !running) open(false); });
        document.addEventListener('keydown', e => { if (e.key === 'Escape' && !back.hidden && !running) open(false); });
        ui.querySelectorAll('[data-tab]').forEach(b => b.onclick = () => { view.tab = b.dataset.tab; renderAll(); });
        ui.querySelector('[data-stop]').onclick = stop;
        ui.querySelector('[data-add]').onclick = safely(() => { invariant(!running, 'Không sửa cài đặt khi đang chạy'); shopEditor(); });
        ui.querySelector('[data-save]').onclick = safely(saveConfig);
        ui.querySelector('[data-excel]').onclick = safely(exportExcel);
        ui.querySelector('[data-copy-log]').onclick = e => { e.preventDefault(); e.stopPropagation(); copyLog(); };
        ui.querySelector('[data-clear-log]').onclick = e => { e.preventDefault(); e.stopPropagation(); journal.length = 0; try { save('journal', journal); } catch { /* bỏ qua */ } log('Đã xóa nhật ký cũ'); };
        if (config.shops.length) view.sales = salesView(selectedRange());
        ui.querySelector('[data-run-sales]').onclick = safely(() => { validateShops(config.shops); selectedRange(); return withSession('sales', s => runSales(s, ui.querySelector('[data-refetch]').checked)); });
        ui.querySelector('[data-img]').onclick = safely(salesImage);
        ui.querySelector('[data-view]').onclick = safely(() => { validateShops(config.shops); view.sales = salesView(selectedRange()); renderSales(); status(view.sales.missing ? `Còn ${view.sales.missing} ngày chưa lấy trong kỳ` : 'Số đã lưu trên máy này', view.sales.missing ? 'warn' : 'ok'); });
        ui.querySelector('[data-print-inv]').onclick = safely(printInventory);
        ui.addEventListener('mousedown', e => { if (!e.target.closest('.kxb-drop')) { view.openDrop = ''; ui.querySelectorAll('.kxb-drop[open]').forEach(d => d.open = false); } });
        ui.querySelector('[data-run-bal]').onclick = safely(() => { validateShops(config.shops); return withSession('balance', runBalance); });
        ui.querySelector('[data-run-inv]').onclick = safely(() => { validateShops(config.shops); return withSession('inventory', runInventory); });
        for (const k of ['category', 'group', 'brand']) ui.querySelector(`[data-f="${k}"]`).onchange = e => { view.invFilter[k] = e.target.value; if (k === 'category') view.invFilter.group = ''; renderInventoryFilters(); renderInventory(); };
        let tq; ui.querySelector('[data-f="q"]').oninput = e => { clearTimeout(tq); tq = setTimeout(() => { view.invFilter.q = e.target.value; renderInventory(); }, 250); };
        ui.querySelector('[data-clear]').onclick = () => { view.invFilter = { category: '', group: '', brand: '', conditions: [], q: '' }; ui.querySelector('[data-f="q"]').value = ''; renderInventoryFilters(); renderInventory(); };
        ui.querySelector('[data-check-update]').onclick = safely(() => checkUpdate(true));
        renderUpdate();
        setTimeout(() => checkUpdate(false), 3000);
        setTimeout(() => { autoYesterday().catch(e => log('Tự đổ hôm qua: ' + e.message, 'error')); }, 6000);
        log(`Mở AutoBI Kho & Xuất Bán V${VERSION}.`);
    }
    window.addEventListener('pagehide', () => { if (running) stop(); });
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mount, { once: true }); else mount();
})();
