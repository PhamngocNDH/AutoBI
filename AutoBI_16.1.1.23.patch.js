/* ===== AutoBI 16.1.1.23 - Trả chậm SAFE CAPTURE =====
   Mục tiêu của bản test này:
   - KHÔNG tự bấm Vùng/Khu vực/Siêu thị.
   - Chỉ ghi dữ liệu khi bảng hiện ĐỒNG THỜI Hải Anh + Trực Cường.
   - Kiểm tra chéo DT trả góp / DT siêu thị / tỷ trọng trước khi ghi cache.
   - Không đụng các module AutoBI đang chạy ổn.
*/
(function () {
    'use strict';
    if (window.top !== window) return;

    const DATA_KEY = 'tgdd_data_cache_v30';
    const STATUS_KEY = 'autobi_tc23_status';
    const TOAST_ONCE_KEY = 'autobi_tc23_toast_once';
    const PATH_TC = '/dashboard/tra-cham';

    const norm = (s) => String(s || '')
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/g, 'd').replace(/Đ/g, 'D')
        .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();

    const text = (el) => (el?.innerText || el?.textContent || '').trim();
    const num = (v) => {
        const n = parseFloat(String(v ?? '').replace(/,/g, '').replace(/%/g, '').trim());
        return Number.isFinite(n) ? n : 0;
    };
    const gmGet = (k, d = null) => { try { return GM_getValue(k, d); } catch (_) { return d; } };
    const gmSet = (k, v) => { try { GM_setValue(k, v); } catch (_) {} };

    function toast(msg, ms = 3800) {
        try {
            const w = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
            if (w.UI && typeof w.UI.showToast === 'function') return w.UI.showToast(msg, ms);
        } catch (_) {}
        console.log('[AutoBI 23 Trả chậm]', msg);
    }

    function ensureBadge() {
        if (!location.pathname.includes(PATH_TC)) return null;
        let box = document.getElementById('autobi-tc23-badge');
        if (box) return box;
        box = document.createElement('div');
        box.id = 'autobi-tc23-badge';
        box.style.cssText = [
            'position:fixed','right:18px','bottom:82px','z-index:2147483647',
            'max-width:430px','padding:10px 12px','border-radius:10px',
            'font:600 12px/1.45 Arial,sans-serif','box-shadow:0 3px 14px rgba(0,0,0,.22)',
            'background:#fff3cd','color:#664d03','border:1px solid #ffecb5'
        ].join(';');
        box.innerHTML = '⏳ AutoBI 23: Chờ bảng <b>Siêu thị</b> có cả <b>Hải Anh + Trực Cường</b>...';
        document.documentElement.appendChild(box);
        return box;
    }

    function setBadgeWaiting(msg) {
        const box = ensureBadge();
        if (!box) return;
        box.style.background = '#fff3cd';
        box.style.color = '#664d03';
        box.style.borderColor = '#ffecb5';
        box.innerHTML = '⏳ AutoBI 23: ' + msg;
    }

    function setBadgeSuccess(rows) {
        const box = ensureBadge();
        if (!box) return;
        box.style.background = '#d1e7dd';
        box.style.color = '#0f5132';
        box.style.borderColor = '#badbcc';
        const ha = rows.shop1, tc = rows.shop2;
        box.innerHTML = `✅ <b>ĐÃ LẤY ĐÚNG TRẢ CHẬM</b><br>` +
            `Hải Anh: ST ${ha.dtSieuThi} | TC ${ha.dtTraGop} | HC ${ha.hc} | FE ${ha.fe} | ${ha.tyTrong}%<br>` +
            `Trực Cường: ST ${tc.dtSieuThi} | TC ${tc.dtTraGop} | HC ${tc.hc} | FE ${tc.fe} | ${tc.tyTrong}%<br>` +
            `<button id="autobi-tc23-open" style="margin-top:7px;padding:5px 10px;border:0;border-radius:7px;background:#198754;color:white;font-weight:700;cursor:pointer">Mở AutoBI</button>`;
        const btn = box.querySelector('#autobi-tc23-open');
        if (btn) btn.onclick = () => { location.href = '/dashboard/thi-dua'; };
    }

    function classifyRow(label) {
        const s = norm(label);
        if (s.includes('hha') || s.includes('hai anh')) return 'shop1';
        if (s.includes('tni') || s.includes('truc cuong')) return 'shop2';
        return null;
    }

    function findCorrectTable() {
        const tables = [...document.querySelectorAll('table')];
        for (const table of tables) {
            const head = norm(text(table.querySelector('thead')) || text(table));
            if (!head.includes('dt tra gop') || !head.includes('dt sieu thi')) continue;
            if (!head.includes('homecredit') || !head.includes('fecredit')) continue;

            let hasHA = false, hasTC = false;
            for (const tr of table.querySelectorAll('tbody tr')) {
                const first = tr.querySelector('td');
                const k = classifyRow(text(first));
                if (k === 'shop1') hasHA = true;
                if (k === 'shop2') hasTC = true;
            }
            if (hasHA && hasTC) return table;
        }
        return null;
    }

    function parseRow(tr) {
        const cells = [...tr.querySelectorAll('td')];
        if (cells.length < 7) return null;
        const key = classifyRow(text(cells[0]));
        if (!key) return null;

        const row = {
            key,
            label: text(cells[0]),
            dtTraGop: num(text(cells[1])),
            dtSieuThi: num(text(cells[2])),
            tyTrong: num(text(cells[3])),
            hc: num(text(cells[4])),
            fe: num(text(cells[6])),
            kredivo: cells[8] ? num(text(cells[8])) : 0
        };

        if (!(row.dtSieuThi > 0) || row.dtTraGop < 0) return null;
        if (!(row.tyTrong >= 0 && row.tyTrong <= 100)) return null;
        const pctCalc = row.dtTraGop / row.dtSieuThi * 100;
        if (Math.abs(pctCalc - row.tyTrong) > 1.25) return null;
        if (row.hc < 0 || row.fe < 0 || row.kredivo < 0) return null;
        if (row.hc > row.dtTraGop * 1.1 || row.fe > row.dtTraGop * 1.1) return null;
        const partnerSum = row.hc + row.fe + row.kredivo;
        if (row.kredivo > 0 && Math.abs(partnerSum - row.dtTraGop) > Math.max(3, row.dtTraGop * 0.03)) return null;
        return row;
    }

    function saveRows(rows) {
        const cache = gmGet(DATA_KEY, {}) || {};
        cache.link7 = cache.link7 || {};
        cache.link2 = cache.link2 || {};

        let sumTra = 0, sumST = 0;
        for (const key of ['shop1', 'shop2']) {
            const r = rows[key];
            cache.link7[key] = cache.link7[key] || {};
            cache.link7[key]._SHOP_DT = r.dtSieuThi;
            cache.link7[key]._SHOP_HC = r.hc;
            cache.link7[key]._SHOP_FE = r.fe;
            cache.link7[key]._SHOP_TOTAL_ = {
                tg: r.tyTrong,
                dt_sieu_thi: r.dtSieuThi,
                dt_tragop: r.dtTraGop,
                hc: r.hc,
                fe: r.fe,
                kredivo: r.kredivo,
                source: 'tra-cham-v23'
            };
            cache.link2[key] = cache.link2[key] || {};
            cache.link2[key].tg = r.tyTrong;
            sumTra += r.dtTraGop;
            sumST += r.dtSieuThi;
        }
        cache.link2.total = cache.link2.total || {};
        if (sumST > 0) cache.link2.total.tg = sumTra / sumST * 100;

        gmSet(DATA_KEY, cache);
        gmSet(STATUS_KEY, {
            ok: true,
            time: Date.now(),
            shop1: rows.shop1,
            shop2: rows.shop2
        });
    }

    let lastSignature = '';
    function captureSafe() {
        if (!location.pathname.includes(PATH_TC)) return false;
        ensureBadge();

        const table = findCorrectTable();
        if (!table) {
            setBadgeWaiting('Chọn <b>Vùng Hà Nội+</b> → <b>Khu vực Nam Định</b> → để <b>Siêu thị không chọn riêng</b> → bấm <b>Xem</b>.');
            return false;
        }

        const rows = {};
        for (const tr of table.querySelectorAll('tbody tr')) {
            const r = parseRow(tr);
            if (r) rows[r.key] = r;
        }
        if (!rows.shop1 || !rows.shop2) {
            setBadgeWaiting('Đã thấy 2 ST nhưng dữ liệu chưa qua kiểm tra an toàn, chưa ghi cache.');
            return false;
        }

        const sig = JSON.stringify([
            rows.shop1.dtTraGop, rows.shop1.dtSieuThi, rows.shop1.tyTrong, rows.shop1.hc, rows.shop1.fe,
            rows.shop2.dtTraGop, rows.shop2.dtSieuThi, rows.shop2.tyTrong, rows.shop2.hc, rows.shop2.fe
        ]);
        if (sig !== lastSignature) {
            lastSignature = sig;
            saveRows(rows);
            setBadgeSuccess(rows);
            toast('✅ Đã lưu đúng Trả chậm Hải Anh + Trực Cường. Bấm Mở AutoBI để kiểm tra.', 5200);
        } else {
            setBadgeSuccess(rows);
        }
        return true;
    }

    function showAutoBIStatusOnce() {
        if (!location.pathname.includes('/dashboard/thi-dua')) return;
        const st = gmGet(STATUS_KEY, null);
        if (!st?.ok || !st.time || Date.now() - st.time > 6 * 60 * 60 * 1000) return;
        const once = gmGet(TOAST_ONCE_KEY, 0);
        if (once === st.time) return;
        gmSet(TOAST_ONCE_KEY, st.time);
        setTimeout(() => toast('✅ Trả chậm đã lưu. Vào Sức khỏe ST để kiểm tra 4 dòng DT ST / DT trả chậm / HomeCredit / FECredit.', 5500), 2200);
    }

    if (location.pathname.includes(PATH_TC)) {
        const start = () => {
            ensureBadge();
            captureSafe();
            new MutationObserver(() => setTimeout(captureSafe, 180)).observe(document.documentElement, { childList: true, subtree: true, characterData: true });
            setInterval(captureSafe, 1600);
        };
        if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
        else start();
    } else {
        showAutoBIStatusOnce();
    }
})();
