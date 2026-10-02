// ==UserScript==
// @name         AutoBI - Phiếu Mua Hàng HNO+
// @namespace    https://github.com/PhamngocNDH/AutoBI/pmh-hnoplus
// @version      1.3.2
// @description  Tự lấy phiếu mua hàng kho 01 (ICT) và 02 (CE) trên trang megalive: bảng lọc theo ngành/hãng, nhật ký, nút dừng, copy Line.
// @author       AutoBI / 38967 - Mr Phạm
// @homepageURL  https://github.com/PhamngocNDH/AutoBI
// @updateURL    https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI_PMH_HNO.user.js
// @downloadURL  https://raw.githubusercontent.com/PhamngocNDH/AutoBI/main/AutoBI_PMH_HNO.user.js
// @match        https://vung.hnoplus.com/megalive/*
// @match        https://vung.hnoplus.com/megalive
// @run-at       document-start
// @noframes
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_setClipboard
// @grant        unsafeWindow
// ==/UserScript==

/*
 * NGUYÊN TẮC
 * - Chỉ chạy ở khung chính (@noframes), mỗi lần chỉ 1 phiên, có nút Dừng.
 * - Tự chạy khi mở trang nhưng có chốt chống lặp: không tự chạy lại trong 3 phút,
 *   và nếu đã có kết quả trong 30 phút gần nhất thì chỉ mở bảng cũ.
 * - Bộ lọc phần tử như AutoBI: bỏ qua phần tử ẩn và phần tử do script tạo ra.
 * - Kho lỗi / hết giờ chờ → ghi Nhật ký, không bịa dữ liệu. Bấm Dừng → bỏ phần đang lấy dở.
 */
(function () {
    'use strict';
    const VERSION = '1.3.2';
    const PREFIX = 'autobi_pmh_v1_';

    /* ===== CẤU HÌNH ===== */
    const KHO = [
        { code: '01', nganh: 'ICT' },
        { code: '02', nganh: 'CE' },
    ];
    const AUTO_RUN = true;          // mở trang là tự lấy
    const WAIT_MAX = 15000;         // chờ tối đa mỗi kho (ms)
    const FRESH_MIN = 30;           // có kết quả trong 30 phút thì không tự lấy lại
    /* ==================== */

    const W = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
    const load = (k, f) => { try { return GM_getValue(PREFIX + k, f); } catch { return f; } };
    const save = (k, v) => { try { GM_setValue(PREFIX + k, v); } catch { /* bỏ qua */ } };
    const clean = v => String(v ?? '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
    const norm = v => clean(v).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/gi, 'd').toLowerCase();
    const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    const pad = n => String(n).padStart(2, '0');
    const dayKey = (t = Date.now()) => { const d = new Date(t); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
    const stamp = t => { const d = new Date(t); return `${pad(d.getHours())}:${pad(d.getMinutes())} ${pad(d.getDate())}/${pad(d.getMonth() + 1)}`; };
    const fmt = n => Number(n || 0).toLocaleString('vi-VN');
    const invariant = (ok, msg) => { if (!ok) throw new Error(msg); };

    /* ================= NHẬT KÝ ================= */
    let ui = null, running = null;
    let journal = load('journal', []).slice(-200);
    function log(msg, kind = 'info') {
        journal.push({ t: Date.now(), kind, msg: clean(msg) });
        if (journal.length > 200) journal.shift();
        save('journal', journal);
        renderLog();
    }
    function renderLog() {
        if (!ui) return;
        ui.querySelector('[data-log]').textContent = journal.slice(-120).reverse()
            .map(x => `${new Date(x.t).toLocaleTimeString('vi-VN')} ${x.kind === 'error' ? '✖' : x.kind === 'ok' ? '✔' : '•'} ${x.msg}`).join('\n') || '(trống)';
    }

    /* ================= BỘ LỌC PHẦN TỬ (như AutoBI) ================= */
    const own = e => !!e?.closest?.('[data-pmh-ui]');
    const visible = e => !!e?.isConnected && !own(e) && e.getClientRects().length > 0
        && getComputedStyle(e).visibility !== 'hidden' && getComputedStyle(e).display !== 'none';
    const all = (s, root = document) => [...root.querySelectorAll(s)].filter(visible);
    const txt = e => clean(e?.innerText || e?.textContent);

    /* ================= THEO DÕI DỮ LIỆU TRANG GỌI =================
       Chỉ ghi đường dẫn + dung lượng vào Nhật ký khi đang chạy, để biết trang lấy dữ liệu từ đâu. */
    function noteNet(url, body) {
        if (!running || !body || body.length < 50) return;
        running.netSeen = Date.now();
        let u = url;
        try { const x = new URL(url, location.href); u = x.host + x.pathname; } catch { /* giữ nguyên */ }
        const key = u + '|' + running.step;
        if (running.netLogged.has(key)) return;
        running.netLogged.add(key);
        log(`Trang tải dữ liệu: ${u.slice(0, 90)} (${Math.max(1, Math.round(body.length / 1024))} KB${/^\s*[[{]/.test(body) ? ', JSON' : ''})`);
    }
    function hookNetwork() {
        try {
            const of = W.fetch;
            if (of && !of.__pmh) {
                const nf = function (...a) {
                    const p = of.apply(this, a);
                    p.then(r => { try { const u = String(a[0]?.url || a[0]); r.clone().text().then(t => noteNet(u, t)).catch(() => {}); } catch { /* bỏ qua */ } }).catch(() => {});
                    return p;
                };
                nf.__pmh = true; W.fetch = nf;
            }
            const P = W.XMLHttpRequest?.prototype;
            if (P && !P.__pmh) {
                const XO = P.open, XS = P.send;
                P.open = function (m, u, ...r) { this.__pmhUrl = u; return XO.call(this, m, u, ...r); };
                P.send = function (...a) {
                    this.addEventListener('load', () => {
                        try { if (!this.responseType || this.responseType === 'text') noteNet(String(this.__pmhUrl), this.responseText); } catch { /* bỏ qua */ }
                    });
                    return XS.apply(this, a);
                };
                P.__pmh = true;
            }
        } catch { /* không theo dõi được thì thôi, không ảnh hưởng việc lấy phiếu */ }
    }

    /* ================= TÌM PHẦN TỬ TRÊN TRANG ================= */
    function findKhoInput() {
        const ins = all('input').filter(i => !['radio', 'checkbox', 'hidden', 'search', 'button', 'submit'].includes(i.type)
            && !/hãng|tìm|deal/iu.test(i.placeholder || ''));
        return ins.find(i => /^\d{1,6}$/.test(i.value)) || ins[0] || null;
    }
    const SEARCH_RE = /s[ăa]n\s*deal/iu;
    // Link thật dẫn sang trang khác (bấm vào là tải lại trang) → không dùng
    const isNavLink = e => {
        if (e.tagName !== 'A') return false;
        const h = (e.getAttribute('href') || '').trim();
        return !!h && !h.startsWith('#') && !/^javascript:/i.test(h);
    };
    // Ưu tiên nút nằm cùng khung với ô mã kho (khung "TÌM KHO XẢ DEAL")
    function findSearchButton() {
        const match = e => SEARCH_RE.test(txt(e) || e.value || '');
        const cands = all('button,[role=button],input[type=button],input[type=submit],a').filter(e => match(e) && !isNavLink(e));
        const input = findKhoInput();
        if (input) {
            for (let el = input.parentElement, lv = 0; el && lv < 8; el = el.parentElement, lv++) {
                const c = cands.find(b => el.contains(b));
                if (c) return c;
            }
        }
        if (cands.length) return cands[0];
        const d = all('div,span').filter(e => match(e) && txt(e).length < 30).sort((x, y) => txt(x).length - txt(y).length);
        return d[0] || null;
    }
    const desc = e => `<${e.tagName.toLowerCase()}${e.type ? ' type=' + e.type : ''}${e.id ? ' #' + e.id : ''}> "${(txt(e) || e.value || '').slice(0, 30)}"${e.form ? ' (nằm trong form)' : ''}`;
    function setValue(input, v) {
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
        input.focus();
        setter.call(input, v);
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
    }

    /* ================= ĐỌC DANH SÁCH PHIẾU ================= */
    const BADGE = /(ch[ỉi]\s*)?c[òo]n\s*:?\s*(\d[\d.,]*)\s*(xu[ấa]t)?/iu;
    const BAD_HEAD = /chọn voucher|nhập tên|tìm kho|săn deal|thông báo|sẵn sàng/iu;
    const stripIcon = s => clean(String(s).replace(/^[^\p{L}\p{N}]+/u, ''));

    // Khung 1 phiếu = khung nhỏ nhất chứa nút chọn + ô "CHỈ CÒN … Xuất" (không lấn sang phiếu khác)
    function radioBox(r) {
        let el = r.parentElement, first = null;
        for (let i = 0; el && i < 8; i++, el = el.parentElement) {
            if (el.querySelectorAll('input[type=radio]').length > 1) break;
            const t = txt(el);
            if (!first && t.length >= 8) first = el;
            if (BADGE.test(t)) return el;
        }
        return first;
    }
    // live = đang đọc trang thật (lọc phần tử ẩn); false = đọc bản HTML tải ngầm
    function itemBoxes(root = document, live = true) {
        const ok = e => !own(e) && (!live || visible(e));
        let boxes = [...root.querySelectorAll('input[type=radio]')].filter(r => !own(r)).map(radioBox).filter(b => b && ok(b));
        if (!boxes.length) {  // dự phòng: trang không dùng nút radio thật
            const isItem = e => { const t = txt(e); return t.length < 250 && BADGE.test(t) && clean(t.replace(BADGE, '')).length >= 8; };
            boxes = [...root.querySelectorAll('label,li,div')].filter(ok).filter(isItem).filter(e => ![...e.querySelectorAll('label,li,div')].some(c => isItem(c)));
        }
        return [...new Set(boxes)];
    }
    function findHeader(el) {
        let node = el;
        for (let lv = 0; node && node.tagName !== 'BODY' && lv < 6; lv++, node = node.parentElement) {
            for (let p = node.previousElementSibling; p; p = p.previousElementSibling) {
                if (own(p) || p.querySelector('input[type=radio]')) continue;
                const raw = txt(p);
                if (!raw || BADGE.test(raw)) continue;
                const t = stripIcon(raw);
                if (t && t.length <= 60 && !BAD_HEAD.test(t)) return t;
            }
        }
        return 'KHÁC';
    }
    function readItems(root = document, live = true) {
        return itemBoxes(root, live).map(box => {
            const t = txt(box), m = t.match(BADGE);
            const left = m ? parseInt(m[2].replace(/\D/g, ''), 10) : null;
            const name = clean(t.replace(BADGE, '')).replace(/^[•\-–\s]+/, '');
            return { group: findHeader(box), name, left: Number.isFinite(left) ? left : null };
        }).filter(x => x.name);
    }
    const sigOf = items => items.map(i => `${i.group}|${i.name}|${i.left}`).join('\n');
    const noDataShown = (root = document, live = true) => [...root.querySelectorAll('div,span,p')].filter(e => !own(e) && (!live || visible(e))).some(e => { const t = txt(e); return t.length < 120 && /không tìm thấy|không có (dữ liệu|deal|voucher|phiếu)/iu.test(t); });

    /* ================= LẤY NGẦM (không tải lại trang) =================
       Nút Săn deal là nút submit của form → trang tải lại. Script tự gửi đúng form đó trong nền,
       đọc HTML trả về, nên màn hình đứng yên. Không đọc được thì quay về cách bấm nút. */
    async function fetchKho(session, k) {
        const input = findKhoInput(), form = input?.form;
        if (!form || !input.name) return { skip: 'ô mã kho không nằm trong form có tên' };
        const btn = findSearchButton();
        let fd;
        try { fd = new FormData(form, btn && btn.form === form ? btn : undefined); }
        catch { fd = new FormData(form); if (btn?.name) fd.append(btn.name, btn.value || ''); }
        fd.set(input.name, k.code);
        const method = (form.getAttribute('method') || 'GET').toUpperCase();
        let url = form.action || location.href;
        const pairs = [...fd].filter(([, v]) => typeof v === 'string');
        const opt = { method, credentials: 'include' };
        if (method === 'GET') { const u = new URL(url, location.href); pairs.forEach(([a, b]) => u.searchParams.set(a, b)); url = u.href; }
        else opt.body = /multipart/i.test(form.enctype) ? fd : new URLSearchParams(pairs);
        const ctl = new AbortController();
        session.abort = () => ctl.abort();
        const timer = setTimeout(() => ctl.abort(), WAIT_MAX);
        try {
            const res = await fetch(url, { ...opt, signal: ctl.signal });
            check(session);
            invariant(res.ok, `HTTP ${res.status}`);
            const doc = new DOMParser().parseFromString(await res.text(), 'text/html');
            const items = readItems(doc, false);
            return { items, noData: !items.length && noDataShown(doc, false) };
        } catch (e) {
            check(session);
            throw e.name === 'AbortError' ? new Error(`quá ${WAIT_MAX / 1000} giây`) : e;
        } finally { clearTimeout(timer); session.abort = null; }
    }

    /* ================= PHIÊN CHẠY ================= */
    const cancelErr = () => Object.assign(new Error('Đã dừng'), { code: 'CANCELLED' });
    function check(s) { if (!s || s.cancelled || s !== running) throw cancelErr(); }
    async function sleep(s, ms) {
        const end = Date.now() + ms;
        while (Date.now() < end) { check(s); await new Promise(r => setTimeout(r, Math.min(100, end - Date.now()))); }
        check(s);
    }
    async function waitFor(s, fn, timeout) {
        const st = Date.now();
        while (Date.now() - st < timeout) { const v = fn(); if (v) return v; await sleep(s, 300); }
        return null;
    }
    async function waitList(s, oldSig, clickedAt) {
        const start = Date.now();
        let last = null, since = Date.now();
        await sleep(s, 800);
        while (Date.now() - start < WAIT_MAX) {
            check(s);
            const items = readItems(), sig = sigOf(items);
            if (sig !== last) { last = sig; since = Date.now(); }
            const elapsed = Date.now() - start, stable = Date.now() - since;
            const fresh = sig !== oldSig || s.netSeen > clickedAt || elapsed > 6000;
            if (items.length && fresh && stable > 1200) return items;
            if (!items.length && elapsed > 4000 && noDataShown()) return [];
            await sleep(s, 300);
        }
        throw new Error(`Hết ${WAIT_MAX / 1000} giây chưa thấy danh sách phiếu`);
    }

    /* ---------- Phiên chạy sống qua lần tải lại trang ----------
       Bấm "Săn deal" có thể làm trang tải lại → script khởi động lại từ đầu.
       Vì vậy tiến độ (đang ở kho nào, đã lấy được gì) được lưu lại; trang tải xong thì đọc tiếp. */
    const JOB_MAX_MS = 3 * 60000;
    const saveJob = j => save('job', j);
    const clearJob = () => save('job', null);

    function stop() {
        clearJob();
        if (!running) { status('Không có phiên nào đang chạy'); return; }
        running.cancelled = true;
        try { running.abort?.(); } catch { /* bỏ qua */ }
        status('Đang dừng…', 'warn');
    }

    async function run(source) {
        if (running) { log('Đang chạy rồi — bỏ qua lệnh mới'); return; }
        const job = { id: Date.now(), startedAt: Date.now(), idx: 0, results: [], phase: 'idle', resumes: 0 };
        saveJob(job);
        log(`Bắt đầu lấy phiếu (${source}) — ${KHO.map(k => `${k.code} ${k.nganh}`).join(', ')}`);
        await drive(job, false);
    }

    async function drive(job, resumed) {
        const session = { cancelled: false, netSeen: 0, netLogged: new Set(), step: '' };
        running = session; setBusy(true); openPanel(); progress(job.idx, KHO.length);
        try {
            invariant(await waitFor(session, findKhoInput, 15000), 'Không thấy ô "Nhập mã kho / siêu thị"');
            while (job.idx < KHO.length) {
                invariant(Date.now() - job.startedAt < JOB_MAX_MS, 'Phiên chạy quá 3 phút — đã hủy để tránh lặp');
                const k = KHO[job.idx];
                check(session); session.step = k.code;
                status(`Đang lấy kho ${k.code} (${k.nganh})… ${job.idx + 1}/${KHO.length}`); progress(job.idx, KHO.length);
                let oldSig = '', items;
                if (!resumed && job.mode !== 'click') {
                    try {
                        const r = await fetchKho(session, k);
                        if (r.skip) { log(`Không lấy ngầm được (${r.skip}) → dùng cách bấm nút`); job.mode = 'click'; }
                        else if (r.items.length && !r.items.some(i => i.left != null)) { log('Dữ liệu tải ngầm không có số xuất → dùng cách bấm nút'); job.mode = 'click'; }
                        else if (r.items.length || r.noData) { items = r.items; log(`Kho ${k.code}: lấy ngầm xong, không tải lại trang`); }
                        else { log('Dữ liệu tải ngầm không có danh sách phiếu (trang dựng bằng JS) → dùng cách bấm nút'); job.mode = 'click'; }
                    } catch (e) {
                        if (e.code === 'CANCELLED') throw e;
                        log(`Lấy ngầm lỗi (${e.message}) → dùng cách bấm nút`); job.mode = 'click';
                    }
                    saveJob(job);
                }
                if (items) { /* đã có từ lấy ngầm */ }
                else if (resumed) {
                    resumed = false;
                    const inp = findKhoInput();
                    log(`Trang đã TẢI LẠI sau khi bấm Săn deal · địa chỉ: ${location.pathname}${location.search}${location.hash} · ô mã kho đang là "${inp ? inp.value : '?'}" · đọc tiếp kho ${k.code}`);
                } else {
                    oldSig = sigOf(readItems());
                    const input = findKhoInput();
                    invariant(input, 'Mất ô nhập mã kho');
                    setValue(input, k.code);
                    await sleep(session, 300);
                    invariant(clean(input.value) === k.code, `Không nhập được mã kho ${k.code} (ô đang là "${input.value}")`);
                    const b = findSearchButton();
                    invariant(b, 'Không thấy nút "SĂN DEAL NGAY"');
                    job.phase = 'submitted'; job.clickedAt = Date.now(); saveJob(job);   // lưu trước khi bấm
                    log(`Kho ${k.code}: nhập mã xong, bấm ${desc(b)}`);
                    b.click();
                }
                let rec;
                try {
                    if (!items) items = await waitList(session, oldSig, job.clickedAt);
                    rec = { ...k, ok: true, items };
                    log(`Kho ${k.code} (${k.nganh}): ${items.length} phiếu, ${new Set(items.map(x => x.group)).size} hãng/nhóm`, 'ok');
                } catch (e) {
                    if (e.code === 'CANCELLED') throw e;
                    rec = { ...k, ok: false, items: [], err: e.message };
                    log(`Kho ${k.code}: ${e.message}`, 'error');
                }
                job.results.push(rec); job.idx++; job.phase = 'idle'; saveJob(job);
            }
            clearJob(); progress(1, 1);
            const result = { time: Date.now(), day: dayKey(), kho: job.results };
            view.data = result; save('last', result); view.groups.clear();
            const bad = result.kho.filter(x => !x.ok).length;
            status(bad ? `Xong, ${bad} kho lỗi — xem Nhật ký` : `Hoàn tất lúc ${stamp(result.time)}`, bad ? 'warn' : 'ok');
            log(bad ? `Kết thúc, ${bad} kho lỗi` : 'Hoàn tất', bad ? 'error' : 'ok');
            renderAll();
        } catch (e) {
            clearJob();
            if (e.code === 'CANCELLED') { status('Đã dừng — bỏ phần đang lấy dở', 'warn'); log('Đã dừng theo yêu cầu'); }
            else { status(e.message, 'err'); log(e.message, 'error'); }
        } finally {
            running = null; setBusy(false);
        }
    }

    // Khi trang mở: nếu có phiên dở dang vừa bấm Săn deal thì đọc tiếp; không thì xét tự chạy
    function resumeOrAuto() {
        const job = load('job', null);
        if (job && job.phase === 'submitted') {
            const okTime = Date.now() - (job.clickedAt || 0) < 60000 && Date.now() - job.startedAt < JOB_MAX_MS;
            job.resumes = (job.resumes || 0) + 1;
            if (okTime && job.resumes <= KHO.length * 2) { saveJob(job); drive(job, true); return; }
            clearJob(); log('Bỏ phiên dở dang cũ (quá giờ hoặc tải lại quá nhiều lần)', 'error');
        } else if (job) clearJob();
        const last = load('last', null);
        if (last) view.data = last;
        maybeAutoRun(last);
    }

    function maybeAutoRun(last) {
        if (!AUTO_RUN) return;
        if (last && Date.now() - last.time < FRESH_MIN * 60000 && last.kho?.every(k => k.ok)) {
            log(`Đã có kết quả lúc ${stamp(last.time)} — mở bảng cũ, bấm "Lấy phiếu" nếu muốn cập nhật`);
            openPanel(); renderAll(); return;
        }
        if (Date.now() - load('lastAuto', 0) < 3 * 60000) { log('Không tự chạy vì vừa tự chạy dưới 3 phút trước (chống lặp)'); return; }
        save('lastAuto', Date.now());
        run('tự động khi mở trang');
    }

    /* ================= DỮ LIỆU CHO BẢNG ================= */
    const view = { data: null, nganh: 'all', groups: new Set(), q: '' };
    const allRows = () => (view.data?.kho || []).flatMap(k => k.items.map(it => ({ nganh: k.nganh, kho: k.code, ...it })));
    function rowsByNganh() { return allRows().filter(r => view.nganh === 'all' || r.nganh === view.nganh); }
    function filteredRows() {
        const q = norm(view.q);
        return rowsByNganh().filter(r => (!view.groups.size || view.groups.has(r.nganh + '|' + r.group))
            && (!q || norm(r.group + ' ' + r.name).includes(q)));
    }
    const lvl = n => n == null ? '' : n < 20 ? 'low' : n < 50 ? 'mid' : 'ok';

    /* ================= GIAO DIỆN ================= */
    const CSS = `
#pmh-launch{position:fixed;right:20px;bottom:20px;z-index:2147483645;background:#087f8c;color:#fff;padding:12px 20px;border:0;border-radius:24px;cursor:pointer;font:600 14px Arial,sans-serif;box-shadow:0 4px 14px #0003}
#pmh-launch.busy{background:#b45309}
#pmh-back{position:fixed;inset:0;z-index:2147483646;background:#0f172a66;display:flex;align-items:center;justify-content:center}
#pmh-back[hidden]{display:none}
#pmh-panel{box-sizing:border-box;width:min(1150px,96vw);height:92vh;display:flex;flex-direction:column;background:#fff;color:#172a3a;border-radius:14px;box-shadow:0 20px 60px #0005;font:14px/1.45 Arial,sans-serif;overflow:hidden}
#pmh-panel *{box-sizing:border-box}
#pmh-panel .top{flex:none;display:flex;flex-wrap:wrap;align-items:center;gap:8px;padding:12px 18px;border-bottom:1px solid #e3e9ed}
#pmh-panel .top strong{font-size:17px}#pmh-panel .top .sp{flex:1}#pmh-panel .ver{font-size:12px;color:#7a8a96}
#pmh-panel .body{flex:1;overflow-y:auto;padding:14px 18px}
#pmh-panel button{min-height:36px;padding:7px 14px;border:1px solid #bac9d1;border-radius:8px;background:#f2f7f8;color:#173047;cursor:pointer;font:inherit}
#pmh-panel button.primary{background:#087f8c;color:#fff;border-color:#087f8c;font-weight:600}
#pmh-panel button.danger{background:#b91c1c;color:#fff;border-color:#b91c1c;font-weight:600}
#pmh-panel .busy-only{display:none}#pmh-panel.busy .busy-only{display:inline-block}#pmh-panel.busy .idle-only{display:none}
#pmh-panel input{min-height:36px;padding:6px 10px;border:1px solid #b8c9ce;border-radius:8px;font:inherit;min-width:240px;flex:1;max-width:420px}
#pmh-panel .st{padding:8px 12px;border-radius:8px;background:#f2f7f8;margin-bottom:6px}
#pmh-panel .st.ok{background:#e6f4ea;color:#14532d}#pmh-panel .st.warn{background:#fff4cf;color:#5c3800}#pmh-panel .st.err{background:#fdecea;color:#8a1c1c}
#pmh-panel .prog{height:4px;background:#e3e9ed;border-radius:2px;overflow:hidden;margin-bottom:10px}#pmh-panel [data-bar]{height:4px;width:0;background:#087f8c;transition:width .3s}
#pmh-panel .muted{color:#6b7c88;font-size:12px}
#pmh-panel .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin:10px 0}
#pmh-panel .kpis>div{background:#f2f6f7;border-radius:10px;padding:10px 14px;display:flex;flex-direction:column}
#pmh-panel .kpis>div.main{background:#087f8c;color:#fff}#pmh-panel .kpis b{font-size:22px}#pmh-panel .kpis small{font-size:12px}
#pmh-panel .kpis>div.ict{background:#e8f0fe}#pmh-panel .kpis>div.ce{background:#fdeef3}
#pmh-panel .comp{border:1px solid #e3e9ed;border-radius:10px;padding:10px 14px;margin-bottom:10px;display:flex;flex-direction:column;gap:6px}
#pmh-panel .filters{border:1px solid #e3e9ed;border-radius:10px;padding:10px 12px;display:flex;flex-direction:column;gap:8px;margin-bottom:10px}
#pmh-panel .group{display:flex;flex-wrap:wrap;align-items:center;gap:6px}
#pmh-panel .group>b{font-size:12px;color:#4a5a66;min-width:80px}
#pmh-panel .chip{min-height:30px;padding:4px 12px;border-radius:16px;background:#fff}
#pmh-panel .chip.on{background:#087f8c;color:#fff;border-color:#087f8c}
#pmh-panel table{width:100%;border-collapse:collapse;font-size:14px}
#pmh-panel th{position:sticky;top:-14px;background:#173047;color:#fff;text-align:left;padding:9px 10px;font-weight:600}
#pmh-panel td{padding:8px 10px;border-bottom:1px solid #e8eef1;vertical-align:middle}
#pmh-panel tr.gstart td{border-top:2px solid #cfdde3}
#pmh-panel td.n,#pmh-panel th.n{text-align:right;white-space:nowrap}
#pmh-panel td.grp{font-weight:700;color:#173047;background:#fafcfd}#pmh-panel td small{display:block;font-weight:400;color:#7a8a96;font-size:11px}
#pmh-panel td.ng{text-align:center;background:#fafcfd;vertical-align:top;padding-top:12px}
#pmh-panel td.grp{vertical-align:top;padding-top:10px}
#pmh-panel .badge{display:inline-block;padding:3px 10px;border-radius:12px;font-weight:700;color:#fff;background:#64748b}
#pmh-panel .badge.ict{background:#1d4ed8}#pmh-panel .badge.ce{background:#be185d}
#pmh-panel .left{display:inline-block;min-width:54px;text-align:center;padding:3px 8px;border-radius:8px;font-weight:700;background:#e6f4ea;color:#14532d}
#pmh-panel .left.mid{background:#fff4cf;color:#5c3800}#pmh-panel .left.low{background:#fdecea;color:#8a1c1c}
#pmh-panel .warnbox{padding:10px;background:#fff4cf;border-radius:8px}
#pmh-panel details{margin-top:14px;border:1px solid #e3e9ed;border-radius:10px;padding:8px 12px}
#pmh-panel summary{cursor:pointer;font-weight:600}
#pmh-panel pre{margin:8px 0 0;max-height:220px;overflow:auto;background:#0f172a;color:#d6e2ea;padding:10px;border-radius:8px;font:12px/1.5 Consolas,monospace;white-space:pre-wrap}
@media (max-width:700px){#pmh-panel{width:100vw;height:100vh;border-radius:0}#pmh-panel table{font-size:13px}#pmh-panel td,#pmh-panel th{padding:6px}}
`;

    function buildUI() {
        if (ui) return;
        const st = document.createElement('style'); st.textContent = CSS; st.setAttribute('data-pmh-ui', '');
        document.head.appendChild(st);

        const launch = document.createElement('button');
        launch.id = 'pmh-launch'; launch.setAttribute('data-pmh-ui', ''); launch.textContent = '🎟️ PMH hôm nay';
        launch.onclick = () => { openPanel(); renderAll(); };
        document.body.appendChild(launch);

        ui = document.createElement('div');
        ui.id = 'pmh-back'; ui.setAttribute('data-pmh-ui', ''); ui.hidden = true;
        ui.innerHTML = `
<div id="pmh-panel">
  <div class="top">
    <strong>🎟️ Phiếu mua hàng HNO+</strong><span class="ver">v${VERSION}</span><span class="sp"></span>
    <button class="primary idle-only" data-act="run">▶ Lấy phiếu</button>
    <button class="danger busy-only" data-act="stop">■ Dừng</button>
    <button data-act="copy">📋 Copy Line</button>
    <button data-act="close">✖ Đóng</button>
  </div>
  <div class="body">
    <div class="st" data-status>Sẵn sàng</div>
    <div class="prog"><div data-bar></div></div>
    <div class="muted" data-meta></div>
    <div class="kpis" data-kpis></div>
    <div class="comp" data-comp></div>
    <div class="filters">
      <div class="group" data-f-nganh></div>
      <div class="group" data-f-group></div>
      <div class="group"><b>Tìm</b><input type="search" data-q placeholder="Tìm hãng, mức giảm, sản phẩm…"><button data-act="clear">Bỏ lọc</button></div>
    </div>
    <div data-table></div>
    <details><summary>📝 Nhật ký (bấm để mở)</summary><pre data-log></pre>
      <div style="margin-top:8px;display:flex;gap:8px"><button data-act="copylog">📋 Sao chép nhật ký</button><button data-act="clearlog">Xóa nhật ký</button></div></details>
  </div>
</div>`;
        document.body.appendChild(ui);

        ui.addEventListener('click', e => {
            const a = e.target.closest('[data-act]')?.dataset.act;
            if (a === 'run') run('bấm tay');
            else if (a === 'stop') stop();
            else if (a === 'copy') copyZalo();
            else if (a === 'close') { if (running) status('Đang chạy — bấm Dừng trước khi đóng', 'warn'); else ui.hidden = true; }
            else if (a === 'clear') { view.nganh = 'all'; view.groups.clear(); view.q = ''; ui.querySelector('[data-q]').value = ''; renderAll(); }
            else if (a === 'copylog') copyText(ui.querySelector('[data-log]').textContent, 'Đã sao chép nhật ký');
            else if (a === 'clearlog') { journal = []; save('journal', journal); renderLog(); }
        });
        ui.querySelector('[data-q]').addEventListener('input', e => { view.q = e.target.value; renderTable(); });
        renderLog();
    }

    function openPanel() { buildUI(); ui.hidden = false; }
    function status(msg, kind = '') { if (!ui) return; const s = ui.querySelector('[data-status]'); s.textContent = msg; s.className = 'st ' + kind; }
    function progress(done, total) { if (ui) ui.querySelector('[data-bar]').style.width = (total ? Math.round(done / total * 100) : 0) + '%'; }
    function setBusy(b) {
        if (!ui) return;
        ui.querySelector('#pmh-panel').classList.toggle('busy', b);
        const l = document.getElementById('pmh-launch');
        if (l) { l.classList.toggle('busy', b); l.textContent = b ? '⏳ Đang lấy PMH…' : '🎟️ PMH hôm nay'; }
    }

    function renderAll() { if (!ui) return; renderSummary(); renderFilters(); renderTable(); }

    function renderSummary() {
        const d = view.data, meta = ui.querySelector('[data-meta]'), kp = ui.querySelector('[data-kpis]'), cp = ui.querySelector('[data-comp]');
        if (!d) { meta.textContent = ''; kp.innerHTML = ''; cp.innerHTML = ''; return; }
        const old = d.day !== dayKey() ? ' ⚠ Dữ liệu của ngày khác, nên bấm "Lấy phiếu" lại.' : '';
        meta.textContent = `Lấy lúc ${stamp(d.time)} · ` + d.kho.map(k => `Kho ${k.code} ${k.nganh} ${k.ok ? '✔' : '✖ ' + (k.err || 'lỗi')}`).join(' · ') + old;
        const rows = allRows();
        const totalLeft = rows.reduce((a, r) => a + (r.left || 0), 0);
        kp.innerHTML = `<div class="main"><small>Tổng phiếu</small><b>${fmt(rows.length)}</b><small>Còn ${fmt(totalLeft)} xuất</small></div>` +
            d.kho.map(k => {
                const g = new Set(k.items.map(i => i.group)).size, left = k.items.reduce((a, i) => a + (i.left || 0), 0);
                return `<div class="${esc(k.nganh.toLowerCase())}"><small>${esc(k.nganh)} · kho ${esc(k.code)}</small><b>${fmt(k.items.length)} phiếu</b><small>${g} hãng/nhóm · còn ${fmt(left)} xuất</small></div>`;
            }).join('');
        cp.innerHTML = d.kho.map(k => {
            const m = new Map();
            k.items.forEach(i => m.set(i.group, (m.get(i.group) || 0) + 1));
            const list = [...m].map(([g, n]) => `${esc(g)} (${n})`).join(', ') || (k.ok ? 'không có phiếu' : 'lỗi khi lấy');
            return `<div><span class="badge ${esc(k.nganh.toLowerCase())}">${esc(k.nganh)}</span> <b>Kho ${esc(k.code)}</b> gồm: ${list}</div>`;
        }).join('');
    }

    function renderFilters() {
        const fn = ui.querySelector('[data-f-nganh]'), fg = ui.querySelector('[data-f-group]');
        const ngs = [...new Set(allRows().map(r => r.nganh))];
        fn.innerHTML = '<b>Ngành</b>' + ['all', ...ngs].map(n =>
            `<button class="chip ${view.nganh === n ? 'on' : ''}" data-ng="${esc(n)}">${n === 'all' ? 'Tất cả' : esc(n)}</button>`).join('');
        fn.querySelectorAll('[data-ng]').forEach(b => b.onclick = () => {
            view.nganh = b.dataset.ng;
            for (const k of [...view.groups]) if (view.nganh !== 'all' && !k.startsWith(view.nganh + '|')) view.groups.delete(k);
            renderFilters(); renderTable();
        });
        const groups = [];
        rowsByNganh().forEach(r => { const k = r.nganh + '|' + r.group; if (!groups.some(g => g.k === k)) groups.push({ k, g: r.group, n: r.nganh }); });
        fg.innerHTML = '<b>Hãng / Nhóm</b>' + (groups.length ? groups.map(g =>
            `<button class="chip ${view.groups.has(g.k) ? 'on' : ''}" data-g="${esc(g.k)}" title="${esc(g.n)}">${esc(g.g)}</button>`).join('') : '<span class="muted">—</span>');
        fg.querySelectorAll('[data-g]').forEach(b => b.onclick = () => {
            const k = b.dataset.g; view.groups.has(k) ? view.groups.delete(k) : view.groups.add(k);
            renderFilters(); renderTable();
        });
    }

    function renderTable() {
        const box = ui.querySelector('[data-table]');
        if (!view.data) { box.innerHTML = '<div class="muted">Chưa có dữ liệu. Bấm ▶ Lấy phiếu.</div>'; return; }
        const rows = filteredRows();
        if (!rows.length) { box.innerHTML = '<div class="warnbox">Không có phiếu khớp bộ lọc.</div>'; return; }
        const same = (a, b, byGroup) => a && b && a.nganh === b.nganh && (!byGroup || a.group === b.group);
        let h = '<table><thead><tr><th class="n">#</th><th>Ngành</th><th>Hãng / Nhóm</th><th>Phiếu mua hàng</th><th class="n">Còn (xuất)</th></tr></thead><tbody>';
        rows.forEach((r, i) => {
            const firstN = !same(rows[i - 1], r, false), firstG = !same(rows[i - 1], r, true);
            let spanN = 0, spanG = 0;
            if (firstN) for (let j = i; j < rows.length && same(rows[j], r, false); j++) spanN++;
            if (firstG) for (let j = i; j < rows.length && same(rows[j], r, true); j++) spanG++;
            const cls = esc(r.nganh.toLowerCase());
            h += `<tr class="${firstG ? 'gstart' : ''}"><td class="n muted">${i + 1}</td>`
                + (firstN ? `<td rowspan="${spanN}" class="ng"><span class="badge ${cls}">${esc(r.nganh)}</span><small>Kho ${esc(r.kho)}</small></td>` : '')
                + (firstG ? `<td rowspan="${spanG}" class="grp">${esc(r.group)}<small>${spanG} phiếu</small></td>` : '')
                + `<td>${esc(r.name)}</td><td class="n"><span class="left ${lvl(r.left)}">${r.left == null ? '—' : fmt(r.left)}</span></td></tr>`;
        });
        h += '</tbody></table>';
        box.innerHTML = `<div class="muted" style="margin-bottom:6px">Đang xem ${rows.length}/${allRows().length} phiếu</div>` + h;
    }

    async function copyText(text, okMsg) {
        try {
            if (typeof GM_setClipboard === 'function') GM_setClipboard(text, 'text');
            else await navigator.clipboard.writeText(text);
        } catch {
            const ta = document.createElement('textarea'); ta.value = text; ta.setAttribute('data-pmh-ui', '');
            document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove();
        }
        status(okMsg, 'ok');
    }

    async function copyZalo() {
        const rows = filteredRows();
        if (!rows.length) { status('Không có phiếu để copy', 'warn'); return; }
        const d = new Date(view.data.time);
        let msg = `📢 PHIẾU MUA HÀNG HÔM NAY ${pad(d.getDate())}/${pad(d.getMonth() + 1)}\n`;
        let ng = null, gr = null;
        rows.forEach(r => {
            if (r.nganh !== ng) { ng = r.nganh; gr = null; msg += `\n━━━━━━━━━━\n🏬 ${r.nganh} (kho ${r.kho})\n`; }
            if (r.group !== gr) { gr = r.group; msg += `\n⚡ ${r.group}\n`; }
            msg += `• ${r.name}\n`;
        });
        msg += `\n👉 Tổng ${rows.length} phiếu. Nhớ tư vấn KH dùng phiếu phù hợp!`;
        await copyText(msg, `Đã copy ${rows.length} phiếu`);
    }

    /* ================= KHỞI ĐỘNG ================= */
    hookNetwork();
    const start = () => { buildUI(); setTimeout(resumeOrAuto, 1200); };
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
    else start();
})();
