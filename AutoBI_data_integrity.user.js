/* AutoBI 16.1.1.53 - đồng bộ trạng thái và bảo vệ dữ liệu báo cáo. */
(function () {
    'use strict';

    const CACHE_KEY = 'tgdd_data_cache_v30';
    const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
    const norm = value => String(value || '')
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/gi, 'd').toLowerCase().replace(/[^a-z0-9]/g, '');
    const text = element => String(element?.innerText || element?.textContent || '').trim();
    const visible = element => !!element && element.getClientRects().length > 0;

    document.addEventListener('click', event => {
        const button = event.target && event.target.closest && event.target.closest('button');
        const label = norm(text(button));
        if (label === 'realtime' || label === 'luyke') {
            window.__AutoBIDataMode53 = label;
        }
    }, true);

    function loading() {
        return Array.from(document.querySelectorAll(
            '.ant-spin-spinning,.loading,.spinner,[aria-busy="true"],.animate-spin'
        )).some(visible);
    }

    function reportSignature() {
        return Array.from(document.querySelectorAll('.report-table table,.ant-table-content table'))
            .slice(0, 12)
            .map(table => Array.from(table.querySelectorAll('thead tr,tbody tr'))
                .filter(row => !row.classList.contains('ant-table-measure-row'))
                .slice(0, 8).map(text).join('|'))
            .join('###');
    }

    function modeButton(mode) {
        const wanted = norm(mode);
        return Array.from(document.querySelectorAll('button'))
            .find(button => visible(button) && norm(text(button)) === wanted) || null;
    }

    function buttonLooksActive(button) {
        if (!button) return false;
        const classes = norm(button.className);
        const state = norm(button.getAttribute('data-state'));
        return button.getAttribute('aria-pressed') === 'true' ||
            state === 'active' || state === 'on' ||
            classes.includes('active') || classes.includes('selected') ||
            classes.includes('bgblue') || classes.includes('bgprimary');
    }

    async function selectMode(mode, options = {}) {
        const requireChange = options.requireChange !== false;
        for (let attempt = 1; attempt <= 3; attempt++) {
            const button = modeButton(mode);
            if (!button) throw new Error('Không tìm thấy nút ' + mode);

            const before = reportSignature();
            button.click();
            const started = Date.now();
            let stable = 0;
            let previous = '';
            let changed = false;

            while (Date.now() - started < 22000) {
                await sleep(350);
                const signature = reportSignature();
                if (signature && signature !== before) changed = true;
                if (signature && signature === previous && !loading()) stable++;
                else stable = 0;
                previous = signature;

                const confirmed = buttonLooksActive(button) || changed || !requireChange;
                if (confirmed && signature && !loading() && stable >= 4 &&
                    (!requireChange || changed || Date.now() - started > 4500)) {
                    window.__AutoBIDataMode53 = norm(mode);
                    return;
                }
            }
            console.warn('[AutoBI 53] Chuyển chế độ chưa ổn định, thử lại:', mode, attempt);
        }
        throw new Error('Bảng ' + mode + ' không tải ổn định');
    }

    function clone(value) {
        return JSON.parse(JSON.stringify(value || {}));
    }

    function dataFingerprint(data) {
        return ['total', 'shop1', 'shop2', 'shop3', 'shop4', 'shop5']
            .map(key => {
                const row = data?.[key] || {};
                return [key, row.r, row.dt, row.t, row.dk, row.dtqd_dk]
                    .map(value => String(value ?? '')).join(':');
            }).join('|');
    }

    function usableRevenue(data) {
        if (!data || typeof data !== 'object') return false;
        return ['total', 'shop1', 'shop2', 'shop3', 'shop4', 'shop5']
            .some(key => {
                const row = data[key];
                return row && Object.values(row).some(value => Number.isFinite(Number(value)) && Number(value) !== 0);
            });
    }

    function healthScore(data) {
        if (!data || typeof data !== 'object') return 0;
        let populated = 0;
        let anchors = 0;
        Object.values(data).forEach(row => {
            if (row && (Number(row.sl) !== 0 || Number(row.dtqd) !== 0)) populated++;
        });
        ['Smartphone', 'Laptop', 'Phụ kiện', 'SIM', 'Điện tử', 'Điện lạnh',
            'Máy giặt', 'Máy lạnh', 'Điện gia dụng', 'Máy lọc nước']
            .forEach(name => {
                const row = data[name];
                if (row && (Number(row.sl) !== 0 || Number(row.dtqd) !== 0)) anchors++;
            });
        return populated * 10 + anchors * 25;
    }

    async function scrapeBestHealth(DATA, UI, label) {
        let best = null;
        let bestScore = -1;
        let repeatedBest = 0;
        for (let attempt = 1; attempt <= 3; attempt++) {
            const candidate = await callbackAsPromise(cb => DATA.scrapeHealthCategoriesDMX(cb), 120000);
            const score = healthScore(candidate);
            if (score > bestScore) {
                best = candidate;
                bestScore = score;
                repeatedBest = 0;
            } else if (score === bestScore && score > 0) {
                repeatedBest++;
            }
            if (attempt >= 2 && repeatedBest >= 1 && bestScore > 0) break;
            UI.showToast('⏳ ' + label + ': đang kiểm tra đủ nhóm ngành hàng (' + attempt + '/3)...', 0);
            await sleep(1600);
        }
        if (!best || bestScore <= 0) throw new Error(label + ' không có dữ liệu ngành hàng');
        console.info('[AutoBI 53] Chọn bảng ngành hàng tốt nhất', label, 'điểm', bestScore);
        return best;
    }

    function readTlpvtc() {
        const nodes = Array.from(document.querySelectorAll('span,div'));
        for (const node of nodes) {
            const label = norm(text(node));
            if (!label.includes('tlpvtc')) continue;
            const card = node.closest('.flex-col,.rounded-xl,div[class*="rounded"]');
            if (!card) continue;
            const values = Array.from(card.querySelectorAll('.tabular-nums,div.text-lg,div[class*="text-lg"],span'));
            for (const item of values) {
                const match = text(item).match(/(-?\d+(?:[.,]\d+)?)\s*%/);
                if (match) return Number(match[1].replace(',', '.')) || 0;
            }
        }
        return 0;
    }

    function callbackAsPromise(invoke, timeout = 30000) {
        return new Promise((resolve, reject) => {
            let done = false;
            const timer = setTimeout(() => {
                if (!done) {
                    done = true;
                    reject(new Error('Hết thời gian chờ thao tác BI'));
                }
            }, timeout);
            try {
                invoke(result => {
                    if (done) return;
                    done = true;
                    clearTimeout(timer);
                    resolve(result);
                });
            } catch (error) {
                clearTimeout(timer);
                reject(error);
            }
        });
    }

    function saveSection(name, value) {
        const cache = GM_getValue(CACHE_KEY, {}) || {};
        cache[name] = clone(value);
        GM_setValue(CACHE_KEY, cache);
    }

    async function robustRevenueRun(DATA, UI, config, done) {
        try {
            UI.showToast('🛡️ Đang đồng bộ bộ lọc và trạng thái báo cáo...', 0);
            await callbackAsPromise(cb => DATA.runFilterAllSequenceDMX(cb), 45000);
            await callbackAsPromise(cb => DATA.ensureRevenueOptionsDMX(cb), 30000);

            UI.showToast('⚡ Đang quét Doanh thu Realtime...', 0);
            await selectMode('Realtime');
            let realtime = DATA.scrapeRevenueTableDMX(config, true);
            if (!usableRevenue(realtime)) {
                await selectMode('Realtime', { requireChange: false });
                realtime = DATA.scrapeRevenueTableDMX(config, true);
            }
            if (!usableRevenue(realtime)) throw new Error('Dữ liệu Realtime rỗng');
            saveSection('link1', realtime);

            UI.showToast('📊 Đang quét Doanh thu Lũy kế...', 0);
            await selectMode('Lũy kế');
            let cumulative = DATA.scrapeRevenueTableDMX(config, false);
            if (!usableRevenue(cumulative) || dataFingerprint(cumulative) === dataFingerprint(realtime)) {
                await sleep(1200);
                await selectMode('Lũy kế', { requireChange: false });
                cumulative = DATA.scrapeRevenueTableDMX(config, false);
            }
            if (!usableRevenue(cumulative)) throw new Error('Dữ liệu Lũy kế rỗng');
            if (dataFingerprint(cumulative) === dataFingerprint(realtime)) {
                throw new Error('Realtime và Lũy kế trùng hoàn toàn; đã chặn lưu nhầm');
            }

            const tlpvtc = readTlpvtc();
            if (cumulative.total) cumulative.total.tlpvtc = tlpvtc;
            ['shop1', 'shop2', 'shop3', 'shop4', 'shop5'].forEach(key => {
                if (cumulative[key]) cumulative[key].tlpvtc = tlpvtc;
            });
            saveSection('link2', cumulative);

            if (window.__AutoBIFastRealtime43?.enabled()) {
                await window.__AutoBIFastRealtime43.finishRevenue(config, done, UI);
                return;
            }

            UI.showToast('👤 Đang mở dữ liệu nhân viên...', 0);
            const staffButton = Array.from(document.querySelectorAll('button'))
                .find(button => visible(button) && (norm(text(button)) === 'nhanvien' || button.querySelector('.lucide-users')));
            if (staffButton) staffButton.click();
            await sleep(700);
            await new Promise(resolve => window.__AutoBIStableWaitDetail(resolve, 2500));

            const staff = DATA.scrapeStaffRevenueTableDMX(config) || {};
            const cache = GM_getValue(CACHE_KEY, {}) || {};
            if (!cache.link6) cache.link6 = {};
            if (!cache.link7) cache.link7 = {};
            Object.keys(staff.staffRevMap || {}).forEach(key => {
                if (!cache.link6[key]) cache.link6[key] = { revenue: {}, competition: {}, crossSell: {} };
                cache.link6[key].revenue = staff.staffRevMap[key];
                cache.link7[key] = (staff.staffInstMap || {})[key];
            });
            GM_setValue(CACHE_KEY, cache);

            const totalHealth = await scrapeBestHealth(DATA, UI, 'Tổng cụm');
            const afterHealth = GM_getValue(CACHE_KEY, {}) || {};
            if (!afterHealth.link8_health) afterHealth.link8_health = {};
            afterHealth.link8_health.total = totalHealth || {};
            GM_setValue(CACHE_KEY, afterHealth);

            await callbackAsPromise(cb => window.__AutoBIHealthPerShop.scan(
                config, DATA.scrapeHealthCategoriesDMX,
                (message, ms) => UI.showToast(message, ms), cb
            ), 420000);

            UI.showToast('✅ Hoàn tất và đã kiểm tra dữ liệu Doanh thu!', 3000);
            if (done) done();
        } catch (error) {
            console.error('[AutoBI 53 Revenue]', error);
            UI.showToast('❌ Đã chặn dữ liệu sai: ' + error.message + '. Hãy chạy lại.', 12000);
            if (done) done();
        }
    }

    function strengthenWaits() {
        window.__AutoBIStableWait = callback => {
            const mode = window.__AutoBIDataMode53;
            if (mode === 'realtime') selectMode('Realtime', { requireChange: false }).then(callback);
            else if (mode === 'luyke') selectMode('Lũy kế', { requireChange: false }).then(callback);
            else window.__AutoBIStable.wait(callback, { minWait: 1200, maxWait: 18000, poll: 300, requireChange: true });
        };
        window.__AutoBIStableWaitDetail = callback => {
            window.__AutoBIStable.wait(callback, { minWait: 1400, maxWait: 22000, poll: 300, requireChange: true });
        };
    }

    function install() {
        const page = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
        const DATA = page.DATA || window.DATA;
        const UI = page.UI || window.UI;
        if (!DATA || !UI || DATA.__integrity53) return false;
        DATA.__integrity53 = true;
        strengthenWaits();
        DATA.runRevenueSequenceDMX = (config, done) => robustRevenueRun(DATA, UI, config, done);
        console.info('[AutoBI 16.1.1.53] Data integrity guard ready');
        return true;
    }

    if (!install()) {
        const timer = setInterval(() => {
            if (install()) clearInterval(timer);
        }, 100);
        setTimeout(() => clearInterval(timer), 30000);
    }
})();
