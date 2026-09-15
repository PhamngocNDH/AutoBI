/* AutoBI 16.1.1.53 - quét nhiều lần và giữ dữ liệu Sức khỏe ST đầy đủ nhất. */
(function () {
    'use strict';

    const DATA_KEY = 'tgdd_data_cache_v30';
    const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
    const slug = value => String(value || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .replace(/đ/g, 'd')
        .replace(/[^a-z0-9]/g, '');

    function visible(element) {
        return !!element && element.getClientRects().length > 0;
    }

    function currentShopLabel() {
        const button = Array.from(document.querySelectorAll('button.filter-field-card, button'))
            .find(item => visible(item) && Array.from(item.querySelectorAll('span'))
                .some(span => slug(span.innerText || span.textContent) === 'sieuthi'));
        return slug(button && (button.innerText || button.textContent));
    }

    function expectedShopTokens(config, index) {
        return [
            config['makho' + index],
            config['shopCode' + index],
            config['code' + index],
            config['shop' + index],
            config['shop' + index + 'Short']
        ].map(slug).filter(Boolean);
    }

    function isExpectedShop(config, index) {
        const label = currentShopLabel();
        return !!label && expectedShopTokens(config, index).some(token => label.includes(token));
    }

    function hasHealthData(data) {
        if (!data || typeof data !== 'object') return false;
        return Object.values(data).some(row => row && (
            Number(row.sl) !== 0 || Number(row.dtqd) !== 0 ||
            Number(row.growth) !== 0 || Number(row.tg_ratio) !== 0
        ));
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

    function scrapeOnce(scrapeFn, timeout = 90000) {
        return new Promise(resolve => {
            let settled = false;
            const timer = setTimeout(() => {
                if (!settled) {
                    settled = true;
                    resolve(null);
                }
            }, timeout);

            try {
                scrapeFn(data => {
                    if (settled) return;
                    settled = true;
                    clearTimeout(timer);
                    resolve(data || null);
                });
            } catch (error) {
                clearTimeout(timer);
                settled = true;
                console.error('[AutoBI Health 53] Lỗi đọc bảng:', error);
                resolve(null);
            }
        });
    }

    function restoreRevenueRows() {
        const cache = GM_getValue(DATA_KEY, {}) || {};
        if (!cache.link2) return;

        let changed = false;
        for (let index = 1; index <= 5; index++) {
            const row = cache.link2['shop' + index];
            if (!row || !Number.isFinite(row.grossQD)) continue;

            const current = Number(row.r);
            const gross = Number(row.grossQD);
            if (Number.isFinite(current) && current !== 0 && Number.isFinite(row.dtqd_dk)) {
                row.dtqd_dk *= gross / current;
            }
            row.r = gross;
            if (Number(row.t) > 0) row.dk = gross / Number(row.t) * 100;
            delete row.grossQD;
            delete row.onlineExcludedQD;
            delete row.onlineExcludedAt;
            changed = true;
        }
        if (changed) GM_setValue(DATA_KEY, cache);
    }

    function installFix() {
        const health = window.__AutoBIHealthPerShop;
        if (!health || health.__shopFix52) return false;

        health.__shopFix52 = true;
        health.scan = async function (config, scrapeFn, toastFn, done) {
            const shops = [];
            for (let index = 1; index <= 5; index++) {
                if (config && config['shop' + index]) shops.push(index);
            }

            restoreRevenueRows();

            try {
                for (let position = 0; position < shops.length; position++) {
                    const index = shops[position];
                    const key = 'shop' + index;
                    const name = config[key + 'Short'] || config[key] || key;

                    if (toastFn) {
                        toastFn('🏬 Sức khỏe ST [' + (position + 1) + '/' + shops.length + ']: ' + name, 3000);
                    }

                    let selected = await health.selectSingleShop(config, index);
                    if (!selected.ok || !isExpectedShop(config, index)) {
                        await sleep(1200);
                        selected = await health.selectSingleShop(config, index);
                    }
                    if (!selected.ok || !isExpectedShop(config, index)) {
                        console.warn('[AutoBI Health 53] Không xác nhận được đúng shop', key, name);
                        continue;
                    }

                    // Không gọi Online scan ở đây: thao tác đó từng trừ nhầm
                    // doanh thu shop. Sức khỏe ST chỉ cần bảng Ngành hàng.
                    await health.waitBI(20000);
                    await sleep(1400);

                    let data = null;
                    let bestScore = -1;
                    let stableScoreCount = 0;
                    for (let attempt = 1; attempt <= 3; attempt++) {
                        const candidate = await scrapeOnce(scrapeFn);
                        const score = healthScore(candidate);
                        if (score > bestScore) {
                            data = candidate;
                            bestScore = score;
                            stableScoreCount = 0;
                        } else if (score === bestScore && score > 0) {
                            stableScoreCount++;
                        }
                        if (attempt >= 2 && stableScoreCount >= 1 && hasHealthData(data)) break;
                        console.warn('[AutoBI Health 53] Đang kiểm tra độ đầy đủ', key, 'lượt', attempt + 1, 'điểm', score);
                        if (toastFn) toastFn('⏳ ' + name + ': đang kiểm tra đủ nhóm ngành hàng...', 3000);
                        await health.waitBI(20000);
                        await sleep(1800);
                    }

                    if (!data || !hasHealthData(data)) {
                        console.warn('[AutoBI Health 53] Không lưu dữ liệu rỗng cho', key);
                        continue;
                    }

                    const cache = GM_getValue(DATA_KEY, {}) || {};
                    if (!cache.link8_health) cache.link8_health = {};
                    cache.link8_health[key] = data;
                    GM_setValue(DATA_KEY, cache);
                    console.info('[AutoBI Health 53] Đã lưu đúng', key, name, data);
                    await sleep(500);
                }
            } catch (error) {
                console.error('[AutoBI Health 53]', error);
                if (toastFn) toastFn('⚠️ Sức khỏe ST: ' + error.message, 8000);
            } finally {
                if (toastFn) toastFn('🔄 Đang khôi phục tất cả siêu thị...', 1800);
                try { await health.restoreAllShops(); } catch (error) {
                    console.warn('[AutoBI Health 53] Không khôi phục được bộ lọc', error);
                }
                if (done) done();
            }
        };
        return true;
    }

    if (!installFix()) {
        const timer = setInterval(() => {
            if (installFix()) clearInterval(timer);
        }, 100);
        setTimeout(() => clearInterval(timer), 30000);
    }
})();
