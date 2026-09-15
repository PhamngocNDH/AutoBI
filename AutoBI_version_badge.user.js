/* AutoBI 16.1.1.53 - nhãn phiên bản; không can thiệp dữ liệu báo cáo. */
(function () {
    'use strict';

    const VERSION = '16.1.1.53';
    const BADGE_ID = 'autobi-version-badge';

    function showVersionBadge() {
        if (!document.body || document.getElementById(BADGE_ID)) return;

        const badge = document.createElement('div');
        badge.id = BADGE_ID;
        badge.textContent = 'AutoBI v' + VERSION;
        badge.title = 'Phiên bản đang sử dụng: AutoBI ' + VERSION + ' (nền ổn định 16.1.1.44)';
        badge.style.cssText = [
            'position:fixed',
            'right:14px',
            'bottom:78px',
            'z-index:2147483646',
            'padding:7px 12px',
            'border-radius:999px',
            'background:linear-gradient(135deg,#111827,#334155)',
            'color:#fff',
            'font:700 12px/1.2 Arial,sans-serif',
            'letter-spacing:.2px',
            'box-shadow:0 4px 14px rgba(15,23,42,.28)',
            'border:1px solid rgba(255,255,255,.28)',
            'pointer-events:none',
            'user-select:none'
        ].join(';');
        document.body.appendChild(badge);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', showVersionBadge, { once: true });
    } else {
        showVersionBadge();
    }

    new MutationObserver(showVersionBadge).observe(document.documentElement, {
        childList: true,
        subtree: true
    });
})();
