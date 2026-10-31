/**
 * DocAnhBI — Đọc ảnh chụp màn hình BI ra số đổ AutoBI
 * Tầng 1 (Thu thập) mới của AutoBI: thay cho đường lấy số bằng API/giao diện BI.
 * Ảnh → AI đọc → số theo đúng dạng "Chép bảng" → file Excel đổ tay / dán vào trang đổ số.
 *
 * CÀI ĐẶT (1 lần):
 *  1. Cài đặt dự án (bánh răng) → Thuộc tính tập lệnh → Thêm:
 *       AI_KEY        = khóa AI: Gemini (aistudio.google.com → Get API key) hoặc Claude (sk-ant-…, console.anthropic.com)
 *       MODEL         = không bắt buộc (mặc định gemini-3.1-flash-lite / claude-sonnet-5-5)
 *       MISTRAL_KEY   = khóa Mistral (console.mistral.ai, miễn phí) — dự phòng khi Gemini quá tải (v1.9)
 *       AUTOBI_KEY    = 1 chuỗi bí mật tự đặt (không bắt buộc, nên có) — AutoBI phải gửi đúng chuỗi này mới lấy được số
 *  2. Triển khai → Tác vụ triển khai mới → Ứng dụng web
 *       Thực thi với tư cách: Tôi · Ai có quyền truy cập: Bất kỳ ai có tài khoản Google
 *  3. Lần sau sửa code: Quản lý các lần triển khai → bút chì → Phiên bản mới (giữ nguyên URL)
 */

var DAB_VERSION = '1.9';

/* v1.1: AutoBI lấy số đã đọc: GET <url /exec>?action=lay&ma=<mã nhận>&key=<AUTOBI_KEY nếu có cài> */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'quyen') return json_(kiemQuyen_(p.ma, ''));
  if (p.action === 'lay') return layKetQua_(p);
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('Đọc ảnh BI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* v1.7: KIỂM QUYỀN theo danh sách AutoBI (tab Auth — Cloud kiểm quyền của AutoBI). Mã không có / hết hạn → chặn đọc ảnh, gửi, lấy số.
   Tắt tạm: Thuộc tính tập lệnh AUTH_CHECK = off. Kết quả nhớ 30 phút (bị chặn: 5 phút) để không gọi Cloud quyền mỗi ảnh. */
var AUTH_API_ = 'https://script.google.com/macros/s/AKfycbzSGNHjQ81dzQJcvk9m1lubSDvkn2bQ2qpSyALpmKqyCBSQXtIH_j3738YLVhIjlxGg/exec';
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function kiemQuyen_(ma, ver) {
  ma = String(ma || '').replace(/\D/g, '');
  if (String(PropertiesService.getScriptProperties().getProperty('AUTH_CHECK') || '').toLowerCase() === 'off') return { ok: true, ma: ma };
  if (!ma) return { ok: false, loi: 'Thiếu mã nhân viên' };
  var c = CacheService.getScriptCache(), k = 'auth_' + ma, v = c.get(k);
  if (v) { try { return JSON.parse(v); } catch (e) { } }
  var d = null;
  try {
    var r = UrlFetchApp.fetch(AUTH_API_ + '?type=check_auth&user=' + encodeURIComponent(ma) + (ver ? '&ver=' + encodeURIComponent(ver) : '') + '&_t=' + Date.now(), { muteHttpExceptions: true, followRedirects: true });
    d = JSON.parse(r.getContentText());
  } catch (e) { return { ok: false, tam: true, loi: 'Chưa kiểm được quyền (mạng chậm) — thử lại' }; }
  var out;
  if (d && d.status === 'success' && d.authorized === true) {
    var ex = d.expireDays;
    out = (ex !== null && ex !== undefined && ex !== '' && Number(ex) < 0) ? { ok: false, loi: 'Quyền AutoBI của mã ' + ma + ' đã hết hạn' } : { ok: true, ma: ma, ten: d.name || '' };
  } else if (d && d.authorized === false) out = { ok: false, loi: 'Mã ' + ma + ' chưa được cấp quyền AutoBI' };
  else return { ok: false, tam: true, loi: 'Máy kiểm quyền trả lỗi — thử lại' };
  c.put(k, JSON.stringify(out), out.ok ? 1800 : 300);
  return out;
}
function canQuyen_(ma, ver) { var q = kiemQuyen_(ma, ver); if (!q.ok) throw new Error(q.loi); return q; }

/* v1.7: AutoBI gửi ảnh thẳng (không qua trang): POST JSON {action:'doc', ma, b64, mime, ver} → {ok, kq}
   · {action:'luu', ma, goi} → lưu gói số (gộp theo phần như trang) · {action:'quyen', ma}. Mã kiểm theo tab Auth. */
function doPost(e) {
  var b = {}, out;
  try { b = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) { return json_({ ok: false, loi: 'Dữ liệu gửi lên không đúng dạng' }); }
  try {
    var q = kiemQuyen_(b.ma, b.ver);
    if (!q.ok) out = { ok: false, quyen: !q.tam, loi: q.loi };
    else if (b.action === 'doc') out = { ok: true, kq: docAnh_(b.b64, b.mime) };
    else if (b.action === 'luu') { out = luuGoi_(q.ma, typeof b.goi === 'string' ? b.goi : JSON.stringify(b.goi)); }
    else if (b.action === 'quyen') out = { ok: true, ten: q.ten || '' };
    else out = { ok: false, loi: 'Thiếu việc cần làm (action)' };
  } catch (er) { out = { ok: false, loi: String((er && er.message) || er) }; }
  return json_(out);
}

function tenFile_(ma) { return 'DocAnhBI_' + String(ma || '').replace(/[^0-9A-Za-z_-]/g, '') + '.json'; }
function maHopLe_(ma) { return /^[0-9A-Za-z_-]{2,20}$/.test(String(ma || '')); }

/** Trang gọi: lưu gói số (JSON) theo mã nhận — ghi đè bản cũ cùng mã. File nằm trong Drive của chủ dự án. */
function luuKetQua(ma, json) {
  ma = String(ma || '').trim();
  canQuyen_(ma, '');
  return luuGoi_(ma, json);
}
function luuGoi_(ma, json) {
  ma = String(ma || '').trim();
  if (!maHopLe_(ma)) throw new Error('Mã nhận số không hợp lệ (2–20 ký tự chữ/số)');
  if (!json || json.length > 2000000) throw new Error('Gói số rỗng hoặc quá lớn');
  JSON.parse(json);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var goi;
  try {
    var it = DriveApp.getFilesByName(tenFile_(ma)), f = it.hasNext() ? it.next() : null, cu = null;
    if (f) { try { cu = JSON.parse(f.getBlob().getDataAsString()); } catch (e) { cu = null; } }
    goi = gopGoi_(cu, JSON.parse(json));
    var txt = JSON.stringify(goi);
    if (f) f.setContent(txt); else DriveApp.createFile(tenFile_(ma), txt, MimeType.PLAIN_TEXT);
  } finally { lock.releaseLock(); }
  return { ok: true, at: Date.now(), atLK: goi.atLK || null, atRT: goi.atRT || null };
}

/* v1.6: GỘP THEO PHẦN TRONG CÙNG NGÀY — sáng gửi lũy kế, trưa gửi realtime, ca chiều gửi cả hai… phần nào gói mới có thì thay,
   phần nào không có thì giữ của lần gửi trước (chỉ khi cùng ngày giờ VN; sang ngày mới thì bỏ hết số cũ). */
var PHAN_LK_ = ['tdST', 'tdNV', 'dtNV', 'tcNV'], PHAN_RT_ = ['tdSTRT', 'dtNVRT', 'tcNVRT'];
var SH_LK_ = ['lkQD', 'lkThuc', 'lkTC', 'lkTraGop', 'target', 'pvB', 'pvK'], SH_RT_ = ['rtQD', 'rtThuc', 'rtTC', 'rtTraGop', 'rtTarget', 'rtPvB', 'rtPvK'];
function vnNgay_(t) { return Utilities.formatDate(new Date(Number(t) || Date.now()), 'Asia/Ho_Chi_Minh', 'yyyyMMdd'); }
function coGi_(o, f) { return f.some(function (k) { return o && o[k] !== null && o[k] !== undefined; }); }
function coPhan_(g, phan, sh) {
  return phan.some(function (k) { return (g[k] || []).length; }) || (g.shops || []).some(function (s) { return coGi_(s, sh); });
}
function gopGoi_(cu, moi) {
  var coLK = coPhan_(moi, PHAN_LK_, SH_LK_), coRT = coPhan_(moi, PHAN_RT_, SH_RT_);
  if (!cu || cu.v !== 1 || vnNgay_(cu.at) !== vnNgay_(moi.at)) {
    moi.atLK = coLK ? moi.at : null; moi.atRT = coRT ? moi.at : null; return moi;
  }
  var out = JSON.parse(JSON.stringify(moi));
  PHAN_LK_.concat(PHAN_RT_).forEach(function (k) { if (!(moi[k] || []).length && (cu[k] || []).length) out[k] = cu[k]; });
  var byMa = {};
  (cu.shops || []).forEach(function (s) { byMa[s.ma] = s; });
  out.shops = (out.shops || []).map(function (s) {
    var o = byMa[s.ma];
    if (o) {
      [SH_LK_, SH_RT_].forEach(function (f) { if (!coGi_(s, f)) f.forEach(function (k) { if (o[k] !== null && o[k] !== undefined) s[k] = o[k]; }); });
      if (!s.bi && o.bi) s.bi = o.bi; if (!s.gon && o.gon) s.gon = o.gon;
      delete byMa[s.ma];
    }
    return s;
  });
  Object.keys(byMa).forEach(function (m) { out.shops.push(byMa[m]); });
  out.atLK = coLK ? moi.at : (cu.atLK || null);
  out.atRT = coRT ? moi.at : (cu.atRT || null);
  if (!out.ngayQua) out.ngayQua = cu.ngayQua || null;
  if (!out.soNgay) out.soNgay = cu.soNgay || null;
  return out;
}

function layKetQua_(p) {
  var key = PropertiesService.getScriptProperties().getProperty('AUTOBI_KEY');
  var out, q0;
  if (key && String(p.key || '') !== key) out = { ok: false, loi: 'Sai khóa AUTOBI_KEY' };
  else if (!maHopLe_(p.ma)) out = { ok: false, loi: 'Thiếu mã nhận số' };
  else if (!(q0 = kiemQuyen_(p.ma, '')).ok) out = { ok: false, quyen: !q0.tam, loi: q0.loi };
  else {
    var it = DriveApp.getFilesByName(tenFile_(p.ma));
    out = it.hasNext() ? { ok: true, data: JSON.parse(it.next().getBlob().getDataAsString()) }
                       : { ok: false, loi: 'Chưa có số cho mã ' + p.ma + ' — vào trang Đọc ảnh BI bấm "Gửi sang AutoBI" trước' };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function thongTin() {
  var p = PropertiesService.getScriptProperties();
  var url = ''; try { url = ScriptApp.getService().getUrl(); } catch (e) { }
  var k = p.getProperty('AI_KEY') || p.getProperty('ANTHROPIC_KEY') || p.getProperty('GEMINI_KEY') || '';
  return { version: DAB_VERSION, coKhoa: !!k, model: p.getProperty('MODEL') || (/^sk-ant-/.test(k) ? 'claude-sonnet-5-5' : 'gemini-3.1-flash-lite'), url: url, coKhoaAutoBI: !!p.getProperty('AUTOBI_KEY') };
}

var PROMPT = [
  'Bạn đọc ẢNH CHỤP MÀN HÌNH báo cáo BI Điện Máy Xanh (app/web, tiếng Việt) và trả về DUY NHẤT một đối tượng JSON, không giải thích, không ```.',
  '',
  'Có 3 loại màn hình:',
  'A) "thi_dua" — tiêu đề "Thi đua": nhiều bảng, mỗi bảng có ô tiêu đề là tên chương trình (vd "T10 - MÁY GIẶT"), cột "DOANH THU" hoặc "SỐ LƯỢNG", có thể có "TARGET", "% HT THÁNG"/"% HT NGÀY", "HẠNG TRONG ST"… Dòng đầu là "TỔNG", sau đó là các dòng siêu thị (vd "ĐMM_NDI_HHA - Hải Anh") hoặc nhân viên (vd "204372 - Nguyễn Quốc Khánh").',
  'C) "tra_cham" — tiêu đề "Báo cáo trả chậm": bảng theo Nhân viên (hoặc Siêu thị) với cột "DT trả góp (triệu)", có thể có "DT siêu thị", "Tỷ lệ/Tỷ trọng trả chậm (%)".',
  'B) "doanh_thu" — tiêu đề "Doanh thu hợp nhất": các thẻ số (DT quy đổi / DT thực, DT dự kiến, % HT target "Target trọn kỳ …", Tỷ trọng trả chậm "x%" kèm "DT trả góp QĐ a / b"), đường dẫn "Toàn công ty > Siêu thị 5243", và có thể có bảng nhân viên (NHÂN VIÊN, SỐ LƯỢNG, DOANH THU QĐ…).',
  '',
  'Dạng JSON:',
  '{"man_hinh":"thi_dua"|"doanh_thu"|"tra_cham"|"khac",',
  ' "che_do":"luy_ke"|"realtime"|null,          // nút đang chọn (màu xanh đậm) Lũy kế / Realtime; không thấy nút thì null',
  ' "loai_dt":"quy_doi"|"thuc"|null,            // chỉ màn doanh_thu: nút DT quy đổi / DT thực đang chọn',
  ' "sieu_thi":"5243"|null,                     // chỉ màn doanh_thu: mã siêu thị ở đường dẫn "Siêu thị xxxx"; không có thì null',
  ' "ngay_qua":số|null,"so_ngay":số|null,        // dòng "Quỹ thời gian: 30/31 ngày" → 30 và 31',
  ' "kpi":{"dt":số,"target":số|null,"tra_cham_pct":số|null,"dt_tra_gop":số|null,"pv_bill":số|null,"pv_khach":số|null,"target_ngay":số|null}|null,   // dt = số lớn ở thẻ DT quy đổi (hoặc DT thực); target = "Target trọn kỳ" (chế độ Lũy kế); target_ngay = "Target ngày" ở thẻ "% HT target ngày" (chế độ Realtime); tra_cham_pct = số % ở thẻ Tỷ trọng trả chậm; dt_tra_gop = số a trong "DT trả góp QĐ a / b"; pv_bill, pv_khach = 2 số trong "495 bill / 3,469 khách" ở thẻ TLPVTC',
  ' "nhan_vien":[{"ten":"204372 - Nguyễn Qu","so_luong":số|null,"doanh_thu":số}],     // bảng nhân viên của màn doanh_thu, đúng thứ tự trên ảnh',
  ' "tra_cham":[{"ten":"Nguyễn Quốc Khánh","dt_tra_gop":số,"dt_sieu_thi":số|null,"ty_le":số|null}],   // chỉ màn tra_cham: từng dòng của bảng (tên NV hoặc siêu thị như hiển thị); dt_tra_gop = cột DT trả góp; dt_sieu_thi, ty_le (%) nếu cột đó hiện đủ, không thấy thì null',
  ' "khoi":[{"chuong_trinh":"T10 - MÁY GIẶT"|null,"don_vi":"DOANH THU"|"SỐ LƯỢNG","tong":số|null,"tong_target":số|null,',
  '          "dong":[{"ten":"…","gia_tri":số,"target":số|null}]}]}',
  '',
  'QUY TẮC:',
  '- Chép số ĐÚNG như trên ảnh. Dấu chấm là thập phân, dấu phẩy là phân cách nghìn: "2,304.58" → 2304.58; "1,247.00" → 1247; số âm giữ dấu trừ.',
  '- Tên: chép nguyên văn như hiển thị (kể cả khi bị cắt cụt), giữ dấu tiếng Việt, gộp chữ xuống dòng thành 1 dòng.',
  '- Bỏ qua các cột phần trăm bị cắt kiểu "184....", cột hạng, TOP/BOTTOM, chữ watermark mờ chéo (vd "38967 Ngọc Phạm Ánh").',
  '- Dòng nằm ở mép trên/dưới ảnh mà SỐ không hiện đủ (bị cắt nửa) thì KHÔNG đưa vào.',
  '- Bảng ở đầu ảnh không thấy ô tên chương trình (cuộn tiếp từ ảnh trước) → "chuong_trinh": null. Nếu ô tiêu đề có hiện (kể cả bị che một phần) thì chép tên.',
  '- "tong" là số ở dòng TỔNG; không thấy dòng TỔNG thì null. "tong_target" là TARGET ở dòng TỔNG nếu có.',
  '- Mọi dòng đúng thứ tự từ trên xuống. Một người/siêu thị có thể xuất hiện 2 lần với số khác nhau — giữ cả 2.',
  '- Màn tra_cham: bỏ dòng "Admin" và dòng tổng; cột bị cắt mép phải không đọc được thì để null.',
  '- Màn không có loại dữ liệu nào thì để mảng rỗng [] hoặc null.'
].join('\n');

/* v1.2: dùng được 2 loại khóa — tự nhận theo dạng khóa:
   sk-ant-…  → Claude (Anthropic)      · MODEL mặc định claude-sonnet-5-5
   còn lại   → Gemini (Google AI Studio) · MODEL mặc định gemini-3.1-flash-lite (miễn phí, nhanh; 31/10/2026 các bản flash lớn hơn hay treo) — máy chủ báo không có model → tự chọn bản flash-lite mới nhất */
function cauHinhAI_() {
  var p = PropertiesService.getScriptProperties();
  var key = p.getProperty('AI_KEY') || p.getProperty('ANTHROPIC_KEY') || p.getProperty('GEMINI_KEY');
  if (!key) throw new Error('Chưa cài khóa AI: Cài đặt dự án → Thuộc tính tập lệnh → AI_KEY');
  var loai = /^sk-ant-/.test(key) ? 'claude' : 'gemini';
  var model = p.getProperty('MODEL') || (loai === 'claude' ? 'claude-sonnet-5-5' : 'gemini-3.1-flash-lite');
  return { key: key, loai: loai, model: model, p: p, mistral: p.getProperty('MISTRAL_KEY') || '' };
}

/* v1.9: MISTRAL làm dự phòng — Gemini nghẽn hết vòng đầu thì đọc bằng Mistral (khóa MISTRAL_KEY, miễn phí ~1 lần/giây).
   Ảnh gửi dạng data URL; trả JSON (response_format json_object). Hết model / quá hạn mức thì trả null để Gemini thử vòng 2. */
function goiMistral_(key, parts, tries) {
  var content = parts.map(function (x) {
    return x.img ? { type: 'image_url', image_url: 'data:' + (x.mime || 'image/jpeg') + ';base64,' + x.img } : { type: 'text', text: x.text };
  });
  var chain = ['ministral-14b-latest', 'mistral-medium-latest', 'mistral-small-latest'], cache = CacheService.getScriptCache();
  for (var k = 0; k < chain.length; k++) {
    if (cache.get('ai_bad_' + chain[k])) continue;
    var code = 0, txt = '';
    try {
      var res = UrlFetchApp.fetch('https://api.mistral.ai/v1/chat/completions', { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + key },
        payload: JSON.stringify({ model: chain[k], temperature: 0, max_tokens: 8192, response_format: { type: 'json_object' }, messages: [{ role: 'user', content: content }] }) });
      code = res.getResponseCode(); txt = res.getContentText();
    } catch (e) { code = 0; txt = String(e && e.message || e); }
    if (code === 200) {
      var d = JSON.parse(txt), m = ((d.choices || [])[0] || {}).message || {};
      var out = typeof m.content === 'string' ? m.content : (m.content || []).map(function (c) { return c.text || ''; }).join('');
      return { text: out, usage: d.usage || null, model: chain[k] };
    }
    tries.push(chain[k] + ':' + code + (code === 200 ? '' : ' ' + String(txt).slice(0, 160)));
    if (code === 401 || code === 403) return null; /* sai khóa Mistral → bỏ qua, không làm hỏng đường Gemini */
    try { if (code === 0 || code === 429 || code >= 500) cache.put('ai_bad_' + chain[k], '1', code === 429 ? 30 : 120); } catch (e) { }
  }
  return null;
}

function goiAI_(cfg, parts) {
  var url, opt, body;
  if (cfg.loai === 'claude') {
    url = 'https://api.anthropic.com/v1/messages';
    body = { model: cfg.model, max_tokens: 4000, temperature: 0, messages: [{ role: 'user', content: parts.map(function (x) {
      return x.img ? { type: 'image', source: { type: 'base64', media_type: x.mime || 'image/jpeg', data: x.img } } : { type: 'text', text: x.text };
    }) }] };
    opt = { headers: { 'x-api-key': cfg.key, 'anthropic-version': '2023-06-01' } };
  } else {
    url = 'https://generativelanguage.googleapis.com/v1beta/models/' + cfg.model + ':generateContent';
    body = { contents: [{ role: 'user', parts: parts.map(function (x) {
      return x.img ? { inline_data: { mime_type: x.mime || 'image/jpeg', data: x.img } } : { text: x.text };
    }) }], generationConfig: { temperature: 0, responseMimeType: 'application/json', maxOutputTokens: 8192 } };
    opt = { headers: { 'x-goog-api-key': cfg.key } };
  }
  var res, code = 0, txt, lastErr = '';
  /* v1.4: Gemini miễn phí hay báo 503 "high demand" theo từng model → xoay vòng sang model khác ngay, hết vòng thì nghỉ rồi thử thêm 1 vòng.
     Bản flash-lite trả lời nhanh; bản flash lớn để cuối (31/10/2026 hay treo ~60s). */
  var chain = cfg.loai === 'gemini'
    ? [cfg.model, 'gemini-3.1-flash-lite', 'gemini-3.5-flash-lite', 'gemini-2.5-flash-lite', 'gemini-flash-lite-latest', 'gemini-3.5-flash', 'gemini-flash-latest']
        .filter(function (m, i, a) { return m && a.indexOf(m) === i; })
    : [cfg.model];
  /* v1.8: nhớ model vừa đọc được (10 phút) để thử trước; model vừa quá tải/treo thì bỏ qua 2 phút ở vòng đầu
     — các ảnh đọc song song không cùng chờ một model đang nghẽn (01/11/2026 có ảnh mất 3–4 phút vì xoay hết vòng). */
  var cache = CacheService.getScriptCache();
  if (cfg.loai === 'gemini') {
    var tot = cache.get('ai_ok');
    if (tot && chain.indexOf(tot) >= 0) chain = [tot].concat(chain.filter(function (m) { return m !== tot; }));
    var nghen = chain.filter(function (m) { return cache.get('ai_bad_' + m); });
    if (nghen.length < chain.length) chain = chain.filter(function (m) { return nghen.indexOf(m) < 0; }).concat(nghen);
  }
  var tries = [];
  for (var vong = 0; vong < 2 && code !== 200; vong++) {
    if (vong) {
      if (cfg.loai === 'gemini' && cfg.mistral) { var mi = goiMistral_(cfg.mistral, parts, tries); if (mi) { cfg.model = mi.model; return mi; } }
      Utilities.sleep(5000);
    }
    for (var k = 0; k < chain.length; k++) {
      if (vong === 0 && cfg.loai === 'gemini' && cfg.mistral && cache.get('ai_bad_' + chain[k])) continue; /* có Mistral thì không chờ model đang nghẽn */
      if (cfg.loai === 'gemini') url = 'https://generativelanguage.googleapis.com/v1beta/models/' + chain[k] + ':generateContent';
      try {
        res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', headers: opt.headers, payload: JSON.stringify(body), muteHttpExceptions: true });
        code = res.getResponseCode(); txt = res.getContentText();
      } catch (e) { code = 0; txt = String(e && e.message || e); }
      if (code === 200) { cfg.model = chain[k]; try { cache.put('ai_ok', chain[k], 600); } catch (e) { } break; }
      try { if (code === 0 || code === 429 || code >= 500) cache.put('ai_bad_' + chain[k], '1', 120); } catch (e) { }
      tries.push(chain[k] + ':' + code);
      lastErr = 'Lỗi AI ' + cfg.loai + ' (mã ' + code + '): ' + txt.slice(0, 200);
      if (code === 401 || code === 403 || code === 400) throw new Error(lastErr); /* sai khóa / sai yêu cầu: đổi model cũng vô ích */
      if (cfg.loai === 'claude') { if (code === 429 || code >= 500 || code === 0) { Utilities.sleep(3000); continue; } throw new Error(lastErr); }
    }
  }
  if (code !== 200) throw new Error('AI đang quá tải, thử lại sau ít phút (' + tries.join(', ') + ') — ' + lastErr);
  var data = JSON.parse(txt), out;
  if (cfg.loai === 'claude') out = (data.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('');
  else out = ((((data.candidates || [])[0] || {}).content || {}).parts || []).map(function (c) { return c.text || ''; }).join('');
  return { text: out, usage: data.usage || data.usageMetadata || null };
}

/** Đọc 1 ảnh (base64 JPEG/PNG) → JSON kết quả. Gọi từ trang bằng google.script.run — v1.7: phải kèm mã NV có quyền */
function docAnh(b64, mime, ma) {
  canQuyen_(ma, '');
  return docAnh_(b64, mime);
}
function docAnh_(b64, mime) {
  if (!b64 || String(b64).length > 8000000) throw new Error('Ảnh rỗng hoặc quá lớn');
  var cfg = cauHinhAI_();
  var r = goiAI_(cfg, [{ img: b64, mime: mime || 'image/jpeg' }, { text: PROMPT }]);
  var out = String(r.text || '').replace(/^[\s\S]*?(\{)/, '$1').replace(/\}[^}]*$/, '}');
  try {
    var obj = JSON.parse(out);
    obj._token = r.usage;
    obj._model = cfg.model;
    return obj;
  } catch (e) {
    throw new Error('AI trả về không đúng dạng JSON — thử chụp lại ảnh rõ hơn');
  }
}

/** Chạy tay trong trình soạn thảo để thử khóa AI (xem Nhật ký thực thi) */
function kiemTraKhoa() {
  var cfg = cauHinhAI_();
  var r = goiAI_(cfg, [{ text: 'Trả về đúng JSON {"ok":true}' }]);
  Logger.log('Khóa ' + cfg.loai + ' · model ' + cfg.model + ' → ' + String(r.text).slice(0, 200));
  return { loai: cfg.loai, model: cfg.model, tra_loi: String(r.text).slice(0, 200) };
}

/** v1.9: chạy tay để thử khóa Mistral (xem Nhật ký thực thi) */
function kiemTraMistral() {
  var k = PropertiesService.getScriptProperties().getProperty('MISTRAL_KEY');
  if (!k) throw new Error('Chưa cài MISTRAL_KEY');
  var t = [], r = goiMistral_(k, [{ text: 'Trả về đúng JSON {"ok":true}' }], t);
  Logger.log(r ? 'Mistral ' + r.model + ' → ' + String(r.text).slice(0, 200) : 'Mistral không trả lời: ' + t.join(' | '));
  return r;
}
