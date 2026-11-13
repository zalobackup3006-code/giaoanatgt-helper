(function () {
'use strict';
var L = window.LESSONS || [], MACH = window.MACH || [];
var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
var KEYK = 'atgt_gemini_key', MODK = 'atgt_gemini_model';
var PREF = ['gemini-3.8-flash', 'gemini-3.5-flash', 'gemini-3.6-flash', 'gemini-3.5-flash-lite', 'gemini-3.1-flash-lite', 'gemini-2.5-flash'];
var FN = { lh: 'Liên hệ', bp: 'Bộ phận', tp: 'Toàn phần', khong: 'Không tích hợp' };
var NGT = ['Bảo đảm mục tiêu bài học', 'Không làm quá tải chương trình, nội dung bài học', 'Không phá vỡ nội dung môn học',
  'Không đưa nội dung tích hợp quá xa lạ với bài học', 'Tích hợp tự nhiên, không gò ép', 'Có liên hệ với thực tiễn địa phương'];
var S = { paras: [], zip: null, xml: null, buf: null, xmlStr: '', roles: {}, idx: null, stat: null, name: '', r: null, form: '', acc: {}, ed: {} };
var $ = function (id) { return document.getElementById(id); };
function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function get(k) { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } }
function put(k, v) { try { if (v) localStorage.setItem(k, v); else localStorage.removeItem(k); } catch (e) {} }
function byId(id) { for (var i = 0; i < L.length; i++) if (L[i].id === id) return L[i]; return null; }
function machOf(id) { for (var i = 0; i < MACH.length; i++) if (MACH[i].id === id) return MACH[i]; return null; }
function lines(t) { return String(t || '').split('\n').map(function (x) { return x.replace(/^\s*-\s*/, '').trim(); }).filter(Boolean); }

/* ===== Khoa Gemini ===== */
function drawKey() {
  var k = get(KEYK), box = $('key');
  if (k) {
    box.innerHTML = '<span class="mute">Khóa Gemini: <span class="ok">đã lưu trên máy này</span>' + (get(MODK) ? ' (' + esc(get(MODK)) + ')' : '') + ' · <button class="link" id="kDel" type="button">đổi khóa</button></span>';
    $('kDel').onclick = function () { put(KEYK, ''); put(MODK, ''); drawKey(); };
    return;
  }
  box.innerHTML = '<div class="item" style="margin-top:0"><b>Khóa Gemini của thầy cô</b> <span class="mute">(làm một lần trên mỗi máy)</span>' +
    '<p class="mute" style="margin:6px 0 8px">Mở <a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noopener">Google AI Studio</a>, đăng nhập tài khoản Google của thầy cô, bấm "Create API key", sao chép rồi dán vào đây. Khóa chỉ lưu trong trình duyệt máy này và gửi thẳng tới Google.</p>' +
    '<div class="row"><input type="text" id="kIn" autocomplete="off" spellcheck="false" placeholder="Dán khóa" style="flex:1 1 240px"><button class="btn" id="kSave" type="button">Lưu</button></div><p id="kMsg" class="mute" style="margin:6px 0 0"></p></div>';
  $('kSave').onclick = function () {
    var v = $('kIn').value.trim(); if (!v) return;
    put(KEYK, v); put(MODK, ''); $('kMsg').textContent = 'Đang kiểm tra...';
    pickModel().then(function (m) { put(MODK, m); drawKey(); })
      .catch(function (e) { put(KEYK, ''); $('kMsg').innerHTML = '<span class="warn">' + esc(e.message) + '</span>'; });
  };
}

/* ===== Doc file Word (giu nguyen cau truc) ===== */
function closest(el, n) { var x = el.parentNode; while (x && x.nodeType === 1) { if (x.namespaceURI === W && x.localName === n) return x; x = x.parentNode; } return null; }
function inside(el, names) { var x = el.parentNode; while (x && x.nodeType === 1) { if (names.indexOf(x.localName) >= 0) return true; x = x.parentNode; } return false; }
function kids(el, n) { var a = []; for (var c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && c.localName === n) a.push(c); return a; }
function textOf(p) {
  var o = '', all = p.getElementsByTagNameNS(W, '*');
  for (var i = 0; i < all.length; i++) { var n = all[i].localName; if (n === 't') o += all[i].textContent; else if (n === 'tab' || n === 'br') o += ' '; }
  return o.replace(/\s+/g, ' ').trim();
}
function span(tc) { var pr = kids(tc, 'tcPr')[0], g = pr && kids(pr, 'gridSpan')[0]; return g ? (parseInt(g.getAttributeNS(W, 'val') || g.getAttribute('w:val'), 10) || 1) : 1; }
function colOf(tc) { var c = 0, cs = kids(tc.parentNode, 'tc'); for (var i = 0; i < cs.length && cs[i] !== tc; i++) c += span(cs[i]); return c; }
function cellAt(tr, col) { var c = 0, cs = kids(tr, 'tc'); for (var i = 0; i < cs.length; i++) { if (c === col) return cs[i]; c += span(cs[i]); } return null; }
function hasImg(el) { return el.getElementsByTagNameNS(W, 'drawing').length + el.getElementsByTagNameNS(W, 'pict').length + el.getElementsByTagNameNS(W, 'object').length; }
function allP(doc) { var ps = doc.getElementsByTagNameNS(W, 'body')[0].getElementsByTagNameNS(W, 'p'), o = []; for (var i = 0; i < ps.length; i++) o.push(ps[i]); return o; }
function readDocx(file) {
  return file.arrayBuffer().then(function (buf) { S.buf = buf; return JSZip.loadAsync(buf); }).then(function (zip) {
    var f = zip.file('word/document.xml'); if (!f) throw new Error('Không phải file Word .docx.');
    return f.async('string').then(function (xml) {
      var doc = new DOMParser().parseFromString(xml, 'application/xml');
      if (doc.getElementsByTagName('parsererror').length) throw new Error('File Word bị lỗi cấu trúc.');
      var ps = allP(doc), a = [], tid = new Map(), roles = {}, idx = new Map();
      ps.forEach(function (p, i) {
        idx.set(p, i);
        var skip = inside(p, ['txbxContent', 'Fallback']), tc = skip ? null : closest(p, 'tc'), o = { text: textOf(p), cell: '', t: -1, r: -1, c: -1, skip: skip, img: !!hasImg(p) };
        if (tc) {
          var tbl = closest(tc, 'tbl'); if (!tid.has(tbl)) tid.set(tbl, tid.size);
          o.t = tid.get(tbl); o.r = kids(tbl, 'tr').indexOf(tc.parentNode); o.c = colOf(tc);
          var R = roles[o.t] = roles[o.t] || {};
          if (o.r <= 2 && o.text.length < 60) {
            if (/giáo viên|teacher/i.test(o.text) && R.gv == null) R.gv = o.c;
            else if (/học sinh|student|pupil/i.test(o.text) && R.hs == null) R.hs = o.c;
            else if (/^(thời gian|thời lượng|time)/i.test(o.text) && R.tg == null) R.tg = o.c;
          }
        }
        a.push(o);
      });
      a.forEach(function (o) { if (o.t >= 0) { var R = roles[o.t]; o.cell = 'bảng ' + (o.t + 1) + ', ' + (R.gv === o.c ? 'cột giáo viên' : R.hs === o.c ? 'cột học sinh' : 'cột ' + (o.c + 1)); } });
      S._els = null; S.zip = zip; S.xml = doc; S.xmlStr = xml; S.paras = a; S.roles = roles; S.idx = idx;
      S.stat = count(doc);
      return a;
    });
  });
}
function count(doc) {
  var g = function (n) { return doc.getElementsByTagNameNS(W, n).length; };
  return { tbl: g('tbl'), tr: g('tr'), tc: g('tc'), img: g('drawing') + g('pict') + g('object'), sect: g('sectPr'), p: g('p') };
}
function numbered() { return S.paras.map(function (p, i) { var t = p.text.length > 300 ? p.text.slice(0, 300) + '...' : p.text; return t && !p.skip ? '[' + i + ']' + (p.cell ? ' (' + p.cell + ')' : '') + ' ' + t : ''; }).filter(Boolean).join('\n'); }

/* ===== Gemini ===== */
function api(path, opt) {
  var k = get(KEYK); if (!k) return Promise.reject(new Error('Chưa có khóa Gemini.'));
  opt = opt || {}; opt.headers = Object.assign({ 'x-goog-api-key': k }, opt.headers || {});
  return fetch('https://generativelanguage.googleapis.com/v1beta/' + path, opt).then(function (r) {
    return r.json().catch(function () { return {}; }).then(function (j) {
      if (r.ok) return j;
      var m = (j.error && j.error.message) || ('Lỗi ' + r.status);
      if (r.status === 400 && /API key/i.test(m)) m = 'Khóa không hợp lệ. Kiểm tra lại khóa đã dán.';
      else if (r.status === 403) m = 'Khóa chưa được phép dùng Gemini. Tạo khóa mới tại Google AI Studio.';
      else if (r.status === 429) m = 'Hết hạn mức trong lúc này. Đợi vài phút rồi thử lại.';
      else if (r.status >= 500) m = 'Gemini đang bận. Thử lại sau ít phút.';
      var e = new Error(m); e.status = r.status; throw e;
    });
  }, function () { throw new Error('Không kết nối được tới Gemini. Kiểm tra mạng.'); });
}
function pickModel() {
  return api('models?pageSize=200').then(function (j) {
    var n = (j.models || []).filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0; }).map(function (m) { return m.name.replace(/^models\//, ''); });
    var chain = PREF.filter(function (x) { return n.indexOf(x) >= 0; });
    n.forEach(function (x) { if (/flash/.test(x) && !/preview|exp|image|tts|live|audio|embed/.test(x) && chain.indexOf(x) < 0) chain.push(x); });
    put('atgt_gemini_models', JSON.stringify(chain.slice(0, 5)));
    if (chain.length) return chain[0];
    var f = n.filter(function (x) { return /flash/.test(x) && !/preview|exp|image|tts|live|audio/.test(x); });
    if (f.length) return f[0]; if (n.length) return n[0];
    throw new Error('Không tìm thấy mô hình Gemini cho khóa này.');
  });
}
var STR = { type: 'STRING' }, INT = { type: 'INTEGER' }, ARR = { type: 'ARRAY', items: { type: 'STRING' } };
function obj(p) { return { type: 'OBJECT', properties: p, required: Object.keys(p), propertyOrdering: Object.keys(p) }; }
var SCHEMA = obj({
  bai_hoc: obj({ mon: STR, lop: INT, ten_bai: STR, thoi_luong_tiet: INT, neo_ten_bai: INT, neo_yccd: INT }),
  cac_hoat_dong: { type: 'ARRAY', items: obj({ ten: STR, thoi_luong: INT, doan_dau: INT, doan_cuoi: INT }) },
  diem_giao: { type: 'ARRAY', items: obj({ ma_bai_atgt: STR, noi_dung_bai_hoc: STR, ly_do: STR }) },
  de_xuat: obj({ hinh_thuc: { type: 'STRING', enum: ['lh', 'bp', 'tp', 'khong'] }, ly_do: STR }),
  phuong_an: { type: 'ARRAY', items: obj({
    hinh_thuc: { type: 'STRING', enum: ['lh', 'bp', 'tp'] }, phu_hop: { type: 'BOOLEAN' }, ly_do: STR,
    yccd: obj({ thuc_hien: ARR, van_dung: ARR, pham_chat: ARR }),
    dieu_chinh_thoi_gian: STR,
    nguyen_tac: { type: 'ARRAY', items: obj({ dat: { type: 'BOOLEAN' }, giai_thich: STR }) },
    chen: { type: 'ARRAY', items: obj({ neo_doan: INT, trich_neo: STR, hoat_dong: STR, ten: STR, thoi_luong: INT, noi_dung: STR, gv: ARR, hs: ARR, hoc_lieu: STR }) }
  }) }
});
var PLAN = SCHEMA.properties.phuong_an.items;
SCHEMA.properties.phuong_an = { type: 'ARRAY', items: obj({ hinh_thuc: PLAN.properties.hinh_thuc, phu_hop: PLAN.properties.phu_hop, ly_do: STR }) };
function sys() {
  return [
    'Bạn hỗ trợ giáo viên tiểu học Việt Nam tích hợp giáo dục an toàn giao thông (ATGT) vào kế hoạch bài dạy có sẵn. Viết tiếng Việt, văn phong sư phạm, ngắn gọn, không dùng biểu tượng cảm xúc.',
    'BA HÌNH THỨC: Liên hệ (lh): đơn giản nhất, khi chỉ có một số nội dung môn học có liên quan; giáo viên khai thác kiến thức môn học và liên hệ với nội dung ATGT; thường 1 - 3 phút, 1 - 2 chỗ. Bộ phận (bp): khi một phần kiến thức của bài học có nội dung ATGT; thường 5 - 10 phút, một hoạt động hoặc bài tập trọn vẹn. Toàn phần (tp): khi hầu hết kiến thức của bài học cũng là kiến thức ATGT; xuyên suốt các hoạt động.',
    'SÁU NGUYÊN TẮC (đánh giá theo đúng thứ tự): ' + NGT.join('; ') + '.',
    'TIẾN HÀNH: xác định nội dung ATGT hợp lí, liên kết với kiến thức môn học, thời lượng; dùng phương pháp tích cực, phương tiện trực quan; nêu rõ hoạt động của học sinh và hoạt động trợ giúp của giáo viên.',
    'NGUỒN ATGT DUY NHẤT: Tài liệu giáo dục ATGT dành cho học sinh lớp 1 - 5 (Bộ GDĐT), 30 bài (mỗi bài có Khởi động, Khám phá, Thực hành, Vận dụng). Chỉ dùng mã bài trong danh sách, ưu tiên cùng lớp; không bịa số trang ngoài trang bắt đầu đã cho.',
    L.map(function (x) { return x.id + ' | Lớp ' + x.lop + ' Bài ' + x.bai + ': ' + x.ten + ' (trang ' + x.trang + '). Mục tiêu: ' + x.muctieu.join(' '); }).join('\n'),
    'ĐẦU RA:',
    '- Giáo án đánh số đoạn [số]. Mọi vị trí là số đoạn có thật. Nội dung chèn đặt NGAY SAU đoạn neo. trich_neo chép nguyên văn 5 - 12 chữ đầu đoạn neo.',
    '- neo_ten_bai: đoạn chứa tên bài. neo_yccd: đoạn CUỐI CÙNG của phần yêu cầu cần đạt / mục tiêu (ngay trước phần đồ dùng dạy học); phần yêu cầu cần đạt về ATGT sẽ chèn ngay sau đoạn này, ở cuối phần yêu cầu cần đạt.',
    '- yccd của mỗi phương án: thuc_hien = học sinh thực hiện được việc gì về ATGT; van_dung = vận dụng được gì vào giải quyết vấn đề ATGT trong thực tế; pham_chat = có cơ hội hình thành, phát triển phẩm chất gì liên quan. Mỗi mục 1 - 2 ý ngắn, bám mục tiêu bài ATGT đã chọn và nội dung bài học; mức độ tương xứng hình thức (liên hệ thì nhẹ nhất).',
    '- Khi phân tích tổng quan: phuong_an ghi đủ 3 hình thức với hinh_thuc, phu_hop, ly_do (một câu ngắn); viết ngắn gọn để trả lời nhanh. Khi được yêu cầu lập chi tiết một hình thức: trả về yccd, dieu_chinh_thoi_gian, nguyen_tac, chen. Không phù hợp thì phu_hop = false, nêu lí do. Không ép toàn phần khi bài không chủ yếu về giao thông. Không có điểm chạm tự nhiên thì de_xuat.hinh_thuc = "khong".',
    '- chen: đặt vào đúng hoạt động phù hợp; nếu giáo án kẻ bảng hai cột thì neo vào đoạn ở cột hoạt động của giáo viên. gv, hs: mỗi phần tử một việc cụ thể, ngắn, hợp lứa tuổi. hoc_lieu: tranh, trò chơi trong tài liệu ATGT lớp mấy bài mấy, hoặc tranh SGK của bài.',
    '- Tổng thời lượng sau tích hợp không vượt thời lượng tiết; nếu thêm thời gian, dieu_chinh_thoi_gian ghi rõ bớt bao nhiêu phút ở hoạt động nào.'
  ].join('\n');
}
function errMsg(st, m) {
  m = m || ('Lỗi ' + st);
  if (st === 400 && /API key/i.test(m)) return 'Khóa không hợp lệ. Kiểm tra lại khóa đã dán.';
  if (st === 403) return 'Khóa chưa được phép dùng Gemini. Tạo khóa mới tại Google AI Studio.';
  if (st === 429) return 'Hết hạn mức trong lúc này. Đợi vài phút rồi thử lại.';
  if (st >= 500) return 'Gemini đang bận. Thử lại sau ít phút.';
  return m;
}
function chain() {
  var c = []; try { c = JSON.parse(get('atgt_gemini_models') || '[]'); } catch (e) {}
  var m = get(MODK);
  if (!c.length) return pickModel().then(function (x) { put(MODK, x); return chain(); });
  if (m && c.indexOf(m) < 0) c.unshift(m);
  if (m) c = [m].concat(c.filter(function (x) { return x !== m; }));
  return Promise.resolve(c);
}
function wait(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
/* Mot lan goi Gemini dang luong: bao tung phan (ket qua, suy nghi) de hien qua trinh */
function once(m, user, schema, onEv, noThink, lim) {
  var cfg = { responseMimeType: 'application/json', responseSchema: schema, temperature: 0.3 };
  if (!noThink) cfg.thinkingConfig = { thinkingLevel: 'low', includeThoughts: true };
  var body = JSON.stringify({ systemInstruction: { parts: [{ text: sys() }] }, contents: [{ role: 'user', parts: [{ text: user }] }], generationConfig: cfg });
  var ac = new AbortController(), got = false, why = '';
  var t1 = setTimeout(function () { if (!got) { why = 'slow'; ac.abort(); } }, 30000);
  var t2 = setTimeout(function () { why = 'long'; ac.abort(); }, lim || 110000);
  function done() { clearTimeout(t1); clearTimeout(t2); }
  return fetch('https://generativelanguage.googleapis.com/v1beta/models/' + m + ':streamGenerateContent?alt=sse', { method: 'POST', headers: { 'Content-Type': 'application/json', 'x-goog-api-key': get(KEYK) }, body: body, signal: ac.signal })
    .then(function (r) {
      if (!r.ok) return r.json().catch(function () { return {}; }).then(function (j) { var raw = (j.error && j.error.message) || ''; var e = new Error(errMsg(r.status, raw)); e.status = r.status; e.raw = raw; throw e; });
      var rd = r.body.getReader(), dec = new TextDecoder(), buf = '', text = '', th = '';
      function pump() {
        return rd.read().then(function (x) {
          if (x.done) return text;
          buf += dec.decode(x.value, { stream: true });
          var ls = buf.split('\n'); buf = ls.pop();
          ls.forEach(function (l) {
            if (l.indexOf('data:') !== 0) return;
            try { var j = JSON.parse(l.slice(5)); var c = j.candidates && j.candidates[0]; (c && c.content && c.content.parts || []).forEach(function (p) { if (!p.text) return; got = true; if (p.thought) th += p.text; else text += p.text; }); } catch (e) {}
          });
          if (onEv) onEv({ text: text, thought: th, model: m });
          return pump();
        });
      }
      return pump();
    }, function () {
      var e = new Error(why ? 'Gemini phản hồi chậm.' : 'Không kết nối được tới Gemini. Kiểm tra mạng.'); e.status = why ? 504 : 0; throw e;
    })
    .then(function (t) {
      done();
      if (!t) { var e = new Error('Gemini không trả lời.'); e.status = 503; throw e; }
      try { return JSON.parse(t); } catch (er) { var e2 = new Error('Kết quả bị cắt ngang.'); e2.status = 503; throw e2; }
    }, function (e) { done(); if (e.name === 'AbortError' || why) { var e3 = new Error('Gemini phản hồi chậm.'); e3.status = 504; throw e3; } throw e; });
}
/* Tu thu lai khi Gemini ban hoac cham, roi chuyen sang mo hinh khac cua cung tai khoan */
function gen(user, schema, onEv, lim) {
  return chain().then(function (ms) {
    var tries = ms.map(function (m) { return { m: m }; });
    var k = 0, noThink = {}, lastErr = null;
    function next() {
      if (k >= tries.length) { var e0 = lastErr || new Error('Gemini đang bận.'); if (e0.raw) e0.message += ' (Chi tiết: ' + e0.raw.slice(0, 140) + ')'; throw e0; }
      var t = tries[k++];
      if (onEv) onEv({ text: '', thought: '', model: t.m, retry: k > 1, start: true });
      return once(t.m, user, schema, onEv, noThink[t.m], lim).catch(function (e) {
        lastErr = e;
        if (e.status === 400 && /think/i.test(e.raw || '') && !noThink[t.m]) { noThink[t.m] = 1; k--; return next(); }
        if (e.status === 404) { if (t.m === get(MODK)) put(MODK, ''); return next(); }
        if (e.status === 0 || e.status === 429 || e.status >= 500) return next();
        throw e;
      });
    }
    return next();
  });
}
function analyze(onEv) { return gen('Phân tích tổng quan.\nKẾ HOẠCH BÀI DẠY:\n' + numbered(), SCHEMA, onEv).then(clean); }
function planFor(f, onEv) {
  var r = S.r, ctx = { bai_hoc: r.bai_hoc, cac_hoat_dong: r.cac_hoat_dong, diem_giao: r.diem_giao };
  return gen('KẾ HOẠCH BÀI DẠY:\n' + numbered() + '\n\nKẾT QUẢ ĐÃ PHÂN TÍCH: ' + JSON.stringify(ctx) + '\n\nHãy lập chi tiết (yccd, dieu_chinh_thoi_gian, nguyen_tac, chen) cho phương án hình thức "' + f + '" (' + FN[f] + '). Trả về một đối tượng phương án.', PLAN, onEv)
    .then(function (p) { p.hinh_thuc = f; p.phu_hop = true; cleanPlan(p); for (var i = 0; i < r.phuong_an.length; i++) if (r.phuong_an[i].hinh_thuc === f) { p.ly_do = p.ly_do || r.phuong_an[i].ly_do; r.phuong_an[i] = p; } return p; });
}

/* ===== Bang tien trinh: the hien Gemini dang lam gi ===== */
function Prog(box, title, steps) {
  var t0 = Date.now(), st = { text: '', thought: '', model: '', retry: false, started: false }, fin = false;
  function grab(re, txt) { var m = re.exec(txt); return m ? m[1] : ''; }
  function all(re, txt) { var o = [], m; re.lastIndex = 0; while ((m = re.exec(txt))) o.push(m[1]); return o; }
  function draw() {
    var txt = st.text, cur = 1;
    if (st.started) cur = 2;
    steps.forEach(function (s, i) { if (i >= 2 && s.key && txt.indexOf('"' + s.key + '"') >= 0) cur = i; });
    if (fin) cur = steps.length;
    var sec = Math.round((Date.now() - t0) / 1000);
    var th = st.thought.replace(/\*\*([^*]+)\*\*/g, '$1: ').replace(/\*\*/g, '').replace(/\s+/g, ' ').trim();
    if (th.length > 260) th = '...' + th.slice(-260);
    box.innerHTML = '<div class="prog"><div class="ph"><b>' + esc(title) + '</b><span class="mute">' + sec + ' giây</span></div>' +
      (st.retry ? '<p class="warn" style="margin:6px 0 0;font-size:.88rem">Mô hình trước bận hoặc chậm, đã chuyển sang ' + esc(st.model) + '.</p>' : '') +
      '<ol class="pst">' + steps.map(function (s, i) {
        var cls = i < cur ? 'done' : i === cur ? 'now' : 'wait';
        var det = typeof s.det === 'function' ? s.det(txt, grab, all, st) : (s.det || '');
        return '<li class="' + cls + '"><span class="dot"></span><span class="lb">' + esc(s.label) + '</span>' + (det ? '<span class="dt">' + esc(det) + '</span>' : '') + '</li>';
      }).join('') + '</ol>' +
      (th && !fin ? '<div class="think"><span class="mute">Gemini đang suy nghĩ:</span> ' + esc(th) + '</div>' : '') +
      (st.text && !fin ? '<p class="mute" style="margin:8px 0 0;font-size:.84rem">Đã nhận ' + st.text.length + ' ký tự kết quả</p>' : '') + '</div>';
  }
  var iv = setInterval(draw, 400); draw();
  return {
    ev: function (e) { if (e.start) { st.started = true; st.text = ''; st.thought = ''; if (e.retry) st.retry = true; } st.model = e.model || st.model; if (!e.start) { st.text = e.text; st.thought = e.thought; } },
    end: function () { fin = true; clearInterval(iv); draw(); return Math.round((Date.now() - t0) / 1000); },
    stop: function () { clearInterval(iv); }
  };
}
var STEPS_A = [
  { label: 'Đọc cấu trúc giáo án trên máy', det: function () { var s0 = S.stat || {}; return S.paras.filter(function (p) { return p.text; }).length + ' đoạn, ' + (s0.tbl || 0) + ' bảng, ' + (s0.img || 0) + ' hình ảnh'; } },
  { label: 'Gửi giáo án tới Gemini', det: function (t, g, a, st) { return st.model; } },
  { label: 'Gemini đọc giáo án, nhận diện bài học và các hoạt động', key: 'bai_hoc', det: function (t, g) { var x = g(/"ten_bai"\s*:\s*"([^"]+)"/, t); return x ? 'Bài học: ' + x : ''; } },
  { label: 'Đối chiếu với 30 bài trong Tài liệu giáo dục ATGT', key: 'diem_giao', det: function (t, g, a) { return a(/"ma_bai_atgt"\s*:\s*"([^"]+)"/g, t).map(byId).filter(Boolean).map(function (x) { return 'Lớp ' + x.lop + ' bài ' + x.bai + ' ' + x.ten; }).join('; '); } },
  { label: 'Chọn hình thức tích hợp phù hợp', key: 'de_xuat', det: function (t) { var i = t.indexOf('"de_xuat"'); if (i < 0) return ''; var m = /"hinh_thuc"\s*:\s*"(\w+)"/.exec(t.slice(i)); return m && FN[m[1]] ? 'Đề xuất: ' + FN[m[1]] : ''; } }
];
var STEPS_B = [
  { label: 'Dùng kết quả phân tích bài học', det: '' },
  { label: 'Gửi yêu cầu lập phương án tới Gemini', det: function (t, g, a, st) { return st.model; } },
  { label: 'Viết yêu cầu cần đạt về ATGT', key: 'yccd' },
  { label: 'Kiểm tra 6 nguyên tắc tích hợp', key: 'nguyen_tac' },
  { label: 'Thiết kế hoạt động tích hợp: vị trí, nội dung, thời lượng, hoạt động giáo viên và học sinh', key: 'chen', det: function (t, g, a) { var i = t.indexOf('"chen"'); if (i < 0) return ''; return a(/"ten"\s*:\s*"([^"]+)"/g, t.slice(i)).join('; '); } }
];
function clamp(n) { n = parseInt(n, 10); return isNaN(n) ? -1 : Math.max(0, Math.min(S.paras.length - 1, n)); }
function nm(s) { return String(s || '').toLowerCase().normalize('NFC').replace(/[^\p{L}\p{N} ]/gu, ' ').replace(/\s+/g, ' ').trim(); }
function fix(i, q) {
  i = clamp(i); q = nm(q).split(' ').slice(0, 6).join(' ');
  if (!q || (i >= 0 && nm(S.paras[i].text).indexOf(q) >= 0)) return i;
  for (var k = 0; k < S.paras.length; k++) if (nm(S.paras[k].text).indexOf(q) >= 0) return k;
  return i;
}
function cleanPlan(p) {
  p.yccd = p.yccd || {}; ['thuc_hien', 'van_dung', 'pham_chat'].forEach(function (k) { p.yccd[k] = p.yccd[k] || []; });
  p.nguyen_tac = p.nguyen_tac || []; p.dieu_chinh_thoi_gian = p.dieu_chinh_thoi_gian || '';
  p.chen = (p.chen || []).map(function (c) { c.neo_doan = fix(c.neo_doan, c.trich_neo); c.gv = c.gv || []; c.hs = c.hs || []; return c; }).filter(function (c) { return c.neo_doan >= 0; });
  p.loaded = p.chen.length > 0 || p.nguyen_tac.length > 0;
  return p;
}
/* Dua vi tri chen yeu cau can dat ATGT xuong cuoi phan yeu cau can dat (ngay truoc phan do dung / thiet bi) */
function endYccd(i) {
  if (i < 0) return i;
  var NEXT = /^(ii|2|b)?[\s.)\-:]*(đồ dùng|thiết bị|chuẩn bị|phương tiện|teaching aids|materials|các hoạt động|hoạt động dạy học|tiến trình)/i;
  for (var k = i + 1; k < S.paras.length && k < i + 80; k++) {
    var p = S.paras[k]; if (p.skip || p.t >= 0 || !p.text) continue;
    if (NEXT.test(p.text)) { for (var j = k - 1; j > i; j--) if (S.paras[j].text && !S.paras[j].skip) return j; return i; }
  }
  return i;
}
function clean(r) {
  r.bai_hoc = r.bai_hoc || {}; r.cac_hoat_dong = r.cac_hoat_dong || [];
  r.diem_giao = (r.diem_giao || []).filter(function (d) { return byId(d.ma_bai_atgt); });
  r.phuong_an = (r.phuong_an || []).filter(function (p) { return FN[p.hinh_thuc]; });
  r.phuong_an.forEach(cleanPlan);
  r.bai_hoc.neo_ten_bai = clamp(r.bai_hoc.neo_ten_bai); r.bai_hoc.neo_yccd = endYccd(clamp(r.bai_hoc.neo_yccd));
  return r;
}

/* ===== Ket qua ===== */
function plan() { var a = S.r.phuong_an; for (var i = 0; i < a.length; i++) if (a[i].hinh_thuc === S.form) return a[i]; return null; }
function K(x) { return S.form + ':' + x; }
function on(x) { return S.acc[K(x)] !== false; }
function v(i, f) { var e = S.ed[K(i)]; if (e && e[f] != null) return e[f]; var c = plan().chen[i]; return Array.isArray(c[f]) ? c[f].join('\n') : (c[f] || ''); }
function yv(f) { var e = S.ed[K('yc')]; if (e && e[f] != null) return e[f]; return (plan().yccd[f] || []).join('\n'); }
function setE(x, f, val) { var k = K(x); S.ed[k] = S.ed[k] || {}; S.ed[k][f] = val; }
function snip(i) { var t = (S.paras[i] || {}).text || ''; return t.length > 80 ? t.slice(0, 80) + '...' : t; }

function result() {
  var r = S.r, b = r.bai_hoc, dx = r.de_xuat || {}, h = '';
  h += '<div class="sec"><h2>' + esc(b.ten_bai) + '</h2><p class="mute">' + esc(b.mon) + ' · Lớp ' + esc(b.lop) + ' · ' + esc(b.thoi_luong_tiet || 35) + ' phút</p></div>';
  h += '<div class="sec"><h2><span class="n">1</span>Hình thức tích hợp</h2><div class="seg" id="seg">' + r.phuong_an.map(function (p) {
    return '<button type="button" data-f="' + p.hinh_thuc + '" class="' + (p.hinh_thuc === S.form ? 'on' : '') + (p.phu_hop ? '' : ' off') + '">' + FN[p.hinh_thuc] + (p.hinh_thuc === dx.hinh_thuc ? ' <small>đề xuất</small>' : '') + '</button>';
  }).join('') + '</div>' + (dx.hinh_thuc === 'khong' ? '<p class="warn" style="margin-top:8px">Hệ thống đề xuất không tích hợp vào bài này: ' + esc(dx.ly_do) + '</p>' : '<p style="margin-top:8px">' + esc(dx.ly_do) + '</p>') + '<div id="why"></div></div>';
  h += '<div class="sec"><h2><span class="n">2</span>Điểm giao và mạch kiến thức</h2>' + (r.diem_giao.length ? r.diem_giao.map(function (d) {
    var x = byId(d.ma_bai_atgt), m = machOf(x.mach), chain = L.filter(function (y) { return y.mach === x.mach; });
    return '<div class="item"><b>Tài liệu ATGT lớp ' + x.lop + ', Bài ' + x.bai + ': ' + esc(x.ten) + '</b> <span class="mute">trang ' + x.trang + '</span>' +
      '<p style="margin:6px 0 0"><span class="mute">Trong bài học:</span> ' + esc(d.noi_dung_bai_hoc) + '</p><p class="mute" style="margin:2px 0 0">' + esc(d.ly_do) + '</p>' +
      (m ? '<p class="mute" style="margin:8px 0 0">Mạch: ' + esc(m.ten) + '</p><div class="chain">' + chain.map(function (y) { return '<span class="' + (y.id === x.id ? 'me' : '') + '">Lớp ' + y.lop + ' · ' + esc(y.ten) + '</span>'; }).join('') + '</div>' : '') + '</div>';
  }).join('') : '<p class="mute">Không tìm thấy điểm giao phù hợp.</p>') + '</div>';
  h += '<div id="pa"></div>';
  $('out').innerHTML = h;
  $('seg').querySelectorAll('button').forEach(function (bt) { bt.onclick = function () { S.form = bt.dataset.f; $('seg').querySelectorAll('button').forEach(function (x) { x.classList.toggle('on', x === bt); }); drawPlan(); }; });
  drawPlan();
}
function drawPlan() {
  var p = plan(), el = $('pa');
  $('why').innerHTML = p ? '<p class="mute" style="margin:8px 0 0">' + FN[p.hinh_thuc] + ': ' + esc(p.ly_do) + '</p>' +
    (p.phu_hop && (p.nguyen_tac || []).length ? '<ul class="list">' + p.nguyen_tac.slice(0, 6).map(function (n, i) {
      return '<li><span class="' + (n.dat ? 'ok' : 'warn') + '">' + (n.dat ? 'Đạt' : 'Lưu ý') + '</span> · ' + NGT[i] + '. <span class="mute">' + esc(n.giai_thich) + '</span></li>'; }).join('') + '</ul>' : '') : '';
  if (p && p.phu_hop && !p.loaded) {
    $('foot').hidden = true;
    var f = p.hinh_thuc;
    var run = function () {
      p._run = true;
      el.innerHTML = '<div class="sec" id="pbox"></div>';
      var pg = Prog($('pbox'), 'Gemini đang lập phương án ' + FN[f].toLowerCase(), STEPS_B);
      planFor(f, pg.ev).then(function () { pg.end(); if (S.form === f) drawPlan(); })
        .catch(function (e) { pg.stop(); p._run = false; el.innerHTML = '<div class="sec"><p class="warn">' + esc(e.message) + '</p><button class="btn pri" id="mk" type="button">Thử lại</button></div>'; $('mk').onclick = run; });
    };
    if (p._run) { el.innerHTML = '<div class="sec"><p class="mute">Đang lập phương án ' + FN[f].toLowerCase() + '...</p></div>'; return; }
    if (S.r.de_xuat && S.r.de_xuat.hinh_thuc === f && !p._auto) { p._auto = true; run(); return; }
    el.innerHTML = '<div class="sec"><p>Phương án ' + FN[f].toLowerCase() + ' chưa được lập chi tiết.</p><button class="btn pri" id="mk" type="button">Lập phương án ' + FN[f].toLowerCase() + '</button></div>';
    $('mk').onclick = run;
    return;
  }
  if (!p || !p.phu_hop || !p.chen.length) { el.innerHTML = '<div class="sec"><p class="mute">Hình thức này không phù hợp với bài học nên không có nội dung chèn.</p></div>'; $('foot').hidden = true; return; }
  var h = '<div class="sec"><h2><span class="n">3</span>Yêu cầu cần đạt về ATGT</h2>' +
    '<label class="chk"><input type="checkbox" id="ycOn"' + (on('yc') ? ' checked' : '') + '> Thêm vào cuối phần yêu cầu cần đạt</label>' +
    '<p class="mute" style="margin:4px 0 0">Vị trí: ngay sau đoạn "' + esc(snip(S.r.bai_hoc.neo_yccd)) + '" (cuối phần yêu cầu cần đạt)</p>' +
    '<div class="yc"><label for="y1">Học sinh thực hiện được</label><textarea id="y1" data-y="thuc_hien">' + esc(yv('thuc_hien')) + '</textarea>' +
    '<label for="y2">Vận dụng vào giải quyết vấn đề ATGT</label><textarea id="y2" data-y="van_dung">' + esc(yv('van_dung')) + '</textarea>' +
    '<label for="y3">Phẩm chất</label><textarea id="y3" data-y="pham_chat">' + esc(yv('pham_chat')) + '</textarea></div></div>';
  h += '<div class="sec"><h2><span class="n">4</span>Hoạt động tích hợp trong bài học</h2><div id="tb"></div><div id="flow" class="flow"></div>' +
    p.chen.map(function (c, i) { return card(c, i); }).join('') + '</div>';
  h += '<div class="sec"><h2><span class="n">5</span>Xem trước trong giáo án</h2><div class="doc" id="doc"></div></div>';
  el.innerHTML = h;
  el.querySelectorAll('textarea[data-y]').forEach(function (t) { t.oninput = function () { setE('yc', t.dataset.y, t.value); doc(); }; });
  $('ycOn').onchange = function () { S.acc[K('yc')] = this.checked; doc(); };
  el.querySelectorAll('[data-a]').forEach(function (c) { c.onchange = function () { S.acc[K(+c.dataset.a)] = c.checked; c.closest('.ins').style.opacity = c.checked ? 1 : .55; sum(); doc(); }; });
  el.querySelectorAll('textarea[data-e]').forEach(function (t) { t.oninput = function () { var a = t.dataset.e.split('|'); setE(+a[0], a[1], t.value); doc(); }; });
  el.querySelectorAll('textarea').forEach(function (t) { var g = function () { t.style.height = 'auto'; t.style.height = (t.scrollHeight + 2) + 'px'; }; g(); t.addEventListener('input', g); });
  $('foot').hidden = false; $('dmsg').textContent = '';
  sum(); doc();
}
function card(c, i) {
  return '<div class="ins" style="opacity:' + (on(i) ? 1 : .55) + '"><div class="hd"><label class="chk"><input type="checkbox" data-a="' + i + '"' + (on(i) ? ' checked' : '') + '> ' + esc(c.ten) + '</label><b style="color:var(--ins);white-space:nowrap">' + esc(c.thoi_luong) + ' phút</b></div>' +
    '<div class="bd"><div class="meta"><span><b>Vị trí:</b> ' + esc(c.hoat_dong) + ', sau đoạn "' + esc(snip(c.neo_doan)) + '"</span></div>' +
    '<div><span class="fl">Nội dung</span><textarea data-e="' + i + '|noi_dung">' + esc(v(i, 'noi_dung')) + '</textarea></div>' +
    '<div class="two"><div><span class="fl">Hoạt động của giáo viên</span><textarea data-e="' + i + '|gv" rows="4">' + esc(v(i, 'gv')) + '</textarea></div>' +
    '<div><span class="fl">Hoạt động của học sinh</span><textarea data-e="' + i + '|hs" rows="4">' + esc(v(i, 'hs')) + '</textarea></div></div>' +
    '<div><span class="fl">Học liệu</span><textarea data-e="' + i + '|hoc_lieu">' + esc(v(i, 'hoc_lieu')) + '</textarea></div></div></div>';
}
function sum() {
  var p = plan(), acts = S.r.cac_hoat_dong, tiet = +S.r.bai_hoc.thoi_luong_tiet || 35;
  var base = acts.reduce(function (a, x) { return a + (+x.thoi_luong || 0); }, 0);
  var add = p.chen.reduce(function (a, c, i) { return a + (on(i) ? (+c.thoi_luong || 0) : 0); }, 0);
  var mx = Math.max(tiet, base + add, 1), over = base + add > tiet;
  $('tb').innerHTML = '<div class="bar"><i style="width:' + (base / mx * 100) + '%;background:var(--accent)"></i><i style="width:' + (add / mx * 100) + '%;background:var(--ins)"></i></div>' +
    '<p class="mute" style="margin:0 0 10px">Giáo án ' + base + ' phút + tích hợp ' + add + ' phút / tiết ' + tiet + ' phút' +
    (over ? '. <span class="warn">' + esc(p.dieu_chinh_thoi_gian || 'Cần bớt thời gian ở hoạt động khác.') + '</span>' : (p.dieu_chinh_thoi_gian ? '. ' + esc(p.dieu_chinh_thoi_gian) : '')) + '</p>';
  var h = '', put0 = {};
  acts.forEach(function (a, ai) {
    h += '<div class="fr"><span class="m">' + esc(a.thoi_luong) + "'</span><div class=\"b\">" + esc(a.ten) + '</div></div>';
    p.chen.forEach(function (c, i) {
      if (put0[i]) return;
      if ((c.neo_doan >= a.doan_dau && c.neo_doan <= a.doan_cuoi) || ai === acts.length - 1) {
        put0[i] = 1;
        h += '<div class="fr i' + (on(i) ? '' : ' x') + '"><span class="m">+' + esc(c.thoi_luong) + "'</span><div class=\"b\">ATGT: " + esc(c.ten) + '</div></div>';
      }
    });
  });
  $('flow').innerHTML = h;
}
function ycLines() {
  if (!on('yc')) return [];
  var o = [], a = lines(yv('thuc_hien')), b = lines(yv('van_dung')), c = lines(yv('pham_chat'));
  if (!a.length && !b.length && !c.length) return [];
  o.push({ t: ycHead(), b: true });
  a.forEach(function (x) { o.push({ t: '- Thực hiện được: ' + x }); });
  b.forEach(function (x) { o.push({ t: '- Vận dụng: ' + x }); });
  c.forEach(function (x) { o.push({ t: '- Phẩm chất: ' + x }); });
  return o;
}
function ycHead() {
  var p = plan(), mins = p.chen.reduce(function (a, c, i) { return a + (on(i) ? (+c.thoi_luong || 0) : 0); }, 0);
  var ids = S.r.diem_giao.map(function (d) { return byId(d.ma_bai_atgt); }).filter(Boolean);
  var g = {}; ids.forEach(function (x) { (g[x.lop] = g[x.lop] || []).indexOf(x.bai) < 0 && g[x.lop].push(x.bai); });
  var tl = Object.keys(g).sort().map(function (l) { return 'lớp ' + l + ' bài ' + g[l].sort(function (m, n) { return m - n; }).join(', '); }).join('; ');
  return 'Tích hợp an toàn giao thông (' + FN[S.form].toLowerCase() + (mins && S.form !== 'tp' ? ' - ' + mins + ' phút' : '') +
    (tl ? '; Tài liệu giáo dục ATGT ' + tl : '') + ')';
}
function title() {
  var ids = S.r.diem_giao.map(function (d) { return byId(d.ma_bai_atgt); }).filter(Boolean);
  return 'Tích hợp ' + FN[S.form].toLowerCase() + ' giáo dục an toàn giao thông' + (ids.length ? ' (Tài liệu giáo dục ATGT ' + ids.map(function (x) { return 'lớp ' + x.lop + ' bài ' + x.bai; }).join(', ') + ')' : '');
}
function hasTitle() { var n = Math.min(S.paras.length, (S.r.bai_hoc.neo_ten_bai || 0) + 4); for (var i = 0; i < n; i++) if (/tích hợp/i.test(S.paras[i].text) && /(ATGT|an toàn giao thông)/i.test(S.paras[i].text)) return true; return false; }
function at() { var m = {}; plan().chen.forEach(function (c, i) { if (on(i)) (m[c.neo_doan] = m[c.neo_doan] || []).push(i); }); return m; }
function insHead(i) { var c = plan().chen[i]; return 'Tích hợp giáo dục ATGT (' + FN[S.form].toLowerCase() + ', ' + c.thoi_luong + ' phút): ' + c.ten; }
/* Vi tri chen: tach cot giao vien / hoc sinh neu giao an ke bang */
function place(ai) {
  var p = S.paras[ai]; if (!p || p.t < 0) return { split: false };
  var R = S.roles[p.t] || {}, gv = R.gv, hs = R.hs;
  if (gv == null || hs == null) { var tc = closest(S.els()[ai], 'tc'), n = tc ? kids(tc.parentNode, 'tc').length : 0; if (n === 2) { gv = 0; hs = 1; } }
  if (gv == null || hs == null || (p.c !== gv && p.c !== hs)) return { split: false };
  return { split: true, own: p.c === gv ? 'gv' : 'hs', gv: gv, hs: hs, tg: R.tg };
}
S.els = function () { return S._els || (S._els = allP(S.xml)); };
function parts(i) {
  var c = plan().chen[i];
  return { head: insHead(i), nd: v(i, 'noi_dung'), gv: lines(v(i, 'gv')), hs: lines(v(i, 'hs')), hl: v(i, 'hoc_lieu'), ten: c.ten, tl: c.thoi_luong };
}
function doc() {
  var m = at(), b = S.r.bai_hoc, yl = ycLines(), els = S.els(), afterP = {}, cellEnd = new Map();
  function addAfter(k, html) { (afterP[k] = afterP[k] || []).push(html); }
  function addCell(tc, html) { if (!tc) return; var a = cellEnd.get(tc) || []; a.push(html); cellEnd.set(tc, a); }
  function box(inner) { return '<div class="add">' + inner + '</div>'; }
  function li(a) { return a.map(function (x) { return '<p>- ' + esc(x) + '</p>'; }).join(''); }
  if (yl.length && b.neo_yccd >= 0) addAfter(b.neo_yccd, box(yl.map(function (l) { return '<p' + (l.b ? ' style="font-weight:600"' : '') + '>' + esc(l.t) + '</p>'; }).join('')));
  Object.keys(m).forEach(function (key) {
    m[key].forEach(function (i) {
      var x = parts(i), pl = place(+key), hd = '<p style="font-weight:600">' + esc(x.head) + '</p>' + (x.nd ? '<p>Nội dung: ' + esc(x.nd) + '</p>' : ''), hl = x.hl ? '<p><i>Học liệu: ' + esc(x.hl) + '</i></p>' : '';
      if (pl.split) {
        var tr = closest(els[+key], 'tc').parentNode;
        addAfter(key, box(hd + li(pl.own === 'gv' ? x.gv : x.hs) + hl));
        addCell(cellAt(tr, pl.own === 'gv' ? pl.hs : pl.gv), box('<p style="font-weight:600">Tích hợp ATGT: ' + esc(x.ten) + '</p>' + li(pl.own === 'gv' ? x.hs : x.gv)));
        if (pl.tg != null) addCell(cellAt(tr, pl.tg), box('<p>+' + esc(x.tl) + ' phút (ATGT)</p>'));
      } else {
        addAfter(key, box(hd + '<p class="mute">Giáo viên</p>' + li(x.gv) + '<p class="mute">Học sinh</p>' + li(x.hs) + hl));
      }
    });
  });
  function para(p) {
    var i = S.idx.get(p), o = S.paras[i] || {}, h = '';
    if (o.img) h += '<p class="mute">[Hình ảnh]</p>';
    if (o.text) h += '<p>' + esc(o.text) + '</p>';
    else if (!o.img) h += '<p class="empty"></p>';
    return h + (afterP[i] || []).join('');
  }
  function block(n) {
    if (n.nodeType !== 1) return '';
    if (n.localName === 'p') return para(n);
    if (n.localName === 'tbl') {
      return '<table class="dt">' + kids(n, 'tr').map(function (tr) {
        return '<tr>' + kids(tr, 'tc').map(function (tc) {
          var s = span(tc), inner = '';
          for (var c = tc.firstChild; c; c = c.nextSibling) inner += block(c);
          inner += (cellEnd.get(tc) || []).join('');
          return '<td' + (s > 1 ? ' colspan="' + s + '"' : '') + '>' + inner + '</td>';
        }).join('') + '</tr>';
      }).join('') + '</table>';
    }
    if (n.localName === 'sdt') { var o2 = ''; var sc = kids(n, 'sdtContent')[0]; if (sc) for (var c2 = sc.firstChild; c2; c2 = c2.nextSibling) o2 += block(c2); return o2; }
    return '';
  }
  var body = S.xml.getElementsByTagNameNS(W, 'body')[0], h = '';
  for (var n = body.firstChild; n; n = n.nextSibling) h += block(n);
  var d = $('doc'); d.innerHTML = h;
  var f = d.querySelector('.add'); if (f) d.scrollTop = Math.max(0, f.offsetTop - d.offsetTop - 60);
}

/* ===== Ghi file Word: chi them doan van, khong doi bang, hinh, dinh dang ===== */
function wEl(d, n) { return d.createElementNS(W, 'w:' + n); }
function mk(d, tmpl, text, o) {
  o = o || {};
  var p = wEl(d, 'p'), pp = tmpl && kids(tmpl, 'pPr')[0];
  if (pp) { pp = pp.cloneNode(true); ['numPr', 'pStyle', 'rPr', 'sectPr', 'framePr', 'pBdr', 'shd'].forEach(function (n) { kids(pp, n).forEach(function (x) { pp.removeChild(x); }); }); p.appendChild(pp); }
  var r = wEl(d, 'r'), rp = wEl(d, 'rPr'), tr = null, rs = tmpl ? tmpl.getElementsByTagNameNS(W, 'r') : [];
  for (var k = 0; k < rs.length && !tr; k++) tr = kids(rs[k], 'rPr')[0] || null;
  if (tr) ['rFonts', 'sz', 'szCs'].forEach(function (n) { var x = kids(tr, n)[0]; if (x) rp.appendChild(x.cloneNode(true)); });
  if (o.b) rp.appendChild(wEl(d, 'b'));
  if (o.i) rp.appendChild(wEl(d, 'i'));
  if (o.c) { var c = wEl(d, 'color'); c.setAttributeNS(W, 'w:val', '1F5FA8'); rp.appendChild(c); }
  r.appendChild(rp);
  var t = wEl(d, 't'); t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve'); t.textContent = text;
  r.appendChild(t); p.appendChild(r); return p;
}
function after(ref, ns) { var nx = ref.nextSibling; ns.forEach(function (n) { ref.parentNode.insertBefore(n, nx); }); return ns.length ? ns[ns.length - 1] : ref; }
function toCell(d, tc, texts) {
  if (!tc) return 0;
  var ps = kids(tc, 'p'), tmpl = ps[ps.length - 1] || null, n = 0;
  texts.forEach(function (x) { var p = mk(d, tmpl, x.t, x.o); var last = tc.lastChild; while (last && last.nodeType !== 1) last = last.previousSibling; tc.appendChild(p); n++; });
  return n;
}
function buildXml(blue) {
  var d = new DOMParser().parseFromString(S.xmlStr, 'application/xml'), b = S.r.bai_hoc, m = at(), o = allP(d), added = 0;
  var C = function (x) { return { c: blue, b: x && x.b, i: x && x.i }; };
  Object.keys(m).forEach(function (key) {
    var a = o[+key], last = a;
    m[key].forEach(function (i) {
      var x = parts(i), pl = place(+key);
      var head = [mk(d, a, x.head, C({ b: 1 }))]; if (x.nd) head.push(mk(d, a, 'Nội dung: ' + x.nd, C()));
      var hl = x.hl ? [mk(d, a, 'Học liệu: ' + x.hl, C({ i: 1 }))] : [];
      var mine = function (arr) { return arr.map(function (y) { return mk(d, a, '- ' + y, C()); }); };
      if (pl.split) {
        var tr = closest(a, 'tc').parentNode, own = pl.own === 'gv' ? x.gv : x.hs, other = pl.own === 'gv' ? x.hs : x.gv;
        var ins1 = head.concat(mine(own), hl); last = after(last, ins1); added += ins1.length;
        added += toCell(d, cellAt(tr, pl.own === 'gv' ? pl.hs : pl.gv), [{ t: 'Tích hợp ATGT: ' + x.ten, o: C({ b: 1 }) }].concat(other.map(function (y) { return { t: '- ' + y, o: C() }; })));
        if (pl.tg != null) added += toCell(d, cellAt(tr, pl.tg), [{ t: '+' + x.tl + ' phút (ATGT)', o: C() }]);
      } else {
        var ins2 = head.concat([mk(d, a, 'Hoạt động của giáo viên:', C({ b: 1 }))], mine(x.gv), [mk(d, a, 'Hoạt động của học sinh:', C({ b: 1 }))], mine(x.hs), hl);
        last = after(last, ins2); added += ins2.length;
      }
    });
  });
  var yl = ycLines();
  if (yl.length && o[b.neo_yccd]) { var y = o[b.neo_yccd]; after(y, yl.map(function (l) { return mk(d, y, l.t, C({ b: l.b })); })); added += yl.length; }
  var s0 = S.stat, s1 = count(d);
  var same = s0.tbl === s1.tbl && s0.tr === s1.tr && s0.tc === s1.tc && s0.img === s1.img && s0.sect === s1.sect && s1.p === s0.p + added;
  if (!same) throw new Error('cấu trúc giáo án bị thay đổi ngoài ý muốn, file chưa được tạo');
  return { d: d, added: added, s0: s0 };
}
var DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
function pack(buf, d) {
  return JSZip.loadAsync(buf).then(function (z) {
    z.file('word/document.xml', new XMLSerializer().serializeToString(d));
    return z.generateAsync({ type: 'blob', mimeType: DOCX, compression: 'DEFLATE' });
  });
}
function save(bl, n) {
  var a2 = document.createElement('a'); a2.href = URL.createObjectURL(bl); a2.download = n; document.body.appendChild(a2); a2.click();
  setTimeout(function () { URL.revokeObjectURL(a2.href); a2.remove(); }, 4000);
}
function build() {
  var x; try { x = buildXml($('blue').checked); } catch (e) { return Promise.reject(e); }
  var s0 = x.s0, added = x.added;
  return pack(S.buf, x.d).then(function (bl) {
    var n = (S.name || 'Giao an') + ' - tich hop ATGT.docx'; save(bl, n);
    $('dmsg').textContent = 'Đã tải: ' + n + '. Giữ nguyên ' + s0.tbl + ' bảng, ' + s0.tc + ' ô, ' + s0.img + ' hình ảnh; chỉ thêm ' + added + ' đoạn tích hợp.';
  });
}

/* ===== Nhieu giao an: tu dong phan tich, chen va xuat file, khong can duyet ===== */
var ONE = obj({ bai_hoc: SCHEMA.properties.bai_hoc, cac_hoat_dong: SCHEMA.properties.cac_hoat_dong, diem_giao: SCHEMA.properties.diem_giao, de_xuat: SCHEMA.properties.de_xuat, chi_tiet: PLAN });
var CK = ['paras', 'zip', 'xml', 'buf', 'xmlStr', 'roles', 'idx', 'stat', '_els', 'name', 'r', 'form', 'acc', 'ed'];
var MAXF = 10, PAR = 3;
var B = { files: [], run: false, out: [] };
/* Doi tam du lieu S sang giao an dang xu li (chi chay dong bo nen khong lan giua cac file) */
function withCtx(c, fn) {
  var keep = {}; CK.forEach(function (k) { keep[k] = S[k]; S[k] = c[k]; });
  try { return fn(); } finally { CK.forEach(function (k) { c[k] = S[k]; S[k] = keep[k]; }); }
}
function snap() { var c = {}; CK.forEach(function (k) { c[k] = S[k]; }); return c; }
function oneUser() {
  return 'Xử lí tự động: phân tích và lập chi tiết luôn trong một lần trả lời.\n' +
    'bai_hoc, cac_hoat_dong, diem_giao, de_xuat: như khi phân tích tổng quan.\n' +
    'chi_tiet: phương án chi tiết (hinh_thuc, phu_hop, ly_do, yccd, dieu_chinh_thoi_gian, nguyen_tac, chen) cho đúng hình thức đã chọn ở de_xuat. Nếu de_xuat là "khong" thì chi_tiet.chen để rỗng.\n' +
    'KẾ HOẠCH BÀI DẠY:\n' + numbered();
}
function fromOne(raw) {
  var dx = raw.de_xuat || {}, ct = raw.chi_tiet;
  raw.phuong_an = dx.hinh_thuc && dx.hinh_thuc !== 'khong' && ct ? [Object.assign(ct, { hinh_thuc: dx.hinh_thuc, phu_hop: true, ly_do: ct.ly_do || dx.ly_do })] : [];
  delete raw.chi_tiet;
  return clean(raw);
}
function tlOf(r) {
  var ids = (r.diem_giao || []).map(function (d) { return byId(d.ma_bai_atgt); }).filter(Boolean);
  return ids.map(function (x) { return 'lớp ' + x.lop + ' bài ' + x.bai; }).filter(function (x, i, a) { return a.indexOf(x) === i; }).join(', ');
}
function bRow(i) { return $('br' + i); }
function bSet(i, st, cls) { var el = bRow(i); if (!el) return; el.querySelector('.bs').innerHTML = st; el.className = 'bi' + (cls ? ' ' + cls : ''); }
function bStage(t) {
  var k = [['chen', 'Thiết kế hoạt động tích hợp'], ['nguyen_tac', 'Kiểm tra 6 nguyên tắc'], ['yccd', 'Viết yêu cầu cần đạt ATGT'], ['chi_tiet', 'Lập phương án chi tiết'], ['de_xuat', 'Chọn hình thức tích hợp'], ['diem_giao', 'Đối chiếu Tài liệu giáo dục ATGT'], ['bai_hoc', 'Nhận diện bài học']];
  for (var i = 0; i < k.length; i++) if (t.indexOf('"' + k[i][0] + '"') >= 0) return k[i][1];
  return 'Gemini đang đọc giáo án';
}
function bDraw() {
  var box = $('blist'); if (!box) return;
  if (!B.files.length) { box.innerHTML = ''; $('bhint').textContent = ''; return; }
  box.innerHTML = B.files.map(function (f, i) {
    return '<li class="bi" id="br' + i + '"><div class="bm"><b class="bn">' + esc(f.name) + '</b><span class="bs mute">' + (f.bad ? '<span class="warn">' + esc(f.bad) + '</span>' : 'Chờ xử lí') + '</span></div><span class="ba"></span></li>';
  }).join('');
  var nOk = B.files.filter(function (f) { return !f.bad; }).length; $('bhint').textContent = nOk + ' giáo án' + (B.files.length >= MAXF ? ' (đã đủ tối đa ' + MAXF + ' file)' : '');
}
function bAdd(list) {
  if (B.run) return;
  if (B.out.length) { B.files = []; B.out = []; $('bres').innerHTML = ''; }
  Array.prototype.forEach.call(list || [], function (f) {
    if (B.files.length >= MAXF) return;
    if (B.files.some(function (x) { return x.name === f.name && x.size === f.size; })) return;
    B.files.push({ file: f, name: f.name, size: f.size, bad: /\.docx$/i.test(f.name) ? '' : 'Chỉ nhận .docx (mở bằng Word, Lưu thành .docx)' });
  });
  bDraw();
}
function bMsg(t, bad) { $('bmsg').innerHTML = bad ? '<span class="warn">' + esc(t) + '</span>' : esc(t); }
function uniq(n, used) { var b = n, k = 2; while (used[n]) n = b.replace(/\.docx$/, '') + ' (' + (k++) + ').docx'; used[n] = 1; return n; }
function runOne(f, i, blue) {
  var c = f.ctx, t0 = Date.now(), raw;
  var tick = function (e) { var t = (e && e.text) || ''; bSet(i, '<span class="spin"></span>' + esc(bStage(t)) + (t ? ' · đã nhận ' + t.length + ' ký tự' : '') + (e && e.model ? ' <span class="mute">(' + esc(e.model) + ')</span>' : '')); };
  bSet(i, '<span class="spin"></span>Gửi tới Gemini');
  var user = withCtx(c, oneUser);
  var call = window.__MOCK_ONE ? window.__MOCK_ONE(f.name, tick) : gen(user, ONE, tick, 150000);
  var again = function (e) { if (window.__MOCK_ONE || !(e.status === 0 || e.status === 429 || e.status >= 500)) throw e; bSet(i, '<span class="spin"></span>Gemini bận, thử lại sau 20 giây'); return wait(20000).then(function () { return gen(user, ONE, tick, 150000); }); };
  return call.catch(again).then(function (j) {
    raw = j;
    var res = withCtx(c, function () {
      S.r = fromOne(raw); S.acc = {}; S.ed = {};
      var p = S.r.phuong_an[0];
      if (!p || (!p.chen.length && !p.yccd.thuc_hien.length)) return { skip: (S.r.de_xuat && S.r.de_xuat.ly_do) || 'Không có điểm giao tự nhiên với nội dung ATGT.' };
      S.form = p.hinh_thuc;
      var x = buildXml(blue), mins = p.chen.reduce(function (a, q) { return a + (+q.thoi_luong || 0); }, 0);
      return { x: x, form: p.hinh_thuc, n: p.chen.length, mins: mins, ten: S.r.bai_hoc.ten_bai || '', tl: tlOf(S.r) };
    });
    if (res.skip) { bSet(i, '<span class="warn">Giữ nguyên, không tích hợp:</span> ' + esc(res.skip), 'skip'); return null; }
    return pack(c.buf, res.x.d).then(function (bl) {
      var sec = Math.round((Date.now() - t0) / 1000), n = c.name + ' - tich hop ATGT.docx';
      var o = { i: i, bl: bl, n: n }; B.out.push(o);
      bSet(i, '<span class="ok">Xong</span> · ' + esc(FN[res.form]) + (res.form !== 'tp' && res.mins ? ' ' + res.mins + ' phút' : '') + ' · ' + res.n + ' hoạt động tích hợp' + (res.tl ? ' · Tài liệu ATGT ' + esc(res.tl) : '') + ' · ' + sec + ' giây' + (res.ten ? '<br><span class="mute">' + esc(res.ten) + '</span>' : ''), 'done');
      var a = bRow(i).querySelector('.ba'); a.innerHTML = '<button class="btn" type="button" style="min-height:34px;padding:5px 12px">Tải</button>';
      a.firstChild.onclick = function () { if (window.__PREVIEW) { bMsg('Bản xem trước không tải file được.'); return; } save(bl, o.n); };
      return o;
    });
  }).catch(function (e) { bSet(i, '<span class="warn">Lỗi: ' + esc(e.message) + '</span>', 'err'); return null; });
}
function runBatch() {
  if (B.run) return;
  var ok = B.files.filter(function (f) { return !f.bad; });
  if (!ok.length) { bMsg('Chọn các file giáo án .docx trước.', true); return; }
  if (!window.__MOCK_ONE && !get(KEYK)) { bMsg('Dán và lưu khóa Gemini trước.', true); return; }
  var blue = $('bblue').checked, bt = $('bgo'), t0 = Date.now(), keep = snap(), done = 0, used = {};
  B.run = true; B.out = []; bt.disabled = true; bt.innerHTML = '<span class="spin"></span>Đang tạo giáo án'; bMsg(''); $('bres').innerHTML = ''; bDraw();
  var total = ok.length;
  var iv = setInterval(function () { $('bprog').innerHTML = '<div class="bar"><i style="width:' + Math.round(done / total * 100) + '%;background:var(--accent)"></i></div><span class="mute">Đã xong ' + done + '/' + total + ' · ' + Math.round((Date.now() - t0) / 1000) + ' giây</span>'; }, 500);
  /* Doc lan luot tung file tren may, giu rieng du lieu cua moi file */
  var rd = Promise.resolve();
  B.files.forEach(function (f, i) {
    if (f.bad) return;
    rd = rd.then(function () {
      bSet(i, '<span class="spin"></span>Đọc file trên máy');
      return readDocx(f.file).then(function () { S.name = f.name.replace(/\.docx$/i, ''); f.ctx = snap(); bSet(i, 'Đã đọc ' + S.paras.filter(function (p) { return p.text; }).length + ' đoạn, ' + S.stat.tbl + ' bảng, ' + S.stat.img + ' hình ảnh · chờ Gemini'); })
        .catch(function (e) { f.ctx = null; bSet(i, '<span class="warn">' + esc(e.message) + '</span>', 'err'); });
    });
  });
  return rd.then(function () {
    CK.forEach(function (k) { S[k] = keep[k]; });
    var q = B.files.map(function (f, i) { return i; }).filter(function (i) { return B.files[i].ctx && !B.files[i].bad; });
    total = q.length || 1;
    if (!q.length) return;
    function worker() { if (!q.length) return Promise.resolve(); var i = q.shift(); return runOne(B.files[i], i, blue).then(function () { done++; return worker(); }); }
    var ws = [], nw = Math.min(PAR, q.length); for (var w = 0; w < nw; w++) ws.push(worker());
    return Promise.all(ws);
  }).then(function () {
    clearInterval(iv);
    var sec = Math.round((Date.now() - t0) / 1000), out = B.out.slice().sort(function (a, b) { return a.i - b.i; });
    $('bprog').innerHTML = '<div class="bar"><i style="width:100%;background:var(' + (out.length ? '--ok' : '--warn') + ')"></i></div><span class="mute">Hoàn thành sau ' + sec + ' giây · tạo được ' + out.length + '/' + ok.length + ' giáo án</span>';
    if (!out.length) { bMsg('Chưa tạo được giáo án nào. Xem lí do ở từng dòng.', true); return; }
    var z = new JSZip(); out.forEach(function (o) { z.file(uniq(o.n, used), o.bl); });
    var zn = 'Giao an tich hop ATGT (' + out.length + ' bai).zip';
    return z.generateAsync({ type: 'blob', compression: 'STORE' }).then(function (zb) {
      $('bres').innerHTML = '<button class="btn pri" id="bzip" type="button">Tải tất cả (.zip)</button> <span class="mute">' + esc(zn) + '</span>';
      var get1 = function () { if (window.__PREVIEW) { bMsg('Bản xem trước không tải file được; bản đăng chính thức tự tải về file .zip chứa các giáo án đã tích hợp.'); return; } save(zb, zn); };
      $('bzip').onclick = get1;
      get1();
      if (!window.__PREVIEW) bMsg('Đã tự động tải ' + zn + '. Nếu trình duyệt chặn, bấm "Tải tất cả" hoặc nút Tải ở từng giáo án.');
    });
  }).catch(function (e) { clearInterval(iv); bMsg(e.message, true); })
    .then(function () { CK.forEach(function (k) { S[k] = keep[k]; }); B.run = false; bt.disabled = false; bt.textContent = 'Tạo giáo án tích hợp'; });
}
function setMode(m) {
  $('mode').querySelectorAll('button').forEach(function (b) { b.classList.toggle('on', b.dataset.m === m); });
  $('one').hidden = m !== 'one'; $('many').hidden = m !== 'many';
  $('foot').hidden = m !== 'one' || !S.r || !(plan() && plan().loaded);
  try { localStorage.setItem('atgt_mode', m); } catch (e) {}
}
$('mode').querySelectorAll('button').forEach(function (b) { b.onclick = function () { setMode(b.dataset.m); }; });
$('bfile').onchange = function (e) { bAdd(e.target.files); e.target.value = ''; };
var bz = $('bdrop');
['dragenter', 'dragover'].forEach(function (ev) { bz.addEventListener(ev, function (e) { e.preventDefault(); bz.classList.add('on'); }); });
['dragleave', 'drop'].forEach(function (ev) { bz.addEventListener(ev, function (e) { e.preventDefault(); bz.classList.remove('on'); }); });
bz.addEventListener('drop', function (e) { bAdd(e.dataTransfer.files); });
$('bclr').onclick = function () { if (B.run) return; B.files = []; B.out = []; $('bres').innerHTML = ''; $('bprog').innerHTML = ''; bMsg(''); bDraw(); };
$('bgo').onclick = runBatch;
(function () { var m = ''; try { m = localStorage.getItem('atgt_mode') || ''; } catch (e) {} if (m === 'many') setMode('many'); })();

/* ===== Khoi dong ===== */
function msg(t, bad) { $('msg').innerHTML = bad ? '<span class="warn">' + esc(t) + '</span>' : esc(t); }
function load(f) {
  if (!f) return;
  if (!/\.docx$/i.test(f.name)) { msg('Chỉ nhận file .docx. File .doc: mở bằng Word, Lưu thành .docx.', true); return; }
  S.name = f.name.replace(/\.docx$/i, ''); $('dropT').textContent = f.name; $('dropS').textContent = 'Đang đọc...';
  readDocx(f).then(function (a) { $('dropS').textContent = a.filter(function (p) { return p.text; }).length + ' đoạn · bấm để chọn file khác'; msg(''); })
    .catch(function (e) { S.paras = []; $('dropS').textContent = 'Không đọc được file'; msg(e.message, true); });
}
$('file').onchange = function (e) { load(e.target.files[0]); };
var dz = $('drop');
['dragenter', 'dragover'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('on'); }); });
['dragleave', 'drop'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('on'); }); });
dz.addEventListener('drop', function (e) { load(e.dataTransfer.files[0]); });
$('go').onclick = function () {
  if (!S.paras.length) { msg('Chọn file giáo án trước.', true); return; }
  if (!get(KEYK)) { msg('Dán và lưu khóa Gemini trước.', true); return; }
  var bt = $('go');
  bt.disabled = true; bt.innerHTML = '<span class="spin"></span>Đang phân tích'; msg('');
  $('out').innerHTML = '<div class="sec" id="pbox"></div>'; $('foot').hidden = true;
  var pg = Prog($('pbox'), 'Gemini đang phân tích giáo án', STEPS_A);
  analyze(pg.ev).then(function (r) { var sec = pg.end(); show(r); msg('Phân tích xong sau ' + sec + ' giây.'); })
    .catch(function (e) { pg.stop(); $('out').innerHTML = ''; msg(e.message, true); })
    .then(function () { bt.disabled = false; bt.textContent = 'Phân tích lại'; });
};
$('dl').onclick = function () { build().catch(function (e) { $('dmsg').textContent = 'Không tạo được file: ' + e.message; }); };
function show(r) {
  S.r = r; S.acc = {}; S.ed = {};
  var rec = r.de_xuat && r.de_xuat.hinh_thuc, ok = r.phuong_an.filter(function (p) { return p.phu_hop; });
  S.form = r.phuong_an.some(function (p) { return p.hinh_thuc === rec; }) ? rec : (ok[0] || r.phuong_an[0] || {}).hinh_thuc;
  if (!S.form) { $('out').innerHTML = '<div class="sec"><p class="warn">Không có phương án tích hợp.</p></div>'; return; }
  result();
}
function examples() {
  var V = window.VIDU || [], box = $('vd'); if (!V.length || !box) return;
  box.innerHTML = '<span class="mute">Xem ví dụ:</span> ' + V.map(function (e, i) { return '<button class="btn" type="button" data-v="' + i + '" style="min-height:34px;padding:5px 12px">' + esc(e.nhan) + ' · ' + esc(e.bai.replace(/^Đạo đức 1 – /, 'Đạo đức 1, ')) + '</button>'; }).join(' ');
  box.querySelectorAll('[data-v]').forEach(function (bt) { bt.onclick = function () {
    var e = V[+bt.dataset.v]; msg('Đang mở ví dụ...');
    fetch(e.file).then(function (r) { if (!r.ok) throw new Error('Không mở được file ví dụ.'); return r.blob(); })
      .then(function (bl) { var f = new File([bl], e.file.split('/').pop()); S.name = f.name.replace(/\.docx$/i, ''); $('dropT').textContent = e.bai; $('dropS').textContent = 'Ví dụ ' + e.nhan.toLowerCase() + ' · bấm để chọn giáo án của thầy cô'; return readDocx(f); })
      .then(function () { show(clean(JSON.parse(JSON.stringify(e.r)))); msg(''); $('out').scrollIntoView({ behavior: 'smooth' }); })
      .catch(function (er) { msg(er.message, true); });
  }; });
}
window.ATGT = { Prog: Prog, STEPS_A: STEPS_A, S: S, show: show, readDocx: readDocx, B: B, bAdd: bAdd, runBatch: runBatch, setMode: setMode };
examples();
drawKey();
})();
