'use strict';
// ============================================================ tiện ích
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const stCls = (s) => (s === 'Chưa test' ? 'Chua' : s || '');
const modeCls = (m) => ({ 'Tự động': 'm-auto', 'Bán tự động': 'm-semi', 'Tay': 'm-manual' }[m] || 'm-auto');
const modeLabel = (m) => ({ 'Tự động': 'Tự động', 'Bán tự động': '⚠ Bán tự động', 'Tay': '✋ Test tay' }[m] || m);

async function api(path, body) {
  const opt = body === undefined ? {} : { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Tool': '1' }, body: JSON.stringify(body) };
  let r;
  try { r = await fetch(path, opt); } catch (e) { throw new Error('Mất kết nối với tool (tool đã tắt?). Hãy mở lại tool.'); }
  const d = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(d.error || ('Lỗi ' + r.status));
  return d;
}
function toast(msg, err) {
  const t = $('toast'); t.textContent = msg; t.className = 'toast show' + (err ? ' err' : '');
  clearTimeout(toast._t); toast._t = setTimeout(() => (t.className = 'toast'), err ? 6000 : 2600);
}
async function guard(fn) { try { return await fn(); } catch (e) { toast(e.message, true); } }

// hộp thoại chung — trả về Promise với giá trị nút được bấm
function dialog({ title, html, buttons = [{ label: 'OK', value: true, cls: 'primary' }], warn = false, cls = '' }) {
  return new Promise((resolve) => {
    const box = $('modal-box');
    box.className = 'modal-box ' + cls;
    box.innerHTML = `<div class="mh ${warn ? 'warn' : ''}"><h2>${title}</h2></div><div class="mb">${html}</div>
      <div class="mf">${buttons.map((b, i) => `<button class="btn ${b.cls || ''}" data-i="${i}">${b.label}</button>`).join('')}</div>`;
    $('modal').hidden = false;
    box.querySelectorAll('.mf button').forEach((b) => (b.onclick = () => {
      $('modal').hidden = true; resolve(buttons[+b.dataset.i].value);
    }));
    const prim = box.querySelector('.mf .primary'); if (prim) prim.focus();
  });
}

// ============================================================ trạng thái
const S = {
  state: null, wb: null, sel: new Set(), open: new Set(), closedGroups: new Set(), kind: 'testcase',
  profile: null, profileDirty: false, poll: null, since: 0, liveKeys: new Set(), manualReq: null, manualShownFor: null,
  runId: null, result: null, rFilter: '', expandAll: false,
};

async function loadState() {
  const st = await api('/api/state');
  S.state = st;
  S.profile = JSON.parse(JSON.stringify(st.profile));
  applySettings(st.settings);
  setWorkbook(st.workbook, true);
  renderGuide();
  renderColumns();
  if (st.run && ['đang chạy', 'chờ bạn kiểm tra', 'chuẩn bị'].includes(st.run.state)) {
    S.since = 0; showProgress(true); startPoll();
  }
  refreshRuns();
}

function applySettings(s) {
  $('o-browser').value = s.browser || 'edge';
  $('o-speed').value = s.speed || 'vua';
  $('o-headless').checked = !!s.headless;
  $('o-shot').checked = !!s.shot_every_step;
  $('o-video').checked = !!s.video;
  $('s-tester').value = s.tester || '';
  $('s-project').value = s.project_name || '';
  $('s-size').value = s.size || '1366x768';
  $('s-timeout').value = s.timeout || 10;
  $('s-autofirst').checked = s.auto_first !== false;
  $('s-stoponfail').checked = s.stop_on_fail !== false;
  const x = s.export || {};
  $('x-pattern').value = x.pattern || '';
  $('x-folder').value = x.folder || '';
  $('x-html').checked = !!x.html; $('x-embed').checked = !!x.embed; $('x-writeback').checked = !!x.writeback;
  $('x-open').checked = x.open_after !== false;
  $('x-round').value = x.round || '1';
  $('o-baseurl').value = s.base_url || '';
  $('s-demo').textContent = S.state.demo_url + '/login.html';
  $('s-data').textContent = S.state.data_dir;
  $('s-version').textContent = `${S.state.app} — phiên bản ${S.state.version}`;
  const rec = (s.recent || []).filter(Boolean);
  $('recent').innerHTML = rec.length ? '<b>Mở gần đây:</b>' + rec.map((p) => `<a class="item" data-path="${esc(p)}">📄 ${esc(p)}</a>`).join('') : '';
  $('recent').querySelectorAll('.item').forEach((a) => (a.onclick = () => openFile(a.dataset.path)));
}

// ============================================================ file Excel
function setWorkbook(wb, keepSel) {
  const old = S.wb;
  S.wb = wb;
  $('empty-state').hidden = !!wb;
  $('run-layout').hidden = !wb;
  $('filebar').hidden = !wb || $('tab-run').hidden;
  if (!wb) return;
  $('file-name').textContent = wb.file_name;
  $('file-name').title = wb.file;
  const nm = wb.testcases.filter((t) => t.mode !== 'Tự động').length;
  $('file-meta').textContent = `${wb.scenarios.length} scenario · ${wb.testcases.length} test case` + (nm ? ` (${nm} cần người kiểm tra)` : '');
  $('file-modified').hidden = !wb.modified;
  const keys = new Set(wb.testcases.map((t) => t.key));
  if (keepSel && old && old.file === wb.file) S.sel = new Set([...S.sel].filter((k) => keys.has(k)));
  else S.sel = new Set();
  // môi trường
  const env = $('o-env');
  env.innerHTML = wb.env_names.map((n) => `<option>${esc(n)}</option>`).join('');
  if (wb.env_names.includes(S.state.settings.env)) env.value = S.state.settings.env;
  showEnvVars();
  // bộ lọc
  const uniq = (k) => [...new Set(wb.testcases.map((t) => t.fields[k]).filter(Boolean))].sort();
  fillSel('f-module', 'Module: tất cả', uniq('module'));
  fillSel('f-type', 'Loại test: tất cả', uniq('loai_test'));
  fillSel('f-prio', 'Priority: tất cả', uniq('priority'));
  // cảnh báo của file
  const warns = [...wb.warnings];
  wb.testcases.forEach((t) => (t.warnings || []).forEach((w) => warns.push(`<b>${esc(t.id)}</b>: ${esc(w)}`)));
  $('wb-warn').hidden = !warns.length;
  $('wb-warn').innerHTML = '<b>⚠ Lưu ý khi đọc file:</b><br>' + warns.map((w) => (w.includes('<b>') ? w : esc(w))).join('<br>');
  renderTree();
  renderMapping();
}
function fillSel(id, first, vals) {
  const el = $(id), cur = el.value;
  el.innerHTML = `<option value="">${first}</option>` + vals.map((v) => `<option>${esc(v)}</option>`).join('');
  if (vals.includes(cur)) el.value = cur;
}
function showEnvVars() {
  const vars = (S.wb && S.wb.envs[$('o-env').value]) || {};
  const shown = Object.entries(vars).map(([k, v]) => `<div><code>\${${esc(k)}}</code> = ${esc(/pass|mat ?khau|password/i.test(k) ? '••••' : (v || '(trống)'))}</div>`);
  $('env-vars').innerHTML = shown.join('') || '<i>Không có sheet Env — chỉ dùng biến có sẵn.</i>';
}

async function openFile(path) {
  const r = await guard(() => api('/api/open', path ? { path } : {}));
  if (!r || r.cancel) return;
  S.state.settings.env = r.workbook.env_names[0];
  setWorkbook(r.workbook, false);
  toast('Đã mở ' + r.workbook.file_name);
  switchTab('run');
}
async function reloadFile() {
  const r = await guard(() => api('/api/reload', {}));
  if (r) { setWorkbook(r.workbook, true); toast('Đã nạp lại file'); }
}
async function makeTemplate(sample) {
  const r = await guard(() => api('/api/template', { sample }));
  if (!r || r.cancel) return;
  if (r.workbook) { setWorkbook(r.workbook, false); switchTab('run'); }
  const go = await dialog({
    title: 'Đã tạo file mẫu', html: `<p>Đã lưu: <b>${esc(r.path)}</b></p>` + (sample ?
      '<p>File có sẵn ví dụ cho web demo <b>Đơn hàng tuyển dụng</b> (đã mở trong tool). Tick vài TC rồi bấm <b>Chạy test</b> để xem tool hoạt động.</p>' +
      '<p class="muted">TC_DHTD_004_02 sẽ <b>Fail</b> vì web demo cố ý có lỗi — để bạn xem tool báo lỗi thế nào.</p>' :
      '<p>Mở file bằng Excel, điền test case theo các cột, lưu lại rồi mở file trong tool.</p>'),
    buttons: [{ label: 'Mở file bằng Excel', value: 'open' }, { label: 'OK', value: false, cls: 'primary' }],
  });
  if (go === 'open') guard(() => api('/api/open_path', { path: r.path }));
}

// kéo-thả file
function setupDrop() {
  const zone = $('drop');
  ['dragenter', 'dragover'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); if (ev === 'drop' || e.target === zone) zone.classList.remove('over'); }));
  document.addEventListener('drop', (e) => {
    if (!$('manual').hidden || !$('fillm').hidden) return;
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (!f) return;
    if (!/\.xlsx?m?$/i.test(f.name)) return toast('Chỉ nhận file Excel .xlsx', true);
    const rd = new FileReader();
    rd.onload = async () => {
      const data = String(rd.result).split(',')[1];
      const r = await guard(() => api('/api/upload', { name: f.name, data }));
      if (r) {
        setWorkbook(r.workbook, false); switchTab('run');
        dialog({ title: 'Đã mở file', html: `<p>Đã mở <b>${esc(f.name)}</b>.</p><p class="muted">File kéo-thả được chép vào thư mục của tool. Khi sửa file gốc trong Excel, hãy <b>kéo-thả lại</b> hoặc dùng nút <b>Mở file Excel</b> (nút này nhớ đúng file gốc để Nạp lại).</p>` });
      }
    };
    rd.readAsDataURL(f);
  });
}

// ============================================================ cây chọn test
function visibleTCs() {
  const q = $('f-search').value.trim().toLowerCase();
  const fm = $('f-module').value, ft = $('f-type').value, fp = $('f-prio').value, fmode = $('f-mode').value;
  return S.wb.testcases.filter((t) => {
    const f = t.fields;
    if (fm && f.module !== fm) return false;
    if (ft && f.loai_test !== ft) return false;
    if (fp && f.priority !== fp) return false;
    if (fmode && t.mode !== fmode) return false;
    if (q) {
      const hay = [t.id, f.tieu_de, f.module, f.scenario_id, f.tag, f.loai_test].join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

function renderTree() {
  if (!S.wb) return;
  const vis = visibleTCs();
  const visKeys = new Set(vis.map((t) => t.key));
  const groups = [];
  const used = new Set();
  for (const s of S.wb.scenarios) {
    const tcs = s.tcs.map((k) => S.wb.testcases.find((t) => t.key === k)).filter((t) => t && visKeys.has(t.key));
    tcs.forEach((t) => used.add(t.key));
    if (tcs.length || !$('f-search').value) groups.push({ id: s.id, title: s.fields.scenario || '', meta: [s.fields.nhom, s.fields.priority, s.fields.ac].filter(Boolean).join(' · '), tcs });
  }
  const rest = vis.filter((t) => !used.has(t.key));
  if (rest.length) groups.push({ id: '', title: S.wb.scenarios.length ? 'Test case không thuộc scenario nào' : 'Tất cả test case', meta: '', tcs: rest });
  const html = groups.map((g) => {
    const n = g.tcs.length, ns = g.tcs.filter((t) => S.sel.has(t.key)).length;
    const gid = g.id || '__rest__';
    const closed = S.closedGroups.has(gid);
    return `<div class="grp ${closed ? 'closed' : ''}" data-g="${esc(gid)}">
      <div class="grp-head"><span class="caret">${closed ? '▸' : '▾'}</span>
        <input type="checkbox" class="gchk" ${n && ns === n ? 'checked' : ''} ${n ? '' : 'disabled'} data-indet="${ns > 0 && ns < n ? 1 : 0}">
        ${g.id ? `<span class="sid">${esc(g.id)}</span>` : ''}<span class="stitle">${esc(g.title)}</span>
        <span class="cnt">${g.meta ? esc(g.meta) + ' · ' : ''}${ns}/${n} TC</span></div>
      <div class="grp-body">${g.tcs.map(tcHtml).join('') || '<div class="muted small" style="padding:8px 28px">Scenario chưa có test case nào (cột Scenario ID trong sheet TestCase).</div>'}</div></div>`;
  }).join('');
  $('tree').innerHTML = html || '<p class="muted">Không có test case khớp bộ lọc.</p>';
  $('tree').querySelectorAll('.gchk').forEach((c) => (c.indeterminate = c.dataset.indet === '1'));
  const allVis = vis.length && vis.every((t) => S.sel.has(t.key));
  $('chk-all').checked = !!allVis;
  $('chk-all').indeterminate = !allVis && vis.some((t) => S.sel.has(t.key));
  updateSummary();
}

function tcHtml(t) {
  const f = t.fields;
  const sizes = (f.kich_thuoc || '').split(/[,;\n]+/).filter((x) => x.trim()).length;
  const open = S.open.has(t.key) || S.expandAll;
  const warn = (t.warnings || []).length;
  return `<div class="tc ${S.sel.has(t.key) ? 'sel' : ''} ${open ? 'open' : ''}" data-k="${esc(t.key)}">
    <div class="tc-row"><input type="checkbox" class="tchk" ${S.sel.has(t.key) ? 'checked' : ''}>
      <span class="tid">${esc(t.id)}</span><span class="ttitle">${esc(f.tieu_de || '')}</span>
      ${warn ? `<span class="wicon" title="${esc(t.warnings.join('\n'))}">⚠ ${warn}</span>` : ''}
      ${f.loai_test ? `<span class="tag">${esc(f.loai_test)}</span>` : ''}${f.priority ? `<span class="tag">${esc(f.priority)}</span>` : ''}
      ${sizes > 1 ? `<span class="tag" title="${esc(f.kich_thuoc)}">🖥 ×${sizes}</span>` : ''}
      <span class="meta">${t.steps.length} bước</span><span class="mode ${modeCls(t.mode)}" title="${esc(t.mode_reason || '')}">${modeLabel(t.mode)}</span></div>
    <div class="tc-detail">${open ? tcDetail(t) : ''}</div></div>`;
}

function tcDetail(t) {
  const f = t.fields;
  let h = '';
  if (t.mode !== 'Tự động') {
    h += `<div class="reason ${t.mode === 'Tay' ? 'manual' : ''}"><b>${modeLabel(t.mode)}:</b> ${esc(t.mode_reason || '')}. ${t.mode === 'Tay' ? 'Tool KHÔNG tự đánh giá — bạn tự kiểm tra rồi chọn Pass/Fail.' : 'Tool chạy phần tự động, đến bước cần người thì DỪNG và hỏi bạn.'}</div>`;
  }
  (t.warnings || []).forEach((w) => (h += `<div class="reason">⚠ ${esc(w)}</div>`));
  const info = [['Precondition', f.precondition], ['Test data', f.test_data], ['Kích thước', f.kich_thuoc], ['Trình duyệt', f.trinh_duyet]].filter((x) => x[1]);
  if (info.length) h += '<div class="small">' + info.map(([k, v]) => `<b>${k}:</b> ${esc(v)}`).join(' · ') + '</div>';
  h += `<table class="steps"><tr><th>#</th><th>Bước thực hiện</th><th>Expected</th><th>Hành động</th><th>Đối tượng</th><th>Giá trị</th></tr>` +
    t.steps.map((s) => `<tr class="${s.kind === 'manual' ? 'st-manual' : s.kind === 'invalid' ? 'st-invalid' : ''}"><td>${esc(s.no)}</td>
      <td>${esc(s.fields.buoc || '')}</td><td>${esc(s.fields.expected || '')}</td>
      <td class="act">${s.kind === 'manual' ? '✋ ' + esc(s.action || 'test tay') : s.kind === 'invalid' ? '❌ ' + esc(s.fields.hanh_dong) : esc(s.action)}</td>
      <td>${esc(s.fields.doi_tuong || '')}</td><td>${esc(s.fields.gia_tri || s.fields.du_lieu_buoc || '')}</td></tr>`).join('') + '</table>';
  return h;
}

function updateSummary() {
  if (!S.wb) { $('btn-run').disabled = true; return; }
  const sel = S.wb.testcases.filter((t) => S.sel.has(t.key));
  const c = { 'Tự động': 0, 'Bán tự động': 0, 'Tay': 0 };
  let runs = 0;
  sel.forEach((t) => {
    c[t.mode]++;
    const n = Math.max(1, (t.fields.kich_thuoc || '').split(/[,;\n]+/).filter((x) => x.trim()).length) *
      Math.max(1, (t.fields.trinh_duyet || '').split(/[,;\n/]+/).filter((x) => /edge|chrome|firefox|coc/i.test(x)).length);
    runs += n;
  });
  $('sel-summary').innerHTML = `<div>Đã chọn <b class="n">${sel.length}</b> test case${runs > sel.length ? ` (${runs} lượt chạy)` : ''}</div>
    <div class="line"><span class="mode m-auto">Tự động</span><b>${c['Tự động']}</b></div>
    <div class="line"><span class="mode m-semi">⚠ Bán tự động</span><b>${c['Bán tự động']}</b></div>
    <div class="line"><span class="mode m-manual">✋ Test tay</span><b>${c['Tay']}</b></div>
    ${c['Bán tự động'] + c['Tay'] ? '<div class="small warn-text" style="margin-top:6px">Có case cần bạn kiểm tra — tool sẽ dừng lại và hỏi.</div>' : ''}`;
  const running = S.poll !== null;
  $('btn-run').disabled = !sel.length || running;
  $('btn-run').textContent = running ? '⏳ Đang chạy…' : `▶ Chạy ${sel.length || ''} test`;
}

// cảnh báo NGAY KHI TICK case không tự động hoàn toàn
async function warnManual(keys) {
  const tcs = keys.map((k) => S.wb.testcases.find((t) => t.key === k)).filter((t) => t && t.mode !== 'Tự động');
  if (!tcs.length) return;
  const tay = tcs.filter((t) => t.mode === 'Tay'), semi = tcs.filter((t) => t.mode === 'Bán tự động');
  let html = '';
  if (tay.length) html += `<p><span class="mode m-manual">✋ TEST TAY</span> — tool <b>KHÔNG tự đánh giá</b>. Khi chạy tới, tool sẽ dừng và bạn phải <b>tự kiểm tra rồi chọn Pass/Fail</b>:</p><ul>${tay.map((t) => `<li><b>${esc(t.id)}</b> ${esc(t.fields.tieu_de || '')} <span class="muted small">— ${esc(t.mode_reason)}</span></li>`).join('')}</ul>`;
  if (semi.length) html += `<p><span class="mode m-semi">⚠ BÁN TỰ ĐỘNG</span> — tool chạy phần tự động, đến bước cần người thì <b>dừng lại để bạn duyệt</b>:</p><ul>${semi.map((t) => `<li><b>${esc(t.id)}</b> ${esc(t.fields.tieu_de || '')} <span class="muted small">— ${esc(t.mode_reason)}</span></li>`).join('')}</ul>`;
  html += '<p class="muted small">Case chưa được bạn kiểm tra sẽ ghi <b>"Chưa test"</b> trong báo cáo — tool không bao giờ tự ghi Pass cho những case này.</p>';
  await dialog({ title: `Bạn vừa chọn ${tcs.length} case cần người kiểm tra`, html, warn: true, buttons: [{ label: 'Đã hiểu', value: true, cls: 'primary' }] });
}

function setupTree() {
  $('tree').addEventListener('click', async (e) => {
    const tcEl = e.target.closest('.tc');
    const grpHead = e.target.closest('.grp-head');
    if (e.target.classList.contains('tchk')) {
      const k = tcEl.dataset.k;
      if (e.target.checked) { S.sel.add(k); renderTree(); await warnManual([k]); } else { S.sel.delete(k); renderTree(); }
      return;
    }
    if (e.target.classList.contains('gchk')) {
      const g = e.target.closest('.grp');
      const keys = [...g.querySelectorAll('.tc')].map((x) => x.dataset.k);
      if (e.target.checked) { const add = keys.filter((k) => !S.sel.has(k)); keys.forEach((k) => S.sel.add(k)); renderTree(); await warnManual(add); } else { keys.forEach((k) => S.sel.delete(k)); renderTree(); }
      return;
    }
    if (grpHead) {
      const gid = grpHead.parentElement.dataset.g;
      S.closedGroups.has(gid) ? S.closedGroups.delete(gid) : S.closedGroups.add(gid);
      renderTree(); return;
    }
    if (tcEl && e.target.closest('.tc-row')) {
      const k = tcEl.dataset.k;
      if (S.open.has(k)) S.open.delete(k); else S.open.add(k);
      S.expandAll = false;
      renderTree();
    }
  });
  ['f-search', 'f-module', 'f-type', 'f-prio', 'f-mode'].forEach((id) => $(id).addEventListener('input', renderTree));
  $('chk-all').onchange = async (e) => {
    const vis = visibleTCs().map((t) => t.key);
    if (e.target.checked) { const add = vis.filter((k) => !S.sel.has(k)); vis.forEach((k) => S.sel.add(k)); renderTree(); await warnManual(add); } else { vis.forEach((k) => S.sel.delete(k)); renderTree(); }
  };
  $('btn-expand').onclick = () => { S.expandAll = true; S.closedGroups.clear(); renderTree(); };
  $('btn-collapse').onclick = () => { S.expandAll = false; S.open.clear(); renderTree(); };
  $('o-env').onchange = () => { showEnvVars(); saveRunOpts(); };
  ['o-browser', 'o-speed', 'o-headless', 'o-shot', 'o-video', 'o-baseurl'].forEach((id) => ($(id).onchange = saveRunOpts));
}
function runOpts() {
  return { env: $('o-env').value, browser: $('o-browser').value, speed: $('o-speed').value, headless: $('o-headless').checked,
    shot_every_step: $('o-shot').checked, video: $('o-video').checked, base_url: $('o-baseurl').value.trim() };
}
function saveRunOpts() { guard(() => api('/api/settings', runOpts()).then((r) => (S.state.settings = r.settings))); }

// ============================================================ chạy
async function confirmRun() {
  const sel = S.wb.testcases.filter((t) => S.sel.has(t.key));
  if (!sel.length) return;
  const by = (m) => sel.filter((t) => t.mode === m);
  const semi = by('Bán tự động'), tay = by('Tay');
  const bad = sel.filter((t) => (t.warnings || []).length);
  const o = runOpts();
  let html = `<div class="count-grid"><div class="cg-auto"><b>${by('Tự động').length}</b>Tự động</div>
    <div class="cg-semi"><b>${semi.length}</b>⚠ Bán tự động</div><div class="cg-manual"><b>${tay.length}</b>✋ Test tay</div></div>
    <p>Môi trường: <b>${esc(o.env)}</b> · Trình duyệt: <b>${esc(o.browser)}</b>${o.headless ? ' (chạy ẩn)' : ''}</p>`;
  if (semi.length || tay.length) {
    html += `<p class="warn-text">Các case dưới đây cần BẠN kiểm tra. Tool sẽ ${S.state.settings.auto_first !== false ? 'chạy hết case tự động trước, rồi ' : ''}dừng lại ở từng case và hiện hộp thoại để bạn chọn Pass/Fail:</p><ul>` +
      [...semi, ...tay].map((t) => `<li><span class="mode ${modeCls(t.mode)}">${modeLabel(t.mode)}</span> <b>${esc(t.id)}</b> ${esc(t.fields.tieu_de || '')}</li>`).join('') + '</ul>';
  }
  if (bad.length) html += `<p class="warn-text">⚠ ${bad.length} TC có cảnh báo khi đọc file:</p><ul>` + bad.map((t) => `<li><b>${esc(t.id)}</b>: ${esc(t.warnings.join(' | '))}</li>`).join('') + '</ul>';
  if (o.headless && (semi.length || tay.length)) html += '<p class="muted small">Đang chọn <b>chạy ẩn</b>: với case test tay bạn chỉ xem được ảnh chụp, không thao tác trên trình duyệt được.</p>';
  const ok = await dialog({ title: `Chạy ${sel.length} test case?`, html, warn: !!(semi.length || tay.length), buttons: [{ label: 'Quay lại', value: false }, { label: '▶ Bắt đầu chạy', value: true, cls: 'primary' }] });
  if (!ok) return;
  const order = S.wb.testcases.filter((t) => S.sel.has(t.key)).map((t) => t.key);
  const r = await guard(() => api('/api/run', { selection: order, options: o }));
  if (!r) return;
  S.runId = r.run_id; S.since = 0; S.liveKeys.clear(); S.manualShownFor = null;
  $('live').innerHTML = '';
  showProgress(true);
  startPoll();
}

function showProgress(on) {
  $('progress').hidden = !on;
  $('btn-goto-results').hidden = true;
  $('btn-stop').hidden = false;
}
function startPoll() {
  if (S.poll) return;
  S.poll = setInterval(poll, 600);
  updateSummary();
  poll();
}
async function poll() {
  let d;
  try { d = await api('/api/run/poll?since=' + S.since); } catch (e) { return; }
  const run = d.run;
  if (!run) return;
  S.runId = run.run_id;
  for (const ev of d.events) {
    S.since = ev.seq;
    if (ev.kind === 'tc_done') addLive(ev.label, ev.status, ev.duration);
    else if (ev.kind === 'warn') addLive(ev.msg, 'warn');
    else if (ev.kind === 'error') addLive(ev.msg, 'err');
    else if (ev.kind === 'info') addLive(ev.msg, 'info');
  }
  const pct = run.total ? Math.round(run.done * 100 / run.total) : 0;
  $('prog-bar').style.width = pct + '%';
  $('prog-count').textContent = `${run.done}/${run.total}`;
  $('prog-state').textContent = { 'đang chạy': '⏳ Đang chạy', 'chờ bạn kiểm tra': '✋ Đang chờ bạn kiểm tra', 'xong': '✅ Đã chạy xong', 'đã dừng': '■ Đã dừng', 'lỗi': '❌ Lỗi', 'chuẩn bị': 'Đang chuẩn bị…' }[run.state] || run.state;
  $('prog-current').textContent = run.current ? `${run.current.label} — ${run.current.title || ''}${run.current.step ? ' · bước ' + run.current.step : ''}` : '';
  $('prog-chips').innerHTML = Object.entries(run.counts).map(([k, v]) => `<span class="st st-${stCls(k)}">${esc(k)}: ${v}</span>`).join('');
  // bước test tay
  if (run.manual) {
    $('btn-manual-reopen').hidden = false;
    if (S.manualShownFor !== run.manual.req_id) { S.manualShownFor = run.manual.req_id; openManual(run.manual); }
  } else {
    $('btn-manual-reopen').hidden = true;
    if (!$('manual').hidden && S.manualReq) { $('manual').hidden = true; S.manualReq = null; }
  }
  document.title = run.manual ? (document.title.startsWith('✋') ? '⏸ Chờ bạn kiểm tra' : '✋ CẦN BẠN KIỂM TRA') : (run.state === 'đang chạy' ? `⏳ ${run.done}/${run.total} — Tool Test` : 'Tool Test Tự Động');
  if (['xong', 'đã dừng', 'lỗi'].includes(run.state)) {
    clearInterval(S.poll); S.poll = null;
    $('btn-stop').hidden = true;
    $('btn-goto-results').hidden = false;
    document.title = 'Tool Test Tự Động';
    updateSummary();
    refreshRuns(run.run_id);
    if (window.scanOnFinish && window.scanOnFinish(run)) return;
    toast(run.state === 'xong' ? 'Đã chạy xong! Bấm "Xem kết quả & xuất file".' : 'Lần chạy đã dừng.');
  }
}
function addLive(label, status, dur) {
  const li = document.createElement('li');
  if (status === 'warn' || status === 'err' || status === 'info') { li.className = status; li.textContent = ({ err: '❌ ', warn: '⚠ ', info: 'ℹ ' })[status] + label; } else li.innerHTML = `<span class="st st-${stCls(status)}">${esc(status)}</span><span class="lbl" title="${esc(label)}">${esc(label)}</span><span class="muted small">${dur ?? ''}s</span>`;
  $('live').appendChild(li);
  $('live').scrollTop = 1e6;
}

// ---------- hộp thoại TEST TAY
function beep() {
  try {
    const ac = new (window.AudioContext || window.webkitAudioContext)();
    [0, 0.25].forEach((t) => { const o = ac.createOscillator(), g = ac.createGain(); o.frequency.value = 880; g.gain.value = 0.08; o.connect(g); g.connect(ac.destination); o.start(ac.currentTime + t); o.stop(ac.currentTime + t + 0.15); });
  } catch (e) { /* bỏ qua */ }
}
function openManual(req) {
  S.manualReq = { req, statuses: req.steps.map(() => ''), images: [] };
  renderManual();
  $('manual').hidden = false;
  beep();
  try { window.focus(); } catch (e) { /* */ }
}
function renderManual() {
  const { req, statuses, images } = S.manualReq;
  if (req.kind === 'login') {
    $('manual-box').innerHTML = `<div class="mh manualh"><h2>🔑 ${esc(req.title || 'Đăng nhập')}</h2></div>
      <div class="mb"><p style="font-size:15px">${esc(req.reason)}</p>
      <p class="muted">Cửa sổ trình duyệt của tool đang mở trang <b>${esc(req.opened || '')}</b>. Đăng nhập xong (kể cả OTP / captcha nếu có) thì quay lại đây bấm nút bên dưới.</p></div>
      <div class="mf"><button class="btn danger" id="m-stop">■ Dừng</button><button class="btn primary big" id="m-login-ok">✔ Đã đăng nhập xong — tiếp tục</button></div>`;
    $('m-login-ok').onclick = async () => { const r = await guard(() => api('/api/run/manual', { req_id: req.req_id, statuses: [] })); if (r) { $('manual').hidden = true; S.manualReq = null; } };
    $('m-stop').onclick = () => { $('manual').hidden = true; S.manualReq = null; guard(() => api('/api/run/stop', {})); };
    return;
  }
  const runImg = (rel) => `/runs/${encodeURIComponent(req.run_id)}/${rel}`;
  const btns = (i) => ['Pass', 'Fail', 'Blocked', 'Chưa test'].map((s) => `<button class="${stCls(s)} ${statuses[i] === s ? 'on' : ''}" data-i="${i}" data-s="${s}">${s === 'Chưa test' ? 'Để sau' : s}</button>`).join('');
  $('manual-box').innerHTML = `<div class="mh manualh"><h2>✋ CẦN BẠN KIỂM TRA — ${esc(req.label)}</h2><span class="mode ${modeCls(req.mode)}" style="background:#fff">${modeLabel(req.mode)}</span></div>
    <div class="mb"><div class="info"><b>Test case</b><span>${esc(req.tc_id)} — ${esc(req.title)}</span><b>Lý do</b><span>${esc(req.reason)}</span>
      ${req.precondition ? `<b>Precondition</b><span>${esc(req.precondition)}</span>` : ''}${req.test_data ? `<b>Test data</b><span>${esc(req.test_data)}</span>` : ''}
      ${req.opened ? `<b>Trình duyệt</b><span>Tool đã mở <a href="#">${esc(req.opened)}</a> trong cửa sổ trình duyệt test — bạn thao tác trực tiếp ở đó.</span>` : ''}</div>
      <table class="steps"><tr><th>#</th><th>Bước thực hiện</th><th>Kết quả mong đợi</th><th>Ảnh tool chụp</th><th style="width:250px">Kết quả của bạn</th></tr>
      ${req.steps.map((s, i) => `<tr class="${S.manualReq.need && !statuses[i] ? 'need' : ''}"><td>${esc(s.no)}</td><td>${esc(s.buoc)}${s.data ? `<div class="small muted">Dữ liệu: ${esc(s.data)}</div>` : ''}</td><td>${esc(s.expected)}</td>
        <td>${s.images.map((im) => `<img src="${runImg(im)}" data-full="${runImg(im)}">`).join('') || '<span class="muted small">—</span>'}</td>
        <td><div class="vbtn">${btns(i)}</div><input placeholder="Ghi chú / kết quả thực tế (không bắt buộc)" data-note="${i}" value="${esc((S.manualReq.notes || [])[i] || '')}"></td></tr>`).join('')}
      </table>
      ${req.steps.length > 1 ? '<div class="row"><button class="btn small" id="m-allpass">Tất cả Pass</button><button class="btn small" id="m-rest">Các bước còn lại = Pass</button></div>' : ''}
      <label>Ghi chú chung</label><textarea id="m-note" rows="2" placeholder="VD: Màu nút Lưu bị nhạt, khó nhìn">${esc(S.manualReq.note || '')}</textarea>
      <div class="paste" id="m-paste">📋 Dán ảnh (Ctrl+V) hoặc kéo-thả ảnh vào đây làm bằng chứng<div class="thumbs">${images.map((im) => `<img src="${runImg(im)}" data-full="${runImg(im)}">`).join('')}</div></div>
    </div><div class="mf"><span class="grow small muted">Phải chọn kết quả cho từng bước. "Để sau" = ghi <b>Chưa test</b> trong báo cáo.</span>
      <button class="btn danger" id="m-stop">■ Dừng cả lần chạy</button><button class="btn primary big" id="m-submit">Lưu kết quả &amp; chạy tiếp →</button></div>`;
  const box = $('manual-box');
  box.querySelectorAll('.vbtn button').forEach((b) => (b.onclick = () => { keepNotes(); S.manualReq.statuses[+b.dataset.i] = b.dataset.s; renderManual(); }));
  const ap = $('m-allpass'); if (ap) ap.onclick = () => { keepNotes(); S.manualReq.statuses = statuses.map(() => 'Pass'); renderManual(); };
  const rp = $('m-rest'); if (rp) rp.onclick = () => { keepNotes(); S.manualReq.statuses = statuses.map((s) => s || 'Pass'); renderManual(); };
  box.querySelectorAll('img[data-full]').forEach((im) => (im.onclick = () => lightbox(im.dataset.full)));
  $('m-submit').onclick = submitManual;
  $('m-stop').onclick = async () => {
    const ok = await dialog({ title: 'Dừng lần chạy?', html: '<p>Các case chưa chạy sẽ ghi <b>Chưa test</b>. Kết quả đã có vẫn được giữ để xuất file.</p>', buttons: [{ label: 'Không', value: false }, { label: 'Dừng', value: true, cls: 'primary' }] });
    if (ok) { $('manual').hidden = true; S.manualReq = null; guard(() => api('/api/run/stop', {})); }
  };
  const paste = $('m-paste');
  paste.ondragover = (e) => { e.preventDefault(); paste.classList.add('over'); };
  paste.ondragleave = () => paste.classList.remove('over');
  paste.ondrop = (e) => { e.preventDefault(); e.stopPropagation(); paste.classList.remove('over'); [...e.dataTransfer.files].forEach(addManualImage); };
}
function keepNotes() {
  if (!S.manualReq) return;
  S.manualReq.notes = [...$('manual-box').querySelectorAll('input[data-note]')].map((x) => x.value);
  const n = $('m-note'); if (n) S.manualReq.note = n.value;
}
function addManualImage(file) {
  if (!file || !/^image\//.test(file.type) || !S.manualReq) return;
  const rd = new FileReader();
  rd.onload = async () => {
    const r = await guard(() => api('/api/run/image', { run_id: S.manualReq.req.run_id, data: rd.result }));
    if (r) { keepNotes(); S.manualReq.images.push(r.rel); renderManual(); toast('Đã thêm ảnh'); }
  };
  rd.readAsDataURL(file);
}
document.addEventListener('paste', (e) => {
  if ($('manual').hidden) return;
  const items = [...(e.clipboardData || {}).items || []].filter((i) => i.type.startsWith('image/'));
  if (items.length) { e.preventDefault(); items.forEach((i) => addManualImage(i.getAsFile())); }
});
async function submitManual() {
  keepNotes();
  const m = S.manualReq;
  if (m.statuses.some((s) => !s)) { m.need = true; renderManual(); toast('Hãy chọn kết quả cho tất cả các bước (hoặc "Để sau").', true); return; }
  const r = await guard(() => api('/api/run/manual', { req_id: m.req.req_id, statuses: m.statuses, notes: m.notes, note: m.note, images: m.images }));
  if (r) { $('manual').hidden = true; S.manualReq = null; document.title = 'Tool Test Tự Động'; }
}

// ============================================================ kết quả
async function refreshRuns(select) {
  const d = await guard(() => api('/api/runs'));
  if (!d) return;
  const sel = $('r-run');
  sel.innerHTML = d.runs.map((r) => `<option value="${esc(r.run_id)}">${esc((r.started || '').replace('T', ' '))} · ${esc(r.project || '')} · ${r.total} lượt (✔${r.counts.Pass || 0} ✘${r.counts.Fail || 0})</option>`).join('') || '<option value="">(chưa có lần chạy nào)</option>';
  if (select) sel.value = select;
  else if (S.runId && d.runs.some((r) => r.run_id === S.runId)) sel.value = S.runId;
  if (sel.value) loadResult(sel.value);
  else renderResult(null);
}
async function loadResult(id) {
  const d = await guard(() => api('/api/run/result?run_id=' + encodeURIComponent(id)));
  if (d) { S.result = d; renderResult(d); previewName(); }
}
function renderResult(d) {
  if (!d) { $('r-cards').innerHTML = ''; $('r-list').innerHTML = '<p class="muted">Chưa có kết quả. Hãy chạy test ở tab <b>Chạy test</b>.</p>'; return; }
  const c = { Pass: 0, Fail: 0, Blocked: 0, 'Chưa test': 0 };
  d.results.forEach((r) => (c[r.status] = (c[r.status] || 0) + 1));
  const tot = d.results.length;
  $('r-cards').innerHTML = `<div class="card"><b>${tot}</b>Tổng lượt</div>` + Object.entries(c).map(([k, v]) => `<div class="card ${stCls(k)}"><b>${v}</b>${k}</div>`).join('') +
    `<div class="card"><b>${tot ? Math.round(c.Pass * 100 / tot) : 0}%</b>Tỷ lệ pass</div><div class="card"><b style="font-size:14px">${esc(d.env || '')}</b>Môi trường<br><span class="small muted">${esc(d.state || '')}</span></div>`;
  const runImg = (rel) => `/runs/${encodeURIComponent(d.run_id)}/${rel}`;
  const list = d.results.filter((r) => !S.rFilter || r.status === S.rFilter);
  $('r-list').innerHTML = list.map((r, i) => `<div class="r-item ${stCls(r.status)} ${r.status !== 'Pass' && list.length < 30 ? '' : ''}" data-i="${i}">
    <div class="r-head"><span class="st st-${stCls(r.status)}">${esc(r.status)}</span><span class="lbl">${esc(r.label)}</span><span class="ttl">${esc(r.title)}</span>
      <span class="mode ${modeCls(r.mode)}">${modeLabel(r.mode)}</span><span class="meta">${r.duration}s</span></div>
    <div class="r-body">${r.note ? `<div class="reason">${esc(r.note)}</div>` : ''}
      ${r.video ? `<p><a href="${runImg(r.video)}" target="_blank">🎬 Xem video</a></p>` : ''}
      <table class="rsteps"><tr><th>#</th><th>Bước</th><th>Mong đợi</th><th>Thực tế</th><th>KQ</th><th>Ảnh</th></tr>
      ${r.steps.map((s) => `<tr><td>${esc(s.no)}</td><td>${esc(s.fields.buoc || s.action || '')}${s.manual ? ' <span class="mode m-manual">người test</span>' : ''}</td><td>${esc(s.fields.expected || '')}</td>
        <td>${esc(s.actual || '')}</td><td><span class="st st-${stCls(s.status)}">${esc(s.status || '')}</span></td>
        <td><div class="thumbs">${(s.images || []).map((im) => `<img src="${runImg(im)}" data-full="${runImg(im)}" data-base="${s.baseline && !im.includes('_khac') ? esc(im) : ''}" title="Bấm để phóng to">`).join('')}</div>
        ${s.baseline && (s.images || []).length ? `<button class="btn link small" data-baseline="${esc(s.images.find((x) => !x.includes('_khac')) || '')}">Dùng ảnh này làm ảnh chuẩn</button>` : ''}</td></tr>`).join('')}</table></div></div>`).join('') || '<p class="muted">Không có kết quả nào khớp bộ lọc.</p>';
  $('r-list').querySelectorAll('.r-head').forEach((h) => (h.onclick = () => h.parentElement.classList.toggle('open')));
  $('r-list').querySelectorAll('img[data-full]').forEach((im) => (im.onclick = () => lightbox(im.dataset.full)));
  $('r-list').querySelectorAll('[data-baseline]').forEach((b) => (b.onclick = async () => {
    const r = await guard(() => api('/api/baseline', { run_id: d.run_id, image: b.dataset.baseline }));
    if (r) toast('Đã cập nhật ảnh chuẩn. Lần chạy sau sẽ so với ảnh này.');
  }));
}
function lightbox(src) { $('lightbox-img').src = src; $('lightbox-cap').textContent = decodeURIComponent(src.split('/').pop()); $('lightbox').hidden = false; }

// ---------- xuất file
function renderVars() {
  $('x-vars').innerHTML = Object.entries(S.state.vars_help).map(([k, v]) => `<button title="${esc(v)}" data-v="{${k}}">{${k}}</button>`).join('');
  $('x-vars').querySelectorAll('button').forEach((b) => (b.onclick = () => {
    const inp = $('x-pattern'); const p = inp.selectionStart ?? inp.value.length;
    inp.value = inp.value.slice(0, p) + b.dataset.v + inp.value.slice(inp.selectionEnd ?? p);
    inp.focus(); previewName();
  }));
}
let pvT;
function previewName() {
  clearTimeout(pvT);
  pvT = setTimeout(async () => {
    const d = await api('/api/export/preview', { run_id: $('r-run').value || '', pattern: $('x-pattern').value, folder: $('x-folder').value }).catch(() => null);
    if (d) { $('x-preview').textContent = d.name; $('x-exists').hidden = !d.exists; }
  }, 200);
}
function exportOpts() {
  return { pattern: $('x-pattern').value, folder: $('x-folder').value, html: $('x-html').checked, embed: $('x-embed').checked, writeback: $('x-writeback').checked, open_after: $('x-open').checked, round: $('x-round').value };
}
async function doExport(overwrite) {
  const id = $('r-run').value;
  if (!id) return toast('Chưa có lần chạy nào để xuất.', true);
  const body = { run_id: id, overwrite: overwrite || 'ask', ...exportOpts() };
  $('x-go').disabled = true; $('x-go').textContent = '⏳ Đang xuất…';
  const r = await guard(() => api('/api/export', body));
  $('x-go').disabled = false; $('x-go').textContent = '⬇ Xuất file Excel';
  if (!r) return;
  if (r.exists) {
    const a = await dialog({ title: 'Trùng tên file', warn: true, html: `<p>Đã có file:<br><b>${esc(r.path)}</b></p><p>Bạn muốn làm gì?</p>`, buttons: [{ label: 'Hủy', value: null }, { label: 'Ghi đè file cũ', value: 'yes' }, { label: 'Lưu thành file mới (thêm số)', value: 'suffix', cls: 'primary' }] });
    if (a) doExport(a);
    return;
  }
  $('x-done').innerHTML = `<div class="ok"><b>✅ Đã xuất ${r.files.length} file</b>${r.images ? ` (kèm ${r.images} ảnh)` : ''}:
    ${r.files.map((f) => `<div class="f"><span class="grow">${esc(f)}</span><button class="btn small" data-p="${esc(f)}">Mở</button></div>`).join('')}
    <div class="f"><button class="btn small" data-p="${esc(r.folder)}">📁 Mở thư mục</button></div></div>`;
  $('x-done').querySelectorAll('[data-p]').forEach((b) => (b.onclick = () => guard(() => api('/api/open_path', { path: b.dataset.p }))));
  previewName();
}

// ============================================================ cấu hình cột
function renderColumns() {
  const st = S.state;
  $('c-profile').innerHTML = st.settings.profiles.map((p) => `<option ${p.name === S.profile.name ? 'selected' : ''}>${esc(p.name)}</option>`).join('');
  document.querySelectorAll('#c-kind button').forEach((b) => b.classList.toggle('on', b.dataset.kind === S.kind));
  const cat = {}; st.catalog[S.kind].forEach((c) => (cat[c.key] = c));
  const rows = S.profile[S.kind];
  const required = S.kind === 'testcase' ? ['tc_id'] : ['ts_id'];
  let html = '<tr><th style="width:30px"></th><th style="width:44px">Bật</th><th>Tên cột (sửa được)</th><th>Tên mặc định</th><th>Nhóm</th><th>Ghi ở</th><th></th><th style="width:70px">Thứ tự</th></tr>';
  rows.forEach((r, i) => {
    const c = cat[r.key];
    html += `<tr class="${r.on ? '' : 'off'}" draggable="true" data-i="${i}"><td class="muted" style="cursor:grab">⋮⋮</td>
      <td><input type="checkbox" data-on="${i}" ${r.on ? 'checked' : ''} ${required.includes(r.key) ? 'disabled title="Cột bắt buộc"' : ''}></td>
      <td><input type="text" data-label="${i}" value="${esc(r.label)}"></td><td class="muted small">${esc(c.label)}</td>
      <td class="small">${esc(c.group)}</td><td class="lv">${S.kind === 'testcase' ? (c.level === 'step' ? 'mỗi dòng bước' : 'dòng đầu TC') : ''}</td>
      <td>${c.result ? '<span class="badge-auto">tool tự điền</span>' : ''}</td>
      <td class="mv"><button data-up="${i}">↑</button> <button data-down="${i}">↓</button></td></tr>`;
  });
  $('c-table').innerHTML = html;
  $('c-count').textContent = `Đang bật ${rows.filter((r) => r.on).length}/${rows.length} cột`;
  const tb = $('c-table');
  tb.querySelectorAll('[data-on]').forEach((x) => (x.onchange = () => { rows[+x.dataset.on].on = x.checked; dirty(); renderColumns(); }));
  tb.querySelectorAll('[data-label]').forEach((x) => (x.oninput = () => { rows[+x.dataset.label].label = x.value; dirty(); }));
  tb.querySelectorAll('[data-up]').forEach((x) => (x.onclick = () => move(+x.dataset.up, -1)));
  tb.querySelectorAll('[data-down]').forEach((x) => (x.onclick = () => move(+x.dataset.down, 1)));
  let from = null;
  tb.querySelectorAll('tr[draggable]').forEach((tr) => {
    tr.ondragstart = () => (from = +tr.dataset.i);
    tr.ondragover = (e) => { e.preventDefault(); tr.classList.add('drag-over'); };
    tr.ondragleave = () => tr.classList.remove('drag-over');
    tr.ondrop = (e) => { e.preventDefault(); const to = +tr.dataset.i; if (from !== null && from !== to) { const [x] = rows.splice(from, 1); rows.splice(to, 0, x); dirty(); renderColumns(); } };
  });
  $('c-save').textContent = S.profileDirty ? '💾 Lưu (chưa lưu)' : '💾 Lưu';
  $('c-layout').value = S.profile.layout || 'rows';
  $('c-sections').value = S.profile.sections ? '1' : '0';
  $('c-matrices').value = S.profile.matrices === false ? '0' : '1';
}
function move(i, d) {
  const rows = S.profile[S.kind]; const j = i + d;
  if (j < 0 || j >= rows.length) return;
  [rows[i], rows[j]] = [rows[j], rows[i]]; dirty(); renderColumns();
}
function dirty() { S.profileDirty = true; $('c-save').textContent = '💾 Lưu (chưa lưu)'; }
async function saveProfile(name, oldName) {
  const r = await guard(() => api('/api/profile/save', { profile: S.profile, name: name || S.profile.name }));
  if (!r) return;
  S.profileDirty = false;
  await loadState();
  toast('Đã lưu mẫu cột "' + r.profile.name + '"');
}
function renderMapping() {
  const wb = S.wb;
  if (!wb) { $('c-mapping').innerHTML = ''; return; }
  const cat = S.state.catalog;
  let h = `<h3>Cột tool nhận được trong file đang mở (${esc(wb.file_name)})</h3><p class="muted small">Cột tool không nhận ra vẫn được giữ nguyên trong báo cáo. Muốn tool hiểu đúng ý nghĩa: chọn cột tương ứng ở bên phải — tool sẽ đổi tên cột đó trong mẫu cột theo tên trong file của bạn và đọc lại file.</p>`;
  for (const kind of ['scenario', 'testcase']) {
    const m = wb.headers[kind];
    if (!m) continue;
    h += `<p><b>${kind === 'scenario' ? 'Sheet Scenario' : 'Sheet Test Case'}</b> — "${esc(wb.sheets[kind].name)}", dòng tiêu đề ${wb.sheets[kind].header_row}</p><table><tr><th>Tên cột trong file</th><th>Tool hiểu là</th></tr>` +
      m.map((x) => `<tr><td>${esc(x.header)}</td><td><select data-kind="${kind}" data-h="${esc(x.header)}" class="${x.key ? '' : 'unmapped'}"><option value="">(không nhận ra — giữ nguyên)</option>${cat[kind].map((c) => `<option value="${c.key}" ${c.key === x.key ? 'selected' : ''}>${esc(c.label)}</option>`).join('')}</select></td></tr>`).join('') + '</table>';
  }
  $('c-mapping').innerHTML = h;
  $('c-mapping').querySelectorAll('select').forEach((s) => (s.onchange = async () => {
    if (!s.value) return;
    const rows = S.profile[s.dataset.kind];
    const row = rows.find((r) => r.key === s.value);
    row.label = s.dataset.h; row.on = true;
    await saveProfile();
  }));
}

// ============================================================ hướng dẫn
function renderGuide() {
  $('g-kw').innerHTML = '<tr><th>Hành động</th><th>Đối tượng</th><th>Giá trị</th><th>Ý nghĩa</th><th>Ví dụ</th></tr>' +
    S.state.keywords.map((k) => `<tr><td class="act ${k.manual ? 'manual' : ''}">${esc(k.name)}</td><td>${esc(k.target || '-')}</td><td>${esc(k.value || '-')}</td><td>${esc(k.desc)}${k.manual ? ' <span class="mode m-manual">test tay</span>' : ''}</td><td class="small">${esc([k.ex_target, k.ex_value].filter(Boolean).join(' → '))}</td></tr>`).join('');
  $('g-types').innerHTML = '<tr><th>Loại test</th><th>Cách test mặc định</th><th>Ý nghĩa</th></tr>' +
    S.state.test_types.map(([n, m, d]) => `<tr><td><b>${esc(n)}</b></td><td><span class="mode ${modeCls(m)}">${modeLabel(m)}</span></td><td>${esc(d)}</td></tr>`).join('');
}

// ============================================================ tab + sự kiện
function placeProgress(t) {
  const host = t === 'scan' ? (window.scanProgressHost ? window.scanProgressHost() : $('scan-progress-host')) : (t === 'run' ? $('run-progress-host') : null);
  if (host && $('progress').parentElement !== host) host.appendChild($('progress'));
}
function switchTab(t) {
  document.querySelectorAll('#tabs button, #tabs2 button').forEach((b) => b.classList.toggle('on', b.dataset.tab === t || (b.dataset.group || '').split(' ').includes(t)));
  document.querySelectorAll('.setnav button').forEach((b) => b.classList.toggle('on', b.dataset.goto === t));
  $('filebar').hidden = !(S.wb && t === 'run');
  document.querySelectorAll('.tab').forEach((s) => (s.hidden = s.id !== 'tab-' + t));
  placeProgress(t);
  if (t === 'scan' && window.scanOnTab) window.scanOnTab();
  if (t === 'results') refreshRuns($('r-run').value || S.runId);
  if (t === 'columns') { renderColumns(); renderMapping(); }
}
function bind() {
  document.querySelectorAll('#tabs button, #tabs2 button').forEach((b) => (b.onclick = () => switchTab(b.dataset.tab)));
  document.querySelectorAll('[data-goto]').forEach((a) => (a.onclick = (e) => { e.preventDefault(); switchTab(a.dataset.goto); }));
  $('btn-open').onclick = () => openFile();
  $('btn-open2').onclick = () => openFile();
  $('btn-reload').onclick = reloadFile;
  $('btn-open-excel').onclick = () => S.wb && guard(() => api('/api/open_path', { path: S.wb.file }));
  $('btn-tpl-sample').onclick = () => makeTemplate(true);
  $('btn-tpl-blank').onclick = () => makeTemplate(false);
  $('g-sample').onclick = () => makeTemplate(true);
  $('c-tpl').onclick = async () => { if (S.profileDirty) await saveProfile(); makeTemplate(false); };
  $('g-demo').onclick = () => window.open(S.state.demo_url + '/login.html', '_blank');
  $('s-demo').onclick = (e) => { e.preventDefault(); window.open(S.state.demo_url + '/login.html', '_blank'); };
  $('s-data').onclick = (e) => { e.preventDefault(); guard(() => api('/api/open_path', { path: S.state.data_dir })); };
  $('btn-run').onclick = confirmRun;
  $('btn-stop').onclick = async () => {
    const ok = await dialog({ title: 'Dừng lần chạy?', html: '<p>Case đang chạy và các case chưa chạy sẽ ghi <b>Chưa test</b>. Kết quả đã có vẫn xuất được.</p>', buttons: [{ label: 'Không', value: false }, { label: 'Dừng', value: true, cls: 'primary' }] });
    if (ok) guard(() => api('/api/run/stop', {}));
  };
  $('btn-goto-results').onclick = () => { switchTab('results'); };
  $('btn-manual-reopen').onclick = () => { if (S.manualReq) $('manual').hidden = false; else S.manualShownFor = null; };
  $('lightbox').onclick = () => ($('lightbox').hidden = true);
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') { $('lightbox').hidden = true; } });
  $('btn-quit').onclick = async () => {
    const ok = await dialog({ title: 'Tắt tool?', html: '<p>Tool sẽ tắt hẳn' + (S.poll ? ' và <b>dừng lần chạy đang chạy</b>' : '') + '. Mở lại bằng cách bấm đúp file tool.</p>', buttons: [{ label: 'Không', value: false }, { label: 'Tắt tool', value: true, cls: 'primary' }] });
    if (!ok) return;
    await api('/api/quit', {}).catch(() => {});
    document.body.innerHTML = '<div style="display:grid;place-items:center;height:100vh;font:16px Segoe UI"><div>Đã tắt Tool Test Tự Động. Bạn có thể đóng cửa sổ này.</div></div>';
    setTimeout(() => window.close(), 400);
  };
  // kết quả
  $('r-run').onchange = () => loadResult($('r-run').value);
  $('r-filter').querySelectorAll('button').forEach((b) => (b.onclick = () => { S.rFilter = b.dataset.st; $('r-filter').querySelectorAll('button').forEach((x) => x.classList.toggle('on', x === b)); renderResult(S.result); }));
  ['x-pattern', 'x-folder'].forEach((id) => $(id).addEventListener('input', previewName));
  ['x-html', 'x-embed', 'x-writeback', 'x-open', 'x-round'].forEach((id) => ($(id).onchange = () => guard(() => api('/api/settings', { export: exportOpts() }))));
  $('c-layout').onchange = () => { S.profile.layout = $('c-layout').value; dirty(); };
  $('c-sections').onchange = () => { S.profile.sections = $('c-sections').value === '1'; dirty(); };
  $('c-matrices').onchange = () => { S.profile.matrices = $('c-matrices').value === '1'; dirty(); };
  $('x-pick').onclick = async () => { const r = await guard(() => api('/api/pick_folder', {})); if (r && r.folder) { $('x-folder').value = r.folder; previewName(); } };
  $('x-go').onclick = () => doExport();
  // cấu hình cột
  $('c-kind').querySelectorAll('button').forEach((b) => (b.onclick = () => { S.kind = b.dataset.kind; renderColumns(); }));
  $('c-profile').onchange = async () => {
    if (S.profileDirty && !(await dialog({ title: 'Bỏ thay đổi?', html: '<p>Mẫu cột hiện tại có thay đổi chưa lưu.</p>', buttons: [{ label: 'Ở lại', value: false }, { label: 'Bỏ thay đổi', value: true, cls: 'primary' }] }))) { $('c-profile').value = S.profile.name; return; }
    await guard(() => api('/api/profile/activate', { name: $('c-profile').value })); S.profileDirty = false; await loadState();
  };
  $('c-save').onclick = () => saveProfile();
  $('c-saveas').onclick = async () => {
    const name = prompt('Tên mẫu cột mới (VD: Dự án Đơn hàng tuyển dụng):', S.profile.name + ' (bản sao)');
    if (name && name.trim()) saveProfile(name.trim());
  };
  $('c-delete').onclick = async () => {
    const ok = await dialog({ title: 'Xóa mẫu cột?', html: `<p>Xóa mẫu "<b>${esc(S.profile.name)}</b>"?</p>`, buttons: [{ label: 'Không', value: false }, { label: 'Xóa', value: true, cls: 'primary' }] });
    if (ok && (await guard(() => api('/api/profile/delete', { name: S.profile.name })))) { S.profileDirty = false; await loadState(); }
  };
  const req = () => (S.kind === 'testcase' ? ['tc_id'] : ['ts_id']);
  $('c-all').onclick = () => { S.profile[S.kind].forEach((r) => (r.on = true)); dirty(); renderColumns(); };
  $('c-none').onclick = () => { S.profile[S.kind].forEach((r) => (r.on = req().includes(r.key))); dirty(); renderColumns(); };
  $('c-default').onclick = () => {
    const cat = S.state.catalog[S.kind];
    S.profile[S.kind] = cat.map((c) => ({ key: c.key, on: c.default, label: c.label })); dirty(); renderColumns();
  };
  // cài đặt
  $('s-save').onclick = async () => {
    const r = await guard(() => api('/api/settings', { tester: $('s-tester').value.trim(), project_name: $('s-project').value.trim(), size: $('s-size').value.trim(), timeout: +$('s-timeout').value || 10, auto_first: $('s-autofirst').checked, stop_on_fail: $('s-stoponfail').checked }));
    if (r) { S.state.settings = r.settings; toast('Đã lưu cài đặt'); }
  };
  window.addEventListener('beforeunload', (e) => { if (S.profileDirty) { e.preventDefault(); e.returnValue = ''; } });
  // kiểm tra file Excel có bị sửa không (mỗi khi quay lại cửa sổ)
  window.addEventListener('focus', async () => {
    if (S.poll) return;
    const st = await api('/api/state').catch(() => null);
    if (!st || !st.workbook) return;
    if (!S.wb || S.wb.file !== st.workbook.file) { S.state = st; setWorkbook(st.workbook, false); return; }
    $('file-modified').hidden = !st.workbook.modified;
  });
}

bind();
setupTree();
setupDrop();
loadState().then(() => { renderVars(); previewName(); switchTab('scan'); }).catch((e) => toast(e.message, true));
