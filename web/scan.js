'use strict';
// ============================================================ Tab "Test web" — MỘT trang cấu hình theo bước:
// 1 chọn Test Scenario / Test Case → 2 cột → 3 nguồn (URL / file) → 4 nhóm test → 5 xem trước → 6 sinh file / điền file.
const SC = { cfg: null, disc: null, plan: [], meta: null, sel: new Set(), selTc: new Set(), known: new Set(), knownTc: new Set(),
  pw: { login: '', perm: '' }, saveT: null, planT: null, adv: 'general', running: null, lastRun: null, runRes: {}, runInfo: '',
  needDiscover: false, F: null };

const clone = (x) => JSON.parse(JSON.stringify(x));
const MODE_BADGE = { 'Tự động': '<span class="mode m-auto">Tự động</span>', 'Bán tự động': '<span class="mode m-semi">⚠ Bán tự động</span>', 'Tay': '<span class="mode m-manual">✋ Test tay</span>' };
const PRIOS = ['High', 'Medium', 'Low'];
const kind = () => (SC.cfg.kind === 'testcase' ? 'testcase' : 'scenario');
const isTC = () => kind() === 'testcase';
const unaccent = (s) => String(s || '').replace(/đ/g, 'd').replace(/Đ/g, 'D').normalize('NFD').replace(/[̀-ͯ]/g, '');

// ------------------------------------------------------------ khởi tạo / lưu cấu hình
function scInit() {
  const st = S.state;
  SC.cfg = clone(st.settings.scan || {});
  const c = SC.cfg;
  c.login = c.login || { mode: 'none' }; c.types = c.types || {}; c.rules = c.rules || []; c.state_rules = c.state_rules || [];
  c.perm = c.perm || { role: '', rows: [] }; c.click = c.click || {}; c.source = c.source || 'auto'; c.scenarios_in = c.scenarios_in || [];
  c.kind = c.kind || 'scenario'; c.edits = c.edits || {}; c.out = c.out || {}; if (c.suggest === undefined) c.suggest = true;
  if (!c.out.folder) c.out.folder = (st.settings.export || {}).folder || '';
  SC.disc = st.discovery || null;
  $('sc-url').value = c.url || '';
  $('sc-login').value = c.login.mode || 'none';
  $('sc-user').value = c.login.user || '';
  $('sc-lurl').value = c.login.url || '';
  $('g-suggest').checked = c.suggest !== false;
  $('x2-folder').value = c.out.folder || '';
  $('x2-round').value = c.out.round || '1';
  if (c.last_run) loadRun(c.last_run, true);
  renderKind();
  scStatus();
  scPlan();
}
function scSave() { clearTimeout(SC.saveT); SC.saveT = setTimeout(() => api('/api/scan/save', { cfg: SC.cfg }).catch(() => {}), 400); }
function scRead() {
  const c = SC.cfg;
  c.url = $('sc-url').value.trim();
  c.login = Object.assign(c.login || {}, { mode: $('sc-login').value, user: $('sc-user').value.trim(), url: $('sc-lurl').value.trim() });
  SC.pw.login = $('sc-pass').value;
}
function scPages() { return (SC.disc && SC.disc.pages) || []; }

// ------------------------------------------------------------ 1. loại tài liệu
function renderKind() {
  const k = kind();
  document.querySelectorAll('#kinds .kind').forEach((b) => b.classList.toggle('on', b.dataset.kind === k));
  if (k === 'scenario') SC.cfg.source = 'auto';
  $('sc-source').hidden = k !== 'testcase';
  document.querySelectorAll('#sc-source button').forEach((b) => b.classList.toggle('on', b.dataset.src === SC.cfg.source));
  $('sc-file-row').hidden = !(k === 'testcase' && SC.cfg.source === 'file');
  $('sc-file-name').textContent = SC.cfg.scenarios_in.length ? `${(SC.cfg.scenarios_file || '').split(/[\\/]/).pop()} — ${SC.cfg.scenarios_in.length} scenario` : 'Chưa chọn file';
  $('sc-auto').hidden = $('sc-login').value !== 'auto';
  $('x2-tc').hidden = k !== 'testcase';
  $('x2-run').hidden = k !== 'testcase';
  $('x2-pattern').value = (SC.cfg.out || {})[k === 'testcase' ? 'pattern_testcase' : 'pattern_scenario'] || (k === 'testcase' ? 'TestCase_{man_hinh}_{ngay}' : 'TestScenario_{man_hinh}_{ngay}');
  renderCols();
  renderGroups();
  renderRunInfo();
  namePreview();
}

// ------------------------------------------------------------ 2. cột
function companyRows(k) { return clone(((S.state.company_profile || {})[k]) || S.profile[k]); }
function renderCols() {
  const k = kind(), rows = S.profile[k];
  const cat = {}; S.state.catalog[k].forEach((c) => (cat[c.key] = c));
  const req = k === 'testcase' ? 'tc_id' : 'ts_id';
  const prof = $('col-prof');
  prof.innerHTML = S.state.settings.profiles.map((p) => `<option ${p.name === S.profile.name ? 'selected' : ''}>${esc(p.name)}</option>`).join('');
  const on = rows.map((r, i) => [r, i]).filter(([r]) => r.on);
  $('col-on').innerHTML = on.map(([r, i], n) => `<span class="colchip" draggable="true" data-i="${i}" title="${esc(cat[r.key] ? 'Tên mặc định: ' + cat[r.key].label : '')}">
      <span class="grip">⋮⋮</span><span class="cno">${n + 1}</span><input value="${esc(r.label)}" data-lab="${i}" size="${Math.max(4, Math.min(24, r.label.length + 1))}">
      <button class="mv" data-mv="-1" title="Sang trái">‹</button><button class="mv" data-mv="1" title="Sang phải">›</button>
      ${r.key === req ? '' : '<button class="rm" title="Bỏ cột này">✕</button>'}</span>`).join('');
  const off = rows.map((r, i) => [r, i]).filter(([r]) => !r.on);
  $('col-off').innerHTML = off.map(([r, i]) => `<button class="addcol" data-i="${i}" title="${esc(cat[r.key] ? cat[r.key].group : '')}">＋ ${esc(r.label)}</button>`).join('') || '<span class="muted small">(đã dùng hết)</span>';
  $('col-count').textContent = `— ${on.length} cột${S.profileDirty ? ' · chưa lưu' : ''}`;
  $('col-save').textContent = S.profileDirty ? '💾 Lưu mẫu cột *' : '💾 Lưu mẫu cột';
  const box = $('col-on');
  const changed = () => { S.profileDirty = true; renderCols(); pvRender(); };
  box.querySelectorAll('[data-lab]').forEach((x) => (x.oninput = () => { rows[+x.dataset.lab].label = x.value; S.profileDirty = true; $('col-save').textContent = '💾 Lưu mẫu cột *'; clearTimeout(SC.colT); SC.colT = setTimeout(pvRender, 300); }));
  box.querySelectorAll('.colchip').forEach((chip) => {
    const i = +chip.dataset.i;
    chip.querySelectorAll('.mv').forEach((b) => (b.onclick = () => {
      const pos = on.findIndex(([, j]) => j === i), t = on[pos + +b.dataset.mv];
      if (!t) return;
      const [x] = rows.splice(i, 1); rows.splice(t[1], 0, x); changed();
    }));
    const rm = chip.querySelector('.rm'); if (rm) rm.onclick = () => { rows[i].on = false; changed(); };
    chip.ondragstart = (e) => { SC.dragCol = i; e.dataTransfer.effectAllowed = 'move'; };
    chip.ondragover = (e) => { e.preventDefault(); chip.classList.add('over'); };
    chip.ondragleave = () => chip.classList.remove('over');
    chip.ondrop = (e) => { e.preventDefault(); e.stopPropagation(); const from = SC.dragCol; SC.dragCol = null; if (from == null || from === i) return; const [x] = rows.splice(from, 1); rows.splice(from < i ? i - 1 : i, 0, x); changed(); };
  });
  $('col-off').querySelectorAll('.addcol').forEach((b) => (b.onclick = () => {
    const i = +b.dataset.i; const [x] = rows.splice(i, 1); x.on = true;
    const lastOn = rows.map((r, j) => (r.on ? j : -1)).filter((j) => j >= 0).pop();
    rows.splice(lastOn === undefined ? 0 : lastOn + 1, 0, x); changed();
  }));
}
async function colSave(name) {
  const r = await guard(() => api('/api/profile/save', { profile: S.profile, name: name || S.profile.name }));
  if (!r) return;
  S.profileDirty = false;
  await loadState();
  renderCols(); pvRender();
  toast('Đã lưu mẫu cột "' + r.profile.name + '"');
}

// ------------------------------------------------------------ 3. nguồn
function scStatus() {
  const pages = scPages();
  const nf = pages.reduce((a, p) => a + (p.forms || []).length, 0), nb = pages.reduce((a, p) => a + (p.buttons || []).length, 0);
  $('sc-status').innerHTML = pages.length
    ? `✔ Đã dò <b>${pages.length}</b> trang · ${nf} form · ${nb} nút — lúc ${esc((SC.disc.when || '').replace('T', ' ').slice(0, 16))} (${pages.map((p) => esc(p.title)).join(', ')}). Trang web thay đổi thì bấm Dò trang lại.`
    : 'Nhập địa chỉ web rồi bấm <b>🔍 Dò trang</b>. Tool tự đọc các trang, nút, form, bảng… rồi lập danh sách bên dưới. <a href="#" id="sc-demo">Dùng thử web demo</a>';
  const d = $('sc-demo');
  if (d) d.onclick = (e) => { e.preventDefault(); scDemo(); };
  renderDiag();
}
function renderDiag() {
  const list = (SC.disc && SC.disc.diagnosis) || [];
  const box = $('sc-diag');
  if (!list.length) { box.innerHTML = ''; return; }
  const acts = new Set(list.flatMap((d) => d.actions || []));
  box.innerHTML = `<div class="diag">${list.map((d) => `<div class="dg"><b>⚠ ${esc(d.title)}</b><div>${esc(d.detail)}</div>
      ${(d.tips || []).length ? `<ul>${d.tips.map((t) => `<li>${esc(t)}</li>`).join('')}</ul>` : ''}</div>`).join('')}
    <div class="row wrap">${acts.has('wait') ? '<button class="btn small" id="dg-wait">🔁 Dò lại – chờ lâu hơn</button>' : ''}
      ${acts.has('manual') ? '<button class="btn small" id="dg-manual">🔑 Tôi tự đăng nhập / vượt captcha rồi dò lại</button>' : ''}</div></div>`;
  const w = $('dg-wait'); if (w) w.onclick = () => { SC.cfg.wait_extra = 25; scSave(); scStart('discover'); };
  const m = $('dg-manual'); if (m) m.onclick = async () => {
    $('sc-login').value = 'manual'; scRead(); renderKind();
    if (S.state.settings.headless) { S.state.settings.headless = false; await guard(() => api('/api/settings', { headless: false })); }
    scSave(); scStart('discover');
  };
}
function scDemo() {
  $('sc-url').value = S.state.demo_url + '/don-hang.html';
  $('sc-login').value = 'auto'; $('sc-user').value = 'admin@demo.vn'; $('sc-lurl').value = S.state.demo_url + '/login.html';
  scRead(); renderKind(); scSave();
  toast('Đã điền web demo. Mật khẩu demo: 123456');
  $('sc-pass').focus();
}

// ------------------------------------------------------------ 4. nhóm test
function ts(key) {
  const t = S.state.scan_types.find((x) => x.key === key);
  if (!SC.cfg.types[key]) SC.cfg.types[key] = { on: t.on, opts: {} };
  const v = SC.cfg.types[key]; v.opts = v.opts || {};
  t.options.forEach((o) => { if (v.opts[o.key] === undefined) v.opts[o.key] = clone(o.default); });
  return v;
}
function optHtml(t, o) {
  const v = ts(t.key).opts[o.key];
  if (o.type === 'select') return `<label>${esc(o.label)}<select data-t="${t.key}" data-o="${o.key}">${o.choices.map(([cv, cl]) => `<option value="${esc(cv)}" ${String(v) === String(cv) ? 'selected' : ''}>${esc(cl)}</option>`).join('')}</select></label>`;
  if (o.type === 'multi') return `<div class="multi"><div class="lbl">${esc(o.label)}</div>${o.choices.map(([cv, cl]) => `<label class="chk"><input type="checkbox" data-t="${t.key}" data-o="${o.key}" data-v="${esc(cv)}" ${(v || []).includes(cv) ? 'checked' : ''}><span>${esc(cl)}</span></label>`).join('')}</div>`;
  if (o.type === 'password') return `<label>${esc(o.label)}<input type="password" data-t="${t.key}" data-o="${o.key}" data-pw="1" value="${esc(SC.pw.perm)}" autocomplete="new-password"></label>`;
  return `<label>${esc(o.label)}<input data-t="${t.key}" data-o="${o.key}" value="${esc(v || '')}" placeholder="${esc(o.hint || '')}"></label>`;
}
function bindOpts(box, after) {
  box.querySelectorAll('select[data-o], input[data-o]:not([type=checkbox])').forEach((el) => (el.oninput = el.onchange = () => {
    if (el.dataset.pw) { SC.pw.perm = el.value; return; }
    ts(el.dataset.t).opts[el.dataset.o] = el.value; scSave(); if (after) after();
  }));
  box.querySelectorAll('input[type=checkbox][data-v]').forEach((el) => (el.onchange = () => {
    const v = ts(el.dataset.t); const arr = new Set(v.opts[el.dataset.o] || []);
    el.checked ? arr.add(el.dataset.v) : arr.delete(el.dataset.v); v.opts[el.dataset.o] = [...arr]; scSave(); if (after) after();
  }));
}
const EXTRA = { business: ['business', 'Khai báo luật…', () => `${SC.cfg.rules.length} luật`], permission: ['permission', 'Ma trận phân quyền…', () => `${(SC.cfg.perm.rows || []).length} dòng`], state: ['state', 'Luồng trạng thái…', () => `${SC.cfg.state_rules.length} luồng`] };
function renderGroups() {
  const fromFile = isTC() && SC.cfg.source === 'file';
  const types = S.state.scan_types;
  const nOn = types.filter((t) => ts(t.key).on).length;
  $('grp-count').textContent = fromFile ? '— đang lấy scenario từ file: tool dùng mọi nhóm để ghép test case' : `— đã chọn ${nOn}/${types.length}`;
  $('grp-list').classList.toggle('dim', fromFile);
  $('grp-list').innerHTML = types.map((t) => {
    const v = ts(t.key), ex = EXTRA[t.key];
    const opts = t.options.filter((o) => t.key !== 'permission' || !['user', 'pass', 'login_mode'].includes(o.key));
    return `<div class="gcard ${v.on ? 'on' : ''}" data-t="${t.key}">
      <label class="gtop"><input type="checkbox" class="g-on" data-t="${t.key}" ${v.on ? 'checked' : ''}><span><b>${esc(t.name)}</b><span class="gvi">${esc(t.vi)}</span></span></label>
      <details><summary>Tùy chọn${ex ? ' · ' + esc(ex[2]()) : ''}</summary><div class="muted small">${esc(t.desc)}</div>
        <div class="opts">${opts.map((o) => optHtml(t, o)).join('')}</div>
        ${ex ? `<button class="btn small" data-adv="${ex[0]}">${esc(ex[1])}</button>` : ''}</details></div>`;
  }).join('');
  const box = $('grp-list');
  box.querySelectorAll('.g-on').forEach((x) => (x.onchange = () => { ts(x.dataset.t).on = x.checked; x.closest('.gcard').classList.toggle('on', x.checked); scSave(); scPlan(); renderGroupCount(); }));
  bindOpts(box, scPlan);
  box.querySelectorAll('[data-adv]').forEach((b) => (b.onclick = () => openAdv(b.dataset.adv)));
}
function renderGroupCount() { const types = S.state.scan_types; $('grp-count').textContent = `— đã chọn ${types.filter((t) => ts(t.key).on).length}/${types.length}`; }

// ------------------------------------------------------------ 5. xem trước
function scPlan() {
  clearTimeout(SC.planT);
  SC.planT = setTimeout(async () => {
    scRead();
    const r = await api('/api/scan/plan', { cfg: planCfg() }).catch((e) => { toast(e.message, true); return null; });
    if (!r) return;
    SC.plan = r.scenarios || [];
    SC.meta = r.meta || null;
    SC.needDiscover = r.need_discover;
    SC.plan.forEach((x) => {
      if (!SC.known.has(x.key)) { SC.known.add(x.key); SC.sel.add(x.key); }
      x.tc_items.forEach((t) => { if (!SC.knownTc.has(t.tc_id)) { SC.knownTc.add(t.tc_id); SC.selTc.add(t.tc_id); } });
    });
    pvRender(); namePreview();
  }, 150);
}
function planCfg() {
  const c = clone(SC.cfg);
  if (c.login) delete c.login.pass;
  return c;
}
function edited(x, f) { const e = SC.cfg.edits[x.key]; return e && e[f] !== undefined ? e[f] : x[f]; }
function pvRows() {
  const nh = $('pv-nhom').value, md = $('pv-mode').value;
  const okMode = (m) => !md || (md === 'Tay' ? m === 'Tay' : m !== 'Tay');
  return SC.plan.filter((x) => (!nh || x.nhom === nh) && (isTC() ? x.tc_items.some((t) => okMode(t.mode)) : okMode(x.mode)));
}
function selPlan() {
  if (!isTC()) return SC.plan.filter((x) => SC.sel.has(x.key));
  return SC.plan.filter((x) => x.tc_items.some((t) => SC.selTc.has(t.tc_id)));
}
function lines(arr) { return (arr || []).filter((s) => s).map((s, i) => `${i + 1}. ${s}`).join('\n'); }
function tcVal(key, x, t) {
  const rr = SC.runRes[t.tc_id];
  const rnd = $('x2-round').value || '1';
  switch (key) {
    case 'tc_id': return t.tc_id;
    case 'scenario_id': return x.id;
    case 'requirement': return `${x.id} / ${x.scenario}`;
    case 'module': return t.module;
    case 'man_hinh': case 'mo_ta': return t.url;
    case 'tieu_de': return t.title;
    case 'loai_test': return t.nhom;
    case 'priority': return t.priority;
    case 'precondition': return t.pre;
    case 'test_data': return t.data;
    case 'data_id': return t.data_id;
    case 'moi_truong': return t.env;
    case 'kich_thuoc': return t.size;
    case 'cach_test': return t.mode;
    case 'buoc': return rr ? lines(rr.steps.map((s) => s.fields.buoc)) : lines(t.steps);
    case 'expected': return rr ? lines(rr.steps.map((s) => s.fields.expected)) : lines(t.expected);
    case 'actual': return rr ? lines(rr.steps.filter((s) => s.status !== 'Pass').map((s) => `[${s.status}] ${s.actual || ''}`)) : '';
    case 'status': return rr ? rr.status : '';
    case 'kq_lan1': case 'kq_lan2': case 'kq_lan3': return rr && key === 'kq_lan' + rnd ? rr.status : '';
    default: return '';
  }
}
function pvRender() {
  if (!SC.cfg) return;
  const t = $('pv-table'), k = kind();
  const nh = $('pv-nhom'), cur = nh.value;
  const groups = [...new Set(SC.plan.map((x) => x.nhom))];
  nh.innerHTML = '<option value="">Nhóm: tất cả</option>' + groups.map((g) => `<option ${g === cur ? 'selected' : ''}>${esc(g)}</option>`).join('');
  const rows = pvRows();
  if (!SC.plan.length) {
    t.innerHTML = '';
    $('pv-empty').innerHTML = isTC() && SC.cfg.source === 'file' && !SC.cfg.scenarios_in.length
      ? 'Bấm <b>📂 Chọn file Test Scenario</b> (bước 3) để lấy danh sách scenario của bạn.'
      : (SC.needDiscover ? '👆 Chưa có dữ liệu — hãy nhập địa chỉ web và bấm <b>🔍 Dò trang</b> ở bước 3.' : 'Không có dòng nào — hãy chọn thêm nhóm test ở bước 4.');
    $('pv-count').textContent = ''; $('pv-hint').textContent = '';
    updButtons(); return;
  }
  $('pv-empty').innerHTML = SC.needDiscover && SC.cfg.source === 'file' ? '<span class="warn-text small">Chưa dò trang: các scenario đang tạm là "Test tay" — bấm 🔍 Dò trang để tool ghép bài kiểm tra tự động.</span>' : '';
  const cols = S.profile[k].filter((c) => c.on);
  let html = `<thead><tr><th class="ck"><input type="checkbox" id="pv-all"></th>${cols.map((c) => `<th>${esc(c.label)}</th>`).join('')}</tr></thead><tbody>`;
  if (k === 'scenario') {
    const selOrder = SC.plan.filter((x) => SC.sel.has(x.key));
    const num = {}; selOrder.forEach((x, i) => (num[x.key] = x.imported ? x.id : `TS_${(SC.meta && SC.meta.code) || 'WEB'}_${String(i + 1).padStart(3, '0')}`));
    let lastLu = null, luNo = 0;
    const luMh = {}; SC.plan.forEach((x) => { (luMh[x.luong] = luMh[x.luong] || new Set()); if (x.mh) luMh[x.luong].add(x.mh); });
    rows.forEach((x) => {
      if (x.luong !== lastLu) {
        lastLu = x.luong; luNo++;
        const inG = rows.filter((y) => y.luong === x.luong), all = inG.every((y) => SC.sel.has(y.key));
        const mh = [...(luMh[x.luong] || [])];
        html += `<tr class="grp-row"><td class="ck"><input type="checkbox" class="pg" data-lu="${x.luong}" ${all ? 'checked' : ''}></td><td colspan="${cols.length}">LUỒNG ${luNo} – ${esc(x.luong_name || 'KHÁC')}${mh.length ? ` (${esc(mh.join(', '))})` : ''}</td></tr>`;
      }
      const on = SC.sel.has(x.key);
      html += `<tr class="${on ? '' : 'off'}" data-k="${esc(x.key)}"><td class="ck"><input type="checkbox" class="pk" ${on ? 'checked' : ''}></td>` + cols.map((c) => {
        switch (c.key) {
          case 'ts_id': return `<td class="nowrap c"><b>${on ? esc(num[x.key]) : '—'}</b></td>`;
          case 'nhom': return `<td class="c"><span class="tag">${esc(edited(x, 'nhom'))}</span></td>`;
          case 'scenario': return `<td class="txt"><div class="ed" contenteditable="true" spellcheck="false" data-f="scenario">${esc(edited(x, 'scenario'))}</div>${x.mode === 'Tay' ? `<span class="hand" title="${esc(x.reason || 'Tool chưa tự kiểm tra được — cần test tay')}">✋ cần test tay</span>` : ''}</td>`;
          case 'ac': return `<td class="c"><div class="ed" contenteditable="true" spellcheck="false" data-f="ac" data-ph="AC_…">${esc(edited(x, 'ac'))}</div></td>`;
          case 'priority': { const p = edited(x, 'priority'); return `<td class="c"><select class="prio p-${esc(p)}" data-f="priority">${[...new Set([...PRIOS, p])].map((v) => `<option ${v === p ? 'selected' : ''}>${esc(v)}</option>`).join('')}</select></td>`; }
          case 'module': return `<td>${esc([x.mh, x.module].filter(Boolean).join(' – '))}</td>`;
          case 'luong': return `<td>${esc(x.luong_name)}</td>`;
          case 'cach_test': return `<td>${MODE_BADGE[x.mode] || ''}</td>`;
          case 'requirement': return '<td></td>';
          default: return '<td class="muted"></td>';
        }
      }).join('') + '</tr>';
    });
  } else {
    rows.forEach((x) => {
      const md = $('pv-mode').value;
      const items = x.tc_items.filter((tt) => !md || (md === 'Tay' ? tt.mode === 'Tay' : tt.mode !== 'Tay'));
      const all = items.every((tt) => SC.selTc.has(tt.tc_id));
      html += `<tr class="grp-row"><td class="ck"><input type="checkbox" class="pg" data-k="${esc(x.key)}" ${all ? 'checked' : ''}></td><td colspan="${cols.length}"><span class="sticky-l">${esc(x.id)} — ${esc(x.scenario)} <span class="muted">(${esc(x.nhom)})</span></span></td></tr>`;
      items.forEach((tt) => {
        const on = SC.selTc.has(tt.tc_id);
        html += `<tr class="${on ? '' : 'off'}" data-tc="${esc(tt.tc_id)}"><td class="ck"><input type="checkbox" class="pt" ${on ? 'checked' : ''}></td>` + cols.map((c) => {
          const v = tcVal(c.key, x, tt);
          if (['status', 'kq_lan1', 'kq_lan2', 'kq_lan3'].includes(c.key)) return `<td class="c">${v ? `<span class="st st-${stCls(v)}">${esc(v)}</span>` : ''}</td>`;
          if (c.key === 'cach_test') return `<td>${MODE_BADGE[v] || ''}</td>`;
          if (c.key === 'tieu_de' || c.key === 'requirement') return `<td class="txt">${esc(v)}</td>`;
          if (['buoc', 'expected', 'actual'].includes(c.key)) return `<td class="pre">${esc(v)}</td>`;
          if (c.key === 'tc_id') return `<td class="nowrap"><b>${esc(v)}</b></td>`;
          return `<td>${esc(v)}</td>`;
        }).join('') + '</tr>';
      });
    });
  }
  t.innerHTML = html + '</tbody>';
  // đếm
  if (k === 'scenario') {
    const sel = SC.plan.filter((x) => SC.sel.has(x.key)), man = sel.filter((x) => x.mode === 'Tay').length;
    const nLu = new Set(sel.map((x) => x.luong)).size;
    $('pv-count').textContent = `— ${SC.plan.length} scenario · đã chọn ${sel.length} (${nLu} luồng${man ? `, ${man} cần test tay` : ''})`;
    $('pv-hint').textContent = 'ID đánh liên tục theo các dòng đã chọn. Bấm vào ô Test Scenario / AC liên quan để sửa trực tiếp.';
  } else {
    const all = SC.plan.flatMap((x) => x.tc_items), sel = all.filter((tt) => SC.selTc.has(tt.tc_id)), man = sel.filter((tt) => tt.mode === 'Tay').length;
    $('pv-count').textContent = `— ${all.length} test case trong ${SC.plan.length} scenario · đã chọn ${sel.length}${man ? ` (${man} cần test tay)` : ''}`;
    $('pv-hint').textContent = Object.keys(SC.runRes).length ? 'Cột kết quả đang hiện kết quả lần chạy gần nhất.' : 'Steps / Kết quả mong đợi là các bước dự kiến — sau khi ▶ Chạy test, tool ghi bước và kết quả thật.';
  }
  const visKeys = k === 'scenario' ? rows.map((x) => x.key) : [...t.querySelectorAll('tr[data-tc]')].map((r) => r.dataset.tc);
  const set = k === 'scenario' ? SC.sel : SC.selTc;
  const allBox = $('pv-all');
  allBox.checked = visKeys.length && visKeys.every((x) => set.has(x));
  allBox.onchange = () => { visKeys.forEach((x) => (allBox.checked ? set.add(x) : set.delete(x))); pvRender(); };
  t.querySelectorAll('.pk').forEach((c) => (c.onchange = () => { const key = c.closest('tr').dataset.k; c.checked ? SC.sel.add(key) : SC.sel.delete(key); pvRender(); }));
  t.querySelectorAll('.pt').forEach((c) => (c.onchange = () => { const id = c.closest('tr').dataset.tc; c.checked ? SC.selTc.add(id) : SC.selTc.delete(id); pvRender(); }));
  t.querySelectorAll('.pg').forEach((c) => (c.onchange = () => {
    if (c.dataset.lu) rows.filter((y) => String(y.luong) === c.dataset.lu).forEach((y) => (c.checked ? SC.sel.add(y.key) : SC.sel.delete(y.key)));
    else { const x = SC.plan.find((y) => y.key === c.dataset.k); x.tc_items.forEach((tt) => (c.checked ? SC.selTc.add(tt.tc_id) : SC.selTc.delete(tt.tc_id))); }
    pvRender();
  }));
  t.querySelectorAll('.ed').forEach((el) => (el.onblur = () => {
    const key = el.closest('tr').dataset.k, f = el.dataset.f, x = SC.plan.find((y) => y.key === key), v = el.innerText.replace(/\s+/g, ' ').trim();
    const e = SC.cfg.edits[key] = SC.cfg.edits[key] || {};
    if (v === (x[f] || '')) delete e[f]; else e[f] = v;
    if (!Object.keys(e).length) delete SC.cfg.edits[key];
    scSave();
  }));
  t.querySelectorAll('.ed').forEach((el) => (el.onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); el.blur(); } }));
  t.querySelectorAll('select.prio').forEach((s) => (s.onchange = () => {
    const key = s.closest('tr').dataset.k, x = SC.plan.find((y) => y.key === key);
    const e = SC.cfg.edits[key] = SC.cfg.edits[key] || {};
    if (s.value === x.priority) delete e.priority; else e.priority = s.value;
    s.className = 'prio p-' + s.value; scSave();
  }));
  updButtons();
}
function updButtons() {
  const n = isTC() ? SC.plan.flatMap((x) => x.tc_items).filter((tt) => SC.selTc.has(tt.tc_id)).length : SC.plan.filter((x) => SC.sel.has(x.key)).length;
  const busy = S.poll !== null;
  $('x2-new').disabled = !n || busy; $('x2-fill').disabled = !n || busy;
  $('x2-run').disabled = !n || busy || SC.cfg.source === 'file' && !SC.cfg.url;
  $('x2-run').textContent = busy && SC.running === 'run' ? '⏳ Đang chạy…' : `▶ Chạy ${n || ''} test case`;
  $('sc-discover').disabled = busy;
  $('x2-new').textContent = n ? `📄 Sinh file Excel mới (${n} ${isTC() ? 'test case' : 'scenario'})` : '📄 Sinh file Excel mới';
}

// ------------------------------------------------------------ 6. xuất file
function namePreview() {
  const k = kind(), m = SC.meta || {};
  const slug = unaccent(m.main || '').split(/[^A-Za-z0-9]+/).filter(Boolean).map((w) => w[0].toUpperCase() + w.slice(1)).join('') || 'Web';
  const d = new Date(), p2 = (n) => String(n).padStart(2, '0');
  const ctx = { man_hinh: slug, du_an: S.state.settings.project_name || slug, ma: m.code || '', ngay: `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}`, gio: `${p2(d.getHours())}${p2(d.getMinutes())}`, nguoi_test: S.state.settings.tester || '', lan_chay: $('x2-round').value || '1', moi_truong: '', trinh_duyet: S.state.settings.browser || '', module: slug };
  let name = ($('x2-pattern').value.trim() || (k === 'testcase' ? 'TestCase_{man_hinh}_{ngay}' : 'TestScenario_{man_hinh}_{ngay}')).replace(/\{([a-z_]+)\}/g, (a, b) => (b in ctx ? ctx[b] : a)).replace(/\.xlsx$/i, '');
  name = name.replace(/[\\/:*?"<>|\r\n\t]/g, '_').replace(/_+/g, '_') + '.xlsx';
  $('x2-preview').textContent = name;
}
function renderRunInfo() {
  const has = isTC() && SC.lastRun && Object.keys(SC.runRes).length;
  $('x2-withrun-wrap').hidden = !has;
  if (has) $('x2-runinfo').innerHTML = `Kèm kết quả lần chạy <b>${esc(SC.runInfo)}</b> (Steps thật, Kết quả thực tế, cột Lần ${esc($('x2-round').value)})`;
}
function saveOut() {
  const k = kind();
  SC.cfg.out[k === 'testcase' ? 'pattern_testcase' : 'pattern_scenario'] = $('x2-pattern').value.trim();
  SC.cfg.out.folder = $('x2-folder').value.trim();
  SC.cfg.out.round = $('x2-round').value;
  scSave(); namePreview();
}
function expBody(extra) {
  const k = kind();
  const b = { kind: k, cfg: planCfg(), profile: S.profile, edits: SC.cfg.edits, folder: $('x2-folder').value.trim(), round: $('x2-round').value,
    pattern_scenario: SC.cfg.out.pattern_scenario, pattern_testcase: SC.cfg.out.pattern_testcase };
  b.only = selPlan().map((x) => x.key);
  if (k === 'testcase') {
    b.only_tc = SC.plan.flatMap((x) => x.tc_items).filter((tt) => SC.selTc.has(tt.tc_id)).map((tt) => tt.tc_id);
    if (SC.lastRun && Object.keys(SC.runRes).length && $('x2-withrun').checked) b.run_id = SC.lastRun;
  }
  return Object.assign(b, extra || {});
}
async function x2New(overwrite) {
  saveOut();
  const btn = $('x2-new'); btn.disabled = true; const old = btn.textContent; btn.textContent = '⏳ Đang tạo file…';
  const r = await guard(() => api('/api/scan/export', expBody({ overwrite: overwrite || 'ask' })));
  btn.textContent = old; updButtons();
  if (!r) return;
  if (r.exists) {
    const a = await dialog({ title: 'Trùng tên file', warn: true, html: `<p>Đã có file:<br><b>${esc(r.path)}</b></p>`, buttons: [{ label: 'Hủy', value: null }, { label: 'Ghi đè file cũ', value: 'yes' }, { label: 'Lưu thành file mới (thêm số)', value: 'suffix', cls: 'primary' }] });
    if (a) x2New(a);
    return;
  }
  doneBox('✅ Đã sinh file Excel mới', r.info, r.path, r.folder);
}
function doneBox(title, info, path, folder) {
  $('x2-done').innerHTML = `<div class="okbox"><b>${esc(title)}</b> — ${esc(info || '')}<div class="f"><span class="grow">${esc(path)}</span>
    <button class="btn small primary" data-p="${esc(path)}">Mở file</button><button class="btn small" data-p="${esc(folder)}">📁 Mở thư mục</button></div></div>`;
  $('x2-done').querySelectorAll('[data-p]').forEach((b) => (b.onclick = () => guard(() => api('/api/open_path', { path: b.dataset.p }))));
  $('x2-done').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// ---------- 📥 điền vào file có sẵn
function fillOpen() {
  saveOut();
  SC.F = SC.F || { save: 'copy' };
  SC.F.mode = isTC() && expBody().run_id ? 'update' : (SC.F.mode === 'update' ? 'append' : SC.F.mode || 'append');
  $('fillm').hidden = false;
  fillRender();
}
function fillRender() {
  const F = SC.F, k = kind(), withRun = isTC() && !!expBody().run_id;
  const kl = k === 'testcase' ? 'Test Case' : 'Test Scenario';
  const sheets = F.sheets || [];
  const okSheets = sheets.filter((s) => s.kind);
  const score = (s) => s.cols.filter((c) => c.key).length * 2 - s.cols.filter((c) => !c.key).length;
  const best = okSheets.filter((s) => s.kind === k).sort((a, b) => score(b) - score(a))[0];
  if (F.sheets && !F.sheet) F.sheet = (best || okSheets[0] || {}).name;
  $('fillm-box').innerHTML = `<div class="mh"><h2>📥 Điền ${esc(kl)} vào file Excel sẵn có</h2></div><div class="mb">
    <div class="fstep"><div class="fs-h">1. Chọn file</div>
      <div class="row"><button class="btn" id="f-pick">📂 Chọn file…</button><input id="f-path" class="grow" placeholder="hoặc dán đường dẫn file .xlsx rồi bấm Mở" value="${esc(F.uploaded ? '' : F.path || '')}"><button class="btn" id="f-open">Mở</button></div>
      <div class="fdrop" id="f-drop">${F.path ? `📄 <b>${esc(F.name)}</b> <span class="muted small">${esc(F.uploaded ? '(file kéo-thả — tool đã chép vào thư mục của tool)' : F.path)}</span>` : 'hoặc <b>kéo-thả file Excel</b> vào đây'}</div></div>
    ${F.sheets ? `<div class="fstep"><div class="fs-h">2. Chọn sheet</div>${sheets.map((s) => `<label class="sheet ${s.kind ? '' : 'dis'} ${s.name === F.sheet ? 'on' : ''}"><input type="radio" name="f-sheet" value="${esc(s.name)}" ${s.name === F.sheet ? 'checked' : ''} ${s.kind ? '' : 'disabled'}>
        <span><b>${esc(s.name)}</b> ${s.kind ? `<span class="tag">${s.kind === 'testcase' ? 'Test Case' : 'Test Scenario'}</span>${s === best ? ' <span class="okc">✔ phù hợp nhất</span>' : ''}` : '<span class="muted small">— không thấy dòng tiêu đề</span>'}
        ${s.kind ? `<div class="muted small">Dòng tiêu đề ${s.header_row} · nhận ${s.cols.filter((c) => c.key).length}/${s.cols.length} cột: ${esc(s.cols.filter((c) => c.key).map((c) => c.header).join(', '))} · đã có ${s.rows} dòng${s.groups ? `, ${s.groups} dòng nhóm` : ''}</div>` : ''}</span></label>`).join('')}</div>
    <div class="fstep"><div class="fs-h">3. Cách ghi</div>${withRun
      ? `<label class="chk"><input type="radio" name="f-mode" value="update" checked><span>Điền kết quả vào cột <b>Kết quả test (Lần ${esc($('x2-round').value)})</b>, Kết quả thực tế, Trạng thái tổng theo <b>Mã TC</b> (và Tên TC). Test case chưa có trong file thì thêm dòng mới ở cuối.</span></label>`
      : `<label class="chk"><input type="radio" name="f-mode" value="append" ${F.mode !== 'replace' ? 'checked' : ''}><span><b>Ghi tiếp</b> vào cuối sheet${k === 'scenario' ? ' — ID nối tiếp file (VD đã có TS_DHTD_043 → ghi từ TS_DHTD_044)' : ''}</span></label>
         <label class="chk"><input type="radio" name="f-mode" value="replace" ${F.mode === 'replace' ? 'checked' : ''}><span><b>Thay toàn bộ</b> dữ liệu cũ trong sheet (giữ tiêu đề và định dạng)</span></label>`}</div>
    <div class="fstep"><div class="fs-h">4. Lưu</div>
      <label class="chk"><input type="radio" name="f-save" value="copy" ${F.save !== 'overwrite' ? 'checked' : ''}><span>Lưu thành <b>BẢN SAO</b> (file gốc giữ nguyên) — khuyên dùng. Bản sao lưu ở thư mục bước 6.</span></label>
      <label class="chk ${F.uploaded ? 'dis' : ''}"><input type="radio" name="f-save" value="overwrite" ${F.save === 'overwrite' ? 'checked' : ''} ${F.uploaded ? 'disabled' : ''}><span>Ghi đè thẳng vào file gốc (tool sẽ hỏi lại trước khi ghi)</span></label></div>` : ''}
    <p class="muted small">Tool nhận cột theo <b>tên tiêu đề</b> và chép định dạng (chữ, viền, màu) từ dòng có sẵn. Ảnh / biểu đồ có sẵn trong file có thể không được giữ khi lưu — vì vậy nên lưu bản sao.</p>
  </div><div class="mf"><button class="btn" id="f-cancel">Hủy</button><button class="btn primary big" id="f-go" ${F.sheet ? '' : 'disabled'}>📥 Điền vào file</button></div>`;
  const box = $('fillm-box');
  $('f-cancel').onclick = () => ($('fillm').hidden = true);
  $('f-pick').onclick = () => fillLoad({});
  $('f-open').onclick = () => { const p = $('f-path').value.trim(); if (p) fillLoad({ path: p }); };
  $('f-path').onkeydown = (e) => { if (e.key === 'Enter') $('f-open').onclick(); };
  const dz = $('f-drop');
  dz.ondragover = (e) => { e.preventDefault(); e.stopPropagation(); dz.classList.add('over'); };
  dz.ondragleave = () => dz.classList.remove('over');
  dz.ondrop = (e) => {
    e.preventDefault(); e.stopPropagation(); dz.classList.remove('over');
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (!f) return;
    if (!/\.xlsm?$/i.test(f.name)) return toast('Chỉ nhận file Excel .xlsx / .xlsm', true);
    const rd = new FileReader(); rd.onload = () => fillLoad({ name: f.name, data: String(rd.result).split(',')[1] }); rd.readAsDataURL(f);
  };
  box.querySelectorAll('input[name=f-sheet]').forEach((x) => (x.onchange = () => { F.sheet = x.value; fillRender(); }));
  box.querySelectorAll('input[name=f-mode]').forEach((x) => (x.onchange = () => (F.mode = x.value)));
  box.querySelectorAll('input[name=f-save]').forEach((x) => (x.onchange = () => (F.save = x.value)));
  $('f-go').onclick = () => fillGo(false);
}
async function fillLoad(body) {
  const r = await guard(() => api('/api/xlsx/sheets', Object.assign({ profile: S.profile }, body)));
  if (!r || r.cancel) return;
  Object.assign(SC.F, { path: r.path, name: r.name, uploaded: r.uploaded, sheets: r.sheets, sheet: null });
  if (r.uploaded) SC.F.save = 'copy';
  fillRender();
}
async function fillGo(confirmed) {
  const F = SC.F;
  const body = expBody({ path: F.path, sheet: F.sheet, mode: F.mode, save: F.save, confirm: !!confirmed });
  const go = $('f-go'); go.disabled = true; go.textContent = '⏳ Đang ghi…';
  const r = await guard(() => api('/api/scan/fill', body));
  if ($('f-go')) { go.disabled = false; go.textContent = '📥 Điền vào file'; }
  if (!r) return;
  if (r.confirm) {
    const ok = await dialog({ title: 'Ghi đè file gốc?', warn: true, html: `<p>Tool sẽ ghi thẳng vào:<br><b>${esc(r.path)}</b></p><p>Không hoàn tác được. Nếu file đang mở trong Excel, hãy đóng file trước.</p>`, buttons: [{ label: 'Không — lưu bản sao', value: 'copy' }, { label: 'Ghi đè file gốc', value: 'yes', cls: 'primary' }] });
    if (ok === 'yes') fillGo(true);
    else if (ok === 'copy') { F.save = 'copy'; fillRender(); }
    return;
  }
  $('fillm').hidden = true;
  const s = r.stats || {};
  const info = `sheet "${s.sheet}": ${s.updated ? `cập nhật ${s.updated} dòng, ` : ''}thêm ${s.appended} dòng${s.groups ? ` (${s.groups} dòng nhóm)` : ''}${s.id_from ? `, ID từ ${s.id_from}` : ''}`;
  doneBox('✅ Đã điền vào file Excel', info, r.path, r.folder);
}

// ------------------------------------------------------------ dò trang / chạy
function scBody(extra) {
  scRead();
  const cfg = clone(SC.cfg);
  cfg.login.pass = SC.pw.login;
  if (cfg.types.permission) cfg.types.permission.opts = Object.assign(cfg.types.permission.opts || {}, { pass: SC.pw.perm });
  Object.assign(cfg, extra || {});
  return { cfg, options: { browser: S.state.settings.browser, speed: S.state.settings.speed, headless: S.state.settings.headless } };
}
function scCheck() {
  scRead();
  if (!SC.cfg.url) { toast('Hãy nhập địa chỉ web cần test (bước 3).', true); $('sc-url').focus(); return false; }
  if (SC.cfg.login.mode === 'auto' && (!SC.cfg.login.user || !SC.pw.login)) { toast('Chọn "Tool tự đăng nhập" thì cần nhập tài khoản và mật khẩu.', true); $('sc-pass').focus(); return false; }
  return true;
}
window.scanProgressHost = () => (SC.running === 'run' ? $('scan-progress-host2') : $('scan-progress-host'));
async function scStart(kd, extra) {
  if (!scCheck()) return;
  const r = await guard(() => api(kd === 'discover' ? '/api/scan/discover' : '/api/scan/run', scBody(extra)));
  if (!r) return;
  SC.running = kd;
  S.runId = r.run_id; S.since = 0; S.manualShownFor = null;
  $('live').innerHTML = '';
  placeProgress('scan'); showProgress(true); startPoll();
  updButtons();
}
async function x2Run() {
  const chosen = SC.plan.flatMap((x) => x.tc_items).filter((tt) => SC.selTc.has(tt.tc_id));
  const man = chosen.filter((tt) => tt.mode === 'Tay').length;
  const w = scWarnings();
  const ok = await dialog({ title: `Chạy ${chosen.length} test case?`, warn: w.length > 0 || man > 0,
    html: `<p>Trên <b>${esc(SC.cfg.url)}</b>. Xong sẽ ghi vào cột <b>Kết quả test (Lần ${esc($('x2-round').value)})</b> khi bạn sinh / điền file.</p>
      ${man ? `<p class="warn-text">${man} test case cần TEST TAY — tool sẽ dừng lại hỏi bạn Pass / Fail.</p>` : ''}
      ${w.length ? `<ul>${w.map((x) => `<li class="warn-text">${esc(x)}</li>`).join('')}</ul>` : ''}`,
    buttons: [{ label: 'Quay lại', value: false }, { label: '▶ Chạy', value: true, cls: 'primary' }] });
  if (!ok) return;
  const all = SC.plan.flatMap((x) => x.tc_items).length === chosen.length;
  scStart('run', { only: [], only_tc: all ? [] : chosen.map((tt) => tt.tc_id) });
}
function scWarnings() {
  const t = (k) => (SC.cfg.types[k] || {}), o = (k, key) => ((t(k).opts || {})[key]);
  const w = [];
  if (t('validation').on && o('validation', 'submit') !== 'no') w.push('Có bấm gửi form với dữ liệu sai — chỉ chạy trên môi trường TEST.');
  if (o('functional', 'form_actions') === 'yes' || SC.cfg.state_rules.length) w.push('Có bấm Lưu / Gửi… → sẽ tạo dữ liệu thật.');
  if (t('permission').on && (o('permission', 'login_mode') || 'manual') === 'auto' && !SC.pw.perm) w.push('Phân quyền: chưa nhập mật khẩu tài khoản thứ 2 (Ma trận phân quyền…).');
  return w;
}
async function loadRun(runId, quiet) {
  const d = await api('/api/run/result?run_id=' + encodeURIComponent(runId)).catch(() => null);
  if (!d || d.kind !== 'scan') { if (!quiet) toast('Không đọc được kết quả lần chạy', true); return; }
  SC.lastRun = runId; SC.runRes = {};
  d.results.forEach((r) => (SC.runRes[r.tc_id] = r));
  SC.runInfo = (d.started || '').replace('T', ' ').slice(0, 16);
  renderRunInfo(); pvRender();
}
window.scanOnFinish = (run) => {
  const kd = SC.running;
  SC.running = null;
  updButtons();
  if (kd === 'discover') {
    guard(async () => {
      const d = await api('/api/scan/discovery', { run_id: run.run_id });
      SC.disc = d.discovery;
      scStatus(); scPlan();
      $('progress').hidden = true;
      const diag = SC.disc.diagnosis || [];
      if (diag.length) { $('sc-diag').scrollIntoView({ behavior: 'smooth', block: 'center' }); toast('⚠ ' + diag[0].title, true); }
      else toast(`Đã dò xong ${scPages().length} trang — xem bảng ở bước 5.`);
    });
    return true;
  }
  if (kd === 'run') {
    SC.cfg.last_run = run.run_id; scSave();
    loadRun(run.run_id, true);
    $('btn-goto-results').hidden = false;
    dialog({ title: run.state === 'xong' ? '✅ Đã chạy xong' : 'Đã dừng', html: `<p>${run.done} test case: ${Object.entries(run.counts).map(([k, v]) => `<span class="st st-${stCls(k)}">${esc(k)}: ${v}</span>`).join(' ')}</p><p class="muted small">Kết quả được ghi vào cột <b>Kết quả test (Lần ${esc($('x2-round').value)})</b>.</p>`,
      buttons: [{ label: 'Xem chi tiết', value: 'detail' }, { label: '📥 Điền vào file sẵn có', value: 'fill' }, { label: '📄 Sinh file Excel mới', value: 'new', cls: 'primary' }] })
      .then((v) => { if (v === 'new') x2New(); if (v === 'fill') fillOpen(); if (v === 'detail') switchTab('results'); });
    return true;
  }
  return false;
};
window.scanOnTab = () => { if (SC.cfg) { renderCols(); pvRender(); } };

// ------------------------------------------------------------ ⚙ cấu hình thêm (hộp thoại)
function needDisc() { return scPages().length ? '' : '<p class="warn-text">Hãy bấm 🔍 Dò trang trước để có danh sách trường / nút / form thật của trang.</p>'; }
const PAGES = {
  general() {
    const c = SC.cfg, st = S.state.settings;
    return `<div class="grid2">
      <label>Phạm vi dò <select id="a-depth"><option value="0">Chỉ trang này</option><option value="1">Trang này + các trang liên kết</option><option value="2">Sâu 2 cấp</option><option value="3">Sâu 3 cấp</option></select></label>
      <label>Số trang tối đa <select id="a-max"><option>5</option><option>10</option><option>20</option><option>50</option></select></label>
      <label>Trình duyệt <select id="a-browser"><option value="edge">Microsoft Edge</option><option value="chrome">Google Chrome</option></select></label>
      <label>Tốc độ <select id="a-speed"><option value="nhanh">Nhanh</option><option value="vua">Vừa (dễ theo dõi)</option><option value="cham">Chậm</option></select></label>
      <label>Chờ thêm khi trang chưa hiện nội dung (web SPA) <select id="a-wait"><option value="8">8 giây</option><option value="15">15 giây</option><option value="25">25 giây</option><option value="45">45 giây</option></select></label>
      </div>
      <label>Bỏ qua link có chữ (không mở các link này) <input id="a-skip" value="${esc(c.skip_words || '')}"></label>
      <label class="chk"><input type="checkbox" id="a-headless" ${st.headless ? 'checked' : ''}><span>Chạy ẩn (không hiện trình duyệt)</span></label>
      <label class="chk"><input type="checkbox" id="a-remember" ${c.login.remember !== false ? 'checked' : ''}><span>Nhớ phiên đăng nhập cho lần sau</span></label>
      <p class="muted small">Case tool không chắc chắn: <select id="a-unsure"><option value="ask">dừng lại hỏi tôi</option><option value="later">ghi "Chưa test" để tôi xem sau</option></select></p>`;
  },
  business() {
    const fields = fieldList(), rdef = S.state.scan_rules;
    return needDisc() + `<p class="muted small">Luật nghiệp vụ của từng trường. Tool tự sinh ca sai / biên / đúng. Ghi <b>AC</b> giống AC trong file scenario để tool ghép đúng scenario.</p>
      <table class="tbl ed" id="br-t"><tr><th>Trường</th><th>Luật</th><th>Giá trị 1</th><th>Giá trị 2</th><th>Thông báo / kết quả mong đợi</th><th>AC</th><th></th></tr>
      ${SC.cfg.rules.map((r, i) => { const d = rdef.find((x) => x[0] === r.rule) || ['', '', '', '', '']; const fk = r.page_url + '|' + r.form_sel + '|' + r.field_sel;
        return `<tr data-i="${i}"><td><select data-k="field">${fields.some((f) => f.key === fk) ? '' : `<option value="${esc(fk)}">${esc(r.field_label || '?')}</option>`}${fields.map((f) => `<option value="${esc(f.key)}" ${f.key === fk ? 'selected' : ''}>${esc(f.text)}</option>`).join('')}</select></td>
          <td><select data-k="rule">${rdef.map((x) => `<option value="${x[0]}" ${x[0] === r.rule ? 'selected' : ''}>${esc(x[1])}</option>`).join('')}</select></td>
          <td><input data-k="v1" value="${esc(r.v1 || '')}" placeholder="${esc(d[2] || '—')}" ${d[2] ? '' : 'disabled'}></td>
          <td><input data-k="v2" value="${esc(r.v2 || '')}" placeholder="${esc(d[3] || '—')}" ${d[3] ? '' : 'disabled'}></td>
          <td><input data-k="msg" value="${esc(r.msg || '')}"></td><td><input data-k="ac" value="${esc(r.ac || '')}" style="width:70px"></td>
          <td><button class="btn small" data-del="${i}">✕</button></td></tr>`; }).join('')}</table>
      <button class="btn small" id="br-add" ${fields.length ? '' : 'disabled'}>＋ Thêm luật</button>`;
  },
  permission() {
    const t = S.state.scan_types.find((x) => x.key === 'permission');
    const items = permItems(); const cur = {}; (SC.cfg.perm.rows || []).forEach((r) => (cur[permKey(r)] = r.expect));
    let last = '';
    return `<div class="grid2">${t.options.filter((o) => ['login_mode', 'user', 'pass'].includes(o.key)).map((o) => optHtml(t, o)).join('')}
      <label>Tên vai trò tài khoản 2 <input id="pm-role" value="${esc(SC.cfg.perm.role || '')}" placeholder="VD: Viewer / Nhân viên"></label></div>` + needDisc() +
      (items.length ? `<div class="row"><b class="grow">Ma trận phân quyền</b><button class="btn small" id="pm-allow">Tất cả = Được thấy</button><button class="btn small" id="pm-clear">Bỏ chọn</button></div>
      <table class="tbl ed perm"><tr><th>Trang</th><th>Vị trí</th><th>Chức năng / Nút</th><th>Vai trò '${esc(SC.cfg.perm.role || 'tài khoản 2')}'</th></tr>
      ${items.map((x, i) => { const v = cur[permKey(x)] || ''; const pg = x.page_title !== last ? x.page_title : ''; last = x.page_title;
        return `<tr class="${v}"><td>${esc(pg)}</td><td class="muted small">${x.where ? 'Popup ' + esc(x.where) : ''}</td><td>${x.button === '__page__' ? '<i>Truy cập trang</i>' : esc(x.button)}</td>
          <td><select data-pi="${i}"><option value="">— Không kiểm tra</option><option value="allow" ${v === 'allow' ? 'selected' : ''}>✔ Được ${x.button === '__page__' ? 'vào' : 'thấy'}</option><option value="deny" ${v === 'deny' ? 'selected' : ''}>✘ KHÔNG được ${x.button === '__page__' ? 'vào' : 'thấy'}</option></select></td></tr>`; }).join('')}</table>` : '');
  },
  state() {
    const forms = formList(); const allSt = [...new Set(scPages().flatMap((p) => p.statusValues || []))];
    return needDisc() + `<p class="muted small">Bấm nút nào trên form thì bản ghi phải ở trạng thái nào (sẽ tạo dữ liệu thật — chỉ dùng môi trường TEST).</p>
      <datalist id="st-values">${allSt.map((v) => `<option value="${esc(v)}">`).join('')}</datalist>
      <table class="tbl ed" id="st-t"><tr><th>Form</th><th>Bấm nút</th><th>Trạng thái mong đợi</th><th>AC</th><th></th></tr>
      ${SC.cfg.state_rules.map((r, i) => { const fk = r.page_url + '|' + r.form_sel; const f = forms.find((x) => x.key === fk);
        return `<tr data-i="${i}"><td><select data-k="form">${f ? '' : `<option value="${esc(fk)}">${esc(r.form_name || '?')}</option>`}${forms.map((x) => `<option value="${esc(x.key)}" ${x.key === fk ? 'selected' : ''}>${esc(x.page_title + ' › ' + x.form_name)}</option>`).join('')}</select></td>
          <td><select data-k="button">${((f && f.buttons) || [r.button]).map((b) => `<option ${b === r.button ? 'selected' : ''}>${esc(b)}</option>`).join('')}</select></td>
          <td><input data-k="expect" list="st-values" value="${esc(r.expect || '')}" placeholder="VD: Nháp"></td>
          <td><input data-k="ac" value="${esc(r.ac || '')}" style="width:70px"></td><td><button class="btn small" data-del="${i}">✕</button></td></tr>`; }).join('')}</table>
      <button class="btn small" id="st-add" ${forms.length ? '' : 'disabled'}>＋ Thêm luồng</button>`;
  },
  clicks() {
    return needDisc() + '<p class="muted small">Tick = cho tool bấm thử nút này. Bỏ tick các nút làm thay đổi dữ liệu thật (Xóa, Duyệt…).</p>' +
      scPages().map((p, pi) => `<div class="lbl">${esc(p.title)}</div><div class="chips-sel">${p.buttons.map((b, bi) => { const on = allowed(p, b);
        return `<label class="chipc ${on ? 'on' : ''}"><input type="checkbox" data-pi="${pi}" data-bi="${bi}" ${on ? 'checked' : ''}>${esc(b.text)}</label>`; }).join('') || '<span class="muted small">Không có nút</span>'}</div>`).join('');
  },
};
function allowed(p, b) { const m = SC.cfg.click[p.url] || {}; return (b.text in m) ? !!m[b.text] : !!b.allowed; }
function fieldList() {
  const out = [];
  scPages().forEach((p) => p.forms.forEach((f) => f.fields.forEach((x) => out.push({ key: p.url + '|' + f.sel + '|' + x.sel, page_url: p.url, page_title: p.title, form_sel: f.sel, field_sel: x.sel, field_label: x.label,
    text: `${p.title} › ${f.name} › ${x.label}${x.required ? ' *' : ''}` }))));
  return out;
}
function formList() {
  const out = [];
  scPages().forEach((p) => p.forms.forEach((f) => out.push({ key: p.url + '|' + f.sel, page_url: p.url, page_title: p.title, form_sel: f.sel, form_name: f.name, buttons: f.buttons })));
  return out;
}
function permItems() {
  const out = [];
  scPages().forEach((p) => {
    out.push({ page_url: p.url, page_title: p.title, where: '', button: '__page__' });
    p.buttons.forEach((b) => out.push({ page_url: p.url, page_title: p.title, where: '', button: b.text }));
    p.popups.forEach((x) => x.buttons.forEach((t) => out.push({ page_url: p.url, page_title: p.title, where: x.opener, button: t })));
  });
  return out;
}
function permKey(x) { return x.page_url + '|' + (x.where || '') + '|' + x.button; }
function openAdv(p) { SC.adv = p || 'general'; $('adv').hidden = false; advRender(); }
function advRender() {
  document.querySelectorAll('#adv-nav button').forEach((b) => b.classList.toggle('on', b.dataset.p === SC.adv));
  const box = $('adv-body');
  box.innerHTML = PAGES[SC.adv]();
  bindOpts(box);
  const c = SC.cfg, st = S.state.settings;
  if (SC.adv === 'general') {
    $('a-depth').value = String(c.depth ?? 1); $('a-max').value = String(c.max_pages || 10);
    $('a-browser').value = st.browser === 'chrome' ? 'chrome' : 'edge'; $('a-speed').value = st.speed || 'vua';
    $('a-wait').value = String(c.wait_extra || 8);
    $('a-unsure').value = ts('validation').opts.unsure || 'ask';
    const saveRun = () => { st.browser = $('a-browser').value; st.speed = $('a-speed').value; st.headless = $('a-headless').checked;
      guard(() => api('/api/settings', { browser: st.browser, speed: st.speed, headless: st.headless })); };
    ['a-browser', 'a-speed', 'a-headless'].forEach((id) => ($(id).onchange = saveRun));
    $('a-depth').onchange = () => { c.depth = +$('a-depth').value; scSave(); };
    $('a-max').onchange = () => { c.max_pages = +$('a-max').value; scSave(); };
    $('a-wait').onchange = () => { c.wait_extra = +$('a-wait').value; scSave(); };
    $('a-skip').oninput = () => { c.skip_words = $('a-skip').value; scSave(); };
    $('a-remember').onchange = () => { c.login.remember = $('a-remember').checked; scSave(); };
    $('a-unsure').onchange = () => { ['validation', 'business'].forEach((k) => (ts(k).opts.unsure = $('a-unsure').value)); scSave(); };
  }
  if (SC.adv === 'business') {
    const fields = fieldList();
    const add = $('br-add'); if (add) add.onclick = () => { const f = fields[0]; c.rules.push({ page_url: f.page_url, page_title: f.page_title, form_sel: f.form_sel, field_sel: f.field_sel, field_label: f.field_label, rule: 'required' }); scSave(); advRender(); };
    box.querySelectorAll('#br-t tr[data-i]').forEach((tr) => { const r = c.rules[+tr.dataset.i];
      tr.querySelectorAll('[data-k]').forEach((el) => (el.onchange = el.oninput = () => {
        if (el.dataset.k === 'field') { const f = fields.find((x) => x.key === el.value); if (f) Object.assign(r, { page_url: f.page_url, page_title: f.page_title, form_sel: f.form_sel, field_sel: f.field_sel, field_label: f.field_label }); }
        else r[el.dataset.k] = el.value;
        scSave(); if (el.dataset.k === 'rule' && el.tagName === 'SELECT') advRender();
      }));
      tr.querySelector('[data-del]').onclick = () => { c.rules.splice(+tr.dataset.i, 1); scSave(); advRender(); };
    });
  }
  if (SC.adv === 'state') {
    const forms = formList();
    const add = $('st-add'); if (add) add.onclick = () => { const f = forms[0]; c.state_rules.push({ page_url: f.page_url, page_title: f.page_title, form_sel: f.form_sel, form_name: f.form_name, button: f.buttons[f.buttons.length - 1] || '', expect: '' }); scSave(); advRender(); };
    box.querySelectorAll('#st-t tr[data-i]').forEach((tr) => { const r = c.state_rules[+tr.dataset.i];
      tr.querySelectorAll('[data-k]').forEach((el) => (el.onchange = el.oninput = () => {
        if (el.dataset.k === 'form') { const f = forms.find((x) => x.key === el.value); if (f) Object.assign(r, { page_url: f.page_url, page_title: f.page_title, form_sel: f.form_sel, form_name: f.form_name, button: f.buttons[f.buttons.length - 1] || '' }); scSave(); advRender(); return; }
        r[el.dataset.k] = el.value; scSave();
      }));
      tr.querySelector('[data-del]').onclick = () => { c.state_rules.splice(+tr.dataset.i, 1); scSave(); advRender(); };
    });
  }
  if (SC.adv === 'permission') {
    const items = permItems();
    const setRows = (fn) => { const cur = {}; (c.perm.rows || []).forEach((r) => (cur[permKey(r)] = r.expect)); fn(cur);
      c.perm.rows = items.filter((x) => cur[permKey(x)]).map((x) => Object.assign({}, x, { expect: cur[permKey(x)] })); scSave(); };
    const role = $('pm-role'); role.oninput = () => { c.perm.role = role.value; scSave(); };
    box.querySelectorAll('.perm select[data-pi]').forEach((sel) => (sel.onchange = () => { setRows((cur) => { const k = permKey(items[+sel.dataset.pi]); if (sel.value) cur[k] = sel.value; else delete cur[k]; }); sel.closest('tr').className = sel.value; }));
    const al = $('pm-allow'); if (al) al.onclick = () => { setRows((cur) => items.forEach((x) => (cur[permKey(x)] = 'allow'))); advRender(); };
    const cl = $('pm-clear'); if (cl) cl.onclick = () => { setRows((cur) => Object.keys(cur).forEach((k) => delete cur[k])); advRender(); };
  }
  if (SC.adv === 'clicks') box.querySelectorAll('.chips-sel input').forEach((x) => (x.onchange = () => {
    const p = scPages()[+x.dataset.pi], b = p.buttons[+x.dataset.bi];
    c.click[p.url] = c.click[p.url] || {}; c.click[p.url][b.text] = x.checked; x.parentElement.classList.toggle('on', x.checked); scSave();
  }));
}

// ------------------------------------------------------------ sự kiện
function scBind() {
  document.querySelectorAll('#kinds .kind').forEach((b) => (b.onclick = () => { SC.cfg.kind = b.dataset.kind; $('x2-done').innerHTML = ''; renderKind(); scSave(); scPlan(); }));
  $('col-prof').onchange = async () => {
    if (S.profileDirty && !(await dialog({ title: 'Bỏ thay đổi?', html: '<p>Mẫu cột hiện tại có thay đổi chưa lưu.</p>', buttons: [{ label: 'Ở lại', value: false }, { label: 'Bỏ thay đổi', value: true, cls: 'primary' }] }))) { $('col-prof').value = S.profile.name; return; }
    await guard(() => api('/api/profile/activate', { name: $('col-prof').value })); S.profileDirty = false; await loadState(); renderCols(); pvRender();
  };
  $('col-save').onclick = () => colSave();
  $('col-saveas').onclick = () => { const name = prompt('Tên mẫu cột mới (VD: Dự án Đơn hàng tuyển dụng):', S.profile.name + ' (bản sao)'); if (name && name.trim()) colSave(name.trim()); };
  $('col-reset').onclick = () => { S.profile[kind()] = companyRows(kind()); S.profileDirty = true; renderCols(); pvRender(); };
  $('sc-login').onchange = () => { scRead(); renderKind(); scSave(); scPlan(); };
  ['sc-user', 'sc-lurl'].forEach((id) => ($(id).onchange = () => { scRead(); scSave(); }));
  $('sc-url').onchange = async () => {
    scRead();
    const r = await api('/api/scan/save', { cfg: { url: SC.cfg.url } }).catch(() => null);
    if (r) { SC.disc = r.discovery; scStatus(); scPlan(); }
  };
  $('sc-url').onkeydown = (e) => { if (e.key === 'Enter') { $('sc-url').onchange(); scStart('discover'); } };
  $('sc-discover').onclick = () => scStart('discover');
  document.querySelectorAll('#sc-source button').forEach((b) => (b.onclick = () => { SC.cfg.source = b.dataset.src; renderKind(); scSave(); scPlan(); }));
  $('sc-import').onclick = async () => {
    const r = await guard(() => api('/api/scan/import', {}));
    if (r && !r.cancel) { SC.cfg.scenarios_in = r.scenarios_in; SC.cfg.scenarios_file = r.file; SC.cfg.source = 'file'; renderKind(); scPlan(); toast(`Đã lấy ${r.count} Test Scenario từ file.`); }
  };
  document.querySelectorAll('[data-preset]').forEach((b) => (b.onclick = () => {
    const pre = b.dataset.preset, P = S.state.scan_presets || {};
    S.state.scan_types.forEach((t) => (ts(t.key).on = pre === 'all' ? true : pre === 'none' ? false : !!(P.seven || {})[t.key]));
    renderGroups(); scSave(); scPlan();
  }));
  $('g-suggest').onchange = () => { SC.cfg.suggest = $('g-suggest').checked; scSave(); scPlan(); };
  document.querySelectorAll('.adv-links [data-adv]').forEach((b) => (b.onclick = () => openAdv(b.dataset.adv)));
  $('adv-close').onclick = () => { $('adv').hidden = true; scSave(); renderGroups(); scPlan(); };
  document.querySelectorAll('#adv-nav button').forEach((b) => (b.onclick = () => { SC.adv = b.dataset.p; advRender(); }));
  $('pv-nhom').onchange = pvRender; $('pv-mode').onchange = pvRender;
  $('x2-pattern').oninput = () => { clearTimeout(SC.outT); SC.outT = setTimeout(saveOut, 300); namePreview(); };
  $('x2-folder').onchange = saveOut;
  $('x2-round').onchange = () => { saveOut(); renderRunInfo(); pvRender(); };
  $('x2-pick').onclick = async () => { const r = await guard(() => api('/api/pick_folder', {})); if (r && r.folder) { $('x2-folder').value = r.folder; saveOut(); } };
  $('x2-new').onclick = () => x2New();
  $('x2-fill').onclick = fillOpen;
  $('x2-run').onclick = x2Run;
}

(function waitState() {
  if (S.state && S.state.scan_types) { scBind(); scInit(); return; }
  setTimeout(waitState, 100);
})();
