"""Chế độ QUÉT TỪ URL: chỉ cần địa chỉ web → tool tự dò các trang, form, nút, bảng… rồi TỰ SINH test case theo
từng loại test đã chọn (Functional, Validation, Business rule, State, Permission, UI, Responsive, Performance,
Security, Accessibility, Compatibility), tự chạy và ghi kết quả (cùng định dạng với chế độ chạy theo Excel).

Không dùng AI: mọi kiểm tra là luật có sẵn, chạy hoàn toàn trên máy (web nội bộ không bị gửi ra ngoài).
"""
import datetime as dt
import json
import os
import random
import re
import shutil
import sys
import time
import traceback
from urllib.parse import urlparse, urlunparse

import columns as C
import engine as E
import excel_io as X
import scan_types as T

# ----------------------------------------------------------------------------- JavaScript chạy trong trang
HELPERS = r"""
const __vis = (el) => { if (!el || !el.getBoundingClientRect) return false; const r = el.getBoundingClientRect();
  if (r.width < 1 || r.height < 1) return false; const cs = getComputedStyle(el);
  return cs.visibility !== 'hidden' && cs.display !== 'none' && parseFloat(cs.opacity || '1') > 0.05; };
const __esc = (s) => (window.CSS && CSS.escape) ? CSS.escape(s) : String(s).replace(/[^\w-]/g, '\\$&');
const __uniq = (s) => { try { return document.querySelectorAll(s).length === 1; } catch (e) { return false; } };
const __sel = (el) => {
  if (el.id && __uniq('#' + __esc(el.id))) return '#' + __esc(el.id);
  const tag = el.tagName.toLowerCase();
  for (const a of ['name', 'data-testid', 'data-test', 'aria-label']) { const v = el.getAttribute(a);
    if (v && v.length < 80) { const s = tag + '[' + a + '="' + v.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"]'; if (__uniq(s)) return s; } }
  const parts = []; let e = el;
  while (e && e.nodeType === 1 && e !== document.documentElement) {
    if (e !== el && e.id && __uniq('#' + __esc(e.id))) { parts.unshift('#' + __esc(e.id)); break; }
    let i = 1, s = e; while ((s = s.previousElementSibling)) if (s.tagName === e.tagName) i++;
    parts.unshift(e.tagName.toLowerCase() + ':nth-of-type(' + i + ')'); e = e.parentElement; }
  return parts.join(' > '); };
const __txt = (el) => ((el.innerText || el.value || el.getAttribute('aria-label') || el.getAttribute('title') ||
  el.getAttribute('alt') || '') + '').replace(/\s+/g, ' ').trim().slice(0, 80);
const __clean = (s) => (s || '').replace(/\s+/g, ' ').replace(/[\s*:：]+$/, '').replace(/^[\s*]+/, '').trim();
const __noCtl = (l) => { const c = l.cloneNode(true); c.querySelectorAll('input,select,textarea,option,button').forEach((x) => x.remove()); return c.textContent || ''; };
const __label = (el) => {
  const pick = (t, src) => ({text: __clean(t).slice(0, 80), star: /\*/.test(t || ''), src});
  let t = el.getAttribute('aria-label'); if (t) return pick(t, 'aria');
  const lb = el.getAttribute('aria-labelledby');
  if (lb) { t = lb.split(/\s+/).map((i) => (document.getElementById(i) || {}).innerText || '').join(' '); if (t.trim()) return pick(t, 'aria'); }
  if (el.id) { const l = document.querySelector('label[for="' + __esc(el.id) + '"]'); if (l && __noCtl(l).trim()) return pick(__noCtl(l), 'label'); }
  const pl = el.closest('label'); if (pl && __noCtl(pl).trim()) return pick(__noCtl(pl), 'label');
  let p = el.previousElementSibling;
  for (let i = 0; i < 3 && p; i++, p = p.previousElementSibling) { const tx = (p.innerText || '').trim();
    if (tx && tx.length < 60 && !/^(input|select|textarea|button)$/i.test(p.tagName)) return pick(tx, 'near'); }
  const par = el.parentElement; if (par) { const pp = par.previousElementSibling; if (pp && pp.innerText && pp.innerText.trim().length < 60) return pick(pp.innerText, 'near'); }
  t = el.getAttribute('placeholder') || el.getAttribute('title') || el.getAttribute('name') || el.id || '';
  return pick(t, 'placeholder'); };
const __fieldOk = (el) => { const t = (el.getAttribute('type') || el.tagName).toLowerCase(); return !['hidden', 'submit', 'button', 'reset', 'image'].includes(t); };
const __field = (el) => { const l = __label(el);
  const t = (el.getAttribute('type') || (el.tagName === 'INPUT' ? 'text' : el.tagName)).toLowerCase();
  const num = (a) => { const v = el.getAttribute(a); return v === null || v === '' || isNaN(Number(v)) ? null : Number(v); };
  return {sel: __sel(el), tag: el.tagName.toLowerCase(), type: t, name: el.getAttribute('name') || el.id || '', label: l.text,
    labelSrc: l.src, required: !!(el.required || el.getAttribute('aria-required') === 'true' || l.star),
    maxlength: el.maxLength > 0 ? el.maxLength : null, minlength: el.minLength > 0 ? el.minLength : null,
    min: num('min'), max: num('max'), pattern: el.getAttribute('pattern') || '', placeholder: el.getAttribute('placeholder') || '',
    accept: el.getAttribute('accept') || '', visible: __vis(el) || (t === 'file' && __vis(el.parentElement)), disabled: !!(el.disabled || el.readOnly),
    options: el.tagName === 'SELECT' ? [...el.options].slice(0, 40).map((o) => ({value: o.value, text: o.text.trim()})) : [],
    role: el.getAttribute('role') || ''}; };
const __SUBMIT_RE = /^(lưu|lưu lại|gửi|gửi đi|đăng nhập|đăng ký|tạo|tạo mới|thêm|thêm mới|xác nhận|cập nhật|tiếp tục|hoàn tất|submit|save|send|login|log in|sign in|sign up|register|create|add|ok|confirm|update|continue|tìm|tìm kiếm|search|lọc|áp dụng|apply|nhập dữ liệu|import|tải lên|upload)$/iu;
const __CANCEL_RE = /^(hủy|huỷ|đóng|close|cancel|quay lại|back|×|✕|x)$/iu;
const __SEARCH_RE = /tìm|search|lọc|filter|từ khóa|từ khoá|keyword/iu;
const __forms = () => { const out = []; const seen = new Set();
  const mk = (c, fl) => {
    const btns = [...c.querySelectorAll('button, input[type=submit], input[type=button], [role=button], a.btn')].filter((b) => !b.disabled);
    let submit = btns.find((b) => (b.getAttribute('type') || '').toLowerCase() === 'submit' || b.matches('input[type=submit]'));
    if (!submit) { const cand = btns.filter((b) => __SUBMIT_RE.test(__txt(b).replace(/^[+\s]+/, ''))); submit = cand[cand.length - 1]; }
    if (!submit && c.tagName === 'FORM') submit = btns.find((b) => b.tagName === 'BUTTON' && !b.getAttribute('type') && !__CANCEL_RE.test(__txt(b)));
    const fields = fl.map(__field); const stext = submit ? __txt(submit) : '';
    const head = c.querySelector('h1,h2,h3,h4,legend,[class*=title]');
    const isSearch = fields.some((f) => f.type === 'search') || (!!submit && __SEARCH_RE.test(stext)) ||
      (fields.length <= 3 && fields.every((f) => f.tag === 'select' || __SEARCH_RE.test(f.label + ' ' + f.placeholder)));
    const form = c.tagName === 'FORM' ? c : c.querySelector('form');
    const allBtns = btns.filter((b) => __txt(b)).map((b) => ({sel: __sel(b), text: __txt(b), type: (b.getAttribute('type') || '').toLowerCase()})).slice(0, 15);
    return {sel: __sel(c), tag: c.tagName.toLowerCase(), isForm: !!form, buttons: allBtns, name: head ? __txt(head) : (c.getAttribute('aria-label') || c.id || ''),
      visible: fl.some(__vis), dialog: !!c.closest('[role=dialog], dialog, [aria-modal=true], .modal, [class*="modal"], [class*="popup"]'),
      fields, submit: submit ? {sel: __sel(submit), text: stext} : null, kind: isSearch ? 'search' : ((!submit && !form) ? 'filter' : 'data'),
      method: ((form && form.getAttribute('method')) || '').toLowerCase(), hasPassword: fields.some((f) => f.type === 'password')}; };
  document.querySelectorAll('form').forEach((f) => { const fl = [...f.querySelectorAll('input, select, textarea')].filter(__fieldOk);
    fl.forEach((x) => seen.add(x)); if (fl.length) out.push(mk(f, fl)); });
  const groups = new Map();
  [...document.querySelectorAll('input, select, textarea')].filter((x) => __fieldOk(x) && !seen.has(x)).forEach((x) => {
    let c = x.closest('[role=dialog], dialog, [aria-modal=true], .modal, [class*="modal"], [class*="popup"], [class*="drawer"]');
    if (!c) { c = x.parentElement; while (c && c !== document.body && !c.querySelector('button, input[type=submit], [role=button]')) c = c.parentElement; }
    c = c || document.body; if (!groups.has(c)) groups.set(c, []); groups.get(c).push(x); });
  groups.forEach((fl, c) => out.push(mk(c, fl)));
  return out; };
const __DIALOG_SEL = '[role=dialog], dialog[open], [aria-modal=true], .modal, [class*="modal"], [class*="popup"], [class*="dialog"], [class*="drawer"], [class*="offcanvas"]';
"""


def _js(body, arg=False):
    return "(%s) => { %s\n%s }" % ("__a" if arg else "", HELPERS, body)


PAGE_INFO_JS = _js(r"""
  const links = [...document.querySelectorAll('a[href]')].map((a) => ({href: a.href, text: __txt(a), vis: __vis(a)})).slice(0, 600);
  const btnEls = [...new Set(document.querySelectorAll('button, input[type=button], input[type=submit], [role=button], [role=tab], [role=switch], [aria-expanded], summary, a.btn, a[class*="btn"]'))];
  const buttons = btnEls.filter((b) => __vis(b) && !b.disabled && b.getAttribute('aria-disabled') !== 'true' &&
      !(b.tagName === 'INPUT' && !['button', 'submit'].includes((b.getAttribute('type') || '').toLowerCase()))).map((b) => {
    const href = b.getAttribute('href') || '';
    return {sel: __sel(b), text: __txt(b), aria: b.getAttribute('aria-label') || '', type: (b.getAttribute('type') || '').toLowerCase(),
      tag: b.tagName.toLowerCase(), inForm: !!b.closest('form'), expanded: b.getAttribute('aria-expanded'), role: b.getAttribute('role') || '',
      href: (href && !href.startsWith('#') && !href.startsWith('javascript')) ? b.href : '', inTable: !!b.closest('table, [role=grid]')}; }).slice(0, 250);
  const TOG = /menu|☰|≡|bộ lọc|lọc nâng cao|nâng cao|advanced|filter|toggle|mở rộng|thu gọn|expand|collapse|hiện|ẩn|show|hide|xem thêm|more/iu;
  const toggles = buttons.filter((b) => b.expanded !== null || b.tag === 'summary' || b.role === 'switch' || TOG.test(b.text + ' ' + b.aria)).slice(0, 20);
  const switches = [...document.querySelectorAll('input[type=checkbox][role=switch], .switch input[type=checkbox], [class*="toggle"] input[type=checkbox], [class*="switch"] input[type=checkbox]')]
    .filter((el) => __vis(el) || __vis(el.parentElement)).slice(0, 10).map((el) => ({sel: __sel(el), text: __label(el).text}));
  const tables = [...document.querySelectorAll('table, [role=grid], [role=table]')].filter(__vis).slice(0, 5).map((t) => {
    const s = __sel(t); const isT = t.tagName === 'TABLE';
    const rowSel = isT ? (t.tBodies.length ? s + ' > tbody > tr' : s + ' tr') : s + ' [role=row]';
    const rows = [...document.querySelectorAll(rowSel)].filter(__vis);
    const ths = isT ? [...t.querySelectorAll('thead th')] : [...t.querySelectorAll('[role=columnheader]')];
    return {sel: s, rowSel, rows: rows.length, first: rows[0] ? [...rows[0].children].map((c) => (c.innerText || '').trim().slice(0, 60)) : [],
      headers: ths.map((h, i) => ({text: __txt(h), sel: __sel(h), idx: i, sortable: h.hasAttribute('aria-sort') ||
        /sort/i.test(h.className + ' ' + (h.getAttribute('data-sort') || '')) || !!h.querySelector('button, a, [class*=sort]') ||
        h.hasAttribute('onclick') || getComputedStyle(h).cursor === 'pointer'}))}; });
  const pager = [...document.querySelectorAll('a, button, [role=button]')].filter(__vis).filter((e) => { const t = __txt(e); const al = e.getAttribute('aria-label') || '';
    return (/^\d{1,3}$/.test(t) && !!e.closest('[class*=pag], [class*=pager], nav, [aria-label*=pag], [aria-label*=rang], ul, [id*=pag]')) ||
      /^(tiếp|tiếp theo|sau|trang sau|next|›|»|>|>>)$/iu.test(t) || /^(trang|page)\s*\d+$/iu.test(al) || /next|trang sau|tiếp/iu.test(al); })
    .slice(0, 20).map((e) => ({sel: __sel(e), text: __txt(e), aria: e.getAttribute('aria-label') || '', current: e.getAttribute('aria-current') || '', cls: e.className + ''}));
  const search = [...document.querySelectorAll('input')].filter(__vis).filter((i) => { const t = (i.getAttribute('type') || 'text').toLowerCase();
      if (!['text', 'search'].includes(t)) return false;
      const l = __label(i).text + ' ' + (i.getAttribute('placeholder') || '') + ' ' + (i.getAttribute('aria-label') || '') + ' ' + (i.name || '') + ' ' + i.id;
      return t === 'search' || __SEARCH_RE.test(l); }).slice(0, 3).map((i) => { let c = i.parentElement, btn = null;
      for (let k = 0; k < 4 && c && !btn; k++, c = c.parentElement) btn = [...c.querySelectorAll('button, input[type=submit], [role=button]')].find((b) => __vis(b) && /tìm|search|lọc|filter|🔍/iu.test(__txt(b) + ' ' + (b.getAttribute('aria-label') || '')));
      return {sel: __sel(i), label: __label(i).text || i.getAttribute('placeholder') || '', button: btn ? {sel: __sel(btn), text: __txt(btn)} : null}; });
  const broken = [...document.images].filter((i) => i.complete && i.naturalWidth === 0 && (i.currentSrc || i.src)).map((i) => i.currentSrc || i.src).slice(0, 20);
  const bodyText = document.body ? document.body.innerText : '';
  const badRe = /(^|[\s:(\[>])(undefined|NaN|\[object Object\]|null)(?=$|[\s,.;)\]<])|\{\{[^}]{1,40}\}\}|lorem ipsum/giu;
  const bad = [...new Set((bodyText.match(badRe) || []).map((s) => s.trim()))].slice(0, 10);
  const errRe = /(Traceback \(most recent call last\)|Exception in thread|NullPointerException|SQLSTATE\[|You have an error in your SQL syntax|ORA-\d{5}|Stack trace:|Fatal error:|Warning: .{0,80} on line \d+|at [\w$.]+\([\w$]+\.(java|cs):\d+\)|Microsoft OLE DB|Unhandled exception|Server Error in '\/')/i;
  const em = bodyText.match(errRe);
  const W = document.documentElement.clientWidth;
  const sw = Math.max(document.documentElement.scrollWidth, document.body ? document.body.scrollWidth : 0);
  const pw = [...document.querySelectorAll('input[type=password]')].map((p) => ({sel: __sel(p), autocomplete: p.getAttribute('autocomplete') || '',
    prefilled: !!p.value, method: ((p.form && p.form.getAttribute('method')) || '').toLowerCase(), action: p.form ? p.form.action : '', vis: __vis(p)}));
  const files = [...document.querySelectorAll('input[type=file]')].map((f) => ({sel: __sel(f), accept: f.getAttribute('accept') || '', label: __label(f).text}));
  const stv = new Set();
  document.querySelectorAll('[class*="badge"], [class*="status"], [class*="state"], [class*="chip"], [class*="tag"]').forEach((e) => {
    if (e.tagName !== 'LABEL' && __vis(e)) { const t = (e.innerText || '').trim(); if (t && t.length <= 30 && !/[\r\n]/.test(t)) stv.add(t); } });
  tables.forEach((t) => t.headers.forEach((h) => { if (/trạng thái|tình trạng|status|state/iu.test(h.text)) {
    document.querySelectorAll(t.rowSel).forEach((r) => { const c = r.children[h.idx]; const v = c && (c.innerText || '').trim(); if (v && v.length <= 30) stv.add(v); }); } }));
  return {title: document.title, url: location.href, lang: document.documentElement.getAttribute('lang') || '',
    h1: [...document.querySelectorAll('h1')].filter(__vis).map(__txt).slice(0, 3), links, buttons, toggles, switches, forms: __forms(),
    tables, pager, search, broken, bad, errText: em ? em[0].slice(0, 200) : '', overflow: {W, sw}, passwords: pw, files,
    textLen: bodyText.length, hasPassword: pw.some((p) => p.vis), statusValues: [...stv].slice(0, 30)};
""")

SNAP_JS = _js(r"""
  const dialogs = [...document.querySelectorAll(__DIALOG_SEL)].filter(__vis).map(__sel);
  const forms = __forms().filter((f) => f.visible).map((f) => f.sel);
  const visCount = [...document.querySelectorAll('a, button, input, select, textarea, li, td, th, label, [role]')].filter(__vis).length;
  const exp = [...document.querySelectorAll('[aria-expanded]')].map((e) => e.getAttribute('aria-expanded')).join(',');
  const chk = [...document.querySelectorAll('input[type=checkbox]')].map((e) => e.checked ? 1 : 0).join('');
  return {url: location.href, dialogs, forms, visCount, exp, chk, textLen: (document.body.innerText || '').length};
""")

PERF_JS = r"""() => { const n = performance.getEntriesByType('navigation')[0]; const res = performance.getEntriesByType('resource');
  const size = (e) => e.transferSize || e.encodedBodySize || 0;
  const total = (n ? size(n) : 0) + res.reduce((a, e) => a + size(e), 0);
  return {load: n ? (n.loadEventEnd || n.domContentLoadedEventEnd || n.duration) : null, dcl: n ? n.domContentLoadedEventEnd : null,
    ttfb: n ? n.responseStart : null, bytes: total, requests: res.length + 1,
    slow: res.filter((e) => e.duration > 1500).sort((a, b) => b.duration - a.duration).slice(0, 8).map((e) => ({name: e.name.slice(0, 120), ms: Math.round(e.duration)}))}; }"""

OVERFLOW_JS = r"""() => { const W = document.documentElement.clientWidth;
  const sw = Math.max(document.documentElement.scrollWidth, document.body ? document.body.scrollWidth : 0); const bad = [];
  if (sw > W + 1) for (const el of document.querySelectorAll('body *')) { const r = el.getBoundingClientRect(); const cs = getComputedStyle(el);
    if (r.width && r.right > W + 1 && cs.position !== 'fixed' && cs.visibility !== 'hidden') { let p = el.parentElement, clipped = false;
      while (p && p !== document.body) { const o = getComputedStyle(p).overflowX; if (o === 'hidden' || o === 'auto' || o === 'scroll' || o === 'clip') { clipped = true; break; } p = p.parentElement; }
      if (!clipped) bad.push(el.tagName.toLowerCase() + (el.id ? '#' + el.id : '') + (typeof el.className === 'string' && el.className.trim() ? '.' + el.className.trim().split(/\s+/)[0] : '') + ' (+' + Math.round(r.right - W) + 'px)'); }
    if (bad.length > 6) break; }
  return {W, sw, bad}; }"""

A11Y_JS = _js(r"""
  const minRatio = __a;
  const imgsNoAlt = [...document.images].filter((i) => __vis(i) && !i.hasAttribute('alt')).map((i) => (i.currentSrc || i.src || '').split('/').pop().slice(0, 60)).slice(0, 10);
  const inputs = [...document.querySelectorAll('input, select, textarea')].filter((x) => __fieldOk(x) && __vis(x));
  const noLabel = inputs.filter((x) => !['aria', 'label'].includes(__label(x).src)).map((x) => (x.getAttribute('placeholder') || x.name || x.id || x.tagName.toLowerCase())).slice(0, 10);
  const noName = [...document.querySelectorAll('button, a[href], [role=button]')].filter(__vis).filter((b) => !__txt(b) && !b.querySelector('img[alt]:not([alt=""]), [aria-label]')).map((b) => b.outerHTML.slice(0, 80)).slice(0, 10);
  const heads = [...document.querySelectorAll('h1,h2,h3,h4,h5,h6')].filter(__vis).map((h) => +h.tagName[1]);
  const jumps = []; for (let i = 1; i < heads.length; i++) if (heads[i] - heads[i - 1] > 1) jumps.push('h' + heads[i - 1] + '→h' + heads[i]);
  const lum = (c) => { const m = (c || '').match(/rgba?\(([^)]+)\)/); if (!m) return null; const p = m[1].split(',').map((s) => parseFloat(s));
    if (p.length > 3 && p[3] < 0.5) return 'transparent'; const f = (v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); };
    return 0.2126 * f(p[0]) + 0.7152 * f(p[1]) + 0.0722 * f(p[2]); };
  const bgOf = (el) => { for (let e = el; e; e = e.parentElement) { const cs = getComputedStyle(e); if (cs.backgroundImage && cs.backgroundImage !== 'none') return null;
    const l = lum(cs.backgroundColor); if (l !== null && l !== 'transparent') return l; } return 1; };
  const low = [];
  if (minRatio > 0) { const els = [...document.querySelectorAll('body *')].filter((e) => [...e.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim().length > 1) && __vis(e)).slice(0, 500);
    for (const e of els) { const cs = getComputedStyle(e); const fg = lum(cs.color); const bg = bgOf(e);
      if (fg === null || fg === 'transparent' || bg === null) continue; const r = (Math.max(fg, bg) + 0.05) / (Math.min(fg, bg) + 0.05);
      const size = parseFloat(cs.fontSize); const large = size >= 24 || (size >= 18.66 && (+cs.fontWeight >= 700)); const need = large ? Math.min(3, minRatio) : minRatio;
      if (r < need) low.push([...e.childNodes].filter((n) => n.nodeType === 3).map((n) => n.textContent.trim()).join(' ').slice(0, 40) + ' (' + (Math.round(r * 100) / 100) + ':1)');
      if (low.length >= 8) break; } }
  return {imgsNoAlt, noLabel, noName, lang: document.documentElement.getAttribute('lang') || '', h1: heads.filter((h) => h === 1).length, jumps: jumps.slice(0, 5), low, title: document.title};
""", arg=True)

FORM_PREP_JS = r"""(fsel) => { window.__at_sub = false; const c = document.querySelector(fsel);
  const form = c && (c.tagName === 'FORM' ? c : c.querySelector('form'));
  (form || document).addEventListener('submit', () => { window.__at_sub = true; }, {capture: true});
  window.__at_before = new Set((document.body.innerText || '').split('\n').map((s) => s.trim()).filter(Boolean)); }"""

FORM_STATE_JS = _js(r"""
  const ERR = /vui lòng|bắt buộc|không hợp lệ|không đúng|không được|phải|tối đa|tối thiểu|ít nhất|vượt quá|chưa nhập|chưa chọn|định dạng|lỗi|sai|không thể|required|invalid|must|error|incorrect|too long|too short|at least|not allowed|please/iu;
  const OK = /thành công|đã lưu|đã tạo|đã thêm|đã gửi|đã cập nhật|đã nhập|hoàn tất|success|saved|created|submitted|updated/iu;
  const before = window.__at_before || new Set();
  const lines = (document.body.innerText || '').split('\n').map((s) => s.trim()).filter(Boolean);
  const fresh = lines.filter((l) => !before.has(l) && l.length < 200);
  const errEls = [...document.querySelectorAll('[role=alert], [aria-live], .error, .err, .invalid-feedback, .help-block, .text-danger, .field-error, .error-message, [class*="error"], [class*="invalid"], [class*="danger"]')]
    .filter(__vis).map((e) => (e.innerText || '').trim()).filter((t) => t && t.length < 200 && !before.has(t));
  const c = document.querySelector(__a.fsel);
  const fields = (__a.fields || []).map((s) => { try { return document.querySelector(s); } catch (e) { return null; } }).filter(Boolean);
  const invalid = fields.filter((f) => (f.checkValidity && !f.checkValidity()) || f.getAttribute('aria-invalid') === 'true' ||
    /(^|\s)(is-invalid|invalid|has-error|error)(\s|$)/.test(f.className + ' ' + (f.parentElement ? f.parentElement.className : '')))
    .map((f) => ({label: __label(f).text, msg: f.validationMessage || ''}));
  const errTexts = [...new Set([...fresh.filter((l) => ERR.test(l) && !OK.test(l)), ...errEls.filter((t) => !OK.test(t))])].slice(0, 8);
  const okTexts = [...new Set(fresh.filter((l) => OK.test(l)))].slice(0, 5);
  return {invalid, errTexts, okTexts, formVisible: !!(c && __vis(c)) && fields.some(__vis), url: location.href, submitted: !!window.__at_sub, fresh: fresh.slice(0, 8)};
""", arg=True)

BTN_IN_JS = _js(r"""
  const out = []; for (const s of __a) { let c = null; try { c = document.querySelector(s); } catch (e) {} if (!c) continue;
    c.querySelectorAll('button, input[type=submit], input[type=button], [role=button], a.btn').forEach((b) => { if (__vis(b) && __txt(b)) out.push(__txt(b)); }); }
  return [...new Set(out)].slice(0, 20);
""", arg=True)

ROWS_JS = _js(r"""
  return [...document.querySelectorAll(__a.rowSel)].filter(__vis).map((r) => (__a.idx === null || __a.idx === undefined) ? (r.innerText || '').replace(/\s+/g, ' ').trim()
    : ((r.children[__a.idx] || {}).innerText || '').replace(/\s+/g, ' ').trim());
""", arg=True)

DIALOG_CLOSE_JS = _js(r"""
  const box = __a ? document.querySelector(__a) : null; const scope = box || document;
  const cands = [...scope.querySelectorAll('button, [role=button], a, .close, [class*="close"], [aria-label]')].filter(__vis);
  const c = cands.find((b) => /^(đóng|hủy|huỷ|close|cancel|×|✕|x|thoát|quay lại)$/iu.test(__txt(b)) || /close|đóng|dismiss/iu.test((b.getAttribute('aria-label') || '') + ' ' + (b.className + '')));
  return c ? {sel: __sel(c), text: __txt(c) || c.getAttribute('aria-label') || 'X'} : null;
""", arg=True)

ERR_PY = re.compile(r"vui lòng|bắt buộc|không hợp lệ|không đúng|không được|phải|tối đa|tối thiểu|ít nhất|vượt quá|chưa "
                    r"nhập|chưa chọn|định dạng|lỗi|sai|không thể|required|invalid|must|error|incorrect|please", re.I)
OK_PY = re.compile(r"thành công|đã lưu|đã tạo|đã thêm|đã gửi|đã cập nhật|đã nhập|hoàn tất|success|saved|created|"
                   r"submitted|updated", re.I)
DENY_RE = re.compile(r"không có quyền|không được phép|truy cập bị từ chối|access denied|forbidden|unauthori[sz]ed|"
                     r"permission denied|403|401", re.I)
WRITE_BTN_RE = re.compile(r"\b(them|tao|sua|xoa|luu|nhap du lieu|import|phe duyet|duyet|cap nhat|add|create|new|edit|"
                          r"delete|remove|save|approve|update)\b")
SKIP_EXT = (".pdf", ".zip", ".rar", ".7z", ".xls", ".xlsx", ".doc", ".docx", ".ppt", ".pptx", ".csv", ".jpg", ".jpeg",
            ".png", ".gif", ".svg", ".webp", ".mp4", ".mp3", ".exe", ".msi", ".apk", ".dmg", ".iso", ".json", ".xml")
TODAY = dt.date.today


def win_name():
    if os.name != "nt":
        return sys.platform
    try:
        return "Windows 11" if sys.getwindowsversion().build >= 22000 else "Windows 10"
    except Exception:
        return "Windows"


BROWSER_NAME = {"edge": "Edge", "chrome": "Chrome", "firefox": "Firefox"}

# Các LUỒNG của file Test Scenario (đúng thứ tự file mẫu của tester). Scenario nào thuộc luồng nào: LUONG_OF.
LUONG = {1: "TRUY CẬP, GIAO DIỆN & PHÂN QUYỀN", 2: "HIỂN THỊ DỮ LIỆU DANH SÁCH", 3: "TÌM KIẾM & BỘ LỌC NÂNG CAO",
         4: "PHỐI HỢP CHỨC NĂNG DANH SÁCH", 5: "THÊM MỚI: NHẬP LIỆU & RÀNG BUỘC", 6: "LƯU & VÒNG ĐỜI PHÊ DUYỆT",
         7: "XEM CHI TIẾT", 8: "HÀNH ĐỘNG TRÊN BẢN GHI", 9: "NHẬP DỮ LIỆU TỪ FILE"}
_LF = {1: "ui responsive labels page login links ease perf a11y perm sec compat",
       2: "mapping api date_col status_col loading",
       3: "search search_input filter toggle dropdown reset no_result",
       4: "combo sorting sort_filter paging page_filter keep_state",
       5: "form_ui required unique format length date_rule rule cancel",
       6: "valid btn_save form_actions flow approve reject lifecycle",
       7: "detail detail_data",
       8: "buttons row_action edit delete copy switch popup export",
       9: "template import_ok file import_bad"}
LUONG_OF = {f: n for n, fs in _LF.items() for f in fs.split()}
FEAT_RANK = {f: i for fs in _LF.values() for i, f in enumerate(fs.split())}


def luong_of_text(text):
    """Scenario lấy từ file (không biết tính năng) → đoán luồng theo chữ."""
    t = C.norm(text)
    for n, rx in ((9, r"nhap du lieu|import|tai file mau|tai tep|file excel|json"),
                  (7, r"chi tiet"), (6, r"luu nhap|phe duyet|tu choi|gui duyet|vong doi|\bluu\b"),
                  (5, r"them moi|tao moi|form|bat buoc|rang buoc|dinh dang|do dai"),
                  (4, r"phan trang|sap xep|ket hop"), (3, r"tim kiem|bo loc|\bloc\b|lam moi"),
                  (2, r"hien thi du lieu|mapping|danh sach|cot "), (8, r"thao tac|hanh dong|sao chep|xoa|sua")):
        if re.search(rx, t):
            return n
    return 1


DIAG_JS = r"""() => { const b = document.body; const txt = b ? (b.innerText || '') : '';
  const q = (s) => { try { return [...document.querySelectorAll(s)].some((e) => { const r = e.getBoundingClientRect(); const cs = getComputedStyle(e);
    return r.width > 30 && r.height > 30 && cs.display !== 'none' && cs.visibility !== 'hidden' && parseFloat(cs.opacity || '1') > 0.1; }); } catch (e) { return false; } };
  const cap = q('[id*=captcha i], [class*=captcha i], [class*=secsdk], [id*=recaptcha i], [class*=geetest i], [class*=turnstile i], iframe[src*=captcha], iframe[src*=recaptcha], iframe[src*=hcaptcha], iframe[src*=turnstile], iframe[src*=challenge]');
  const all = b ? b.querySelectorAll('*') : [];
  return {text: txt.slice(0, 3000), textLen: txt.length, iframes: [...document.querySelectorAll('iframe')].map((f) => f.src || '').slice(0, 10),
    video: document.querySelectorAll('video').length, canvas: document.querySelectorAll('canvas').length, cap, nEls: all.length,
    shadow: [...all].filter((e) => e.shadowRoot).length, title: document.title, url: location.href}; }"""
CAPTCHA_RE = re.compile(r"captcha|verify you are human|i'?m not a robot|xác minh bạn|xác thực bạn không phải|unusual traffic|"
                        r"安全验证|拖动滑块|滑块验证|请完成验证|请完成安全验证|人机验证|just a moment|checking your browser|attention required|"
                        r"access denied|request blocked|bot detected|are you a robot", re.I)
LOGIN_WALL_RE = re.compile(r"đăng nhập|log ?in|sign ?in|登录|扫码|login to continue|vui lòng đăng nhập", re.I)


def words_re(text):
    ws = [C.norm(w) for w in re.split(r"[,;\n]+", text or "") if C.norm(w)]
    if not ws:
        return None
    return re.compile(r"(^|\b)(" + "|".join(re.escape(w) for w in ws) + r")(\b|$)")


class Scanner(E.Runner):
    def __init__(self, cfg, options, run_dir, data_dir):   # noqa: C901 (không gọi Runner.__init__)
        import threading
        self.cfg = cfg
        self.url = cfg["url"].strip()
        if not re.match(r"^https?://", self.url, re.I):
            self.url = "http://" + self.url
        p = urlparse(self.url)
        self.origin = (p.scheme.lower(), p.netloc.lower())
        self.discover_only = bool(cfg.get("discover_only"))
        self.wb = {"project": cfg.get("project") or p.netloc or "web", "file": None, "scenarios": [], "testcases": [],
                   "headers": {}, "elements": {}, "envs": {}, "env_names": []}
        self.tc_by_key, self.tc_by_id, self.selection = {}, {}, []
        self.opt = options
        self.run_dir = run_dir
        self.data_dir = data_dir
        self.img_dir = os.path.join(run_dir, "anh")
        os.makedirs(self.img_dir, exist_ok=True)
        self.base_dir = os.path.join(data_dir, "anh-chuan", E.slug(p.netloc or "web"))
        self.timeout = max(3.0, float(options.get("timeout") or 10))
        self.lock = threading.Lock()
        self.events, self.results = [], []
        self.state = "chuẩn bị"
        self.stop_flag = threading.Event()
        self.manual_req = self.manual_ans = None
        self.manual_evt = threading.Event()
        self.current = None
        self.started = dt.datetime.now()
        self.finished = None
        self.thread = None
        self.pw = None
        self.browsers = {}
        self.warn_browser = set()
        self.instances = []
        self.total = 0
        self.types = cfg.get("types") or T.default_type_opts()
        self.rules = cfg.get("rules") or []
        self.pages = []
        self.login_rec = None
        self.logged_in = False
        self.storage = None
        self.counters = {}
        self.scn_by_key = {}
        self.codes, self.scn_count, self.used_tc = {}, {}, set()
        self.imported = [x for x in (cfg.get("scenarios_in") or []) if x.get("id") or x.get("scenario")]
        self.only = set(cfg.get("only") or []) or None
        self.force_all = bool(self.imported)
        self.data_n = 0
        self.vars = {}
        self.dialogs, self.console_errors, self.resp_log = [], [], []
        self.dismiss_next = False
        self.req_done = set()      # request đã tải xong (chỉ đọc JSON của request này — tránh treo ở luồng stream)
        self.inst = {"n": 0, "label": "quet", "tc": {"id": "quet", "mode": "Tự động", "fields": {}}, "browser": "edge"}
        self.discovery = []
        self.skip_link = words_re(cfg.get("skip_words") or T.SKIP_LINK_WORDS)
        self.unsafe = words_re(T.UNSAFE_WORDS + ", " + (cfg.get("unsafe_extra") or ""))
        self.env_label = "%s / %s" % (BROWSER_NAME.get(options.get("browser", "edge"), "Edge"), win_name())
        self.rnd = str(random.randint(10000, 99999))

    # ------------------------------------------------------------------ trạng thái cho giao diện
    def on(self, key):
        if getattr(self, "force_all", False) and key not in ("compat",):
            return True     # lấy Test Scenario từ file: sinh đủ các nhóm rồi ghép theo scenario của tester
        return bool((self.types.get(key) or {}).get("on"))

    def o(self, key, opt):
        t = self.types.get(key) or {}
        v = (t.get("opts") or {}).get(opt)
        if v is None:
            for x in T.BY_KEY[key]["options"]:
                if x["key"] == opt:
                    return x["default"]
        return v

    def snapshot(self):
        s = super().snapshot()
        s["total"] = max(self.total, len(self.results))
        s["mode"] = "scan"
        return s

    def result_data(self):
        d = super().result_data()
        d.update({"kind": "scan", "url": self.url, "env": self.env_label, "discovery": self.discovery,
                  "diagnosis": getattr(self, "diagnosis", []),
                  "pages": [{"url": p["url"], "title": p.get("title", ""), "status": p.get("status")} for p in self.pages],
                  "cfg": self._safe_cfg()})
        return d

    def _safe_cfg(self):
        """Cấu hình lưu kèm kết quả — bỏ mật khẩu."""
        c = json.loads(json.dumps({k: v for k, v in self.cfg.items() if k != "login"}, ensure_ascii=False))
        perm = ((c.get("types") or {}).get("permission") or {}).get("opts") or {}
        if "pass" in perm:
            perm["pass"] = ""
        return c

    def _info(self, msg):
        self.emit("info", msg=msg)
        with self.lock:
            if self.current is not None:
                self.current["step"] = ""
            self.current = {"label": "Đang dò", "title": msg, "mode": "Tự động", "step": "", "n": 0}

    # ------------------------------------------------------------------ sự kiện trình duyệt
    def _hook_page(self, page):
        if id(page) in self._hooked:
            return
        self._hooked.add(id(page))
        page.on("dialog", self._on_dialog)
        page.on("console", self._on_console)
        page.on("pageerror", lambda e: self.console_errors.append(str(e)[:300]))
        page.on("response", self._on_response)
        page.on("requestfinished", lambda q: self.req_done.add(q))
        page.on("requestfailed", lambda r: self.resp_log.append({"url": r.url, "status": 0, "type": r.resource_type,
                                                                 "failed": (r.failure or "")[:100], "resp": None}))

    def _on_console(self, m):
        # "Failed to load resource" = tài nguyên lỗi (đã kiểm tra riêng ở bước tài nguyên) → không tính là lỗi JS
        if m.type == "error" and not m.text.startswith("Failed to load resource") and "favicon" not in m.text:
            self.console_errors.append(m.text[:300])

    def _on_response(self, r):
        try:
            rt = r.request.resource_type
            ent = {"url": r.url, "status": r.status, "type": rt, "resp": None}
            if rt in ("xhr", "fetch") and "json" in (r.headers.get("content-type") or ""):
                ent["resp"] = r
                ent["req"] = r.request
            self.resp_log.append(ent)
        except Exception:
            pass

    def _on_dialog(self, d):
        """Khi quét: hộp thoại xác nhận (confirm) luôn bấm HỦY để không xóa/duyệt nhầm dữ liệu thật."""
        self.dialogs.append(d.message)
        try:
            if d.type in ("confirm", "prompt"):
                d.dismiss()
            else:
                d.accept()
        except Exception:
            pass

    def _new_ctx(self, browser=None, storage=None, size=None):
        b = self._launch(browser or self.opt.get("browser", "edge"))
        w, h = size or X.size_of(self.opt.get("size", "")) or (1366, 768)
        kw = {"viewport": {"width": w, "height": h}, "ignore_https_errors": True, "locale": "vi-VN"}
        if storage:
            kw["storage_state"] = storage
        ctx = b.new_context(**kw)
        ctx.set_default_timeout(self.timeout * 1000)
        ctx.set_default_navigation_timeout(max(30, self.timeout * 3) * 1000)
        if not hasattr(self, "_hooked"):
            self._hooked = set()
        ctx.on("page", self._hook_page)
        return ctx

    def _ensure_page(self):
        try:
            if self.page and not self.page.is_closed():
                return self.page
        except Exception:
            pass
        self.page = self.ctx.new_page()
        self._hook_page(self.page)
        return self.page

    def _settle(self, ms=300, idle=2500):
        try:
            self.page.wait_for_load_state("networkidle", timeout=idle)
        except Exception:
            pass
        self._pump(ms)

    def _goto(self, url):
        page = self._ensure_page()
        for extra in self.ctx.pages:
            if extra is not page:
                try:
                    extra.close()
                except Exception:
                    pass
        resp = page.goto(url, wait_until="load")
        self._settle()
        return resp

    def _is_login_page(self):
        try:
            return self.page.locator("input[type=password]").filter(visible=True).count() > 0
        except Exception:
            return False

    # ------------------------------------------------------------------ kết quả / scenario
    def _code(self, module):
        """'Đơn hàng tuyển dụng' -> 'DHTD' (mã màn hình dùng cho TS_DHTD_001, TC_DHTD_001)."""
        module = module or "Hệ thống"
        if module in self.codes:
            return self.codes[module]
        words = [w for w in C.norm(module).upper().split() if w.isalnum()]
        c = ("".join(w[0] for w in words)[:6] if len(words) > 1 else (words[0][:4] if words else "WEB")) or "WEB"
        base, i = c, 2
        while c in self.codes.values():
            c = "%s%d" % (base, i)
            i += 1
        self.codes[module] = c
        return c

    def _S(self, tkey, feature, module, title, ac="", prio=None, luong=None):
        """Một TEST SCENARIO: mã theo màn hình (TS_DHTD_001), cột Nhóm như file tester (functional, validation…),
        thuộc 1 LUỒNG (truy cập & giao diện, hiển thị danh sách, tìm kiếm & lọc… — xem LUONG)."""
        key = "%s|%s|%s" % (tkey, feature, title)
        if key in self.scn_by_key:
            return self.scn_by_key[key]
        code = self._code(module)
        self.scn_count[code] = self.scn_count.get(code, 0) + 1
        sid = "TS_%s_%03d" % (code, self.scn_count[code])
        sc = {"id": sid, "key": key, "type": tkey, "feature": feature, "code": code, "module": module or "",
              "luong": luong or LUONG_OF.get(feature) or luong_of_text(title), "n": len(self.wb["scenarios"]),
              "fields": {"ts_id": sid, "nhom": T.NHOM.get(tkey, tkey), "module": module or "", "scenario": title,
                         "priority": prio or T.BY_KEY[tkey]["priority"], "ac": ac},
              "extra": {}, "row": None, "tcs": []}
        self.wb["scenarios"].append(sc)
        self.scn_by_key[key] = sc
        return sc

    def _scn(self, tkey, title, module=""):
        k = (tkey, title)
        if k in self.scn_by_key:
            return self.scn_by_key[k]
        t = T.BY_KEY[tkey]
        n = sum(1 for s in self.wb["scenarios"] if s["fields"].get("nhom") == t["name"]) + 1
        sid = "TS_%s_%02d" % (t["prefix"], n)
        s = {"id": sid, "fields": {"ts_id": sid, "nhom": t["name"], "module": module, "scenario": title,
                                   "priority": t["priority"], "ac": ""}, "extra": {}, "row": None, "tcs": []}
        self.wb["scenarios"].append(s)
        self.scn_by_key[k] = s
        return s

    def _res(self, tkey, title, scn, page=None, priority=None, mode="Tự động", pre="", data="", size="",
             browser=None, module=None, tc_id=None):
        t = T.BY_KEY[tkey]
        if tc_id and tc_id not in self.used_tc:      # mã TC đã đánh sẵn khi lập kế hoạch → giống bảng xem trước
            self.used_tc.add(tc_id)
            return self._res_dict(tkey, title, scn, tc_id, page, priority, mode, pre, data, size, browser, module)
        code = self._scn_code(scn)
        self.counters[code] = self.counters.get(code, 0) + 1
        tc_id = ("TC_%s_%03d" % (code, self.counters[code])) if code else "TC%03d" % self.counters[code]
        while tc_id in self.used_tc:
            self.counters[code] += 1
            tc_id = ("TC_%s_%03d" % (code, self.counters[code])) if code else "TC%03d" % self.counters[code]
        self.used_tc.add(tc_id)
        return self._res_dict(tkey, title, scn, tc_id, page, priority, mode, pre, data, size, browser, module)

    @staticmethod
    def _scn_code(scn):
        return scn.get("code") or ("_".join(scn["id"].split("_")[1:-1]) if scn["id"].count("_") >= 2 else "")

    def _res_dict(self, tkey, title, scn, tc_id, page, priority, mode, pre, data, size, browser, module):
        t = T.BY_KEY[tkey]
        b = browser or self.opt.get("browser", "edge")
        pg = page or {}
        mod = module if module is not None else (pg.get("name") or "")
        data_id = ""
        if data:
            self.data_n += 1
            data_id = "TD-%03d" % self.data_n
        fields = {"tc_id": tc_id, "scenario_id": scn["id"], "requirement": "%s / %s" % (scn["id"], scn["fields"]["scenario"]),
                  "module": mod, "man_hinh": pg.get("path", ""), "tieu_de": title, "loai_test": scn["fields"].get("nhom") or t["name"],
                  "priority": priority or t["priority"], "precondition": pre or self._pre(), "test_data": data,
                  "data_id": data_id, "kich_thuoc": size, "trinh_duyet": BROWSER_NAME.get(b, b), "cach_test": mode,
                  "moi_truong": "%s / %s" % (BROWSER_NAME.get(b, b), win_name()), "mo_ta": pg.get("url", "")}
        res = {"key": tc_id, "tc_id": tc_id, "label": tc_id, "title": title, "scenario_id": scn["id"], "mode": mode,
               "mode_reason": "", "fields": fields, "extra": {}, "size": size, "browser": b, "env": self.env_label,
               "steps": [], "status": "", "note": "", "start": dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
               "duration": 0, "video": "", "tester": self.opt.get("tester", ""), "row": None, "type": tkey,
               "page_url": pg.get("url", "")}
        scn["tcs"].append(tc_id)
        return res

    def _pre(self):
        return "Đã đăng nhập (tài khoản chính)" if self.logged_in else "Không cần đăng nhập"

    @staticmethod
    def _st(res, buoc, expected="", data=""):
        s = {"no": str(len(res["steps"]) + 1), "fields": {"buoc": buoc, "expected": expected, "du_lieu_buoc": data},
             "extra": {}, "action": "", "kind": "auto", "status": "", "actual": "", "images": [], "duration": 0,
             "row": None}
        res["steps"].append(s)
        return s

    def _ok(self, s, actual):
        s["status"], s["actual"] = "Pass", actual

    def _bad(self, s, actual, shot=True, status="Fail"):
        s["status"], s["actual"] = status, actual
        if shot:
            self._shot(s, "loi")

    def _done(self, res, t0):
        res["duration"] = round(time.time() - t0, 1)
        for s in res["steps"]:
            if not s["status"]:
                s["status"], s["actual"] = "Skip", s["actual"] or "Không chạy"
        res["status"] = E.tc_status([s["status"] for s in res["steps"]]) if res["steps"] else "Chưa test"
        with self.lock:
            self.results.append(res)
        self.emit("tc_done", label="%s %s" % (res["tc_id"], res["title"][:60]), status=res["status"],
                  key=res["key"], duration=res["duration"])

    def _run_item(self, item):
        """item = (tkey, title, scn, kwargs của _res, hàm chạy(res))."""
        tkey, title, scn, kw, fn = item
        res = self._res(tkey, title, scn, **kw)
        self.inst = {"n": len(self.results), "label": res["tc_id"], "browser": res["browser"],
                     "tc": {"id": res["tc_id"], "mode": res["mode"], "fields": res["fields"]}}
        with self.lock:
            self.current = {"label": res["tc_id"], "title": title, "mode": res["mode"], "step": "",
                            "n": len(self.results) + 1}
        t0 = time.time()
        try:
            fn(res)
        except E.Stopped:
            for s in res["steps"]:
                if not s["status"]:
                    s["status"], s["actual"] = "Chưa test", "Đã bấm Dừng"
            if not res["steps"]:
                self._st(res, "Chạy kiểm tra")["status"] = "Chưa test"
            res["note"] = "Đã bấm Dừng"
            self._done(res, t0)
            raise
        except Exception as e:
            msg = self._friendly(e, {})
            s = next((x for x in res["steps"] if not x["status"]), None) or self._st(res, "Chạy kiểm tra")
            self._bad(s, "Không chạy được: %s" % msg, status="Blocked")
            res["note"] = msg
            self.emit("warn", msg="%s: %s" % (res["tc_id"], msg))
        self._done(res, t0)
        self.save()

    # ------------------------------------------------------------------ chạy tổng
    def _run_all(self):
        self.state = "đang chạy"
        self.emit("start", total=0)
        lm = (self.cfg.get("login") or {}).get("mode", "none")
        pm = self.o("permission", "login_mode") if self.on("permission") else "none"
        if self.opt.get("headless") and ("manual" in (lm, pm)):
            self.opt["headless"] = False
            self.emit("warn", msg="Bạn chọn tự đăng nhập bằng tay nên tool HIỆN trình duyệt (không chạy ẩn).")
        try:
            from playwright.sync_api import sync_playwright
            self.pw = sync_playwright().start()
        except Exception as e:
            self.state = "lỗi"
            self.emit("error", msg="Không khởi động được bộ điều khiển trình duyệt: %s" % e)
            self._finish()
            return
        try:
            self.ctx = self._new_ctx()
            self.page = None
            self._ensure_page()
            self._phase_login()
            reuse = self.cfg.get("reuse_pages")
            if reuse and os.path.isfile(reuse):
                with open(reuse, encoding="utf-8") as f:
                    self.pages = json.load(f)
                self.emit("info", msg="Dùng kết quả dò trang đã có (%d trang)" % len(self.pages))
                self.discovery = [self._disc_page(p) for p in self.pages]
            else:
                self._phase_discover()
                try:
                    with open(os.path.join(self.run_dir, "pages.json"), "w", encoding="utf-8") as f:
                        json.dump([{k: v for k, v in pg.items() if k != "json"} | {"json": pg.get("json", [])[:3]}
                                   for pg in self.pages], f, ensure_ascii=False, default=str)
                except Exception as e:
                    self.emit("warn", msg="Không lưu được kết quả dò: %s" % e)
            if self.discover_only:
                return
            plan = self._make_plan()
            self.total = len(self.results) + len(plan)
            self.emit("info", msg="Đã sinh %d test case — bắt đầu chạy" % len(plan))
            for i, item in enumerate(plan):
                if self.stop_flag.is_set():
                    tkey, title, scn, kw, _ = item
                    res = self._res(tkey, title, scn, **kw)
                    self._st(res, "Chạy kiểm tra")["status"] = "Chưa test"
                    res["note"] = "Đã bấm Dừng"
                    self._done(res, time.time())
                    continue
                try:
                    self._run_item(item)
                except E.Stopped:
                    continue
        except E.Stopped:
            pass
        except E.StepBlocked as e:
            self.emit("error", msg=str(e))
        except Exception as e:
            self.emit("error", msg="Lỗi không mong muốn: %s" % e, detail=traceback.format_exc())
        finally:
            for b in self.browsers.values():
                try:
                    b.close()
                except Exception:
                    pass
            try:
                self.pw.stop()
            except Exception:
                pass
            self._finish()

    # ------------------------------------------------------------------ 1. ĐĂNG NHẬP
    def _state_file(self, suffix=""):
        d = os.path.join(self.data_dir, "phien")
        os.makedirs(d, exist_ok=True)
        return os.path.join(d, E.slug("%s_%s%s" % (self.origin[0], self.origin[1], suffix)) + ".json")

    def _wait_human(self, req, banner):
        req.setdefault("req_id", "h-%d" % int(time.time() * 1000))
        req.setdefault("run_id", os.path.basename(self.run_dir))
        req.setdefault("steps", [])
        self._banner(True, banner)
        with self.lock:
            self.manual_ans = None
            self.manual_req = req
            self.state = "chờ bạn kiểm tra"
        self.manual_evt.clear()
        self.emit("manual", label=req.get("label", ""), reason=req.get("reason", ""))
        while not self.manual_evt.wait(0.2):
            self._pump(300)
        with self.lock:
            ans = self.manual_ans
            self.manual_req = None
            if self.state == "chờ bạn kiểm tra":
                self.state = "đang chạy"
        self._banner(False)
        if ans is None:
            raise E.Stopped()
        return ans

    def _login_into(self, ctx_page_setter, mode, login_url, user, pw, who="tài khoản chính"):
        """Đăng nhập trên self.page. Trả về (ok, mô tả)."""
        page = self._ensure_page()
        self._goto(login_url)
        if mode == "manual":
            for attempt in range(2):
                self._wait_human({"kind": "login", "label": "Đăng nhập", "title": "Đăng nhập %s" % who,
                                  "reason": "Hãy đăng nhập %s trong cửa sổ trình duyệt mà tool vừa mở, xong bấm "
                                            "'Đã đăng nhập xong'." % who, "opened": login_url},
                                 "🔑 Hãy đăng nhập ở đây, xong bấm 'Đã đăng nhập xong' trong cửa sổ Tool Test")
                self._settle(300, 3000)
                if not self._is_login_page():
                    return True, "Người test tự đăng nhập — vào được %s" % self.page.url
            return False, "Vẫn đang ở trang đăng nhập (có ô mật khẩu)"
        # tự điền
        try:
            self.page.locator("input[type=password]").filter(visible=True).first.wait_for(state="visible",
                                                                                         timeout=self.timeout * 1000)
        except Exception:
            return False, "Không thấy ô mật khẩu ở trang %s" % login_url
        info = self.page.evaluate(_js(r"""
          const p = [...document.querySelectorAll('input[type=password]')].find(__vis);
          const scope = (p && (p.form || p.closest('div, section, main'))) || document;
          const ins = [...scope.querySelectorAll('input')].filter((i) => __vis(i) && ['text', 'email', 'tel', ''].includes((i.getAttribute('type') || 'text').toLowerCase()));
          const before = ins.filter((i) => p && (i.compareDocumentPosition(p) & Node.DOCUMENT_POSITION_FOLLOWING));
          const u = before[before.length - 1] || ins[0];
          const btns = [...scope.querySelectorAll('button, input[type=submit], [role=button]')].filter(__vis);
          const b = btns.find((x) => (x.getAttribute('type') || '').toLowerCase() === 'submit') || btns.find((x) => /đăng nhập|login|log in|sign in|vào/iu.test(__txt(x))) || btns[0];
          return {user: u ? __sel(u) : null, pass: p ? __sel(p) : null, btn: b ? __sel(b) : null};
        """))
        if not info.get("pass"):
            return False, "Không thấy ô mật khẩu"
        if info.get("user") and user:
            self.page.locator(info["user"]).first.fill(user)
        self.page.locator(info["pass"]).first.fill(pw or "")
        url0 = self.page.url
        if info.get("btn"):
            self.page.locator(info["btn"]).first.click()
        else:
            self.page.locator(info["pass"]).first.press("Enter")
        deadline = time.time() + self.timeout
        while time.time() < deadline:
            self._pump(300)
            if not self._is_login_page():
                break
        self._settle(300, 3000)
        if self._is_login_page():
            txt = ""
            try:
                txt = " ".join(t for t in self.page.inner_text("body").split("\n") if ERR_PY.search(t))[:200]
            except Exception:
                pass
            return False, "Đăng nhập không thành công%s" % (": " + txt if txt else " (vẫn ở trang đăng nhập)")
        return True, "Đăng nhập thành công, đang ở %s" % (self.page.url if self.page.url != url0 else url0)

    def _phase_login(self):
        lg = self.cfg.get("login") or {}
        mode = lg.get("mode", "none")
        if mode == "none":
            return
        login_url = (lg.get("url") or "").strip() or self.url
        scn = self._S("functional", "login", "Đăng nhập", "Kiểm tra đăng nhập hệ thống bằng tài khoản hợp lệ")
        sf = self._state_file()
        self._info("Đăng nhập…")
        if lg.get("remember", True) and os.path.exists(sf):
            try:
                ctx = self._new_ctx(storage=sf)
                self.ctx.close()
                self.ctx, self.page = ctx, None
                self._ensure_page()
                self._goto(self.url)
                if not self._is_login_page():
                    self.logged_in = True
                    self.storage = sf

                    def fn(res):
                        s = self._st(res, "Mở %s bằng phiên đăng nhập đã lưu" % self.url, "Vào được hệ thống")
                        self._ok(s, "Phiên đăng nhập còn hiệu lực — vào được %s" % self.page.url)
                    self._run_item(("functional", "Đăng nhập bằng phiên đã lưu", scn, {"module": "Đăng nhập",
                                                                                       "pre": "Có phiên lưu"}, fn))
                    return
            except Exception:
                pass
        # đăng nhập mới
        try:
            self.page.goto(login_url, wait_until="load")
            self._settle()
            if self._is_login_page():
                self.login_rec = {"url": self.page.url, "info": self.page.evaluate(PAGE_INFO_JS)}
        except Exception:
            pass
        holder = {}

        def fn(res):
            s1 = self._st(res, "Mở trang đăng nhập", "Hiện form đăng nhập", login_url)
            s2 = self._st(res, "Nhập tài khoản / mật khẩu và bấm Đăng nhập" if mode == "auto" else
                          "Người test tự đăng nhập trong trình duyệt", "Đăng nhập thành công, vào trang chính",
                          ("Tài khoản: %s" % lg.get("user", "")) if mode == "auto" else "")
            ok, msg = self._login_into(None, mode, login_url, lg.get("user", ""), lg.get("pass", ""))
            self._ok(s1, "Đã mở %s" % login_url)
            if ok:
                self._ok(s2, msg)
                self._shot(s2, "dangnhap")
            else:
                self._bad(s2, msg)
            holder["ok"] = ok
        self._run_item(("functional", "Đăng nhập với tài khoản đã cấu hình", scn,
                        {"module": "Đăng nhập", "pre": "Có tài khoản hợp lệ",
                         "data": ("Tài khoản: %s" % lg.get("user", "")) if mode == "auto" else ""}, fn))
        if not holder.get("ok"):
            raise E.StepBlocked("Đăng nhập không thành công nên không quét tiếp được. Kiểm tra tài khoản / cách "
                                "đăng nhập rồi chạy lại.")
        self.logged_in = True
        try:
            if lg.get("remember", True):
                self.ctx.storage_state(path=sf)
                self.storage = sf
            else:
                self.storage = self.ctx.storage_state()
        except Exception:
            self.storage = None
        if self.page.url.rstrip("/") != self.url.rstrip("/"):
            pass

    # ------------------------------------------------------------------ 2. DÒ TRANG
    def _norm_url(self, u):
        try:
            p = urlparse(u)
        except Exception:
            return None
        if p.scheme.lower() not in ("http", "https"):
            return None
        frag = p.fragment if p.fragment.startswith(("/", "!/")) else ""
        return urlunparse((p.scheme.lower(), p.netloc.lower(), p.path or "/", "", p.query, frag))

    def _internal(self, u):
        p = urlparse(u)
        return (p.scheme.lower(), p.netloc.lower()) == self.origin

    def _skip(self, url, text=""):
        low = C.norm(url + " " + (text or ""))
        if self.skip_link and self.skip_link.search(low):
            return True
        path = urlparse(url).path.lower()
        return path.endswith(SKIP_EXT)

    def _visit(self, url, collect=True):
        self.resp_log = []
        c0 = len(self.console_errors)
        rec = {"url": url, "path": urlparse(url).path + (("#" + urlparse(url).fragment) if urlparse(url).fragment
                                                           else "")}
        try:
            resp = self._goto(url)
        except Exception as e:
            rec.update({"error": self._friendly(e, {}), "status": 0, "console": [], "failed": [], "title": url,
                        "name": url})
            return rec
        rec["final_url"] = self.page.url
        rec["status"] = resp.status if resp else 200
        try:
            rec["headers"] = resp.all_headers() if resp else {}
        except Exception:
            rec["headers"] = {}
        rec["console"] = list(self.console_errors[c0:])
        rec["failed"] = [{"url": r["url"], "status": r["status"], "type": r["type"], "failed": r.get("failed", "")}
                         for r in self.resp_log if (r["status"] >= 400 or r.get("failed")) and
                         not r["url"].split("?")[0].endswith("favicon.ico")]
        rec["json"] = []
        for r in self.resp_log:
            if r.get("resp") is not None and r.get("req") in self.req_done and len(rec["json"]) < 8:
                try:
                    rec["json"].append(r["resp"].json())
                except Exception:
                    pass
        if collect:
            try:
                rec["info"] = self.page.evaluate(PAGE_INFO_JS)
                if self._thin(rec["info"]):
                    # web dạng SPA hay vẽ nội dung chậm → chờ thêm rồi đọc lại (không dùng time.sleep: _pump)
                    wait = max(3.0, min(60.0, float(self.cfg.get("wait_extra") or 8)))
                    self._info("Trang chưa hiện nút / form nào — chờ thêm tối đa %g giây cho trang tải xong…" % wait)
                    t_end = time.time() + wait
                    while time.time() < t_end:
                        self._check_stop()
                        self._pump(1000)
                        rec["info"] = self.page.evaluate(PAGE_INFO_JS)
                        if not self._thin(rec["info"]):
                            self._settle(300, 2000)
                            rec["info"] = self.page.evaluate(PAGE_INFO_JS)
                            break
                try:
                    rec["diag"] = self.page.evaluate(DIAG_JS)
                except Exception:
                    rec["diag"] = {}
                rec["final_url"] = self.page.url
                rec["perf"] = self.page.evaluate(PERF_JS)
            except Exception as e:
                rec["info"], rec["perf"] = {}, {}
                rec["error"] = "Không đọc được nội dung trang: %s" % self._friendly(e, {})
            info = rec["info"] or {}
            rec["title"] = info.get("title") or url
            rec["name"] = (info.get("h1") or [None])[0] or info.get("title") or rec["path"]
        return rec

    def _phase_discover(self):
        depth = int(self.cfg.get("depth", 1))
        maxp = int(self.cfg.get("max_pages", 10))
        start = self._norm_url(self.url)
        queue, seen = [(start, 0)], {start}
        need_probe = (self.on("functional") and self.o("functional", "buttons") != "none") or self.on("state") or \
            self.on("validation") or self.on("business")
        while queue and len(self.pages) < maxp:
            self._check_stop()
            url, d = queue.pop(0)
            self._info("Dò trang %d/%d: %s" % (len(self.pages) + 1, maxp, url))
            rec = self._visit(url)
            rec["depth"] = d
            if self.logged_in and rec.get("info", {}).get("hasPassword") and len(self.pages) == 0:
                self.emit("warn", msg="Trang đầu vẫn hiện ô mật khẩu — có thể phiên đăng nhập đã hết hạn.")
            self.pages.append(rec)
            if rec.get("error"):
                continue
            nexts = []
            if self._blocked(rec):
                self.emit("warn", msg="Trang %s hiện captcha / chặn công cụ tự động — tool không bấm thử nút trên trang "
                                      "này." % url)
                continue
            if need_probe:
                nexts += self._probe(rec)
            for lk in (rec.get("info") or {}).get("links", []):
                nexts.append((lk["href"], lk.get("text", "")))
            if d < depth:
                for href, text in nexts:
                    n = self._norm_url(href)
                    if not n or n in seen or not self._internal(n) or self._skip(n, text):
                        continue
                    seen.add(n)
                    queue.append((n, d + 1))
        self.discovery = [self._disc_page(p) for p in self.pages]
        self.emit("info", msg="Đã dò %d trang, %d form, %d nút" % (len(self.pages), sum(
            len(p["forms"]) for p in self.discovery), sum(len(p["buttons"]) for p in self.discovery)))
        self.diagnosis = self._diagnose()
        for d in self.diagnosis:
            self.emit("warn", msg="%s — %s" % (d["title"], d["detail"]))

    @staticmethod
    def _blocked(rec):
        dg = rec.get("diag") or {}
        return bool(dg.get("cap") or CAPTCHA_RE.search(dg.get("text") or "") or rec.get("status") in (401, 403, 429))

    @staticmethod
    def _thin(info):
        info = info or {}
        named = [b for b in info.get("buttons", []) if (b.get("text") or b.get("aria"))]
        return not named and not info.get("forms") and not info.get("tables") and len(info.get("links", [])) <= 3

    def _diagnose(self):
        """Dò ra 0 nút / 0 form, gặp captcha, trang chặn tool hoặc cần đăng nhập → nói rõ LÝ DO và CÁCH XỬ LÝ."""
        out = []
        if not self.pages:
            return out
        first = self.pages[0]
        lm = (self.cfg.get("login") or {}).get("mode", "none")
        T_WAIT = "Bấm 'Dò lại – chờ lâu hơn': tool chờ trang tải xong tới 25 giây (web dạng SPA tải chậm)."
        T_MAN = ("Chọn cách đăng nhập 'Tôi tự đăng nhập / vượt captcha trên trình duyệt' rồi Dò lại: tool mở trình duyệt, "
                 "bạn tự đăng nhập hoặc giải captcha, xong bấm 'Đã đăng nhập xong' để tool dò tiếp.")
        T_HEAD = "Tắt 'Chạy ẩn' (⚙ Phạm vi dò & chạy) — nhiều trang chặn trình duyệt chạy ẩn."
        T_PUB = ("Trang công cộng lớn (Douyin, TikTok, Facebook, YouTube…) thường chặn công cụ tự động. Tool dùng để test "
                 "web nội bộ / môi trường test của công ty.")

        def add(code, title, detail, tips, actions=()):
            out.append({"code": code, "title": title, "detail": detail, "tips": [t for t in tips if t],
                        "actions": list(actions)})
        if first.get("error"):
            add("open", "Không mở được trang", first["error"],
                ["Kiểm tra lại địa chỉ (URL), mạng / VPN của công ty.", "Thử mở địa chỉ này bằng trình duyệt thường."])
            return out
        dg = first.get("diag") or {}
        text = dg.get("text") or ""
        st = first.get("status") or 0
        n_btn = sum(len(d["buttons"]) for d in self.discovery)
        n_form = sum(len(d["forms"]) for d in self.discovery)
        head = T_HEAD if self.opt.get("headless") else ""
        a, b = urlparse(first["url"]).netloc.lower(), urlparse(first.get("final_url") or first["url"]).netloc.lower()
        if a and b and a != b:
            add("redirect", "Trang tự chuyển sang tên miền khác",
                "Mở %s nhưng trang chuyển sang %s. Tool chỉ dò trong cùng tên miền với địa chỉ bạn nhập." % (a, b),
                ["Nếu đó là trang đăng nhập chung (SSO): " + T_MAN, "Hoặc nhập thẳng địa chỉ sau khi chuyển."],
                ["manual"])
        if dg.get("cap") or CAPTCHA_RE.search(text) or st == 429:
            add("captcha", "Trang hiện captcha / chặn công cụ tự động",
                "Trang yêu cầu xác minh 'không phải robot' (captcha, trượt ghép hình…) hoặc từ chối trình duyệt do tool "
                "điều khiển%s, nên tool không đọc được nút / form." % (" (HTTP %d)" % st if st >= 400 else ""),
                [T_MAN, head, T_PUB], ["manual"])
        elif st in (401, 403):
            add("denied", "Trang từ chối truy cập (HTTP %d)" % st,
                "Máy chủ không cho tool vào trang — cần đăng nhập, không có quyền, hoặc trang chặn công cụ tự động.",
                [T_MAN, head, T_PUB], ["manual"])
        elif st >= 400:
            add("http", "Trang trả về lỗi HTTP %d" % st, "Máy chủ báo lỗi khi mở trang.",
                ["Kiểm tra lại địa chỉ, thử mở bằng trình duyệt thường."])
        few = n_form == 0 and n_btn <= 8           # dò ra quá ít phần tử → giải thích lý do
        if lm == "none" and ((first.get("info") or {}).get("hasPassword") or
                             (few and LOGIN_WALL_RE.search(text[:2000]))) and not out:
            add("login", "Trang cần đăng nhập",
                "Dò được rất ít chức năng (%d form, %d nút) và trang đang yêu cầu / mời đăng nhập — phần lớn chức năng chỉ "
                "hiện sau khi đăng nhập." % (n_form, n_btn),
                ["Chọn 'Tool tự đăng nhập' (nhập tài khoản) hoặc '" + "Tôi tự đăng nhập / vượt captcha trên trình duyệt'"
                 " rồi Dò lại.", head, T_PUB], ["manual", "wait"])
        if few and not any(x["code"] in ("captcha", "denied", "open", "login") for x in out):
            why = []
            if (dg.get("textLen") or 0) < 150:
                why.append("trang gần như không có chữ (web dạng SPA chưa tải xong, hoặc nội dung vẽ bằng video / canvas)")
            if dg.get("iframes"):
                why.append("nội dung nằm trong %d khung iframe (tool chưa dò bên trong iframe)" % len(dg["iframes"]))
            if dg.get("video") or dg.get("canvas"):
                why.append("trang chủ yếu là video / hình vẽ canvas, không có nút / form HTML thường")
            if dg.get("shadow"):
                why.append("trang dùng Shadow DOM (thành phần đóng gói) nên không đọc được nút bên trong")
            if LOGIN_WALL_RE.search(text[:2000]) and lm == "none":
                why.append("nhiều chức năng chỉ hiện sau khi đăng nhập")
            if not why:
                why.append("trang không có nút / form HTML mà tool nhận ra, hoặc trang chặn công cụ tự động")
            add("empty", "Dò được %d trang nhưng chỉ có %d form, %d nút" % (len(self.pages), n_form, n_btn),
                "Lý do có thể: " + "; ".join(why) + ".", [T_WAIT, T_MAN, head, T_PUB], ["wait", "manual"])
        return out

    def _disc_page(self, p):
        """Những gì DÒ ĐƯỢC trên trang — giao diện dùng làm dropdown cho tester chọn (không đoán)."""
        info = p.get("info") or {}
        seen, buttons = set(), []
        for b in info.get("buttons", []):
            t = b["text"] or b["aria"]
            if t and t.lower() not in seen:
                seen.add(t.lower())
                buttons.append({"text": t, "allowed": b.get("allowed", self._safe_button(b)),
                                "inTable": b.get("inTable"), "inForm": b.get("inForm")})
        forms = []
        for f in info.get("forms", []):
            if f.get("kind") != "data" or not (f.get("visible") or f.get("opener")):
                continue
            forms.append({"sel": f["sel"], "name": f.get("name") or (f.get("submit") or {}).get("text", "") or "form",
                          "opener": (f.get("opener") or {}).get("text", ""),
                          "buttons": [b["text"] for b in f.get("buttons", [])],
                          "fields": [{"sel": x["sel"], "label": x["label"] or x["name"], "type": x["type"],
                                      "required": x["required"], "maxlength": x["maxlength"], "min": x["min"],
                                      "max": x["max"], "tag": x["tag"]} for x in f["fields"]
                                     if self._rows_valid(x) or x.get("disabled")]})
        return {"url": p["url"], "path": p.get("path", ""), "title": p.get("name", ""), "status": p.get("status"),
                "error": p.get("error", ""), "buttons": buttons,
                "popups": [{"opener": x["opener"]["text"], "buttons": x.get("buttons", [])} for x in p.get("popups", [])],
                "forms": forms, "statusValues": info.get("statusValues", [])}

    def _safe_button(self, b):
        txt = C.norm("%s %s" % (b.get("text", ""), b.get("aria", "")))
        if not txt:
            return False
        if b.get("type") == "submit" and not re.search(r"\b(tim|search|loc|filter)\b", txt):
            return False
        if self.unsafe and self.unsafe.search(txt):
            return False
        return not b.get("href")

    def _allowed(self, p, text, default):
        """Tester đã tick/bỏ tick nút này ở bước 'Dò trang' chưa? Chưa thì dùng mặc định."""
        m = (self.cfg.get("click") or {}).get(self._norm_url(p["url"])) or {}
        return bool(m[text]) if text in m else default

    def _probe(self, rec):
        """Bấm thử từng nút an toàn: ghi lại lỗi, popup/form mở ra, trang chuyển tới. Trả về link mới để dò."""
        info = rec["info"]
        limit = int(self.o("functional", "max_buttons") or 15)
        seen, cands, skipped = set(), [], []
        for b in info.get("buttons", []):
            key = (b["text"] or b["aria"]).strip().lower()
            if not key or key in seen:
                continue
            seen.add(key)
            label = b["text"] or b["aria"]
            b["allowed"] = self._allowed(rec, label, self._safe_button(b))
            if b["allowed"]:
                cands.append(b)
            else:
                skipped.append(label)
        rec["clicks"], rec["skipped"], rec["popups"] = [], skipped[:20], []
        nexts = []
        dirty = True
        todo = cands[:limit]
        t_end = time.time() + 150          # trang nặng / chậm: không bấm thử quá 2,5 phút mỗi trang
        for bi, b in enumerate(todo, start=1):
            self._check_stop()
            if time.time() > t_end:
                self.emit("warn", msg="Trang %s tải chậm — tool dừng bấm thử sau %d/%d nút để không chờ quá lâu."
                          % (rec["url"], bi - 1, len(todo)))
                break
            self._info("Bấm thử nút %d/%d trên trang %s: '%s'" % (bi, len(todo), rec.get("name") or rec["url"],
                                                                (b["text"] or b["aria"])[:40]))
            try:
                if dirty:
                    self._goto(rec["url"])
                before = self.page.evaluate(SNAP_JS)
                c0, d0, n0 = len(self.console_errors), len(self.dialogs), len(self.resp_log)
                pages0 = len(self.ctx.pages)
                err = ""
                try:
                    self.page.locator(b["sel"]).first.click(timeout=4000)
                except Exception as e:
                    err = self._friendly(e, {})
                self._settle(500, 1500)
                newtab = []
                for p in self.ctx.pages[pages0:]:
                    try:
                        p.wait_for_load_state("load", timeout=5000)
                        newtab.append(p.url)
                        p.close()
                    except Exception:
                        pass
                after = self.page.evaluate(SNAP_JS)
            except E.Stopped:
                raise
            except Exception as e:
                rec["clicks"].append({"text": b["text"] or b["aria"], "sel": b["sel"], "error": self._friendly(e, {}),
                                      "js": [], "http": [], "nav": "", "newtab": [], "opened": [], "dialogs": []})
                dirty = True
                continue
            opened = [s for s in after["forms"] + after["dialogs"] if s not in before["forms"] + before["dialogs"]]
            nav = after["url"] if self._norm_url(after["url"]) != self._norm_url(before["url"]) else ""
            changed = bool(opened or nav or after["visCount"] != before["visCount"] or after["exp"] != before["exp"]
                           or abs(after["textLen"] - before["textLen"]) > 20)
            http = ["%s (%d)" % (r["url"][-80:], r["status"]) for r in self.resp_log[n0:] if r["status"] >= 500]
            click = {"text": b["text"] or b["aria"], "sel": b["sel"], "error": err, "js": self.console_errors[c0:],
                     "http": http, "nav": nav, "newtab": newtab, "opened": opened, "dialogs": self.dialogs[d0:],
                     "changed": changed, "inTable": b.get("inTable")}
            rec["clicks"].append(click)
            if nav:
                nexts.append((nav, b["text"]))
            nexts += [(u, b["text"]) for u in newtab]
            for f in info.get("forms", []):
                if f["sel"] in opened and not f.get("visible") and not f.get("opener"):
                    f["opener"] = {"text": b["text"] or b["aria"], "sel": b["sel"]}
            dlg = [s for s in after["dialogs"] if s not in before["dialogs"]]
            if dlg or opened:
                try:
                    inside = self.page.evaluate(BTN_IN_JS, (dlg or opened)[:1])
                except Exception:
                    inside = []
                rec["popups"].append({"opener": {"text": b["text"] or b["aria"], "sel": b["sel"]},
                                      "dialog": (dlg or opened)[0], "buttons": inside, "is_dialog": bool(dlg)})
            dirty = changed or bool(err)
        return nexts

    # ------------------------------------------------------------------ 3. SINH TEST CASE
    def _make_plan(self):
        """Sinh TEST SCENARIO theo từng màn hình (mỗi tính năng 1 scenario, có cột Nhóm), test case nằm dưới
        scenario. Nếu tester đưa file Test Scenario → ghép từng scenario của tester với các bài kiểm tra phù hợp."""
        pages = list(self.pages)
        ok_pages = [p for p in pages if not p.get("error")]
        P = []
        for p in pages:
            good = not p.get("error")
            if good and self.on("ui"):
                P += self._plan_ui(p)
            if good and self.on("responsive"):
                P += self._plan_responsive(p)
            if good and self.on("validation"):
                P += self._plan_validation(p)
            if self.on("functional"):
                P += self._plan_functional(p)
            if good and self.on("integration"):
                P += self._plan_integration(p)
            if self.on("business"):
                P += self._plan_business(page=p)
            if good and self.on("state"):
                P += self._plan_state(p)
                P += self._plan_state_flows(page=p)
            if good and self.on("usability"):
                P += self._plan_usability(p)
            if good and self.on("performance"):
                P += self._plan_perf(p)
            if good and self.on("accessibility"):
                P += self._plan_a11y(p)
        if self.on("business"):
            P += self._plan_business(page=None)
        if self.on("state"):
            P += self._plan_state_flows(page=None)
        if self.on("permission"):
            P += self._plan_permission(ok_pages)
        if self.on("security"):
            P += self._plan_security(ok_pages)
        if self.on("compat"):
            P += self._plan_compat(ok_pages)
        if self.imported:
            P = self._map_imported(P)
        else:
            if self.cfg.get("suggest", True) is not False:
                for p in ok_pages:
                    P += self._plan_suggest(p)
            P = self._order(P)
        self._number_tcs(P)
        if self.only:
            P = [it for it in P if it[2].get("key") in self.only]
        only_tc = set(self.cfg.get("only_tc") or [])
        if only_tc:
            P = [it for it in P if it[3].get("tc_id") in only_tc]
        keep = {it[2]["id"] for it in P} | {sc["id"] for sc in self.wb["scenarios"] if sc["tcs"]}
        self.wb["scenarios"] = [sc for sc in self.wb["scenarios"] if sc["id"] in keep]
        return P

    def main_screen(self):
        ok = [p for p in self.pages if not p.get("error")]
        return (ok[0].get("name") if ok else "") or urlparse(self.url).netloc or "Web"

    def screens(self):
        """Mã màn hình MH_01, MH_02… theo thứ tự trang dò được."""
        out, seen = [], set()
        for p in self.pages:
            nm = p.get("name") or p["url"]
            if p.get("error") or nm in seen:
                continue
            seen.add(nm)
            out.append({"mh": "MH_%02d" % (len(out) + 1), "name": nm, "url": p["url"]})
        return out

    def _order(self, P):
        """Sắp scenario theo LUỒNG (như file mẫu), đánh lại ID liên tục TS_<mã màn hình chính>_001…; test case
        xếp theo đúng thứ tự scenario."""
        mh = {s["name"]: i for i, s in enumerate(self.screens())}
        code = self._code(self.main_screen())
        used = {it[2]["id"] for it in P}
        scns = [sc for sc in self.wb["scenarios"] if sc["id"] in used and not sc["tcs"]]
        scns.sort(key=lambda sc: (sc.get("luong") or 9, mh.get(sc.get("module"), 999),
                                  FEAT_RANK.get(sc.get("feature"), 50), sc.get("n", 0)))
        for i, sc in enumerate(scns, start=1):
            sc["id"] = sc["fields"]["ts_id"] = "TS_%s_%03d" % (code, i)
            sc["code"] = code
        pos = {id(sc): i for i, sc in enumerate(scns)}
        self.wb["scenarios"] = [sc for sc in self.wb["scenarios"] if id(sc) not in pos] + scns
        return sorted(P, key=lambda it: pos.get(id(it[2]), -1))

    def _number_tcs(self, P):
        """Đánh mã TC ngay khi lập kế hoạch (trước khi lọc theo dòng đã tick) → mã ở bảng xem trước, file Excel sinh
        ra và lần chạy luôn trùng nhau."""
        cnt = dict(self.counters)
        for it in P:
            if it[3].get("tc_id"):
                continue
            code = self._scn_code(it[2])
            while True:
                cnt[code] = cnt.get(code, 0) + 1
                tid = ("TC_%s_%03d" % (code, cnt[code])) if code else "TC%03d" % cnt[code]
                if tid not in self.used_tc:
                    break
            it[3]["tc_id"] = tid

    # ---------- ghép Test Scenario của tester (đọc từ file Excel) với bài kiểm tra tự động
    KW = {
        "functional": [("search", r"tim kiem|search"), ("paging", r"phan trang|trang tiep|paging"),
                       ("sorting", r"sap xep|sort"), ("buttons", r"cot hanh dong|thao tac|button|nut|click"),
                       ("links", r"\blink\b|lien ket"), ("form_actions", r"\bluu\b|gui|tao moi|them moi|phe duyet|"
                                                                     r"nhap du lieu|tai tep|import|tai len"),
                       ("login", r"dang nhap|login")],
        "validation": [("required", r"bo trong|bat buoc|\(\*\)|required"),
                       ("format", r"dinh dang|email|so dien thoai|sdt|format"),
                       ("length", r"do dai|ky tu|gioi han|toi da|toi thieu|min|max|so luong|gia tri"),
                       ("file", r"tep|file|excel|json|tai len|upload|nhap du lieu"),
                       ("search_input", r"tim kiem|search")],
        "integration": [("mapping", r"mapping|du lieu tra ve|khop|api"), ("dropdown", r"dropdown|danh sach chon"),
                        ("loading", r"dang tai|loading|treo"), ("api", r"\bapi\b|backend|may chu")],
        "state": [("toggle", r"bo loc|menu|mo/dong|mo dong|mo rong|thu gon|an/hien"),
                  ("popup", r"popup|hop thoai|bieu tuong x|nut dong"), ("switch", r"toggle|bat/tat|bat tat|nut gat|an bai"),
                  ("flow", r"trang thai|cho duyet|phe duyet|het han|ban nhap|luu nhap")],
        "usability": [("ui", r"giao dien|bo cuc|mac dinh|layout"),
                      ("responsive", r"kich thuoc man hinh|thay doi kich thuoc|responsive|mobile|dien thoai"),
                      ("ease", r"che khuat|de bam|de thao tac|nut gat|toggle|kho bam|qua nho"),
                      ("popup", r"popup")],
    }
    NHOM_TYPES = {"functional": ["functional"], "function": ["functional"], "chuc nang": ["functional"],
                  "validation": ["validation"], "integration": ["integration"], "tich hop": ["integration"],
                  "business rule": ["business"], "business": ["business"], "nghiep vu": ["business"],
                  "permission": ["permission"], "phan quyen": ["permission"], "state": ["state"],
                  "trang thai": ["state"], "usability": ["usability", "ui", "responsive", "state"],
                  "ui": ["ui", "usability"], "giao dien": ["ui", "usability"], "performance": ["performance"],
                  "hieu nang": ["performance"], "security": ["security"], "bao mat": ["security"],
                  "accessibility": ["accessibility"], "compatibility": ["compat"]}

    @staticmethod
    def _words(t):
        stop = {"kiem", "tra", "trang", "cua", "va", "cac", "khi", "tren", "the", "theo", "voi", "duoc", "co", "khong",
                "la", "mot", "nhung", "cho", "tai", "form", "chuc", "nang", "du", "lieu"}
        return {w for w in C.norm(t).split() if len(w) >= 2 and w not in stop}

    FEAT_NAME = {"search": "ô tìm kiếm", "paging": "phân trang", "sorting": "cột sắp xếp được", "buttons": "nút thao tác",
                 "links": "liên kết", "form_actions": "nút lưu / gửi của form", "login": "trang đăng nhập",
                 "required": "trường bắt buộc (*)", "format": "trường email / số điện thoại", "length":
                 "trường có giới hạn độ dài / min-max", "file": "ô tải tệp", "search_input": "ô tìm kiếm",
                 "mapping": "dữ liệu API để đối chiếu", "dropdown": "dropdown", "loading": "dữ liệu tải",
                 "api": "API", "toggle": "bộ lọc / menu mở-đóng", "popup": "popup", "switch": "nút bật/tắt",
                 "flow": "luồng trạng thái đã khai báo", "ui": "giao diện", "responsive": "trang",
                 "ease": "nút / nút gạt", "rule": "luật nghiệp vụ đã khai báo", "perm": "phân quyền"}
    GROUP = {"ui": "usability", "responsive": "usability", "usability": "usability"}
    REAL = {"ui": "ui", "responsive": "responsive", "ease": "usability", "popup": "state"}

    def _pairs(self, nt, tkeys):
        pairs = set()
        for tk in tkeys:
            g = self.GROUP.get(tk, tk)
            for feat, rx in self.KW.get(g, []):
                if re.search(rx, nt):
                    pairs.add((self.REAL.get(feat, g) if g == "usability" else g, feat))
        return pairs

    def _map_imported(self, P):
        groups = {}
        for it in P:
            groups.setdefault(it[2]["key"], []).append(it)
        gen = [sc for sc in self.wb["scenarios"] if sc.get("key") in groups]
        codes = set(self.codes.values())
        out, new_scns = [], [sc for sc in self.wb["scenarios"] if sc["tcs"]]   # giữ scenario đăng nhập đã chạy
        for x in self.imported:
            text = x.get("scenario") or ""
            nt, ng = C.norm(text), C.norm(x.get("nhom") or "")
            tkeys = self.NHOM_TYPES.get(ng) or next((v for k, v in self.NHOM_TYPES.items() if k in ng), []) or \
                ["functional"]
            pairs = self._pairs(nt, tkeys)
            if not pairs and "business" not in tkeys:   # nhóm ghi một đằng, nội dung là việc khác (VD 'validation' nhưng nói mapping API)
                pairs = self._pairs(nt, ["functional", "validation", "integration", "state", "usability"])
            if "business" in tkeys:
                pairs.add(("business", "rule"))
            if "permission" in tkeys:
                pairs.add(("permission", "perm"))
            sid = x.get("id") or "TS_%03d" % (len(new_scns) + 1)
            parts = sid.split("_")
            code = "_".join(parts[1:-1]) if len(parts) >= 3 else ""
            module = next((m for m, c in self.codes.items() if c == code), "")
            reason = ""
            cands = [sc for sc in gen if (sc["type"], sc["feature"]) in pairs]
            if "business" in tkeys:
                ac = C.norm(x.get("ac") or "")
                tw = self._words(text)
                cands = [sc for sc in cands if sc["type"] != "business" or (ac and C.norm(sc["fields"].get("ac") or "")
                                                                         == ac) or
                         len(self._words(sc["fields"]["scenario"]) & tw) >= 2]
            if code and code in codes:              # mã màn hình của scenario (TS_DHTD_…) → đúng màn hình đó
                cands = [sc for sc in cands if sc["code"] == code or sc["type"] in ("permission", "security",
                                                                                    "compat")]
            else:
                tw = self._words(text)
                sc_best = max([len(self._words(sc.get("module") or "") & tw) for sc in cands] or [0])
                if sc_best:
                    cands = [sc for sc in cands if len(self._words(sc.get("module") or "") & tw) == sc_best]
            if "form" in nt or "popup" in nt:
                tw = self._words(text)
                fm = [sc for sc in cands if len(self._words(sc["fields"]["scenario"]) & tw) >= 3]
                if fm:
                    cands = fm
            items = []
            for sc in cands:
                items += groups.get(sc["key"], [])
            quotes = re.findall(r"['\"“”‘’]([^'\"“”‘’]{2,60})['\"“”‘’]", text)
            # tên trong ngoặc ('Ẩn bài viết', 'Nhập liệu dữ liệu') = phần tử cụ thể → chỉ nhận đúng phần tử đó
            quotes = [q for q in quotes if self._words(q) and not re.search(r"^(het han|hoat dong|da duyet|cho duyet|"
                                                                            r"ban nhap|nhap)$", C.norm(q))]
            named = {"popup", "switch", "toggle", "form_actions", "buttons", "file", "dropdown"}
            if quotes and items:
                qw = set().union(*[self._words(q) for q in quotes])
                hit = [it for it in items if it[2]["feature"] not in named or self._words(it[1]) & qw]
                if not hit:
                    reason = "Không tìm thấy %s trên trang" % ", ".join("'%s'" % q for q in quotes)
                items = hit
            if not items and not reason:
                if "business" in tkeys:
                    reason = "Chưa khai báo luật nghiệp vụ cho scenario này (mục Business rule)"
                elif ("state", "flow") in pairs and len(pairs) == 1:
                    reason = "Chưa khai báo luồng trạng thái cho scenario này (mục State)"
                elif pairs:
                    reason = "Không tìm thấy %s trên %s" % (", ".join(sorted({self.FEAT_NAME.get(f, f) for _, f in pairs if f != "flow"} or {"luồng trạng thái đã khai báo"})),
                                                            ("trang " + module) if module else "các trang đã dò")
                else:
                    reason = "Tool chưa tự động được nội dung này (luồng nghiệp vụ nhiều bước / dữ liệu riêng)"
            S = {"id": sid, "key": "IN|" + sid, "type": tkeys[0], "feature": ",".join(sorted(f for _, f in pairs)),
                 "code": code, "module": module or (cands[0]["module"] if cands else ""), "imported": True,
                 "fields": {"ts_id": sid, "nhom": x.get("nhom") or "", "module": module or (cands[0]["module"] if cands
                                                                                           else ""),
                            "scenario": text, "ac": x.get("ac") or "", "priority": x.get("priority") or ""},
                 "extra": {}, "row": None, "tcs": [], "reason": reason}
            items = [(tk, title, S, dict(kw, priority=x.get("priority") or kw.get("priority")), fn)
                     for (tk, title, _, kw, fn) in items]
            if not items:
                items = [self._manual_item(S, x, reason)]
            out += items
            new_scns.append(S)
        self.wb["scenarios"] = new_scns
        return out

    # ---------- scenario GỢI Ý theo tính năng dò được (văn phong file mẫu) — tool chưa tự động được → test tay
    def _plan_suggest(self, p):
        info = p.get("info") or {}
        name = p.get("name") or p["url"]
        out = []
        have = {(sc["feature"], sc["module"]) for sc in self.wb["scenarios"]}
        why = "Tool chưa tự động được — cần người test kiểm tra với dữ liệu / nghiệp vụ thật"

        def add(tk, feat, title, prio="High", luong=None):
            if not self.on(tk):
                return
            sc = self._S(tk, feat, name, title, prio=prio, luong=luong)
            if sc["tcs"] or any(it[2] is sc for it in out):
                return
            sc["reason"] = why
            x = {"scenario": title, "ac": "", "priority": prio}
            it = self._manual_item(sc, x, why, url=p["url"])
            it[3]["page"] = p
            out.append(it)

        N = C.norm
        btns = [b for b in info.get("buttons", []) if (b.get("text") or b.get("aria"))]
        bt = lambda b: (b.get("text") or b.get("aria") or "").strip()  # noqa: E731
        tables = [t for t in info.get("tables", []) if t.get("headers") or t.get("rows")]
        forms = info.get("forms", [])
        dforms = self._data_forms(p)
        filt = [x for f in forms if f.get("kind") in ("filter", "search") for x in f.get("fields", [])
                if x["tag"] == "select" or x["type"] in ("date", "datetime-local", "month", "checkbox", "radio")]
        flabels = list(dict.fromkeys(x["label"] for x in filt if x.get("label")))[:7]
        search = info.get("search", [])
        sortable = any(h.get("sortable") for t in tables for h in t.get("headers", []))
        stv = info.get("statusValues", [])
        files = [x for f in forms for x in f.get("fields", []) if x["type"] == "file"]
        popups = [pp["opener"]["text"] for pp in p.get("popups", [])] + \
            [f["opener"]["text"] for f in forms if f.get("opener")]
        q = lambda s: "'%s'" % s  # noqa: E731
        # LUỒNG 1
        if tables or dforms:
            add("usability", "labels", "Kiểm tra nhãn, placeholder và tên trạng thái hiển thị đúng ngôn ngữ, thống nhất "
                                       "giữa danh sách, bộ lọc, form và chi tiết của trang %s" % name, "Medium")
        # LUỒNG 2 — hiển thị danh sách
        for t in tables[:1]:
            heads = [h["text"] for h in t.get("headers", []) if h.get("text")]
            dates = [h for h in heads if re.search(r"\bngay\b|date|thoi gian|\bhan\b", N(h))]
            if dates:
                add("validation", "date_col", "Kiểm tra cột %s hiển thị đúng định dạng dd/MM/yyyy và hợp lệ với dữ liệu"
                    % ", ".join(dates[:4]))
            sts = [h for h in heads if re.search(r"trang thai|tinh trang|status", N(h))]
            if sts:
                add("functional", "status_col", "Kiểm tra cột %s hiển thị đúng giá trị%s theo dữ liệu và điều kiện "
                                                "nghiệp vụ" % (", ".join(sts[:2]), (" (%s)" % ", ".join(stv[:6])) if stv
                                                               else ""))
            if ("mapping", name) not in have:
                add("integration", "mapping", "Kiểm tra dữ liệu hiển thị trên danh sách %s được mapping đúng với dữ "
                                              "liệu trả về" % (name[:1].lower() + name[1:]))
        # LUỒNG 3 — tìm kiếm & bộ lọc
        if flabels:
            add("functional", "filter", "Kiểm tra lọc theo từng tiêu chí %s và thời điểm điều kiện lọc được áp dụng"
                % ", ".join(flabels))
        reset = next((bt(b) for b in btns if re.match(r"^(lam moi|reset|xoa bo loc|xoa loc|dat lai|clear|bo loc mac "
                                                      r"dinh)", N(bt(b)))), "")
        if reset and (search or flabels):
            add("functional", "reset", "Kiểm tra nút %s xóa toàn bộ điều kiện lọc, từ khóa tìm kiếm và tải lại danh "
                                       "sách mặc định" % reset, "Medium")
        if search:
            add("validation", "no_result", "Kiểm tra tìm kiếm với từ khóa không tồn tại và thông báo khi không có kết "
                                           "quả", "Medium")
        # LUỒNG 4 — phối hợp chức năng danh sách
        if tables and (search or flabels):
            if search and flabels:
                add("functional", "combo", "Kiểm tra kết hợp từ khóa tìm kiếm với nhiều tiêu chí lọc cùng lúc cho kết "
                                           "quả thỏa mãn đồng thời tất cả điều kiện")
            if sortable:
                add("functional", "sort_filter", "Kiểm tra sắp xếp tăng/giảm từng cột khi danh sách đang được tìm "
                                                 "kiếm hoặc lọc", "Medium")
            if info.get("pager"):
                add("functional", "page_filter", "Kiểm tra phân trang khi đang tìm kiếm hoặc lọc: số trang, tổng số "
                                                 "mục và quay về trang 1 khi thay đổi điều kiện", "Medium")
            if dforms or any(re.search(r"\b(xem|chi tiet|view|detail)\b", N(bt(b))) for b in btns):
                add("functional", "keep_state", "Kiểm tra giữ nguyên từ khóa, điều kiện lọc, sắp xếp và trang hiện tại "
                                                "khi quay lại danh sách từ Chi tiết hoặc Thêm mới", "Low")
        # LUỒNG 5 / 6 — form thêm mới, lưu & vòng đời
        manual_btns = self.o("functional", "form_actions") != "yes"
        for f in dforms:
            if any(x["type"] == "file" for x in f["fields"]):
                continue
            fname = f.get("name") or (f.get("submit") or {}).get("text") or "form"
            fl = [x for x in f["fields"] if x.get("label")]
            add("usability", "form_ui", "Kiểm tra giao diện, bố cục và giá trị mặc định của form %s" % fname, "Medium")
            for x in fl:
                if re.search(r"^ma\b|\bma (don|so|nhan vien|hang|khach)|\bcode\b", N(x["label"])):
                    add("validation", "unique", "Kiểm tra không cho lưu khi %s trùng với bản ghi đã tồn tại" % x["label"])
                    break
            dates = [x["label"] for x in fl if x["type"] in ("date", "datetime-local") or re.search(r"\bngay\b", N(x["label"]))]
            if len(dates) >= 2:
                add("validation", "date_rule", "Kiểm tra ràng buộc giữa %s trên form %s" % (", ".join(dates[:4]), fname))
            fb = [b["text"] for b in f.get("buttons", []) if b.get("text")]
            if manual_btns:
                for b in fb:
                    nb = N(b)
                    if self.CANCEL_RE.match(nb):
                        add("functional", "cancel", "Kiểm tra nút %s trên form %s không lưu dữ liệu và xác nhận trước khi "
                                                    "rời form đã nhập" % (b, fname), "Medium")
                    elif re.search(r"luu nhap|draft", nb):
                        st = next((v for v in stv if re.search(r"nhap|draft", N(v))), "Bản nháp")
                        add("functional", "btn_save", "Kiểm tra %s tạo bản ghi ở trạng thái %s và hiển thị trên danh sách"
                            % (b, st))
                    elif re.search(r"gui (phe )?duyet|trinh duyet|submit for approval", nb):
                        st = next((v for v in stv if re.search(r"cho duyet|pending", N(v))), "Chờ duyệt")
                        add("functional", "btn_save", "Kiểm tra %s tạo bản ghi ở trạng thái %s và hiển thị trên danh sách"
                            % (b, st))
                    elif re.match(r"^(luu|luu lai|them|them moi|tao|tao moi|cap nhat|save|create|add|submit)$", nb):
                        add("functional", "btn_save", "Kiểm tra %s lưu bản ghi thành công, hiển thị thông báo và bản ghi "
                                                      "mới trên danh sách" % b)
        if any(re.search(r"\b(phe duyet|duyet|approve)\b", N(bt(b))) and not re.search(r"\bgui\b", N(bt(b)))
               for b in btns):
            add("functional", "approve", "Kiểm tra phê duyệt bản ghi đang chờ duyệt và các thay đổi của bản ghi sau khi "
                                         "được duyệt")
        if any(re.search(r"tu choi|reject", N(bt(b))) for b in btns):
            add("functional", "reject", "Kiểm tra từ chối bản ghi đang chờ duyệt và nhánh xử lý sau khi từ chối (lý do, "
                                        "chỉnh sửa, gửi lại)")
        if len(stv) >= 2 and any(re.search(r"duyet|nhap|het han|hoat dong|approved|pending", N(v)) for v in stv):
            add("functional", "lifecycle", "Kiểm tra các thao tác được phép hoặc bị chặn theo từng trạng thái của bản "
                                           "ghi (%s)" % ", ".join(stv[:6]), luong=6)
        # LUỒNG 7 — xem chi tiết
        view = next((bt(b) for b in btns if b.get("inTable") and re.search(r"\b(xem|chi tiet|view|detail)\b",
                                                                             N(bt(b)))), "")
        if view:
            add("functional", "detail", "Kiểm tra mở trang Chi tiết đúng bản ghi từ nút %s trên danh sách" % q(view))
            add("validation", "detail_data", "Kiểm tra dữ liệu và bố cục trang Chi tiết khớp với bản ghi đã lưu, các "
                                             "trường ở chế độ chỉ xem")
        # LUỒNG 8 — hành động trên bản ghi
        rowb = list(dict.fromkeys(bt(b) for b in btns if b.get("inTable") and bt(b) != view))
        used = set()
        for b in rowb:
            nb = N(b)
            if re.search(r"\b(sua|chinh sua|edit|cap nhat)\b", nb):
                add("functional", "edit", "Kiểm tra sửa bản ghi từ nút %s: mở đúng dữ liệu, lưu thay đổi và cập nhật "
                                          "trên danh sách" % q(b))
            elif re.search(r"\b(xoa|delete|remove)\b", nb):
                add("functional", "delete", "Kiểm tra xóa bản ghi từ nút %s: hiện xác nhận, xóa thành công và cập nhật "
                                            "danh sách; hủy xác nhận thì không xóa" % q(b))
            elif re.search(r"sao chep|copy", nb):
                add("functional", "copy", "Kiểm tra %s: nội dung sao chép, thông báo và kết quả khi dùng" % q(b), "Medium")
            else:
                continue
            used.add(b)
        other = [b for b in rowb if b not in used and not re.search(r"duyet|tu choi|approve|reject", N(b)) and
                 len(b) <= 30][:5]
        if other:
            add("functional", "row_action", "Kiểm tra các thao tác %s trên từng bản ghi được phép hoặc bị chặn theo "
                                            "trạng thái của bản ghi" % ", ".join(q(b) for b in other))
        exp = next((bt(b) for b in btns if re.search(r"\b(xuat|export)\b", N(bt(b)))), "")
        if exp:
            add("functional", "export", "Kiểm tra %s dữ liệu danh sách ra file đúng dữ liệu đang hiển thị / đang lọc"
                % q(exp), "Medium")
        # LUỒNG 9 — nhập dữ liệu từ file
        imp = next((t for t in [bt(b) for b in btns] + popups if re.search(r"nhap du lieu|import|tai len|upload|nhap tu "
                                                                           r"file|nhap excel", N(t))), "")
        if imp or files:
            acc = ", ".join(dict.fromkeys(x.get("accept") for x in files if x.get("accept"))) or "Excel"
            tpl = next((bt(b) for b in btns if re.search(r"file mau|tai mau|template|mau excel", N(bt(b)))), "")
            if tpl:
                add("functional", "template", "Kiểm tra tải file mẫu từ nút %s: đúng định dạng, đủ cột, mở được"
                    % q(tpl), "Medium")
            add("functional", "import_ok", "Kiểm tra nhập dữ liệu%s từ file %s hợp lệ" % (
                (" bằng " + q(imp)) if imp else "", acc))
            if ("file", name) not in have:
                add("validation", "import_bad", "Kiểm tra nhập dữ liệu khi không chọn file, sai định dạng, file rỗng "
                                                "hoặc có dòng lỗi (thiếu trường bắt buộc, trùng mã)")
        return out

    def _manual_item(self, S, x, why="", url=None):
        """Scenario tool chưa tự động được → test case TEST TAY có ghi rõ lý do, tool dừng hỏi người test."""
        why = why or "Tool chưa tự động được scenario này"
        tk = S["type"] if S["type"] in T.BY_KEY else "functional"

        def fn(res):
            res["mode"] = res["fields"]["cach_test"] = "Tay"
            res["mode_reason"] = why
            res["note"] = "✋ Test tay — " + why
            s = self._st(res, x.get("scenario") or "", "Đúng theo yêu cầu %s" % (x.get("ac") or "của scenario"))
            try:
                self._goto(url or self.url)
                self._shot(s, "tay")
            except Exception:
                pass
            if self.o("validation", "unsure") == "ask":
                self._ask_manual(self.inst, res, [s], reason="✋ TEST TAY — %s" % why)
            else:
                s["status"], s["actual"] = "Chưa test", "✋ Cần test tay: %s" % why
        return (tk, "✋ Test tay: %s" % (x.get("scenario") or "")[:80], S,
                {"module": S["module"], "mode": "Tay", "priority": x.get("priority") or "Medium"}, fn)

    def plan_preview(self):
        """Không mở trình duyệt: trả về danh sách Test Scenario + test case sẽ chạy (cho bảng ở giao diện)."""
        plan = self._make_plan()
        by = {}
        mh = {s["name"]: s["mh"] for s in self.screens()}
        for tk, title, sc, kw, fn in plan:
            by.setdefault(sc["id"], {"titles": [], "modes": set(), "items": []})
            by[sc["id"]]["titles"].append(title)
            by[sc["id"]]["modes"].add(kw.get("mode") or "Tự động")
            res = self._res(tk, title, sc, **kw)
            steps, exp = self._plan_steps(tk, sc, title, kw)
            f = res["fields"]
            by[sc["id"]]["items"].append({"tc_id": res["tc_id"], "title": title, "mode": kw.get("mode") or "Tự động",
                                          "module": f["module"], "priority": f["priority"], "pre": f["precondition"],
                                          "data": f["test_data"], "data_id": f["data_id"], "env": f["moi_truong"],
                                          "size": f["kich_thuoc"], "url": f["mo_ta"], "steps": steps,
                                          "expected": exp, "nhom": f["loai_test"]})
        out = []
        for sc in self.wb["scenarios"]:
            if sc["id"] not in by:
                continue
            m = by[sc["id"]]["modes"]
            mode = "Tay" if m == {"Tay"} else ("Bán tự động" if ("Tay" in m or "Bán tự động" in m) else "Tự động")
            lu = sc.get("luong") or luong_of_text(sc["fields"]["scenario"])
            out.append({"key": sc.get("key"), "id": sc["id"], "nhom": sc["fields"]["nhom"],
                        "scenario": sc["fields"]["scenario"], "ac": sc["fields"].get("ac", ""),
                        "priority": sc["fields"].get("priority", ""), "module": sc.get("module", ""),
                        "mh": mh.get(sc.get("module"), ""), "luong": lu, "luong_name": LUONG.get(lu, ""),
                        "tcs": by[sc["id"]]["titles"], "tc_items": by[sc["id"]]["items"], "mode": mode,
                        "reason": sc.get("reason", ""), "imported": bool(sc.get("imported"))})
        return out

    def plan_meta(self):
        main = self.main_screen()
        lg = (self.cfg.get("login") or {})
        return {"main": main, "code": self._code(main), "screens": self.screens(), "luong": LUONG, "url": self.url,
                "role": lg.get("user") if lg.get("mode") == "auto" else ("Người test tự đăng nhập" if lg.get("mode") ==
                                                                          "manual" else "Khách (không đăng nhập)")}

    # ---------- các bước dự kiến của test case (cho file Excel sinh TRƯỚC khi chạy; chạy xong tool ghi bước thật)
    def _plan_steps(self, tk, sc, title, kw):
        pg = kw.get("page") or {}
        url = pg.get("url") or self.url
        feat = sc.get("feature") or ""
        data = kw.get("data") or ""
        pre = kw.get("pre") or ""
        form = re.search(r"mở form '([^']*)'", pre)
        form = form.group(1) if form else ""
        q = re.search(r"'([^']+)'", title)
        thing = q.group(1) if q else ""
        open_ = "Mở trang %s" % url
        if kw.get("mode") == "Tay" or sc.get("imported") and title.startswith("✋"):
            return ([open_, sc["fields"]["scenario"]], ["Trang mở được", "Đúng theo yêu cầu%s" % (
                " " + sc["fields"]["ac"] if sc["fields"].get("ac") else "")])
        if tk == "validation" and feat == "search_input":
            return ([open_, "Nhập vào ô tìm kiếm: %s" % data, "Bấm Tìm kiếm / Enter"],
                    ["Có ô tìm kiếm", "", "Không lỗi hệ thống, không thực thi mã, hiển thị thông báo phù hợp"])
        if tk in ("validation", "business"):
            st = [open_] + (["Mở form '%s'" % form] if form else []) + ["Nhập dữ liệu: %s" % (data or title),
                                                                         "Bấm nút gửi / lưu form"]
            ok = feat == "valid" or "hợp lệ và gửi" in title
            return st, [""] * (len(st) - 1) + ["Lưu thành công, hiển thị thông báo thành công" if ok else
                                               "Hiển thị thông báo lỗi tại trường, không lưu dữ liệu"]
        T_ = {
            "page": ([open_, "Quan sát trang và Console của trình duyệt", "Kiểm tra ảnh, CSS, JS, API của trang"],
                     ["Trang tải thành công (HTTP < 400)", "Không có lỗi JavaScript", "Không có tài nguyên lỗi 4xx/5xx"]),
            "login": (["Mở trang đăng nhập", "Nhập tài khoản / mật khẩu và bấm Đăng nhập"],
                      ["Hiện form đăng nhập", "Đăng nhập thành công, vào trang chính"]),
            "links": ([open_, "Mở lần lượt các link trên trang"], ["", "Các link mở đúng trang, không lỗi 404/500"]),
            "buttons": ([open_, "Bấm lần lượt các nút thao tác (bỏ qua nút làm thay đổi dữ liệu)"],
                        ["", "Hệ thống phản hồi, không phát sinh lỗi"]),
            "search": ([open_, "Nhập từ khóa có trong danh sách (%s)" % data, "Bấm Tìm kiếm / Enter"],
                       ["Danh sách có dữ liệu", "", "Chỉ hiện các bản ghi chứa từ khóa"]),
            "paging": ([open_, "Bấm sang trang tiếp theo", "Quay lại trang trước"],
                       ["Hiển thị trang 1", "Hiển thị dữ liệu trang 2, khác trang 1", "Hiển thị lại dữ liệu trang 1"]),
            "sorting": ([open_, "Bấm tiêu đề cột '%s'" % thing, "Bấm tiêu đề cột '%s' lần nữa" % thing],
                        ["", "Dữ liệu sắp xếp tăng dần theo cột", "Dữ liệu sắp xếp giảm dần theo cột"]),
            "form_actions": ([open_, "Mở form '%s', nhập dữ liệu hợp lệ" % form, "Bấm '%s'" % thing],
                             ["", "", "Form đóng / không lưu" if "không lưu" in title else "Lưu thành công"]),
            "api": ([open_, "Theo dõi các API được gọi (tab Network)"], ["", "Các API trả về thành công, không lỗi"]),
            "dropdown": ([open_, "Mở từng dropdown"], ["", "Dropdown có dữ liệu lấy từ hệ thống, không trống"]),
            "loading": ([open_, "Chờ trang tải dữ liệu"], ["", "Biểu tượng đang tải biến mất trong thời gian cho phép"]),
            "mapping": ([open_, "So sánh dữ liệu trên danh sách với dữ liệu API trả về"],
                        ["", "Dữ liệu hiển thị khớp với dữ liệu trả về"]),
            "toggle": ([open_, "Bấm '%s' để mở" % thing, "Bấm '%s' lần nữa để đóng" % thing],
                       ["", "Nội dung mở ra", "Nội dung đóng lại"]),
            "switch": ([open_, "Bấm nút gạt '%s'" % thing, "Bấm lại nút gạt"],
                       ["", "Trạng thái nút gạt thay đổi", "Trạng thái trở lại như ban đầu"]),
            "popup": ([open_, "Bấm '%s' để mở popup" % thing, "Bấm nút Đóng / biểu tượng X"],
                      ["", "Popup hiển thị", "Popup đóng, trang trở về trạng thái ban đầu"]),
            "flow": ([open_, "Mở form, nhập dữ liệu hợp lệ", "Bấm '%s'" % thing, "Xem trạng thái bản ghi trên danh sách"],
                     ["", "", "Lưu thành công", title.split("→")[-1].strip()]),
            "ui": ([open_, "Quan sát giao diện: ảnh, chữ, bố cục, tiêu đề tab"],
                   ["", "Không có ảnh lỗi, chữ lỗi (undefined/null/NaN), không tràn ngang, có tiêu đề"]),
            "responsive": (["Mở trang %s ở kích thước %s" % (url, kw.get("size") or ""), "Quan sát bố cục, cuộn ngang"],
                           ["", "Không tràn ngang, không lỗi JavaScript"]),
            "ease": (["Mở trang %s%s" % (url, (" ở kích thước " + kw["size"]) if kw.get("size") else ""),
                      "Kiểm tra từng nút, ô nhập, nút gạt"], ["", "Không bị che khuất, đủ lớn để bấm, chữ không quá nhỏ"]),
            "perf": ([open_, "Đo thời gian tải và dung lượng trang"], ["", "Thời gian tải và dung lượng trong giới hạn"]),
            "a11y": ([open_, "Kiểm tra ảnh, ô nhập, nút, tiêu đề, độ tương phản"],
                     ["", "Ảnh có mô tả, ô nhập có nhãn, nút có tên, đủ tương phản"]),
            "compat": ([open_ + " trên %s" % BROWSER_NAME.get(kw.get("browser"), kw.get("browser") or ""),
                        "Quan sát trang"], ["", "Trang tải được, không lỗi JavaScript, không tràn"]),
            "perm": (["Đăng nhập bằng tài khoản vai trò thứ 2", title], ["Đăng nhập thành công",
                                                                         "Đúng quyền theo ma trận phân quyền"]),
            "sec": ([title], ["Đạt yêu cầu bảo mật cơ bản"]),
        }
        return T_.get(feat) or ([open_, title], ["", "Đạt yêu cầu"])

    def _kw(self, p, **extra):
        d = {"page": p}
        d.update(extra)
        return d

    # ---------- Functional
    def _plan_functional(self, p):
        out = []
        name = p.get("name") or p["url"]
        lo = name[:1].lower() + name[1:]
        F = lambda feat, title: self._S("functional", feat, name, title)  # noqa: E731
        scn = F("page", "Kiểm tra truy cập trang %s: mở được, hiển thị đầy đủ, không phát sinh lỗi JavaScript và tài nguyên" % name)

        def load(res):
            s1 = self._st(res, "Mở trang %s" % p["url"], "Trang tải thành công (HTTP < 400)")
            if p.get("error"):
                return self._bad(s1, p["error"], shot=False)
            if p["status"] >= 400:
                self._bad(s1, "Trang trả về lỗi HTTP %d" % p["status"], shot=False)
            else:
                self._ok(s1, "HTTP %d, tải trong %.1f giây" % (p["status"], ((p.get("perf") or {}).get("load") or 0)
                                                               / 1000.0))
            s2 = self._st(res, "Kiểm tra lỗi JavaScript", "Không có lỗi JavaScript")
            if p["console"]:
                msg = "Có %d lỗi JavaScript: %s" % (len(p["console"]), " | ".join(p["console"][:3]))
                if self.o("functional", "js_error") == "fail":
                    self._goto(p["url"])
                    self._bad(s2, msg)
                else:
                    self._ok(s2, "(chỉ ghi chú) " + msg)
            else:
                self._ok(s2, "Không có lỗi JavaScript")
            s3 = self._st(res, "Kiểm tra tài nguyên (ảnh, CSS, JS, API) tải được", "Không có tài nguyên lỗi 4xx/5xx")
            bad = [f for f in p["failed"] if f["type"] != "document" or f["url"] != p["url"]]
            if bad:
                self._bad(s3, "%d tài nguyên lỗi: %s" % (len(bad), "; ".join("%s (%s)" % (b["url"][-70:], b["status"] or
                                                                                          b["failed"]) for b in bad[:5])),
                          shot=False)
            else:
                self._ok(s3, "Tất cả tài nguyên tải được")
        out.append(("functional", "Trang '%s' mở được, không lỗi" % name, scn, self._kw(p, priority="Critical"),
                    load))
        if p.get("error"):
            return out
        info = p["info"]
        mode_links = self.o("functional", "links")
        links = [lk for lk in info.get("links", []) if lk["href"].startswith("http")]
        if mode_links != "none" and links:
            def links_fn(res, links=links):
                uniq = []
                for lk in links:
                    n = self._norm_url(lk["href"]) or lk["href"]
                    if n not in [u for u, _ in uniq] and (mode_links == "all" or self._internal(n)) and \
                            not self._skip(n, lk["text"]):
                        uniq.append((n, lk["text"]))
                s = self._st(res, "Kiểm tra %d link trên trang" % len(uniq), "Không có link hỏng (404/500)")
                bad = []
                known = {self._norm_url(x["url"]): x.get("status") for x in self.pages}
                for u, text in uniq[:150]:
                    self._check_stop()
                    st = known.get(u)
                    if st is None:
                        try:
                            r = self.ctx.request.get(u, timeout=10000, max_redirects=5)
                            st = r.status
                        except Exception as e:
                            st = str(e).splitlines()[0][:50]
                        known[u] = st
                    if not isinstance(st, int) or st >= 400:
                        bad.append("%s → %s" % (text or u[-60:], st))
                if bad:
                    self._bad(s, "%d/%d link lỗi: %s" % (len(bad), len(uniq), "; ".join(bad[:8])), shot=False)
                else:
                    self._ok(s, "Đã kiểm tra %d link, không có link hỏng" % len(uniq))
            out.append(("functional", "Link trên trang '%s' không bị hỏng" % name,
                        F("links", "Kiểm tra các liên kết, menu điều hướng trên trang %s mở đúng trang và không bị hỏng" % name), self._kw(p, priority="Medium"),
                        links_fn))
        if self.o("functional", "buttons") != "none" and (p.get("clicks") or p.get("skipped")):
            def btn_fn(res):
                for c in p.get("clicks", []):
                    s = self._st(res, "Bấm nút '%s'" % c["text"], "Không phát sinh lỗi, hệ thống phản hồi")
                    errs = c["js"] + c["http"]
                    if c["error"]:
                        self._goto(p["url"])
                        self._bad(s, "Không bấm được: %s" % c["error"], shot=False, status="Blocked")
                    elif errs:
                        self._bad(s, "Có lỗi sau khi bấm: %s" % " | ".join(errs[:3]), shot=False)
                    else:
                        what = []
                        if c["opened"]:
                            what.append("mở popup/form")
                        if c["nav"]:
                            what.append("chuyển tới %s" % c["nav"])
                        if c["newtab"]:
                            what.append("mở tab mới %s" % c["newtab"][0])
                        if c["dialogs"]:
                            what.append("hiện hộp thoại '%s' (tool bấm Hủy)" % c["dialogs"][0][:60])
                        if not what and c["changed"]:
                            what.append("nội dung trang thay đổi")
                        self._ok(s, "Không lỗi — " + (", ".join(what) if what else "không thấy thay đổi trên trang"))
                if p.get("skipped"):
                    s = self._st(res, "Các nút KHÔNG bấm thử (có thể thay đổi dữ liệu thật)",
                                 "Cần người test kiểm tra riêng")
                    s["status"], s["actual"] = "Skip", ", ".join(p["skipped"])
            out.append(("functional", "Các nút bấm trên trang '%s' hoạt động" % name,
                        F("buttons", "Kiểm tra các thao tác (nút bấm, icon) trên trang %s hoạt động và không phát sinh lỗi" % name),
                        self._kw(p, priority="High"), btn_fn))
        if self.o("functional", "search") == "yes":
            for sv in info.get("search", [])[:1]:
                out.append(("functional", "Tìm kiếm theo từ khóa trên trang '%s'" % name,
                            F("search", "Kiểm tra tìm kiếm %s theo từ khóa và kết quả hiển thị đúng" % lo),
                            self._kw(p, data="Ô: %s" % sv["label"]), lambda res, sv=sv: self._t_search(res, p, sv)))
        tables = info.get("tables", [])
        if self.o("functional", "paging") == "yes" and info.get("pager") and tables:
            out.append(("functional", "Phân trang danh sách trên trang '%s'" % name,
                        F("paging", "Kiểm tra phân trang danh sách %s: chuyển trang, số trang và dữ liệu từng trang" % lo), self._kw(p, priority="Medium"),
                        lambda res: self._t_paging(res, p)))
        if self.o("functional", "sorting") == "yes":
            for t in tables[:1]:
                for h in [h for h in t["headers"] if h["sortable"] and h["text"]][:4]:
                    out.append(("functional", "Sắp xếp theo cột '%s'" % h["text"],
                                F("sorting", "Kiểm tra sắp xếp tăng/giảm theo các cột hỗ trợ sắp xếp trên trang %s" % name),
                                self._kw(p, priority="Medium"),
                                lambda res, t=t, h=h: self._t_sort(res, p, t, h)))
        if self.o("functional", "form_actions") == "yes":
            out += self._plan_form_actions(p)
        return out

    def _rows(self, row_sel, idx=None):
        try:
            return self.page.evaluate(ROWS_JS, {"rowSel": row_sel, "idx": idx})
        except Exception:
            return []

    def _t_search(self, res, p, sv):
        t = (p["info"].get("tables") or [None])[0]
        kw = ""
        if t and t["first"]:
            kw = next((c for c in t["first"] if 3 <= len(c) <= 40 and not re.match(r"^[\d\s.,:/-]+$", c)), "") or \
                next((c for c in t["first"] if 2 <= len(c) <= 40), "")
        c0 = len(self.console_errors)

        def do(term):
            self._goto(p["url"])
            box = self.page.locator(sv["sel"]).first
            box.fill(term)
            if sv.get("button"):
                self.page.locator(sv["button"]["sel"]).first.click()
            else:
                box.press("Enter")
            self._settle(600, 2500)
        s0 = self._st(res, "Mở trang %s" % p["url"], "Trang hiển thị ô tìm kiếm '%s'" % sv["label"])
        self._ok(s0, "Có ô tìm kiếm")
        if kw:
            s1 = self._st(res, "Nhập từ khóa có trong danh sách và tìm", "Chỉ hiện các dòng chứa từ khóa", kw)
            do(kw)
            rows = self._rows(t["rowSel"])
            if not rows:
                self._bad(s1, "Không còn dòng nào dù từ khóa '%s' có trong danh sách" % kw)
            else:
                wrong = [r for r in rows if kw.lower() not in r.lower()]
                if wrong:
                    self._bad(s1, "%d/%d dòng không chứa '%s' (VD: '%s')" % (len(wrong), len(rows), kw, wrong[0][:80]))
                else:
                    self._ok(s1, "Còn %d dòng, tất cả chứa '%s'" % (len(rows), kw))
        s2 = self._st(res, "Nhập từ khóa không tồn tại và tìm", "Không có kết quả / hiện thông báo không có dữ liệu",
                      "zz_khong_ton_tai_%s" % self.rnd)
        do("zz_khong_ton_tai_%s" % self.rnd)
        if t:
            rows = [r for r in self._rows(t["rowSel"]) if r]
            body = self.page.inner_text("body")
            empty_msg = re.search(r"không có|không tìm thấy|chưa có dữ liệu|no data|no result|not found|0 kết quả",
                                  body, re.I)
            if not rows or (len(rows) == 1 and empty_msg):
                self._ok(s2, "Không có kết quả" + (" — hiện '%s'" % empty_msg.group(0) if empty_msg else ""))
            else:
                self._bad(s2, "Vẫn hiện %d dòng với từ khóa không tồn tại" % len(rows))
        else:
            self._ok(s2, "Tìm kiếm chạy được (trang không có bảng để đối chiếu)")
        s3 = self._st(res, "Kiểm tra lỗi khi tìm kiếm", "Không phát sinh lỗi JavaScript")
        errs = self.console_errors[c0:]
        if errs:
            self._bad(s3, "Lỗi JavaScript: %s" % " | ".join(errs[:3]), shot=False)
        else:
            self._ok(s3, "Không có lỗi")

    def _t_paging(self, res, p):
        t = p["info"]["tables"][0]
        pager = p["info"]["pager"]
        target = next((x for x in pager if x["text"] == "2"), None) or next(
            (x for x in pager if re.search(r"tiếp|sau|next|›|»|>", x["text"] + x["aria"], re.I)), None)
        s1 = self._st(res, "Mở trang, ghi nhận dữ liệu trang 1", "Danh sách có dữ liệu")
        self._goto(p["url"])
        r1 = self._rows(t["rowSel"])
        if not r1:
            return self._bad(s1, "Danh sách trống")
        self._ok(s1, "Trang 1 có %d dòng, dòng đầu: %s" % (len(r1), r1[0][:60]))
        s2 = self._st(res, "Bấm '%s' để sang trang tiếp" % (target["text"] or target["aria"]) if target else
                      "Bấm sang trang 2", "Hiện dữ liệu trang 2, khác trang 1")
        if not target:
            return self._bad(s2, "Không tìm thấy nút sang trang 2 / Tiếp", status="Blocked")
        c0 = len(self.console_errors)
        self.page.locator(target["sel"]).first.click()
        self._settle(500, 2500)
        r2 = self._rows(t["rowSel"])
        if not r2:
            self._bad(s2, "Trang 2 không có dòng nào")
        elif r2 == r1:
            self._bad(s2, "Dữ liệu trang 2 giống hệt trang 1")
        else:
            self._ok(s2, "Trang 2 có %d dòng, dòng đầu: %s" % (len(r2), r2[0][:60]))
        s3 = self._st(res, "Kiểm tra lỗi khi phân trang", "Không phát sinh lỗi JavaScript")
        errs = self.console_errors[c0:]
        self._bad(s3, " | ".join(errs[:3]), shot=False) if errs else self._ok(s3, "Không có lỗi")

    @staticmethod
    def _sort_key(v):
        s = v.strip()
        m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
        if m:
            return (0, int(m.group(3)) * 10000 + int(m.group(2)) * 100 + int(m.group(1)), "")
        t = s.replace(".", "").replace(",", ".").replace(" ", "")
        if re.match(r"^-?\d+(\.\d+)?$", t):
            return (0, float(t), "")
        return (1, 0, C.norm(s))

    def _t_sort(self, res, p, t, h):
        self._goto(p["url"])
        before = self._rows(t["rowSel"], h["idx"])
        verdicts = []
        for n in (1, 2):
            s = self._st(res, "Bấm tiêu đề cột '%s' (lần %d)" % (h["text"], n), "Dữ liệu cột được sắp xếp %s" % (
                "tăng hoặc giảm dần" if n == 1 else "theo chiều ngược lại"))
            self.page.locator(h["sel"]).first.click()
            self._settle(400, 2000)
            vals = [v for v in self._rows(t["rowSel"], h["idx"]) if v]
            asc = desc = False
            for kf in (self._sort_key, lambda v: v.casefold(), lambda v: v):   # bỏ dấu / chữ thường / nguyên gốc
                keys = [kf(v) for v in vals]
                asc = asc or keys == sorted(keys)
                desc = desc or keys == sorted(keys, reverse=True)
            if len(vals) < 2:
                self._bad(s, "Không đủ dữ liệu để kiểm tra sắp xếp", status="Blocked")
                continue
            if vals == before and n == 1 and not (asc or desc):
                self._bad(s, "Bấm tiêu đề nhưng thứ tự không đổi và chưa được sắp xếp")
            elif asc or desc:
                verdicts.append("tăng" if asc else "giảm")
                if n == 2 and len(verdicts) == 2 and verdicts[0] == verdicts[1] and len(set(vals)) > 1:
                    self._bad(s, "Bấm lần 2 vẫn sắp xếp %s dần (không đảo chiều)" % verdicts[1])
                else:
                    self._ok(s, "Đã sắp xếp %s dần: %s…" % (verdicts[-1], ", ".join(vals[:4])))
            else:
                self._bad(s, "Cột chưa được sắp xếp: %s…" % ", ".join(vals[:5]))

    @staticmethod
    def _api_objects(bodies):
        found = []

        def walk(x, d=0):
            if d > 3 or len(found) >= 1:
                return
            if isinstance(x, list) and x and all(isinstance(i, dict) for i in x[:5]):
                found.append(x[:5])
                return
            if isinstance(x, dict):
                for v in x.values():
                    walk(v, d + 1)
        for b in bodies:
            walk(b)
        return found[0] if found else None

    def _t_api(self, res, p, objs):
        s1 = self._st(res, "Mở trang và lấy dữ liệu API trả về", "API trả về danh sách dữ liệu")
        self._goto(p["url"])
        self._ok(s1, "API trả về %d+ bản ghi" % len(objs))
        s2 = self._st(res, "Đối chiếu giá trị trong %d bản ghi đầu với nội dung hiển thị" % len(objs),
                      "Mỗi bản ghi đều hiển thị trên trang")
        body = re.sub(r"\s+", " ", self.page.inner_text("body"))
        shown, missing = 0, []
        for o in objs:
            vals = [str(v) for k, v in o.items() if isinstance(v, (str, int, float)) and not isinstance(v, bool)
                    and 2 <= len(str(v)) <= 80 and not re.match(r"^(https?:|\d{4}-\d\d-\d\dT)", str(v))]
            if not vals:
                continue
            hit = [v for v in vals if v in body]
            if len(hit) >= max(1, len(vals) // 2):
                shown += 1
            else:
                missing.append(", ".join(vals[:3]))
        if not shown and not missing:
            return self._ok(s2, "Không có giá trị chữ/số để đối chiếu")
        if missing:
            self._bad(s2, "%d/%d bản ghi không thấy trên trang (VD: %s)" % (len(missing), shown + len(missing),
                                                                         missing[0][:100]))
        else:
            self._ok(s2, "Cả %d bản ghi đều hiển thị đúng" % shown)

    # ---------- Validation (form)
    def _rows_valid(self, f):
        return f.get("type") not in ("hidden", "submit", "button", "reset", "image") and not f.get("disabled")

    def _valid_value(self, f):
        t = f.get("type") or "text"
        lab = C.norm("%s %s %s" % (f.get("label", ""), f.get("name", ""), f.get("placeholder", "")))
        today = TODAY()
        if f.get("tag") == "select":
            opts = [o for o in f.get("options", []) if o["value"] and not re.search(r"chọn|select|--|tất cả", o["text"],
                                                                                   re.I)]
            return opts[0]["text"] if opts else None
        if t in ("checkbox", "radio"):
            return True
        if t == "file":
            return None
        if t == "email" or "email" in lab or "thu dien tu" in lab:
            return "test%s@example.com" % self.rnd
        if t == "tel" or any(k in lab for k in ("dien thoai", "sdt", "phone", "mobile", "so dt", "di dong")):
            return "0901234567"
        if t in ("number", "range"):
            lo, hi = f.get("min"), f.get("max")
            v = lo if lo is not None else (1 if hi is None or hi >= 1 else hi)
            return str(int(v)) if float(v).is_integer() else str(v)
        if t == "date":
            return (today + dt.timedelta(days=7)).isoformat()
        if t == "datetime-local":
            return (today + dt.timedelta(days=7)).isoformat() + "T09:00"
        if t == "time":
            return "09:00"
        if t == "month":
            return today.strftime("%Y-%m")
        if t == "url" or "website" in lab or re.search(r"\burl\b", lab):
            return "https://example.com"
        if t == "password":
            return "Test@12345"
        if t == "color":
            return "#336699"
        if re.search(r"\b(ho ten|ten|name|fullname)\b", lab):
            s = "Nguyễn Văn Test"
        elif re.search(r"\b(ma|code|ma so)\b", lab):
            s = "AUTO%s" % self.rnd
        elif "dia chi" in lab or "address" in lab:
            s = "123 Đường Kiểm Thử"
        elif re.search(r"\b(ngay|date)\b", lab):
            s = (today + dt.timedelta(days=7)).strftime("%d/%m/%Y")
        elif re.search(r"\b(so luong|quantity|so)\b", lab):
            s = "1"
        else:
            s = "Kiểm thử tự động %s" % self.rnd
        if f.get("minlength") and len(s) < f["minlength"]:
            s = (s + "x" * f["minlength"])[:max(f["minlength"], len(s))]
        if f.get("maxlength"):
            s = s[:f["maxlength"]]
        return s

    def _open_form(self, p, form):
        self._goto(p["url"])
        if not form.get("visible") and form.get("opener"):
            self.page.locator(form["opener"]["sel"]).first.click()
            self._pump(400)
        first = next((f for f in form["fields"] if self._rows_valid(f) and f["type"] != "file"), None)
        if first:
            self.page.locator(first["sel"]).first.wait_for(state="visible", timeout=5000)

    def _fill(self, form, values):
        """values: sel -> giá trị (None = để trống, KEEP = giá trị hợp lệ mặc định). Trả về (mô tả, đặc biệt)."""
        desc, special = [], {}
        for f in form["fields"]:
            if not self._rows_valid(f):
                continue
            v = values[f["sel"]] if f["sel"] in values else self._valid_value(f)
            loc = self.page.locator(f["sel"]).first
            try:
                if f["type"] == "file":
                    if v:
                        loc.set_input_files(v)
                        desc.append("%s = tệp %s" % (f["label"] or f["name"], os.path.basename(v)))
                    continue
                if not loc.is_visible():
                    continue
                if f["tag"] == "select":
                    if v in (None, ""):
                        try:
                            loc.select_option(value="")
                        except Exception:
                            loc.select_option(index=0)
                        desc.append("%s = (không chọn)" % (f["label"] or f["name"]))
                    else:
                        try:
                            loc.select_option(label=v)
                        except Exception:
                            loc.select_option(value=v)
                        desc.append("%s = %s" % (f["label"] or f["name"], v))
                elif f["type"] in ("checkbox", "radio"):
                    loc.set_checked(bool(v))
                    desc.append("%s = %s" % (f["label"] or f["name"], "tick" if v else "không tick"))
                else:
                    loc.fill("" if v is None else str(v))
                    shown = "(để trống)" if not v else (str(v) if len(str(v)) <= 40 else str(v)[:20] + "…(%d ký tự)"
                                                        % len(str(v)))
                    desc.append("%s = %s" % (f["label"] or f["name"], shown))
                    if v:
                        try:
                            special[f["sel"]] = loc.input_value()
                        except Exception:
                            pass
            except Exception as e:
                msg = str(e)
                if "Cannot type text" in msg or "Malformed value" in msg:
                    special[f["sel"]] = "__not_typable__"
                    desc.append("%s = %s (ô không cho nhập)" % (f["label"] or f["name"], v))
                else:
                    desc.append("%s: không nhập được (%s)" % (f["label"] or f["name"], msg.splitlines()[0][:60]))
        return "; ".join(desc), special

    def _submit_and_judge(self, form, form_was_visible=True, button=None):
        self.page.evaluate(FORM_PREP_JS, form["sel"])
        url0 = self.page.url
        d0 = len(self.dialogs)
        c0 = len(self.console_errors)
        if button:
            self.page.locator(button["sel"]).first.click(timeout=5000)
        elif form.get("submit"):
            self.page.locator(form["submit"]["sel"]).first.click(timeout=5000)
        else:
            last = [f for f in form["fields"] if self._rows_valid(f) and f["tag"] == "input"]
            self.page.locator(last[-1]["sel"] if last else form["fields"][0]["sel"]).first.press("Enter")
        self._settle(700, 3000)
        st = self.page.evaluate(FORM_STATE_JS, {"fsel": form["sel"], "fields": [f["sel"] for f in form["fields"]]})
        dlg = self.dialogs[d0:]
        errs = st["errTexts"] + [d for d in dlg if ERR_PY.search(d) and not OK_PY.search(d)]
        oks = st["okTexts"] + [d for d in dlg if OK_PY.search(d)]
        blocked = bool(st["invalid"]) and not st["submitted"] and form.get("isForm")
        nav = self._norm_url(st["url"]) != self._norm_url(url0)
        closed = form_was_visible and not st["formVisible"]
        rejected = bool(errs) or blocked
        accepted = bool(oks) or nav or (closed and not errs)
        detail_r = ("Hệ thống báo lỗi: '%s'" % "' | '".join(errs[:3])) if errs else (
            "Trình duyệt chặn gửi form: %s" % "; ".join("%s — %s" % (x["label"], x["msg"]) for x in st["invalid"][:3])
            if blocked else "")
        detail_a = ("Hệ thống CHẤP NHẬN: '%s'" % "' | '".join(oks[:2])) if oks else (
            "Trang chuyển sang %s" % st["url"] if nav else ("Form đã đóng (đã được lưu?)" if closed else ""))
        js = self.console_errors[c0:]
        if rejected and not accepted:
            return "rejected", detail_r, errs + [x["msg"] for x in st["invalid"]], js
        if accepted and not rejected:
            return "accepted", detail_a, oks, js
        return "unknown", ("Không thấy báo lỗi hay thông báo thành công rõ ràng%s" % (
            " (%s; %s)" % (detail_r, detail_a) if (detail_r and detail_a) else "")), errs, js

    def _form_case(self, res, p, form, values, expect, exp_msg="", note_data="", pre_check=None, meta=None):
        """Một ca kiểm tra form: mở form → nhập → bấm gửi → so với mong đợi (reject/accept)."""
        name = form.get("name") or (form.get("submit") or {}).get("text") or "form"
        s1 = self._st(res, "Mở trang %s%s" % (p["path"], " và bấm '%s'" % form["opener"]["text"] if
                                               form.get("opener") and not form.get("visible") else ""),
                      "Hiện form '%s'" % name)
        self._open_form(p, form)
        self._ok(s1, "Đã mở form")
        s2 = self._st(res, "Nhập dữ liệu", "", note_data)
        desc, special = self._fill(form, values)
        s2["fields"]["du_lieu_buoc"] = note_data or desc
        self._ok(s2, desc or "Đã nhập")
        m = dict(meta or {})
        m.update({"screen": p.get("name"), "form": name, "expect": "Từ chối" if expect == "reject" else "Chấp nhận",
                  "exp_msg": exp_msg})
        if pre_check:
            r = pre_check(special)
            if r:
                s3 = self._st(res, r[0], r[1])
                s3["m"] = dict(m, act_msg=r[3])
                (self._ok if r[2] else self._bad)(s3, r[3])
                if r[2] is not None:
                    return
        want = "Hệ thống báo lỗi, KHÔNG cho lưu" + (" — thông báo '%s'" % exp_msg if exp_msg else "") if \
            expect == "reject" else "Lưu / gửi thành công"
        s3 = self._st(res, "Bấm '%s'" % ((form.get("submit") or {}).get("text") or "Enter"), want)
        s3["m"] = m
        if self.o("validation", "submit") == "no" and res["type"] == "validation":
            s3["status"], s3["actual"] = "Skip", "Không bấm gửi (tùy chọn: chỉ kiểm tra thuộc tính ô)"
            return
        verdict, detail, msgs, js = self._submit_and_judge(form)
        m["act_msg"] = detail
        if verdict == "unknown":
            return self._unsure(res, s3, detail, res["type"])
        if expect == "reject":
            if verdict == "rejected":
                if exp_msg and not any(C.norm(exp_msg) in C.norm(m) for m in msgs):
                    self._bad(s3, "Có báo lỗi nhưng khác mong đợi: %s" % detail)
                else:
                    self._ok(s3, detail)
            else:
                self._bad(s3, "Dữ liệu KHÔNG hợp lệ nhưng %s" % detail[0].lower() + detail[1:])
        else:
            if verdict == "accepted":
                self._ok(s3, detail)
            else:
                self._bad(s3, "Dữ liệu hợp lệ nhưng bị từ chối — %s" % detail)
        if js:
            s4 = self._st(res, "Kiểm tra lỗi JavaScript", "Không có lỗi")
            self._bad(s4, " | ".join(js[:3]), shot=False)

    def _unsure(self, res, s, detail, tkey):
        self._shot(s, "duyet")
        if self.o(tkey if tkey in ("validation", "business") else "validation", "unsure") == "ask":
            res["mode"] = res["fields"]["cach_test"] = "Bán tự động"
            self._ask_manual(self.inst, res, [s], reason="Tool không chắc kết quả: %s" % detail)
        else:
            s["status"], s["actual"] = "Chưa test", "Tool không chắc — cần xem ảnh: %s" % detail
            res["mode"] = res["fields"]["cach_test"] = "Bán tự động"

    def _data_forms(self, p):
        return [f for f in (p.get("info") or {}).get("forms", []) if f["kind"] == "data" and
                (f.get("visible") or f.get("opener")) and not (self.logged_in and f.get("hasPassword") and
                                                                len(f["fields"]) <= 3)]

    def _plan_validation(self, p):
        out = []
        cases = set(self.o("validation", "cases") or [])
        limit = int(self.o("validation", "max_cases") or 20)
        for form in self._data_forms(p):
            fname = form.get("name") or (form.get("submit") or {}).get("text") or "form"
            sub = (form.get("submit") or {}).get("text") or "Lưu"
            VT = {"required": "Kiểm tra thông báo lỗi khi bỏ trống các trường bắt buộc (*) và bấm %s trên form %s"
                              % (sub, fname),
                  "format": "Kiểm tra định dạng dữ liệu nhập (email, số điện thoại) trên form %s" % fname,
                  "length": "Kiểm tra độ dài, giá trị giới hạn (min / max) của các trường trên form %s" % fname,
                  "file": "Kiểm tra nhập dữ liệu khi chọn tệp sai định dạng trên form %s" % fname,
                  "valid": "Kiểm tra lưu form %s với dữ liệu hợp lệ" % fname}
            fields = [f for f in form["fields"] if self._rows_valid(f)]
            items = []
            req = [f for f in fields if f["required"]]
            pre = self._pre() + "; mở form '%s'" % fname

            def add(title, values, expect, data="", check=None, prio="High", meta=None, feat="required"):
                scn = self._S("validation", feat, p.get("name"), VT[feat])
                items.append(("validation", title, scn, self._kw(p, pre=pre, data=data, priority=prio),
                              lambda res, v=values, e=expect, d=data, c=check, f=form, mt=meta:
                              self._form_case(res, p, f, v, e, note_data=d, pre_check=c, meta=mt)))

            def mt(f, rule, inp):
                return {"field": f["label"] or f["name"] if f else "(tất cả các ô)", "ftype": f["type"] if f else "",
                        "required": ("Có" if f["required"] else "Không") if f else "", "rule": rule, "input": inp}
            if "empty_all" in cases and fields:
                add("Gửi form '%s' khi để trống tất cả" % fname, {f["sel"]: None for f in fields}, "reject",
                    "Tất cả các ô để trống", meta=mt(None, "Gửi form trống", "(để trống)"))
            if "required" in cases:
                for f in req:
                    add("Bỏ trống trường bắt buộc '%s'" % f["label"], {f["sel"]: None}, "reject",
                        "%s = (để trống), các ô khác hợp lệ" % f["label"], meta=mt(f, "Bắt buộc nhập", "(để trống)"))
            if "spaces" in cases:
                for f in req:
                    if f["tag"] in ("input", "textarea") and f["type"] in ("text", "textarea", "email", "search",
                                                                           "tel"):
                        add("Trường bắt buộc '%s' chỉ nhập khoảng trắng" % f["label"], {f["sel"]: "   "}, "reject",
                            "%s = '   '" % f["label"], prio="Medium", meta=mt(f, "Bắt buộc (không nhận khoảng trắng)",
                                                                           "'   '"), feat="required")
            for f in fields:
                lab = C.norm(f["label"] + " " + f["name"])
                if "email" in cases and (f["type"] == "email" or "email" in lab):
                    for bad in ("abc", "abc@", "abc@xyz"):
                        add("Email '%s' sai định dạng: %s" % (f["label"], bad), {f["sel"]: bad}, "reject",
                            "%s = %s" % (f["label"], bad), meta=mt(f, "Định dạng email", bad), feat="format")
                if "phone" in cases and (f["type"] == "tel" or re.search(r"dien thoai|sdt|phone|di dong|mobile",
                                                                          lab)):
                    for bad in ("12345", "abcdefghij"):
                        add("Số điện thoại '%s' sai định dạng: %s" % (f["label"], bad), {f["sel"]: bad}, "reject",
                            "%s = %s" % (f["label"], bad), meta=mt(f, "Định dạng số điện thoại", bad), feat="format")
                if "maxlength" in cases and f.get("maxlength") and f["tag"] in ("input", "textarea"):
                    n = f["maxlength"]

                    def chk(special, f=f, n=n):
                        got = special.get(f["sel"])
                        if isinstance(got, str) and got != "__not_typable__" and len(got) <= n:
                            return ("Kiểm tra số ký tự trong ô '%s'" % f["label"], "Ô chỉ nhận tối đa %d ký tự" % n,
                                    True, "Ô tự cắt còn %d ký tự (giới hạn %d)" % (len(got), n))
                        return None
                    add("Trường '%s' nhập %d ký tự (vượt tối đa %d)" % (f["label"], n + 1, n), {f["sel"]: "a" * (n + 1)},
                        "reject", "%s = %d ký tự" % (f["label"], n + 1), check=chk, prio="Medium",
                        meta=mt(f, "Độ dài tối đa %d ký tự" % n, "%d ký tự" % (n + 1)), feat="length")
                if "number_range" in cases and f["type"] == "number":
                    if f.get("min") is not None:
                        v = f["min"] - 1
                        add("Trường '%s' nhập %g (dưới tối thiểu %g)" % (f["label"], v, f["min"]), {f["sel"]: "%g" % v},
                            "reject", "%s = %g" % (f["label"], v), meta=mt(f, "Giá trị nhỏ nhất %g" % f["min"], "%g" % v),
                            feat="length")
                    if f.get("max") is not None:
                        v = f["max"] + 1
                        add("Trường '%s' nhập %g (vượt tối đa %g)" % (f["label"], v, f["max"]), {f["sel"]: "%g" % v},
                            "reject", "%s = %g" % (f["label"], v), meta=mt(f, "Giá trị lớn nhất %g" % f["max"], "%g" % v),
                            feat="length")
                if "file_type" in cases and f["type"] == "file":
                    path = self._bad_file(f.get("accept", ""))
                    add("Tải tệp sai định dạng vào '%s'" % (f["label"] or "tệp"), {f["sel"]: path}, "reject",
                        "Tệp %s (cho phép: %s)" % (os.path.basename(path), f.get("accept") or "không ghi"),
                        meta=mt(f, "Định dạng tệp (%s)" % (f.get("accept") or "?"), os.path.basename(path)), feat="file")
            if "valid" in cases and fields:
                add("Nhập đầy đủ dữ liệu hợp lệ và gửi form '%s'" % fname, {}, "accept",
                    "Dữ liệu hợp lệ do tool tạo", prio="High", meta=mt(None, "Dữ liệu hợp lệ", "(hợp lệ)"), feat="valid")
            out += items[:limit]
        if "search" in cases or True:
            for sv in (p.get("info") or {}).get("search", [])[:1]:
                scn = self._S("validation", "search_input", p.get("name"),
                              "Kiểm tra tìm kiếm với từ khóa chỉ khoảng trắng, ký tự đặc biệt, mã script, chuỗi quá dài trên "
                              "trang %s" % p.get("name"))
                for desc, val in (("ký tự đặc biệt", "<script>alert(1)</script>"), ("chuỗi rất dài (300 ký tự)", "a" * 300),
                                  ("chỉ khoảng trắng", "   "), ("dấu nháy và ký tự SQL", "' OR 1=1 --")):
                    out.append(("validation", "Tìm kiếm với %s" % desc, scn, self._kw(p, data="%s" % val[:60],
                                                                                   priority="Medium"),
                                lambda res, sv=sv, val=val, desc=desc: self._t_search_input(res, p, sv, val, desc)))
        return out

    def _t_search_input(self, res, p, sv, val, desc):
        s1 = self._st(res, "Mở trang %s" % p["path"], "Có ô tìm kiếm '%s'" % sv["label"])
        self._goto(p["url"])
        self._ok(s1, "Đã mở")
        c0, n0 = len(self.console_errors), len(self.resp_log)
        s2 = self._st(res, "Nhập %s vào ô tìm kiếm và tìm" % desc, "Hệ thống không lỗi, không hiện lỗi kỹ thuật, "
                      "trang vẫn dùng được", val[:60])
        box = self.page.locator(sv["sel"]).first
        box.fill(val)
        if sv.get("button"):
            self.page.locator(sv["button"]["sel"]).first.click()
        else:
            box.press("Enter")
        self._settle(600, 2500)
        s2["m"] = {"screen": p.get("name"), "form": "Tìm kiếm", "field": sv["label"], "ftype": "search", "required": "",
                   "rule": "Dữ liệu đầu vào tìm kiếm", "input": desc, "expect": "Không lỗi"}
        errs = self.console_errors[c0:] + ["%s (%d)" % (r["url"][-60:], r["status"]) for r in self.resp_log[n0:]
                                          if r["status"] >= 500]
        body = self.page.inner_text("body")
        tech = re.search(r"Traceback|Exception|SQLSTATE|SQL syntax|ORA-\d{5}|Stack trace|Fatal error", body)
        dialog_xss = any("1" == d.strip() for d in self.dialogs[-3:])
        if dialog_xss:
            s2["m"]["act_msg"] = "Chạy được mã script"
            self._bad(s2, "Trang CHẠY đoạn script người dùng nhập (lỗi XSS)")
        elif tech:
            s2["m"]["act_msg"] = "Lộ lỗi kỹ thuật"
            self._bad(s2, "Trang hiện lỗi kỹ thuật: %s" % tech.group(0))
        elif errs:
            s2["m"]["act_msg"] = "Có lỗi"
            self._bad(s2, "Có lỗi: %s" % " | ".join(errs[:3]))
        else:
            s2["m"]["act_msg"] = "Không lỗi"
            self._ok(s2, "Không lỗi, trang vẫn hoạt động")

    def _bad_file(self, accept):
        acc = accept.lower()
        ext = ".txt" if ".txt" not in acc and "text/" not in acc else ".abc"
        path = os.path.join(self.run_dir, "tep_sai_dinh_dang" + ext)
        if not os.path.exists(path):
            with open(path, "w", encoding="utf-8") as f:
                f.write("Tệp kiểm thử sai định dạng do Tool Test tạo.\n")
        return path

    # ---------- Business rule
    def _find_field(self, rule):
        pu = self._norm_url(rule.get("page_url") or self.url)
        page = next((p for p in self.pages if self._norm_url(p["url"]) == pu), None)
        if page is None:
            page = self._visit(pu)
            page["depth"] = 0
            if page.get("info"):
                self._probe(page)
            self.pages.append(page)
        for form in (page.get("info") or {}).get("forms", []):
            for f in form["fields"]:
                if f["sel"] == rule.get("field_sel") or (rule.get("field_label") and C.norm(f["label"]) == C.norm(
                        rule["field_label"])):
                    return page, form, f
        return page, None, None

    def _date_val(self, f, d):
        return d.isoformat() if f["type"] == "date" else d.strftime("%d/%m/%Y")

    def _plan_business(self, page=None):
        out = []
        valid = self.o("business", "valid_cases") == "yes"
        for rule in self.rules:
            if not rule.get("rule"):
                continue
            label = rule.get("field_label") or rule.get("field_sel") or "?"
            rname = T.RULE_NAME.get(rule["rule"], rule["rule"])
            title = "%s: %s%s" % (label, rname, (" (%s%s)" % (rule.get("v1", ""), " – " + rule["v2"] if rule.get("v2")
                                                             else "")) if rule.get("v1") else "")
            if page is not None and self._norm_url(rule.get("page_url") or self.url) != self._norm_url(page["url"]):
                continue
            if page is None and any(self._norm_url(x["url"]) == self._norm_url(rule.get("page_url") or self.url)
                                    for x in self.pages):
                continue
            scn = self._S("business", "rule", rule.get("page_title") or "", "Kiểm tra quy tắc %s" % title,
                          ac=rule.get("ac", ""))
            v1, v2 = (rule.get("v1") or "").strip(), (rule.get("v2") or "").strip()
            msg = (rule.get("msg") or "").strip()
            r = rule["rule"]
            today = TODAY()
            cases = []   # (mô tả, giá trị hoặc hàm(f), mong đợi)

            def num(x):
                try:
                    return float(x.replace(",", "."))
                except Exception:
                    return None
            if r == "required":
                cases.append(("Bỏ trống '%s'" % label, None, "reject"))
                if valid:
                    cases.append(("Nhập '%s' hợp lệ" % label, "__valid__", "accept"))
            elif r in ("maxlen", "minlen"):
                n = int(num(v1) or 0)
                if n > 0:
                    if r == "maxlen":
                        cases.append(("'%s' nhập %d ký tự (vượt tối đa %d)" % (label, n + 1, n), "a" * (n + 1),
                                      "reject"))
                        if valid:
                            cases.append(("'%s' nhập đúng %d ký tự (biên trong)" % (label, n), "a" * n, "accept"))
                    else:
                        cases.append(("'%s' nhập %d ký tự (dưới tối thiểu %d)" % (label, n - 1, n), "a" * max(0, n - 1),
                                      "reject"))
                        if valid:
                            cases.append(("'%s' nhập đúng %d ký tự (biên trong)" % (label, n), "a" * n, "accept"))
            elif r == "range":
                a, b = num(v1), num(v2)
                if a is not None:
                    cases.append(("'%s' = %g (dưới %g)" % (label, a - 1, a), "%g" % (a - 1), "reject"))
                if b is not None:
                    cases.append(("'%s' = %g (vượt %g)" % (label, b + 1, b), "%g" % (b + 1), "reject"))
                if valid and a is not None:
                    cases.append(("'%s' = %g (biên dưới)" % (label, a), "%g" % a, "accept"))
                if valid and b is not None:
                    cases.append(("'%s' = %g (biên trên)" % (label, b), "%g" % b, "accept"))
            elif r == "email":
                cases += [("'%s' = %s" % (label, x), x, "reject") for x in ("abc", "abc@", "abc@xyz", "@xyz.com")]
                if valid:
                    cases.append(("'%s' = email hợp lệ" % label, "test%s@example.com" % self.rnd, "accept"))
            elif r == "phone":
                cases += [("'%s' = %s" % (label, x), x, "reject") for x in ("12345", "abcdefghij", "1234567890",
                                                                             "090123456789")]
                if valid:
                    cases.append(("'%s' = 0901234567" % label, "0901234567", "accept"))
            elif r == "digits":
                cases += [("'%s' = %s" % (label, x), x, "reject") for x in ("abc", "12a3", "1.5")]
            elif r == "nospecial":
                cases += [("'%s' = %s" % (label, x), x, "reject") for x in ("<script>alert(1)</script>", "@#$%^&*",
                                                                             "'; DROP TABLE x;--")]
            elif r == "no_past":
                cases.append(("'%s' = hôm qua" % label, lambda f: self._date_val(f, today - dt.timedelta(days=1)),
                              "reject"))
                if valid:
                    cases.append(("'%s' = ngày mai" % label, lambda f: self._date_val(f, today + dt.timedelta(days=1)),
                                  "accept"))
            elif r == "no_future":
                cases.append(("'%s' = ngày mai" % label, lambda f: self._date_val(f, today + dt.timedelta(days=1)),
                              "reject"))
                if valid:
                    cases.append(("'%s' = hôm nay" % label, lambda f: self._date_val(f, today), "accept"))
            elif r == "after_field":
                cases.append(("'%s' TRƯỚC '%s'" % (label, v1), ("__after__", v1, -5), "reject"))
                if valid:
                    cases.append(("'%s' SAU '%s'" % (label, v1), ("__after__", v1, 5), "accept"))
            elif r in ("reject", "accept"):
                for x in [x.strip() for x in v1.split(";") if x.strip()]:
                    cases.append(("'%s' = %s" % (label, x), x, r))
            elif r == "calc":
                cases.append(("Tính '%s' khi nhập %s" % (label, v1), ("__calc__", v1, v2), "calc"))
            for desc, val, expect in cases:
                out.append(("business", desc, scn, {"module": rule.get("page_title") or "", "pre": self._pre(),
                                                    "data": desc, "priority": "High"},
                            lambda res, rule=rule, val=val, expect=expect, msg=msg: self._t_rule(res, rule, val,
                                                                                               expect, msg)))
        return out

    def _t_rule(self, res, rule, val, expect, msg):
        page, form, f = self._find_field(rule)
        res["fields"]["man_hinh"] = page.get("path", "")
        rname = T.RULE_NAME.get(rule["rule"], rule["rule"])
        meta = {"ac": rule.get("ac", ""), "field": rule.get("field_label") or "", "rule": "%s%s" % (
            rname, (" (%s%s)" % (rule.get("v1"), " – " + rule["v2"] if rule.get("v2") else "")) if rule.get("v1") else "")}
        if isinstance(val, tuple) and val[0] == "__calc__":
            return self._t_calc(res, page, form, f, rule, val[1], val[2], meta)
        if not f:
            s = self._st(res, "Tìm trường '%s' trên trang %s" % (rule.get("field_label"), rule.get("page_url")),
                         "Có trường này")
            return self._bad(s, "Không tìm thấy trường (trang đã thay đổi?). Hãy bấm 'Dò form' lại.", shot=False,
                             status="Blocked")
        values = {}
        if isinstance(val, tuple) and val[0] == "__after__":
            other = next((x for x in form["fields"] if C.norm(x["label"]) == C.norm(val[1])), None)
            if not other:
                s = self._st(res, "Tìm trường '%s'" % val[1], "Có trường này")
                return self._bad(s, "Không có trường '%s' trong form" % val[1], shot=False, status="Blocked")
            base = TODAY() + dt.timedelta(days=10)
            values[other["sel"]] = self._date_val(other, base)
            values[f["sel"]] = self._date_val(f, base + dt.timedelta(days=val[2]))
        elif callable(val):
            values[f["sel"]] = val(f)
        elif val != "__valid__":
            values[f["sel"]] = val
        data = "; ".join("%s = %s" % (x["label"], values[x["sel"]]) for x in form["fields"] if x["sel"] in values)

        def chk(special):
            got = special.get(f["sel"])
            if got == "__not_typable__" and expect == "reject":
                return ("Kiểm tra ô '%s'" % f["label"], "Không cho nhập giá trị sai", True,
                        "Ô không cho nhập giá trị này")
            if rule["rule"] == "maxlen" and expect == "reject" and isinstance(got, str):
                n = int(float(rule.get("v1") or 0))
                if len(got) <= n:
                    return ("Kiểm tra số ký tự trong ô", "Tối đa %d ký tự" % n, True,
                            "Ô tự cắt còn %d ký tự" % len(got))
            return None
        meta["input"] = data or "(hợp lệ)"
        self._form_case(res, page, form, values, expect, exp_msg=msg, note_data=data, pre_check=chk, meta=meta)

    def _t_calc(self, res, page, form, f, rule, inputs, expected, meta):
        """Luật TỰ TÍNH: nhập các trường đầu vào → đọc trường kết quả → so với giá trị mong đợi."""
        pairs = [x.split("=", 1) for x in re.split(r"[;\n]+", inputs or "") if "=" in x]
        forms = (page.get("info") or {}).get("forms", [])
        if form is None:
            want = {C.norm(k) for k, _ in pairs}
            form = next((fm for fm in forms if want & {C.norm(x["label"]) for x in fm["fields"]}), None)
        s1 = self._st(res, "Mở form '%s'" % ((form or {}).get("name") or ""), "Form hiện ra")
        meta["input"] = inputs
        if not form:
            s1["m"] = dict(meta, expect=expected, act_msg="Không tìm thấy form")
            return self._bad(s1, "Không tìm thấy form có các trường %s" % inputs, shot=False, status="Blocked")
        self._open_form(page, form)
        self._ok(s1, "Đã mở form")
        s2 = self._st(res, "Nhập: %s" % inputs, "", inputs)
        miss = []
        for k, v in pairs:
            fld = next((x for x in form["fields"] if C.norm(x["label"]) == C.norm(k)), None)
            if not fld:
                miss.append(k.strip())
                continue
            loc = self.page.locator(fld["sel"]).first
            if fld["tag"] == "select":
                loc.select_option(label=v.strip())
            else:
                loc.fill(v.strip())
                loc.press("Tab")
        self._pump(600)
        if miss:
            s2["m"] = dict(meta, expect=expected, act_msg="Thiếu trường")
            return self._bad(s2, "Không có trường: %s" % ", ".join(miss), status="Blocked")
        self._ok(s2, "Đã nhập")
        label = rule.get("field_label") or ""
        s3 = self._st(res, "Đọc giá trị '%s'" % label, "= %s" % expected)
        got = None
        try:
            if f:
                loc = self.page.locator(f["sel"]).first
                try:
                    got = loc.input_value(timeout=1500)
                except Exception:
                    got = loc.inner_text(timeout=1500)
            else:
                loc = self.find(label, timeout=3)
                try:
                    got = loc.input_value(timeout=1500)
                except Exception:
                    got = loc.inner_text(timeout=1500)
        except Exception as e:
            s3["m"] = dict(meta, expect=expected, act_msg="Không đọc được")
            return self._bad(s3, "Không đọc được '%s': %s" % (label, self._friendly(e, {})), status="Blocked")
        got = (got or "").strip()

        def num(x):
            t = re.sub(r"[^\d,.\-]", "", x or "").replace(",", ".")
            try:
                return float(t)
            except ValueError:
                return None
        same = (num(got) is not None and num(expected) is not None and abs(num(got) - num(expected)) < 1e-6) or \
            C.norm(got) == C.norm(expected)
        s3["m"] = dict(meta, expect=expected, act_msg=got)
        (self._ok if same else self._bad)(s3, "Đang hiện: %s" % (got or "(trống)"))

    # ---------- State
    def _plan_state(self, p):
        out = []
        info = p["info"]
        name = p.get("name")
        if self.o("state", "toggles") == "yes":
            done = set()
            for t in info.get("toggles", [])[:8]:
                key = (t["text"] or t["aria"]).lower()
                if not key or key in done or not self._allowed(p, t["text"] or t["aria"], self._safe_button(t)):
                    continue
                done.add(key)
                out.append(("state", "Bấm mở / đóng '%s'" % (t["text"] or t["aria"]),
                            self._S("state", "toggle", name, "Kiểm tra mở/đóng bộ lọc nâng cao, menu trên trang %s"
                                    % name), self._kw(p),
                            lambda res, t=t: self._t_toggle(res, p, t)))
            for sw in info.get("switches", [])[:5]:
                out.append(("state", "Bật / tắt '%s'" % (sw["text"] or "nút gạt"),
                            self._S("state", "switch", name, "Kiểm tra trạng thái và chức năng của nút bật/tắt (toggle) "
                                                           "trên trang %s" % name), self._kw(p),
                            lambda res, sw=sw: self._t_switch(res, p, sw)))
        if self.o("state", "popups") == "yes":
            seen = set()
            pops = list(p.get("popups", []))
            for f in info.get("forms", []):
                if f.get("opener") and not f.get("visible"):
                    pops.append({"opener": f["opener"], "dialog": f["sel"]})
            for pp in pops:
                k = pp["opener"]["text"].lower()
                if k in seen:
                    continue
                seen.add(k)
                op = C.norm(pp["opener"]["text"])
                lu = 9 if re.search(r"nhap du lieu|import|tai len|upload|tep|file", op) else \
                    5 if re.search(r"them|tao|new|add|create", op) else 7 if re.search(r"xem|chi tiet|view|detail", op) \
                    else 8
                out.append(("state", "Mở và đóng popup '%s'" % pp["opener"]["text"],
                            self._S("state", "popup", name, "Kiểm tra mở/đóng popup '%s' trên trang %s bằng nút Đóng "
                                                          "hoặc biểu tượng X" % (pp["opener"]["text"], name), luong=lu),
                            self._kw(p), lambda res, pp=pp: self._t_popup(res, p, pp)))
        return out

    def _snap(self):
        return self.page.evaluate(SNAP_JS)

    @staticmethod
    def _diff(a, b):
        return (a["dialogs"] != b["dialogs"] or a["forms"] != b["forms"] or a["exp"] != b["exp"] or
                a["chk"] != b["chk"] or abs(a["visCount"] - b["visCount"]) >= 1 or abs(a["textLen"] - b["textLen"]) > 20)

    def _t_toggle(self, res, p, t):
        self._goto(p["url"])
        c0 = len(self.console_errors)
        s0 = self._snap()
        s1 = self._st(res, "Bấm '%s' lần 1" % (t["text"] or t["aria"]), "Trạng thái thay đổi (mở ra / hiện thêm)")
        s1["m"] = {"screen": p.get("name"), "object": t["text"] or t["aria"], "from": "Đóng / thu gọn",
                   "action": "Bấm lần 1", "expect": "Mở / hiện ra"}
        self.page.locator(t["sel"]).first.click()
        self._pump(500)
        a = self._snap()
        s1["m"]["actual"] = "Mở / hiện ra" if self._diff(s0, a) else "Không đổi"
        if self._diff(s0, a):
            self._ok(s1, "Trạng thái đã thay đổi" + (" (aria-expanded: %s)" % a["exp"] if a["exp"] != s0["exp"] else ""))
            self._shot(s1, "mo")
        else:
            self._bad(s1, "Bấm nhưng không thấy gì thay đổi")
        s2 = self._st(res, "Bấm '%s' lần 2" % (t["text"] or t["aria"]), "Trở về trạng thái ban đầu (đóng / thu gọn)")
        s2["m"] = {"screen": p.get("name"), "object": t["text"] or t["aria"], "from": "Mở", "action": "Bấm lần 2",
                   "expect": "Đóng / thu gọn"}
        try:
            self.page.locator(t["sel"]).first.click(timeout=4000)
        except Exception as e:
            return self._bad(s2, "Không bấm lại được (nút bị che khi đang mở?): %s" % self._friendly(e, {}))
        self._pump(500)
        b = self._snap()
        s2["m"]["actual"] = "Đóng / thu gọn" if not self._diff(s0, b) else "Vẫn mở"
        if not self._diff(s0, b):
            self._ok(s2, "Đã trở về trạng thái ban đầu")
        else:
            self._bad(s2, "Chưa trở về trạng thái ban đầu")
        s3 = self._st(res, "Kiểm tra lỗi", "Không có lỗi JavaScript")
        errs = self.console_errors[c0:]
        self._bad(s3, " | ".join(errs[:3]), shot=False) if errs else self._ok(s3, "Không có lỗi")

    def _t_switch(self, res, p, sw):
        self._goto(p["url"])
        loc = self.page.locator(sw["sel"]).first
        v0 = loc.is_checked()
        for n in (1, 2):
            s = self._st(res, "Bấm nút gạt '%s' lần %d" % (sw["text"], n), "Trạng thái chuyển %s" % (
                "BẬT" if (not v0) == (n == 1) else "TẮT"))
            try:
                loc.click(timeout=4000)
            except Exception:
                loc.click(force=True)
            self._pump(400)
            v = loc.is_checked()
            want = (not v0) if n == 1 else v0
            s["m"] = {"screen": p.get("name"), "object": sw["text"], "from": "BẬT" if (v0 if n == 1 else not v0) else "TẮT",
                      "action": "Bấm lần %d" % n, "expect": "BẬT" if want else "TẮT", "actual": "BẬT" if v else "TẮT"}
            (self._ok if v == want else self._bad)(s, "Đang %s" % ("BẬT" if v else "TẮT"))

    def _t_popup(self, res, p, pp):
        self._goto(p["url"])
        s1 = self._st(res, "Bấm '%s'" % pp["opener"]["text"], "Popup hiện ra")
        s1["m"] = {"screen": p.get("name"), "object": "Popup '%s'" % pp["opener"]["text"], "from": "Đóng",
                   "action": "Bấm '%s'" % pp["opener"]["text"], "expect": "Mở"}
        self.page.locator(pp["opener"]["sel"]).first.click()
        self._pump(500)
        snap = self._snap()
        opened = pp["dialog"] in snap["dialogs"] + snap["forms"]
        s1["m"]["actual"] = "Mở" if opened else "Không mở"
        if not opened:
            return self._bad(s1, "Popup không hiện ra")
        self._ok(s1, "Popup đã hiện")
        self._shot(s1, "popup")
        close = self.page.evaluate(DIALOG_CLOSE_JS, pp["dialog"])
        s2 = self._st(res, "Bấm nút '%s' trên popup" % (close["text"] if close else "Đóng / X"), "Popup đóng lại")
        s2["m"] = {"screen": p.get("name"), "object": "Popup '%s'" % pp["opener"]["text"], "from": "Mở",
                   "action": "Bấm '%s'" % (close["text"] if close else "Đóng / X"), "expect": "Đóng", "actual": ""}
        if not close:
            self._bad(s2, "Không thấy nút Đóng / Hủy / X trên popup")
        else:
            self.page.locator(close["sel"]).first.click()
            self._pump(500)
            snap = self._snap()
            still = pp["dialog"] in snap["dialogs"] + snap["forms"]
            s2["m"]["actual"] = "Vẫn mở" if still else "Đóng"
            if still:
                self._bad(s2, "Bấm '%s' nhưng popup vẫn hiện" % close["text"])
            else:
                self._ok(s2, "Popup đã đóng")
        if self.o("state", "esc") == "yes":
            s3 = self._st(res, "Mở lại popup và nhấn phím Esc", "Popup đóng lại")
            self._goto(p["url"])
            self.page.locator(pp["opener"]["sel"]).first.click()
            self._pump(400)
            self.page.keyboard.press("Escape")
            self._pump(400)
            snap = self._snap()
            if pp["dialog"] in snap["dialogs"] + snap["forms"]:
                self._bad(s3, "Nhấn Esc nhưng popup không đóng")
            else:
                self._ok(s3, "Popup đã đóng bằng phím Esc")

    # ---------- Permission: MA TRẬN PHÂN QUYỀN (vai trò × nút/chức năng DÒ ĐƯỢC, tester chọn Được thấy / Không)
    def _plan_permission(self, pages):
        perm = self.cfg.get("perm") or {}
        role = (perm.get("role") or "").strip() or "Tài khoản 2"
        rows = [r for r in perm.get("rows", []) if r.get("expect") in ("allow", "deny")]
        mod = (pages[0].get("name") if pages else "") or "Hệ thống"
        scn = self._S("permission", "perm", mod, "Kiểm tra phạm vi chức năng của vai trò %s trên các màn hình (xem, thêm mới, sửa, xóa, "
                                                       "phê duyệt)" % role)
        holder = {}

        def login2(res):
            mode = self.o("permission", "login_mode")
            s = self._st(res, "Đăng nhập bằng tài khoản vai trò '%s'" % role, "Đăng nhập thành công",
                         "Tài khoản: %s" % (self.o("permission", "user") or "(tự đăng nhập)"))
            self.ctx2 = self._new_ctx()
            main_ctx, main_page = self.ctx, self.page
            self.ctx, self.page = self.ctx2, None
            self._ensure_page()
            try:
                if mode == "none":
                    self._ok(s, "Không đăng nhập (khách)")
                    holder["ok"] = True
                    return
                lg = self.cfg.get("login") or {}
                ok, msg = self._login_into(None, mode, (lg.get("url") or "").strip() or self.url,
                                           self.o("permission", "user"), self.o("permission", "pass"),
                                           who="tài khoản vai trò '%s'" % role)
                holder["ok"] = ok
                (self._ok if ok else self._bad)(s, msg)
            finally:
                holder["page2"] = self.page
                self.ctx, self.page = main_ctx, main_page
        out = [("permission", "Đăng nhập bằng tài khoản vai trò '%s'" % role, scn, {"module": "Phân quyền"}, login2)]
        if rows:
            order, groups = [], {}
            for r in rows:
                k = self._norm_url(r.get("page_url") or self.url)
                if k not in groups:
                    order.append(k)
                    groups[k] = []
                groups[k].append(r)
            for k in order:
                title = groups[k][0].get("page_title") or k
                out.append(("permission", "Vai trò '%s' trên trang '%s'" % (role, title), scn,
                            {"module": title, "pre": "Đã đăng nhập vai trò '%s'" % role},
                            lambda res, k=k, rs=groups[k], title=title: self._t_perm_matrix(res, k, title, rs, role,
                                                                                            holder)))
        else:
            expect = self.o("permission", "expect")
            for pg in pages[:15]:
                out.append(("permission", "Vai trò '%s' mở trang '%s'" % (role, pg.get("name")), scn, self._kw(
                    pg, pre="Đã đăng nhập vai trò '%s'" % role),
                    lambda res, pg=pg: self._t_perm(res, pg, expect, holder, role)))
        return out

    def _as_acc2(self, holder):
        main = (self.ctx, self.page)
        self.ctx, self.page = self.ctx2, holder.get("page2")
        self._ensure_page()
        return main

    def _t_perm_matrix(self, res, url, title, rows, role, holder):
        s0 = self._st(res, "Mở trang '%s' bằng vai trò '%s'" % (title, role), "Mở được trang (nếu vai trò có quyền)")
        if not holder.get("ok"):
            return self._bad(s0, "Chưa đăng nhập được tài khoản vai trò '%s'" % role, shot=False, status="Blocked")
        main = self._as_acc2(holder)
        try:
            resp = self._goto(url)
            st = resp.status if resp else 200
            denied = st in (401, 403) or self._is_login_page() or bool(DENY_RE.search(self.page.inner_text("body")[:3000]))
            self._ok(s0, "Bị chặn (HTTP %d)" % st if denied else "Mở được (HTTP %d)" % st)
            visible = {"": set() if denied else {C.norm(x) for x in self.page.evaluate(BTN_IN_JS, ["body"])}}
            opened_note = {}
            for where in sorted({r.get("where") or "" for r in rows} - {""}):
                visible[where] = set()
                if denied:
                    continue
                try:
                    self._goto(url)
                    before = self._snap()
                    loc = self.page.get_by_role("button", name=where, exact=True)
                    if not loc.count():
                        loc = self.page.get_by_text(where, exact=True)
                    if not loc.filter(visible=True).count():
                        opened_note[where] = "không thấy nút mở '%s'" % where
                        continue
                    loc.filter(visible=True).first.click(timeout=4000)
                    self._pump(500)
                    after = self._snap()
                    new = [x for x in after["dialogs"] + after["forms"] if x not in before["dialogs"] + before["forms"]]
                    visible[where] = {C.norm(x) for x in self.page.evaluate(BTN_IN_JS, new or ["body"])}
                except Exception as e:
                    opened_note[where] = "không mở được: %s" % self._friendly(e, {})
            for r in rows:
                where = r.get("where") or ""
                btn = r["button"]
                is_page = btn == "__page__"
                want = r["expect"] == "allow"
                label = "Mở trang" if is_page else ("Nút '%s'%s" % (btn, " (trong popup '%s')" % where if where else ""))
                s = self._st(res, "Kiểm tra: %s" % label, ("Được phép" if is_page else "ĐƯỢC thấy") if want else
                             ("KHÔNG được vào" if is_page else "KHÔNG được thấy"))
                seen = (not denied) if is_page else (C.norm(btn) in visible.get(where, set()))
                s["m"] = {"screen": title, "where": where or "(trang)", "action": "Truy cập trang" if is_page else btn,
                          "role": role, "expect": "Được phép" if want else "Không được phép",
                          "actual": ("Vào được" if seen else "Bị chặn") if is_page else ("Thấy" if seen else "Không thấy")}
                note = (" — " + opened_note[where]) if where in opened_note else ""
                if seen == want:
                    self._ok(s, s["m"]["actual"] + note)
                elif is_page:
                    self._bad(s, "Vai trò '%s' %s%s" % (role, "VẪN VÀO ĐƯỢC trang" if seen else "bị chặn không vào được",
                                                        note))
                else:
                    self._bad(s, "Vai trò '%s' %s%s" % (role, "VẪN THẤY nút này" if seen else "KHÔNG thấy nút này", note))
        finally:
            holder["page2"] = self.page
            self.ctx, self.page = main

    def _t_perm(self, res, p, expect, holder, role):
        s1 = self._st(res, "Mở %s bằng vai trò '%s'" % (p["url"], role), {
            "view_only": "Xem được trang", "no_access": "Bị chặn (chuyển về đăng nhập / báo không có quyền)",
            "same": "Xem được trang như tài khoản chính"}[expect])
        if not holder.get("ok"):
            return self._bad(s1, "Chưa đăng nhập được tài khoản 2", shot=False, status="Blocked")
        main = self._as_acc2(holder)
        try:
            resp = self._goto(p["url"])
            status = resp.status if resp else 200
            body = self.page.inner_text("body")
            denied = status in (401, 403) or self._is_login_page() or bool(DENY_RE.search(body[:3000]))
            s1["m"] = {"screen": p.get("name"), "where": "(trang)", "action": "Truy cập trang", "role": role,
                       "expect": "Không được phép" if expect == "no_access" else "Được phép",
                       "actual": "Bị chặn" if denied else "Vào được"}
            if expect == "no_access":
                if denied:
                    self._ok(s1, "Bị chặn — HTTP %d" % status)
                else:
                    self._bad(s1, "Vai trò '%s' VẪN XEM ĐƯỢC trang (HTTP %d)" % (role, status))
                return
            if denied:
                return self._bad(s1, "Không xem được trang (HTTP %d / bị chặn)" % status)
            self._ok(s1, "Xem được trang (HTTP %d)" % status)
            if expect == "view_only":
                s2 = self._st(res, "Kiểm tra các nút thao tác ghi (chưa chọn ma trận — tool dựa vào tên nút)",
                              "Vai trò chỉ xem KHÔNG thấy các nút này")
                texts = self.page.evaluate(BTN_IN_JS, ["body"])
                w = sorted({t for t in texts if WRITE_BTN_RE.search(C.norm(t))})
                s2["m"] = {"screen": p.get("name"), "where": "(trang)", "action": ", ".join(w) or "(nút ghi)",
                           "role": role, "expect": "Không được phép", "actual": "Thấy" if w else "Không thấy"}
                if w:
                    self._bad(s2, "Vẫn thấy nút: %s" % ", ".join(w[:8]))
                else:
                    self._ok(s2, "Không thấy nút thao tác ghi")
        finally:
            holder["page2"] = self.page
            self.ctx, self.page = main

    # ---------- Usability / UI
    def _plan_ui(self, p):
        name = p.get("name")
        scn = self._S("ui", "ui", name, "Kiểm tra giao diện, bố cục mặc định, breadcrumb và menu đang chọn của trang %s" % name)
        review = self.o("ui", "review") == "yes"
        return [("ui", "Giao diện trang '%s' hiển thị đúng" % name, scn,
                 self._kw(p, mode="Bán tự động" if review else "Tự động"), lambda res: self._t_ui(res, p, review))]

    def _t_ui(self, res, p, review):
        checks = set(self.o("ui", "checks") or [])
        info = p["info"]
        s0 = self._st(res, "Mở trang %s và chụp ảnh" % p["url"], "Trang hiển thị đầy đủ")
        self._goto(p["url"])
        self._shot(s0, "trang", full=self.o("ui", "shot") == "full")
        self._ok(s0, "Đã chụp ảnh trang")
        if "title" in checks:
            s = self._st(res, "Kiểm tra tiêu đề tab trình duyệt", "Có tiêu đề trang")
            self._ok(s, "Tiêu đề: %s" % info["title"]) if info.get("title") else self._bad(s, "Trang không có tiêu đề",
                                                                                         shot=False)
        if "images" in checks:
            s = self._st(res, "Kiểm tra ảnh trên trang", "Không có ảnh bị lỗi")
            self._bad(s, "%d ảnh lỗi: %s" % (len(info["broken"]), ", ".join(u[-60:] for u in info["broken"][:5]))) \
                if info.get("broken") else self._ok(s, "Không có ảnh lỗi")
        if "badtext" in checks:
            s = self._st(res, "Kiểm tra chữ hiển thị", "Không hiện giá trị lỗi (undefined, null, NaN, {{…}})")
            self._bad(s, "Trang hiện: %s" % ", ".join(info["bad"])) if info.get("bad") else self._ok(s, "Không có")
        if "overflow" in checks:
            s = self._st(res, "Kiểm tra tràn ngang", "Không có thanh cuộn ngang")
            ov = self.page.evaluate(OVERFLOW_JS)
            if ov["sw"] > ov["W"] + 1:
                self._bad(s, "Tràn ngang %dpx > %dpx: %s" % (ov["sw"], ov["W"], ", ".join(ov["bad"][:5])))
            else:
                self._ok(s, "Không tràn")
        cmp_ = self.o("ui", "compare")
        if cmp_ and cmp_ != "no":
            s = self._st(res, "So sánh với ảnh lần quét trước", "Khác không quá %s%%" % cmp_)
            self._compare(s, "ui_" + E.slug(p["path"] or "trang"), float(cmp_))
        if review:
            s = self._st(res, "Người test duyệt giao diện", "Bố cục, màu sắc, chữ hiển thị đúng thiết kế")
            self._shot(s, "duyet")
            self._ask_manual(self.inst, res, [s], reason="Duyệt giao diện trang %s" % p.get("name"))

    def _compare(self, s, key, limit):
        from PIL import Image, ImageChops
        os.makedirs(self.base_dir, exist_ok=True)
        vp = self.page.viewport_size or {"width": 0, "height": 0}
        base = os.path.join(self.base_dir, "%s_%dx%d_%s.png" % (key, vp["width"], vp["height"], self.inst["browser"]))
        cur = self._shot(s, "sosanh")
        s["baseline"] = base
        if not cur:
            return self._bad(s, "Không chụp được ảnh", shot=False)
        if not os.path.exists(base):
            shutil.copyfile(cur, base)
            return self._ok(s, "Lần đầu: đã lưu làm ảnh chuẩn cho lần sau")
        a, b = Image.open(base).convert("RGB"), Image.open(cur).convert("RGB")
        if a.size != b.size:
            return self._bad(s, "Kích thước ảnh khác ảnh chuẩn", shot=False)
        diff = ImageChops.difference(a, b).convert("L").point(lambda v: 255 if v > 40 else 0)
        ratio = sum(1 for v in diff.getdata() if v) * 100.0 / (a.size[0] * a.size[1])
        if ratio > limit:
            hl = Image.blend(b, Image.composite(Image.new("RGB", b.size, (255, 0, 0)), b, diff), 0.6)
            dn = os.path.basename(cur).replace(".png", "_khac.png")
            hl.save(os.path.join(self.img_dir, dn))
            s["images"].append("anh/" + dn)
            return self._bad(s, "Khác ảnh lần trước %.2f%% (cho phép %g%%)" % (ratio, limit), shot=False)
        self._ok(s, "Giống ảnh lần trước (khác %.2f%%)" % ratio)

    # ---------- Responsive
    def _sizes(self):
        sizes = []
        for x in (self.o("responsive", "sizes") or []) + re.split(r"[,;\s]+", self.o("responsive", "custom") or ""):
            s = X.size_of(x)
            if s and s not in sizes:
                sizes.append(s)
        return sizes

    def _plan_responsive(self, p):
        name = p.get("name")
        scn = self._S("responsive", "responsive", name, "Kiểm tra khả năng hiển thị và hoạt động của trang %s khi "
                                                         "thay đổi kích thước màn hình" % name)
        review = self.o("responsive", "review") == "yes"
        return [("responsive", "Trang '%s' ở kích thước %dx%d" % (name, w, h), scn,
                 self._kw(p, size="%dx%d" % (w, h), mode="Bán tự động" if review else "Tự động"),
                 lambda res, w=w, h=h: self._t_resp(res, p, w, h, review)) for (w, h) in self._sizes()]

    def _t_resp(self, res, p, w, h, review):
        default = X.size_of(self.opt.get("size", "")) or (1366, 768)
        try:
            self.page.set_viewport_size({"width": default[0], "height": default[1]})
            self._goto(p["url"])
            c0 = len(self.console_errors)
            s1 = self._st(res, "Đổi kích thước cửa sổ sang %dx%d khi đang mở trang" % (w, h),
                          "Không phát sinh lỗi JavaScript khi đổi kích thước")
            self.page.set_viewport_size({"width": w, "height": h})
            self._pump(700)
            errs = self.console_errors[c0:]
            self._bad(s1, "Lỗi JavaScript: %s" % " | ".join(errs[:3])) if errs else self._ok(s1, "Không có lỗi")
            s2 = self._st(res, "Tải lại trang ở %dx%d" % (w, h), "Trang hiển thị, không tràn ngang")
            c1 = len(self.console_errors)
            self._goto(p["url"])
            ov = self.page.evaluate(OVERFLOW_JS)
            if ov["sw"] > ov["W"] + 1:
                self._bad(s2, "Tràn ngang %dpx > %dpx: %s" % (ov["sw"], ov["W"], ", ".join(ov["bad"][:5])))
            else:
                self._ok(s2, "Không tràn ngang")
                self._shot(s2, "kt")
            errs = self.console_errors[c1:]
            if errs:
                s3 = self._st(res, "Kiểm tra lỗi khi tải ở %dx%d" % (w, h), "Không có lỗi JavaScript")
                self._bad(s3, " | ".join(errs[:3]), shot=False)
            if review:
                s4 = self._st(res, "Người test duyệt bố cục ở %dx%d" % (w, h), "Bố cục gọn, không vỡ, đủ chức năng")
                self._shot(s4, "duyet", full=True)
                self._ask_manual(self.inst, res, [s4], reason="Duyệt bố cục %s ở %dx%d" % (p.get("name"), w, h))
        finally:
            try:
                self.page.set_viewport_size({"width": default[0], "height": default[1]})
            except Exception:
                pass

    # ---------- Performance
    def _plan_perf(self, p):
        name = p.get("name")
        scn = self._S("performance", "perf", name, "Kiểm tra thời gian tải và dung lượng trang %s trong giới hạn cho phép" % name)
        return [("performance", "Tốc độ tải trang '%s'" % name, scn, self._kw(p), lambda res: self._t_perf(res, p))]

    def _t_perf(self, res, p):
        limit = float(self.o("performance", "max_load") or 3)
        maxmb = float(self.o("performance", "max_weight") or 0)
        runs = int(self.o("performance", "runs") or 1)
        loads, perf = [], p.get("perf") or {}
        if perf.get("load"):
            loads.append(perf["load"] / 1000.0)
        for _ in range(max(0, runs - len(loads))):
            self._goto(p["url"])
            perf = self.page.evaluate(PERF_JS)
            if perf.get("load"):
                loads.append(perf["load"] / 1000.0)
        s1 = self._st(res, "Đo thời gian tải trang (%d lần)" % max(1, len(loads)), "Dưới %g giây" % limit)
        if not loads:
            self._bad(s1, "Không đo được", shot=False, status="Blocked")
        else:
            avg = sum(loads) / len(loads)
            (self._ok if avg <= limit else self._bad)(s1, "Trung bình %.2f giây" % avg, ) if avg <= limit else \
                self._bad(s1, "Tải mất %.2f giây (chậm hơn %g giây)" % (avg, limit), shot=False)
        s2 = self._st(res, "Đo dung lượng trang", "Dưới %g MB" % maxmb if maxmb else "Ghi nhận dung lượng")
        mb = (perf.get("bytes") or 0) / 1048576.0
        if maxmb and mb > maxmb:
            self._bad(s2, "%.2f MB, %d request" % (mb, perf.get("requests", 0)), shot=False)
        else:
            self._ok(s2, "%.2f MB, %d request" % (mb, perf.get("requests", 0)))
        s3 = self._st(res, "Kiểm tra tài nguyên chậm", "Không có tài nguyên tải quá %g giây" % limit)
        slow = [x for x in perf.get("slow", []) if x["ms"] > limit * 1000]
        if slow:
            self._bad(s3, "; ".join("%s (%.1fs)" % (x["name"][-60:], x["ms"] / 1000.0) for x in slow[:5]), shot=False)
        else:
            self._ok(s3, "Không có" + (" (chậm nhất: %s %.1fs)" % (perf["slow"][0]["name"][-50:],
                                                                 perf["slow"][0]["ms"] / 1000.0) if perf.get("slow")
                                         else ""))

    # ---------- Security
    def _plan_security(self, pages):
        checks = set(self.o("security", "checks") or [])
        scn = self._S("security", "sec", "Hệ thống", "Kiểm tra truy cập trực tiếp URL khi chưa đăng nhập và các yêu cầu bảo "
                                                   "mật cơ bản (lỗi hệ thống, mật khẩu, cookie, header, HTTPS)")
        out = []
        first = pages[0] if pages else None
        if "anon" in checks and self.logged_in:
            out.append(("security", "Truy cập các trang khi CHƯA đăng nhập", scn, {"module": "Bảo mật",
                                                                                   "pre": "Chưa đăng nhập"},
                        lambda res: self._t_anon(res, pages[:15])))
        if "errors" in checks:
            out.append(("security", "Không lộ thông báo lỗi hệ thống trên các trang", scn, {"module": "Bảo mật"},
                        lambda res: self._t_errs(res, pages)))
        if "password" in checks:
            recs = [p for p in pages if (p.get("info") or {}).get("passwords")]
            if self.login_rec:
                recs = [self.login_rec] + recs
            if recs:
                out.append(("security", "Ô mật khẩu được bảo vệ", scn, {"module": "Bảo mật"},
                            lambda res: self._t_pw(res, recs)))
        if first and "cookies" in checks:
            out.append(("security", "Cookie phiên đăng nhập có HttpOnly / Secure / SameSite", scn,
                        {"module": "Bảo mật", "priority": "Medium"}, lambda res: self._t_cookies(res)))
        if first and "headers" in checks:
            out.append(("security", "Header bảo mật của máy chủ", scn, {"module": "Bảo mật", "priority": "Medium"},
                        lambda res: self._t_headers(res, first)))
        if first and "https" in checks:
            out.append(("security", "Hệ thống dùng HTTPS", scn, {"module": "Bảo mật", "priority": "Medium"},
                        lambda res: self._t_https(res, pages)))
        return out

    def _t_anon(self, res, pages):
        ctx = self._new_ctx()
        try:
            pg = ctx.new_page()
            for p in pages:
                self._check_stop()
                s = self._st(res, "Mở %s khi chưa đăng nhập" % p["path"], "Bị chuyển về trang đăng nhập / báo cần "
                                                                          "đăng nhập")
                try:
                    r = pg.goto(p["url"], wait_until="load")
                    try:
                        pg.wait_for_load_state("networkidle", timeout=2500)
                    except Exception:
                        pass
                    st = r.status if r else 200
                    has_pw = pg.locator("input[type=password]").filter(visible=True).count() > 0
                    body = pg.inner_text("body")[:3000]
                    if st in (401, 403) or has_pw or DENY_RE.search(body) or self._norm_url(pg.url) != self._norm_url(
                            p["url"]):
                        self._ok(s, "Bị chặn → %s (HTTP %d)" % (pg.url, st))
                    else:
                        name = "%03d_%s_anon.png" % (self.inst["n"] + 1, E.slug(p["path"], 30))
                        pg.screenshot(path=os.path.join(self.img_dir, name))
                        s["images"].append("anh/" + name)
                        self._bad(s, "XEM ĐƯỢC trang khi chưa đăng nhập (HTTP %d)" % st, shot=False)
                except Exception as e:
                    self._bad(s, "Không mở được: %s" % self._friendly(e, {}), shot=False, status="Blocked")
        finally:
            ctx.close()

    def _t_errs(self, res, pages):
        for p in pages:
            s = self._st(res, "Kiểm tra nội dung trang %s" % p["path"], "Không hiện chi tiết lỗi kỹ thuật (stack "
                                                                        "trace, SQL…)")
            e = (p.get("info") or {}).get("errText")
            self._bad(s, "Lộ thông tin: %s" % e, shot=False) if e else self._ok(s, "Không có")

    def _t_pw(self, res, recs):
        for p in recs:
            for pw in (p.get("info") or {}).get("passwords", []):
                s = self._st(res, "Ô mật khẩu ở %s" % urlparse(p["url"]).path, "Gửi bằng POST, không điền sẵn mật khẩu")
                bad = []
                if pw.get("method") == "get":
                    bad.append("form gửi bằng GET (mật khẩu lộ trên URL)")
                if pw.get("prefilled"):
                    bad.append("mật khẩu được điền sẵn")
                if pw.get("action", "").startswith("http://") and self.origin[0] == "https":
                    bad.append("gửi mật khẩu qua HTTP")
                self._bad(s, "; ".join(bad), shot=False) if bad else self._ok(s, "An toàn (autocomplete=%s)" % (
                    pw.get("autocomplete") or "mặc định"))

    def _t_cookies(self, res):
        s = self._st(res, "Đọc cookie của hệ thống", "Cookie phiên có HttpOnly, SameSite (và Secure nếu HTTPS)")
        try:
            cks = [c for c in self.ctx.cookies() if self.origin[1].split(":")[0].endswith(c["domain"].lstrip("."))]
        except Exception:
            cks = []
        sess = [c for c in cks if re.search(r"sess|sid|token|auth|jwt|login|connect|asp|php|jsession|remember",
                                            c["name"], re.I)] or cks
        if not sess:
            return self._ok(s, "Không có cookie (có thể dùng token trong bộ nhớ trình duyệt)")
        bad = []
        for c in sess:
            miss = []
            if not c.get("httpOnly"):
                miss.append("HttpOnly")
            if self.origin[0] == "https" and not c.get("secure"):
                miss.append("Secure")
            if (c.get("sameSite") or "None") == "None":
                miss.append("SameSite")
            if miss:
                bad.append("%s thiếu %s" % (c["name"], "/".join(miss)))
        self._bad(s, "; ".join(bad[:6]), shot=False) if bad else self._ok(s, "%d cookie đều đủ thuộc tính" % len(sess))

    def _t_headers(self, res, p):
        h = {k.lower(): v for k, v in (p.get("headers") or {}).items()}
        need = [("x-frame-options", "Chống nhúng trang (X-Frame-Options hoặc CSP frame-ancestors)",
                 lambda: "x-frame-options" in h or "frame-ancestors" in h.get("content-security-policy", "")),
                ("x-content-type-options", "X-Content-Type-Options: nosniff",
                 lambda: "nosniff" in h.get("x-content-type-options", "").lower()),
                ("content-security-policy", "Content-Security-Policy", lambda: "content-security-policy" in h)]
        if self.origin[0] == "https":
            need.append(("strict-transport-security", "Strict-Transport-Security (HSTS)",
                         lambda: "strict-transport-security" in h))
        for key, desc, ok in need:
            s = self._st(res, "Kiểm tra header %s" % key, "Có %s" % desc)
            self._ok(s, "Có: %s" % h.get(key, "(trong CSP)")[:80]) if ok() else self._bad(
                s, "Máy chủ không trả về header %s" % key, shot=False)
        s = self._st(res, "Kiểm tra header lộ phiên bản máy chủ", "Không lộ phiên bản (Server, X-Powered-By)")
        leak = ["%s: %s" % (k, h[k]) for k in ("server", "x-powered-by", "x-aspnet-version") if k in h and
                re.search(r"\d", h[k])]
        self._bad(s, "; ".join(leak), shot=False) if leak else self._ok(s, "Không lộ")

    def _t_https(self, res, pages):
        s = self._st(res, "Kiểm tra giao thức", "Dùng HTTPS")
        if self.origin[0] == "https":
            self._ok(s, "Dùng HTTPS")
            s2 = self._st(res, "Kiểm tra tài nguyên HTTP trên trang HTTPS", "Không có (mixed content)")
            self._ok(s2, "Không phát hiện")
        else:
            self._bad(s, "Hệ thống dùng HTTP — mật khẩu và dữ liệu gửi đi không được mã hóa", shot=False)

    # ---------- Accessibility
    def _plan_a11y(self, p):
        name = p.get("name")
        scn = self._S("accessibility", "a11y", name, "Kiểm tra khả năng truy cập (accessibility) trang %s: mô tả ảnh, nhãn ô "
                                                       "nhập, độ tương phản" % name)
        return [("accessibility", "Khả năng truy cập trang '%s'" % name, scn, self._kw(p),
                 lambda res: self._t_a11y(res, p))]

    def _t_a11y(self, res, p):
        lv = self.o("accessibility", "level")
        ratio = {"AA": 4.5, "A": 3.0}.get(lv, 0)
        self._goto(p["url"])
        a = self.page.evaluate(A11Y_JS, ratio)
        rows = [("Ảnh có mô tả (alt)", "Mọi ảnh đều có alt", a["imgsNoAlt"], "Ảnh thiếu alt: %s"),
                ("Ô nhập có nhãn", "Mọi ô nhập có nhãn gắn với ô (label/aria-label)", a["noLabel"],
                 "Ô thiếu nhãn: %s"),
                ("Nút / link có tên", "Mọi nút, link có chữ hoặc aria-label", a["noName"], "Không có tên: %s")]
        for buoc, exp, bad, fmt in rows:
            s = self._st(res, buoc, exp)
            self._bad(s, fmt % ", ".join(bad[:6]), shot=False) if bad else self._ok(s, "Đạt")
        s = self._st(res, "Khai báo ngôn ngữ trang", "Thẻ html có thuộc tính lang")
        self._ok(s, "lang=%s" % a["lang"]) if a["lang"] else self._bad(s, "Thiếu thuộc tính lang", shot=False)
        s = self._st(res, "Cấu trúc tiêu đề", "Có 1 tiêu đề h1, không nhảy cấp")
        probs = ([] if a["h1"] == 1 else ["%d thẻ h1" % a["h1"]]) + a["jumps"]
        self._bad(s, "; ".join(probs), shot=False) if probs else self._ok(s, "Đạt")
        if ratio:
            s = self._st(res, "Độ tương phản chữ", "Tỷ lệ ≥ %g:1" % ratio)
            self._bad(s, "Chữ khó đọc: %s" % "; ".join(a["low"][:5])) if a["low"] else self._ok(s, "Đạt")

    # ---------- Compatibility
    def _plan_compat(self, pages):
        out = []
        for b in self.o("compat", "browsers") or []:
            scn = self._S("compat", "compat", "Hệ thống", "Kiểm tra hệ thống hiển thị và hoạt động đúng trên trình "
                                                       "duyệt %s" % BROWSER_NAME.get(b, b))
            for p in pages[:10]:
                out.append(("compat", "Trang '%s' trên %s" % (p.get("name"), BROWSER_NAME.get(b, b)), scn,
                            self._kw(p, browser=b), lambda res, p=p, b=b: self._t_compat(res, p, b)))
        return out

    def _t_compat(self, res, p, b):
        ctx = self._new_ctx(browser=b, storage=self.storage)
        main_ctx, main_page = self.ctx, self.page
        self.ctx, self.page = ctx, None
        try:
            self._ensure_page()
            c0 = len(self.console_errors)
            s1 = self._st(res, "Mở %s trên %s" % (p["path"], BROWSER_NAME.get(b, b)), "Trang tải thành công")
            r = self._goto(p["url"])
            st = r.status if r else 200
            self._ok(s1, "HTTP %d" % st) if st < 400 else self._bad(s1, "HTTP %d" % st)
            s2 = self._st(res, "Kiểm tra lỗi JavaScript", "Không có lỗi")
            errs = self.console_errors[c0:]
            self._bad(s2, " | ".join(errs[:3]), shot=False) if errs else self._ok(s2, "Không có lỗi")
            s3 = self._st(res, "Kiểm tra hiển thị", "Không tràn ngang")
            ov = self.page.evaluate(OVERFLOW_JS)
            if ov["sw"] > ov["W"] + 1:
                self._bad(s3, "Tràn ngang: %s" % ", ".join(ov["bad"][:4]))
            else:
                self._ok(s3, "Hiển thị bình thường")
                self._shot(s3, "tt")
        finally:
            self.ctx, self.page = main_ctx, main_page
            try:
                ctx.close()
            except Exception:
                pass

    # ---------- Functional: các nút của form (Lưu, Lưu nháp, Gửi phê duyệt, Hủy…) — lấy đúng nút DÒ ĐƯỢC trong form
    ACT_RE = re.compile(r"\b(luu|gui|duyet|nhap|tao|them|cap nhat|xac nhan|hoan tat|save|submit|send|create|add|"
                        r"update|confirm)\b")
    CANCEL_RE = re.compile(r"^(huy|dong|close|cancel|quay lai|back|x)$")

    def _plan_form_actions(self, p):
        out = []
        for form in self._data_forms(p):
            fname = form.get("name") or (form.get("submit") or {}).get("text") or "form"
            scn = self._S("functional", "form_actions", p.get("name"), "Kiểm tra chức năng các nút trên form %s (Lưu, Lưu nháp, Gửi "
                                                                     "phê duyệt, Hủy…)" % fname)
            pre = self._pre() + "; mở form '%s'" % fname
            for b in form.get("buttons", []):
                t = C.norm(b["text"]).strip()
                if not t or not self._allowed(p, b["text"], default=True):
                    continue
                if self.CANCEL_RE.match(t):
                    out.append(("functional", "Bấm '%s' trên form '%s': đóng form, không lưu" % (b["text"], fname), scn,
                                self._kw(p, pre=pre, priority="Medium"),
                                lambda res, f=form, b=b: self._t_form_button(res, p, f, b, cancel=True)))
                else:
                    out.append(("functional", "Bấm '%s' trên form '%s' với dữ liệu hợp lệ" % (b["text"], fname), scn,
                                self._kw(p, pre=pre, data="Dữ liệu hợp lệ do tool tạo"),
                                lambda res, f=form, b=b: self._t_form_button(res, p, f, b)))
        return out

    def _t_form_button(self, res, p, form, b, cancel=False):
        name = form.get("name") or "form"
        s1 = self._st(res, "Mở form '%s'" % name, "Form hiện ra")
        self._open_form(p, form)
        self._ok(s1, "Đã mở form")
        s2 = self._st(res, "Nhập dữ liệu hợp lệ", "")
        desc, _ = self._fill(form, {})
        s2["fields"]["du_lieu_buoc"] = desc
        self._ok(s2, desc or "Đã nhập")
        s3 = self._st(res, "Bấm '%s'" % b["text"], "Form đóng lại, KHÔNG lưu dữ liệu" if cancel else
                      "Hệ thống thực hiện đúng hành động '%s' (báo thành công / chuyển trang / đóng form)" % b["text"])
        s3["m"] = {"screen": p.get("name"), "function": "%s › %s" % (name, b["text"]), "action": "Bấm '%s'" % b["text"]}
        verdict, detail, msgs, js = self._submit_and_judge(form, button=b)
        if cancel:
            st = self._snap()
            if form["sel"] in st["forms"] + st["dialogs"]:
                self._bad(s3, "Bấm '%s' nhưng form vẫn hiện" % b["text"])
            elif verdict == "accepted" and "CHẤP NHẬN" in detail:
                self._bad(s3, "Form đóng nhưng có vẻ ĐÃ LƯU: %s" % detail)
            else:
                self._ok(s3, "Form đã đóng, không có thông báo lưu")
        elif verdict == "accepted":
            self._ok(s3, detail)
        elif verdict == "rejected":
            self._bad(s3, "Dữ liệu hợp lệ nhưng hệ thống báo lỗi — %s" % detail)
        else:
            self._unsure(res, s3, detail, "validation")
        if js:
            s4 = self._st(res, "Kiểm tra lỗi JavaScript", "Không có lỗi")
            self._bad(s4, " | ".join(js[:3]), shot=False)

    # ---------- Integration (Frontend – Backend)
    SPIN_JS = _js(r"""
      return [...document.querySelectorAll('[class*="loading"], [class*="spinner"], [class*="skeleton"], [aria-busy="true"], .loader, [class*="loader"]')]
        .filter(__vis).map((e) => e.tagName.toLowerCase() + (e.className && typeof e.className === 'string' ? '.' + e.className.trim().split(/\s+/)[0] : '')).slice(0, 5);
    """)
    OPTIONS_JS = _js(r"""
      return __a.map((s) => { let el = null; try { el = document.querySelector(s); } catch (e) {}
        if (!el || !el.options) return null; const opts = [...el.options].filter((o) => o.value !== '' && !/^\s*(--|chọn|select|vui lòng chọn|tất cả)/iu.test(o.text));
        return {sel: s, label: __label(el).text, n: opts.length, sample: opts.slice(0, 4).map((o) => o.text.trim())}; }).filter(Boolean);
    """, arg=True)
    COMBO_JS = _js(r"""
      return [...document.querySelectorAll('[role=combobox], [aria-haspopup=listbox]')].filter((e) => __vis(e) && e.tagName !== 'SELECT')
        .slice(0, 6).map((e) => ({sel: __sel(e), label: __label(e).text || __txt(e)}));
    """)

    def _plan_integration(self, p):
        out = []
        name = p.get("name")
        I = lambda feat, title: self._S("integration", feat, name, title)  # noqa: E731
        lo = name[:1].lower() + name[1:]
        if self.o("integration", "api_errors") == "yes":
            out.append(("integration", "API gọi khi mở trang '%s' không lỗi" % name,
                        I("api", "Kiểm tra các API được gọi khi mở trang %s trả về thành công, không lỗi" % name), self._kw(p),
                        lambda res: self._t_api_errors(res, p)))
        if self.o("integration", "dropdowns") == "yes":
            forms = [f for f in (p["info"].get("forms") or []) if (f.get("visible") or f.get("opener")) and
                     any(x["tag"] == "select" for x in f["fields"])]
            if forms:
                out.append(("integration", "Dropdown trên trang '%s' có dữ liệu từ hệ thống" % name,
                            I("dropdown", "Kiểm tra dữ liệu các dropdown trên trang %s được nạp đầy đủ từ hệ thống" % name),
                            self._kw(p),
                            lambda res, forms=forms: self._t_dropdowns(res, p, forms)))
        if self.o("integration", "loading") not in (None, "no"):
            out.append(("integration", "Trang '%s' tải xong dữ liệu, không treo 'đang tải'" % name,
                        I("loading", "Kiểm tra trang %s tải xong dữ liệu, không bị treo biểu tượng đang tải" % name),
                        self._kw(p, priority="Medium"), lambda res: self._t_loading(res, p)))
        if self.o("integration", "api_map") == "yes" and p.get("json"):
            objs = self._api_objects(p["json"])
            if objs:
                out.append(("integration", "Dữ liệu hiển thị trên '%s' khớp dữ liệu API trả về" % name,
                            I("mapping", "Kiểm tra dữ liệu hiển thị trên danh sách %s được mapping đúng với dữ liệu trả "
                                         "về" % lo),
                            self._kw(p), lambda res, objs=objs: self._t_api(res, p, objs)))
        return out

    def _t_api_errors(self, res, p):
        s1 = self._st(res, "Mở trang %s và theo dõi các lời gọi API" % p["path"], "Trang gọi API để lấy dữ liệu")
        self.resp_log = []
        self._goto(p["url"])
        calls = [r for r in self.resp_log if r["type"] in ("xhr", "fetch")]
        self._ok(s1, ("%d lời gọi API" % len(calls)) if calls else "Trang không gọi API riêng (dữ liệu có sẵn trong trang)")
        for r in calls[:15]:
            s = self._st(res, "API %s" % r["url"].split("?")[0][-80:], "Trả về thành công (2xx/3xx)")
            s["m"] = {"screen": p.get("name"), "component": "API", "api": r["url"].split("?")[0][-120:],
                      "http": r["status"] or r.get("failed"), "api_data": "", "ui_data": ""}
            if r["status"] >= 400 or r.get("failed"):
                self._bad(s, "Lỗi %s" % (r["status"] or r.get("failed")), shot=False)
            else:
                self._ok(s, "HTTP %d" % r["status"])
        for c in [c for c in p.get("clicks", []) if c.get("http")][:4]:
            s = self._st(res, "Bấm '%s' — kiểm tra API" % c["text"], "Không có lỗi máy chủ 5xx")
            s["m"] = {"screen": p.get("name"), "component": "Nút '%s'" % c["text"], "api": c["http"][0], "http": "5xx"}
            self._bad(s, "Lỗi máy chủ: %s" % "; ".join(c["http"][:2]), shot=False)

    def _t_dropdowns(self, res, p, forms):
        for form in forms:
            if not form.get("visible") and form.get("opener"):
                self._open_form(p, form)
                where = " (trong popup '%s')" % form["opener"]["text"]
            else:
                self._goto(p["url"])
                where = ""
            sels = [x["sel"] for x in form["fields"] if x["tag"] == "select"]
            for o in self.page.evaluate(self.OPTIONS_JS, sels):
                s = self._st(res, "Mở dropdown '%s'%s" % (o["label"] or "?", where), "Có dữ liệu để chọn")
                s["m"] = {"screen": p.get("name"), "component": "Dropdown '%s'%s" % (o["label"], where), "api": "",
                          "http": "", "api_data": "", "ui_data": "%d lựa chọn" % o["n"]}
                if o["n"] == 0:
                    self._bad(s, "Dropdown TRỐNG — backend chưa trả dữ liệu hoặc frontend chưa gắn dữ liệu")
                else:
                    self._ok(s, "%d lựa chọn: %s%s" % (o["n"], ", ".join(o["sample"]), "…" if o["n"] > 4 else ""))
        self._goto(p["url"])
        for cb in self.page.evaluate(self.COMBO_JS):
            s = self._st(res, "Mở dropdown '%s'" % cb["label"], "Hiện danh sách lựa chọn")
            s["m"] = {"screen": p.get("name"), "component": "Dropdown '%s'" % cb["label"]}
            try:
                self.page.locator(cb["sel"]).first.click(timeout=4000)
                self._pump(700)
                n = self.page.locator("[role=option]").filter(visible=True).count()
                self.page.keyboard.press("Escape")
                s["m"]["ui_data"] = "%d lựa chọn" % n
                (self._ok if n else self._bad)(s, "%d lựa chọn" % n if n else "Danh sách TRỐNG")
            except Exception as e:
                self._bad(s, "Không mở được: %s" % self._friendly(e, {}), shot=False, status="Blocked")

    def _t_loading(self, res, p):
        limit = float(self.o("integration", "loading") or 10)
        s = self._st(res, "Mở trang %s và chờ dữ liệu tải" % p["path"], "Không còn biểu tượng 'đang tải' sau %g giây"
                     % limit)
        s["m"] = {"screen": p.get("name"), "component": "Biểu tượng đang tải"}
        t0 = time.time()
        self._goto(p["url"])
        spin = self.page.evaluate(self.SPIN_JS)
        while spin and time.time() - t0 < limit:
            self._pump(500)
            spin = self.page.evaluate(self.SPIN_JS)
        if spin:
            self._bad(s, "Sau %g giây vẫn hiện biểu tượng đang tải: %s" % (limit, ", ".join(spin)))
        else:
            self._ok(s, "Dữ liệu tải xong sau %.1f giây" % (time.time() - t0))

    # ---------- State: bảng chuyển trạng thái do tester chọn (form ▾, nút ▾, trạng thái mong đợi ▾)
    ROW_JS = _js(r"""
      const els = [...document.querySelectorAll('tr, li, [role=row], [class*="card"], [class*="item"]')].filter(__vis)
        .filter((e) => (e.innerText || '').includes(__a));
      els.sort((a, b) => (a.innerText || '').length - (b.innerText || '').length);
      return els[0] ? (els[0].innerText || '').replace(/\s+/g, ' ').trim().slice(0, 300) : null;
    """, arg=True)

    def _plan_state_flows(self, page=None):
        out = []
        for i, r in enumerate(self.cfg.get("state_rules") or []):
            if not r.get("button") or not r.get("expect"):
                continue
            pu = self._norm_url(r.get("page_url") or self.url)
            if page is not None and pu != self._norm_url(page["url"]):
                continue
            if page is None and any(self._norm_url(x["url"]) == pu for x in self.pages):
                continue
            scn = self._S("state", "flow", r.get("page_title") or "", "Kiểm tra trạng thái bản ghi chuyển sang '%s' khi "
                                                                       "bấm '%s'" % (r["expect"], r["button"]),
                          ac=r.get("ac", ""))
            out.append(("state", "%s → bấm '%s' → trạng thái '%s'" % (r.get("from") or "Tạo mới", r["button"],
                                                                      r["expect"]), scn,
                        {"module": r.get("page_title") or "", "pre": self._pre() + "; mở form '%s'" % (
                            r.get("form_name") or ""), "data": "Dữ liệu hợp lệ + mã nhận diện do tool tạo",
                         "priority": "High"}, lambda res, r=r, i=i: self._t_flow(res, r, i)))
        return out

    def _find_form(self, rule):
        pu = self._norm_url(rule.get("page_url") or self.url)
        page = next((p for p in self.pages if self._norm_url(p["url"]) == pu), None)
        if page is None:
            page = self._visit(pu)
            if page.get("info"):
                self._probe(page)
            self.pages.append(page)
        forms = (page.get("info") or {}).get("forms", [])
        form = next((f for f in forms if f["sel"] == rule.get("form_sel")), None) or next(
            (f for f in forms if rule.get("form_name") and C.norm(f.get("name", "")) == C.norm(rule["form_name"])),
            None)
        if form is None:
            want = C.norm(rule.get("button", ""))
            form = next((f for f in forms if any(C.norm(b["text"]) == want for b in f.get("buttons", []))), None)
        return page, form

    def _t_flow(self, res, rule, i):
        page, form = self._find_form(rule)
        res["fields"]["man_hinh"] = page.get("path", "")
        m = {"screen": page.get("name"), "object": rule.get("form_name") or (form or {}).get("name", ""),
             "from": rule.get("from") or "(tạo mới)", "action": "Bấm '%s'" % rule["button"], "expect": rule["expect"],
             "actual": ""}
        if not form:
            s = self._st(res, "Tìm form có nút '%s'" % rule["button"], "Có form")
            s["m"] = m
            return self._bad(s, "Không tìm thấy form trên trang (hãy bấm 'Dò trang' lại)", shot=False,
                             status="Blocked")
        marker = "KTTT%s%02d" % (self.rnd, i)
        cand = [f for f in form["fields"] if self._rows_valid(f) and f["tag"] in ("input", "textarea") and
                f["type"] in ("text", "textarea")]
        text_f = next((f for f in cand if f["required"]), None) or (cand[0] if cand else None)
        values = {text_f["sel"]: ("Kiểm thử trạng thái " + marker)[:text_f.get("maxlength") or 200]} if text_f else {}
        s1 = self._st(res, "Mở form '%s'" % (form.get("name") or ""), "Form hiện ra")
        self._open_form(page, form)
        self._ok(s1, "Đã mở form")
        s2 = self._st(res, "Nhập dữ liệu hợp lệ (mã nhận diện: %s)" % marker, "")
        desc, _ = self._fill(form, values)
        s2["fields"]["du_lieu_buoc"] = desc
        self._ok(s2, desc)
        btn = next((b for b in form.get("buttons", []) if C.norm(b["text"]) == C.norm(rule["button"])), None)
        s3 = self._st(res, "Bấm '%s'" % rule["button"], "Hệ thống lưu bản ghi")
        if not btn:
            s3["m"] = m
            return self._bad(s3, "Form không có nút '%s'" % rule["button"], status="Blocked")
        verdict, detail, _, js = self._submit_and_judge(form, button=btn)
        if verdict == "rejected":
            s3["m"] = dict(m, actual="Không lưu được")
            return self._bad(s3, "Không lưu được — %s" % detail)
        self._ok(s3, detail or "Đã bấm")
        s4 = self._st(res, "Xem trạng thái bản ghi vừa tạo", "Trạng thái là '%s'" % rule["expect"])
        s4["m"] = m
        row = self.page.evaluate(self.ROW_JS, marker) if text_f else None
        if row is None and text_f:
            self._goto(page["url"])
            row = self.page.evaluate(self.ROW_JS, marker)
        if row is None:
            m["actual"] = "Không thấy bản ghi"
            return self._unsure(res, s4, "Không tìm thấy bản ghi vừa tạo (mã %s) trong danh sách" % marker, "business")
        found = [v for v in (page.get("info") or {}).get("statusValues", []) if C.norm(v) and C.norm(v) in C.norm(row)]
        m["actual"] = rule["expect"] if C.norm(rule["expect"]) in C.norm(row) else (found[0] if found else row[:80])
        if C.norm(rule["expect"]) in C.norm(row):
            self._ok(s4, "Bản ghi: %s" % row[:160])
        else:
            self._bad(s4, "Trạng thái KHÔNG phải '%s' (đang: %s). Bản ghi: %s" % (rule["expect"], m["actual"],
                                                                               row[:120]))

    # ---------- Usability
    USA_JS = _js(r"""
      const out = {covered: [], small: [], font: [], switches: []};
      const name = (e) => (__txt(e) || e.getAttribute('placeholder') || e.getAttribute('name') || e.id || e.tagName.toLowerCase()).slice(0, 40);
      const desc = (t) => t.tagName.toLowerCase() + (t.id ? '#' + t.id : '') + (typeof t.className === 'string' && t.className.trim() ? '.' + t.className.trim().split(/\s+/)[0] : '');
      const hit = (el, vis) => { vis.scrollIntoView({block: 'center', inline: 'nearest'}); const r = vis.getBoundingClientRect();
        const t = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
        if (!t || t === el || t === vis || el.contains(t) || vis.contains(t) || t.contains(vis)) return null;
        const lab = el.closest && el.closest('label'); if (lab && lab.contains(t)) return null;
        if (el.labels && [...el.labels].some((l) => l.contains(t))) return null; return t; };
      const els = [...document.querySelectorAll('button, a[href], input, select, textarea, [role=button], [role=tab]')]
        .filter((e) => __vis(e) && (e.getAttribute('type') || '') !== 'hidden').slice(0, 150);
      for (const el of els) { const r = el.getBoundingClientRect();
        const par = el.closest('p, li'); const inlineLink = el.tagName === 'A' && par && (par.innerText || '').length > (el.innerText || '').length + 20;
        const tp = (el.getAttribute('type') || '').toLowerCase();
        if (__a.target && !inlineLink && !['checkbox', 'radio'].includes(tp) && ((r.width < 24 && r.height < 24) || r.width < 16 || r.height < 16)) out.small.push({el: name(el), why: Math.round(r.width) + '×' + Math.round(r.height) + 'px'});
        if (__a.covered) { const t = hit(el, el); if (t) out.covered.push({el: name(el), why: 'bị che bởi ' + desc(t)}); }
        if (out.covered.length > 8 && out.small.length > 8) break; }
      if (__a.switch) for (const el of document.querySelectorAll('input[type=checkbox], input[type=radio], [role=switch]')) {
        const vis = __vis(el) ? el : (el.closest('label') || el.parentElement); if (!vis || !__vis(vis)) continue;
        const r = vis.getBoundingClientRect(); const t = hit(el, vis); const nm = __label(el).text || name(el);
        if (t) out.switches.push({el: nm, why: 'bị che bởi ' + desc(t)});
        else if (r.width < 16 || r.height < 16) out.switches.push({el: nm, why: 'quá nhỏ ' + Math.round(r.width) + '×' + Math.round(r.height) + 'px'});
        if (out.switches.length > 8) break; }
      if (__a.font) { const seen = new Set();
        for (const e of [...document.querySelectorAll('body *')].filter((e) => [...e.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim().length > 2) && __vis(e)).slice(0, 600)) {
          const fs = parseFloat(getComputedStyle(e).fontSize); if (fs < 12) { const t = [...e.childNodes].filter((n) => n.nodeType === 3).map((n) => n.textContent.trim()).join(' ').slice(0, 30);
            if (!seen.has(t)) { seen.add(t); out.font.push({el: t, why: fs + 'px'}); } } if (out.font.length >= 8) break; } }
      window.scrollTo(0, 0); return out;
    """, arg=True)

    def _plan_usability(self, p):
        name = p.get("name")
        scn = self._S("usability", "ease", name, "Kiểm tra các nút, ô nhập, nút gạt trên trang %s dễ thao tác, không bị che "
                                                  "khuất" % name)
        where = self.o("usability", "where")
        sizes = [("Máy tính", None)] if where == "desktop" else [("Điện thoại", (375, 812))] if where == "mobile" \
            else [("Máy tính", None), ("Điện thoại", (375, 812))]
        return [("usability", "Trang '%s' dễ thao tác trên %s" % (name, lab.lower()), scn,
                 self._kw(p, size=("%dx%d" % sz) if sz else ""),
                 lambda res, sz=sz, lab=lab: self._t_usability(res, p, sz, lab)) for lab, sz in sizes]

    def _t_usability(self, res, p, size, lab):
        checks = set(self.o("usability", "checks") or [])
        default = X.size_of(self.opt.get("size", "")) or (1366, 768)
        try:
            if size:
                self.page.set_viewport_size({"width": size[0], "height": size[1]})
            s0 = self._st(res, "Mở trang %s trên %s" % (p["path"], lab.lower()), "Trang hiển thị")
            self._goto(p["url"])
            self._ok(s0, "Đã mở")
            r = self.page.evaluate(self.USA_JS, {"covered": "covered" in checks, "target": "target" in checks,
                                                 "switch": "switch" in checks, "font": "font" in checks})
            rows = [("covered", "Nút / ô nhập không bị che khuất", "Không bị thanh điều khiển hay phần tử khác che",
                     r["covered"]),
                    ("target", "Kích thước nút bấm", "Nút đủ lớn để bấm (chuẩn WCAG 2.2: vùng bấm ≥ 24px)", r["small"]),
                    ("switch", "Nút gạt / checkbox bấm được", "Không bị che, không quá nhỏ", r["switches"]),
                    ("font", "Cỡ chữ dễ đọc", "Chữ không nhỏ hơn 12px", r["font"])]
            for key, buoc, exp, bad in rows:
                if key not in checks:
                    continue
                s = self._st(res, buoc, exp)
                s["m"] = {"screen": p.get("name"), "size": lab + (" %dx%d" % size if size else ""), "criterion": buoc,
                          "element": ", ".join(x["el"] for x in bad[:6]), "issue": ", ".join(x["why"] for x in bad[:6])}
                if bad:
                    self._bad(s, "; ".join("'%s' %s" % (x["el"], x["why"]) for x in bad[:6]),
                              shot=key in ("covered", "switch"))
                else:
                    self._ok(s, "Đạt")
        finally:
            try:
                self.page.set_viewport_size({"width": default[0], "height": default[1]})
            except Exception:
                pass
