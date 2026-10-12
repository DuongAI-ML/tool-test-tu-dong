// Web giả lập "Đơn hàng tuyển dụng" — dữ liệu chỉ nằm trong trình duyệt (sessionStorage).
// LỖI CỐ Ý để tập test: ô Số lượng nhận cả giá trị > 100 (yêu cầu chỉ cho 1–100).
const PAGE_SIZE = 10;
const SEED = [
  ['DH001', 'Lập trình viên Java', 'Phòng IT', 3, 'Đang tuyển', '2026-11-30'],
  ['DH002', 'Nhân viên kinh doanh', 'Phòng Kinh doanh', 5, 'Đang tuyển', '2026-11-15'],
  ['DH003', 'Kế toán tổng hợp', 'Phòng Kế toán', 1, 'Tạm dừng', '2026-12-01'],
  ['DH004', 'Tester (QC)', 'Phòng IT', 2, 'Đang tuyển', '2026-11-20'],
  ['DH005', 'Chuyên viên tuyển dụng', 'Phòng Nhân sự', 1, 'Đã đóng', '2026-10-01'],
  ['DH006', 'Thiết kế UI/UX', 'Phòng IT', 1, 'Đang tuyển', '2026-12-10'],
  ['DH007', 'Trưởng nhóm bán hàng', 'Phòng Kinh doanh', 1, 'Đang tuyển', '2026-11-25'],
  ['DH008', 'Thực tập sinh IT', 'Phòng IT', 4, 'Tạm dừng', '2026-12-15'],
  ['DH009', 'Nhân viên hành chính', 'Phòng Nhân sự', 2, 'Đang tuyển', '2026-11-05'],
  ['DH010', 'DevOps Engineer', 'Phòng IT', 1, 'Đang tuyển', '2026-12-20'],
  ['DH011', 'Kế toán thuế', 'Phòng Kế toán', 1, 'Đã đóng', '2026-09-30'],
  ['DH012', 'Chăm sóc khách hàng', 'Phòng Kinh doanh', 6, 'Đang tuyển', '2026-11-28'],
].map(([code, title, dept, qty, status, deadline]) => ({ code, title, dept, qty, status, deadline }));

let orders = JSON.parse(sessionStorage.getItem('hr_orders') || 'null') || SEED;
let page = 1;
const $ = (id) => document.getElementById(id);
const save = () => sessionStorage.setItem('hr_orders', JSON.stringify(orders));

// LỖI CỐ Ý 2 (phân quyền): tài khoản viewer bị ẩn nút Tạo nhưng vẫn thấy nút Xóa.
const ROLE = ((document.cookie.match(/(^|; )hr_session=([^;]+)/) || [])[2]) || 'admin';
$('user-name').textContent = ROLE === 'viewer' ? 'Viewer' : 'Admin';
if (ROLE === 'viewer') $('btn-create').hidden = true;
$('logout').onclick = (e) => { e.preventDefault(); document.cookie = 'hr_session=; path=/; max-age=0'; location.href = 'login.html'; };
$('btn-menu').onclick = () => $('side').classList.toggle('open');

function filtered() {
  const q = $('q').value.trim().toLowerCase();
  const st = $('f-status').value;
  return orders.filter((o) => (!q || o.code.toLowerCase().includes(q) || o.title.toLowerCase().includes(q))
    && (!st || o.status === st));
}

function fmt(d) { if (!d) return ''; const [y, m, dd] = d.split('-'); return `${dd}/${m}/${y}`; }

function render() {
  const list = filtered();
  const pages = Math.max(1, Math.ceil(list.length / PAGE_SIZE));
  if (page > pages) page = pages;
  const rows = list.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);
  const tb = document.querySelector('#bang-don-hang tbody');
  tb.innerHTML = '';
  for (const o of rows) {
    const tr = document.createElement('tr');
    const cls = { 'Đang tuyển': 'on', 'Tạm dừng': 'pause', 'Đã đóng': 'off', 'Nháp': 'draft', 'Chờ duyệt': 'wait' }[o.status] || '';
    tr.innerHTML = `<td data-l="Mã đơn">${o.code}</td><td data-l="Vị trí">${esc(o.title)}</td>
      <td data-l="Phòng ban">${o.dept}</td><td data-l="Số lượng">${o.qty}</td>
      <td data-l="Trạng thái"><span class="badge ${cls}">${o.status}</span></td><td data-l="Hạn nộp">${fmt(o.deadline)}</td>
      <td data-l="Thao tác"><button class="link" data-del="${o.code}">Xóa</button></td>`;
    tb.appendChild(tr);
  }
  $('empty').hidden = list.length > 0;
  const pg = $('pager');
  pg.innerHTML = '';
  if (pages > 1) {
    for (let i = 1; i <= pages; i++) {
      const b = document.createElement('button');
      b.textContent = i; b.className = 'btn' + (i === page ? ' primary' : '');
      b.setAttribute('aria-label', 'Trang ' + i);
      b.onclick = () => { page = i; render(); };
      pg.appendChild(b);
    }
  }
}

function esc(s) { return String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

document.querySelector('#bang-don-hang').addEventListener('click', (e) => {
  const code = e.target.getAttribute && e.target.getAttribute('data-del');
  if (!code) return;
  if (confirm(`Bạn có chắc muốn xóa đơn hàng ${code}?`)) {
    orders = orders.filter((o) => o.code !== code); save(); render(); toast(`Đã xóa đơn hàng ${code}`);
  }
});

$('btn-search').onclick = () => { page = 1; render(); };
$('q').addEventListener('keydown', (e) => { if (e.key === 'Enter') { page = 1; render(); } });

function openModal() {
  $('form-order').reset();
  ['e-title', 'e-qty'].forEach((i) => ($(i).textContent = ''));
  $('modal').hidden = false;
  setTimeout(() => $('o-title').focus(), 50);
}
$('btn-create').onclick = openModal;
$('btn-cancel').onclick = () => ($('modal').hidden = true);
const calc = () => { const a = $('o-score1').value, b = $('o-score2').value;
  $('o-total').value = (a === '' && b === '') ? '' : String(Number(a || 0) + Number(b || 0)); };
$('o-score1').oninput = calc; $('o-score2').oninput = calc;

$('form-order').addEventListener('submit', (e) => {
  e.preventDefault();
  const title = $('o-title').value.trim();
  const qtyRaw = $('o-qty').value;
  const qty = Number(qtyRaw);
  $('e-title').textContent = $('e-qty').textContent = '';
  let ok = true;
  if (!title) { $('e-title').textContent = 'Vui lòng nhập vị trí tuyển'; ok = false; }
  // LỖI CỐ Ý: thiếu kiểm tra qty > 100
  if (qtyRaw === '' || !Number.isInteger(qty) || qty < 1) { $('e-qty').textContent = 'Số lượng từ 1 đến 100'; ok = false; }
  if (!ok) return;
  const n = orders.reduce((m, o) => Math.max(m, parseInt(o.code.slice(2), 10)), 0) + 1;
  const code = 'DH' + String(n).padStart(3, '0');
  const draft = e.submitter && e.submitter.id === 'btn-draft';
  orders.unshift({ code, title, dept: $('o-dept').value, qty, status: draft ? 'Nháp' : 'Chờ duyệt', deadline: $('o-deadline').value });
  save();
  $('modal').hidden = true;
  page = 1; $('q').value = ''; $('f-status').value = '';
  render();
  toast(draft ? 'Đã lưu nháp đơn hàng ' + code : 'Tạo đơn hàng thành công — đã gửi phê duyệt');
});

let tt;
function toast(msg) {
  const t = $('toast'); t.textContent = msg; t.classList.add('show');
  clearTimeout(tt); tt = setTimeout(() => t.classList.remove('show'), 3000);
}

render();
