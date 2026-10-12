"""Dò web demo 1 lần, rồi xem tool sinh / ghép TEST SCENARIO thế nào:
  1) Tool tự sinh scenario theo màn hình (cột Nhóm như file tester)
  2) Lấy đúng 22 scenario trong file của tester → mỗi scenario được ghép với bài kiểm tra nào
Dùng: python tests/scenario_map.py [port] [--run]   (--run = chạy thật các scenario của tester)"""
import os
import sys
import tempfile
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import scan_types as T  # noqa: E402
import scanner as S  # noqa: E402

PORT = int(next((a for a in sys.argv[1:] if a.isdigit()), "18620"))
BASE = "http://127.0.0.1:%d/demo" % PORT
TESTER = [  # đúng ảnh Test Scenario tester gửi
    ("TS_DHTD_001", "usability", "Kiểm tra giao diện và bố cục mặc định của trang Đơn hàng tuyển dụng", "AC_01", "Medium"),
    ("TS_DHTD_002", "usability", "Kiểm tra trang không phát sinh lỗi JavaScript khi thay đổi kích thước màn hình", "AC_02", "High"),
    ("TS_DHTD_003", "validation", "Kiểm tra dữ liệu hiển thị trên danh sách được mapping đúng với dữ liệu trả về", "AC_03", "High"),
    ("TS_DHTD_004", "validation", "Kiểm tra validation dữ liệu đầu vào của chức năng tìm kiếm", "AC_04", "Medium"),
    ("TS_DHTD_005", "functional", "Kiểm tra chức năng tìm kiếm đơn hàng tuyển dụng theo từ khóa", "AC_05", "High"),
    ("TS_DHTD_006", "functional", "Kiểm tra các chức năng thao tác tại cột HÀNH ĐỘNG", "AC_06", "High"),
    ("TS_DHTD_007", "functional", "Kiểm tra chức năng phân trang danh sách đơn hàng tuyển dụng", "AC_07", "Medium"),
    ("TS_DHTD_008", "functional", "Kiểm tra chức năng sắp xếp dữ liệu theo các cột", "AC_07", "Medium"),
    ("TS_DHTD_009", "business rule", "Kiểm tra quy tắc xác định và hiển thị trạng thái đơn hàng, trạng thái phê duyệt", "AC_08", "High"),
    ("TS_DHTD_010", "state", "Kiểm tra trạng thái mở/đóng của bộ lọc nâng cao và menu Nhập dữ liệu", "AC_09", "Low"),
    ("TS_DHTD_011", "permission", "Kiểm tra quyền truy cập trang bằng tài khoản Viewer", "AC_10", "High"),
    ("TS_DHTD_012", "functional", "Kiểm tra chức năng click vào ô xem button hoạt động không.", "AC_11", "High"),
    ("TS_DHTD_014", "functional", "Kiểm tra luồng vòng đời đơn hàng: Từ khi tạo mới -> Phê duyệt -> Trạng thái 'Hoạt động' -> Ứng viên có thể nhìn thấy", "AC_13", "High"),
    ("TS_DHTD_015", "business rule", "Kiểm tra sự thay đổi trạng thái tự động thành 'Hết hạn' khi thời gian hệ thống vượt qua 'Ngày đóng'", "AC_14", "High"),
    ("TS_DHTD_016", "business rule", "Kiểm tra ràng buộc không thể gắn ứng viên cho một đơn hàng có 'Trạng thái phê duyệt' khác Đã duyệt hoặc trạng thái 'Hết hạn'", "AC_15", "High"),
    ("TS_DHTD_017", "validation", "Kiểm tra validation bỏ trống các trường bắt buộc (*) trên form Tạo mới đơn hàng tuyển dụng (Thông tin cơ bản)", "AC_11", "High"),
    ("TS_DHTD_018", "business rule", "Kiểm tra quy tắc logic thời gian (Ngày phỏng vấn, Ngày đóng, Ngày dự kiến xuất cảnh) trên form Tạo mới", "AC_11", "High"),
    ("TS_DHTD_019", "function", "Kiểm tra trạng thái và chức năng của nút toggle 'Ẩn bài viết' khi tạo mới đơn hàng tuyển dụng", "AC_11", "Medium"),
    ("TS_DHTD_020", "validation", "Kiểm tra validation khi click\"Nhập dữ liệu\" chưa chọn tệp hoặc tải lên tệp sai định dạng (không phải Excel, JSON)", "AC_12", "High"),
    ("TS_DHTD_021", "function", "Kiểm tra chức năng tải tệp lên và nhập dữ liệu đơn hàng thành công từ file hợp lệ (.xlsx, .json)", "AC_12", "High"),
    ("TS_DHTD_022", "usability", "Kiểm tra đóng popup 'Nhập liệu dữ liệu' khi bấm nút Đóng hoặc biểu tượng X", "AC_12", "Low"),
]
types = T.default_type_opts()
types["permission"]["opts"].update({"login_mode": "auto", "user": "viewer@demo.vn", "pass": "123456"})
types["validation"]["opts"]["unsure"] = "later"
base_cfg = {"url": BASE + "/don-hang.html", "depth": 1, "max_pages": 5, "types": types,
            "login": {"mode": "auto", "url": BASE + "/login.html", "user": "admin@demo.vn", "pass": "123456",
                      "remember": True},
            "perm": {"role": "Viewer", "rows": []}}
data_dir = tempfile.mkdtemp(prefix="scn_data_")
opts = {"browser": "edge", "headless": True, "speed": "nhanh", "timeout": 6, "tester": "Demo"}


def run(cfg, label):
    rd = tempfile.mkdtemp(prefix="scn_")
    sc = S.Scanner(cfg, opts, rd, data_dir)
    sc.start()
    while sc.thread.is_alive():
        m = sc.manual_req
        if m:
            sc.answer_manual({"req_id": m["req_id"], "statuses": ["Pass"] * len(m.get("steps") or []), "note": "auto"})
        time.sleep(0.3)
    print("== %s: %d TC" % (label, len(sc.results)))
    return sc, rd


d, rd = run(dict(base_cfg, discover_only=True), "Dò trang")
pages_file = os.path.join(rd, "pages.json")
print("pages.json:", os.path.exists(pages_file))

print("\n######## 1) TOOL TỰ SINH TEST SCENARIO")
pv = S.Scanner(dict(base_cfg), opts, tempfile.mkdtemp(), data_dir)
import json  # noqa: E402
pv.pages = json.load(open(pages_file, encoding="utf-8"))
pv.logged_in = True
for x in pv.plan_preview():
    print("%-12s %-14s %-95s %2d TC  %s" % (x["id"], x["nhom"], x["scenario"][:95], len(x["tcs"]), x["mode"]))

print("\n######## 2) LẤY 22 SCENARIO CỦA TESTER")
cfg2 = dict(base_cfg, scenarios_in=[{"id": a, "nhom": b, "scenario": c, "ac": d_, "priority": e}
                                    for a, b, c, d_, e in TESTER])
pv2 = S.Scanner(cfg2, opts, tempfile.mkdtemp(), data_dir)
pv2.pages = pv.pages
pv2.logged_in = True
for x in pv2.plan_preview():
    print("%-12s %-14s %-70s → %2d TC [%s] %s" % (x["id"], x["nhom"], x["scenario"][:70], len(x["tcs"]), x["mode"],
                                                 " | ".join(t[:40] for t in x["tcs"][:3])))
if "--run" in sys.argv:
    sc, rd2 = run(dict(cfg2, reuse_pages=pages_file), "Chạy scenario của tester")
    for r in sc.results:
        bad = [s for s in r["steps"] if s["status"] not in ("Pass", "Skip")]
        print("%-14s %-12s %-10s %-55s %s" % (r["tc_id"], r["scenario_id"], r["status"], r["title"][:55],
                                             (bad[0]["actual"][:80] if bad else "")))
    print("Thư mục:", rd2)
