"""Chạy thử chế độ QUÉT TỪ URL trên web demo (cần tool đang chạy để phục vụ /demo).
Dùng: python tests/scan_demo.py [port] [--show]   → in ra các test case sinh ra + kết quả."""
import json
import os
import sys
import tempfile
import threading
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import scan_types as T  # noqa: E402
import scanner as S  # noqa: E402

PORT = int(next((a for a in sys.argv[1:] if a.isdigit()), "18620"))
BASE = "http://127.0.0.1:%d/demo" % PORT
types = T.default_type_opts()
for k in types:
    types[k]["on"] = True
types["compat"]["opts"]["browsers"] = ["chrome"]
types["permission"]["opts"].update({"login_mode": "auto", "user": "viewer@demo.vn", "pass": "123456",
                                    "expect": "view_only"})
types["usability"]["opts"]["compare"] = "3"
types["validation"]["opts"]["unsure"] = "later"
types["functional"]["opts"]["form_actions"] = "yes"
cfg = {"url": BASE + "/don-hang.html", "depth": 1, "max_pages": 5,
       "login": {"mode": "auto", "url": BASE + "/login.html", "user": "admin@demo.vn", "pass": "123456",
                 "remember": True},
       "types": types,
       "rules": [{"page_url": BASE + "/don-hang.html", "field_label": "Số lượng", "rule": "range", "v1": "1",
                  "v2": "100", "msg": "Số lượng từ 1 đến 100", "ac": "AC_11"},
                 {"page_url": BASE + "/don-hang.html", "field_label": "Vị trí tuyển", "rule": "required",
                  "msg": "Vui lòng nhập vị trí tuyển"},
                 {"page_url": BASE + "/don-hang.html", "field_label": "Tổng điểm", "rule": "calc",
                  "v1": "Điểm ngành nghề=8; Điểm kinh nghiệm=7", "v2": "15", "ac": "AC_15"}],
       "state_rules": [{"page_url": BASE + "/don-hang.html", "form_name": "Tạo đơn hàng", "button": "Lưu nháp",
                        "expect": "Nháp"},
                       {"page_url": BASE + "/don-hang.html", "form_name": "Tạo đơn hàng",
                        "button": "Lưu & Gửi phê duyệt", "expect": "Chờ duyệt"}],
       "perm": {"role": "Viewer", "rows": [
           {"page_url": BASE + "/don-hang.html", "page_title": "Đơn hàng tuyển dụng", "button": "__page__",
            "expect": "allow"},
           {"page_url": BASE + "/don-hang.html", "page_title": "Đơn hàng tuyển dụng", "button": "+ Tạo đơn hàng",
            "expect": "deny"},
           {"page_url": BASE + "/don-hang.html", "page_title": "Đơn hàng tuyển dụng", "button": "Xóa",
            "expect": "deny"},
           {"page_url": BASE + "/don-hang.html", "page_title": "Đơn hàng tuyển dụng", "button": "Tìm kiếm",
            "expect": "allow"}]}}
run_dir = tempfile.mkdtemp(prefix="scan_")
data_dir = tempfile.mkdtemp(prefix="scan_data_")
opts = {"browser": "edge", "headless": "--show" not in sys.argv, "speed": "nhanh", "timeout": 6, "tester": "Demo"}
sc = S.Scanner(cfg, opts, run_dir, data_dir)
sc.start()
seen = 0
t0 = time.time()
while sc.thread.is_alive():
    for e in sc.events_since(seen):
        seen = e["seq"]
        if e["kind"] in ("info", "warn", "error"):
            print("  [%s] %s" % (e["kind"], e.get("msg")))
    m = sc.manual_req
    if m:
        print("  [hỏi người] %s — %s" % (m.get("label"), m.get("reason")))
        sc.answer_manual({"req_id": m["req_id"], "statuses": ["Pass"] * len(m.get("steps") or []),
                          "note": "tự trả lời"})
    time.sleep(0.3)
for e in sc.events_since(seen):
    if e["kind"] in ("info", "warn", "error"):
        print("  [%s] %s" % (e["kind"], e.get("msg")))
print("Xong sau %.0fs — %d TC, %d scenario. Thư mục: %s" % (time.time() - t0, len(sc.results),
                                                            len(sc.wb["scenarios"]), run_dir))
cnt = {}
for r in sc.results:
    cnt[r["status"]] = cnt.get(r["status"], 0) + 1
    bad = [s for s in r["steps"] if s["status"] not in ("Pass", "Skip")]
    print("%-8s %-10s %-62s %s" % (r["tc_id"], r["status"], r["title"][:62],
                                   ("| b%s: %s" % (bad[0]["no"], bad[0]["actual"][:110])) if bad else ""))
print(cnt)
with open(os.path.join(run_dir, "result.json"), encoding="utf-8") as f:
    d = json.load(f)
print("discovery forms:", [(p["url"][-20:], [(f["name"], f["opener"], len(f["fields"])) for f in p["forms"]])
                           for p in d["discovery"]])
