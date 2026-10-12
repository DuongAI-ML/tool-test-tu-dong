"""Chạy thử toàn bộ luồng qua API: tạo file mẫu có ví dụ → chạy tất cả TC (ẩn trình duyệt) → tự trả lời các bước
test tay → xuất Excel + HTML + ghi ngược. Dùng: python tests/smoke_api.py [port]"""
import json
import os
import sys
import tempfile
import time
import urllib.request

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 18620
BASE = "http://127.0.0.1:%d" % PORT


def call(path, body=None):
    req = urllib.request.Request(BASE + path, data=json.dumps(body).encode() if body is not None else None,
                                 headers={"X-Tool": "1", "Content-Type": "application/json"},
                                 method="POST" if body is not None else "GET")
    for attempt in range(20):
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            return {"http_error": e.code, **json.loads(e.read().decode("utf-8"))}
        except urllib.error.URLError:
            time.sleep(1)   # máy hết cổng tạm thời (chương trình khác mở quá nhiều kết nối) → thử lại
    raise RuntimeError("Không kết nối được tool")


out = tempfile.mkdtemp(prefix="tooltest_")
xlsx = os.path.join(out, "TestCase_Mau_CoViDu.xlsx")
r = call("/api/template", {"sample": True, "path": xlsx})
wb = r["workbook"]
print("Mở file:", wb["file_name"], len(wb["testcases"]), "TC", wb["warnings"])
call("/api/settings", {"headless": True, "speed": "nhanh", "tester": "Tester Demo", "timeout": 5,
                       "export": {"folder": out, "open_after": False, "writeback": True, "html": True,
                                  "embed": True}})
sel = [t["key"] for t in wb["testcases"]]
r = call("/api/run", {"selection": sel})
print("Chạy:", r)
since, answered = 0, set()
t0 = time.time()
while True:
    p = call("/api/run/poll?since=%d" % since)
    for e in p["events"]:
        since = e["seq"]
        if e["kind"] in ("tc_done", "warn", "error", "manual"):
            print("  ", e["kind"], e.get("label", ""), e.get("status", ""), e.get("msg", ""), e.get("reason", ""))
    run = p["run"]
    m = run.get("manual")
    if m and m["req_id"] not in answered:
        answered.add(m["req_id"])
        sts = ["Pass"] * len(m["steps"])
        if m["tc_id"] == "TC_DHTD_001_02":
            sts[-1] = "Fail"
        call("/api/run/manual", {"req_id": m["req_id"], "statuses": sts, "note": "Tự trả lời (smoke test)"})
    if run["state"] in ("xong", "đã dừng", "lỗi"):
        break
    if time.time() - t0 > 600:
        print("QUÁ GIỜ")
        call("/api/run/stop", {})
        break
    time.sleep(0.3)
print("Kết quả:", run["counts"], "trong %.0fs" % (time.time() - t0))
res = call("/api/run/result?run_id=" + run["run_id"])
for x in res["results"]:
    bad = [s for s in x["steps"] if s["status"] not in ("Pass",)]
    print("  %-28s %-10s %s" % (x["label"], x["status"], (bad[0]["no"], bad[0]["actual"][:150]) if bad else ""))
e = call("/api/export", {"run_id": run["run_id"], "pattern": "KetQua_{du_an}_{ngay}_{moi_truong}_lan{lan_chay}",
                         "overwrite": "suffix"})
print("Xuất:", e)
e2 = call("/api/export", {"run_id": run["run_id"], "pattern": "KetQua_{du_an}_{ngay}_{moi_truong}_lan{lan_chay}"})
print("Xuất lại (phải hỏi trùng tên):", e2)
