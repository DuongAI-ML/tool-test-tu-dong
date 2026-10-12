"""Bộ chạy test: điều khiển Edge/Chrome có sẵn trên máy bằng Playwright, chạy từng bước theo từ khóa.

Chạy trong 1 luồng riêng (Playwright sync API phải dùng trong cùng 1 luồng). Giao diện hỏi trạng thái qua
Runner.snapshot() / Runner.events_since(); bước test tay được trả lời qua Runner.answer_manual().
"""
import datetime as dt
import json
import os
import random
import re
import shutil
import threading
import time
import traceback

import columns as C
import excel_io as X
import keywords as K

STATUS_ORDER = ["Fail", "Blocked", "Chưa test", "Pass", "Skip"]
SPEEDS = {"nhanh": 0, "vua": 250, "cham": 700}
HINTS = [  # (từ đầu trong tên đối tượng (đã bỏ dấu), vai trò ưu tiên)
    ("nut bam", "button"), ("nut", "button"), ("button", "button"), ("btn", "button"),
    ("lien ket", "link"), ("duong link", "link"), ("link", "link"),
    ("o nhap", "textbox"), ("o tim kiem", "searchbox"), ("truong", "textbox"), ("field", "textbox"),
    ("input", "textbox"), ("textbox", "textbox"), ("o", "textbox"),
    ("dropdown", "combobox"), ("combobox", "combobox"), ("danh sach chon", "combobox"), ("select", "combobox"),
    ("hop chon", "combobox"),
    ("checkbox", "checkbox"), ("o tick", "checkbox"), ("o chon", "checkbox"), ("radio", "radio"),
    ("tab", "tab"), ("menu", "menuitem"), ("tieu de", "heading"), ("heading", "heading"),
    ("anh", "img"), ("hinh", "img"), ("image", "img"), ("bang", "table"),
    ("dong chu", "text"), ("chu", "text"), ("text", "text"), ("thong bao", "text"), ("nhan", "text"),
]
QUOTES = r"[\"“”'‘’«»]"


class StepFail(Exception):
    pass


class StepBlocked(Exception):
    pass


class Stopped(Exception):
    pass


def slug(s, maxlen=60):
    s = C.norm(s).replace(" ", "-")
    return (s or "x")[:maxlen]


def tc_status(statuses):
    st = [s for s in statuses if s]
    if not st:
        return "Chưa test"
    for s in ("Fail", "Blocked", "Chưa test"):
        if s in st:
            return s
    return "Pass" if "Pass" in st else "Chưa test"


class Runner:
    def __init__(self, wb, selection, options, run_dir, data_dir):
        self.wb = wb
        self.tc_by_key = {t["key"]: t for t in wb["testcases"]}
        self.tc_by_id = {}
        for t in wb["testcases"]:
            self.tc_by_id.setdefault(t["id"], t)
        self.selection = [k for k in selection if k in self.tc_by_key]
        self.opt = options
        self.run_dir = run_dir
        self.img_dir = os.path.join(run_dir, "anh")
        os.makedirs(self.img_dir, exist_ok=True)
        self.base_dir = os.path.join(data_dir, "anh-chuan", slug(wb.get("project", "du-an")))
        self.timeout = max(1.0, float(options.get("timeout") or 10))
        self.lock = threading.Lock()
        self.events = []
        self.state = "chuẩn bị"
        self.stop_flag = threading.Event()
        self.manual_req = None
        self.manual_ans = None
        self.manual_evt = threading.Event()
        self.results = []
        self.current = None
        self.started = dt.datetime.now()
        self.finished = None
        self.thread = None
        self.pw = None
        self.browsers = {}
        self.warn_browser = set()
        self.instances = self._expand()

    # ------------------------------------------------------------------ chuẩn bị danh sách lượt chạy
    def _expand(self):
        out = []
        for key in self.selection:
            tc = self.tc_by_key[key]
            sizes = X.parse_sizes(tc["fields"].get("kich_thuoc", "")) or [None]
            browsers = X.parse_browsers(tc["fields"].get("trinh_duyet", "")) or [self.opt.get("browser", "edge")]
            for b in browsers:
                for s in sizes:
                    label = tc["id"]
                    if len(sizes) > 1 or s:
                        label += " @%dx%d" % s if s else ""
                    if len(browsers) > 1:
                        label += " [%s]" % b
                    out.append({"tc": tc, "size": s, "browser": b, "label": label})
        if self.opt.get("auto_first", True):
            out.sort(key=lambda i: X.MODE_RANK.get(i["tc"]["mode"], 0))
        for n, i in enumerate(out):
            i["n"] = n
        return out

    # ------------------------------------------------------------------ giao tiếp với giao diện
    def emit(self, kind, **data):
        with self.lock:
            data.update({"seq": len(self.events) + 1, "kind": kind, "t": time.strftime("%H:%M:%S")})
            self.events.append(data)

    def events_since(self, seq):
        with self.lock:
            return [e for e in self.events if e["seq"] > seq]

    def snapshot(self):
        with self.lock:
            done = len(self.results)
            counts = {}
            for r in self.results:
                counts[r["status"]] = counts.get(r["status"], 0) + 1
            return {"state": self.state, "total": len(self.instances), "done": done, "counts": counts,
                    "current": self.current, "manual": self.manual_req, "run_dir": self.run_dir,
                    "run_id": os.path.basename(self.run_dir), "seq": len(self.events),
                    "started": self.started.strftime("%d/%m/%Y %H:%M:%S"),
                    "finished": self.finished.strftime("%d/%m/%Y %H:%M:%S") if self.finished else None}

    def answer_manual(self, ans):
        with self.lock:
            if not self.manual_req or ans.get("req_id") != self.manual_req["req_id"]:
                return False
            self.manual_ans = ans
        self.manual_evt.set()
        return True

    def stop(self):
        self.stop_flag.set()
        self.manual_evt.set()

    def start(self):
        self.thread = threading.Thread(target=self._run_all, daemon=True, name="runner")
        self.thread.start()

    # ------------------------------------------------------------------ chạy
    def _run_all(self):
        self.state = "đang chạy"
        self.emit("start", total=len(self.instances))
        try:
            from playwright.sync_api import sync_playwright
            self.pw = sync_playwright().start()
        except Exception as e:
            self.state = "lỗi"
            self.emit("error", msg="Không khởi động được bộ điều khiển trình duyệt: %s" % e)
            self._finish()
            return
        try:
            for inst in self.instances:
                if self.stop_flag.is_set():
                    self._record_not_run(inst, "Đã bấm Dừng")
                    continue
                self._run_instance(inst)
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

    def _finish(self):
        self.finished = dt.datetime.now()
        if self.state != "lỗi":
            self.state = "đã dừng" if self.stop_flag.is_set() else "xong"
        self.current = None
        self.save()
        self.emit("finish", state=self.state)

    def save(self):
        data = self.result_data()
        tmp = os.path.join(self.run_dir, "result.json.tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        os.replace(tmp, os.path.join(self.run_dir, "result.json"))

    def result_data(self):
        return {"run_id": os.path.basename(self.run_dir), "project": self.wb.get("project"),
                "file": self.wb.get("file"), "options": self.opt, "env": self.opt.get("env"),
                "started": self.started.isoformat(timespec="seconds"),
                "finished": self.finished.isoformat(timespec="seconds") if self.finished else None,
                "state": self.state, "results": self.results,
                "scenarios": self.wb.get("scenarios", []), "headers": self.wb.get("headers", {})}

    def _record_not_run(self, inst, why):
        tc = inst["tc"]
        res = self._new_result(inst)
        for s in tc["steps"]:
            res["steps"].append(self._step_res(s, "Chưa test", why))
        res["status"] = "Chưa test"
        res["note"] = why
        with self.lock:
            self.results.append(res)
        self.emit("tc_done", label=inst["label"], status="Chưa test", key=tc["key"])

    def _new_result(self, inst):
        tc = inst["tc"]
        return {"key": tc["key"], "tc_id": tc["id"], "label": inst["label"], "title": tc["fields"].get("tieu_de", ""),
                "scenario_id": tc["fields"].get("scenario_id", ""), "mode": tc["mode"],
                "mode_reason": tc.get("mode_reason", ""), "fields": tc["fields"], "extra": tc.get("extra", {}),
                "size": "%dx%d" % inst["size"] if inst["size"] else "", "browser": inst["browser"],
                "env": self.opt.get("env", ""), "steps": [], "status": "", "note": "", "start": "", "duration": 0,
                "video": "", "tester": self.opt.get("tester", ""), "row": tc.get("row"),
                "steps_merged": tc.get("steps_merged", False)}

    def _step_res(self, s, status="", actual=""):
        return {"no": s["no"], "fields": s["fields"], "extra": s.get("extra", {}), "action": s.get("action", ""),
                "kind": s.get("kind", ""), "status": status, "actual": actual, "images": [], "duration": 0,
                "row": s.get("row")}

    def _launch(self, name):
        if name in self.browsers:
            return self.browsers[name]
        headless = bool(self.opt.get("headless"))
        slow = SPEEDS.get(self.opt.get("speed", "vua"), 250)
        tries = {"edge": ["msedge", "chrome", None], "chrome": ["chrome", "msedge", None],
                 "firefox": ["firefox"]}.get(name, ["msedge", "chrome", None])
        errs = []
        for ch in tries:
            try:
                args = {"headless": headless, "slow_mo": slow}
                if ch == "firefox":
                    b = self.pw.firefox.launch(**args)
                elif ch:
                    b = self.pw.chromium.launch(channel=ch, args=["--start-maximized"] if not headless else [],
                                                **args)
                else:
                    b = self.pw.chromium.launch(**args)
                if ch != tries[0] and name not in self.warn_browser:
                    self.warn_browser.add(name)
                    self.emit("warn", msg="Không mở được %s → dùng %s thay thế." % (name, ch or "Chromium"))
                self.browsers[name] = b
                return b
            except Exception as e:
                errs.append("%s: %s" % (ch or "chromium", str(e).splitlines()[0][:200]))
        raise StepBlocked("Không mở được trình duyệt %s. Máy cần có Microsoft Edge hoặc Google Chrome.\n%s"
                          % (name, "\n".join(errs)))

    def _vars_for(self, inst):
        env_name = inst["tc"]["fields"].get("moi_truong", "")
        if env_name not in self.wb["envs"]:
            env_name = self.opt.get("env") or (self.wb["env_names"][0] if self.wb["env_names"] else "")
        v = {"demo_url": self.opt.get("demo_url", ""), "hom_nay": dt.date.today().strftime("%d/%m/%Y"),
             "ngay": dt.date.today().strftime("%d/%m/%Y"), "ngau_nhien": str(random.randint(10000, 99999)),
             "moi_truong": env_name}
        for k, val in self.wb["envs"].get(env_name, {}).items():
            v[k] = val
        if self.opt.get("base_url"):        # ô "URL web cần test" trên giao diện ghi đè ${base_url} của sheet Env
            v["base_url"] = self.opt["base_url"].rstrip("/")
        v.update(X.parse_test_data(inst["tc"]["fields"].get("test_data", "")))
        # biến lồng nhau: base_url = ${demo_url}
        for _ in range(3):
            for k in list(v):
                v[k] = self._subst(v[k], v)
        return v

    @staticmethod
    def _subst(text, v):
        if not text or "${" not in text:
            return text

        def rep(m):
            name = m.group(1).strip()
            return str(v[name]) if name in v else m.group(0)
        return re.sub(r"\$\{([^}]+)\}", rep, text)

    def _run_instance(self, inst):
        tc = inst["tc"]
        res = self._new_result(inst)
        res["start"] = dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        t0 = time.time()
        with self.lock:
            self.current = {"label": inst["label"], "title": res["title"], "mode": tc["mode"], "step": "",
                            "n": inst["n"] + 1}
        self.emit("tc_start", label=inst["label"], title=res["title"], mode=tc["mode"], key=tc["key"])
        ctx = page = None
        self.vars = self._vars_for(inst)
        self.dialogs = []
        self.dismiss_next = False
        self.console_errors = []
        self.inst = inst
        self.res = res
        try:
            browser = self._launch(inst["browser"])
            size = inst["size"] or X.size_of(self.opt.get("size", "")) or (1366, 768)
            kw = {"viewport": {"width": size[0], "height": size[1]}, "ignore_https_errors": True, "locale": "vi-VN"}
            if self.opt.get("video"):
                kw["record_video_dir"] = os.path.join(self.run_dir, "video")
                kw["record_video_size"] = {"width": size[0], "height": size[1]}
            ctx = browser.new_context(**kw)
            ctx.set_default_timeout(self.timeout * 1000)
            ctx.set_default_navigation_timeout(max(30, self.timeout * 3) * 1000)
            self._hooked = set()
            ctx.on("page", self._hook_page)
            self.ctx = ctx
            page = ctx.new_page()
            self._hook_page(page)
            self.page = page
            if tc["mode"] == "Tay":
                self._run_manual_tc(inst, res)
            else:
                self._run_steps(inst, res)
        except Stopped:
            for s in tc["steps"][len(res["steps"]):]:
                res["steps"].append(self._step_res(s, "Chưa test", "Đã bấm Dừng"))
            res["note"] = "Đã bấm Dừng giữa chừng"
        except StepBlocked as e:
            done = len(res["steps"])
            for s in tc["steps"][done:]:
                res["steps"].append(self._step_res(s, "Blocked", str(e)))
            res["note"] = str(e)
            self.emit("warn", msg="%s: %s" % (inst["label"], e))
        except Exception as e:
            res["note"] = "Lỗi tool: %s" % e
            self.emit("error", msg="%s: lỗi không mong muốn %s" % (inst["label"], e), detail=traceback.format_exc())
            for s in tc["steps"][len(res["steps"]):]:
                res["steps"].append(self._step_res(s, "Blocked", "Lỗi tool: %s" % e))
        finally:
            video = None
            try:
                if page and self.opt.get("video"):
                    video = page.video
            except Exception:
                video = None
            try:
                if ctx:
                    ctx.close()
            except Exception:
                pass
            if video:
                try:
                    src = video.path()
                    dst = os.path.join(self.run_dir, "video", "%03d_%s.webm" % (inst["n"] + 1, slug(inst["label"])))
                    if os.path.exists(src):
                        os.replace(src, dst)
                        res["video"] = os.path.relpath(dst, self.run_dir).replace("\\", "/")
                except Exception:
                    pass
        res["duration"] = round(time.time() - t0, 1)
        res["status"] = tc_status([s["status"] for s in res["steps"]])
        if not res["steps"]:
            res["status"] = "Chưa test"
            res["note"] = res["note"] or "TC không có bước nào"
        with self.lock:
            self.results.append(res)
        self.save()
        self.emit("tc_done", label=inst["label"], status=res["status"], key=tc["key"], duration=res["duration"])

    def _hook_page(self, page):
        if id(page) in self._hooked:
            return
        self._hooked.add(id(page))
        page.on("dialog", self._on_dialog)
        page.on("console", lambda m: self.console_errors.append(m.text) if m.type == "error" else None)
        page.on("pageerror", lambda e: self.console_errors.append(str(e)))

    def _on_dialog(self, d):
        self.dialogs.append(d.message)
        try:
            if self.dismiss_next:
                self.dismiss_next = False
                d.dismiss()
            else:
                d.accept()
        except Exception:
            pass

    def _pump(self, ms):
        """Chờ nhưng vẫn để Playwright xử lý sự kiện (time.sleep sẽ làm page.url, hộp thoại… không cập nhật)."""
        try:
            self.page.wait_for_timeout(ms)
        except Exception:
            time.sleep(ms / 1000.0)

    def _check_stop(self):
        if self.stop_flag.is_set():
            raise Stopped()

    def _run_steps(self, inst, res):
        tc = inst["tc"]
        failed = False
        for s in tc["steps"]:
            self._check_stop()
            sr = self._step_res(s)
            res["steps"].append(sr)
            with self.lock:
                if self.current:
                    self.current["step"] = s["no"]
            t0 = time.time()
            if failed and self.opt.get("stop_on_fail", True):
                sr["status"], sr["actual"] = "Skip", "Không chạy vì bước trước bị lỗi"
                continue
            if s["kind"] == "invalid":
                sr["status"] = "Fail"
                sr["actual"] = "Hành động '%s' không có trong danh sách. Xem sheet HuongDan." % s["fields"].get(
                    "hanh_dong")
                failed = True
                continue
            if s["kind"] == "manual":
                if self.page.url != "about:blank":
                    self._shot(sr, "duyet")
                self._ask_manual(inst, res, [sr], reason="Bước %s cần bạn kiểm tra" % s["no"])
                if sr["status"] in ("Fail", "Blocked"):
                    failed = True
                continue
            try:
                actual = self._exec(s["action"], s["fields"], sr, depth=0)
                sr["status"] = "Pass"
                sr["actual"] = actual or "Đúng như mong đợi"
                if self.opt.get("shot_every_step"):
                    self._shot(sr, "b")
            except (StepFail, StepBlocked) as e:
                sr["status"] = "Fail"
                sr["actual"] = str(e)
                failed = True
                self._shot(sr, "loi")
            except Stopped:
                sr["status"], sr["actual"] = "Chưa test", "Đã bấm Dừng"
                raise
            except Exception as e:
                sr["status"] = "Fail"
                sr["actual"] = self._friendly(e, s["fields"])
                failed = True
                self._shot(sr, "loi")
            sr["duration"] = round(time.time() - t0, 2)
            self.emit("step", label=inst["label"], no=s["no"], status=sr["status"], actual=sr["actual"])
        if tc.get("review_at_end") and not self.stop_flag.is_set():
            last = {"no": "Duyệt", "fields": {"buoc": "Duyệt kết quả các bước", "expected":
                                              "Kết quả trên màn hình đúng như mong đợi"}, "kind": "manual",
                    "action": "chụp_và_duyệt"}
            sr = self._step_res(last)
            self._shot(sr, "duyet")
            res["steps"].append(sr)
            self._ask_manual(inst, res, [sr], reason="TC Bán tự động: tool đã chạy xong các bước, bạn duyệt lại")

    def _run_manual_tc(self, inst, res):
        tc = inst["tc"]
        url = self.vars.get("base_url", "")
        opened = ""
        if url and "${" not in url:
            try:
                self.page.goto(url)
                opened = "" if self.opt.get("headless") else url
            except Exception:
                opened = ""
        srs = [self._step_res(s) for s in tc["steps"]]
        res["steps"].extend(srs)
        if not srs:
            return
        if self.page.url != "about:blank":
            self._shot(srs[0], "tay")
        self._ask_manual(inst, res, srs, reason="TC TEST TAY (%s)" % (tc.get("mode_reason") or "người đánh giá"),
                         opened=opened)

    def _ask_manual(self, inst, res, srs, reason, opened=""):
        if self.stop_flag.is_set():
            raise Stopped()
        tc = inst["tc"]
        req_id = "%d-%d" % (inst["n"], int(time.time() * 1000))
        req = {"req_id": req_id, "label": inst["label"], "tc_id": tc["id"], "title": res["title"],
               "mode": tc["mode"], "reason": reason, "opened": opened,
               "precondition": tc["fields"].get("precondition", ""), "test_data": tc["fields"].get("test_data", ""),
               "steps": [{"no": sr["no"], "buoc": sr["fields"].get("buoc", ""),
                          "expected": sr["fields"].get("expected", ""), "data": sr["fields"].get("du_lieu_buoc", "")
                          or sr["fields"].get("gia_tri", ""), "images": list(sr["images"])} for sr in srs],
               "run_id": os.path.basename(self.run_dir)}
        self._banner(True)
        with self.lock:
            self.manual_ans = None
            self.manual_req = req
            self.state = "chờ bạn kiểm tra"
        self.manual_evt.clear()
        self.emit("manual", label=inst["label"], reason=reason)
        while not self.manual_evt.wait(0.2):
            self._pump(300)   # vẫn xử lý sự kiện trình duyệt (hộp thoại, tab mới) trong lúc chờ người test
        with self.lock:
            ans = self.manual_ans
            self.manual_req = None
            if self.state == "chờ bạn kiểm tra":
                self.state = "đang chạy"
        self._banner(False)
        if not ans:
            for sr in srs:
                sr["status"], sr["actual"] = "Chưa test", "Đã bấm Dừng trước khi kiểm tra"
            raise Stopped()
        statuses = ans.get("statuses") or []
        note = (ans.get("note") or "").strip()
        for i, sr in enumerate(srs):
            st = statuses[i] if i < len(statuses) else ""
            if st not in ("Pass", "Fail", "Blocked", "Chưa test"):
                st = ans.get("all") if ans.get("all") in ("Pass", "Fail", "Blocked", "Chưa test") else "Chưa test"
            sr["status"] = st
            per = (ans.get("notes") or [])
            sr["actual"] = (per[i] if i < len(per) and per[i] else "") or note or (
                "Người test xác nhận: " + st if st != "Chưa test" else "Chưa kiểm tra (để sau)")
            sr["manual"] = True
        for img in ans.get("images") or []:
            if srs and img:
                srs[-1]["images"].append(img)
        self.emit("manual_done", label=inst["label"], statuses=[sr["status"] for sr in srs])

    def _banner(self, on, text="⏸ Tool Test đang chờ bạn chọn Pass/Fail — xem cửa sổ Tool Test"):
        """Dải chữ nhỏ ở góc trang đang test, nhắc người dùng xem cửa sổ tool (không có trong ảnh chụp)."""
        try:
            if on:
                self.page.evaluate("""(t) => { if (document.getElementById('__autotest_banner')) return;
                  const d = document.createElement('div'); d.id='__autotest_banner';
                  d.textContent = t;
                  d.style.cssText = 'position:fixed;left:8px;bottom:8px;z-index:2147483647;background:#c2410c;'
                    + 'color:#fff;font:600 13px Segoe UI,Arial;padding:6px 10px;border-radius:6px;'
                    + 'box-shadow:0 2px 8px rgba(0,0,0,.3);pointer-events:none;opacity:.92';
                  document.documentElement.appendChild(d); }""", text)
            else:
                self.page.evaluate("() => { const d=document.getElementById('__autotest_banner'); d && d.remove(); }")
        except Exception:
            pass

    def _shot(self, sr, tag, full=False, locator=None):
        try:
            self._banner(False)
            name = "%03d_%s_b%s_%s_%d.png" % (self.inst["n"] + 1, slug(self.inst["label"], 40), slug(str(sr["no"]), 8),
                                              tag, len(sr["images"]) + 1)
            path = os.path.join(self.img_dir, name)
            if locator is not None:
                locator.screenshot(path=path)
            else:
                self.page.screenshot(path=path, full_page=full)
            rel = "anh/" + name
            sr["images"].append(rel)
            return path
        except Exception:
            return None

    # ------------------------------------------------------------------ tìm phần tử
    def _parse_target(self, target):
        t = (target or "").strip()
        nth = None
        m = re.search(r"\s#(\d+)\s*$", t)
        if m:
            nth = int(m.group(1)) - 1
            t = t[:m.start()].strip()
        el = self.wb.get("elements", {})
        if t in el:
            return {"selector": el[t], "nth": nth, "raw": t}
        low = {k.lower(): v for k, v in el.items()}
        if t.lower() in low:
            return {"selector": low[t.lower()], "nth": nth, "raw": t}
        if re.match(r"^(css=|xpath=|text=|id=|role=|internal:)", t) or re.match(r"^(#[\w-]|\.[\w-]|\[|//|\(//|"
                                                                               r"[a-z]+[#.\[][\w-])", t):
            if t.startswith("id="):
                t = "#" + t[3:]
            return {"selector": t, "nth": nth, "raw": t}
        hint, name = None, t
        qm = re.search(QUOTES + r"(.+?)" + QUOTES, t)
        if qm:
            name = qm.group(1).strip()
            before = C.norm(t[:qm.start()])
            for h, role in HINTS:
                if before == h or before.endswith(" " + h) or before.startswith(h + " ") or before == h:
                    hint = role
                    break
        else:
            n = C.norm(t)
            for h, role in HINTS:
                if n.startswith(h + " "):
                    hint = role
                    words = len(h.split())
                    name = " ".join(t.split()[words:]).strip() or t
                    break
        return {"name": name, "hint": hint, "nth": nth, "raw": t}

    def _strategies(self, page, spec):
        if "selector" in spec:
            return [lambda: page.locator(spec["selector"])]
        n, hint, raw = spec["name"], spec["hint"], spec["raw"]
        S = []
        if hint in ("button", "link", "tab", "menuitem", "checkbox", "radio", "heading", "combobox"):
            roles = [hint] + (["link", "menuitem"] if hint == "button" else []) + (["button"] if hint in (
                "link", "menuitem", "tab") else [])
            for r in roles:
                S.append(lambda r=r: page.get_by_role(r, name=n, exact=True))
            if hint in ("checkbox", "radio", "combobox"):
                S.append(lambda: page.get_by_label(n, exact=True))
            for r in roles:
                S.append(lambda r=r: page.get_by_role(r, name=n))
        lab = re.compile(r"^\s*" + re.escape(n) + r"\s*[*:：]*\s*$", re.I)
        if hint in ("textbox", "searchbox"):
            S += [lambda: page.get_by_label(n, exact=True), lambda: page.get_by_label(lab),
                  lambda: page.get_by_role("spinbutton", name=lab), lambda: page.get_by_placeholder(n, exact=True),
                  lambda: page.get_by_role("textbox", name=n, exact=True),
                  lambda: page.get_by_role("searchbox", name=n), lambda: page.get_by_label(n),
                  lambda: page.get_by_placeholder(n)]
        if hint == "img":
            S += [lambda: page.get_by_alt_text(n, exact=True), lambda: page.get_by_alt_text(n),
                  lambda: page.get_by_role("img", name=n)]
        if hint == "table":
            S += [lambda: page.get_by_role("table", name=n), lambda: page.locator("table").filter(has_text=n)]
        if hint == "text":
            S += [lambda: page.get_by_text(n, exact=True), lambda: page.get_by_text(n)]
        S += [
            lambda: page.get_by_label(n, exact=True),
            lambda: page.get_by_label(lab),
            lambda: page.get_by_placeholder(n, exact=True),
            lambda: page.get_by_role("button", name=n, exact=True),
            lambda: page.get_by_role("link", name=n, exact=True),
            lambda: page.get_by_role("heading", name=n, exact=True),
            lambda: page.get_by_text(n, exact=True),
            lambda: page.get_by_title(n, exact=True),
            lambda: page.get_by_alt_text(n, exact=True),
            lambda: page.get_by_test_id(n),
            lambda: page.locator("[name=%s]" % json.dumps(n)),
        ]
        if re.match(r"^[A-Za-z][\w-]*$", n):
            S.append(lambda: page.locator("#" + n))
        S += [lambda: page.get_by_label(n), lambda: page.get_by_placeholder(n),
              lambda: page.get_by_role("button", name=n), lambda: page.get_by_role("link", name=n),
              lambda: page.get_by_text(n)]
        if raw != n:
            S += [lambda: page.get_by_text(raw, exact=True), lambda: page.get_by_label(raw),
                  lambda: page.get_by_placeholder(raw)]
        return S

    def find(self, target, want_visible=True, timeout=None):
        """Trả về Locator của phần tử. Thử nhiều cách tìm (nhãn, placeholder, nút, chữ…) cho đến khi hết giờ."""
        spec = self._parse_target(self._subst(target, self.vars))
        page = self.page
        deadline = time.time() + (timeout if timeout is not None else self.timeout)
        fallback = None
        while True:
            self._check_stop()
            for strat in self._strategies(page, spec):
                try:
                    loc = strat()
                    cnt = loc.count()
                except Exception:
                    continue
                if not cnt:
                    continue
                vis = loc.filter(visible=True)
                try:
                    vc = vis.count()
                except Exception:
                    vc = 0
                pick = vis if vc else loc
                total = vc if vc else cnt
                if spec["nth"] is not None:
                    if spec["nth"] >= total:
                        continue
                    pick = pick.nth(spec["nth"])
                else:
                    pick = pick.first
                if vc or not want_visible:
                    return pick
                if fallback is None:
                    fallback = pick
            if time.time() >= deadline:
                if fallback is not None:
                    return fallback
                raise StepFail("Không tìm thấy đối tượng '%s' sau %g giây." % (target, self.timeout if timeout is None
                                                                                else timeout))
            self._pump(250)

    def _any_visible(self, target):
        spec = self._parse_target(self._subst(target, self.vars))
        for strat in self._strategies(self.page, spec):
            try:
                loc = strat()
                if loc.count() and loc.filter(visible=True).count():
                    return True
            except Exception:
                continue
        return False

    # ------------------------------------------------------------------ thực thi 1 hành động
    def _wait_until(self, fn, seconds=None):
        deadline = time.time() + (seconds if seconds is not None else self.timeout)
        while True:
            self._check_stop()
            try:
                ok, info = fn()
            except Exception as e:
                ok, info = False, str(e).splitlines()[0][:200]
            if ok:
                return info
            if time.time() >= deadline:
                raise StepFail(info)
            self._pump(250)

    def _friendly(self, e, f):
        msg = str(e).splitlines()[0] if str(e) else e.__class__.__name__
        if "Timeout" in msg:
            return "Quá thời gian chờ khi thao tác với '%s' (%s)." % (f.get("doi_tuong") or f.get("gia_tri") or "",
                                                                   msg[:160])
        if "net::ERR" in msg:
            m = re.search(r"net::ERR_\w+", msg)
            return "Không vào được trang (%s). Kiểm tra địa chỉ, mạng nội bộ/VPN." % (m.group(0) if m else msg)
        if "Element is not" in msg or "not enabled" in msg:
            return "Phần tử không thao tác được: %s" % msg[:200]
        return msg[:300]

    def _exec(self, action, f, sr, depth=0):
        target = f.get("doi_tuong", "")
        value = self._subst(f.get("gia_tri", "") or f.get("du_lieu_buoc", ""), self.vars)
        page = self.page
        tsec = self.timeout * 1000

        def need_t():
            if not target:
                raise StepFail("Bước này cần điền cột Đối tượng.")

        def need_v():
            if value == "":
                raise StepFail("Bước này cần điền cột Giá trị.")

        if action == "mở_trang":
            need_v()
            url = value
            if not re.match(r"^[a-z]+://", url) and not url.startswith("about:"):
                if url.startswith("/") and self.vars.get("base_url"):
                    url = self.vars["base_url"].rstrip("/") + url
                else:
                    url = "http://" + url
            if "${" in url:
                raise StepFail("Biến chưa có giá trị: %s (kiểm tra sheet Env / môi trường đã chọn)." % url)
            try:
                page.goto(url, wait_until="load")
            except Exception as e:
                raise StepFail(self._friendly(e, f))
            return "Đã mở %s" % page.url
        if action in ("bấm", "bấm_đúp", "bấm_phải", "di_chuột"):
            need_t()
            loc = self.find(target)
            try:
                if action == "bấm":
                    loc.click(timeout=tsec)
                elif action == "bấm_đúp":
                    loc.dblclick(timeout=tsec)
                elif action == "bấm_phải":
                    loc.click(button="right", timeout=tsec)
                else:
                    loc.hover(timeout=tsec)
            except Exception as e:
                raise StepFail(self._friendly(e, f))
            try:
                page.wait_for_load_state("domcontentloaded", timeout=5000)
            except Exception:
                pass
            self._pump(150)   # cho trang kịp mở hộp thoại / tự focus ô đầu tiên…
            return {"bấm": "Đã bấm", "bấm_đúp": "Đã bấm đúp", "bấm_phải": "Đã bấm phải",
                    "di_chuột": "Đã đưa chuột lên"}[action] + " '%s'" % target
        if action in ("nhập", "gõ_từng_phím"):
            need_t()
            loc = self.find(target)
            got = None
            for attempt in range(3):
                try:
                    if action == "nhập":
                        loc.fill(value, timeout=tsec)
                    else:
                        loc.fill("", timeout=tsec)
                        loc.press_sequentially(value, delay=60)
                except Exception as e:
                    raise StepFail(self._friendly(e, f))
                try:      # đọc lại: trang có thể tự chuyển con trỏ sang ô khác giữa chừng
                    got = loc.input_value(timeout=2000)
                except Exception:
                    got = None    # không phải ô input (contenteditable…) → không kiểm tra được
                if got is None or got == value or (value and got.replace(",", "") == value.replace(",", "")):
                    break
                self._pump(300)
            else:
                raise StepFail("Đã nhập '%s' vào '%s' nhưng ô đang có '%s' (ô giới hạn ký tự / định dạng, hoặc "
                               "trang tự xóa?)" % (value, target, got))
            return "Đã nhập '%s' vào '%s'" % (value, target)
        if action == "xóa_nội_dung":
            need_t()
            self.find(target).fill("", timeout=tsec)
            return "Đã xóa nội dung '%s'" % target
        if action == "nhấn_phím":
            need_v()
            key = {"enter": "Enter", "tab": "Tab", "esc": "Escape", "escape": "Escape", "xoa": "Backspace",
                   "backspace": "Backspace", "space": "Space", "cach": "Space"}.get(C.norm(value), value)
            if target:
                self.find(target).press(key, timeout=tsec)
            else:
                page.keyboard.press(key)
            return "Đã nhấn phím %s" % key
        if action == "chọn":
            need_t()
            need_v()
            loc = self.find(target)
            try:
                loc.select_option(label=value, timeout=tsec)
            except Exception:
                try:
                    loc.select_option(value=value, timeout=3000)
                except Exception:
                    try:  # dropdown tự làm (không phải <select>): bấm mở rồi bấm vào lựa chọn
                        loc.click(timeout=tsec)
                        opt = page.get_by_role("option", name=value)
                        if not opt.count():
                            opt = page.get_by_text(value, exact=True)
                        opt.filter(visible=True).first.click(timeout=tsec)
                    except Exception as e:
                        raise StepFail("Không chọn được '%s' trong '%s': %s" % (value, target, self._friendly(e, f)))
            return "Đã chọn '%s' trong '%s'" % (value, target)
        if action in ("tick", "bỏ_tick"):
            need_t()
            loc = self.find(target, want_visible=False)
            try:
                loc.set_checked(action == "tick", timeout=tsec)
            except Exception:
                try:
                    loc.set_checked(action == "tick", timeout=tsec, force=True)
                except Exception as e:
                    raise StepFail(self._friendly(e, f))
            return ("Đã tick '%s'" if action == "tick" else "Đã bỏ tick '%s'") % target
        if action == "tải_file_lên":
            need_t()
            need_v()
            files = [p.strip().strip('"') for p in value.split(";") if p.strip()]
            missing = [p for p in files if not os.path.isfile(p)]
            if missing:
                raise StepFail("Không thấy file: %s" % ", ".join(missing))
            loc = self.find(target, want_visible=False)
            try:
                loc.set_input_files(files, timeout=tsec)
            except Exception:
                with page.expect_file_chooser(timeout=tsec) as fc:
                    loc.click()
                fc.value.set_files(files)
            return "Đã chọn file %s" % ", ".join(os.path.basename(p) for p in files)
        if action == "chờ":
            try:
                sec = float(value.replace(",", ".") or 1)
            except ValueError:
                raise StepFail("Giá trị của 'chờ' phải là số giây.")
            end = time.time() + min(sec, 600)
            while time.time() < end:
                self._check_stop()
                page.wait_for_timeout(min(250, max(1, (end - time.time()) * 1000)))
            return "Đã chờ %g giây" % sec
        if action in ("chờ_thấy", "chờ_ẩn"):
            need_t()
            sec = float(value) if re.match(r"^\d+(\.\d+)?$", value or "") else max(self.timeout, 15)
            if action == "chờ_thấy":
                self.find(target, timeout=sec)
                return "Đã thấy '%s'" % target
            self._wait_until(lambda: (not self._any_visible(target), "'%s' vẫn còn hiển thị sau %g giây" % (
                target, sec)), sec)
            return "'%s' đã ẩn" % target
        if action == "cuộn_tới":
            need_t()
            self.find(target).scroll_into_view_if_needed(timeout=tsec)
            return "Đã cuộn tới '%s'" % target
        if action == "quay_lại":
            page.go_back()
            return "Đã quay lại trang trước"
        if action == "tải_lại_trang":
            page.reload()
            return "Đã tải lại trang"
        if action == "chuyển_tab":
            pages = self.ctx.pages
            if value and value.isdigit():
                i = int(value) - 1
                if i >= len(pages):
                    raise StepFail("Chỉ có %d tab, không có tab thứ %s." % (len(pages), value))
                self.page = pages[i]
            else:
                if len(pages) < 2:
                    deadline = time.time() + self.timeout
                    while len(self.ctx.pages) < 2 and time.time() < deadline:
                        page.wait_for_timeout(250)
                    pages = self.ctx.pages
                self.page = pages[-1]
            self.page.bring_to_front()
            try:
                self.page.wait_for_load_state("load")
            except Exception:
                pass
            return "Đã chuyển sang tab: %s" % self.page.url
        if action == "kiểm_tra_hiển_thị":
            need_t()
            loc = self.find(target)
            self._wait_until(lambda: (loc.is_visible(), "'%s' có trong trang nhưng KHÔNG hiển thị" % target))
            return "'%s' đang hiển thị" % target
        if action == "kiểm_tra_ẩn":
            need_t()
            self._wait_until(lambda: (not self._any_visible(target), "'%s' vẫn đang hiển thị" % target))
            return "'%s' không hiển thị" % target
        if action in ("kiểm_tra_text", "kiểm_tra_không_có_text"):
            need_v()
            want = action == "kiểm_tra_text"
            loc = self.find(target) if target else None

            def chk():
                txt = (loc.inner_text(timeout=2000) if loc is not None else page.inner_text("body", timeout=2000))
                txt2 = re.sub(r"\s+", " ", txt)
                has = value in txt or value in txt2 or C.norm(value) and value.lower() in txt2.lower()
                if want:
                    return bool(has), "Không thấy chữ '%s' %s.%s" % (
                        value, "trong '%s'" % target if target else "trên trang",
                        (" Trong '%s' đang có: '%s'" % (target, txt2[:120])) if target else "")
                return (not has), "Vẫn thấy chữ '%s' %s" % (value, "trong '%s'" % target if target else "trên trang")
            self._wait_until(chk)
            return ("Thấy chữ '%s'" if want else "Không có chữ '%s'") % value
        if action == "kiểm_tra_giá_trị":
            need_t()
            loc = self.find(target)
            self._wait_until(lambda: (loc.input_value(timeout=2000) == value, "Ô '%s' đang có giá trị '%s', mong đợi"
                                      " '%s'" % (target, loc.input_value(timeout=2000), value)))
            return "Ô '%s' có giá trị '%s'" % (target, value)
        if action == "kiểm_tra_url":
            need_v()
            self._wait_until(lambda: (value in self.page.url, "Địa chỉ hiện tại là '%s', không chứa '%s'" % (
                self.page.url, value)))
            return "Địa chỉ trang: %s" % self.page.url
        if action == "kiểm_tra_tiêu_đề":
            need_v()
            self._wait_until(lambda: (value.lower() in page.title().lower(), "Tiêu đề trang là '%s'" % page.title()))
            return "Tiêu đề trang: %s" % page.title()
        if action == "kiểm_tra_số_lượng":
            need_t()
            need_v()
            m = re.match(r"^\s*(>=|<=|=|==|>|<)?\s*(\d+)\s*$", value)
            if not m:
                raise StepFail("Giá trị phải dạng: 5, =5, >=1, <10")
            op, num = m.group(1) or "=", int(m.group(2))
            spec = self._parse_target(self._subst(target, self.vars))

            def count():
                best = 0
                for strat in self._strategies(page, spec):
                    try:
                        c = strat().filter(visible=True).count()
                    except Exception:
                        continue
                    if c:
                        return c
                return best

            def chk():
                c = count()
                ok = {"=": c == num, "==": c == num, ">=": c >= num, "<=": c <= num, ">": c > num, "<": c < num}[op]
                return ok, "Đếm được %d '%s', mong đợi %s%d" % (c, target, op, num)
            info = self._wait_until(chk)
            return info.replace(", mong đợi", " — đúng mong đợi")
        if action in ("kiểm_tra_bật", "kiểm_tra_tắt"):
            need_t()
            loc = self.find(target)
            want = action == "kiểm_tra_bật"
            self._wait_until(lambda: (loc.is_enabled() == want, "'%s' đang %s" % (
                target, "bị khóa" if want else "bật (bấm được)")))
            return "'%s' đang %s" % (target, "bật" if want else "bị khóa")
        if action == "kiểm_tra_đã_tick":
            need_t()
            loc = self.find(target, want_visible=False)
            self._wait_until(lambda: (loc.is_checked(), "'%s' chưa được tick" % target))
            return "'%s' đã tick" % target
        if action == "kiểm_tra_hộp_thoại":
            need_v()
            self._wait_until(lambda: (any(value.lower() in d.lower() for d in self.dialogs),
                                      "Không thấy hộp thoại chứa '%s'. Các hộp thoại đã hiện: %s" % (
                                          value, self.dialogs or "(không có)")), min(self.timeout, 5))
            return "Hộp thoại: %s" % [d for d in self.dialogs if value.lower() in d.lower()][-1]
        if action == "hủy_hộp_thoại_tiếp_theo":
            self.dismiss_next = True
            return "Hộp thoại tiếp theo sẽ bấm Hủy"
        if action == "đổi_kích_thước":
            need_v()
            s = X.size_of(value)
            if not s:
                raise StepFail("Kích thước không hợp lệ: '%s' (VD: 1366x768, mobile, tablet)" % value)
            page.set_viewport_size({"width": s[0], "height": s[1]})
            page.wait_for_timeout(400)
            return "Đã đổi kích thước %dx%d" % s
        if action == "kiểm_tra_không_tràn":
            page.wait_for_timeout(300)
            info = page.evaluate("""() => { const W = document.documentElement.clientWidth;
              const sw = Math.max(document.documentElement.scrollWidth, document.body ? document.body.scrollWidth : 0);
              const bad = [];
              for (const el of document.querySelectorAll('body *')) {
                const r = el.getBoundingClientRect(); const cs = getComputedStyle(el);
                if (r.width && r.right > W + 1 && cs.position !== 'fixed' && cs.visibility !== 'hidden') {
                  let p = el.parentElement, clipped = false;
                  while (p && p !== document.body) { const o = getComputedStyle(p).overflowX;
                    if (o === 'hidden' || o === 'auto' || o === 'scroll' || o === 'clip') { clipped = true; break; }
                    p = p.parentElement; }
                  if (!clipped) bad.push((el.tagName.toLowerCase() + (el.id ? '#' + el.id : '') +
                     (el.className && typeof el.className === 'string' ? '.' + el.className.trim().split(/\\s+/)[0] : '')
                     + ' (' + Math.round(r.right - W) + 'px)'));
                }
                if (bad.length > 8) break;
              }
              return {W, sw, bad}; }""")
            if info["sw"] > info["W"] + 1:
                self._shot(sr, "tran")
                raise StepFail("Trang bị tràn ngang: rộng %dpx > màn hình %dpx. Phần tử tràn: %s" % (
                    info["sw"], info["W"], ", ".join(info["bad"][:6]) or "?"))
            return "Không tràn ngang (rộng %dpx)" % info["W"]
        if action == "chụp_màn_hình":
            full = "toan" in C.norm(value)
            self._shot(sr, "chup", full=full, locator=self.find(target) if target else None)
            return "Đã chụp ảnh"
        if action == "so_sánh_ảnh":
            return self._visual(sr, target, value)
        if action == "đo_thời_gian_tải":
            limit = float((value or "3").replace(",", "."))
            t = page.evaluate("""() => { const n = performance.getEntriesByType('navigation')[0];
               return n ? (n.loadEventEnd || n.domContentLoadedEventEnd || n.duration) / 1000 : -1; }""")
            if t < 0:
                raise StepFail("Không đo được thời gian tải trang (hãy mở trang trước).")
            if t > limit:
                raise StepFail("Trang tải mất %.2f giây, chậm hơn mức cho phép %g giây." % (t, limit))
            return "Trang tải trong %.2f giây (cho phép %g giây)" % (t, limit)
        if action == "kiểm_tra_không_lỗi_js":
            page.wait_for_timeout(300)
            if self.console_errors:
                raise StepFail("Có %d lỗi JavaScript: %s" % (len(self.console_errors),
                                                             " | ".join(self.console_errors[:3])[:400]))
            return "Không có lỗi JavaScript"
        if action == "kiểm_tra_link_hỏng":
            links = page.evaluate("""() => [...new Set([...document.querySelectorAll('a[href]')].map(a => a.href)
                .filter(h => h.startsWith('http')))].slice(0, 150)""")
            bad = []
            for u in links:
                self._check_stop()
                try:
                    r = self.ctx.request.get(u, timeout=10000, max_redirects=5)
                    if r.status >= 400:
                        bad.append("%s (%d)" % (u, r.status))
                except Exception as e:
                    bad.append("%s (%s)" % (u, str(e).splitlines()[0][:60]))
            if bad:
                raise StepFail("%d/%d link lỗi: %s" % (len(bad), len(links), "; ".join(bad[:8])))
            return "Đã kiểm tra %d link, không có link hỏng" % len(links)
        if action == "lưu_text":
            need_t()
            need_v()
            loc = self.find(target)
            try:
                txt = loc.input_value(timeout=1000)
            except Exception:
                txt = loc.inner_text(timeout=tsec).strip()
            name = value.strip().strip("$").strip("{}").strip()
            self.vars[name] = txt
            return "Đã lưu ${%s} = '%s'" % (name, txt)
        if action == "chạy_tc":
            need_v()
            if depth >= 5:
                raise StepFail("Gọi TC lồng nhau quá sâu (có thể TC tự gọi lại chính nó).")
            sub = self.tc_by_id.get(value) or self.tc_by_key.get(value)
            if not sub:
                raise StepFail("Không có TC '%s' trong file." % value)
            done = []
            for s in sub["steps"]:
                self._check_stop()
                if s["kind"] != "auto":
                    raise StepFail("TC '%s' bước %s không tự động được (test tay / thiếu hành động) nên không gọi "
                                   "được." % (value, s["no"]))
                try:
                    self._exec(s["action"], s["fields"], sr, depth + 1)
                except StepFail as e:
                    raise StepFail("TC %s, bước %s lỗi: %s" % (value, s["no"], e))
                except Stopped:
                    raise
                except Exception as e:
                    raise StepFail("TC %s, bước %s lỗi: %s" % (value, s["no"], self._friendly(e, s["fields"])))
                done.append(s["no"])
            return "Đã chạy %s (%d bước)" % (value, len(done))
        if action == "chạy_js":
            need_v()
            try:
                out = page.evaluate("() => { %s }" % value if "return" in value else "() => (%s)" % value)
            except Exception as e:
                raise StepFail("Lỗi JavaScript: %s" % str(e).splitlines()[0][:200])
            if out is False:
                raise StepFail("Đoạn JavaScript trả về false")
            return "JavaScript trả về: %s" % (json.dumps(out, ensure_ascii=False)[:200] if out is not None else "OK")
        raise StepFail("Hành động '%s' chưa được hỗ trợ." % action)

    def _visual(self, sr, target, value):
        from PIL import Image, ImageChops
        try:
            limit = float((value or "1").replace("%", "").replace(",", "."))
        except ValueError:
            limit = 1.0
        vp = self.page.viewport_size or {"width": 0, "height": 0}
        name = "%s_b%s_%dx%d_%s.png" % (slug(self.inst["tc"]["id"]), slug(str(sr["no"]), 8), vp["width"],
                                        vp["height"], self.inst["browser"])
        os.makedirs(self.base_dir, exist_ok=True)
        base = os.path.join(self.base_dir, name)
        self.page.wait_for_timeout(300)
        cur = self._shot(sr, "sosanh", locator=self.find(target) if target else None)
        if not cur:
            raise StepFail("Không chụp được ảnh để so sánh.")
        sr["baseline"] = base
        if not os.path.exists(base):
            shutil.copyfile(cur, base)
            return "Lần đầu: đã lưu ảnh này làm ẢNH CHUẨN cho các lần sau"
        a = Image.open(base).convert("RGB")
        b = Image.open(cur).convert("RGB")
        if a.size != b.size:
            raise StepFail("Kích thước ảnh khác ảnh chuẩn (%dx%d so với %dx%d)." % (b.size + a.size))
        diff = ImageChops.difference(a, b).convert("L").point(lambda p: 255 if p > 40 else 0)
        changed = sum(1 for p in diff.getdata() if p)
        ratio = changed * 100.0 / (a.size[0] * a.size[1])
        if ratio > limit:
            red = Image.new("RGB", b.size, (255, 0, 0))
            hl = Image.composite(red, b, diff)
            hl = Image.blend(b, hl, 0.6)
            dname = os.path.basename(cur).replace(".png", "_khac.png")
            hl.save(os.path.join(self.img_dir, dname))
            sr["images"].append("anh/" + dname)
            raise StepFail("Giao diện khác ảnh chuẩn %.2f%% (cho phép %g%%). Xem ảnh tô đỏ chỗ khác." % (ratio, limit))
        return "Giống ảnh chuẩn (khác %.2f%%, cho phép %g%%)" % (ratio, limit)
