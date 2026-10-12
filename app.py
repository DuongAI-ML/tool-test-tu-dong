#!/usr/bin/env python3
"""
Tool Test Tự Động — đọc test case/test scenario từ Excel, tự chạy trên web bằng Edge/Chrome có sẵn, xuất kết quả.

Chạy:  python app.py                 (mở cửa sổ tool ở http://127.0.0.1:8620)
       python app.py --no-browser --port 8620 --data D:\\du-lieu
Cần:   pip install playwright openpyxl pillow   (KHÔNG cần 'playwright install' — dùng Edge/Chrome của máy)
"""
import argparse
import base64
import datetime as dt
import json
import logging
import mimetypes
import os
import re
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from logging.handlers import RotatingFileHandler
from urllib.parse import parse_qs, unquote, urlparse

import columns as C
import engine as E
import excel_io as X
import keywords as K
import report as R
import sample
import scan_types as ST
import scanner as SC

APP_NAME = "Tool Test Tự Động"
VERSION = "1.0.0"
DEFAULT_PORT = 8620
DATA_FOLDER = "ToolTestTuDong"

FROZEN = getattr(sys, "frozen", False)
BASE = os.path.dirname(os.path.abspath(sys.executable if FROZEN else __file__))
RES = getattr(sys, "_MEIPASS", BASE)
WEB_DIR = os.path.join(RES, "web")
DEMO_DIR = os.path.join(RES, "demo")
IS_WIN = os.name == "nt"

log = logging.getLogger("tooltest")
log.setLevel(logging.INFO)

DATA = None
PORT = None
SETTINGS = {}
SET_LOCK = threading.Lock()
WB = None               # file Excel đang mở (dict từ excel_io.read_workbook)
WB_LOCK = threading.Lock()
RUNNER = None           # engine.Runner đang/đã chạy gần nhất
TK_LOCK = threading.Lock()
EXPORTED = set()        # đường dẫn đã xuất (cho phép mở bằng /api/open_path)
MAX_BODY = 60_000_000


class ApiError(Exception):
    def __init__(self, msg, code=400):
        super().__init__(msg)
        self.code = code


# ----------------------------------------------------------------------------- cài đặt
def documents_dir():
    if IS_WIN:
        try:
            import ctypes
            from ctypes import wintypes
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            if ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf) == 0 and buf.value:
                return buf.value
        except Exception:
            pass
    return os.path.join(os.path.expanduser("~"), "Documents")


def default_settings():
    return {"tester": "", "project_name": "", "browser": "edge", "headless": False, "speed": "vua", "timeout": 10,
            "shot_every_step": False, "video": False, "auto_first": True, "stop_on_fail": True, "size": "1366x768",
            "env": "", "last_file": "", "recent": [],
            "export": {"pattern": "KetQua_{du_an}_{ngay}_{gio}", "folder": os.path.join(documents_dir(), "KetQuaTest"),
                       "embed": True, "html": True, "writeback": False, "open_after": True, "round": "1"},
            "active_profile": C.COMPANY_NAME, "profiles": [C.company_profile(), C.default_profile()],
            "base_url": "",
            "scan": {"url": "", "login": {"mode": "none", "url": "", "user": "", "remember": True}, "depth": 1,
                     "max_pages": 10, "skip_words": ST.SKIP_LINK_WORDS, "types": ST.default_type_opts(),
                     "rules": [], "state_rules": [], "perm": {"role": "", "rows": []}, "click": {},
                     "source": "auto", "scenarios_in": [], "scenarios_file": "", "only": [],
                     "kind": "scenario", "suggest": True, "wait_extra": 8, "edits": {},
                     "out": {"pattern_scenario": "TestScenario_{man_hinh}_{ngay}",
                             "pattern_testcase": "TestCase_{man_hinh}_{ngay}", "folder": "", "round": "1"}}}


def settings_path():
    return os.path.join(DATA, "settings.json")


def load_settings():
    global SETTINGS
    s = default_settings()
    try:
        with open(settings_path(), encoding="utf-8") as f:
            saved = json.load(f)
        if isinstance(saved, dict):
            exp = dict(s["export"], **(saved.get("export") or {}))
            s.update(saved)
            s["export"] = exp
    except FileNotFoundError:
        pass
    except Exception as e:
        log.warning("settings.json lỗi, dùng mặc định: %s", e)
    profs = [C.fix_profile(p) for p in s.get("profiles") or [] if isinstance(p, dict)]
    if not profs:
        profs = [C.default_profile()]
    # nâng cấp: chưa có mẫu cột theo file tester (và người dùng chưa từng xóa nó) → thêm và dùng làm mặc định
    if C.COMPANY_NAME not in [p["name"] for p in profs] and not s.get("company_removed"):
        profs.insert(0, C.company_profile())
        s["active_profile"] = C.COMPANY_NAME
    base_scan = default_settings()["scan"]
    sc = dict(base_scan, **(s.get("scan") or {}))
    sc["out"] = dict(base_scan["out"], **(sc.get("out") if isinstance(sc.get("out"), dict) else {}))
    types = ST.default_type_opts()
    for k, v in (sc.get("types") or {}).items():
        if k in types and isinstance(v, dict):
            types[k]["on"] = bool(v.get("on"))
            types[k]["opts"].update({kk: vv for kk, vv in (v.get("opts") or {}).items() if kk in types[k]["opts"]})
    sc["types"] = types
    s["scan"] = sc
    s["profiles"] = profs
    if s.get("active_profile") not in [p["name"] for p in profs]:
        s["active_profile"] = profs[0]["name"]
    SETTINGS = s


def save_settings():
    with SET_LOCK:
        tmp = settings_path() + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(SETTINGS, f, ensure_ascii=False, indent=1)
        os.replace(tmp, settings_path())


def active_profile():
    for p in SETTINGS["profiles"]:
        if p["name"] == SETTINGS["active_profile"]:
            return p
    return SETTINGS["profiles"][0]


def demo_url():
    return "http://127.0.0.1:%d/demo" % PORT


# ----------------------------------------------------------------------------- hộp thoại chọn file (tkinter)
def tk_dialog(kind, **kw):
    with TK_LOCK:
        try:
            import tkinter as tk
            from tkinter import filedialog
        except Exception:
            raise ApiError("Máy không mở được hộp thoại chọn file. Hãy kéo-thả file vào cửa sổ tool.")
        root = tk.Tk()
        root.withdraw()
        try:
            root.attributes("-topmost", True)
            root.update()
            if kind == "open":
                return filedialog.askopenfilename(parent=root, **kw) or ""
            if kind == "save":
                return filedialog.asksaveasfilename(parent=root, **kw) or ""
            if kind == "folder":
                return filedialog.askdirectory(parent=root, **kw) or ""
        finally:
            try:
                root.destroy()
            except Exception:
                pass
    return ""


# ----------------------------------------------------------------------------- file Excel
def open_workbook(path):
    global WB
    wb = X.read_workbook(path, active_profile())
    with WB_LOCK:
        WB = wb
    SETTINGS["last_file"] = wb["file"]
    rec = [wb["file"]] + [p for p in SETTINGS.get("recent", []) if p != wb["file"]]
    SETTINGS["recent"] = rec[:8]
    if SETTINGS.get("env") not in wb["env_names"]:
        SETTINGS["env"] = wb["env_names"][0]
    save_settings()
    log.info("Mở file %s: %d scenario, %d TC", path, len(wb["scenarios"]), len(wb["testcases"]))
    return wb


def wb_view(wb):
    if not wb:
        return None
    return {k: wb[k] for k in ("file", "file_name", "project", "scenarios", "testcases", "env_names", "envs",
                               "elements", "headers", "sheets", "warnings")} | {
        "modified": os.path.exists(wb["file"]) and os.path.getmtime(wb["file"]) > wb["mtime"] + 0.5}


# ----------------------------------------------------------------------------- lần chạy
def runs_dir():
    return os.path.join(DATA, "runs")


def run_path(run_id):
    if not re.match(r"^[\w-]{6,40}$", run_id or ""):
        raise ApiError("Mã lần chạy không hợp lệ")
    p = os.path.join(runs_dir(), run_id)
    if not os.path.isdir(p):
        raise ApiError("Không thấy lần chạy %s" % run_id, 404)
    return p


def load_result(run_id):
    if RUNNER and os.path.basename(RUNNER.run_dir) == run_id:
        return RUNNER.result_data()
    p = os.path.join(run_path(run_id), "result.json")
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        raise ApiError("Lần chạy chưa có kết quả", 404)


def list_runs(limit=40):
    out = []
    try:
        names = sorted(os.listdir(runs_dir()), reverse=True)
    except FileNotFoundError:
        return out
    for n in names[:limit]:
        p = os.path.join(runs_dir(), n, "result.json")
        try:
            with open(p, encoding="utf-8") as f:
                d = json.load(f)
        except Exception:
            continue
        c = R.counts(d)
        out.append({"run_id": n, "project": d.get("project"), "started": d.get("started"), "state": d.get("state"),
                    "env": d.get("env"), "counts": c, "total": sum(c.values())})
    return out


def run_number(data):
    """Lần chạy thứ mấy trong ngày của cùng dự án (cho biến {lan_chay})."""
    day = (data.get("started") or "")[:10]
    same = [r for r in list_runs(500) if r["project"] == data.get("project") and (r["started"] or "")[:10] == day]
    same.sort(key=lambda r: r["started"] or "")
    for i, r in enumerate(same, start=1):
        if r["run_id"] == data.get("run_id"):
            return i
    return len(same) + 1


def start_run(body):
    global RUNNER
    if RUNNER and RUNNER.thread and RUNNER.thread.is_alive():
        raise ApiError("Đang có lần chạy khác. Hãy đợi xong hoặc bấm Dừng.")
    with WB_LOCK:
        wb = WB
    if not wb:
        raise ApiError("Chưa mở file Excel test case.")
    sel = [k for k in body.get("selection") or [] if isinstance(k, str)]
    if not sel:
        raise ApiError("Chưa tick test case nào.")
    opts = {k: SETTINGS.get(k) for k in ("tester", "browser", "headless", "speed", "timeout", "shot_every_step",
                                          "video", "auto_first", "stop_on_fail", "size", "env")}
    opts.update({k: v for k, v in (body.get("options") or {}).items() if k in opts})
    opts["demo_url"] = demo_url()
    opts["base_url"] = (body.get("options") or {}).get("base_url", SETTINGS.get("base_url", "")).strip()
    if opts.get("env") not in wb["env_names"]:
        opts["env"] = wb["env_names"][0]
    run_id = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    d = os.path.join(runs_dir(), run_id)
    i = 1
    while os.path.exists(d):
        i += 1
        d = os.path.join(runs_dir(), "%s_%d" % (run_id, i))
    os.makedirs(d)
    RUNNER = E.Runner(wb, sel, opts, d, DATA)
    RUNNER.start()
    log.info("Bắt đầu chạy %s: %d TC, %d lượt", os.path.basename(d), len(sel), len(RUNNER.instances))
    return {"run_id": os.path.basename(d), "total": len(RUNNER.instances)}


def do_export(body):
    run_id = body.get("run_id") or (os.path.basename(RUNNER.run_dir) if RUNNER else "")
    if not run_id:
        raise ApiError("Chưa có lần chạy nào để xuất.")
    if RUNNER and os.path.basename(RUNNER.run_dir) == run_id and RUNNER.thread and RUNNER.thread.is_alive():
        raise ApiError("Lần chạy chưa xong. Hãy đợi chạy xong (hoặc bấm Dừng) rồi xuất.")
    data = load_result(run_id)
    exp = dict(SETTINGS["export"])
    for k in ("pattern", "folder", "embed", "html", "writeback", "open_after", "round"):
        if k in body:
            exp[k] = body[k]
    SETTINGS["export"] = exp
    save_settings()
    folder = (exp.get("folder") or "").strip() or os.path.join(documents_dir(), "KetQuaTest")
    try:
        os.makedirs(folder, exist_ok=True)
    except OSError as e:
        raise ApiError("Không tạo được thư mục lưu: %s\n%s" % (folder, e))
    ctx = R.name_context(data, SETTINGS, run_number(data))
    name = R.build_name(exp.get("pattern"), ctx)
    path = os.path.join(folder, name)
    mode = body.get("overwrite", "ask")
    if os.path.exists(path):
        if mode == "ask":
            return {"exists": True, "path": path, "name": name}
        if mode == "suffix":
            path = R.unique_path(path)
    rd = run_path(run_id)
    out = R.export_excel(data, active_profile(), path, rd, embed_images=bool(exp.get("embed")),
                         opts={"round": exp.get("round") or "1"})
    files = [path]
    if exp.get("html"):
        hp = os.path.splitext(path)[0] + ".html"
        R.export_html(data, hp, os.path.splitext(os.path.basename(path))[0] + "_anh")
        files.append(hp)
    if exp.get("writeback") and data.get("file"):
        src = data.get("file")
        wp = R.unique_path(os.path.join(folder, "%s_CoKetQua_%s.xlsx" % (
            os.path.splitext(os.path.basename(src or "TestCase"))[0], ctx["ngay"])))
        R.write_back(data, src, wp, str(exp.get("round") or "1"))
        files.append(wp)
    for f in files:
        EXPORTED.add(os.path.normcase(os.path.abspath(f)))
    EXPORTED.add(os.path.normcase(os.path.abspath(folder)))
    if exp.get("open_after"):
        open_path(path)
    log.info("Xuất kết quả %s -> %s", run_id, files)
    return {"ok": True, "files": files, "folder": folder, "images": out["images"]}


def open_path(p):
    try:
        if IS_WIN:
            os.startfile(p)
        else:
            webbrowser.open("file://" + p)
    except Exception as e:
        raise ApiError("Không mở được: %s (%s)" % (p, e))


# ----------------------------------------------------------------------------- HTTP
class Handler(BaseHTTPRequestHandler):
    server_version = "ToolTest/" + VERSION
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):
        pass

    def _host_ok(self):
        host = (self.headers.get("Host") or "").split(":")[0].lower()
        return host in ("127.0.0.1", "localhost", "[::1]")

    def _send(self, code, body, ctype="application/json; charset=utf-8", extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False).encode("utf-8")
        elif isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _file(self, root, rel):
        rel = unquote(rel).replace("\\", "/").lstrip("/")
        full = os.path.normpath(os.path.join(root, rel))
        if not full.startswith(os.path.normpath(root) + os.sep) and full != os.path.normpath(root):
            return self._send(403, {"error": "forbidden"})
        if os.path.isdir(full):
            full = os.path.join(full, "index.html")
        if not os.path.isfile(full):
            return self._send(404, {"error": "Không thấy: " + rel})
        ctype = mimetypes.guess_type(full)[0] or "application/octet-stream"
        if ctype.startswith("text/") or ctype in ("application/javascript",):
            ctype += "; charset=utf-8"
        with open(full, "rb") as f:
            data = f.read()
        self._send(200, data, ctype)

    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        if not self._host_ok():
            return self._send(403, {"error": "host"})
        u = urlparse(self.path)
        p = u.path
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        try:
            if p in ("/", "/index.html"):
                return self._file(WEB_DIR, "index.html")
            if p == "/favicon.ico":
                return self._file(WEB_DIR, "icon.svg")
            if p.startswith("/web/"):
                return self._file(WEB_DIR, p[5:])
            if p == "/demo" or p == "/demo/":
                self.send_response(302)
                self.send_header("Location", "/demo/login.html")
                self.send_header("Content-Length", "0")
                self.end_headers()
                return
            if p.startswith("/demo/"):
                return self._file(DEMO_DIR, p[6:])
            if p.startswith("/runs/"):
                parts = p[6:].split("/", 1)
                if len(parts) == 2:
                    return self._file(run_path(parts[0]), parts[1])
                return self._send(404, {"error": "?"})
            if p == "/api/state":
                return self._send(200, api_state())
            if p == "/api/run/poll":
                return self._send(200, api_poll(int(q.get("since") or 0)))
            if p == "/api/runs":
                return self._send(200, {"runs": list_runs()})
            if p == "/api/run/result":
                return self._send(200, load_result(q.get("run_id", "")))
            if p == "/api/ping":
                return self._send(200, {"app": "tooltest", "version": VERSION})
            return self._send(404, {"error": "Không có trang này"})
        except ApiError as e:
            return self._send(e.code, {"error": str(e)})
        except Exception as e:
            log.exception("GET %s", p)
            return self._send(500, {"error": "Lỗi: %s" % e})

    def do_POST(self):
        if not self._host_ok() or self.headers.get("X-Tool") != "1":
            return self._send(403, {"error": "forbidden"})
        try:
            n = int(self.headers.get("Content-Length") or 0)
            if n > MAX_BODY:
                raise ApiError("Dữ liệu gửi lên quá lớn", 413)
            raw = self.rfile.read(n) if n else b""
            body = json.loads(raw.decode("utf-8")) if raw else {}
            if not isinstance(body, dict):
                raise ApiError("Dữ liệu không hợp lệ")
            fn = ROUTES.get(urlparse(self.path).path)
            if not fn:
                raise ApiError("Không có API này", 404)
            return self._send(200, fn(body))
        except ApiError as e:
            return self._send(e.code, {"error": str(e)})
        except X.ExcelError as e:
            return self._send(400, {"error": str(e)})
        except Exception as e:
            log.exception("POST %s", self.path)
            return self._send(500, {"error": "Lỗi: %s" % e})


def api_state():
    snap = RUNNER.snapshot() if RUNNER else None
    with WB_LOCK:
        wb = WB
    return {"app": APP_NAME, "version": VERSION, "settings": SETTINGS, "profile": active_profile(),
            "catalog": C.CATALOG, "keywords": K.KEYWORDS, "test_types": C.TEST_TYPES, "workbook": wb_view(wb),
            "run": snap, "demo_url": demo_url(), "vars_help": R.VARS_HELP, "data_dir": DATA,
            "scan_types": ST.TYPES, "scan_rules": ST.RULES, "scan_presets": ST.PRESETS,
            "company_profile": C.company_profile(),
            "discovery": load_discovery(SETTINGS["scan"].get("url"))}


def disc_path(url):
    d = os.path.join(DATA, "do-trang")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, E.slug(url or "x", 90) + ".json")


def load_discovery(url):
    if not url:
        return None
    try:
        with open(disc_path(url), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def clean_scan_cfg(cfg):
    """Cấu hình quét để LƯU lại: bỏ mật khẩu."""
    c = json.loads(json.dumps(cfg or {}, ensure_ascii=False))
    (c.get("login") or {}).pop("pass", None)
    perm = ((c.get("types") or {}).get("permission") or {}).get("opts") or {}
    perm.pop("pass", None)
    return c


def r_scan_save(b):
    cfg = b.get("cfg") or {}
    keep = SETTINGS["scan"]
    for k in ("url", "login", "depth", "max_pages", "skip_words", "types", "rules", "state_rules", "perm", "click",
              "source", "scenarios_in", "scenarios_file", "only", "kind", "suggest", "wait_extra", "edits", "out",
              "last_run"):
        if k in cfg:
            keep[k] = clean_scan_cfg({k: cfg[k]})[k]
    save_settings()
    return {"ok": True, "scan": keep, "discovery": load_discovery(keep.get("url"))}


def _start_scan(b, discover_only):
    global RUNNER
    if RUNNER and RUNNER.thread and RUNNER.thread.is_alive():
        raise ApiError("Đang có lần chạy khác. Hãy đợi xong hoặc bấm Dừng.")
    cfg = b.get("cfg") or {}
    url = (cfg.get("url") or "").strip()
    if not url:
        raise ApiError("Hãy nhập địa chỉ (URL) trang web cần test.")
    if not re.match(r"^https?://", url, re.I):
        url = "http://" + url
    cfg["url"] = url
    r_scan_save({"cfg": cfg})
    cfg["discover_only"] = discover_only
    cfg["project"] = SETTINGS.get("project_name") or ""
    if not discover_only:
        disc = load_discovery(url) or {}
        if disc.get("pages_file") and os.path.isfile(disc["pages_file"]) and cfg.get("reuse", True):
            cfg["reuse_pages"] = disc["pages_file"]       # dùng đúng kết quả Dò trang mà bảng scenario dựa vào
        if cfg.get("source") != "file":
            cfg["scenarios_in"] = []
    opts = {k: SETTINGS.get(k) for k in ("tester", "browser", "headless", "speed", "timeout", "size")}
    opts.update({k: v for k, v in (b.get("options") or {}).items() if k in opts})
    for k in ("browser", "headless", "speed"):
        if k in (b.get("options") or {}):
            SETTINGS[k] = b["options"][k]
    save_settings()
    run_id = dt.datetime.now().strftime("%Y%m%d_%H%M%S") + ("_do" if discover_only else "_quet")
    d = os.path.join(runs_dir(), run_id)
    os.makedirs(d, exist_ok=True)
    RUNNER = SC.Scanner(cfg, opts, d, DATA)
    RUNNER.on_done = None
    RUNNER.start()
    log.info("%s %s", "Dò trang" if discover_only else "Quét & test", url)
    return {"run_id": run_id}


def _scan_cfg_for(b):
    cfg = json.loads(json.dumps(SETTINGS["scan"], ensure_ascii=False))
    cfg.update(b.get("cfg") or {})
    if cfg.get("source") != "file":
        cfg["scenarios_in"] = []
    return cfg


def _plan_scanner(cfg):
    """Scanner KHÔNG mở trình duyệt, dùng kết quả Dò trang gần nhất của URL → (scanner, cần dò?, kết quả dò)."""
    disc = load_discovery(cfg.get("url")) or {}
    pages = []
    if disc.get("pages_file") and os.path.isfile(disc["pages_file"]):
        with open(disc["pages_file"], encoding="utf-8") as f:
            pages = json.load(f)
    if not pages and not cfg.get("scenarios_in"):
        return None, True, disc
    tmp = os.path.join(DATA, "tmp-plan")
    os.makedirs(tmp, exist_ok=True)
    for k in ("only", "only_tc"):
        cfg.pop(k, None)
    if not re.match(r"^https?://", cfg.get("url") or "", re.I):
        cfg["url"] = "http://" + (cfg.get("url") or "localhost")
    sc = SC.Scanner(cfg, {"browser": SETTINGS.get("browser", "edge")}, tmp, DATA)
    sc.pages = pages
    sc.logged_in = (cfg.get("login") or {}).get("mode", "none") != "none"
    return sc, not pages, disc


def r_scan_plan(b):
    """Bảng xem trước (Test Scenario theo LUỒNG + test case của từng scenario), không mở trình duyệt."""
    cfg = _scan_cfg_for(b)
    r_scan_save({"cfg": {k: cfg[k] for k in ("source", "scenarios_in", "scenarios_file", "kind", "suggest")
                         if k in cfg}})
    sc, need, disc = _plan_scanner(cfg)
    if not sc:
        return {"scenarios": [], "need_discover": True, "meta": None, "diagnosis": disc.get("diagnosis") or []}
    return {"scenarios": sc.plan_preview(), "meta": sc.plan_meta(), "need_discover": need,
            "diagnosis": disc.get("diagnosis") or []}


def _selected_plan(b):
    """Kế hoạch đã lọc theo các dòng người dùng tick ở bảng xem trước + chỉnh sửa tay (AC, Priority…)."""
    cfg = _scan_cfg_for(b)
    sc, need, _ = _plan_scanner(cfg)
    if not sc:
        raise ApiError("Chưa có dữ liệu dò trang — hãy bấm 🔍 Dò trang trước.")
    plan, meta = sc.plan_preview(), sc.plan_meta()
    keys = set(b.get("only") or [])
    tcs = set(b.get("only_tc") or [])
    edits = b.get("edits") or {}
    out = []
    for x in plan:
        if keys and x["key"] not in keys:
            continue
        x = dict(x)
        x["tc_items"] = [t for t in x["tc_items"] if not tcs or t["tc_id"] in tcs]
        if tcs and not x["tc_items"]:
            continue
        for k, v in (edits.get(x["key"]) or {}).items():
            if k in ("scenario", "ac", "priority", "nhom") and isinstance(v, str):
                x[k] = v.strip()
        out.append(x)
    if not out:
        raise ApiError("Chưa chọn dòng nào ở bảng xem trước.")
    return out, meta, sc


def _ts_rows(plan, meta, renumber=True):
    """Dòng Test Scenario để ghi file: ID đánh LIÊN TỤC cả file (TS_DHTD_001…), giữ thứ tự luồng."""
    rows = []
    for i, x in enumerate(plan, start=1):
        sid = ("TS_%s_%03d" % (meta["code"], i)) if renumber and not x.get("imported") else x["id"]
        rows.append({"key": x["key"], "luong": x["luong"], "luong_name": x["luong_name"], "mh": x.get("mh", ""),
                     "module": x.get("module", ""), "mode": x.get("mode", ""),
                     "fields": {"ts_id": sid, "nhom": x["nhom"], "scenario": x["scenario"], "ac": x.get("ac", ""),
                                "priority": x.get("priority", ""), "module": x.get("module", "")}})
    return rows


def _ts_meta(rows, meta):
    names = list(dict.fromkeys(s["name"] for s in meta["screens"] if any(r["mh"] == s["mh"] for r in rows)))
    scr = "; ".join("%s = %s" % (s["mh"], s["name"]) for s in meta["screens"] if s["name"] in names)
    note = "Vai trò khảo sát: %s. Nguồn: tool dò tự động %s lúc %s. Màn hình: %s. Truy vết chi tiết ở sheet " \
           "\"Ma tran truy vet\"." % (meta.get("role") or "", meta.get("url") or "",
                                      dt.datetime.now().strftime("%d/%m/%Y %H:%M"), scr or meta.get("main", ""))
    return {"title": ", ".join(names) or meta.get("main", ""), "note": note, "screens": meta["screens"]}


def _profile_from(b):
    p = b.get("profile")
    return C.fix_profile(p) if isinstance(p, dict) else active_profile()


def _tc_data(plan, meta, b):
    """Dữ liệu test case CHƯA chạy (để sinh file Test Case trước khi chạy) — cùng dạng kết quả lần chạy."""
    results, scns = [], []
    for x in plan:
        scns.append({"id": x["id"], "fields": {"ts_id": x["id"], "nhom": x["nhom"], "scenario": x["scenario"],
                                               "ac": x.get("ac", ""), "priority": x.get("priority", ""),
                                               "module": x.get("module", "")}, "tcs": []})
        for t in x["tc_items"]:
            steps = [{"no": str(i), "fields": {"buoc": s, "expected": e}, "status": "", "actual": "", "images": []}
                     for i, (s, e) in enumerate(zip(t["steps"], t["expected"]), start=1)]
            fields = {"tc_id": t["tc_id"], "scenario_id": x["id"], "requirement": "%s / %s" % (x["id"], x["scenario"]),
                      "module": t["module"], "man_hinh": "", "tieu_de": t["title"], "loai_test": t["nhom"],
                      "priority": t["priority"], "precondition": t["pre"], "test_data": t["data"],
                      "data_id": t["data_id"], "kich_thuoc": t["size"], "cach_test": t["mode"], "moi_truong": t["env"],
                      "mo_ta": t["url"]}
            results.append({"key": t["tc_id"], "tc_id": t["tc_id"], "label": t["tc_id"], "title": t["title"],
                            "scenario_id": x["id"], "mode": t["mode"], "fields": fields, "extra": {}, "steps": steps,
                            "status": "", "note": "", "duration": 0, "env": t["env"], "size": t["size"],
                            "browser": "", "tester": SETTINGS.get("tester", ""), "start": ""})
            scns[-1]["tcs"].append(t["tc_id"])
    return {"project": meta.get("main", ""), "scenarios": scns, "results": results, "headers": {}, "options": {},
            "started": dt.datetime.now().isoformat(timespec="seconds")}


def _out_path(b, kind, meta):
    out = dict(SETTINGS["scan"].get("out") or {})
    for k in ("pattern_scenario", "pattern_testcase", "folder", "round"):
        if k in b:
            out[k] = b[k]
    SETTINGS["scan"]["out"] = out
    save_settings()
    folder = (out.get("folder") or "").strip() or SETTINGS["export"].get("folder") or \
        os.path.join(documents_dir(), "KetQuaTest")
    try:
        os.makedirs(folder, exist_ok=True)
    except OSError as e:
        raise ApiError("Không tạo được thư mục lưu: %s\n%s" % (folder, e))
    now = dt.datetime.now()
    slug = "".join(w[:1].upper() + w[1:] for w in C.norm(meta.get("main", "")).split()) or "Web"
    ctx = {"du_an": SETTINGS.get("project_name") or slug, "man_hinh": slug,
           "ma": meta.get("code", ""), "ngay": now.strftime("%Y-%m-%d"), "gio": now.strftime("%H%M"),
           "nguoi_test": SETTINGS.get("tester", ""), "lan_chay": str(out.get("round") or "1"), "moi_truong": "",
           "trinh_duyet": SETTINGS.get("browser", ""), "module": meta.get("main", "")}
    pat = out.get("pattern_scenario" if kind == "scenario" else "pattern_testcase")
    return folder, os.path.join(folder, R.build_name(pat, ctx)), out


def r_scan_export(b):
    """⬇ Sinh file Excel MỚI: Test Scenario (đúng bố cục file mẫu) hoặc Test Case (theo mẫu cột công ty)."""
    kind = "testcase" if b.get("kind") == "testcase" else "scenario"
    plan, meta, _ = _selected_plan(b)
    folder, path, out = _out_path(b, kind, meta)
    mode = b.get("overwrite", "ask")
    if os.path.exists(path):
        if mode == "ask":
            return {"exists": True, "path": path, "name": os.path.basename(path)}
        if mode == "suffix":
            path = R.unique_path(path)
    prof = _profile_from(b)
    if kind == "scenario":
        cols = X.enabled_columns(prof, "scenario")
        rows = _ts_rows(plan, meta)
        res = R.export_scenarios(path, rows, cols, _ts_meta(rows, meta))
        info = "%d scenario, %d luồng" % (res["rows"], res["groups"])
    else:
        run_id = b.get("run_id") or ""
        if run_id:
            data = load_result(run_id)
            R.export_excel(data, prof, path, run_path(run_id), embed_images=True,
                           opts={"round": str(out.get("round") or "1")})
            info = "%d test case có kết quả (ghi vào cột Lần %s)" % (len(data["results"]), out.get("round") or "1")
        else:
            data = _tc_data(plan, meta, b)
            tmp = os.path.join(DATA, "tmp-plan")
            os.makedirs(tmp, exist_ok=True)
            R.export_excel(data, prof, path, tmp, opts={"plan_only": True})
            info = "%d test case (chưa chạy)" % len(data["results"])
    EXPORTED.add(os.path.normcase(os.path.abspath(path)))
    EXPORTED.add(os.path.normcase(os.path.abspath(folder)))
    log.info("Sinh file %s: %s", kind, path)
    return {"ok": True, "path": path, "folder": folder, "info": info}


def _upload_path(name, data):
    up = os.path.join(DATA, "tai-len")
    os.makedirs(up, exist_ok=True)
    name = os.path.basename(name or "file.xlsx")
    if not name.lower().endswith((".xlsx", ".xlsm")):
        raise ApiError("Chỉ nhận file .xlsx / .xlsm")
    path = os.path.join(up, name)
    with open(path, "wb") as f:
        f.write(base64.b64decode(data))
    return path


def r_xlsx_sheets(b):
    """Chọn file Excel CÓ SẴN (hộp thoại / kéo-thả / dán đường dẫn) → danh sách sheet và cột tool nhận ra."""
    path = (b.get("path") or "").strip().strip('"')
    uploaded = False
    if not path and b.get("data"):
        path, uploaded = _upload_path(b.get("name"), b["data"]), True
    if not path:
        path = tk_dialog("open", title="Chọn file Excel có sẵn để điền", initialdir=documents_dir(),
                         filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tất cả", "*.*")])
        if not path:
            return {"cancel": True}
    if not os.path.isfile(path):
        raise ApiError("Không thấy file: %s" % path)
    if not path.lower().endswith((".xlsx", ".xlsm")):
        raise ApiError("Chỉ điền được file .xlsx / .xlsm")
    EXPORTED.add(os.path.normcase(os.path.abspath(path)))
    return {"path": os.path.abspath(path), "name": os.path.basename(path), "uploaded": uploaded,
            "sheets": R.sheet_info(path, _profile_from(b))}


def r_scan_fill(b):
    """📥 Điền vào file Excel SẴN CÓ: nhận cột theo tên tiêu đề, giữ định dạng, ghi vào BẢN SAO (hoặc ghi đè file
    gốc khi người dùng đã xác nhận)."""
    kind = "testcase" if b.get("kind") == "testcase" else "scenario"
    src = (b.get("path") or "").strip()
    if not src or not os.path.isfile(src):
        raise ApiError("Hãy chọn file Excel có sẵn.")
    plan, meta, _ = _selected_plan(b)
    prof = _profile_from(b)
    sheets = R.sheet_info(src, prof)
    sheet = b.get("sheet") or next((s["name"] for s in sheets if s["kind"] == kind), "")
    if not sheet:
        raise ApiError("Không thấy sheet %s trong file (cần dòng tiêu đề như %s)." % (
            "Test Scenario" if kind == "scenario" else "Test Case",
            "ID | Nhóm | Test Scenario" if kind == "scenario" else "Mã TC | Tên TC | Steps"))
    mode = b.get("mode") if b.get("mode") in ("append", "replace", "update") else "append"
    rnd = str(b.get("round") or (SETTINGS["scan"].get("out") or {}).get("round") or "1")
    items, id_key, upd = [], None, ()
    if kind == "scenario":
        rows = _ts_rows(plan, meta)
        for title, grp in R.scenario_groups(rows):
            items.append(("group", title))
            items += [("row", dict(r["fields"], luong=title, cach_test=r["mode"])) for r in grp]
        id_key = "ts_id"
    else:
        run_id = b.get("run_id") or ""
        data = load_result(run_id) if run_id else _tc_data(plan, meta, b)
        scn = {s["id"]: s for s in data.get("scenarios", [])}
        last = None
        for r in data["results"]:
            if r.get("scenario_id") != last and r.get("scenario_id") in scn:
                last = r["scenario_id"]
                s = scn[last]["fields"]
                items.append(("group", "%s — %s%s" % (last, s.get("scenario", ""), (" (%s)" % s["nhom"])
                                                      if s.get("nhom") else "")))
            steps = r["steps"] or []
            vals = {c["key"]: R._merged_value(c["key"], r, data, steps, rnd) for c in C.TESTCASE_COLUMNS
                    if c["key"] != "anh"}
            if not run_id:
                for k in ("status", "kq_lan1", "kq_lan2", "kq_lan3", "actual", "status_buoc"):
                    vals[k] = ""
            items.append(("row", vals))
        id_key = "tc_id"
        if run_id:
            mode = "update"
            upd = ("kq_lan" + rnd, "status", "actual")
    if b.get("save") == "overwrite":
        if not b.get("confirm"):
            return {"confirm": True, "path": src}
        if os.path.normcase(os.path.abspath(src)).startswith(os.path.normcase(os.path.abspath(DATA))):
            raise ApiError("File kéo-thả là bản chép trong thư mục tool — không ghi đè được file gốc. Hãy chọn file "
                           "bằng nút 📂 hoặc dán đường dẫn, hoặc chọn 'Lưu bản sao'.")
        out_file = src
    else:
        folder = (b.get("folder") or (SETTINGS["scan"].get("out") or {}).get("folder") or "").strip() or \
            os.path.dirname(src)
        os.makedirs(folder, exist_ok=True)
        stem, ext = os.path.splitext(os.path.basename(src))
        out_file = R.unique_path(os.path.join(folder, "%s_DaDien_%s%s" % (stem, dt.datetime.now().strftime(
            "%Y-%m-%d_%H%M"), ext)))
    st = R.fill_existing(src, out_file, sheet, kind, prof, items, mode=mode, id_key=id_key, update_keys=upd, rnd=rnd,
                         renumber="always" if kind == "scenario" else "collision")
    EXPORTED.add(os.path.normcase(os.path.abspath(out_file)))
    EXPORTED.add(os.path.normcase(os.path.abspath(os.path.dirname(out_file))))
    log.info("Điền vào file có sẵn %s -> %s (%s)", src, out_file, st)
    return {"ok": True, "path": out_file, "folder": os.path.dirname(out_file), "stats": st}


def r_scan_import(b):
    """Đọc file Excel TEST SCENARIO của tester (cột ID, Nhóm, Test Scenario, AC liên quan, Priority)."""
    path = b.get("path")
    if not path and b.get("data"):
        up = os.path.join(DATA, "tai-len")
        os.makedirs(up, exist_ok=True)
        path = os.path.join(up, os.path.basename(b.get("name") or "scenario.xlsx"))
        with open(path, "wb") as f:
            f.write(base64.b64decode(b["data"]))
    if not path:
        path = tk_dialog("open", title="Chọn file Excel Test Scenario", initialdir=documents_dir(),
                         filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tất cả", "*.*")])
        if not path:
            return {"cancel": True}
    wb = X.read_workbook(path, active_profile())
    items = [{"id": s["id"], "nhom": s["fields"].get("nhom", ""), "scenario": s["fields"].get("scenario", ""),
              "ac": s["fields"].get("ac", ""), "priority": s["fields"].get("priority", "")} for s in wb["scenarios"]]
    if not items:
        raise ApiError("Không thấy Test Scenario trong file. File cần có các cột ID, Nhóm, Test Scenario "
                       "(AC liên quan, Priority nếu có).")
    SETTINGS["scan"].update({"source": "file", "scenarios_in": items, "scenarios_file": path})
    save_settings()
    return {"ok": True, "count": len(items), "file": path, "scenarios_in": items}


def r_scan_discover(b):
    return _start_scan(b, True)


def r_scan_run(b):
    return _start_scan(b, False)


def r_scan_discovery(b):
    """Sau khi dò xong: lưu kết quả dò của URL để lần sau mở lại không cần dò."""
    data = load_result(b.get("run_id", ""))
    pf = os.path.join(run_path(b.get("run_id", "")), "pages.json")
    disc = {"url": data.get("url"), "when": data.get("started"), "pages": data.get("discovery") or [],
            "pages_file": pf if os.path.isfile(pf) else "", "diagnosis": data.get("diagnosis") or []}
    with open(disc_path(data.get("url")), "w", encoding="utf-8") as f:
        json.dump(disc, f, ensure_ascii=False)
    return {"discovery": disc}


def api_poll(since):
    if not RUNNER:
        return {"run": None, "events": []}
    return {"run": RUNNER.snapshot(), "events": RUNNER.events_since(since),
            "results": [_brief(r) for r in RUNNER.results]}


def _brief(r):
    return {"key": r["key"], "label": r["label"], "title": r["title"], "status": r["status"], "mode": r["mode"],
            "duration": r["duration"], "note": r.get("note", "")}


def r_settings(b):
    allowed = {"tester", "project_name", "browser", "headless", "speed", "timeout", "shot_every_step", "video",
               "auto_first", "stop_on_fail", "size", "env", "base_url"}
    for k, v in b.items():
        if k in allowed:
            SETTINGS[k] = v
        elif k == "export" and isinstance(v, dict):
            SETTINGS["export"].update({kk: vv for kk, vv in v.items() if kk in SETTINGS["export"]})
    try:
        SETTINGS["timeout"] = max(1, min(120, float(SETTINGS.get("timeout") or 10)))
    except (TypeError, ValueError):
        SETTINGS["timeout"] = 10
    save_settings()
    return {"ok": True, "settings": SETTINGS}


def r_profile_save(b):
    p = C.fix_profile(b.get("profile") or {})
    name = (b.get("name") or p["name"] or "").strip()[:60]
    if not name:
        raise ApiError("Hãy đặt tên cho mẫu cột.")
    p["name"] = name
    profs = [x for x in SETTINGS["profiles"] if x["name"] != name]
    old_name = b.get("old_name")
    if old_name and old_name != name and b.get("rename"):
        profs = [x for x in profs if x["name"] != old_name]
    profs.append(p)
    SETTINGS["profiles"] = profs
    SETTINGS["active_profile"] = name
    save_settings()
    _reparse()
    return {"ok": True, "profile": p, "profiles": [x["name"] for x in profs]}


def r_profile_activate(b):
    if b.get("name") not in [x["name"] for x in SETTINGS["profiles"]]:
        raise ApiError("Không có mẫu cột này")
    SETTINGS["active_profile"] = b["name"]
    save_settings()
    _reparse()
    return {"ok": True}


def r_profile_delete(b):
    if len(SETTINGS["profiles"]) <= 1:
        raise ApiError("Phải giữ lại ít nhất 1 mẫu cột.")
    SETTINGS["profiles"] = [x for x in SETTINGS["profiles"] if x["name"] != b.get("name")]
    if b.get("name") == C.COMPANY_NAME:
        SETTINGS["company_removed"] = True       # đã chủ động xóa → không tự thêm lại
    if SETTINGS["active_profile"] == b.get("name"):
        SETTINGS["active_profile"] = SETTINGS["profiles"][0]["name"]
    save_settings()
    return {"ok": True}


def _reparse():
    """Đổi tên cột trong mẫu → đọc lại file để nhận cột theo tên mới."""
    with WB_LOCK:
        wb = WB
    if wb and os.path.isfile(wb["file"]):
        try:
            open_workbook(wb["file"])
        except Exception:
            pass


def r_open(b):
    path = b.get("path")
    if not path:
        start = os.path.dirname(SETTINGS.get("last_file") or "") or documents_dir()
        path = tk_dialog("open", title="Chọn file Excel test case", initialdir=start,
                         filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tất cả", "*.*")])
        if not path:
            return {"cancel": True}
    return {"workbook": wb_view(open_workbook(path))}


def r_reload(b):
    with WB_LOCK:
        wb = WB
    if not wb:
        raise ApiError("Chưa mở file nào.")
    return {"workbook": wb_view(open_workbook(wb["file"]))}


def r_upload(b):
    name = os.path.basename(b.get("name") or "file.xlsx")
    if not name.lower().endswith((".xlsx", ".xlsm")):
        raise ApiError("Chỉ nhận file .xlsx / .xlsm")
    try:
        raw = base64.b64decode(b.get("data") or "")
    except Exception:
        raise ApiError("File gửi lên bị lỗi")
    up = os.path.join(DATA, "tai-len")
    os.makedirs(up, exist_ok=True)
    path = os.path.join(up, name)
    with open(path, "wb") as f:
        f.write(raw)
    wb = open_workbook(path)
    v = wb_view(wb)
    v["uploaded"] = True
    return {"workbook": v}


def r_template(b):
    with_sample = bool(b.get("sample"))
    default = "TestCase_Mau_CoViDu.xlsx" if with_sample else "TestCase_Mau.xlsx"
    path = b.get("path") or tk_dialog("save", title="Lưu file Excel mẫu", initialfile=default,
                                      initialdir=documents_dir(), defaultextension=".xlsx",
                                      filetypes=[("Excel", "*.xlsx")])
    if not path:
        return {"cancel": True}
    if not path.lower().endswith(".xlsx"):
        path += ".xlsx"
    prof = active_profile()
    if with_sample:
        # file ví dụ luôn có đủ các cột ví dụ dùng tới
        prof = json.loads(json.dumps(prof))
        need = {"tc_id", "scenario_id", "module", "tieu_de", "loai_test", "priority", "precondition", "test_data",
                "cach_test", "kich_thuoc", "step_no", "buoc", "expected", "hanh_dong", "doi_tuong", "gia_tri",
                "ghi_chu", "ts_id", "nhom", "scenario", "ac"}
        for kind in ("scenario", "testcase"):
            for x in prof[kind]:
                if x["key"] in need:
                    x["on"] = True
    X.create_template(path, prof, sample.sample_data() if with_sample else None)
    EXPORTED.add(os.path.normcase(os.path.abspath(path)))
    res = {"ok": True, "path": path}
    if b.get("open_in_tool", with_sample):
        res["workbook"] = wb_view(open_workbook(path))
    return res


def r_run(b):
    return start_run(b)


def r_stop(b):
    if RUNNER:
        RUNNER.stop()
    return {"ok": True}


def r_manual(b):
    if not RUNNER:
        raise ApiError("Không có lần chạy nào")
    if not RUNNER.answer_manual(b):
        raise ApiError("Câu hỏi này đã hết hiệu lực (có thể đã trả lời).")
    return {"ok": True}


def r_image(b):
    rd = run_path(b.get("run_id", ""))
    m = re.match(r"^data:image/(png|jpeg|jpg|gif|webp);base64,(.+)$", b.get("data") or "", re.S)
    if not m:
        raise ApiError("Ảnh không hợp lệ")
    ext = "jpg" if m.group(1) in ("jpeg", "jpg") else m.group(1)
    os.makedirs(os.path.join(rd, "anh"), exist_ok=True)
    name = "dan_%s_%03d.%s" % (time.strftime("%H%M%S"), int(time.time() * 1000) % 1000, ext)
    with open(os.path.join(rd, "anh", name), "wb") as f:
        f.write(base64.b64decode(m.group(2)))
    return {"ok": True, "rel": "anh/" + name}


def r_export(b):
    return do_export(b)


def r_preview_name(b):
    data = None
    run_id = b.get("run_id")
    if run_id:
        try:
            data = load_result(run_id)
        except ApiError:
            data = None
    if not data:
        with WB_LOCK:
            wb = WB
        data = {"project": wb["project"] if wb else "DuAn", "results": [], "env": SETTINGS.get("env", ""),
                "started": dt.datetime.now().isoformat(timespec="seconds"),
                "options": {"tester": SETTINGS.get("tester", ""), "browser": SETTINGS.get("browser", "")}}
    ctx = R.name_context(data, SETTINGS, run_number(data) if data.get("run_id") else 1)
    name = R.build_name(b.get("pattern"), ctx)
    folder = b.get("folder") or SETTINGS["export"].get("folder")
    return {"name": name, "exists": bool(folder) and os.path.exists(os.path.join(folder, name)), "ctx": ctx}


def r_pick_folder(b):
    p = tk_dialog("folder", title="Chọn thư mục lưu kết quả",
                  initialdir=SETTINGS["export"].get("folder") or documents_dir())
    if not p:
        return {"cancel": True}
    SETTINGS["export"]["folder"] = os.path.normpath(p)
    save_settings()
    return {"folder": SETTINGS["export"]["folder"]}


def r_open_path(b):
    p = os.path.normcase(os.path.abspath(b.get("path") or ""))
    allowed = p in EXPORTED or p.startswith(os.path.normcase(os.path.abspath(DATA))) or (
        os.path.normcase(os.path.abspath(SETTINGS["export"].get("folder") or "")) == p)
    if WB and p == os.path.normcase(os.path.abspath(WB["file"])):
        allowed = True
    if not allowed or not os.path.exists(b.get("path") or ""):
        raise ApiError("Không mở được đường dẫn này")
    open_path(b["path"])
    return {"ok": True}


def r_baseline(b):
    """Dùng ảnh của lần chạy này làm ẢNH CHUẨN mới cho bước so_sánh_ảnh."""
    data = load_result(b.get("run_id", ""))
    rd = run_path(b.get("run_id", ""))
    for r in data["results"]:
        for s in r["steps"]:
            if s.get("baseline") and b.get("image") in s.get("images", []):
                src = os.path.join(rd, b["image"])
                if not os.path.isfile(src) or "_khac" in b["image"]:
                    raise ApiError("Hãy chọn ảnh chụp gốc (không phải ảnh tô đỏ).")
                os.makedirs(os.path.dirname(s["baseline"]), exist_ok=True)
                import shutil
                shutil.copyfile(src, s["baseline"])
                return {"ok": True}
    raise ApiError("Ảnh này không thuộc bước so sánh ảnh.")


def r_quit(b):
    log.info("Tắt tool")
    if RUNNER:
        RUNNER.stop()
    threading.Timer(0.5, lambda: os._exit(0)).start()
    return {"ok": True}


ROUTES = {
    "/api/settings": r_settings, "/api/profile/save": r_profile_save, "/api/profile/activate": r_profile_activate,
    "/api/profile/delete": r_profile_delete, "/api/open": r_open, "/api/reload": r_reload, "/api/upload": r_upload,
    "/api/template": r_template, "/api/run": r_run, "/api/run/stop": r_stop, "/api/run/manual": r_manual,
    "/api/run/image": r_image, "/api/export": r_export, "/api/export/preview": r_preview_name,
    "/api/pick_folder": r_pick_folder, "/api/open_path": r_open_path, "/api/baseline": r_baseline,
    "/api/quit": r_quit, "/api/scan/save": r_scan_save, "/api/scan/discover": r_scan_discover,
    "/api/scan/run": r_scan_run, "/api/scan/discovery": r_scan_discovery, "/api/scan/plan": r_scan_plan,
    "/api/scan/import": r_scan_import, "/api/scan/export": r_scan_export, "/api/scan/fill": r_scan_fill,
    "/api/xlsx/sheets": r_xlsx_sheets,
}


# ----------------------------------------------------------------------------- khởi động
def ping(port):
    try:
        import urllib.request
        with urllib.request.urlopen("http://127.0.0.1:%d/api/ping" % port, timeout=1) as r:
            return json.loads(r.read().decode()).get("app") == "tooltest"
    except Exception:
        return False


def find_edge():
    for base in (os.environ.get("ProgramFiles(x86)"), os.environ.get("ProgramFiles"), os.environ.get("LOCALAPPDATA")):
        if base:
            p = os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe")
            if os.path.isfile(p):
                return p
    for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)"), os.environ.get("LOCALAPPDATA")):
        if base:
            p = os.path.join(base, "Google", "Chrome", "Application", "chrome.exe")
            if os.path.isfile(p):
                return p
    return None


def open_window(url):
    if os.environ.get("TOOLTEST_NO_BROWSER"):
        return
    exe = find_edge() if IS_WIN else None
    if exe:
        try:
            subprocess.Popen([exe, "--app=" + url, "--window-size=1400,900"], close_fds=True)
            return
        except Exception:
            pass
    try:
        if IS_WIN:
            os.startfile(url)
        else:
            webbrowser.open(url)
    except Exception:
        pass


def message_box(text, error=False):
    if sys.stdout:
        try:
            print(text)
        except Exception:
            pass
    if IS_WIN and FROZEN:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, text, APP_NAME, 0x10 if error else 0x40)
        except Exception:
            pass


def default_data_dir():
    if os.environ.get("TOOLTEST_DATA"):
        return os.environ["TOOLTEST_DATA"]
    base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Local")
    return os.path.join(base, DATA_FOLDER)


def main(argv=None):
    global DATA, PORT
    ap = argparse.ArgumentParser(description=APP_NAME)
    ap.add_argument("--port", type=int, default=int(os.environ.get("TOOLTEST_PORT", DEFAULT_PORT)))
    ap.add_argument("--data", default=None)
    ap.add_argument("--no-browser", action="store_true")
    args, _ = ap.parse_known_args(argv)
    if args.no_browser:
        os.environ["TOOLTEST_NO_BROWSER"] = "1"
    DATA = os.path.abspath(args.data or default_data_dir())
    try:
        os.makedirs(runs_dir(), exist_ok=True)
    except OSError as e:
        message_box("Không tạo được thư mục dữ liệu:\n%s\n\n%s" % (DATA, e), error=True)
        return 1
    fh = RotatingFileHandler(os.path.join(DATA, "app.log"), maxBytes=1_000_000, backupCount=2, encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    log.addHandler(fh)
    if sys.stdout:
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
        log.addHandler(logging.StreamHandler(sys.stdout))

    run_file = os.path.join(DATA, "running.json")
    try:
        with open(run_file, encoding="utf-8") as f:
            old = json.load(f)
        if ping(int(old.get("port") or 0)):     # tool đã mở → chỉ mở lại cửa sổ
            open_window("http://127.0.0.1:%d/" % int(old["port"]))
            return 0
    except Exception:
        pass
    load_settings()
    httpd = None
    for p in range(args.port, args.port + 15):
        try:
            httpd = ThreadingHTTPServer(("127.0.0.1", p), Handler)
            PORT = p
            break
        except OSError:
            continue
    if not httpd:
        message_box("Không tìm thấy cổng mạng trống (%d–%d)." % (args.port, args.port + 14), error=True)
        return 1
    httpd.daemon_threads = True
    try:
        with open(run_file, "w", encoding="utf-8") as f:
            json.dump({"port": PORT, "pid": os.getpid()}, f)
    except OSError:
        pass
    if SETTINGS.get("last_file") and os.path.isfile(SETTINGS["last_file"]):
        try:
            open_workbook(SETTINGS["last_file"])
        except Exception as e:
            log.warning("Không mở lại được file lần trước: %s", e)
    url = "http://127.0.0.1:%d/" % PORT
    log.info("%s v%s — %s — dữ liệu: %s", APP_NAME, VERSION, url, DATA)
    threading.Timer(0.3, lambda: open_window(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
