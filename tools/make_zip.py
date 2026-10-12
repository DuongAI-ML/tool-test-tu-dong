"""Nén dist/ToolTestTuDong → dist/ToolTestTuDong_v<phiên bản>_Windows.zip (file gửi cho người dùng / đưa lên GitHub
Releases)."""
import os
import re
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "dist", "ToolTestTuDong")
ver = re.search(r'^VERSION = "([^"]+)"', open(os.path.join(ROOT, "app.py"), encoding="utf-8").read(), re.M).group(1)
out = os.path.join(ROOT, "dist", "ToolTestTuDong_v%s_Windows.zip" % ver)
if not os.path.isfile(os.path.join(SRC, "ToolTestTuDong.exe")):
    raise SystemExit("Chưa có %s — hãy chạy build_exe.bat" % SRC)
if os.path.exists(out):
    os.remove(out)
n = 0
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
    for d, _, files in os.walk(SRC):
        for f in files:
            full = os.path.join(d, f)
            z.write(full, os.path.join("ToolTestTuDong", os.path.relpath(full, SRC)))
            n += 1
print("Đã nén %d file → %s (%.1f MB)" % (n, out, os.path.getsize(out) / 1048576))
