"""Danh mục cột chuẩn cho sheet Test Scenario và Test Case.

Mỗi cột: key (nội bộ), label (tên mặc định), group (nhóm hiển thị), level ("tc" = ghi ở dòng đầu của TC,
"step" = ghi ở mỗi dòng bước), result (True = tool tự điền sau khi chạy), syn (các tên khác để tự nhận cột khi
đọc Excel), default (bật sẵn trong mẫu mặc định).
"""
import re
import unicodedata


def norm(s):
    """'Kết quả mong đợi' -> 'ket qua mong doi' (bỏ dấu, chữ thường, bỏ ký tự lạ) để so khớp tên cột."""
    if s is None:
        return ""
    s = str(s).replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^a-zA-Z0-9]+", " ", s).strip().lower()
    return s


def _c(key, label, group, level="tc", result=False, syn=(), default=True):
    return {"key": key, "label": label, "group": group, "level": level, "result": result,
            "syn": list(syn), "default": default}


SCENARIO_COLUMNS = [
    _c("ts_id", "ID", "Định danh", syn=["id", "ma", "ma scenario", "scenario id", "ts id", "ma ts", "stt id"]),
    _c("nhom", "Nhóm", "Mô tả", syn=["nhom", "group", "category", "nhom test", "loai", "loai test", "test type"]),
    _c("module", "Module/Chức năng", "Định danh", syn=["module", "chuc nang", "feature", "man hinh"], default=False),
    _c("scenario", "Test Scenario", "Mô tả", syn=["test scenario", "scenario", "kich ban", "kich ban test", "mo ta",
                                                   "noi dung", "description", "ten scenario"]),
    _c("ac", "AC liên quan", "Định danh", syn=["ac", "ac lien quan", "acceptance criteria", "tieu chi chap nhan"]),
    _c("requirement", "Requirement/User Story", "Định danh",
       syn=["requirement", "user story", "us", "req", "yeu cau", "req id", "story"], default=False),
    _c("priority", "Priority", "Mô tả", syn=["priority", "do uu tien", "muc uu tien", "uu tien"]),
    _c("rui_ro", "Mức rủi ro", "Mô tả", syn=["muc rui ro", "rui ro", "risk"], default=False),
    _c("luong", "Luồng", "Mô tả", syn=["luong", "luong nghiep vu", "flow"], default=False),
    _c("cach_test", "Cách test", "Mô tả", syn=["cach test", "tu dong", "automation", "hinh thuc test"],
       default=False),
    _c("nguoi_viet", "Người viết", "Quản lý", syn=["nguoi viet", "nguoi tao", "author", "created by"], default=False),
    _c("ngay_tao", "Ngày tạo", "Quản lý", syn=["ngay tao", "ngay viet", "created date", "created"], default=False),
    _c("trang_thai", "Trạng thái", "Quản lý", syn=["trang thai", "state"], default=False),
    _c("ghi_chu", "Ghi chú", "Quản lý", syn=["ghi chu", "note", "notes", "comment", "remark"], default=False),
    # ---- tool tự điền
    _c("tong_tc", "Tổng TC", "Kết quả", result=True, syn=["tong tc", "so tc", "total"]),
    _c("so_pass", "Pass", "Kết quả", result=True, syn=["pass", "so pass", "passed"]),
    _c("so_fail", "Fail", "Kết quả", result=True, syn=["fail", "so fail", "failed"]),
    _c("so_blocked", "Blocked", "Kết quả", result=True, syn=["blocked", "so blocked"]),
    _c("so_chua_test", "Chưa test", "Kết quả", result=True, syn=["chua test", "not run", "chua chay"]),
    _c("ty_le", "Tỷ lệ pass (%)", "Kết quả", result=True, syn=["ty le pass", "ty le", "pass rate"]),
    _c("ket_luan", "Kết luận", "Kết quả", result=True, syn=["ket luan", "ket qua", "result", "status"]),
    _c("ngay_chay", "Ngày chạy", "Kết quả", result=True, syn=["ngay chay", "ngay test"], default=False),
    _c("nguoi_chay", "Người chạy", "Kết quả", result=True, syn=["nguoi chay", "nguoi test", "tester"], default=False),
]

TESTCASE_COLUMNS = [
    _c("tc_id", "TC ID", "Định danh", syn=["tc id", "id", "ma tc", "ma test case", "test case id", "tcid", "ma"]),
    _c("scenario_id", "Scenario ID", "Định danh", syn=["scenario id", "ts id", "ma scenario", "scenario", "ma ts",
                                                        "test scenario id"]),
    _c("requirement", "Requirement/AC", "Định danh", syn=["requirement", "ac", "ac lien quan", "user story", "req",
                                                          "yeu cau", "requirement scenario"], default=False),
    _c("module", "Module", "Định danh", syn=["module", "chuc nang", "feature"]),
    _c("man_hinh", "Màn hình", "Định danh", syn=["man hinh", "screen", "page", "trang"], default=False),
    _c("tieu_de", "Tiêu đề", "Mô tả", syn=["tieu de", "ten test case", "ten tc", "test case", "title",
                                            "test case title", "summary", "muc dich"]),
    _c("mo_ta", "Mô tả", "Mô tả", syn=["mo ta", "description", "noi dung"], default=False),
    _c("loai_test", "Loại test", "Mô tả", syn=["loai test", "test type", "loai", "nhom", "type", "kieu test"]),
    _c("ky_thuat", "Kỹ thuật test", "Mô tả", syn=["ky thuat test", "ky thuat", "technique"], default=False),
    _c("priority", "Priority", "Mô tả", syn=["priority", "do uu tien", "muc uu tien", "uu tien"]),
    _c("severity", "Severity", "Mô tả", syn=["severity", "muc do nghiem trong", "muc do"], default=False),
    _c("tag", "Tag", "Mô tả", syn=["tag", "tags", "nhan"], default=False),
    _c("precondition", "Precondition", "Điều kiện", syn=["precondition", "preconditions", "pre condition",
                                                         "dieu kien tien quyet",
                                                         "dieu kien", "tien dieu kien", "pre requisite"]),
    _c("data_id", "Data ID", "Điều kiện", syn=["data id", "ma du lieu", "ma data", "id du lieu"], default=False),
    _c("test_data", "Test data", "Điều kiện", syn=["test data", "du lieu test", "du lieu", "data"]),
    _c("postcondition", "Postcondition", "Điều kiện", syn=["postcondition", "post condition", "dieu kien sau"],
       default=False),
    _c("cach_test", "Cách test", "Tự động hóa", syn=["cach test", "tu dong", "automation", "automated",
                                                     "automation status", "hinh thuc test"]),
    _c("moi_truong", "Môi trường", "Cấu hình chạy", syn=["moi truong", "environment", "env"], default=False),
    _c("trinh_duyet", "Trình duyệt", "Cấu hình chạy", syn=["trinh duyet", "browser"], default=False),
    _c("kich_thuoc", "Kích thước màn hình", "Cấu hình chạy", syn=["kich thuoc man hinh", "kich thuoc", "viewport",
                                                                   "do phan giai", "resolution", "screen size"]),
    _c("build", "Build/Version", "Cấu hình chạy", syn=["build", "version", "phien ban", "build version"],
       default=False),
    _c("step_no", "Step #", "Bước", level="step", syn=["step", "step no", "buoc so", "stt buoc", "so buoc", "#",
                                                       "step #", "stt"]),
    _c("buoc", "Bước thực hiện", "Bước", level="step", syn=["buoc thuc hien", "cac buoc thuc hien", "test steps",
                                                            "steps", "buoc", "thao tac", "step description",
                                                            "cac buoc"]),
    _c("du_lieu_buoc", "Dữ liệu bước", "Bước", level="step", syn=["du lieu buoc", "step data", "input"],
       default=False),
    _c("expected", "Expected result", "Bước", level="step", syn=["expected result", "expected", "ket qua mong doi",
                                                                 "kq mong doi", "ket qua mong muon"]),
    _c("hanh_dong", "Hành động", "Tự động hóa", level="step", syn=["hanh dong", "action", "keyword", "tu khoa"]),
    _c("doi_tuong", "Đối tượng", "Tự động hóa", level="step", syn=["doi tuong", "target", "element", "locator",
                                                                    "selector"]),
    _c("gia_tri", "Giá trị", "Tự động hóa", level="step", syn=["gia tri", "value", "gia tri nhap"]),
    # ---- tool tự điền
    _c("actual", "Actual result", "Kết quả", level="step", result=True,
       syn=["actual result", "actual", "ket qua thuc te", "kq thuc te"]),
    _c("status_buoc", "Status bước", "Kết quả", level="step", result=True, syn=["status buoc", "step status"]),
    _c("anh", "Ảnh chụp", "Kết quả", level="step", result=True, syn=["anh chup", "anh", "screenshot", "evidence",
                                                                      "bang chung"]),
    _c("kq_lan1", "Kết quả test (Lần 1)", "Kết quả", result=True, syn=["ket qua test lan 1", "lan 1", "ket qua lan 1",
                                                                   "round 1", "kq lan 1"], default=False),
    _c("kq_lan2", "Kết quả test (Lần 2)", "Kết quả", result=True, syn=["ket qua test lan 2", "lan 2", "ket qua lan 2",
                                                                   "round 2", "kq lan 2"], default=False),
    _c("kq_lan3", "Kết quả test (Lần 3)", "Kết quả", result=True, syn=["ket qua test lan 3", "lan 3", "ket qua lan 3",
                                                                   "round 3", "kq lan 3"], default=False),
    _c("status", "Status", "Kết quả", result=True, syn=["status", "trang thai", "ket qua", "result", "pass fail",
                                                        "trang thai test tong", "trang thai tong", "ket qua tong"]),
    _c("thoi_gian", "Thời gian chạy", "Kết quả", result=True, syn=["thoi gian chay", "thoi gian", "duration"],
       default=False),
    _c("lan_chay", "Lần chạy", "Quản lý", syn=["lan chay", "round", "cycle", "lan test"], default=False),
    _c("nguoi_viet", "Người viết", "Quản lý", syn=["nguoi viet", "nguoi tao", "author", "created by"], default=False),
    _c("ngay_viet", "Ngày viết", "Quản lý", syn=["ngay viet", "ngay tao", "created date"], default=False),
    _c("nguoi_review", "Người review", "Quản lý", syn=["nguoi review", "reviewer"], default=False),
    _c("nguoi_test", "Người test", "Kết quả", result=True, syn=["nguoi test", "tester", "executed by"]),
    _c("ngay_test", "Ngày test", "Kết quả", result=True, syn=["ngay test", "ngay chay", "executed date", "test date"]),
    _c("bug_id", "Bug ID", "Quản lý", syn=["bug id", "bug", "ma loi", "defect id", "defect"]),
    _c("link_bug", "Link bug", "Quản lý", syn=["link bug", "bug link", "link loi"], default=False),
    _c("ghi_chu", "Ghi chú", "Quản lý", syn=["ghi chu", "note", "notes", "comment", "remark"]),
]

CATALOG = {"scenario": SCENARIO_COLUMNS, "testcase": TESTCASE_COLUMNS}
BY_KEY = {k: {c["key"]: c for c in cols} for k, cols in CATALOG.items()}

PRIORITIES = ["Critical", "High", "Medium", "Low"]
SEVERITIES = ["Critical", "Major", "Minor", "Trivial"]
CACH_TEST = ["Tự động", "Bán tự động", "Tay"]

# Loại test: (tên, cách test mặc định khi cột "Cách test" để trống, ghi chú)
TEST_TYPES = [
    ("Functional", "Tự động", "Chức năng chạy đúng nghiệp vụ"),
    ("Validation", "Tự động", "Kiểm tra nhập liệu: bỏ trống, sai định dạng, quá độ dài"),
    ("Boundary", "Tự động", "Giá trị biên: min, max, min-1, max+1"),
    ("Negative", "Tự động", "Trường hợp nhập/thao tác sai"),
    ("Smoke", "Tự động", "Kiểm tra nhanh luồng chính"),
    ("Sanity", "Tự động", "Kiểm tra nhanh sau khi sửa"),
    ("Regression", "Tự động", "Chạy lại sau khi sửa lỗi"),
    ("E2E", "Tự động", "Luồng đầu–cuối qua nhiều màn hình"),
    ("Integration", "Tự động", "Tích hợp giữa các chức năng"),
    ("UI", "Bán tự động", "Giao diện: hiển thị, vị trí, text — tool so ảnh, người duyệt"),
    ("Responsive", "Bán tự động", "Thay đổi kích thước màn hình — tool chụp, người duyệt"),
    ("Compatibility", "Tự động", "Tương thích trình duyệt"),
    ("Usability", "Tay", "Dễ dùng, dễ nhìn, màu sắc — người đánh giá"),
    ("Performance", "Tự động", "Thời gian tải trang (cơ bản)"),
    ("Security", "Bán tự động", "Phân quyền, truy cập khi chưa đăng nhập (cơ bản)"),
    ("Accessibility", "Bán tự động", "Thiếu alt, nhãn, tương phản — tool quét, người duyệt"),
    ("Localization", "Bán tự động", "Ngôn ngữ, định dạng ngày/tiền"),
]
TYPE_DEFAULT_MODE = {norm(t[0]): t[1] for t in TEST_TYPES}
TYPE_DEFAULT_MODE.update({"giao dien": "Bán tự động", "kha nang su dung": "Tay", "chuc nang": "Tự động",
                          "hieu nang": "Tự động", "bao mat": "Bán tự động", "tuong thich": "Tự động",
                          "hoi quy": "Tự động", "gui": "Bán tự động", "ui ux": "Bán tự động", "ux": "Tay"})


def mode_of(text):
    """Chuẩn hóa giá trị cột 'Cách test' -> 'Tự động' | 'Bán tự động' | 'Tay' | '' (không rõ)."""
    n = norm(text)
    if not n:
        return ""
    if n.startswith("ban") or "semi" in n or "mot phan" in n:
        return "Bán tự động"
    if n in ("tay", "manual", "khong", "no", "n", "thu cong", "test tay") or n.startswith("tay") or "manual" in n:
        return "Tay"
    if n in ("tu dong", "auto", "automated", "automation", "co", "yes", "y", "x") or n.startswith("tu dong"):
        return "Tự động"
    return ""


def default_profile():
    return {
        "name": "Mặc định", "layout": "rows", "sections": False, "matrices": True,
        "scenario": [{"key": c["key"], "on": c["default"], "label": c["label"]} for c in SCENARIO_COLUMNS],
        "testcase": [{"key": c["key"], "on": c["default"], "label": c["label"]} for c in TESTCASE_COLUMNS],
    }


# Mẫu theo đúng file của tester (ảnh Test case + Test scenario đã gửi): 1 dòng / TC, các bước gộp 1 ô,
# có dòng nhóm theo loại test, cột Kết quả Lần 1/2/3 + Trạng thái Test tổng, cột Bug # chứa ảnh lỗi.
COMPANY_TC = [("tc_id", "Mã TC"), ("module", "Chức năng"), ("requirement", "Requirement / Scenario"),
              ("tieu_de", "Tên TC"), ("precondition", "Preconditions"), ("moi_truong", "Environment"),
              ("buoc", "Steps"), ("expected", "Kết quả mong đợi"), ("actual", "Kết quả thực tế"),
              ("data_id", "Data ID"), ("test_data", "Test data"), ("priority", "Priority"),
              ("kq_lan1", "Kết quả test (Lần 1)"), ("kq_lan2", "Kết quả test (Lần 2)"),
              ("kq_lan3", "Kết quả test (Lần 3)"), ("status", "Trạng thái Test tổng"), ("anh", "Bug #")]
COMPANY_TS = [("ts_id", "ID"), ("nhom", "Nhóm"), ("scenario", "Test Scenario"), ("ac", "AC liên quan"),
              ("priority", "Priority")]
COMPANY_NAME = "Mẫu công ty (theo file tester)"


def company_profile():
    def build(pairs, cols):
        keys = [k for k, _ in pairs]
        rows = [{"key": k, "on": True, "label": lab} for k, lab in pairs]
        rows += [{"key": c["key"], "on": False, "label": c["label"]} for c in cols if c["key"] not in keys]
        return rows
    return {"name": COMPANY_NAME, "layout": "merged", "sections": True, "matrices": True,
            "scenario": build(COMPANY_TS, SCENARIO_COLUMNS), "testcase": build(COMPANY_TC, TESTCASE_COLUMNS)}


def fix_profile(p):
    """Bổ sung cột mới (khi nâng cấp tool) / bỏ cột không còn vào hồ sơ cột đã lưu."""
    base = default_profile()
    out = {"name": (p or {}).get("name") or base["name"],
           "layout": (p or {}).get("layout") if (p or {}).get("layout") in ("rows", "merged") else "rows",
           "sections": bool((p or {}).get("sections")), "matrices": (p or {}).get("matrices", True) is not False}
    for kind in ("scenario", "testcase"):
        have = [x for x in (p or {}).get(kind, []) if isinstance(x, dict) and x.get("key") in BY_KEY[kind]]
        seen = {x["key"] for x in have}
        rows = [{"key": x["key"], "on": bool(x.get("on")), "label": str(x.get("label") or
                                                                       BY_KEY[kind][x["key"]]["label"])[:80]}
                for x in have]
        for x in base[kind]:
            if x["key"] not in seen:
                rows.append(dict(x, on=False))
        out[kind] = rows
    return out
