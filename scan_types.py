"""Các LOẠI TEST của chế độ "Quét từ URL" và tùy chọn riêng của từng loại (giao diện tự vẽ dropdown từ đây)."""

YES_NO = [["yes", "Có"], ["no", "Không"]]
REVIEW = {"key": "review", "label": "Người duyệt", "type": "select", "default": "no",
          "choices": [["no", "Không — tool tự đánh giá"], ["yes", "Có — tool chụp ảnh rồi DỪNG cho tôi xem"]]}
UNSURE = {"key": "unsure", "label": "Khi tool không chắc kết quả", "type": "select", "default": "ask",
          "choices": [["ask", "Dừng lại hỏi tôi Pass/Fail"], ["later", "Ghi 'Chưa test' + ảnh để tôi xem sau"]]}

TYPES = [
    {"key": "functional", "name": "Functional", "vi": "Chức năng", "prefix": "FUN", "priority": "High", "on": True,
     "desc": "Trang mở được, không lỗi JS, link không hỏng, nút bấm chạy, tìm kiếm, phân trang, sắp xếp cột.",
     "options": [
         {"key": "js_error", "label": "Lỗi JavaScript", "type": "select", "default": "fail",
          "choices": [["fail", "Có lỗi là Fail"], ["warn", "Chỉ ghi chú, không Fail"]]},
         {"key": "links", "label": "Kiểm tra link", "type": "select", "default": "internal",
          "choices": [["internal", "Link trong web"], ["all", "Cả link ra ngoài"], ["none", "Không kiểm tra"]]},
         {"key": "buttons", "label": "Bấm thử các nút (cột Hành động, thanh công cụ…)", "type": "select",
          "default": "safe", "choices": [["safe", "Chỉ nút an toàn (bỏ qua Xóa, Lưu, Gửi, Duyệt, Đăng xuất…)"],
                                         ["none", "Không bấm"]]},
         {"key": "max_buttons", "label": "Số nút tối đa mỗi trang", "type": "select", "default": "15",
          "choices": [["5", "5"], ["15", "15"], ["30", "30"]]},
         {"key": "search", "label": "Tìm kiếm theo từ khóa", "type": "select", "default": "yes", "choices": YES_NO},
         {"key": "paging", "label": "Phân trang", "type": "select", "default": "yes", "choices": YES_NO},
         {"key": "sorting", "label": "Sắp xếp theo cột", "type": "select", "default": "yes", "choices": YES_NO},
         {"key": "form_actions", "label": "Bấm các nút của form (Lưu, Lưu nháp, Gửi phê duyệt, Hủy…) với dữ liệu hợp lệ",
          "type": "select", "default": "no",
          "choices": [["no", "Không (không tạo dữ liệu)"], ["yes", "Có — sẽ TẠO dữ liệu thật (chỉ môi trường TEST)"]]},
     ]},
    {"key": "integration", "name": "Integration", "vi": "Tích hợp Frontend – Backend", "prefix": "INT",
     "priority": "High", "on": True,
     "desc": "Dropdown có bị trống (backend chưa trả dữ liệu), API gọi lên có lỗi, dữ liệu trên trang có khớp API, "
             "biểu tượng đang tải có bị treo.",
     "options": [
         {"key": "dropdowns", "label": "Dropdown có dữ liệu (kể cả dropdown trong popup)", "type": "select",
          "default": "yes", "choices": YES_NO},
         {"key": "api_errors", "label": "API trả lỗi (4xx/5xx, mất kết nối)", "type": "select", "default": "yes",
          "choices": YES_NO},
         {"key": "api_map", "label": "Dữ liệu hiển thị khớp dữ liệu API trả về", "type": "select", "default": "yes",
          "choices": YES_NO},
         {"key": "loading", "label": "Biểu tượng 'đang tải' không bị treo quá", "type": "select", "default": "10",
          "choices": [["5", "5 giây"], ["10", "10 giây"], ["20", "20 giây"], ["no", "Không kiểm tra"]]},
     ]},
    {"key": "validation", "name": "Validation", "vi": "Kiểm tra dữ liệu nhập", "prefix": "VAL", "priority": "High",
     "on": True, "desc": "Tự tìm các form, thử bỏ trống ô bắt buộc (*), sai định dạng, vượt độ dài, số ngoài khoảng, "
                         "tải tệp sai định dạng.",
     "options": [
         {"key": "submit", "label": "Bấm nút gửi form để kiểm tra", "type": "select", "default": "yes",
          "choices": [["yes", "Có (chỉ dùng trên môi trường TEST)"], ["no", "Không — chỉ kiểm tra thuộc tính ô"]]},
         {"key": "cases", "label": "Các ca kiểm tra", "type": "multi",
          "default": ["empty_all", "required", "email", "phone", "maxlength", "number_range", "file_type"],
          "choices": [["empty_all", "Gửi form trống"], ["required", "Bỏ trống từng ô bắt buộc (*)"],
                      ["email", "Email sai định dạng"], ["phone", "Số điện thoại sai định dạng"],
                      ["maxlength", "Vượt độ dài tối đa"], ["number_range", "Số ngoài khoảng min/max"],
                      ["file_type", "Tải tệp sai định dạng"], ["spaces", "Chỉ nhập khoảng trắng"],
                      ["valid", "Dữ liệu hợp lệ — gửi thật (sẽ TẠO dữ liệu)"]]},
         {"key": "max_cases", "label": "Số ca tối đa mỗi form", "type": "select", "default": "20",
          "choices": [["10", "10"], ["20", "20"], ["50", "50"]]},
         UNSURE,
     ]},
    {"key": "business", "name": "Business rule", "vi": "Luật nghiệp vụ", "prefix": "BR", "priority": "High",
     "on": False, "rules": True,
     "desc": "Bạn chọn trường + luật bằng dropdown (bắt buộc, độ dài, khoảng số, email, SĐT, ngày…), tool tự sinh ca "
             "đúng / sai / giá trị biên và chạy.",
     "options": [
         {"key": "valid_cases", "label": "Kiểm tra cả ca hợp lệ (giá trị biên trong)", "type": "select",
          "default": "no", "choices": [["no", "Không — không tạo dữ liệu"], ["yes", "Có — sẽ gửi form thật"]]},
         UNSURE,
     ]},
    {"key": "state", "name": "State", "vi": "Trạng thái", "prefix": "STA", "priority": "Medium", "on": True,
     "state_rules": True,
     "desc": "Mở/đóng bộ lọc, menu, toggle, popup. Thêm LUỒNG TRẠNG THÁI: bấm 'Lưu nháp' → đơn ở trạng thái 'Nháp', "
             "bấm 'Gửi phê duyệt' → 'Chờ duyệt'… (khai báo bên dưới, sẽ tạo dữ liệu thật).",
     "options": [
         {"key": "toggles", "label": "Nút bật/tắt, mở rộng/thu gọn, menu", "type": "select", "default": "yes",
          "choices": YES_NO},
         {"key": "popups", "label": "Popup: mở rồi đóng bằng nút Đóng / X", "type": "select", "default": "yes",
          "choices": YES_NO},
         {"key": "esc", "label": "Popup đóng được bằng phím Esc", "type": "select", "default": "no",
          "choices": [["no", "Không kiểm tra"], ["yes", "Có kiểm tra"]]},
     ]},
    {"key": "permission", "name": "Permission", "vi": "Phân quyền", "prefix": "PER", "priority": "High", "on": False,
     "desc": "Đăng nhập bằng TÀI KHOẢN THỨ 2 (VD Viewer), mở lại các trang đã quét và kiểm tra quyền.",
     "options": [
         {"key": "expect", "label": "Tài khoản thứ 2 được phép", "type": "select", "default": "view_only",
          "choices": [["view_only", "Chỉ xem — không thấy nút Thêm/Sửa/Xóa/Duyệt/Nhập"],
                      ["no_access", "Không được vào các trang này"], ["same", "Giống tài khoản chính"]]},
         {"key": "login_mode", "label": "Đăng nhập tài khoản 2", "type": "select", "default": "manual",
          "choices": [["manual", "Tôi tự đăng nhập trong trình duyệt"], ["auto", "Tool tự điền tài khoản bên dưới"],
                      ["none", "Không đăng nhập (khách)"]]},
         {"key": "user", "label": "Tài khoản 2", "type": "text", "default": ""},
         {"key": "pass", "label": "Mật khẩu 2 (không lưu lại)", "type": "password", "default": ""},
         {"key": "deny", "label": "Nút tài khoản 2 KHÔNG được thấy (cách nhau bằng ;)", "type": "text",
          "default": "", "hint": "VD: Lưu & Gửi phê duyệt; Phê duyệt; Xóa"},
         {"key": "must", "label": "Nút tài khoản 2 PHẢI thấy (cách nhau bằng ;)", "type": "text", "default": "",
          "hint": "VD: Lưu nháp"},
         {"key": "popups", "label": "Kiểm tra cả nút trong popup / form", "type": "select", "default": "yes",
          "choices": YES_NO},
     ]},
    {"key": "usability", "name": "Usability", "vi": "Dễ sử dụng", "prefix": "USA", "priority": "Medium", "on": True,
     "desc": "Nút / nút gạt có bấm được không, có bị thanh điều khiển hay phần tử khác che khuất không, nút quá nhỏ "
             "khó bấm, chữ quá nhỏ khó đọc.",
     "options": [
         {"key": "checks", "label": "Kiểm tra", "type": "multi", "default": ["covered", "target", "switch", "font"],
          "choices": [["covered", "Nút / ô nhập bị che khuất"], ["target", "Nút quá nhỏ khó bấm (< 24px)"],
                      ["switch", "Nút gạt / checkbox bấm được"], ["font", "Chữ quá nhỏ (< 12px)"]]},
         {"key": "where", "label": "Kiểm tra ở màn hình", "type": "select", "default": "both",
          "choices": [["desktop", "Máy tính"], ["mobile", "Điện thoại (375×812)"], ["both", "Cả hai"]]},
     ]},
    {"key": "ui", "name": "UI", "vi": "Giao diện", "prefix": "UI", "priority": "Medium",
     "on": True, "desc": "Ảnh lỗi, hiện chữ lỗi (undefined, null, NaN, {{…}}), tràn ngang, tiêu đề trang; chụp ảnh "
                         "từng trang, so với lần quét trước.",
     "options": [
         {"key": "checks", "label": "Kiểm tra", "type": "multi", "default": ["images", "badtext", "overflow", "title"],
          "choices": [["images", "Ảnh bị lỗi"], ["badtext", "Hiện chữ lỗi (undefined, null, NaN, {{…}})"],
                      ["overflow", "Bị tràn ngang"], ["title", "Có tiêu đề tab"]]},
         {"key": "compare", "label": "So với ảnh lần quét trước", "type": "select", "default": "no",
          "choices": [["no", "Không"], ["1", "Có — khác > 1% là Fail"], ["3", "Có — khác > 3%"],
                      ["5", "Có — khác > 5%"]]},
         {"key": "shot", "label": "Ảnh chụp", "type": "select", "default": "full",
          "choices": [["full", "Cả trang dài"], ["view", "Phần màn hình"]]},
         REVIEW,
     ]},
    {"key": "responsive", "name": "Responsive", "vi": "Nhiều kích thước màn hình", "prefix": "RES",
     "priority": "High", "on": False,
     "desc": "Mở từng trang ở các kích thước: không tràn ngang, không phát sinh lỗi JavaScript khi đổi kích thước.",
     "options": [
         {"key": "sizes", "label": "Kích thước", "type": "multi", "default": ["1366x768", "768x1024", "375x812"],
          "choices": [["1920x1080", "Desktop 1920×1080"], ["1366x768", "Laptop 1366×768"],
                      ["768x1024", "Tablet 768×1024"], ["390x844", "iPhone 390×844"],
                      ["375x812", "Mobile 375×812"], ["360x800", "Android 360×800"]]},
         {"key": "custom", "label": "Kích thước khác (VD 1280x720, 412x915)", "type": "text", "default": ""},
         REVIEW,
     ]},
    {"key": "performance", "name": "Performance", "vi": "Hiệu năng cơ bản", "prefix": "PRF", "priority": "Medium",
     "on": True, "desc": "Thời gian tải, dung lượng, tài nguyên chậm.",
     "options": [
         {"key": "max_load", "label": "Thời gian tải tối đa", "type": "select", "default": "3",
          "choices": [["1", "1 giây"], ["2", "2 giây"], ["3", "3 giây"], ["5", "5 giây"], ["10", "10 giây"]]},
         {"key": "max_weight", "label": "Dung lượng trang tối đa", "type": "select", "default": "5",
          "choices": [["1", "1 MB"], ["2", "2 MB"], ["5", "5 MB"], ["10", "10 MB"], ["0", "Không giới hạn"]]},
         {"key": "runs", "label": "Số lần đo", "type": "select", "default": "1",
          "choices": [["1", "1 lần"], ["3", "3 lần (lấy trung bình)"]]},
     ]},
    {"key": "security", "name": "Security", "vi": "Bảo mật cơ bản", "prefix": "SEC", "priority": "High", "on": False,
     "desc": "Vào trang khi chưa đăng nhập, lộ lỗi hệ thống, header bảo mật, cookie phiên, HTTPS, ô mật khẩu.",
     "options": [
         {"key": "checks", "label": "Kiểm tra", "type": "multi",
          "default": ["anon", "errors", "password", "cookies", "headers", "https"],
          "choices": [["anon", "Vào trang khi CHƯA đăng nhập"], ["errors", "Lộ thông báo lỗi hệ thống"],
                      ["password", "Ô mật khẩu an toàn"], ["cookies", "Cookie phiên (HttpOnly, Secure)"],
                      ["headers", "Header bảo mật"], ["https", "Dùng HTTPS"]]},
     ]},
    {"key": "accessibility", "name": "Accessibility", "vi": "Khả năng truy cập", "prefix": "ACC", "priority": "Low",
     "on": False, "desc": "Ảnh thiếu mô tả, ô nhập thiếu nhãn, nút không có tên, ngôn ngữ trang, tiêu đề, tương phản.",
     "options": [
         {"key": "level", "label": "Độ tương phản chữ", "type": "select", "default": "AA",
          "choices": [["AA", "WCAG AA (4.5:1)"], ["A", "Nới lỏng (3:1)"], ["off", "Không kiểm tra"]]},
     ]},
    {"key": "compat", "name": "Compatibility", "vi": "Tương thích trình duyệt", "prefix": "CMP", "priority": "Medium",
     "on": False, "desc": "Mở lại các trang trên trình duyệt khác: tải được, không lỗi JS, không tràn.",
     "options": [
         {"key": "browsers", "label": "Trình duyệt", "type": "multi", "default": ["edge", "chrome"],
          "choices": [["edge", "Microsoft Edge"], ["chrome", "Google Chrome"]]},
     ]},
]
ORDER = ["validation", "functional", "integration", "business", "permission", "state", "usability", "ui",
         "responsive", "performance", "security", "accessibility", "compat"]   # 7 loại chính của tester đứng đầu
TYPES.sort(key=lambda t: ORDER.index(t["key"]))
BY_KEY = {t["key"]: t for t in TYPES}

RULES = [  # luật nghiệp vụ cho Business rule: key, tên, cần giá trị 1, cần giá trị 2, gợi ý
    ["required", "Bắt buộc nhập", "", "", ""],
    ["maxlen", "Độ dài tối đa N ký tự", "N", "", "VD 100"],
    ["minlen", "Độ dài tối thiểu N ký tự", "N", "", "VD 6"],
    ["range", "Số trong khoảng từ A đến B", "A", "B", "VD 1 và 100"],
    ["email", "Đúng định dạng email", "", "", ""],
    ["phone", "Đúng định dạng số điện thoại (10 số, bắt đầu 0)", "", "", ""],
    ["digits", "Chỉ được nhập số", "", "", ""],
    ["nospecial", "Không cho ký tự đặc biệt", "", "", ""],
    ["no_past", "Ngày không được ở quá khứ", "", "", ""],
    ["no_future", "Ngày không được ở tương lai", "", "", ""],
    ["after_field", "Ngày phải SAU ngày của trường khác", "Tên trường kia", "", "VD Ngày phỏng vấn"],
    ["calc", "Tự tính: nhập các trường → trường này phải bằng", "Trường A=giá trị; Trường B=giá trị",
     "Kết quả mong đợi", "VD: Điểm ngành nghề=8; Điểm kinh nghiệm=7 → 15"],
    ["reject", "Các giá trị sau phải BỊ TỪ CHỐI", "giá trị; giá trị", "", "cách nhau bằng ;"],
    ["accept", "Các giá trị sau phải ĐƯỢC CHẤP NHẬN (gửi thật)", "giá trị; giá trị", "", "cách nhau bằng ;"],
]
RULE_NAME = {r[0]: r[1] for r in RULES}

PRESETS = {
    "smoke": {"functional": True, "validation": False, "integration": True, "business": False, "state": False,
              "permission": False, "usability": False, "ui": True, "responsive": False, "performance": True,
              "security": False, "accessibility": False, "compat": False},
    "seven": {"functional": True, "validation": True, "integration": True, "business": True, "state": True,
              "permission": True, "usability": True, "ui": False, "responsive": False, "performance": False,
              "security": False, "accessibility": False, "compat": False},
    "full": {"functional": True, "validation": True, "integration": True, "business": True, "state": True,
             "permission": False, "usability": True, "ui": True, "responsive": True, "performance": True,
             "security": True, "accessibility": True, "compat": False},
}

UNSAFE_WORDS = ("xóa, xoá, delete, remove, hủy, huỷ, cancel, đăng xuất, logout, log out, sign out, thoát, thanh toán, "
                "pay, checkout, gửi, submit, send, lưu, save, xác nhận, confirm, duyệt, approve, từ chối, reject, khóa, "
                "khoá, lock, reset, import, nhập dữ liệu, upload, tải lên, đặt hàng, mua, buy, chuyển tiền, kích hoạt, "
                "activate, vô hiệu, disable, đồng ý, accept, cập nhật, update, phê duyệt, đóng đơn, kết thúc, gỡ, "
                "revoke, archive, lưu trữ")
SKIP_LINK_WORDS = "logout, log-out, signout, sign-out, dang-xuat, dangxuat, đăng xuất, thoát, delete, xoa, remove"


def default_type_opts():
    out = {}
    for t in TYPES:
        out[t["key"]] = {"on": t["on"], "opts": {o["key"]: o["default"] for o in t["options"]}}
    return out

# Cột "Nhóm" của Test Scenario — đúng chữ tester dùng trong file (ui / responsive gộp vào usability như file mẫu)
NHOM = {"functional": "functional", "validation": "validation", "integration": "integration",
        "business": "business rule", "permission": "permission", "state": "state", "usability": "usability",
        "ui": "usability", "responsive": "usability", "performance": "performance", "security": "security",
        "accessibility": "accessibility", "compat": "compatibility"}
