"""Dữ liệu cho file Excel MẪU CÓ VÍ DỤ — chạy được ngay trên web giả lập "Đơn hàng tuyển dụng" có sẵn trong tool."""

SCENARIOS = [
    {"ts_id": "TS_DHTD_001", "nhom": "Usability", "scenario": "Kiểm tra giao diện và bố cục mặc định của trang Đơn "
     "hàng tuyển dụng", "ac": "AC_01", "priority": "Medium"},
    {"ts_id": "TS_DHTD_002", "nhom": "Usability", "scenario": "Kiểm tra khả năng hiển thị và hoạt động của trang khi "
     "thay đổi kích thước màn hình", "ac": "AC_02", "priority": "High"},
    {"ts_id": "TS_DHTD_003", "nhom": "Functional", "scenario": "Kiểm tra tạo mới, tìm kiếm, xóa đơn hàng tuyển dụng",
     "ac": "AC_03", "priority": "High"},
    {"ts_id": "TS_DHTD_004", "nhom": "Validation", "scenario": "Kiểm tra ràng buộc dữ liệu form Tạo đơn hàng",
     "ac": "AC_04", "priority": "High"},
    {"ts_id": "TS_LOGIN_001", "nhom": "Functional", "scenario": "Kiểm tra đăng nhập hệ thống", "ac": "AC_00",
     "priority": "Critical"},
]


def _tc(tc, steps):
    rows = []
    for i, s in enumerate(steps, start=1):
        buoc, expected, action, target, value = (list(s) + ["", "", "", "", ""])[:5]
        row = {"step_no": i, "buoc": buoc, "expected": expected, "hanh_dong": action, "doi_tuong": target,
               "gia_tri": value}
        if i == 1:
            row.update(tc)
        rows.append(row)
    return rows


LOGIN = ("Đăng nhập bằng tài khoản admin", "Vào trang Đơn hàng tuyển dụng", "chạy_tc", "", "TC_LOGIN_01")

TESTCASES = (
    _tc({"tc_id": "TC_LOGIN_01", "scenario_id": "TS_LOGIN_001", "module": "Đăng nhập", "tieu_de": "Đăng nhập thành "
         "công với tài khoản đúng", "loai_test": "Functional", "priority": "Critical", "tag": "smoke",
         "precondition": "Có tài khoản admin", "test_data": "", "cach_test": "Tự động"}, [
        ("Mở trang đăng nhập", "Hiện form đăng nhập", "mở_trang", "", "${base_url}/login.html"),
        ("Nhập email", "", "nhập", "Ô Email", "${user}"),
        ("Nhập mật khẩu", "", "nhập", "Ô Mật khẩu", "${pass}"),
        ("Bấm nút Đăng nhập", "", "bấm", 'Nút "Đăng nhập"', ""),
        ("Kiểm tra vào trang chính", "Chuyển sang trang Đơn hàng tuyển dụng", "kiểm_tra_url", "", "don-hang.html"),
    ])
    + _tc({"tc_id": "TC_LOGIN_02", "scenario_id": "TS_LOGIN_001", "module": "Đăng nhập", "tieu_de": "Đăng nhập sai "
           "mật khẩu", "loai_test": "Negative", "priority": "High", "test_data": "sai_mk=abc999",
           "cach_test": "Tự động"}, [
        ("Mở trang đăng nhập", "", "mở_trang", "", "${base_url}/login.html"),
        ("Nhập email đúng", "", "nhập", "Ô Email", "${user}"),
        ("Nhập mật khẩu sai", "", "nhập", "Ô Mật khẩu", "${sai_mk}"),
        ("Bấm Đăng nhập", "", "bấm", 'Nút "Đăng nhập"', ""),
        ("Kiểm tra thông báo lỗi", "Hiện 'Email hoặc mật khẩu không đúng'", "kiểm_tra_text", "",
         "Email hoặc mật khẩu không đúng"),
        ("Vẫn ở trang đăng nhập", "URL vẫn là login", "kiểm_tra_url", "", "login.html"),
    ])
    + _tc({"tc_id": "TC_LOGIN_03", "scenario_id": "TS_LOGIN_001", "module": "Đăng nhập", "tieu_de": "Bỏ trống email "
           "và mật khẩu", "loai_test": "Validation", "priority": "Medium", "cach_test": "Tự động"}, [
        ("Mở trang đăng nhập", "", "mở_trang", "", "${base_url}/login.html"),
        ("Bấm Đăng nhập khi chưa nhập gì", "", "bấm", 'Nút "Đăng nhập"', ""),
        ("Kiểm tra lỗi email", "Hiện 'Vui lòng nhập email'", "kiểm_tra_text", "", "Vui lòng nhập email"),
        ("Kiểm tra lỗi mật khẩu", "Hiện 'Vui lòng nhập mật khẩu'", "kiểm_tra_text", "", "Vui lòng nhập mật khẩu"),
    ])
    + _tc({"tc_id": "TC_DHTD_001_01", "scenario_id": "TS_DHTD_001", "module": "Đơn hàng tuyển dụng",
           "tieu_de": "Bố cục mặc định hiển thị đủ thành phần", "loai_test": "UI", "priority": "Medium",
           "precondition": "Đã đăng nhập", "cach_test": "Bán tự động"}, [
        LOGIN,
        ("Kiểm tra lời chào", "Hiện 'Xin chào, Admin'", "kiểm_tra_text", "", "Xin chào"),
        ("Kiểm tra tiêu đề trang", "Tiêu đề 'Đơn hàng tuyển dụng'", "kiểm_tra_hiển_thị", 'Tiêu đề "Đơn hàng tuyển dụng"'),
        ("Kiểm tra ô tìm kiếm", "Có ô tìm kiếm", "kiểm_tra_hiển_thị", "Ô Tìm theo mã hoặc vị trí"),
        ("Kiểm tra nút Tạo đơn hàng", "Có nút Tạo đơn hàng", "kiểm_tra_hiển_thị", 'Nút "+ Tạo đơn hàng"'),
        ("Kiểm tra bảng có dữ liệu", "Bảng có 10 dòng (trang 1)", "kiểm_tra_số_lượng", "Dòng trong bảng", "10"),
        ("So sánh giao diện với ảnh chuẩn", "Giống ảnh chuẩn", "so_sánh_ảnh", "", "2"),
    ])
    + _tc({"tc_id": "TC_DHTD_001_02", "scenario_id": "TS_DHTD_001", "module": "Đơn hàng tuyển dụng",
           "tieu_de": "Màu sắc, cỡ chữ dễ nhìn, căn lề đều", "loai_test": "Usability", "priority": "Low",
           "precondition": "Đã đăng nhập", "cach_test": "Tay"}, [
        ("Quan sát màu sắc tổng thể", "Màu hài hòa, đúng nhận diện công ty"),
        ("Quan sát cỡ chữ, font", "Chữ dễ đọc, không quá nhỏ"),
        ("Quan sát căn lề các thành phần", "Các thành phần thẳng hàng"),
    ])
    + _tc({"tc_id": "TC_DHTD_002_01", "scenario_id": "TS_DHTD_002", "module": "Đơn hàng tuyển dụng",
           "tieu_de": "Hiển thị đúng ở các kích thước màn hình", "loai_test": "Responsive", "priority": "High",
           "precondition": "Đã đăng nhập", "kich_thuoc": "1920x1080, 1366x768, 768x1024, 375x812",
           "cach_test": "Bán tự động"}, [
        LOGIN,
        ("Kiểm tra không tràn ngang", "Không có thanh cuộn ngang", "kiểm_tra_không_tràn"),
        ("Kiểm tra nút Tạo đơn hàng vẫn hiển thị", "Nút hiển thị", "kiểm_tra_hiển_thị", 'Nút "+ Tạo đơn hàng"'),
        ("Chụp màn hình để duyệt bố cục", "Bố cục gọn, không vỡ", "chụp_và_duyệt"),
    ])
    + _tc({"tc_id": "TC_DHTD_003_01", "scenario_id": "TS_DHTD_003", "module": "Đơn hàng tuyển dụng",
           "tieu_de": "Tạo đơn hàng hợp lệ", "loai_test": "Functional", "priority": "High",
           "precondition": "Đã đăng nhập", "test_data": "vi_tri=Tester tự động ${ngau_nhien}",
           "cach_test": "Tự động"}, [
        LOGIN,
        ("Bấm Tạo đơn hàng", "Hiện form Tạo đơn hàng", "bấm", 'Nút "+ Tạo đơn hàng"'),
        ("Nhập vị trí tuyển", "", "nhập", "Ô Vị trí tuyển", "${vi_tri}"),
        ("Nhập số lượng 5", "", "nhập", "Ô Số lượng", "5"),
        ("Chọn phòng ban", "", "chọn", "Dropdown Phòng ban", "Phòng IT"),
        ("Bấm Lưu & Gửi phê duyệt", "", "bấm", 'Nút "Lưu & Gửi phê duyệt"'),
        ("Kiểm tra thông báo", "Hiện 'Tạo đơn hàng thành công'", "kiểm_tra_text", "", "Tạo đơn hàng thành công"),
        ("Đơn mới nằm đầu bảng", "Dòng đầu là vị trí vừa tạo", "kiểm_tra_text", "Bảng đơn hàng", "${vi_tri}"),
    ])
    + _tc({"tc_id": "TC_DHTD_003_02", "scenario_id": "TS_DHTD_003", "module": "Đơn hàng tuyển dụng",
           "tieu_de": "Tìm kiếm theo mã đơn", "loai_test": "Functional", "priority": "Medium",
           "cach_test": "Tự động"}, [
        LOGIN,
        ("Nhập mã DH004 vào ô tìm kiếm", "", "nhập", "Ô Tìm theo mã hoặc vị trí", "DH004"),
        ("Bấm Tìm kiếm", "", "bấm", 'Nút "Tìm kiếm"'),
        ("Kiểm tra còn 1 kết quả", "Bảng còn 1 dòng", "kiểm_tra_số_lượng", "Dòng trong bảng", "1"),
        ("Kiểm tra đúng đơn", "Dòng là 'Tester (QC)'", "kiểm_tra_text", "Bảng đơn hàng", "Tester (QC)"),
    ])
    + _tc({"tc_id": "TC_DHTD_003_03", "scenario_id": "TS_DHTD_003", "module": "Đơn hàng tuyển dụng",
           "tieu_de": "Xóa đơn hàng có hộp thoại xác nhận", "loai_test": "Functional", "priority": "Medium",
           "cach_test": "Tự động"}, [
        LOGIN,
        ("Bấm Xóa ở dòng đầu tiên", "Hiện hộp thoại xác nhận", "bấm", 'Nút "Xóa" #1'),
        ("Kiểm tra nội dung hộp thoại", "Hỏi 'Bạn có chắc muốn xóa'", "kiểm_tra_hộp_thoại", "", "Bạn có chắc"),
        ("Kiểm tra thông báo đã xóa", "Hiện 'Đã xóa đơn hàng DH001'", "kiểm_tra_text", "", "Đã xóa đơn hàng DH001"),
    ])
    + _tc({"tc_id": "TC_DHTD_004_01", "scenario_id": "TS_DHTD_004", "module": "Đơn hàng tuyển dụng",
           "tieu_de": "Bỏ trống Vị trí tuyển", "loai_test": "Validation", "priority": "High",
           "cach_test": "Tự động"}, [
        LOGIN,
        ("Bấm Tạo đơn hàng", "", "bấm", 'Nút "+ Tạo đơn hàng"'),
        ("Nhập số lượng 2, bỏ trống vị trí", "", "nhập", "Ô Số lượng", "2"),
        ("Bấm Lưu & Gửi phê duyệt", "", "bấm", 'Nút "Lưu & Gửi phê duyệt"'),
        ("Kiểm tra lỗi", "Hiện 'Vui lòng nhập vị trí tuyển'", "kiểm_tra_text", "", "Vui lòng nhập vị trí tuyển"),
    ])
    + _tc({"tc_id": "TC_DHTD_004_02", "scenario_id": "TS_DHTD_004", "module": "Đơn hàng tuyển dụng",
           "tieu_de": "Số lượng = 101 (vượt giá trị biên trên)", "loai_test": "Boundary", "priority": "High",
           "test_data": "so_luong=101", "cach_test": "Tự động",
           "ghi_chu": "Web demo CỐ Ý có lỗi này để xem tool báo Fail"}, [
        LOGIN,
        ("Bấm Tạo đơn hàng", "", "bấm", 'Nút "+ Tạo đơn hàng"'),
        ("Nhập vị trí", "", "nhập", "Ô Vị trí tuyển", "Kiểm thử biên"),
        ("Nhập số lượng 101", "", "nhập", "Ô Số lượng", "${so_luong}"),
        ("Bấm Lưu & Gửi phê duyệt", "", "bấm", 'Nút "Lưu & Gửi phê duyệt"'),
        ("Kiểm tra báo lỗi", "Hiện 'Số lượng từ 1 đến 100', không tạo đơn", "kiểm_tra_text", "",
         "Số lượng từ 1 đến 100"),
    ])
    + _tc({"tc_id": "TC_DHTD_004_03", "scenario_id": "TS_DHTD_004", "module": "Đơn hàng tuyển dụng",
           "tieu_de": "Số lượng = 0 (dưới giá trị biên)", "loai_test": "Boundary", "priority": "Medium",
           "cach_test": "Tự động"}, [
        LOGIN,
        ("Bấm Tạo đơn hàng", "", "bấm", 'Nút "+ Tạo đơn hàng"'),
        ("Nhập vị trí", "", "nhập", "Ô Vị trí tuyển", "Kiểm thử biên"),
        ("Nhập số lượng 0", "", "nhập", "Ô Số lượng", "0"),
        ("Bấm Lưu & Gửi phê duyệt", "", "bấm", 'Nút "Lưu & Gửi phê duyệt"'),
        ("Kiểm tra báo lỗi", "Hiện 'Số lượng từ 1 đến 100'", "kiểm_tra_text", "", "Số lượng từ 1 đến 100"),
    ])
    + _tc({"tc_id": "TC_SEC_01", "scenario_id": "TS_LOGIN_001", "module": "Đăng nhập",
           "tieu_de": "Vào trang Đơn hàng khi chưa đăng nhập", "loai_test": "Security", "priority": "High",
           "cach_test": "Tự động"}, [
        ("Mở thẳng trang đơn hàng", "", "mở_trang", "", "${base_url}/don-hang.html"),
        ("Kiểm tra bị chuyển về đăng nhập", "Chuyển về login.html", "kiểm_tra_url", "", "login.html"),
    ])
    + _tc({"tc_id": "TC_PERF_01", "scenario_id": "TS_DHTD_001", "module": "Đơn hàng tuyển dụng",
           "tieu_de": "Trang đăng nhập tải dưới 3 giây, không lỗi JS", "loai_test": "Performance",
           "priority": "Low", "cach_test": "Tự động"}, [
        ("Mở trang đăng nhập", "", "mở_trang", "", "${base_url}/login.html"),
        ("Đo thời gian tải", "Dưới 3 giây", "đo_thời_gian_tải", "", "3"),
        ("Kiểm tra không lỗi JavaScript", "Không có lỗi", "kiểm_tra_không_lỗi_js"),
    ])
)

ENV_NAMES = ["Demo", "Staging"]
ENV_ROWS = [
    ["base_url", "${demo_url}", "http://staging.congty.local"],
    ["user", "admin@demo.vn", ""],
    ["pass", "123456", ""],
]
ELEMENTS = [
    ["Dòng trong bảng", "css=#bang-don-hang tbody tr", "Mỗi dòng đơn hàng"],
    ["Bảng đơn hàng", "#bang-don-hang", ""],
]


def sample_data():
    return {"scenarios": SCENARIOS, "testcases": TESTCASES, "env_names": ENV_NAMES, "env_rows": ENV_ROWS,
            "elements": ELEMENTS}
