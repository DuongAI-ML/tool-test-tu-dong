"""Danh sách HÀNH ĐỘNG (từ khóa) cô tester ghi vào cột "Hành động" của mỗi bước.

Tên có thể viết có dấu/không dấu, dùng gạch dưới hoặc dấu cách: "kiểm_tra_text" = "kiem tra text" = "Kiểm tra text".
"""
from columns import norm

# name, aliases, target ("" | "cần" | "tùy chọn"), value ("" | "cần" | "tùy chọn"), mô tả, ví dụ đối tượng, ví dụ giá trị,
# manual (True = bước này người test tự đánh giá)
_K = [
    ("mở_trang", ["open", "goto", "mo url", "truy cap", "vao trang", "mo trang web", "navigate"], "", "cần",
     "Mở một địa chỉ web", "", "${base_url}/login.html", False),
    ("bấm", ["click", "nhan", "click chuot", "bam vao", "nhan vao", "chon nut"], "cần", "",
     "Bấm vào nút / link / phần tử", 'Nút "Đăng nhập"', "", False),
    ("bấm_đúp", ["double click", "dblclick", "nhap dup", "bam dup"], "cần", "", "Bấm đúp chuột", "Dòng đầu tiên",
     "", False),
    ("bấm_phải", ["right click", "click phai", "chuot phai"], "cần", "", "Bấm chuột phải", "Ảnh đại diện", "", False),
    ("di_chuột", ["hover", "re chuot", "dua chuot"], "cần", "", "Đưa chuột lên phần tử (hiện menu, tooltip)",
     'Menu "Tài khoản"', "", False),
    ("nhập", ["fill", "type", "nhap lieu", "dien", "go", "nhap gia tri", "input"], "cần", "cần",
     "Xóa nội dung cũ rồi nhập giá trị vào ô", "Ô Email", "${user}", False),
    ("gõ_từng_phím", ["type slowly", "go tung phim", "go cham"], "cần", "cần",
     "Gõ từng phím một (cho ô có gợi ý / tự hoàn thành)", "Ô Tìm kiếm", "DH001", False),
    ("xóa_nội_dung", ["clear", "xoa trang", "xoa o"], "cần", "", "Xóa trắng ô nhập", "Ô Email", "", False),
    ("nhấn_phím", ["press", "phim", "an phim", "bam phim"], "tùy chọn", "cần",
     "Nhấn phím (Enter, Tab, Escape, Control+A…)", "Ô Tìm kiếm", "Enter", False),
    ("chọn", ["select", "chon dropdown", "chon gia tri", "chon trong danh sach", "chon option"], "cần", "cần",
     "Chọn giá trị trong dropdown (theo chữ hiển thị)", "Dropdown Phòng ban", "Phòng IT", False),
    ("tick", ["check", "danh dau", "chon checkbox", "tich"], "cần", "", "Tick ô checkbox / radio",
     "Checkbox Ghi nhớ", "", False),
    ("bỏ_tick", ["uncheck", "bo chon", "bo danh dau", "bo tich"], "cần", "", "Bỏ tick checkbox", "Checkbox Ghi nhớ",
     "", False),
    ("tải_file_lên", ["upload", "tai len", "chon file", "dinh kem"], "cần", "cần",
     "Chọn file để tải lên (nhiều file cách nhau bằng ;)", "Ô Đính kèm", r"C:\TestData\cv.pdf", False),
    ("chờ", ["wait", "doi", "sleep", "tam dung"], "", "cần", "Chờ số giây", "", "2", False),
    ("chờ_thấy", ["wait for", "doi hien thi", "cho hien thi", "doi thay"], "cần", "tùy chọn",
     "Chờ đến khi phần tử hiện ra (giá trị = số giây tối đa)", "Bảng đơn hàng", "15", False),
    ("chờ_ẩn", ["wait hidden", "cho bien mat", "doi an"], "cần", "tùy chọn", "Chờ đến khi phần tử biến mất",
     "Biểu tượng tải", "15", False),
    ("cuộn_tới", ["scroll", "scroll to", "cuon den", "keo toi"], "cần", "", "Cuộn trang tới phần tử",
     "Nút \"Lưu\"", "", False),
    ("quay_lại", ["back", "go back", "tro lai"], "", "", "Quay lại trang trước", "", "", False),
    ("tải_lại_trang", ["reload", "refresh", "f5", "lam moi"], "", "", "Tải lại trang (F5)", "", "", False),
    ("chuyển_tab", ["switch tab", "doi tab", "sang tab"], "", "tùy chọn",
     "Chuyển sang tab trình duyệt khác (số thứ tự, hoặc để trống = tab mới nhất)", "", "2", False),
    ("kiểm_tra_hiển_thị", ["assert visible", "kiem tra hien", "kiem tra co", "kiem tra thay", "verify visible",
                           "kiem tra xuat hien"], "cần", "", "Kiểm tra phần tử ĐANG HIỂN THỊ",
     'Nút "Tạo đơn hàng"', "", False),
    ("kiểm_tra_ẩn", ["assert hidden", "kiem tra khong hien thi", "kiem tra khong co", "kiem tra bien mat",
                     "verify hidden"], "cần", "", "Kiểm tra phần tử KHÔNG hiển thị", "Hộp thoại lỗi", "", False),
    ("kiểm_tra_text", ["assert text", "kiem tra noi dung", "kiem tra chu", "kiem tra co chu", "verify text",
                       "kiem tra thong bao"], "tùy chọn", "cần",
     "Kiểm tra có dòng chữ (trong phần tử, hoặc cả trang nếu bỏ trống đối tượng)", "", "Xin chào", False),
    ("kiểm_tra_không_có_text", ["assert no text", "kiem tra khong co chu", "kiem tra khong co text"], "tùy chọn",
     "cần", "Kiểm tra KHÔNG có dòng chữ", "", "Lỗi", False),
    ("kiểm_tra_giá_trị", ["assert value", "kiem tra gia tri o"], "cần", "cần",
     "Kiểm tra giá trị đang có trong ô nhập", "Ô Email", "admin@demo.vn", False),
    ("kiểm_tra_url", ["assert url", "kiem tra duong dan", "kiem tra dia chi"], "", "cần",
     "Kiểm tra địa chỉ trang có chứa đoạn chữ", "", "/don-hang", False),
    ("kiểm_tra_tiêu_đề", ["assert title", "kiem tra title", "kiem tra tieu de trang"], "", "cần",
     "Kiểm tra tiêu đề tab trình duyệt có chứa chữ", "", "Đơn hàng", False),
    ("kiểm_tra_số_lượng", ["assert count", "dem", "kiem tra so dong", "dem so luong"], "cần", "cần",
     "Đếm số phần tử (=5, >=1, <10…)", "Dòng trong bảng", ">=1", False),
    ("kiểm_tra_bật", ["assert enabled", "kiem tra bam duoc", "kiem tra enable"], "cần", "",
     "Kiểm tra phần tử đang bật (bấm/nhập được)", 'Nút "Lưu"', "", False),
    ("kiểm_tra_tắt", ["assert disabled", "kiem tra vo hieu", "kiem tra disable", "kiem tra mo"], "cần", "",
     "Kiểm tra phần tử đang bị khóa (disabled)", 'Nút "Lưu"', "", False),
    ("kiểm_tra_đã_tick", ["assert checked", "kiem tra da chon", "kiem tra da tich"], "cần", "",
     "Kiểm tra checkbox/radio đang được tick", "Checkbox Ghi nhớ", "", False),
    ("kiểm_tra_hộp_thoại", ["assert dialog", "kiem tra popup", "kiem tra alert", "kiem tra confirm"], "", "cần",
     "Kiểm tra hộp thoại alert/confirm vừa hiện có chứa chữ (tool tự bấm OK)", "", "Bạn có chắc", False),
    ("hủy_hộp_thoại_tiếp_theo", ["dismiss dialog", "bam huy hop thoai", "cancel dialog"], "", "",
     "Hộp thoại confirm tiếp theo sẽ bấm Hủy (mặc định tool bấm OK)", "", "", False),
    ("đổi_kích_thước", ["resize", "viewport", "doi man hinh", "kich thuoc"], "", "cần",
     "Đổi kích thước màn hình: 1366x768 hoặc mobile / tablet / laptop / desktop", "", "768x1024", False),
    ("kiểm_tra_không_tràn", ["assert no overflow", "kiem tra khong tran ngang", "kiem tra tran", "khong tran"],
     "", "", "Kiểm tra trang không bị tràn ngang (không có thanh cuộn ngang)", "", "", False),
    ("chụp_màn_hình", ["screenshot", "chup anh", "chup man hinh"], "tùy chọn", "tùy chọn",
     "Chụp ảnh (giá trị 'toàn trang' = chụp cả trang dài)", "", "toàn trang", False),
    ("so_sánh_ảnh", ["visual compare", "so sanh giao dien", "so anh", "so sanh anh chuan"], "tùy chọn", "tùy chọn",
     "So với ảnh chuẩn (lần đầu chạy tự lưu làm ảnh chuẩn). Giá trị = % khác cho phép, mặc định 1", "", "1",
     False),
    ("đo_thời_gian_tải", ["assert load time", "do toc do", "thoi gian tai", "performance"], "", "tùy chọn",
     "Kiểm tra trang tải xong trong N giây (mặc định 3)", "", "3", False),
    ("kiểm_tra_không_lỗi_js", ["assert no console error", "kiem tra loi console", "khong loi console",
                                "kiem tra khong loi"], "", "", "Kiểm tra trang không có lỗi JavaScript", "", "",
     False),
    ("kiểm_tra_link_hỏng", ["check broken links", "kiem tra link", "link hong", "kiem tra lien ket"], "", "",
     "Kiểm tra các link trên trang không bị lỗi 404/500", "", "", False),
    ("lưu_text", ["store text", "luu bien", "lay text", "luu gia tri"], "cần", "cần",
     "Lấy chữ của phần tử lưu vào biến, dùng lại bằng ${tên_biến}", "Mã đơn dòng đầu", "ma_don", False),
    ("chạy_tc", ["call", "goi tc", "chay test case", "goi flow", "chay flow", "dung lai tc"], "", "cần",
     "Chạy lại các bước của TC khác (VD: đăng nhập)", "", "TC_LOGIN_01", False),
    ("chạy_js", ["eval", "javascript", "js", "run js"], "tùy chọn", "cần", "Chạy đoạn JavaScript (cho người biết code)",
     "", "localStorage.clear()", False),
    ("kiểm_tra_tay", ["manual", "kiem tra bang mat", "test tay", "nguoi kiem tra", "tu kiem tra"], "", "",
     "BƯỚC TEST TAY: tool dừng lại, bạn tự kiểm tra rồi chọn Pass/Fail", "", "", True),
    ("chụp_và_duyệt", ["screenshot review", "chup va duyet", "chup de duyet", "duyet anh"], "tùy chọn", "tùy chọn",
     "Tool chụp ảnh rồi DỪNG để bạn xem ảnh và chọn Pass/Fail", "", "", True),
]

KEYWORDS = []
LOOKUP = {}
for name, aliases, target, value, desc, ex_t, ex_v, manual in _K:
    k = {"name": name, "aliases": aliases, "target": target, "value": value, "desc": desc,
         "ex_target": ex_t, "ex_value": ex_v, "manual": manual}
    KEYWORDS.append(k)
    for a in [name] + aliases:
        LOOKUP.setdefault(norm(a), name)


def lookup(text):
    """Tên hành động người dùng ghi -> tên chuẩn, hoặc None nếu không nhận ra. Chuỗi rỗng -> ''."""
    n = norm(text)
    if not n:
        return ""
    return LOOKUP.get(n)


def info(name):
    for k in KEYWORDS:
        if k["name"] == name:
            return k
    return None


MANUAL = {k["name"] for k in KEYWORDS if k["manual"]}
