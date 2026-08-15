#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
stamp_pdf_mistral.py — Dùng Mistral AI (model có khả năng Vision) để đọc &
phân loại từng trang của 1 file PDF scan hồ sơ nhân sự (nhiều loại văn bản
gộp chung), rồi TYPEWRITER (chèn chữ) mã viết tắt PL02 màu đỏ vào
GÓC TRÊN BÊN TRÁI của trang đầu tiên mỗi văn bản — giống mẫu
file_scan_note.pdf (vd: "HDTHUVIEC", "HDXDTH-01", "HDXDTH-02"...).

KHÔNG tách file PDF — giữ nguyên toàn bộ số trang, chỉ thêm nhãn để một
công cụ khác dựa vào đó tách file sau.

CÁCH DÙNG:
    python stamp_pdf_mistral.py "duong_dan_file.pdf" --outfile "da_danh_dau.pdf"

YÊU CẦU:
    pip install -U pymupdf mistralai pydantic
    Cần biến môi trường MISTRAL_API_KEY (hoặc dùng --api-key)
    Lấy API key tại: https://console.mistral.ai/ (mục "API Keys")
    LƯU Ý: gói Free của Mistral giới hạn khá thấp (khoảng 1 request/giây),
    nên mặc định --delay được đặt là 3 giây/trang; nếu đã nâng cấp gói trả
    phí (pay-as-you-go / Scale) có thể giảm xuống.

Ý TƯỞNG THUẬT TOÁN:
    1. Render từng trang PDF thành ảnh PNG (PyMuPDF) để gửi cho Mistral.
    2. Gửi trang đang xét cho Mistral qua Chat Completions API (client.chat.parse),
       KÈM THEO cả ảnh của TRANG NGAY TRƯỚC ĐÓ (để model so sánh trực quan 2
       trang liên tiếp, tránh nhầm trang tiếp nối — cùng logo/header lặp lại —
       thành văn bản mới), cùng DANH MỤC LOẠI GIẤY TỜ (PL02) và quy tắc phân
       biệt trang tiếp nối/văn bản mới, yêu cầu trả JSON theo schema
       (is_new_document_start, ma_loai_giay_to, title).
    3. Gom các trang liên tiếp cùng 1 văn bản (is_new_document_start=True
       đánh dấu ranh giới văn bản mới).
    4. Đánh số thứ tự (STT) cho các loại giấy tờ xuất hiện >= 2 lần trong
       file (đúng quy định PL02: chỉ thêm STT khi có từ 2 văn bản cùng loại
       trở lên).
    5. Chèn text (typewriter) màu đỏ vào góc trên-trái của TRANG
       ĐẦU mỗi văn bản, ngay trên chính file PDF gốc (không mở file mới,
       không tách trang) — dùng page.insert_textbox() của PyMuPDF.
    6. Lưu ra 1 file PDF duy nhất (đủ số trang như file gốc, chỉ có thêm
       nhãn) + file log CSV để đối chiếu lại (LLM có thể đọc/phân loại sai).
"""

import argparse
import base64
import csv
import json
import os
import sys
import time

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("Thiếu thư viện PyMuPDF. Cài bằng: pip install pymupdf")

try:
    from mistralai.client import Mistral
except ImportError:
    sys.exit("Thiếu thư viện mistralai. Cài bằng: pip install -U mistralai")

try:
    from pydantic import BaseModel
except ImportError:
    sys.exit("Thiếu thư viện pydantic. Cài bằng: pip install -U pydantic")


# Model Vision khuyến nghị của Mistral (2026): mistral-small-latest (nhanh, rẻ,
# đủ dùng cho đọc chữ/scan); có thể đổi sang "mistral-medium-latest",
# "mistral-large-latest" hoặc "pixtral-large-latest" nếu cần đọc chữ khó hơn.
DEFAULT_MODEL = "mistral-small-latest"
RENDER_DPI = 200  # tăng lên 300 nếu chữ scan quá nhỏ/mờ khi model đọc

# ---------------------------------------------------------------------------
# Vị trí & kiểu chữ con dấu "typewriter" — chỉnh qua tham số dòng lệnh nếu cần
# Toạ độ tính bằng point (1/72 inch), gốc (0,0) ở GÓC TRÊN-TRÁI trang, giống
# mẫu file_scan_note.pdf: mã viết tắt nằm sát mép trên-trái, cỡ chữ nhỏ.
# ---------------------------------------------------------------------------
STAMP_MARGIN_X = 36      # cách mép trái ~0.5 inch
STAMP_MARGIN_Y = 12      # cách mép trên
STAMP_BOX_WIDTH = 240
STAMP_BOX_HEIGHT = 28     # ĐÃ SỬA: 16 -> 28. Với fontsize=15, PyMuPDF cần
                          # tối thiểu ~26pt chiều cao để insert_textbox() vẽ
                          # được 1 dòng chữ; để 16 khiến mode="text" ÂM THẦM
                          # KHÔNG chèn được chữ nào (insert_textbox trả về số
                          # âm = thiếu chỗ, nhưng không raise lỗi).
                          # mode="annot" (mặc định) không bị ảnh hưởng bởi
                          # tham số này vì FreeText annotation tự bao chữ.
STAMP_FONTSIZE = 15
STAMP_FONT = "hebo"      # Helvetica-Bold (font built-in của PyMuPDF, đủ cho mã ASCII)
# STAMP_COLOR = (0.09, 0.30, 0.72)  # xanh dương giống mẫu
STAMP_COLOR = (1.0, 0.0, 0.0) # ĐỎ CHUẨN
# ---------------------------------------------------------------------------
# DANH MỤC VIẾT TẮT LOẠI GIẤY TỜ (PL02) + QUY TẮC PHÂN BIỆT
# ---------------------------------------------------------------------------
# QUAN TRỌNG: nội dung này KHÔNG còn hard-code trong script nữa — nó được
# đọc từ file JSON bên ngoài "danh_muc_pl02.json" (nằm CÙNG THƯ MỤC với
# script/exe, hoặc chỉ định qua --config). Nhờ vậy, khi công ty cập nhật
# PL02 hoặc cần chỉnh quy tắc nhận diện cho Mistral, CHỈ CẦN SỬA FILE JSON
# NÀY BẰNG NOTEPAD — KHÔNG CẦN BUILD LẠI FILE .EXE.
#
# Nếu chưa có file config, script sẽ tự tạo file mẫu (từ danh mục mặc định
# bên dưới) ngay lần chạy đầu tiên để bạn có sẵn mà chỉnh sửa.
# ---------------------------------------------------------------------------

CONFIG_FILENAME = "danh_muc_pl02.json"

# Danh mục MẶC ĐỊNH — chỉ dùng để (a) sinh file config mẫu lần đầu, và
# (b) làm fallback nếu file config bị thiếu/lỗi. Muốn thay đổi lâu dài,
# hãy sửa file danh_muc_pl02.json, đừng sửa ở đây.
DANH_MUC_MAC_DINH = [
    ("QDTUYENDUNG", "Quyết định tuyển dụng lao động thử việc VÀ các quyết định tiếp theo trong "
     "cùng chuỗi tuyển dụng — bao gồm cả 'Quyết định V/v chấm dứt thời gian thử việc, ký hợp đồng "
     "lao động xác định/không xác định thời hạn' (quyết định kết thúc thử việc để chuyển sang ký "
     "HĐLĐ chính thức). Nhận diện qua tiêu đề có cụm 'Tuyển dụng lao động thử việc' hoặc 'chấm dứt "
     "thời gian thử việc, ký hợp đồng lao động...'. KHÔNG dùng QDCHAMDUT cho loại này."),
    ("THONGBAOLUONG", "Thông báo lương / thông báo kết quả xếp lương, tăng lương"),
    ("SOYEULYLICH", "Sơ yếu lý lịch do cán bộ tự khai và có xác nhận của cơ quan nhà nước có thẩm quyền "
     "(thường có dán ảnh 4x6, tiêu đề 'SƠ YẾU LÝ LỊCH')"),
    ("GIAYKHAISINH", "Bản sao Giấy khai sinh có chứng thực của cơ quan có thẩm quyền (mẫu của UBND "
     "xã/phường, tiêu đề 'GIẤY KHAI SINH')"),
    ("CCCD", "Bản sao Căn cước công dân/Căn cước có chứng thực của cơ quan có thẩm quyền (ảnh thẻ CCCD "
     "có mã QR, quốc huy, dòng chữ 'CĂN CƯỚC CÔNG DÂN' hoặc 'CĂN CƯỚC')"),
    ("CMND", "Bản sao Chứng minh nhân dân (mẫu CMND cũ, 9 hoặc 12 số) có chứng thực của cơ quan có thẩm quyền"),
    ("HOCHIEU", "Bản sao hộ chiếu có chứng thực của cơ quan có thẩm quyền"),
    ("DINHDANH", "Bản sao mã định danh điện tử/thông báo định danh cá nhân có chứng thực của cơ quan có thẩm quyền"),
    ("BANGDAIHOC", "Bằng đại học / bằng cử nhân (văn bằng tốt nghiệp, có quốc huy, tên trường, ngành học, xếp loại)"),
    ("BANGDIEMDH", "Bảng điểm đại học (bảng liệt kê 'STT | Mã HP/Unit code | Tên học phần/Unit title' kèm điểm số)"),
    ("BANGTIENSI", "Bằng tiến sĩ"),
    ("BANGTHACSI", "Bằng thạc sĩ"),
    ("CHUNGCHIKHAC", "Chứng chỉ khác, bằng cấp khác không thuộc các mã trên (vd: chứng chỉ ngoại ngữ/tin "
     "học, TOEIC, IELTS, chứng chỉ nghiệp vụ ngân hàng, chứng chỉ đào tạo ngắn hạn...)"),
    ("GIAYKHAMSK", "Giấy khám sức khỏe (mẫu của cơ sở y tế/bệnh viện, có các mục khám thể lực, khám lâm "
     "sàng, tiêu đề 'GIẤY KHÁM SỨC KHỎE')"),
    ("CV", "Hồ sơ/thông tin ứng viên nộp khi ứng tuyển (mẫu nội bộ 'THÔNG TIN ỨNG VIÊN', kỹ năng, quá "
     "trình công tác — khác với Sơ yếu lý lịch có xác nhận nhà nước)"),
    ("QDCHAMDUT", "Quyết định chấm dứt HỢP ĐỒNG LAO ĐỘNG / chấm dứt quan hệ lao động, cho thôi việc, nghỉ "
     "việc HẲN (kết thúc quan hệ lao động với công ty). CHỈ dùng mã này khi văn bản thật sự chấm dứt "
     "việc làm; KHÔNG dùng cho quyết định kết thúc thời gian thử việc để ký HĐLĐ tiếp theo (loại đó "
     "thuộc QDTUYENDUNG, xem mô tả ở trên)."),
    ("CAMKETBAOMAT", "Cam kết/Hợp đồng bảo mật thông tin (tiêu đề có thể là 'CAM KẾT BẢO MẬT THÔNG TIN' "
     "hoặc 'HỢP ĐỒNG BẢO MẬT THÔNG TIN')"),
    ("CAMKETKHAC", "Bản cam kết khác không phải bảo mật thông tin (vd: cam kết không cạnh tranh, cam kết "
     "thử thách công việc...)"),
    ("HDTHUVIEC", "Hợp đồng thử việc (tiêu đề có cụm 'HỢP ĐỒNG THỬ VIỆC' hoặc 'HĐTV')"),
    ("HDXDTH", "Hợp đồng lao động XÁC ĐỊNH THỜI HẠN (tiêu đề có cụm 'HỢP ĐỒNG LAO ĐỘNG' kèm 'XÁC ĐỊNH "
     "THỜI HẠN' hoặc mã hiệu 'HDLD.XDTH')"),
    ("HDKXDTH", "Hợp đồng lao động KHÔNG XÁC ĐỊNH THỜI HẠN (tiêu đề có cụm 'KHÔNG XÁC ĐỊNH THỜI HẠN' hoặc "
     "mã hiệu 'HDLD.KXDTH')"),
    ("PHULUCHD", "Phụ lục hợp đồng lao động (văn bản sửa đổi/bổ sung một số điều khoản của HĐLĐ đã ký, "
     "tiêu đề 'PHỤ LỤC HỢP ĐỒNG LAO ĐỘNG')"),
    ("HDDAOTAO", "Hợp đồng đào tạo / cam kết đào tạo (ràng buộc nghĩa vụ làm việc sau khi được cử đi đào tạo)"),
    ("QDDIEUDONG", "Quyết định điều động / Quyết định chuyển đổi vị trí, đơn vị công tác (không phải bổ "
     "nhiệm chức danh quản lý)"),
    ("QDBONHIEM", "Quyết định bổ nhiệm chức danh quản lý / điều động kèm bổ nhiệm"),
    ("QDTAMHOAN", "Quyết định tạm hoãn thực hiện hợp đồng lao động"),
    ("QDDIEUCHINH", "Quyết định điều chỉnh cấp bậc, ngạch/bậc lương, chức danh công việc (không kèm đổi "
     "vị trí/đơn vị)"),
    ("QDNGHIHUU", "Quyết định nghỉ hưu"),
    ("QDKYLUAT", "Quyết định/biên bản xử lý kỷ luật lao động"),
    ("KEKHAITS", "Bản kê khai tài sản, thu nhập"),
]

QUY_TAC_PHAN_BIET_MAC_DINH = """1. "Quyết định V/v chấm dứt thời gian thử việc, ký hợp đồng lao động xác
   định/không xác định thời hạn" là quyết định KẾT THÚC THỬ VIỆC để chuyển
   sang ký HĐLĐ chính thức — vẫn thuộc nhóm QDTUYENDUNG (tuyển dụng), TUYỆT
   ĐỐI KHÔNG gán QDCHAMDUT dù tiêu đề có chữ "chấm dứt". Chỉ gán QDCHAMDUT
   khi văn bản thật sự chấm dứt HỢP ĐỒNG LAO ĐỘNG / quan hệ lao động (cho
   thôi việc hẳn, nghỉ việc).
2. Phân biệt 3 loại hợp đồng lao động bằng đúng cụm từ trong tiêu đề:
   "THỬ VIỆC" -> HDTHUVIEC; "XÁC ĐỊNH THỜI HẠN" -> HDXDTH;
   "KHÔNG XÁC ĐỊNH THỜI HẠN" -> HDKXDTH. Văn bản chỉ sửa đổi/bổ sung một
   vài điều khoản của hợp đồng đã ký (không phải hợp đồng mới) -> PHULUCHD.
3. CCCD (thẻ căn cước có mã QR, quốc huy) khác CMND (chứng minh nhân dân
   mẫu cũ) khác GIAYKHAISINH (giấy khai sinh do UBND cấp) khác SOYEULYLICH
   (sơ yếu lý lịch có dán ảnh, do cán bộ tự khai) — đọc kỹ tiêu đề/loại
   giấy tờ trên ảnh, không suy đoán theo vị trí xuất hiện.
4. CAMKETBAOMAT dùng cho cả văn bản tiêu đề "CAM KẾT BẢO MẬT THÔNG TIN" lẫn
   "HỢP ĐỒNG BẢO MẬT THÔNG TIN" (cùng nội dung bảo mật). CAMKETKHAC chỉ
   dùng cho cam kết KHÔNG liên quan bảo mật thông tin.
5. QDDIEUDONG (điều động/chuyển đổi vị trí) khác QDBONHIEM (bổ nhiệm chức
   danh quản lý) khác QDDIEUCHINH (chỉ điều chỉnh cấp bậc/ngạch lương,
   không đổi vị trí công tác)."""


def _duong_dan_thu_muc_goc():
    """Thư mục chứa file config: cùng thư mục với file .exe khi đã build
    bằng PyInstaller (sys.frozen), hoặc cùng thư mục với script .py khi
    chạy trực tiếp bằng python."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _tao_file_config_mau(duong_dan):
    du_lieu = {
        "danh_muc_loai_giay_to": [
            {"ma": ma, "mo_ta": mo_ta} for ma, mo_ta in DANH_MUC_MAC_DINH
        ],
        "quy_tac_phan_biet": QUY_TAC_PHAN_BIET_MAC_DINH,
    }
    with open(duong_dan, "w", encoding="utf-8") as f:
        json.dump(du_lieu, f, ensure_ascii=False, indent=2)


def load_danh_muc_config(config_path=None):
    """Đọc danh mục PL02 + quy tắc phân biệt từ file JSON bên ngoài.
    Nếu chưa chỉ định --config, sẽ tìm file 'danh_muc_pl02.json' cùng thư
    mục với script/exe; nếu chưa tồn tại, tự tạo file mẫu để người dùng
    chỉnh sửa (không cần build lại exe).
    Trả về: (danh_sach_tuple[(ma, mo_ta), ...], quy_tac_phan_biet_str)
    """
    duong_dan = config_path or os.path.join(_duong_dan_thu_muc_goc(), CONFIG_FILENAME)

    if not os.path.isfile(duong_dan):
        try:
            _tao_file_config_mau(duong_dan)
            print(f"[i] Chưa có file cấu hình, đã tạo file mẫu: {duong_dan}")
            print("    Bạn có thể mở file này bằng Notepad để sửa danh mục/quy tắc")
            print("    MÀ KHÔNG CẦN build lại chương trình.")
        except Exception as e:  # noqa: BLE001
            print(f"[!] Không tạo được file cấu hình mẫu ({e}). Dùng danh mục mặc định trong code.")
            return DANH_MUC_MAC_DINH, QUY_TAC_PHAN_BIET_MAC_DINH

    try:
        with open(duong_dan, "r", encoding="utf-8") as f:
            du_lieu = json.load(f)
        danh_sach = [(item["ma"].strip(), item["mo_ta"].strip()) for item in du_lieu["danh_muc_loai_giay_to"]]
        quy_tac = du_lieu.get("quy_tac_phan_biet", QUY_TAC_PHAN_BIET_MAC_DINH)
        if not danh_sach:
            raise ValueError("danh_muc_loai_giay_to rỗng")
        print(f"[i] Đã nạp {len(danh_sach)} mã loại giấy tờ từ file cấu hình: {duong_dan}")
        return danh_sach, quy_tac
    except Exception as e:  # noqa: BLE001
        print(f"[!] Lỗi đọc file cấu hình '{duong_dan}': {e}")
        print("    Dùng tạm danh mục mặc định trong code. Hãy sửa lại file JSON cho đúng định dạng.")
        return DANH_MUC_MAC_DINH, QUY_TAC_PHAN_BIET_MAC_DINH


# ---------------------------------------------------------------------------
# QUAN TRỌNG (GUI/--windowed): KHÔNG được nạp config (và không được print())
# ở cấp module — vì gui_stamp_pdf.py làm "import stamp_pdf_mistral" NGAY LÚC
# MỞ APP, trước khi GUI kịp gắn StdoutRedirector; nếu build bằng PyInstaller
# --windowed thì lúc đó sys.stdout có thể là None -> gọi print() sẽ LÀM CRASH
# APP NGAY KHI VỪA MỞ. Vì vậy việc nạp config được để DẠNG LAZY, chỉ thực sự
# chạy bên trong hàm stamp_pdf() (tức là lúc đã bấm "Bắt đầu" trong GUI, khi
# sys.stdout đã được thay bằng StdoutRedirector an toàn).
# ---------------------------------------------------------------------------
DANH_MUC_LOAI_GIAY_TO = None   # sẽ được nap_cau_hinh() gán
QUY_TAC_PHAN_BIET = None       # sẽ được nap_cau_hinh() gán
MA_LOAI_HOP_LE = None          # sẽ được nap_cau_hinh() gán
_danh_muc_text_lines = None    # sẽ được nap_cau_hinh() gán


def nap_cau_hinh(config_path=None):
    """Nạp (hoặc nạp LẠI) danh mục PL02 + quy tắc phân biệt vào các biến
    module-level. Gọi hàm này ở ĐẦU stamp_pdf() (không gọi ở cấp module) để
    tránh crash khi bị import trong ứng dụng GUI --windowed, đồng thời cho
    phép mỗi lần chạy đều đọc bản JSON mới nhất (sửa JSON xong không cần
    khởi động lại app)."""
    global DANH_MUC_LOAI_GIAY_TO, QUY_TAC_PHAN_BIET, MA_LOAI_HOP_LE, _danh_muc_text_lines
    DANH_MUC_LOAI_GIAY_TO, QUY_TAC_PHAN_BIET = load_danh_muc_config(config_path)
    MA_LOAI_HOP_LE = {ma for ma, _ten in DANH_MUC_LOAI_GIAY_TO} | {"KHAC"}
    _danh_muc_text_lines = "\n".join(
        f'  - "{ma}": {ten}' for ma, ten in DANH_MUC_LOAI_GIAY_TO
    )


class PhanLoaiTrang(BaseModel):
    """Schema JSON bắt buộc mà Mistral phải trả về cho mỗi trang."""
    is_new_document_start: bool
    ma_loai_giay_to: str
    title: str


def build_prompt_text(prev_ma, has_prev_image):
    """Dựng nội dung prompt gửi Mistral, dùng danh mục/quy tắc ĐANG được nạp
    (từ nap_cau_hinh()). Xây dựng động (không phải hằng số module-level) để
    luôn phản ánh đúng bản danh_muc_pl02.json mới nhất, và để tránh phụ
    thuộc vào biến chưa được gán lúc import module.

    has_prev_image: True nếu có gửi kèm ẢNH trang trước (không chỉ mã loại
    dạng text) để model so sánh trực quan 2 trang liên tiếp — giúp phát
    hiện đúng trang tiếp nối (VD: logo/header lặp lại ở mỗi trang) thay vì
    tưởng nhầm là văn bản mới."""
    if has_prev_image:
        anh_note = (
            "Bạn được cung cấp 2 ẢNH: ẢNH THỨ NHẤT là TRANG TRƯỚC (trang liền kề, "
            "đứng ngay trước trang đang xét trong file gốc), ẢNH THỨ HAI là TRANG "
            "ĐANG XÉT (trang cần phân loại). HÃY SO SÁNH TRỰC QUAN 2 ẢNH này — nhìn "
            "bố cục, phông chữ, số hiệu văn bản, có chữ ký/con dấu ở ảnh trước hay "
            "chưa — để quyết định trang đang xét có phải TIẾP NỐI của đúng văn bản ở "
            "ảnh trước hay là một văn bản HOÀN TOÀN MỚI."
        )
    else:
        anh_note = (
            "Đây là TRANG ĐẦU TIÊN của file (không có trang trước để so sánh), nên "
            "chỉ có 1 ảnh duy nhất — ảnh của trang đang xét."
        )
    return f"""Bạn đang xem ảnh chụp/scan của (các) TRANG trong bộ hồ sơ nhân sự tiếng Việt
(hồ sơ cán bộ ngân hàng: quyết định, sơ yếu lý lịch, bằng cấp, hợp đồng...).

{anh_note}

LƯU Ý QUAN TRỌNG: các trang trong file này hầu như trang nào cũng có LOGO/HEADER
"VietinBank" lặp lại ở đầu trang — đây là header cố định của công ty, KHÔNG PHẢI
dấu hiệu của một văn bản mới. Tuyệt đối không kết luận is_new_document_start=true
chỉ vì thấy logo/header; phải xem có TIÊU ĐỀ VĂN BẢN + SỐ HIỆU VĂN BẢN MỚI hay
không (xem quy tắc chi tiết bên dưới).

Dưới đây là DANH MỤC MÃ LOẠI GIẤY TỜ chính thức (trích Phụ lục 02), kèm mô tả
chi tiết và các cụm từ/tiêu đề thường gặp trên chính văn bản thật để bạn đối
chiếu. Bạn PHẢI chọn đúng 1 mã trong danh sách này cho TRANG ĐANG XÉT, dựa
vào TIÊU ĐỀ và NỘI DUNG đọc được trên ảnh (ưu tiên đọc đúng tiêu đề/trích yếu
ở phần đầu trang hơn là đoán theo bố cục):
{_danh_muc_text_lines}
  - "KHAC": chỉ dùng khi trang không khớp với bất kỳ mã nào ở trên (vd:
    trang trắng, trang chỉ có logo/watermark ngân hàng, phụ lục biểu mẫu
    không xác định được loại).

QUY TẮC PHÂN BIỆT CÁC LOẠI DỄ NHẦM LẪN (đọc kỹ trước khi chọn mã):
{QUY_TAC_PHAN_BIET}

Trang trước đó (nếu có) được xác định là mã: {prev_ma or "(không có, đây là trang đầu tiên của file)"}

CÁCH XÁC ĐỊNH RANH GIỚI VĂN BẢN MỚI (is_new_document_start) — ĐỌC KỸ:

A. DẤU HIỆU CỦA TRANG TIẾP NỐI (is_new_document_start = FALSE) — rất phổ biến
   trong hồ sơ nhiều trang như hợp đồng, phụ lục hợp đồng:
   - Có đánh số dạng "Trang X/Y", "X/Y", hoặc số thứ tự trang ở góc/chân trang
     (vd: "1/4", "2/4"...) — đây là bằng chứng RẤT MẠNH cho thấy đây là 1 trong
     nhiều trang của CÙNG một văn bản, KHÔNG phải văn bản mới.
   - Trang bắt đầu bằng nội dung ĐIỀU KHOẢN tiếp theo, không có tiêu đề văn bản
     mới: vd bắt đầu bằng "Điều 4.", "Điều 5.", hoặc các gạch đầu dòng/mục nhỏ
     tiếp nối như "b)", "c)", "d)", "- ...", hoặc đoạn văn không viết hoa toàn bộ
     mở đầu trang (khác với tiêu đề văn bản luôn in hoa/đậm, căn giữa).
   - KHÔNG xuất hiện dòng "Số: .../QĐ-...", "Số HĐLĐ: ...", "CỘNG HÒA XÃ HỘI CHỦ
     NGHĨA VIỆT NAM" ở đầu trang (những dòng mở đầu bắt buộc của 1 văn bản mới).
   - Chỉ có chữ ký/con dấu/họ tên các bên mà KHÔNG có tiêu đề + số hiệu mới phía
     trên — đây là trang ký kết thúc của văn bản đang xét, không phải văn bản mới.
   - Khi so sánh với ẢNH TRANG TRƯỚC: nếu bố cục/phông chữ/kiểu trình bày giống
     hệt trang trước và mạch nội dung nối tiếp logic (vd trang trước kết ở "Điều
     3", trang này mở đầu "Điều 4") thì gần như chắc chắn là trang tiếp nối.

B. DẤU HIỆU CỦA VĂN BẢN MỚI THẬT SỰ (is_new_document_start = TRUE):
   - Có tiêu đề in hoa/đậm, thường căn giữa đầu trang (vd "HỢP ĐỒNG LAO ĐỘNG",
     "QUYẾT ĐỊNH", "PHỤ LỤC HỢP ĐỒNG LAO ĐỘNG"...) KÈM số hiệu văn bản MỚI (vd
     "Số: .../QĐ-...", "Số HĐLĐ: ...", "Mã hiệu: BM...") khác với văn bản ở
     trang trước.
   - Số hiệu văn bản (nếu đọc được) khác với số hiệu của văn bản đang xét ở
     trang trước — đây là bằng chứng chắc chắn nhất.
   - Trang không có đánh số "X/Y" nối tiếp từ trang trước (vd trang trước là
     "3/4" mà trang này không phải "4/4" mà lại có tiêu đề riêng).
   - Lưu ý: một văn bản có thể có PHỤ LỤC đính kèm ngay sau — phụ lục hợp đồng
     ("PHỤ LỤC HỢP ĐỒNG LAO ĐỘNG") có tiêu đề + số hiệu riêng nên VẪN được coi
     là is_new_document_start=true (là 1 văn bản riêng, mã PHULUCHD) dù nó đi
     kèm ngay sau 1 hợp đồng — trừ khi phụ lục đó chính là 1 trong các trang
     đánh số "X/Y" của văn bản trước (khi đó là tiếp nối, không phải phụ lục
     độc lập).

Khi không chắc chắn 100%, ưu tiên dựa vào: (1) số trang "X/Y" nếu có, (2) có
tiêu đề + số hiệu MỚI hay không, (3) so sánh trực quan bố cục 2 ảnh.

Hãy trả lời theo đúng schema JSON:
- is_new_document_start: true nếu trang đang xét là TRANG ĐẦU TIÊN của 1 văn
  bản KHÁC với trang trước (false nếu chỉ là trang tiếp theo / mặt sau / chữ
  ký / phần nối tiếp của CÙNG văn bản đang xét — xem mục A ở trên).
- ma_loai_giay_to: PHẢI là một trong các mã ở danh mục trên (viết đúng
  chính xác, in hoa, không dấu, không thêm ký tự khác), hoặc "KHAC".
- title: tiêu đề / trích yếu ngắn gọn đọc được trên trang (chỉ cần điền nếu
  đây là trang đầu văn bản, nếu không để chuỗi rỗng "").
"""


def render_page_to_png_bytes(page, dpi=RENDER_DPI):
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    return pix.tobytes("png")


def ask_mistral_about_page(client, model, png_bytes, prev_ma, prev_png_bytes=None, so_lan_thu_lai=5):
    """Gửi trang đang xét cho Mistral phân loại. Nếu có prev_png_bytes (ảnh
    render của trang liền trước trong file gốc), gửi KÈM CẢ 2 ảnh (trang
    trước + trang đang xét) để model so sánh trực quan, tránh nhầm trang
    tiếp nối (cùng logo/header lặp lại) thành văn bản mới."""
    has_prev_image = prev_png_bytes is not None
    prompt = build_prompt_text(prev_ma, has_prev_image)
    b64_data = base64.b64encode(png_bytes).decode("utf-8")

    content = [{"type": "text", "text": prompt}]
    if has_prev_image:
        prev_b64 = base64.b64encode(prev_png_bytes).decode("utf-8")
        content.append({"type": "text", "text": "== ẢNH 1: TRANG TRƯỚC =="})
        content.append({"type": "image_url", "image_url": f"data:image/png;base64,{prev_b64}"})
        content.append({"type": "text", "text": "== ẢNH 2: TRANG ĐANG XÉT (cần phân loại) =="})
        content.append({"type": "image_url", "image_url": f"data:image/png;base64,{b64_data}"})
    else:
        content.append({"type": "image_url", "image_url": f"data:image/png;base64,{b64_data}"})

    messages = [
        {
            "role": "system",
            "content": "Bạn là trợ lý phân loại tài liệu scan. Chỉ trả lời theo đúng schema JSON được yêu cầu.",
        },
        {
            "role": "user",
            "content": content,
        },
    ]

    for attempt in range(so_lan_thu_lai):
        try:
            resp = client.chat.parse(
                model=model,
                messages=messages,
                response_format=PhanLoaiTrang,
                max_tokens=300,
                temperature=0,
            )
            msg = resp.choices[0].message
            parsed = getattr(msg, "parsed", None)
            if parsed is not None:
                info = parsed.model_dump()
            else:
                info = json.loads(msg.content)

            if info.get("ma_loai_giay_to") not in MA_LOAI_HOP_LE:
                info["ma_loai_giay_to"] = "KHAC"
            return info
        except Exception as e:  # noqa: BLE001 - muốn bắt cả lỗi API của Mistral
            loi_text = str(e)
            loi_lower = loi_text.lower()
            if "429" in loi_text or "rate limit" in loi_lower or "capacity" in loi_lower:
                cho_giay = 20
                print(f"    [!] Vượt quota/rate limit (429), lần {attempt + 1}/{so_lan_thu_lai}. Đợi {cho_giay}s rồi thử lại...")
                time.sleep(cho_giay)
            elif ("model" in loi_lower and ("not found" in loi_lower or "invalid" in loi_lower)) or "404" in loi_text:
                sys.exit(
                    f"\n[LỖI] Model '{model}' không tồn tại hoặc không hỗ trợ.\n"
                    f"Hãy kiểm tra danh sách model hiện hành tại https://docs.mistral.ai/models\n"
                    f"rồi chạy lại với --model <tên_model_mới>.\nChi tiết lỗi: {loi_text}"
                )
            elif "401" in loi_text or "unauthorized" in loi_lower:
                sys.exit(f"\n[LỖI] API key không hợp lệ hoặc hết hạn. Chi tiết lỗi: {loi_text}")
            else:
                print(f"    [!] Lỗi lần {attempt + 1}/{so_lan_thu_lai}: {e}. Thử lại sau 5s...")
                time.sleep(5)

    print("    [!] Vẫn lỗi sau nhiều lần thử, tạm gán tiếp nối văn bản trước.")
    return {
        "is_new_document_start": False,
        "ma_loai_giay_to": prev_ma or "KHAC",
        "title": "",
    }


def group_pages_into_documents(page_infos):
    docs = []
    current = None
    for i, info in enumerate(page_infos):
        if info["is_new_document_start"] or current is None:
            if current is not None:
                docs.append(current)
            current = {
                "start": i,
                "end": i,
                "ma_loai": info.get("ma_loai_giay_to") or "KHAC",
                "title": info.get("title") or "",
            }
        else:
            current["end"] = i
            if not current["title"] and info.get("title"):
                current["title"] = info["title"]
    if current is not None:
        docs.append(current)
    return docs


def gan_nhan_theo_pl02(documents):
    """Trả về list nhãn (chuỗi) tương ứng từng văn bản, theo đúng quy tắc
    PL02: chỉ thêm hậu tố -STT khi có từ 2 văn bản cùng loại trở lên,
    ví dụ: "HDTHUVIEC" (chỉ có 1) nhưng "HDXDTH-01", "HDXDTH-02" (có 2)."""
    dem_theo_loai = {}
    for d in documents:
        dem_theo_loai[d["ma_loai"]] = dem_theo_loai.get(d["ma_loai"], 0) + 1

    stt_dang_dung = {}
    nhan_list = []
    for d in documents:
        ma_loai = d["ma_loai"]
        can_them_stt = dem_theo_loai[ma_loai] >= 2
        if can_them_stt:
            stt_dang_dung[ma_loai] = stt_dang_dung.get(ma_loai, 0) + 1
            nhan = f"{ma_loai}-{stt_dang_dung[ma_loai]:02d}"
        else:
            nhan = ma_loai
        nhan_list.append(nhan)
    return nhan_list


def dong_dau_typewriter(page, text, mode="annot"):
    """Chèn (typewriter) nhãn màu đỏ vào góc trên-trái của trang.

    mode="annot" (mặc định): tạo PDF FreeText Annotation — giống công cụ
        "Typewriter"/"Add Text" của Acrobat. Ưu điểm: sau khi lưu file, có
        thể MỞ LẠI bằng Acrobat/Foxit/PDF-XChange... và CLICK VÀO ĐỂ SỬA chữ
        trực tiếp, không cần chạy lại script. Nhược điểm: một số công cụ
        tách file dựa vào lớp text thô của PDF (không phải OCR trên ảnh
        render) có thể KHÔNG đọc được chữ trong annotation.
    mode="text": in đè chữ thẳng vào nội dung trang (như bản cũ) — chắc
        chắn được mọi công cụ đọc được (kể cả trích xuất text thô lẫn OCR
        trên ảnh render), nhưng KHÔNG sửa lại được sau khi lưu, muốn sửa
        phải chạy lại script.
    """
    rect = fitz.Rect(
        STAMP_MARGIN_X,
        STAMP_MARGIN_Y,
        STAMP_MARGIN_X + STAMP_BOX_WIDTH,
        STAMP_MARGIN_Y + STAMP_BOX_HEIGHT,
    )
    if mode == "annot":
        # Lưu ý: FreeText annotation của PyMuPDF chỉ hỗ trợ font Base-14
        # (Helvetica/Times/Courier), không có bản đậm (bold) từ MuPDF v1.16+.
        annot = page.add_freetext_annot(
            rect,
            text,
            fontsize=STAMP_FONTSIZE,
            fontname="helv",
            text_color=STAMP_COLOR,
            fill_color=None,   # nền trong suốt, không che nội dung gốc
            border_width=0,    # không viền
            align=0,           # căn trái
        )
        annot.set_border(width=0)
        annot.update()
    else:
        page.insert_textbox(
            rect,
            text,
            fontsize=STAMP_FONTSIZE,
            fontname=STAMP_FONT,
            color=STAMP_COLOR,
            align=0,  # căn trái
        )


def stamp_pdf(input_path, output_path, api_key, model, delay_giay, stamp_mode="annot", config_path=None):
    # Nạp danh mục PL02 + quy tắc phân biệt NGAY TẠI ĐÂY (không phải lúc
    # import module) — an toàn cho cả CLI lẫn GUI --windowed, và luôn đọc
    # đúng bản danh_muc_pl02.json mới nhất mỗi lần chạy.
    nap_cau_hinh(config_path)

    client = Mistral(api_key=api_key)

    doc = fitz.open(input_path)
    n_pages = len(doc)
    print(f"Tổng số trang: {n_pages}")

    page_infos = []
    prev_ma = None
    prev_png_bytes = None  # ảnh trang trước, gửi kèm để model so sánh trực quan
    for i in range(n_pages):
        print(f"  Đang đọc trang {i + 1}/{n_pages}...")
        png_bytes = render_page_to_png_bytes(doc[i])
        info = ask_mistral_about_page(client, model, png_bytes, prev_ma, prev_png_bytes=prev_png_bytes)
        page_infos.append(info)
        prev_ma = info.get("ma_loai_giay_to") or prev_ma
        prev_png_bytes = png_bytes
        if delay_giay > 0 and i < n_pages - 1:
            time.sleep(delay_giay)  # giãn cách chủ động để không vượt quota RPM

    documents = group_pages_into_documents(page_infos)
    print(f"\nPhát hiện {len(documents)} văn bản riêng biệt.\n")

    nhan_list = gan_nhan_theo_pl02(documents)

    log_rows = []
    for d, nhan in zip(documents, nhan_list):
        trang_dau = doc[d["start"]]
        dong_dau_typewriter(trang_dau, nhan, mode=stamp_mode)

        so_trang = d["end"] - d["start"] + 1
        print(f"  -> Trang {d['start']+1}-{d['end']+1} ({so_trang} trang): đóng dấu \"{nhan}\"")
        log_rows.append({
            "nhan_da_dong_dau": nhan,
            "trang_bat_dau": d["start"] + 1,
            "trang_ket_thuc": d["end"] + 1,
            "so_trang": so_trang,
            "ma_loai_giay_to": d["ma_loai"],
            "tieu_de_doc_duoc": d["title"],
        })

    doc.save(output_path)
    doc.close()

    log_path = os.path.splitext(output_path)[0] + "_log_danh_dau.csv"
    with open(log_path, "w", newline="", encoding="utf-8-sig") as f:
        writer_csv = csv.DictWriter(f, fieldnames=list(log_rows[0].keys()))
        writer_csv.writeheader()
        writer_csv.writerows(log_rows)

    print(f"\nXong. File đã đánh dấu: {output_path}")
    print(f"File log kiểm tra: {log_path}")
    print("LƯU Ý: hãy mở file log để đối chiếu lại mã loại giấy tờ, đặc biệt")
    print("các trường hợp nhiều bằng cấp/hợp đồng cùng loại — AI chỉ đánh STT")
    print("theo THỨ TỰ TRANG xuất hiện trong file gốc, cần tự kiểm tra lại")
    print("trước khi dùng công cụ khác tách file dựa theo nhãn này.")


def main():
    parser = argparse.ArgumentParser(
        description="Dùng Mistral AI (Vision) nhận diện loại giấy tờ và typewriter mã viết tắt PL02 "
                     "(màu đỏ) vào góc trên-trái trang đầu mỗi văn bản. KHÔNG tách file."
    )
    parser.add_argument("input_pdf", help="Đường dẫn file PDF gốc (đã gộp nhiều văn bản)")
    parser.add_argument("--outfile", default=None,
                         help="Đường dẫn file PDF kết quả (mặc định: <tên_file_gốc>_da_danh_dau.pdf)")
    parser.add_argument("--config", default=None,
                         help="Đường dẫn file cấu hình JSON danh mục PL02 + quy tắc phân biệt (mặc định: "
                              f"'{CONFIG_FILENAME}' cùng thư mục với script/exe). Sửa file này để cập nhật "
                              "danh mục/quy tắc MÀ KHÔNG CẦN build lại chương trình.")
    parser.add_argument("--api-key", default=None,
                         help="Mistral API key (nếu không truyền, sẽ thử lấy biến môi trường MISTRAL_API_KEY, "
                              "nếu vẫn không có sẽ hỏi trực tiếp)")
    parser.add_argument("--model", default=DEFAULT_MODEL,
                         help=f"Model Mistral (có Vision) dùng để đọc trang (mặc định: {DEFAULT_MODEL})")
    parser.add_argument("--delay", type=float, default=3.0,
                         help="Số giây nghỉ giữa mỗi trang để tránh vượt quota (mặc định 3s, phù hợp gói Free "
                              "~1 request/giây của Mistral). Nếu dùng gói trả phí (pay-as-you-go/Scale), có thể giảm xuống 0.")
    parser.add_argument("--stamp-mode", choices=["annot", "text"], default="annot",
                         help="'annot' (mặc định): tạo con dấu dạng FreeText annotation, MỞ LẠI BẰNG "
                              "Acrobat/Foxit... ĐỂ SỬA ĐƯỢC trực tiếp. 'text': in đè cố định vào trang, "
                              "không sửa lại được nhưng chắc chắn mọi công cụ đọc text/OCR đều thấy được.")
    parser.add_argument("--font-size", type=float, default=None, help="Cỡ chữ con dấu (mặc định 15)")
    parser.add_argument("--margin-x", type=float, default=None, help="Khoảng cách từ mép trái, đơn vị point (mặc định 36)")
    parser.add_argument("--margin-y", type=float, default=None, help="Khoảng cách từ mép trên, đơn vị point (mặc định 12)")
    args = parser.parse_args()

    global STAMP_FONTSIZE, STAMP_MARGIN_X, STAMP_MARGIN_Y
    if args.font_size is not None:
        STAMP_FONTSIZE = args.font_size
    if args.margin_x is not None:
        STAMP_MARGIN_X = args.margin_x
    if args.margin_y is not None:
        STAMP_MARGIN_Y = args.margin_y

    api_key = args.api_key or os.environ.get("MISTRAL_API_KEY") or input(
        "Nhập Mistral API key (lấy tại https://console.mistral.ai/ mục API Keys): "
    ).strip()
    if not api_key:
        sys.exit("Chưa có API key. Truyền --api-key, đặt biến môi trường MISTRAL_API_KEY, hoặc nhập khi được hỏi.")
    if not os.path.isfile(args.input_pdf):
        sys.exit(f"Không tìm thấy file: {args.input_pdf}")

    outfile = args.outfile
    if not outfile:
        goc, _ext = os.path.splitext(args.input_pdf)
        outfile = f"{goc}_da_danh_dau.pdf"

    stamp_pdf(args.input_pdf, outfile, api_key, args.model, args.delay,
              stamp_mode=args.stamp_mode, config_path=args.config)


if __name__ == "__main__":
    main()
