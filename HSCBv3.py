#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HSCB_split_stamp.py — Giao diện (GUI) tách file PDF hồ sơ nhân sự theo mã
PL02, dùng cho file ĐÃ ĐƯỢC ĐÓNG DẤU SẴN bởi stamp_pdf_mistral.py.

KHÔNG dùng OCR, KHÔNG dò màu pixel. Vì stamp_pdf_mistral.py chèn nhãn bằng
page.add_freetext_annot() (FreeText Annotation chứa text thật) hoặc
page.insert_textbox() (chữ thật in thẳng vào trang) — cả 2 đều là LỚP TEXT
THẬT trên PDF, không phải hình ảnh — nên chỉ cần đọc thẳng nội dung đó bằng
PyMuPDF là chính xác 100%, nhanh hơn nhiều và không cần cài Tesseract.

CÁCH BUILD FILE .EXE (chạy trên máy Windows có cài Python):
    1. pip install pymupdf pypdf pyinstaller
    2. Mở Command Prompt tại thư mục chứa file này, chạy:
       pyinstaller --onefile --windowed --name TachHoSoPDF HSCB_split_stamp.py
    3. File .exe nằm trong thư mục "dist\\TachHoSoPDF.exe"
"""

import csv
import json
import os
import re
import sys
import threading
import queue
import unicodedata
import difflib
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ---------------------------------------------------------------------------
# BẢNG MÀU VIETINBANK (xanh dương - đỏ)
# ---------------------------------------------------------------------------
VTB_BLUE_DARK = "#003C71"
VTB_BLUE = "#005B94"
VTB_BLUE_LIGHT = "#EAF2FA"
VTB_RED = "#D31145"
VTB_RED_DARK = "#A50D37"
VTB_GREEN = "#1E8449"
VTB_BG = "#F4F7FB"
VTB_CARD = "#FFFFFF"
VTB_BORDER = "#D7E0EA"
VTB_TEXT = "#1A2733"
VTB_MUTED = "#6B7A8D"

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    PdfReader = PdfWriter = None


# ---------------------------------------------------------------------------
# VỊ TRÍ NHÃN — PHẢI KHỚP ĐÚNG với stamp_pdf_mistral.py (STAMP_MARGIN_X/Y,
# STAMP_BOX_WIDTH/HEIGHT). Chỉ dùng để trích text dự phòng khi công cụ đóng
# dấu chạy ở mode="text" (in thẳng chữ, không phải annotation).
# ---------------------------------------------------------------------------
STAMP_MARGIN_X = 36
STAMP_MARGIN_Y = 12
STAMP_BOX_WIDTH = 240
STAMP_BOX_HEIGHT = 28
# Vùng dò dự phòng nới rộng hơn box gốc một chút, phòng khi chữ dài tràn nhẹ
_VUNG_DU_PHONG = fitz.Rect(
    0, 0, STAMP_MARGIN_X + STAMP_BOX_WIDTH + 60, STAMP_MARGIN_Y + STAMP_BOX_HEIGHT + 12
) if fitz else None

MATCH_ACCEPT_THRESHOLD = 0.35
REVIEW_FLAG_THRESHOLD = 0.85

# Danh mục MẶC ĐỊNH — chỉ dùng làm fallback khi không có file cấu hình
# 'danh_muc_pl02.json' (CÙNG NGUỒN với stamp_pdf_mistral.py). Khi có file
# cấu hình, hàm nap_danh_muc() sẽ nạp danh sách mã từ file đó để bộ tách
# luôn nhận diện đúng mọi mã mà bộ đóng dấu có thể tạo ra (vd: CAMKETKHAC).
DANH_MUC_MA_MAC_DINH = {
    "QDTUYENDUNG": "Quyết định tuyển dụng",
    "THONGBAOLUONG": "Thông báo lương",
    "SOYEULYLICH": "Sơ yếu lý lịch",
    "GIAYKHAISINH": "Giấy khai sinh",
    "CCCD": "Căn cước công dân",
    "CMND": "Chứng minh nhân dân",
    "HOCHIEU": "Hộ chiếu",
    "DINHDANH": "Định danh điện tử",
    "BANGDIEMDH": "Bảng điểm đại học",
    "BANGTIENSI": "Bằng tiến sĩ",
    "BANGTHACSI": "Bằng thạc sĩ",
    "BANGDAIHOC": "Bằng đại học",
    "CHUNGCHIKHAC": "Chứng chỉ khác",
    "GIAYKHAMSK": "Giấy khám sức khỏe",
    "CV": "Phiếu thông tin ứng viên (CV)",
    "QDCHAMDUT": "Quyết định chấm dứt",
    "CAMKETBAOMAT": "Cam kết bảo mật thông tin",
    "HDTHUVIEC": "Hợp đồng thử việc",
    "HDXDTH": "HĐLĐ xác định thời hạn",
    "HDKXDTH": "HĐLĐ không xác định thời hạn",
    "PHULUCHD": "Phụ lục hợp đồng",
    "HDDAOTAO": "Hợp đồng đào tạo",
    "QDBONHIEM": "Quyết định Bổ nhiệm/Điều động bổ nhiệm",
    "QDDIEUDONG": "Quyết định điều động",
    "QDCHUYENDOI": "Quyết định chuyển đổi công việc",
    "QDTAMHOAN": "Quyết định tạm hoãn HĐLĐ",
    "QDDIEUCHINH": "Quyết định điều chỉnh lương",
    "QDNGHIHUU": "Quyết định nghỉ hưu",
    "QDKYLUAT": "Kỷ luật",
    "KEKHAITS": "Kê khai tài sản",
    "KHAC": "Khác (không thuộc danh mục PL02 chuẩn)",
}

# Được nạp runtime bởi nap_danh_muc(): ưu tiên từ file danh_muc_pl02.json,
# fallback sang danh mục mặc định ở trên.
DANH_MUC_MA = {}
MA_LIST = []

CONFIG_FILENAME = "danh_muc_pl02.json"


def _duong_dan_thu_muc_goc():
    """Thư mục chứa file cấu hình: cùng thư mục với file .exe khi đã build
    bằng PyInstaller (sys.frozen), hoặc cùng thư mục với script .py."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def nap_danh_muc(duong_dan=None):
    """Nạp danh mục PL02 từ file 'danh_muc_pl02.json' — CÙNG NGUỒN với
    stamp_pdf_mistral.py — để bộ tách nhận diện đúng mọi mã mà bộ đóng dấu
    có thể tạo ra. Nếu chưa có file hoặc đọc lỗi, dùng danh mục mặc định
    nhúng trong code. Tên hiển thị ưu tiên tên ngắn có trong danh mục nhúng
    (CSV gọn); mã mới chỉ có mô tả dài trong cấu hình thì dùng mô tả đó."""
    global DANH_MUC_MA, MA_LIST
    duong_dan = duong_dan or os.path.join(_duong_dan_thu_muc_goc(), CONFIG_FILENAME)
    danh_sach = []
    if os.path.isfile(duong_dan):
        try:
            with open(duong_dan, "r", encoding="utf-8") as f:
                du_lieu = json.load(f)
            danh_sach = [
                (item["ma"].strip(), item["mo_ta"].strip())
                for item in du_lieu.get("danh_muc_loai_giay_to", [])
                if item.get("ma") and item["ma"].strip()
            ]
        except Exception:  # noqa: BLE001 - file lỗi thì dùng danh mục mặc định
            danh_sach = []
    if not danh_sach:
        danh_sach = list(DANH_MUC_MA_MAC_DINH.items())

    moi = {}
    for ma, mo_ta in danh_sach:
        moi[ma] = DANH_MUC_MA_MAC_DINH.get(ma, mo_ta)
    moi.setdefault("KHAC", DANH_MUC_MA_MAC_DINH.get("KHAC", "Khác"))
    DANH_MUC_MA = moi
    MA_LIST = list(DANH_MUC_MA.keys())


nap_danh_muc()


def chuan_hoa_khong_dau(text):
    if not text:
        return ""
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = text.replace("Đ", "D").replace("đ", "d")
    return text.upper()


def _tach_phan_hau_to(text):
    """Tách phần chữ (mã loại) khỏi phần số thứ tự (vd "HDXDTH-01" -> "HDXDTH")."""
    m = re.match(r"^([A-Z]+)-?([0-9OIL]{0,3})$", text)
    if m:
        return m.group(1), m.group(2)
    return text, ""


def so_khop_ma_pl02(raw_text):
    """So khớp text đọc được với danh mục 29 mã PL02 (dự phòng cho trường
    hợp nhãn bị gõ sai/thiếu ký tự), trả về (mã_khớp_nhất, tỷ_lệ_khớp, raw)."""
    raw = re.sub(r"[^A-Z0-9-]", "", raw_text.upper())
    if not raw:
        return None, 0.0, raw
    phan_chu, _phan_so = _tach_phan_hau_to(raw)
    if not phan_chu:
        phan_chu = raw
    best = difflib.get_close_matches(phan_chu, MA_LIST, n=1, cutoff=0.0)
    if not best:
        return None, 0.0, raw
    ty_le = difflib.SequenceMatcher(None, phan_chu, best[0]).ratio()
    return best[0], ty_le, raw


def doc_nhan_da_dong_dau(page):
    """Đọc nhãn PL02 mà stamp_pdf_mistral.py đã đóng dấu trên trang, KHÔNG
    dùng OCR — đọc thẳng lớp text/annotation thật nên chính xác tuyệt đối,
    miễn nhãn còn nguyên. Trả về text thô (str) hoặc None nếu trang không
    có nhãn (thuộc cùng văn bản với trang trước)."""
    # 1) Ưu tiên FreeText Annotation — chế độ mặc định (mode="annot") của
    #    công cụ đóng dấu, cho phép mở lại bằng Acrobat/Foxit để sửa tay.
    for annot in (page.annots() or []):
        if annot.type[1] == "FreeText":
            text = (annot.info.get("content") or "").strip()
            if text:
                return text

    # 2) Dự phòng: chế độ "text" (in thẳng chữ vào nội dung trang) — trích
    #    text thô đúng vùng góc trên-trái nơi công cụ đóng dấu chèn chữ.
    if _VUNG_DU_PHONG is not None:
        text = page.get_text("text", clip=_VUNG_DU_PHONG).strip()
        if text:
            return text.splitlines()[0].strip()

    return None


def doc_thong_tin_nhan(page):
    """Đọc + so khớp nhãn của 1 trang. Trả về dict thông tin, hoặc None nếu
    trang không có nhãn (nối tiếp văn bản trước)."""
    nhan_goc = doc_nhan_da_dong_dau(page)
    if nhan_goc is None:
        return None

    ma, ty_le, raw = so_khop_ma_pl02(nhan_goc)
    if ty_le >= MATCH_ACCEPT_THRESHOLD and ma is not None:
        ma_loai = ma
    else:
        ma_loai = raw if raw else "KHAC"

    return {
        "ma_loai": ma_loai,
        "ty_le_khop": round(ty_le, 2),
        "nhan_goc": nhan_goc,
        "can_kiem_tra": ty_le < REVIEW_FLAG_THRESHOLD,
    }


def gom_trang_thanh_van_ban(page_infos):
    docs = []
    current = None
    for i, info in enumerate(page_infos):
        if info is not None or current is None:
            if current is not None:
                docs.append(current)
            if info is None:
                info = {"ma_loai": "KHAC", "ty_le_khop": 0.0, "nhan_goc": "", "can_kiem_tra": True}
            current = {
                "start": i,
                "end": i,
                "ma_loai": info["ma_loai"],
                "ty_le_khop": info["ty_le_khop"],
                "nhan_goc": info["nhan_goc"],
                "can_kiem_tra": info["can_kiem_tra"],
            }
        else:
            current["end"] = i
    if current is not None:
        docs.append(current)
    return docs


def dat_ten_file_theo_pl02(documents, ma_can_bo, ho_ten_slug):
    dem_theo_loai = {}
    for d in documents:
        dem_theo_loai[d["ma_loai"]] = dem_theo_loai.get(d["ma_loai"], 0) + 1

    stt_dang_dung = {}
    ten_file_list = []
    for d in documents:
        ma_loai = d["ma_loai"]
        can_them_stt = dem_theo_loai[ma_loai] >= 2
        if can_them_stt:
            stt_dang_dung[ma_loai] = stt_dang_dung.get(ma_loai, 0) + 1
            phan_stt = f"-{stt_dang_dung[ma_loai]:02d}"
        else:
            phan_stt = ""
        ten_file_list.append(f"{ma_can_bo}-{ho_ten_slug}-{ma_loai}{phan_stt}")
    return ten_file_list


class NgungXuLy(Exception):
    """Báo hiệu người dùng bấm Hủy giữa chừng."""
    pass


def split_pdf(input_path, outdir, ma_can_bo, ho_ten, log_fn=print, should_stop=None):
    nap_danh_muc()
    os.makedirs(outdir, exist_ok=True)
    ho_ten_slug = chuan_hoa_khong_dau(ho_ten).replace(" ", "-")
    ho_ten_slug = re.sub(r"[^A-Z0-9-]", "", ho_ten_slug)

    doc = fitz.open(input_path)
    n_pages = len(doc)
    log_fn(f"Tổng số trang: {n_pages}")

    page_infos = []
    for i in range(n_pages):
        if should_stop and should_stop():
            doc.close()
            raise NgungXuLy()
        info = doc_thong_tin_nhan(doc[i])
        if info is None:
            log_fn(f"  Trang {i + 1}/{n_pages}: không có nhãn -> nối tiếp văn bản trước")
        else:
            canh_bao = "  [CẦN KIỂM TRA]" if info["can_kiem_tra"] else ""
            log_fn(
                f"  Trang {i + 1}/{n_pages}: nhãn='{info['nhan_goc']}' "
                f"-> mã={info['ma_loai']} (khớp {info['ty_le_khop']}){canh_bao}"
            )
        page_infos.append(info)
    doc.close()

    documents = gom_trang_thanh_van_ban(page_infos)
    log_fn(f"\nPhát hiện {len(documents)} văn bản riêng biệt.\n")

    ten_file_list = dat_ten_file_theo_pl02(documents, ma_can_bo, ho_ten_slug)

    reader = PdfReader(input_path)
    log_rows = []
    used_names = set()

    for d, ten_goc in zip(documents, ten_file_list):
        writer = PdfWriter()
        for p in range(d["start"], d["end"] + 1):
            writer.add_page(reader.pages[p])

        name = ten_goc
        n = 1
        while name in used_names:
            n += 1
            name = f"{ten_goc}_dup{n}"
        used_names.add(name)

        out_path = os.path.join(outdir, f"{name}.pdf")
        with open(out_path, "wb") as f:
            writer.write(f)

        so_trang = d["end"] - d["start"] + 1
        log_fn(f"  -> {name}.pdf  (trang {d['start']+1}-{d['end']+1}, {so_trang} trang)")
        log_rows.append({
            "file": f"{name}.pdf",
            "trang_bat_dau": d["start"] + 1,
            "trang_ket_thuc": d["end"] + 1,
            "so_trang": so_trang,
            "ma_loai_giay_to": d["ma_loai"],
            "ten_day_du_loai": DANH_MUC_MA.get(d["ma_loai"], ""),
            "nhan_da_doc": d["nhan_goc"],
            "ty_le_khop": d["ty_le_khop"],
            "can_kiem_tra": "CO" if d["can_kiem_tra"] else "",
        })

    log_path = os.path.join(outdir, "_log_tach_file.csv")
    with open(log_path, "w", newline="", encoding="utf-8-sig") as f:
        writer_csv = csv.DictWriter(f, fieldnames=list(log_rows[0].keys()))
        writer_csv.writeheader()
        writer_csv.writerows(log_rows)

    log_fn(f"\nXong. File log kiểm tra: {log_path}")
    log_fn("*** QUAN TRỌNG: hãy mở file log (cột 'can_kiem_tra' = CO) và các")
    log_fn("PDF vừa tách để kiểm tra lại trước khi dùng chính thức.")
    return outdir, log_path


# ---------------------------------------------------------------------------
# GIAO DIỆN (GUI)
# ---------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tách hồ sơ PDF theo mã PL02 CV 7761 — VietinBank chi nhánh Bình Thuận")
        self.geometry("820x640")
        self.minsize(700, 520)
        self.configure(bg=VTB_BG)

        self.msg_queue = queue.Queue()
        self.worker_thread = None
        self.stop_flag = threading.Event()
        self._outdir_manual = False
        self._dang_cap_nhat_tu_dong = False

        self._setup_style()
        self._build_ui()
        self._kiem_tra_thu_vien()
        self.after(100, self._poll_queue)

    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", background=VTB_BG, foreground=VTB_TEXT, font=("Segoe UI", 10))
        style.configure("TFrame", background=VTB_BG)
        style.configure("Card.TFrame", background=VTB_CARD)
        style.configure("Header.TFrame", background=VTB_BLUE_DARK)

        style.configure("TLabel", background=VTB_BG, foreground=VTB_TEXT)
        style.configure("Card.TLabel", background=VTB_CARD, foreground=VTB_TEXT)
        style.configure("Field.TLabel", background=VTB_CARD, foreground=VTB_BLUE_DARK,
                         font=("Segoe UI", 10, "bold"))
        style.configure("Header.TLabel", background=VTB_BLUE_DARK, foreground="#FFFFFF",
                         font=("Segoe UI", 15, "bold"))
        style.configure("SubHeader.TLabel", background=VTB_BLUE_DARK, foreground="#CFE3F5",
                         font=("Segoe UI", 9))

        style.configure("Hyperlink.TLabel", background=VTB_BLUE_DARK,
                        foreground="#8FD3FF", font=("Segoe UI", 9, "underline"))
        style.configure("Muted.TLabel", background=VTB_CARD, foreground=VTB_MUTED)
        style.configure("Warn.TLabel", background=VTB_CARD, foreground=VTB_RED,
                         font=("Segoe UI", 9, "bold"))

        style.configure("TEntry", fieldbackground="#FFFFFF", foreground=VTB_TEXT,
                         bordercolor=VTB_BORDER, lightcolor=VTB_BORDER, darkcolor=VTB_BORDER,
                         padding=5)
        style.map("TEntry", bordercolor=[("focus", VTB_BLUE)])

        style.configure("Primary.TButton", background=VTB_BLUE, foreground="#FFFFFF",
                         font=("Segoe UI", 10, "bold"), padding=(14, 8), borderwidth=0)
        style.map("Primary.TButton",
                  background=[("disabled", "#9FB9CC"), ("active", VTB_BLUE_DARK)],
                  foreground=[("disabled", "#EFEFEF")])

        style.configure("Secondary.TButton", background="#E7EEF5", foreground=VTB_BLUE_DARK,
                         font=("Segoe UI", 10), padding=(12, 7), borderwidth=0)
        style.map("Secondary.TButton",
                  background=[("disabled", "#EFEFEF"), ("active", "#D6E3F0")],
                  foreground=[("disabled", VTB_MUTED)])

        style.configure("Danger.TButton", background=VTB_RED, foreground="#FFFFFF",
                         font=("Segoe UI", 10, "bold"), padding=(12, 7), borderwidth=0)
        style.map("Danger.TButton",
                  background=[("disabled", "#E9AEBD"), ("active", VTB_RED_DARK)],
                  foreground=[("disabled", "#F6F6F6")])

        style.configure("TProgressbar", background=VTB_RED, troughcolor="#E1E9F1",
                         bordercolor="#E1E9F1", lightcolor=VTB_RED, darkcolor=VTB_RED,
                         thickness=8)

    def _build_ui(self):
        header = ttk.Frame(self, style="Header.TFrame")
        header.pack(fill="x", side="top")
        inner_header = ttk.Frame(header, style="Header.TFrame")
        inner_header.pack(fill="x", padx=18, pady=(14, 12))
        ttk.Label(inner_header, text="📜 TÁCH HỒ SƠ PDF THEO MÃ PL02 CV 7761",
                  style="Header.TLabel").pack(anchor="w")

        link_row = ttk.Frame(inner_header, style="Header.TFrame")
        link_row.pack(anchor="w")
        ttk.Label(link_row, text="🌍 VIETINBANK CHI NHÁNH BÌNH THUẬN - ",
                  style="SubHeader.TLabel").pack(side="left")
        self.website_link = ttk.Label(link_row, text="https://vietinbank.cn600.website",
                                       style="Hyperlink.TLabel", cursor="hand2")
        self.website_link.pack(side="left")
        self.website_link.bind("<Button-1>", lambda e: webbrowser.open("https://vietinbank.cn600.website"))
        self.website_link.bind("<Enter>", lambda e: self.website_link.configure(foreground="#FFFFFF"))
        self.website_link.bind("<Leave>", lambda e: self.website_link.configure(foreground="#8FD3FF"))

        ttk.Label(
            inner_header,
            text="⚡ Đọc thẳng nhãn PL02 đã đóng dấu sẵn (annotation/text thật — KHÔNG cần OCR) "
                 "— chạy hoàn toàn local, không gửi dữ liệu ra ngoài",
            style="SubHeader.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        accent_bar = tk.Frame(self, bg=VTB_RED, height=3)
        accent_bar.pack(fill="x", side="top")

        body = ttk.Frame(self, style="TFrame")
        body.pack(fill="both", expand=True)
        pad = {"padx": 16, "pady": 8}

        # ── Thẻ: thông tin cán bộ ──
        card_info = ttk.Frame(body, style="Card.TFrame")
        card_info.pack(fill="x", **pad)
        frm_top = ttk.Frame(card_info, style="Card.TFrame")
        frm_top.pack(fill="x", padx=14, pady=12)

        ttk.Label(frm_top, text="Mã cán bộ", style="Field.TLabel").grid(row=0, column=0, sticky="w")
        self.entry_ma_can_bo = ttk.Entry(frm_top, width=30)
        self.entry_ma_can_bo.grid(row=1, column=0, sticky="we", padx=(0, 20), pady=(4, 0))

        ttk.Label(frm_top, text="Họ và tên", style="Field.TLabel").grid(row=0, column=1, sticky="w")
        self.entry_ho_ten = ttk.Entry(frm_top, width=30)
        self.entry_ho_ten.grid(row=1, column=1, sticky="we", pady=(4, 0))
        self.entry_ho_ten.bind("<KeyRelease>", self._cap_nhat_outdir_tu_dong)

        frm_top.columnconfigure(0, weight=1)
        frm_top.columnconfigure(1, weight=1)

        # ── Thẻ: chọn file / thư mục ──
        card_paths = ttk.Frame(body, style="Card.TFrame")
        card_paths.pack(fill="x", **pad)
        inner_paths = ttk.Frame(card_paths, style="Card.TFrame")
        inner_paths.pack(fill="x", padx=14, pady=12)

        ttk.Label(inner_paths, text="File PDF cần tách (đã đóng dấu)", style="Field.TLabel").grid(
            row=0, column=0, sticky="w", columnspan=2)
        self.entry_pdf = ttk.Entry(inner_paths)
        self.entry_pdf.grid(row=1, column=0, sticky="we", padx=(0, 8), pady=(4, 10))
        ttk.Button(inner_paths, text="Chọn file...", style="Secondary.TButton",
                   command=self._chon_file_pdf).grid(row=1, column=1, pady=(4, 10))

        ttk.Label(inner_paths, text="Thư mục lưu kết quả", style="Field.TLabel").grid(
            row=2, column=0, sticky="w", columnspan=2)
        self.entry_outdir = ttk.Entry(inner_paths)
        self.entry_outdir.grid(row=3, column=0, sticky="we", padx=(0, 8), pady=(4, 0))
        self.entry_outdir.bind("<KeyRelease>", self._outdir_da_sua_tay)
        ttk.Button(inner_paths, text="Chọn thư mục...", style="Secondary.TButton",
                   command=self._chon_outdir).grid(row=3, column=1, pady=(4, 0))

        inner_paths.columnconfigure(0, weight=1)

        # ── Thẻ: hành động ──
        card_actions = ttk.Frame(body, style="Card.TFrame")
        card_actions.pack(fill="x", **pad)
        inner_actions = ttk.Frame(card_actions, style="Card.TFrame")
        inner_actions.pack(fill="x", padx=14, pady=12)

        frm_btn = ttk.Frame(inner_actions, style="Card.TFrame")
        frm_btn.pack(fill="x")
        self.btn_run = ttk.Button(frm_btn, text="▶  Tách file PDF", style="Primary.TButton",
                                   command=self._bat_dau)
        self.btn_run.pack(side="left")
        self.btn_stop = ttk.Button(frm_btn, text="■  Hủy", style="Danger.TButton",
                                    command=self._huy, state="disabled")
        self.btn_stop.pack(side="left", padx=8)
        self.btn_open_outdir = ttk.Button(frm_btn, text="📂  Mở thư mục kết quả", style="Secondary.TButton",
                                           command=self._mo_outdir, state="disabled")
        self.btn_open_outdir.pack(side="left")

        self.lbl_trang_thai = ttk.Label(inner_actions, text="", style="Muted.TLabel",
                                         wraplength=740, justify="left")
        self.lbl_trang_thai.pack(fill="x", pady=(10, 0))

        self.progress = ttk.Progressbar(inner_actions, mode="indeterminate", style="TProgressbar")
        self.progress.pack(fill="x", pady=(10, 0))

        # ── Thẻ: nhật ký ──
        card_log = ttk.Frame(body, style="Card.TFrame")
        card_log.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        ttk.Label(card_log, text="Nhật ký xử lý", style="Field.TLabel").pack(anchor="w", padx=14, pady=(12, 4))

        frm_log = tk.Frame(card_log, bg=VTB_CARD, highlightbackground=VTB_BORDER, highlightthickness=1)
        frm_log.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self.txt_log = tk.Text(frm_log, wrap="word", state="disabled", bg="#FBFCFE",
                                fg=VTB_TEXT, insertbackground=VTB_TEXT, relief="flat",
                                padx=10, pady=8, font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(frm_log, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scrollbar.set)
        self.txt_log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.txt_log.tag_configure("canh_bao", foreground=VTB_RED)
        self.last_outdir = None

    def _kiem_tra_thu_vien(self):
        thieu = []
        if fitz is None:
            thieu.append("pymupdf")
        if PdfReader is None:
            thieu.append("pypdf")
        if thieu:
            messagebox.showerror(
                "Thiếu thư viện",
                "Thiếu thư viện: " + ", ".join(thieu) + "\nCài bằng: pip install pymupdf pypdf"
            )
        else:
            self.lbl_trang_thai.configure(
                text="✔ Sẵn sàng — đọc trực tiếp nhãn đã đóng dấu, không cần OCR.",
                style="Muted.TLabel",
            )

    def _ghi_log(self, dong):
        self.msg_queue.put(("log", dong))

    def _chon_file_pdf(self):
        path = filedialog.askopenfilename(
            title="Chọn file PDF cần tách (đã đóng dấu)",
            filetypes=[("File PDF", "*.pdf"), ("Tất cả file", "*.*")],
        )
        if path:
            self.entry_pdf.delete(0, tk.END)
            self.entry_pdf.insert(0, path)
            self._cap_nhat_outdir_tu_dong()

    def _ten_thu_muc_tu_ho_ten(self):
        ho_ten = self.entry_ho_ten.get().strip()
        if not ho_ten:
            return ""
        slug = chuan_hoa_khong_dau(ho_ten).replace(" ", "-")
        slug = re.sub(r"[^A-Z0-9-]", "", slug)
        slug = re.sub(r"-{2,}", "-", slug).strip("-")
        return slug

    def _cap_nhat_outdir_tu_dong(self, event=None):
        if self._outdir_manual:
            return
        pdf_path = self.entry_pdf.get().strip()
        if not pdf_path:
            return
        ten_thu_muc = self._ten_thu_muc_tu_ho_ten()
        if not ten_thu_muc:
            return
        goi_y = os.path.join(os.path.dirname(pdf_path), ten_thu_muc)
        self._dang_cap_nhat_tu_dong = True
        self.entry_outdir.delete(0, tk.END)
        self.entry_outdir.insert(0, goi_y)
        self._dang_cap_nhat_tu_dong = False

    def _outdir_da_sua_tay(self, event=None):
        if self._dang_cap_nhat_tu_dong:
            return
        self._outdir_manual = True

    def _chon_outdir(self):
        path = filedialog.askdirectory(title="Chọn thư mục lưu kết quả")
        if path:
            self._outdir_manual = True
            self.entry_outdir.delete(0, tk.END)
            self.entry_outdir.insert(0, path)

    def _mo_outdir(self):
        if self.last_outdir and os.path.isdir(self.last_outdir):
            if sys.platform.startswith("win"):
                os.startfile(self.last_outdir)  # noqa
            elif sys.platform == "darwin":
                os.system(f'open "{self.last_outdir}"')
            else:
                os.system(f'xdg-open "{self.last_outdir}"')

    def _bat_dau(self):
        ma_can_bo = self.entry_ma_can_bo.get().strip()
        ho_ten = self.entry_ho_ten.get().strip()
        pdf_path = self.entry_pdf.get().strip()
        outdir = self.entry_outdir.get().strip()

        if not ma_can_bo or not ho_ten:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng nhập Mã cán bộ và Họ và tên.")
            return
        if not pdf_path or not os.path.isfile(pdf_path):
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn file PDF hợp lệ.")
            return
        if not outdir:
            ten_thu_muc = self._ten_thu_muc_tu_ho_ten() or "ket_qua_tach"
            outdir = os.path.join(os.path.dirname(pdf_path), ten_thu_muc)

        if fitz is None or PdfReader is None:
            messagebox.showerror("Thiếu thư viện", "Vui lòng cài đủ thư viện rồi thử lại.")
            return

        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state="disabled")

        self.btn_run.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.btn_open_outdir.configure(state="disabled")
        self.progress.start(10)
        self.stop_flag.clear()

        self.worker_thread = threading.Thread(
            target=self._chay_nen, args=(pdf_path, outdir, ma_can_bo, ho_ten), daemon=True
        )
        self.worker_thread.start()

    def _chay_nen(self, pdf_path, outdir, ma_can_bo, ho_ten):
        try:
            out, _log = split_pdf(
                pdf_path, outdir, ma_can_bo, ho_ten,
                log_fn=self._ghi_log, should_stop=self.stop_flag.is_set,
            )
            self.msg_queue.put(("done", out))
        except NgungXuLy:
            self.msg_queue.put(("stopped", None))
        except Exception as e:  # noqa
            self.msg_queue.put(("error", str(e)))

    def _huy(self):
        self.stop_flag.set()
        self._ghi_log("\nĐang dừng...")

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "log":
                    self.txt_log.configure(state="normal")
                    if "CẦN KIỂM TRA" in payload or payload.startswith("*** QUAN TRỌNG"):
                        self.txt_log.insert(tk.END, payload + "\n", "canh_bao")
                    else:
                        self.txt_log.insert(tk.END, payload + "\n")
                    self.txt_log.see(tk.END)
                    self.txt_log.configure(state="disabled")
                elif kind == "done":
                    self.progress.stop()
                    self.btn_run.configure(state="normal")
                    self.btn_stop.configure(state="disabled")
                    self.btn_open_outdir.configure(state="normal")
                    self.last_outdir = payload
                    messagebox.showinfo("Hoàn tất", f"Đã tách xong. Kết quả lưu tại:\n{payload}")
                elif kind == "stopped":
                    self.progress.stop()
                    self.btn_run.configure(state="normal")
                    self.btn_stop.configure(state="disabled")
                    self._ghi_log("Đã hủy theo yêu cầu.")
                elif kind == "error":
                    self.progress.stop()
                    self.btn_run.configure(state="normal")
                    self.btn_stop.configure(state="disabled")
                    messagebox.showerror("Lỗi", payload)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
