#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gui_stamp_pdf.py — Giao diện Tkinter cho stamp_pdf_mistral.py.

Bắt buộc file này nằm CÙNG THƯ MỤC với stamp_pdf_mistral.py.

CÀI ĐẶT:
    pip install -U pymupdf mistralai pydantic

CHẠY:
    python gui_stamp_pdf.py

ĐÓNG GÓI THÀNH .EXE (Windows, dùng PyInstaller):
    pip install -U pyinstaller
    pyinstaller --noconfirm --onefile --windowed --name "sHSCB" --icon=shscb.ico --add-data "stamp_pdf_mistral.py;." gui_stamp_pdf.py
 
     File .exe sẽ nằm trong thư mục dist/. Có thể thêm --icon=ten_file.ico
    nếu muốn gắn icon riêng.
"""

import os
import re
import sys
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

try:
    import fitz  # PyMuPDF - chỉ dùng để đếm số trang cho thanh tiến độ
except ImportError:
    fitz = None

try:
    import stamp_pdf_mistral as spm
except ImportError as e:
    raise SystemExit(
        "Không tìm thấy stamp_pdf_mistral.py. Hãy đặt gui_stamp_pdf.py cùng "
        f"thư mục với stamp_pdf_mistral.py.\nChi tiết lỗi: {e}"
    )

# ---------------------------------------------------------------------------
# Bảng màu nhận diện VietinBank (theo tài liệu ý nghĩa logo)
# ---------------------------------------------------------------------------
BLUE = "#0072CE"
BLUE_DARK = "#00539C"
RED = "#ED1C24"
WHITE = "#FFFFFF"
GRAY_BG = "#F2F4F7"
TEXT_DARK = "#1A1A1A"
TEXT_MUTED = "#6B7280"


class StdoutRedirector:
    """Chuyển print() trong luồng xử lý nền vào một queue để GUI đọc."""

    def __init__(self, q):
        self.q = q

    def write(self, text):
        if text:
            self.q.put(("log", text))

    def flush(self):
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("VietinBank — Công cụ đánh dấu hồ sơ theo PL02 cv 7761")
        self.geometry("880x720")
        self.minsize(760, 620)
        self.configure(bg=WHITE)

        self.log_queue = queue.Queue()
        self.worker_thread = None
        self.total_pages = 0

        self._build_style()
        self._build_ui()
        self.after(120, self._poll_log_queue)

    # ------------------------------------------------------------------
    # Giao diện
    # ------------------------------------------------------------------
    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background=WHITE)
        style.configure("Card.TFrame", background=WHITE)
        style.configure("TLabel", background=WHITE, foreground=TEXT_DARK, font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=WHITE, foreground=TEXT_MUTED, font=("Segoe UI", 9))
        style.configure("Section.TLabel", background=WHITE, foreground=BLUE_DARK,
                         font=("Segoe UI", 11, "bold"))
        style.configure("TEntry", padding=6)
        style.configure("TCombobox", padding=6)
        style.configure("TRadiobutton", background=WHITE, font=("Segoe UI", 9))
        style.configure("Horizontal.TProgressbar", troughcolor=GRAY_BG, background=BLUE,
                         bordercolor=GRAY_BG, lightcolor=BLUE, darkcolor=BLUE)

    def _build_ui(self):
        # ---- Header (xanh dương, viền đỏ bên phải) ----
        header = tk.Frame(self, bg=BLUE, height=68)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        title_box = tk.Frame(header, bg=BLUE)
        title_box.pack(side="left", padx=20, pady=10)
        tk.Label(title_box, text="VietinBank", bg=BLUE, fg=WHITE,
                 font=("Segoe UI", 17, "bold")).pack(anchor="w")
        tk.Label(title_box, text="Đánh dấu hồ sơ nhân sự theo PL02 cv 7761 — powered by Mistral AI — made by TCTH-VTB600",
                 bg=BLUE, fg="#DCEBFF", font=("Segoe UI", 9)).pack(anchor="w")

        tk.Frame(header, bg=RED, width=6).pack(side="right", fill="y")

        # ---- Nội dung chính (scroll nếu cửa sổ nhỏ) ----
        outer = tk.Frame(self, bg=WHITE)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, bg=WHITE, highlightthickness=0)
        vscroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")

        body = ttk.Frame(canvas, style="TFrame", padding=(24, 18, 24, 18))
        body_id = canvas.create_window((0, 0), window=body, anchor="nw")

        def _on_body_configure(_event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfig(body_id, width=event.width)

        body.bind("<Configure>", _on_body_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        body.columnconfigure(1, weight=1)
        row = 0

        # ---- File đầu vào ----
        ttk.Label(body, text="1. File PDF hồ sơ", style="Section.TLabel").grid(
            row=row, column=0, columnspan=3, sticky="w", pady=(0, 6))
        row += 1

        self.input_var = tk.StringVar()
        ttk.Label(body, text="File đầu vào:").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(body, textvariable=self.input_var).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        ttk.Button(body, text="Chọn file...", command=self._choose_input).grid(row=row, column=2, pady=4)
        row += 1

        self.output_var = tk.StringVar()
        ttk.Label(body, text="File kết quả:").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(body, textvariable=self.output_var).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        ttk.Button(body, text="Chọn nơi lưu...", command=self._choose_output).grid(row=row, column=2, pady=4)
        row += 1

        ttk.Separator(body).grid(row=row, column=0, columnspan=3, sticky="ew", pady=14)
        row += 1

        # ---- API key & model ----
        ttk.Label(body, text="2. Kết nối Mistral AI", style="Section.TLabel").grid(
            row=row, column=0, columnspan=3, sticky="w", pady=(0, 6))
        row += 1

        self.api_key_var = tk.StringVar(value=os.environ.get("MISTRAL_API_KEY", ""))
        ttk.Label(body, text="Mistral API key:").grid(row=row, column=0, sticky="w", pady=4)
        self.api_key_entry = ttk.Entry(body, textvariable=self.api_key_var, show="•")
        self.api_key_entry.grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        self.show_key_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(body, text="Hiện", variable=self.show_key_var,
                         command=self._toggle_key_visibility).grid(row=row, column=2, sticky="w", pady=4)
        row += 1

        ttk.Label(body, text="", style="Muted.TLabel").grid(row=row, column=0, sticky="w")
        ttk.Label(body, text="Lấy API key tại console.mistral.ai (mục API Keys)",
                  style="Muted.TLabel").grid(row=row, column=1, sticky="w", padx=8)
        row += 1

        self.model_var = tk.StringVar(value=spm.DEFAULT_MODEL)
        ttk.Label(body, text="Model (Vision):").grid(row=row, column=0, sticky="w", pady=4)
        model_combo = ttk.Combobox(body, textvariable=self.model_var, values=(
            "mistral-small-latest",
            "pixtral-large-latest",
            "mistral-medium-latest",
            "mistral-large-latest",
        ))
        model_combo.grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        row += 1

        self.delay_var = tk.StringVar(value="3.0")
        ttk.Label(body, text="Độ trễ giữa các trang (giây):").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(body, textvariable=self.delay_var, width=10).grid(row=row, column=1, sticky="w", padx=8, pady=4)
        row += 1

        ttk.Separator(body).grid(row=row, column=0, columnspan=3, sticky="ew", pady=14)
        row += 1

        # ---- Con dấu ----
        ttk.Label(body, text="3. Kiểu con dấu (typewriter)", style="Section.TLabel").grid(
            row=row, column=0, columnspan=3, sticky="w", pady=(0, 6))
        row += 1

        self.stamp_mode_var = tk.StringVar(value="annot")
        mode_frame = ttk.Frame(body)
        mode_frame.grid(row=row, column=0, columnspan=3, sticky="w")
        ttk.Radiobutton(mode_frame, text="Annotation (sửa lại được bằng Acrobat/Foxit)",
                         variable=self.stamp_mode_var, value="annot").pack(anchor="w")
        ttk.Radiobutton(mode_frame, text="Text (in đè cố định, chắc chắn mọi công cụ đọc được)",
                         variable=self.stamp_mode_var, value="text").pack(anchor="w")
        row += 1

        # ---- Tuỳ chọn nâng cao ----
        adv_toggle = ttk.Checkbutton(body, text="Tuỳ chọn nâng cao (cỡ chữ / vị trí)",
                                      variable=tk.BooleanVar(), command=None)
        self.adv_visible = tk.BooleanVar(value=False)
        adv_btn = ttk.Checkbutton(body, text="Hiện tuỳ chọn nâng cao (cỡ chữ / vị trí)",
                                   variable=self.adv_visible, command=lambda: self._toggle_advanced())
        adv_btn.grid(row=row, column=0, columnspan=3, sticky="w", pady=(10, 2))
        row += 1

        self.adv_frame = ttk.Frame(body)
        self.adv_frame.grid(row=row, column=0, columnspan=3, sticky="ew")
        self.adv_frame.grid_remove()
        row += 1

        self.font_size_var = tk.StringVar()
        self.margin_x_var = tk.StringVar()
        self.margin_y_var = tk.StringVar()
        ttk.Label(self.adv_frame, text="Cỡ chữ (mặc định 15):").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(self.adv_frame, textvariable=self.font_size_var, width=10).grid(row=0, column=1, sticky="w", padx=8)
        ttk.Label(self.adv_frame, text="Cách mép trái - point (mặc định 36):").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(self.adv_frame, textvariable=self.margin_x_var, width=10).grid(row=1, column=1, sticky="w", padx=8)
        ttk.Label(self.adv_frame, text="Cách mép trên - point (mặc định 12):").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Entry(self.adv_frame, textvariable=self.margin_y_var, width=10).grid(row=2, column=1, sticky="w", padx=8)

        ttk.Separator(body).grid(row=row, column=0, columnspan=3, sticky="ew", pady=14)
        row += 1

        # ---- Nút chạy + tiến độ ----
        run_frame = tk.Frame(body, bg=WHITE)
        run_frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(4, 10))
        run_frame.columnconfigure(0, weight=1)

        self.run_btn = tk.Button(
            run_frame, text="▶  Bắt đầu đánh dấu", command=self._start_processing,
            bg=BLUE, fg=WHITE, activebackground=BLUE_DARK, activeforeground=WHITE,
            font=("Segoe UI", 11, "bold"), relief="flat", padx=18, pady=10, cursor="hand2",
        )
        self.run_btn.grid(row=0, column=0, sticky="w")

        self.status_var = tk.StringVar(value="Sẵn sàng")
        ttk.Label(run_frame, textvariable=self.status_var, style="Muted.TLabel").grid(
            row=0, column=1, sticky="e", padx=10)
        row += 1

        self.progress = ttk.Progressbar(body, orient="horizontal", mode="determinate")
        self.progress.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        row += 1

        # ---- Nhật ký ----
        ttk.Label(body, text="Nhật ký xử lý", style="Section.TLabel").grid(
            row=row, column=0, columnspan=3, sticky="w", pady=(4, 6))
        row += 1

        self.log_box = scrolledtext.ScrolledText(
            body, height=14, bg=GRAY_BG, fg=TEXT_DARK, font=("Consolas", 9),
            relief="flat", borderwidth=1,
        )
        self.log_box.grid(row=row, column=0, columnspan=3, sticky="ew")
        row += 1

    # ------------------------------------------------------------------
    # Sự kiện
    # ------------------------------------------------------------------
    def _toggle_key_visibility(self):
        self.api_key_entry.config(show="" if self.show_key_var.get() else "•")

    def _toggle_advanced(self):
        if self.adv_visible.get():
            self.adv_frame.grid()
        else:
            self.adv_frame.grid_remove()

    def _choose_input(self):
        path = filedialog.askopenfilename(
            title="Chọn file PDF hồ sơ",
            filetypes=[("PDF files", "*.pdf"), ("Tất cả file", "*.*")],
        )
        if path:
            self.input_var.set(path)
            if not self.output_var.get():
                goc, _ext = os.path.splitext(path)
                self.output_var.set(f"{goc}_da_danh_dau.pdf")

    def _choose_output(self):
        path = filedialog.asksaveasfilename(
            title="Chọn nơi lưu file kết quả",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
        )
        if path:
            self.output_var.set(path)

    def _start_processing(self):
        input_pdf = self.input_var.get().strip()
        output_pdf = self.output_var.get().strip()
        api_key = self.api_key_var.get().strip()
        model = self.model_var.get().strip() or spm.DEFAULT_MODEL
        stamp_mode = self.stamp_mode_var.get()

        if not input_pdf or not os.path.isfile(input_pdf):
            messagebox.showerror("Lỗi", "Vui lòng chọn file PDF đầu vào hợp lệ.")
            return
        if not api_key:
            messagebox.showerror("Lỗi", "Vui lòng nhập Mistral API key.")
            return
        try:
            delay = float(self.delay_var.get())
        except ValueError:
            messagebox.showerror("Lỗi", "Độ trễ (delay) phải là số.")
            return

        if not output_pdf:
            goc, _ext = os.path.splitext(input_pdf)
            output_pdf = f"{goc}_da_danh_dau.pdf"
            self.output_var.set(output_pdf)

        try:
            if self.font_size_var.get().strip():
                spm.STAMP_FONTSIZE = float(self.font_size_var.get())
            if self.margin_x_var.get().strip():
                spm.STAMP_MARGIN_X = float(self.margin_x_var.get())
            if self.margin_y_var.get().strip():
                spm.STAMP_MARGIN_Y = float(self.margin_y_var.get())
        except ValueError:
            messagebox.showerror("Lỗi", "Cỡ chữ / khoảng cách phải là số.")
            return

        self.total_pages = 0
        if fitz is not None:
            try:
                doc = fitz.open(input_pdf)
                self.total_pages = len(doc)
                doc.close()
            except Exception:
                self.total_pages = 0

        self.progress.config(maximum=max(self.total_pages, 1), value=0)
        self.log_box.delete("1.0", tk.END)
        self.run_btn.config(state="disabled", bg=TEXT_MUTED)
        self.status_var.set("Đang xử lý...")

        self.worker_thread = threading.Thread(
            target=self._run_stamp,
            args=(input_pdf, output_pdf, api_key, model, delay, stamp_mode),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_stamp(self, input_pdf, output_pdf, api_key, model, delay, stamp_mode):
        old_stdout = sys.stdout
        sys.stdout = StdoutRedirector(self.log_queue)
        try:
            spm.stamp_pdf(input_pdf, output_pdf, api_key, model, delay, stamp_mode=stamp_mode)
            self.log_queue.put(("done_ok", output_pdf))
        except SystemExit as e:
            self.log_queue.put(("log", f"\n[LỖI] {e}\n"))
            self.log_queue.put(("done_err", str(e)))
        except Exception as e:  # noqa: BLE001
            self.log_queue.put(("log", f"\n[LỖI] {e}\n"))
            self.log_queue.put(("done_err", str(e)))
        finally:
            sys.stdout = old_stdout

    def _poll_log_queue(self):
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()
                if kind == "log":
                    self.log_box.insert(tk.END, payload)
                    self.log_box.see(tk.END)
                    m = re.search(r"Đang đọc trang (\d+)/(\d+)", payload)
                    if m:
                        self.progress["value"] = int(m.group(1))
                elif kind == "done_ok":
                    self.run_btn.config(state="normal", bg=BLUE)
                    self.status_var.set("Hoàn tất ✅")
                    self.progress["value"] = self.progress["maximum"]
                    messagebox.showinfo("Hoàn tất", f"Đã đánh dấu xong file:\n{payload}")
                elif kind == "done_err":
                    self.run_btn.config(state="normal", bg=BLUE)
                    self.status_var.set("Có lỗi xảy ra")
                    messagebox.showerror("Lỗi", payload)
        except queue.Empty:
            pass
        self.after(120, self._poll_log_queue)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
