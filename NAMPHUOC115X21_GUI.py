#!/usr/bin/python
# -*- coding: utf8 -*-

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import threading
import os
import re
from PIL import Image, ImageDraw, ImageFont, ImageTk
import openpyxl


# ─────────────────────────────────────────────
#  Core processing function
# ─────────────────────────────────────────────

def process_images(cfg, log, progress_cb, done_cb):
    try:
        wb = openpyxl.load_workbook(cfg["excel_file"], read_only=True, data_only=True)
        ws = wb["Sheet1"]

        arr = []
        for index, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
            if 0 <= index < 48:
                obj = {}
                obj["name"] = row[0]
                obj["imageName"] = (
                    str(row[1])
                    .replace("x2", "")
                    .replace(" .", "")
                    .replace("(", "")
                    .replace(")", "")
                    .replace("X2 ", "")
                )
                obj["imageName"] = re.sub(" *", "", obj["imageName"])
                obj["name"] = re.sub(" +", " ", obj["name"])
                arr.append(obj)

        os.makedirs(cfg["output_folder"], exist_ok=True)
        im1 = Image.open(cfg["base_image"])
        total = len(arr)

        for i, e in enumerate(arr):
            is_long = len(e["name"]) > cfg["tong_chu_dai"]
            font_size = cfg["size_chu_dai"] if is_long else cfg["size_chu"]
            if cfg.get("font_file") and os.path.isfile(cfg["font_file"]):
                font = ImageFont.truetype(cfg["font_file"], font_size)
            else:
                font = ImageFont.load_default()

            back_im = im1.copy()
            draw = ImageDraw.Draw(back_im)

            x = cfg["chu_trai_dai"] if is_long else cfg["chu_trai"]
            y = cfg["chu_tren_dai"] if is_long else cfg["chu_tren"]
            name_text = e["name"].title()
            if e["name"].endswith("y"):
                name_text += " "

            draw.text((x, y), name_text, cfg["text_color"], font=font)

            out_path = os.path.join(cfg["output_folder"], e["imageName"] + ".jpg")
            back_im.save(out_path, quality=100)

            log(f"[{i+1}/{total}] Saved: {e['imageName']}.jpg")
            progress_cb((i + 1) / total * 100)

        done_cb(True, f"Hoàn thành! {total} ảnh đã lưu tại:\n{cfg['output_folder']}")
    except Exception as exc:
        done_cb(False, str(exc))


# ─────────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────────

PREVIEW_W, PREVIEW_H = 420, 280

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tạo Ảnh Hàng Loạt")
        self.resizable(False, False)
        self._preview_job = None
        self._tk_img = None
        self._build_ui()
        self._schedule_preview()

    # ── helpers ───────────────────────────────
    def _lbl(self, parent, text, row, col, **kw):
        tk.Label(parent, text=text, anchor="w", **kw).grid(
            row=row, column=col, sticky="w", padx=6, pady=3
        )

    def _entry(self, parent, var, row, col, width=38):
        e = ttk.Entry(parent, textvariable=var, width=width)
        e.grid(row=row, column=col, sticky="w", padx=6, pady=3)
        return e

    def _spin(self, parent, var, row, col, from_=0, to=9999):
        s = ttk.Spinbox(parent, textvariable=var, from_=from_, to=to, width=8)
        s.grid(row=row, column=col, sticky="w", padx=6, pady=3)
        return s

    def _browse_file(self, var, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            var.set(path)

    def _browse_folder(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    # ── build ─────────────────────────────────
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # ── File paths ────────────────────────
        f_files = ttk.LabelFrame(self, text="Tệp tin", padding=8)
        f_files.grid(row=0, column=0, columnspan=2, sticky="ew", **pad)

        self.v_excel  = tk.StringVar()
        self.v_base   = tk.StringVar()
        self.v_font   = tk.StringVar()
        self.v_outdir = tk.StringVar()

        rows_files = [
            ("File Excel (.xlsx):", self.v_excel,  [("Excel", "*.xlsx *.xls")]),
            ("Ảnh nền (.jpg):",     self.v_base,   [("Hình ảnh", "*.jpg *.jpeg *.png")]),
            ("File font (.ttf):",   self.v_font,   [("Font", "*.ttf *.otf")]),
        ]
        for i, (lbl, var, ft) in enumerate(rows_files):
            self._lbl(f_files, lbl, i, 0)
            self._entry(f_files, var, i, 1)
            ttk.Button(
                f_files, text="Chọn…",
                command=lambda v=var, f=ft: self._browse_file(v, f)
            ).grid(row=i, column=2, padx=4, pady=3)

        self._lbl(f_files, "Thư mục xuất:", 3, 0)
        self._entry(f_files, self.v_outdir, 3, 1)
        ttk.Button(
            f_files, text="Chọn…",
            command=lambda: self._browse_folder(self.v_outdir)
        ).grid(row=3, column=2, padx=4, pady=3)

        # ── Text position params ───────────────
        f_pos = ttk.LabelFrame(self, text="Vị Trí & Cỡ Chữ", padding=8)
        f_pos.grid(row=1, column=0, sticky="nsew", **pad)

        self.v_chu_trai      = tk.IntVar(value=900)
        self.v_chu_tren      = tk.IntVar(value=785)
        self.v_chu_trai_dai  = tk.IntVar(value=850)
        self.v_chu_tren_dai  = tk.IntVar(value=780)
        self.v_size_chu      = tk.IntVar(value=75)
        self.v_size_chu_dai  = tk.IntVar(value=70)
        self.v_tong_chu_dai  = tk.IntVar(value=19)

        pos_rows = [
            ("Lề trái (thường):",    self.v_chu_trai),
            ("Lề trên (thường):",    self.v_chu_tren),
            ("Lề trái (tên dài):",   self.v_chu_trai_dai),
            ("Lề trên (tên dài):",   self.v_chu_tren_dai),
            ("Cỡ chữ (thường):",     self.v_size_chu),
            ("Cỡ chữ (tên dài):",    self.v_size_chu_dai),
            ("Ngưỡng tên dài:",      self.v_tong_chu_dai),
        ]
        for i, (lbl, var) in enumerate(pos_rows):
            self._lbl(f_pos, lbl, i, 0)
            self._spin(f_pos, var, i, 1)

        # color picker
        self.v_text_color = tk.StringVar(value="#000000")
        self._lbl(f_pos, "Màu chữ:", len(pos_rows), 0)
        color_row = ttk.Frame(f_pos)
        color_row.grid(row=len(pos_rows), column=1, sticky="w", padx=6, pady=3)
        self.color_preview = tk.Label(color_row, bg="#000000", width=4, relief="solid")
        self.color_preview.pack(side="left")
        ttk.Button(color_row, text="Chọn màu…", command=self._pick_color).pack(side="left", padx=6)
        self.v_text_color.trace_add("write", lambda *_: self._schedule_preview())

        # ── Live preview ───────────────────────
        f_prev = ttk.LabelFrame(self, text="Xem Trước", padding=8)
        f_prev.grid(row=1, column=1, sticky="nsew", **pad)

        self.v_preview_name = tk.StringVar(value="Nguyen Van A")
        name_row = ttk.Frame(f_prev)
        name_row.pack(fill="x", pady=(0, 6))
        tk.Label(name_row, text="Tên mẫu:").pack(side="left")
        ttk.Entry(name_row, textvariable=self.v_preview_name, width=24).pack(side="left", padx=6)

        self.preview_canvas = tk.Canvas(f_prev, width=PREVIEW_W, height=PREVIEW_H,
                                        bg="#2b2b2b", highlightthickness=1,
                                        highlightbackground="#555")
        self.preview_canvas.pack()
        self.preview_canvas.create_text(
            PREVIEW_W // 2, PREVIEW_H // 2,
            text="Chọn ảnh nền để xem trước",
            fill="#888", justify="center", font=("Arial", 12)
        )

        # trace all vars that affect the preview
        for v in (self.v_base, self.v_font, self.v_preview_name,
                  self.v_chu_trai, self.v_chu_tren,
                  self.v_chu_trai_dai, self.v_chu_tren_dai,
                  self.v_size_chu, self.v_size_chu_dai, self.v_tong_chu_dai):
            v.trace_add("write", lambda *_: self._schedule_preview())

        # ── Run / progress ─────────────────────
        f_run = ttk.Frame(self, padding=8)
        f_run.grid(row=2, column=0, columnspan=2, sticky="ew", **pad)

        self.btn_run = ttk.Button(f_run, text="▶  Tạo Ảnh", command=self._run)
        self.btn_run.pack(side="left", padx=4)

        self.progress = ttk.Progressbar(f_run, length=340, mode="determinate")
        self.progress.pack(side="left", padx=8)

        self.lbl_status = tk.Label(f_run, text="Sẵn sàng", fg="gray", width=12)
        self.lbl_status.pack(side="left")

        # ── Log ───────────────────────────────
        f_log = ttk.LabelFrame(self, text="Nhật ký", padding=6)
        f_log.grid(row=3, column=0, columnspan=2, sticky="ew", **pad)

        self.log_box = tk.Text(f_log, height=10, width=72, state="disabled",
                               bg="#1e1e1e", fg="#d4d4d4", font=("Courier", 10))
        sb = ttk.Scrollbar(f_log, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=sb.set)
        self.log_box.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

    def _pick_color(self):
        result = colorchooser.askcolor(color=self.v_text_color.get(), title="Chọn màu chữ")
        if result and result[1]:
            self.v_text_color.set(result[1])
            self.color_preview.config(bg=result[1])

    # ── preview ───────────────────────────────
    def _schedule_preview(self):
        if self._preview_job:
            self.after_cancel(self._preview_job)
        self._preview_job = self.after(250, self._update_preview)

    def _update_preview(self):
        self._preview_job = None
        base = self.v_base.get()
        font_path = self.v_font.get()
        if not base or not os.path.isfile(base):
            return
        try:
            name = self.v_preview_name.get().strip() or "Sample Name"
            is_long = len(name) > self.v_tong_chu_dai.get()
            x = self.v_chu_trai_dai.get() if is_long else self.v_chu_trai.get()
            y = self.v_chu_tren_dai.get() if is_long else self.v_chu_tren.get()
            size = self.v_size_chu_dai.get() if is_long else self.v_size_chu.get()

            img = Image.open(base).convert("RGB")
            iw, ih = img.size

            # scale image first so text is drawn at preview resolution
            scale = min(PREVIEW_W / iw, PREVIEW_H / ih)
            new_w, new_h = int(iw * scale), int(ih * scale)
            img = img.resize((new_w, new_h), Image.LANCZOS)

            px, py = int(x * scale), int(y * scale)
            scaled_size = max(8, int(size * scale))

            draw = ImageDraw.Draw(img)
            if font_path and os.path.isfile(font_path):
                font = ImageFont.truetype(font_path, scaled_size)
            else:
                font = ImageFont.load_default()

            draw.text((px, py), name.title(), self.v_text_color.get(), font=font)

            # crosshair marker
            draw.line([(px - 6, py), (px + 6, py)], fill="red", width=1)
            draw.line([(px, py - 6), (px, py + 6)], fill="red", width=1)

            self._tk_img = ImageTk.PhotoImage(img)
            self.preview_canvas.delete("all")
            ox = (PREVIEW_W - new_w) // 2
            oy = (PREVIEW_H - new_h) // 2
            self.preview_canvas.create_image(ox, oy, anchor="nw", image=self._tk_img)
        except Exception:
            pass

    # ── actions ───────────────────────────────
    def _log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _set_progress(self, pct):
        self.progress["value"] = pct
        self.lbl_status.config(text=f"{int(pct)}%", fg="blue")
        self.update_idletasks()

    def _done(self, ok, msg):
        self.btn_run.config(state="normal")
        if ok:
            self.lbl_status.config(text="Xong!", fg="green")
            self._log("✓ " + msg)
            messagebox.showinfo("Hoàn thành", msg)
        else:
            self.lbl_status.config(text="Lỗi", fg="red")
            self._log("✗ Lỗi: " + msg)
            messagebox.showerror("Lỗi", msg)

    def _run(self):
        # Validate required fields
        missing = []
        if not self.v_excel.get():  missing.append("File Excel")
        if not self.v_base.get():   missing.append("Ảnh nền")
        if not self.v_outdir.get(): missing.append("Thư mục xuất")
        if missing:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng điền:\n• " + "\n• ".join(missing))
            return

        cfg = {
            "excel_file":    self.v_excel.get(),
            "base_image":    self.v_base.get(),
            "font_file":     self.v_font.get(),
            "output_folder": self.v_outdir.get(),
            "chu_trai":      self.v_chu_trai.get(),
            "chu_tren":      self.v_chu_tren.get(),
            "chu_trai_dai":  self.v_chu_trai_dai.get(),
            "chu_tren_dai":  self.v_chu_tren_dai.get(),
            "size_chu":      self.v_size_chu.get(),
            "size_chu_dai":  self.v_size_chu_dai.get(),
            "tong_chu_dai":  self.v_tong_chu_dai.get(),
            "text_color":    self.v_text_color.get(),
        }

        self.btn_run.config(state="disabled")
        self.progress["value"] = 0
        self.lbl_status.config(text="Đang chạy…", fg="blue")
        self._log(f"\n── Bắt đầu: {cfg['excel_file']}")

        threading.Thread(
            target=process_images,
            args=(cfg, self._log, self._set_progress, self._done),
            daemon=True,
        ).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()
