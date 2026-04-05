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
#  Helpers
# ─────────────────────────────────────────────

PHOTO_EXTS = [".jpg", ".jpeg", ".png", ".JPG", ".JPEG", ".PNG"]


def _build_photo_map(photo_dir):
    mapping = {}
    if not photo_dir or not os.path.isdir(photo_dir):
        return mapping
    for fname in sorted(os.listdir(photo_dir)):
        if not any(fname.lower().endswith(e) for e in (".jpg", ".jpeg", ".png")):
            continue
        m = re.search(r'\((\d+)\)', fname) or re.search(r'(\d+)', fname)
        if m:
            n = int(m.group(1))
            if n not in mapping:
                mapping[n] = os.path.join(photo_dir, fname)
    return mapping


def _find_photo(photo_dir, index_1based, name, photo_map):
    if index_1based in photo_map:
        return photo_map[index_1based]
    for stem in [str(index_1based), name,
                 re.sub(r"\s+", "", name), re.sub(r"\s+", "_", name)]:
        for ext in PHOTO_EXTS:
            p = os.path.join(photo_dir, stem + ext)
            if os.path.isfile(p):
                return p
    return None


def _paste_photo(base_im, photo_path, x, y, w, h):
    photo = Image.open(photo_path).convert("RGBA")
    photo = photo.resize((w, h), Image.LANCZOS)
    bg = Image.new("RGB", (w, h), (255, 255, 255))
    bg.paste(photo, mask=photo.split()[3])
    base_im.paste(bg, (x, y))


def _make_font(font_file, size):
    if font_file and os.path.isfile(font_file):
        return ImageFont.truetype(font_file, size)
    return ImageFont.load_default()


# ─────────────────────────────────────────────
#  Core processing
# ─────────────────────────────────────────────

def process_images(cfg, log, progress_cb, done_cb):
    try:
        wb = openpyxl.load_workbook(cfg["excel_file"], read_only=True, data_only=True)
        ws = wb["Sheet1"]

        arr = []
        for index, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
            if 0 <= index < 48:
                obj = {
                    "name":      re.sub(r" +", " ", str(row[0]) if row[0] else ""),
                    "imageName": re.sub(r" *", "", str(row[1])
                                        .replace("x2", "").replace(" .", "")
                                        .replace("(", "").replace(")", "").replace("X2 ", "")),
                    "index":     index + 1,
                    "row":       row,
                }
                arr.append(obj)

        os.makedirs(cfg["output_folder"], exist_ok=True)
        im1 = Image.open(cfg["base_image"]).convert("RGB")
        total = len(arr)

        use_photo = (cfg.get("use_photo") and cfg.get("photo_dir")
                     and os.path.isdir(cfg["photo_dir"])
                     and cfg.get("photo_w", 0) > 0 and cfg.get("photo_h", 0) > 0)
        photo_map = _build_photo_map(cfg.get("photo_dir", "")) if use_photo else {}

        for i, e in enumerate(arr):
            back_im = im1.copy()

            if use_photo:
                path = _find_photo(cfg["photo_dir"], e["index"], e["name"], photo_map)
                if path:
                    _paste_photo(back_im, path,
                                 cfg["photo_x"], cfg["photo_y"],
                                 cfg["photo_w"], cfg["photo_h"])
                else:
                    log(f"  ⚠ Không tìm thấy ảnh: {e['index']} / {e['name']}")

            draw = ImageDraw.Draw(back_im)
            is_long = len(e["name"]) > cfg["tong_chu_dai"]
            font = _make_font(cfg.get("font_file"),
                              cfg["size_chu_dai"] if is_long else cfg["size_chu"])
            x = cfg["chu_trai_dai"] if is_long else cfg["chu_trai"]
            y = cfg["chu_tren_dai"] if is_long else cfg["chu_tren"]
            name_text = e["name"].title()
            if e["name"].endswith("y"):
                name_text += " "
            draw.text((x, y), name_text, cfg["text_color"], font=font)

            for field in cfg.get("extra_fields", []):
                if not field["enabled"]:
                    continue
                col_idx = field["col"] - 1
                row_data = e["row"]
                if col_idx < 0 or col_idx >= len(row_data) or row_data[col_idx] is None:
                    continue
                text = str(row_data[col_idx]).strip()
                if text:
                    draw.text((field["x"], field["y"]), text, cfg["text_color"],
                              font=_make_font(cfg.get("font_file"), field["size"]))

            back_im.save(os.path.join(cfg["output_folder"], e["imageName"] + ".jpg"),
                         quality=100)
            log(f"[{i+1}/{total}] Saved: {e['imageName']}.jpg")
            progress_cb((i + 1) / total * 100)

        done_cb(True, f"Hoàn thành! {total} ảnh đã lưu tại:\n{cfg['output_folder']}")
    except Exception as exc:
        done_cb(False, str(exc))


# ─────────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────────

PREVIEW_W, PREVIEW_H = 460, 300
DRAG_RADIUS = 14   # px — how close to a handle to start dragging


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tạo Ảnh Hàng Loạt")
        self.resizable(False, False)
        self._preview_job = None
        self._tk_img = None
        # drag state
        self._dragging = False
        self._drag_handle = None   # {"x_var", "y_var", "ox", "oy"} — ox/oy = drag offset
        self._preview_scale = 1.0
        self._preview_ox = 0       # canvas x-offset for centered image
        self._preview_oy = 0
        self._handles = []         # list of handle dicts built each frame
        self._build_ui()
        self._schedule_preview()

    # ── helpers ───────────────────────────────
    def _lbl(self, parent, text, row, col, **kw):
        tk.Label(parent, text=text, anchor="w", **kw).grid(
            row=row, column=col, sticky="w", padx=4, pady=2)

    def _entry(self, parent, var, row, col, width=38):
        e = ttk.Entry(parent, textvariable=var, width=width)
        e.grid(row=row, column=col, sticky="w", padx=4, pady=2)
        return e

    def _spin(self, parent, var, row, col, from_=0, to=9999, width=8):
        s = ttk.Spinbox(parent, textvariable=var, from_=from_, to=to, width=width)
        s.grid(row=row, column=col, sticky="w", padx=4, pady=2)
        return s

    def _browse_file(self, var, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            var.set(path)

    def _browse_folder(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def _trace(self, *vars_):
        for v in vars_:
            v.trace_add("write", lambda *_: self._schedule_preview())

    # ── build ─────────────────────────────────
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # ── Files ─────────────────────────────
        f_files = ttk.LabelFrame(self, text="Tệp tin", padding=8)
        f_files.grid(row=0, column=0, columnspan=2, sticky="ew", **pad)

        self.v_excel    = tk.StringVar()
        self.v_base     = tk.StringVar()
        self.v_font     = tk.StringVar()
        self.v_outdir   = tk.StringVar()
        self.v_photodir = tk.StringVar()

        for i, (lbl, var, ft) in enumerate([
            ("File Excel (.xlsx):", self.v_excel,   [("Excel", "*.xlsx *.xls")]),
            ("Ảnh nền (.jpg):",     self.v_base,    [("Hình ảnh", "*.jpg *.jpeg *.png")]),
            ("File font (.ttf):",   self.v_font,    [("Font", "*.ttf *.otf")]),
        ]):
            self._lbl(f_files, lbl, i, 0)
            self._entry(f_files, var, i, 1)
            ttk.Button(f_files, text="Chọn…",
                       command=lambda v=var, f=ft: self._browse_file(v, f)
                       ).grid(row=i, column=2, padx=4, pady=2)

        self._lbl(f_files, "Thư mục xuất:", 3, 0)
        self._entry(f_files, self.v_outdir, 3, 1)
        ttk.Button(f_files, text="Chọn…",
                   command=lambda: self._browse_folder(self.v_outdir)
                   ).grid(row=3, column=2, padx=4, pady=2)

        self._lbl(f_files, "Thư mục ảnh HS:", 4, 0)
        self._entry(f_files, self.v_photodir, 4, 1)
        ttk.Button(f_files, text="Chọn…",
                   command=lambda: self._browse_folder(self.v_photodir)
                   ).grid(row=4, column=2, padx=4, pady=2)

        # ── Left column ────────────────────────
        left = ttk.Frame(self)
        left.grid(row=1, column=0, sticky="nsew", **pad)

        # ── Text position ──────────────────────
        f_pos = ttk.LabelFrame(left, text="Vị Trí & Cỡ Chữ", padding=8)
        f_pos.pack(fill="x")

        self.v_chu_trai     = tk.IntVar(value=900)
        self.v_chu_tren     = tk.IntVar(value=785)
        self.v_chu_trai_dai = tk.IntVar(value=850)
        self.v_chu_tren_dai = tk.IntVar(value=780)
        self.v_size_chu     = tk.IntVar(value=75)
        self.v_size_chu_dai = tk.IntVar(value=70)
        self.v_tong_chu_dai = tk.IntVar(value=19)

        for i, (lbl, var) in enumerate([
            ("Lề trái (thường):",  self.v_chu_trai),
            ("Lề trên (thường):",  self.v_chu_tren),
            ("Lề trái (tên dài):", self.v_chu_trai_dai),
            ("Lề trên (tên dài):", self.v_chu_tren_dai),
            ("Cỡ chữ (thường):",   self.v_size_chu),
            ("Cỡ chữ (tên dài):",  self.v_size_chu_dai),
            ("Ngưỡng tên dài:",    self.v_tong_chu_dai),
        ]):
            self._lbl(f_pos, lbl, i, 0)
            self._spin(f_pos, var, i, 1)

        self.v_text_color = tk.StringVar(value="#000000")
        self._lbl(f_pos, "Màu chữ:", 7, 0)
        cr = ttk.Frame(f_pos)
        cr.grid(row=7, column=1, sticky="w", padx=4, pady=2)
        self.color_preview = tk.Label(cr, bg="#000000", width=4, relief="solid")
        self.color_preview.pack(side="left")
        ttk.Button(cr, text="Chọn màu…", command=self._pick_color).pack(side="left", padx=6)
        self.v_text_color.trace_add("write", lambda *_: self._schedule_preview())

        # ── Photo settings ─────────────────────
        f_photo = ttk.LabelFrame(left, text="Ảnh Học Sinh", padding=8)
        f_photo.pack(fill="x", pady=(8, 0))

        self.v_use_photo = tk.BooleanVar(value=False)
        ttk.Checkbutton(f_photo, text="Bật tính năng ảnh học sinh",
                        variable=self.v_use_photo,
                        command=self._schedule_preview
                        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=4, pady=(0, 4))

        self.v_photo_x = tk.IntVar(value=100)
        self.v_photo_y = tk.IntVar(value=400)
        self.v_photo_w = tk.IntVar(value=260)
        self.v_photo_h = tk.IntVar(value=320)

        for i, (lbl, var) in enumerate([
            ("Lề trái (X):", self.v_photo_x),
            ("Lề trên (Y):", self.v_photo_y),
            ("Chiều rộng:",  self.v_photo_w),
            ("Chiều cao:",   self.v_photo_h),
        ], start=1):
            self._lbl(f_photo, lbl, i, 0)
            self._spin(f_photo, var, i, 1)

        self.v_sample_photo = tk.StringVar()
        self._lbl(f_photo, "Ảnh mẫu (preview):", 5, 0)
        self._entry(f_photo, self.v_sample_photo, 5, 1, width=20)
        ttk.Button(f_photo, text="Chọn…",
                   command=lambda: self._browse_file(self.v_sample_photo,
                       [("Hình ảnh", "*.jpg *.jpeg *.png")])
                   ).grid(row=5, column=2, padx=4, pady=2)

        # ── Preview ────────────────────────────
        f_prev = ttk.LabelFrame(self, text="Xem Trước  (kéo thả để đổi vị trí)", padding=8)
        f_prev.grid(row=1, column=1, sticky="nsew", **pad)

        self.v_preview_name = tk.StringVar(value="Nguyen Van A")
        nr = ttk.Frame(f_prev)
        nr.pack(fill="x", pady=(0, 6))
        tk.Label(nr, text="Tên mẫu:").pack(side="left")
        ttk.Entry(nr, textvariable=self.v_preview_name, width=24).pack(side="left", padx=6)

        self.preview_canvas = tk.Canvas(f_prev, width=PREVIEW_W, height=PREVIEW_H,
                                        bg="#2b2b2b", highlightthickness=1,
                                        highlightbackground="#555", cursor="fleur")
        self.preview_canvas.pack()
        self.preview_canvas.create_text(PREVIEW_W // 2, PREVIEW_H // 2,
                                        text="Chọn ảnh nền để xem trước",
                                        fill="#888", justify="center", font=("Arial", 12))
        self.preview_canvas.bind("<Button-1>",       self._drag_start)
        self.preview_canvas.bind("<B1-Motion>",      self._drag_move)
        self.preview_canvas.bind("<ButtonRelease-1>", self._drag_end)

        # ── Extra fields ───────────────────────
        f_extra = ttk.LabelFrame(self, text="Thông Tin Bổ Sung", padding=8)
        f_extra.grid(row=2, column=0, columnspan=2, sticky="ew", **pad)

        for col, txt in enumerate(["", "Trường", "Cột Excel", "X", "Y", "Cỡ chữ", "Mẫu (preview)"]):
            tk.Label(f_extra, text=txt, font=("", 9, "bold"), anchor="w"
                     ).grid(row=0, column=col, padx=3)
        ttk.Separator(f_extra, orient="horizontal").grid(
            row=1, column=0, columnspan=7, sticky="ew", pady=2)

        self.extra_fields = []
        for i, (label, def_col, def_x, def_y, def_sz, sample) in enumerate([
            ("Ngày sinh", 3, 550, 600, 50, "01/01/2014"),
            ("Giới tính", 4, 950, 600, 50, "Nam"),
            ("Địa chỉ",   5, 550, 670, 50, "Đà Nẵng"),
        ]):
            row = i + 2
            ef = {
                "label":   label,
                "enabled": tk.BooleanVar(value=False),
                "col":     tk.IntVar(value=def_col),
                "x":       tk.IntVar(value=def_x),
                "y":       tk.IntVar(value=def_y),
                "size":    tk.IntVar(value=def_sz),
                "sample":  tk.StringVar(value=sample),
            }
            self.extra_fields.append(ef)
            ttk.Checkbutton(f_extra, variable=ef["enabled"],
                            command=self._schedule_preview
                            ).grid(row=row, column=0, padx=3)
            tk.Label(f_extra, text=label, width=12, anchor="w").grid(row=row, column=1, padx=3)
            ttk.Spinbox(f_extra, textvariable=ef["col"],  from_=1, to=50,   width=6).grid(row=row, column=2, padx=3)
            ttk.Spinbox(f_extra, textvariable=ef["x"],    from_=0, to=9999, width=6).grid(row=row, column=3, padx=3)
            ttk.Spinbox(f_extra, textvariable=ef["y"],    from_=0, to=9999, width=6).grid(row=row, column=4, padx=3)
            ttk.Spinbox(f_extra, textvariable=ef["size"], from_=6, to=300,  width=6).grid(row=row, column=5, padx=3)
            ttk.Entry(f_extra, textvariable=ef["sample"], width=16).grid(row=row, column=6, padx=3)
            self._trace(ef["col"], ef["x"], ef["y"], ef["size"], ef["sample"])

        # ── Run / progress ─────────────────────
        f_run = ttk.Frame(self, padding=8)
        f_run.grid(row=3, column=0, columnspan=2, sticky="ew", **pad)
        self.btn_run = ttk.Button(f_run, text="▶  Tạo Ảnh", command=self._run)
        self.btn_run.pack(side="left", padx=4)
        self.progress = ttk.Progressbar(f_run, length=340, mode="determinate")
        self.progress.pack(side="left", padx=8)
        self.lbl_status = tk.Label(f_run, text="Sẵn sàng", fg="gray", width=12)
        self.lbl_status.pack(side="left")

        # ── Log ───────────────────────────────
        f_log = ttk.LabelFrame(self, text="Nhật ký", padding=6)
        f_log.grid(row=4, column=0, columnspan=2, sticky="ew", **pad)
        self.log_box = tk.Text(f_log, height=7, width=80, state="disabled",
                               bg="#1e1e1e", fg="#d4d4d4", font=("Courier", 10))
        sb = ttk.Scrollbar(f_log, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=sb.set)
        self.log_box.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        # trace all preview vars
        self._trace(
            self.v_base, self.v_font, self.v_preview_name,
            self.v_chu_trai, self.v_chu_tren,
            self.v_chu_trai_dai, self.v_chu_tren_dai,
            self.v_size_chu, self.v_size_chu_dai, self.v_tong_chu_dai,
            self.v_photo_x, self.v_photo_y, self.v_photo_w, self.v_photo_h,
            self.v_sample_photo, self.v_photodir,
        )

    # ── color picker ──────────────────────────
    def _pick_color(self):
        result = colorchooser.askcolor(color=self.v_text_color.get(), title="Chọn màu chữ")
        if result and result[1]:
            self.v_text_color.set(result[1])
            self.color_preview.config(bg=result[1])

    # ── drag & drop ───────────────────────────
    def _drag_start(self, event):
        best, best_dist = None, DRAG_RADIUS
        for h in self._handles:
            if h["type"] == "point":
                d = ((event.x - h["cx"])**2 + (event.y - h["cy"])**2) ** 0.5
                if d < best_dist:
                    best, best_dist = h, d
            elif h["type"] == "rect":
                rx, ry, rw, rh = h["cx"], h["cy"], h["cw"], h["ch"]
                if rx <= event.x <= rx + rw and ry <= event.y <= ry + rh:
                    # inside the box → drag from click point
                    best = h
                    best_dist = 0
                    break
        if best:
            self._dragging = True
            self._drag_handle = {
                "x_var": best["x_var"],
                "y_var": best["y_var"],
                "ox": event.x - best["cx"],   # offset within handle
                "oy": event.y - best["cy"],
            }

    def _drag_move(self, event):
        if not self._dragging or not self._drag_handle:
            return
        h = self._drag_handle
        s = self._preview_scale
        # convert canvas pos → image pos
        new_cx = event.x - h["ox"]
        new_cy = event.y - h["oy"]
        img_x = round((new_cx - self._preview_ox) / s)
        img_y = round((new_cy - self._preview_oy) / s)
        h["x_var"].set(max(0, img_x))
        h["y_var"].set(max(0, img_y))
        self._update_preview()   # immediate, no debounce

    def _drag_end(self, event):
        self._dragging = False
        self._drag_handle = None

    # ── preview ───────────────────────────────
    def _schedule_preview(self):
        if self._dragging:
            return
        if self._preview_job:
            self.after_cancel(self._preview_job)
        self._preview_job = self.after(250, self._update_preview)

    def _update_preview(self):
        self._preview_job = None
        base = self.v_base.get()
        if not base or not os.path.isfile(base):
            return
        try:
            img = Image.open(base).convert("RGB")
            iw, ih = img.size
            scale = min(PREVIEW_W / iw, PREVIEW_H / ih)
            new_w, new_h = int(iw * scale), int(ih * scale)
            img = img.resize((new_w, new_h), Image.LANCZOS)
            draw = ImageDraw.Draw(img)
            font_path = self.v_font.get()
            color = self.v_text_color.get()

            ox = (PREVIEW_W - new_w) // 2
            oy = (PREVIEW_H - new_h) // 2
            self._preview_scale = scale
            self._preview_ox = ox
            self._preview_oy = oy

            handles = []

            # ── photo ────────────────────────────
            if self.v_use_photo.get():
                pw = self.v_photo_w.get()
                ph_h = self.v_photo_h.get()
                spw = max(4, int(pw * scale))
                sph = max(4, int(ph_h * scale))
                px = int(self.v_photo_x.get() * scale)
                py = int(self.v_photo_y.get() * scale)
                sample = self.v_sample_photo.get()
                if sample and os.path.isfile(sample):
                    ph_img = Image.open(sample).convert("RGBA")
                    ph_img = ph_img.resize((spw, sph), Image.LANCZOS)
                    bg = Image.new("RGB", (spw, sph), (255, 255, 255))
                    bg.paste(ph_img, mask=ph_img.split()[3])
                    img.paste(bg, (px, py))
                else:
                    draw.rectangle([px, py, px + spw, py + sph], outline="#ff6600", width=2)
                    draw.line([px, py, px + spw, py + sph], fill="#ff6600", width=1)
                    draw.line([px + spw, py, px, py + sph], fill="#ff6600", width=1)
                handles.append({
                    "type": "rect",
                    "x_var": self.v_photo_x, "y_var": self.v_photo_y,
                    "cx": ox + px, "cy": oy + py, "cw": spw, "ch": sph,
                })

            # ── name text ────────────────────────
            name = self.v_preview_name.get().strip() or "Tên Mẫu"
            is_long = len(name) > self.v_tong_chu_dai.get()
            x_var = self.v_chu_trai_dai if is_long else self.v_chu_trai
            y_var = self.v_chu_tren_dai if is_long else self.v_chu_tren
            nsz = self.v_size_chu_dai.get() if is_long else self.v_size_chu.get()
            tx = int(x_var.get() * scale)
            ty = int(y_var.get() * scale)
            font = _make_font(font_path, max(8, int(nsz * scale)))
            draw.text((tx, ty), name.title(), color, font=font)
            draw.line([(tx-7, ty), (tx+7, ty)], fill="red", width=2)
            draw.line([(tx, ty-7), (tx, ty+7)], fill="red", width=2)
            handles.append({
                "type": "point",
                "x_var": x_var, "y_var": y_var,
                "cx": ox + tx, "cy": oy + ty,
            })

            # ── extra fields ─────────────────────
            for ef in self.extra_fields:
                if not ef["enabled"].get():
                    continue
                text = ef["sample"].get().strip()
                ex = int(ef["x"].get() * scale)
                ey = int(ef["y"].get() * scale)
                esz = max(6, int(ef["size"].get() * scale))
                efont = _make_font(font_path, esz)
                if text:
                    draw.text((ex, ey), text, color, font=efont)
                draw.line([(ex-6, ey), (ex+6, ey)], fill="#00aaff", width=2)
                draw.line([(ex, ey-6), (ex, ey+6)], fill="#00aaff", width=2)
                handles.append({
                    "type": "point",
                    "x_var": ef["x"], "y_var": ef["y"],
                    "cx": ox + ex, "cy": oy + ey,
                })

            self._handles = handles
            self._tk_img = ImageTk.PhotoImage(img)
            self.preview_canvas.delete("all")
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
            "use_photo":     self.v_use_photo.get(),
            "photo_dir":     self.v_photodir.get(),
            "photo_x":       self.v_photo_x.get(),
            "photo_y":       self.v_photo_y.get(),
            "photo_w":       self.v_photo_w.get(),
            "photo_h":       self.v_photo_h.get(),
            "extra_fields": [
                {"enabled": ef["enabled"].get(), "col": ef["col"].get(),
                 "x": ef["x"].get(), "y": ef["y"].get(), "size": ef["size"].get()}
                for ef in self.extra_fields
            ],
        }

        self.btn_run.config(state="disabled")
        self.progress["value"] = 0
        self.lbl_status.config(text="Đang chạy…", fg="blue")
        self._log(f"\n── Bắt đầu: {cfg['excel_file']}")

        threading.Thread(target=process_images,
                         args=(cfg, self._log, self._set_progress, self._done),
                         daemon=True).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()
