"""
Orientation Inspector
---------------------
Windows GUI tool that reports whether an image's rotation is baked into
the pixel data or only declared via the EXIF Orientation flag (tag 274).

Requirements:
    pip install pillow

Usage:
    python orientation_inspector.py
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageOps, ImageTk

EXIF_ORIENTATION_TAG = 274

ORIENTATION_MEANINGS = {
    1: "Normal (no rotation)",
    2: "Mirrored horizontal",
    3: "Rotated 180°",
    4: "Mirrored vertical",
    5: "Mirrored horizontal, rotated 270° CW",
    6: "Rotated 90° CW",
    7: "Mirrored horizontal, rotated 90° CW",
    8: "Rotated 270° CW (90° CCW)",
}

SUPPORTED_EXTENSIONS = {
    ".jpg", ".jpeg", ".jpe", ".jfif",
    ".tif", ".tiff",
    ".png", ".webp", ".heic", ".heif",
}

PREVIEW_SIZE = 240


def inspect_image(path):
    """Return a dict describing the orientation state of one image."""
    info = {
        "path": path,
        "name": os.path.basename(path),
        "format": "?",
        "size": "?",
        "orientation": None,
        "meaning": "-",
        "verdict": "",
        "flag_only": False,
        "error": None,
    }
    try:
        with Image.open(path) as img:
            info["format"] = img.format or "?"
            info["size"] = f"{img.width} x {img.height}"
            try:
                exif = img.getexif()
                orientation = exif.get(EXIF_ORIENTATION_TAG)
            except Exception:
                orientation = None

            info["orientation"] = orientation

            if orientation is None:
                info["meaning"] = "No flag present"
                info["verdict"] = ("Pixels as stored — no orientation flag. "
                                   "Any rotation is baked in.")
            elif orientation == 1:
                info["meaning"] = ORIENTATION_MEANINGS[1]
                info["verdict"] = ("Flag = 1 (normal). Pixels are upright as "
                                   "stored; any rotation is baked in.")
            elif orientation in ORIENTATION_MEANINGS:
                info["meaning"] = ORIENTATION_MEANINGS[orientation]
                info["flag_only"] = True
                info["verdict"] = (f"FLAG-ONLY rotation (tag {orientation}: "
                                   f"{ORIENTATION_MEANINGS[orientation]}). "
                                   "Pixels are stored UN-rotated; viewers "
                                   "rotate on display. RIPs or apps that "
                                   "ignore EXIF will show it sideways.")
            else:
                info["meaning"] = f"Unknown value ({orientation})"
                info["verdict"] = ("Non-standard orientation value — treat "
                                   "with caution.")
    except Exception as exc:
        info["error"] = str(exc)
        info["verdict"] = f"Could not read file: {exc}"
    return info


def bake_and_strip(path):
    """Apply the EXIF orientation to the pixels, remove the flag, and save
    a copy alongside the original. Returns the new path."""
    with Image.open(path) as img:
        icc = img.info.get("icc_profile")
        dpi = img.info.get("dpi")
        fmt = img.format

        transposed = ImageOps.exif_transpose(img)
        if transposed is None:  # older Pillow returns None when no change
            transposed = img.copy()

        exif = transposed.getexif()
        if EXIF_ORIENTATION_TAG in exif:
            del exif[EXIF_ORIENTATION_TAG]

        root, ext = os.path.splitext(path)
        out_path = f"{root}_baked{ext}"

        save_kwargs = {"exif": exif.tobytes()}
        if icc:
            save_kwargs["icc_profile"] = icc
        if dpi:
            save_kwargs["dpi"] = dpi
        if fmt == "JPEG":
            # 'keep' isn't valid on a transposed copy; use high quality
            # with 4:4:4 subsampling to minimise recompression loss.
            save_kwargs.update(quality=95, subsampling=0)

        transposed.save(out_path, format=fmt, **save_kwargs)
    return out_path


class OrientationInspector(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Orientation Inspector")
        self.geometry("980x680")
        self.minsize(820, 560)

        self.records = {}          # tree item id -> info dict
        self._photo_refs = []      # keep PhotoImage references alive

        self._build_ui()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        toolbar = ttk.Frame(self, padding=(8, 8, 8, 4))
        toolbar.pack(fill="x")

        ttk.Button(toolbar, text="Add files...",
                   command=self.add_files).pack(side="left")
        ttk.Button(toolbar, text="Add folder...",
                   command=self.add_folder).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="Clear list",
                   command=self.clear_list).pack(side="left", padx=(6, 0))
        self.bake_btn = ttk.Button(
            toolbar, text="Bake rotation + strip flag (save copy)",
            command=self.bake_selected, state="disabled")
        self.bake_btn.pack(side="right")

        # --- results table -------------------------------------------------
        table_frame = ttk.Frame(self, padding=(8, 4))
        table_frame.pack(fill="both", expand=True)

        columns = ("format", "size", "flag", "meaning", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns,
                                 show="tree headings", selectmode="browse")
        self.tree.heading("#0", text="File")
        self.tree.heading("format", text="Format")
        self.tree.heading("size", text="Stored pixels (W x H)")
        self.tree.heading("flag", text="EXIF flag")
        self.tree.heading("meaning", text="Flag meaning")
        self.tree.heading("status", text="Status")

        self.tree.column("#0", width=260, anchor="w")
        self.tree.column("format", width=70, anchor="center")
        self.tree.column("size", width=140, anchor="center")
        self.tree.column("flag", width=70, anchor="center")
        self.tree.column("meaning", width=220, anchor="w")
        self.tree.column("status", width=140, anchor="center")

        vsb = ttk.Scrollbar(table_frame, orient="vertical",
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.tag_configure("flag_only", background="#fff3cd")
        self.tree.tag_configure("error", background="#f8d7da")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        # --- detail / preview pane ------------------------------------------
        detail = ttk.LabelFrame(self, text="Details", padding=8)
        detail.pack(fill="x", padx=8, pady=(0, 8))

        self.verdict_var = tk.StringVar(
            value="Add images to inspect their orientation state.")
        verdict_label = ttk.Label(detail, textvariable=self.verdict_var,
                                  wraplength=920, justify="left")
        verdict_label.pack(anchor="w", pady=(0, 8))

        previews = ttk.Frame(detail)
        previews.pack(anchor="w")

        left = ttk.Frame(previews)
        left.pack(side="left", padx=(0, 24))
        ttk.Label(left, text="Stored pixel data (flag ignored)").pack()
        self.raw_canvas = tk.Label(left, relief="groove", width=34, height=16)
        self.raw_canvas.pack(pady=4)

        right = ttk.Frame(previews)
        right.pack(side="left")
        ttk.Label(right, text="As displayed (flag applied)").pack()
        self.display_canvas = tk.Label(right, relief="groove",
                                       width=34, height=16)
        self.display_canvas.pack(pady=4)

    # ------------------------------------------------------------ actions
    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select images",
            filetypes=[
                ("Images", "*.jpg *.jpeg *.jpe *.jfif *.tif *.tiff "
                           "*.png *.webp *.heic *.heif"),
                ("All files", "*.*"),
            ])
        self._ingest(paths)

    def add_folder(self):
        folder = filedialog.askdirectory(title="Select folder")
        if not folder:
            return
        paths = []
        for entry in sorted(os.listdir(folder)):
            full = os.path.join(folder, entry)
            if (os.path.isfile(full)
                    and os.path.splitext(entry)[1].lower()
                    in SUPPORTED_EXTENSIONS):
                paths.append(full)
        if not paths:
            messagebox.showinfo("No images",
                                "No supported image files in that folder.")
            return
        self._ingest(paths)

    def _ingest(self, paths):
        for path in paths:
            info = inspect_image(path)
            flag = ("-" if info["orientation"] is None
                    else str(info["orientation"]))
            if info["error"]:
                status, tags = "Read error", ("error",)
            elif info["flag_only"]:
                status, tags = "FLAG-ONLY", ("flag_only",)
            else:
                status, tags = "Baked / upright", ()
            item = self.tree.insert(
                "", "end", text=info["name"],
                values=(info["format"], info["size"], flag,
                        info["meaning"], status),
                tags=tags)
            self.records[item] = info

    def clear_list(self):
        self.tree.delete(*self.tree.get_children())
        self.records.clear()
        self._photo_refs.clear()
        self.raw_canvas.configure(image="", text="")
        self.display_canvas.configure(image="", text="")
        self.verdict_var.set("Add images to inspect their orientation state.")
        self.bake_btn.configure(state="disabled")

    def on_select(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        info = self.records.get(sel[0])
        if not info:
            return

        self.verdict_var.set(f"{info['name']}  —  {info['verdict']}")
        self.bake_btn.configure(
            state="normal" if info["flag_only"] else "disabled")

        self._photo_refs.clear()
        if info["error"]:
            self.raw_canvas.configure(image="", text="unreadable")
            self.display_canvas.configure(image="", text="unreadable")
            return

        try:
            with Image.open(info["path"]) as img:
                raw = img.copy()
            raw.thumbnail((PREVIEW_SIZE, PREVIEW_SIZE))
            raw_photo = ImageTk.PhotoImage(raw)

            with Image.open(info["path"]) as img:
                shown = ImageOps.exif_transpose(img)
                shown = (shown or img).copy()
            shown.thumbnail((PREVIEW_SIZE, PREVIEW_SIZE))
            shown_photo = ImageTk.PhotoImage(shown)

            self._photo_refs.extend([raw_photo, shown_photo])
            self.raw_canvas.configure(image=raw_photo, text="",
                                      width=raw_photo.width(),
                                      height=raw_photo.height())
            self.display_canvas.configure(image=shown_photo, text="",
                                          width=shown_photo.width(),
                                          height=shown_photo.height())
        except Exception as exc:
            self.raw_canvas.configure(image="", text=f"preview failed\n{exc}")
            self.display_canvas.configure(image="", text="")

    def bake_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        info = self.records.get(item)
        if not info or not info["flag_only"]:
            return
        try:
            out_path = bake_and_strip(info["path"])
        except Exception as exc:
            messagebox.showerror("Bake failed", str(exc))
            return
        messagebox.showinfo(
            "Saved",
            f"Rotation baked into pixels and flag removed.\n\n{out_path}")


if __name__ == "__main__":
    app = OrientationInspector()
    app.mainloop()
