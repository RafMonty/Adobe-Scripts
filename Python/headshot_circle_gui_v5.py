#!/usr/bin/env python3
"""
headshot_circle_gui.py  (v5 — selectable face-size metric)
----------------------------------------------------------
Windows GUI for batch headshot alignment + circle cropping.

New in v5: a "Scale by" option controlling how head size is measured.
    Eyes            — inter-eye distance (v4 behaviour)
    Eye-mouth       — vertical face height, eye line to mouth
    Combined        — geometric mean of both (recommended; balances
                      face width and height so wide-short and
                      narrow-long faces scale consistently)

Positioning is unchanged: the eye midpoint is always placed at the target
point, so eyes stay level across the whole set — only the size
calculation differs.

Also includes: optional reference image for targets, optional backgrounds
folder (random or stable-per-filename assignment), horizontal-centring
override, flattening of transparent sources, anti-aliased circular crop,
transparent PNG output with filename suffix.

Face/eye detection uses OpenCV's built-in YuNet detector. The ~230 KB
model file is downloaded automatically on first run.

Dependencies:
    pip install opencv-python Pillow numpy

Run:
    python headshot_circle_gui.py
"""

import hashlib
import math
import queue
import random
import threading
import tkinter as tk
import urllib.request
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, ttk

import cv2
import numpy as np
from PIL import Image, ImageDraw

VALID_EXT = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp"}

# Canvas-relative defaults used when no reference image is supplied
DEFAULT_EYE_DIST_FRAC = 0.28    # inter-eye distance / canvas width
DEFAULT_EYE_MOUTH_FRAC = 0.30   # eye-midpoint to mouth-midpoint / canvas width
DEFAULT_EYE_Y_FRAC = 0.42       # eye line height from canvas top
DEFAULT_EYE_X_FRAC = 0.50       # eye midpoint horizontal position

SCALE_MODES = ("Combined (recommended)", "Eyes only", "Eye-mouth height")

MODEL_NAME = "face_detection_yunet_2023mar.onnx"
MODEL_URL = ("https://media.githubusercontent.com/media/opencv/opencv_zoo/"
             "main/models/face_detection_yunet/" + MODEL_NAME)
DETECT_MAX_DIM = 1024


# ===========================================================================
# Face detection (OpenCV YuNet)
# ===========================================================================

def ensure_model(log=print):
    model_path = Path(__file__).resolve().parent / MODEL_NAME
    if model_path.exists() and model_path.stat().st_size > 100_000:
        return model_path
    log(f"Downloading face detection model ({MODEL_NAME})…")
    tmp = model_path.with_suffix(".tmp")
    urllib.request.urlretrieve(MODEL_URL, tmp)
    if tmp.stat().st_size < 100_000:
        tmp.unlink(missing_ok=True)
        raise RuntimeError(
            "Model download failed (got a stub file).\n"
            f"Please download it manually from:\n{MODEL_URL}\n"
            f"and save it next to this script as {MODEL_NAME}"
        )
    tmp.rename(model_path)
    log(f"Model saved to {model_path}")
    return model_path


class FaceDetector:
    """Wraps cv2.FaceDetectorYN; returns eye + mouth landmarks."""

    def __init__(self, model_path):
        self.det = cv2.FaceDetectorYN_create(
            str(model_path), "", (320, 320),
            score_threshold=0.6, nms_threshold=0.3, top_k=500,
        )

    def detect(self, img_bgr):
        """Return landmark dict in original-image pixels, or None.

        Keys: right_eye, left_eye, mouth (midpoint of mouth corners).
        """
        h, w = img_bgr.shape[:2]
        scale = 1.0
        det_img = img_bgr
        if max(h, w) > DETECT_MAX_DIM:
            scale = DETECT_MAX_DIM / max(h, w)
            det_img = cv2.resize(img_bgr, (int(w * scale), int(h * scale)),
                                 interpolation=cv2.INTER_AREA)
        dh, dw = det_img.shape[:2]
        self.det.setInputSize((dw, dh))
        _, faces = self.det.detect(det_img)
        if faces is None or len(faces) == 0:
            return None
        best = max(faces, key=lambda f: f[2] * f[3])
        # Layout: [x,y,w,h, re_x,re_y, le_x,le_y, nose_x,nose_y,
        #          mouth_r_x,mouth_r_y, mouth_l_x,mouth_l_y, score]
        inv = 1.0 / scale
        return {
            "right_eye": (best[4] * inv, best[5] * inv),
            "left_eye": (best[6] * inv, best[7] * inv),
            "mouth": (
                (best[10] + best[12]) / 2.0 * inv,
                (best[11] + best[13]) / 2.0 * inv,
            ),
        }


# ===========================================================================
# Geometry
# ===========================================================================

def eye_midpoint(lm):
    (rx, ry), (lx, ly) = lm["right_eye"], lm["left_eye"]
    return ((rx + lx) / 2.0, (ry + ly) / 2.0)


def face_metric(lm, mode):
    """Face-size measurement in pixels, per the selected scale mode."""
    (rx, ry), (lx, ly) = lm["right_eye"], lm["left_eye"]
    eye_dist = math.hypot(lx - rx, ly - ry)
    mx, my = eye_midpoint(lm)
    em_dist = math.hypot(lm["mouth"][0] - mx, lm["mouth"][1] - my)
    if mode == "Eyes only":
        return eye_dist
    if mode == "Eye-mouth height":
        return em_dist
    # Combined: geometric mean balances face width and height
    return math.sqrt(max(eye_dist, 1e-6) * max(em_dist, 1e-6))


def default_target_metric(size, mode):
    if mode == "Eyes only":
        return size * DEFAULT_EYE_DIST_FRAC
    if mode == "Eye-mouth height":
        return size * DEFAULT_EYE_MOUTH_FRAC
    return size * math.sqrt(DEFAULT_EYE_DIST_FRAC * DEFAULT_EYE_MOUTH_FRAC)


# ===========================================================================
# Image helpers
# ===========================================================================

def flatten_alpha(pil_img, bg=(255, 255, 255)):
    if pil_img.mode in ("RGBA", "LA") or (
        pil_img.mode == "P" and "transparency" in pil_img.info
    ):
        rgba = pil_img.convert("RGBA")
        base = Image.new("RGBA", rgba.size, bg + (255,))
        return Image.alpha_composite(base, rgba).convert("RGB")
    return pil_img.convert("RGB")


def load_background_paths(folder):
    folder = Path(folder)
    return sorted(p for p in folder.iterdir()
                  if p.suffix.lower() in VALID_EXT and p.is_file())


def pick_background(bg_paths, filename, stable):
    if stable:
        digest = hashlib.md5(filename.encode("utf-8")).hexdigest()
        return bg_paths[int(digest, 16) % len(bg_paths)]
    return random.choice(bg_paths)


def cover_crop(pil_img, size):
    img = pil_img.convert("RGB")
    w, h = img.size
    scale = size / min(w, h)
    img = img.resize((max(size, round(w * scale)), max(size, round(h * scale))),
                     Image.LANCZOS)
    w, h = img.size
    left = (w - size) // 2
    top = (h - size) // 2
    return img.crop((left, top, left + size, top + size))


def warp_to_canvas(img_np, anchor_xy, scale, canvas_size, target_xy,
                   border_value):
    """Uniformly scale the image and place anchor_xy at target_xy."""
    tx = target_xy[0] - anchor_xy[0] * scale
    ty = target_xy[1] - anchor_xy[1] * scale
    M = np.array([[scale, 0, tx], [0, scale, ty]], dtype=np.float32)
    return cv2.warpAffine(
        img_np, M, (canvas_size, canvas_size),
        flags=cv2.INTER_LANCZOS4,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=border_value,
    )


def circle_mask(size, supersample=4):
    big = size * supersample
    mask = Image.new("L", (big, big), 0)
    ImageDraw.Draw(mask).ellipse((0, 0, big - 1, big - 1), fill=255)
    return mask.resize((size, size), Image.LANCZOS)


def reference_targets(ref_path, detector, size, bg, mode):
    """Derive (target_metric, target_eye_xy) from a reference image."""
    ref = flatten_alpha(Image.open(ref_path), bg=bg)
    ref_bgr = cv2.cvtColor(np.array(ref), cv2.COLOR_RGB2BGR)
    lm = detector.detect(ref_bgr)
    if lm is None:
        raise ValueError("No face detected in reference image.")
    rw, rh = ref.size
    ref_min = min(rw, rh)
    scale = size / ref_min
    off_x = (rw - ref_min) / 2.0
    off_y = (rh - ref_min) / 2.0
    metric = face_metric(lm, mode) * scale
    mx, my = eye_midpoint(lm)
    xy = ((mx - off_x) * scale, (my - off_y) * scale)
    return metric, xy


# ===========================================================================
# Batch worker
# ===========================================================================

def run_batch(cfg, log, progress, done):
    try:
        in_dir = Path(cfg["input"])
        out_dir = Path(cfg["output"]) if cfg["output"] else in_dir / "circled"
        out_dir.mkdir(parents=True, exist_ok=True)
        size = cfg["size"]
        bg = cfg["bg"]
        suffix = cfg["suffix"]
        mode = cfg["scale_mode"]

        bg_paths = []
        if cfg["backgrounds"]:
            bg_paths = load_background_paths(cfg["backgrounds"])
            if not bg_paths:
                log("Warning: no images found in backgrounds folder — "
                    "falling back to solid fill colour.")
            else:
                assign = "stable per file" if cfg["stable_bg"] else "random"
                log(f"Loaded {len(bg_paths)} background(s), assignment: {assign}")

        detector = FaceDetector(ensure_model(log))

        log(f"Scale metric: {mode}")
        target_metric = default_target_metric(size, mode)
        target_eye_xy = (size * DEFAULT_EYE_X_FRAC, size * DEFAULT_EYE_Y_FRAC)
        if cfg["reference"]:
            target_metric, target_eye_xy = reference_targets(
                cfg["reference"], detector, size, bg, mode
            )
            log(f"Reference targets: metric {target_metric:.0f}px, "
                f"eye midpoint ({target_eye_xy[0]:.0f}, {target_eye_xy[1]:.0f})")

        if cfg["center_x"]:
            target_eye_xy = (size / 2.0, target_eye_xy[1])
            log("Horizontal centring ON — eye midpoint forced to canvas centre.")

        mask = circle_mask(size)

        files = sorted(p for p in in_dir.iterdir()
                       if p.suffix.lower() in VALID_EXT and p.is_file())
        if not files:
            log("No images found in input folder.")
            done(0, 0)
            return

        ok = fail = 0
        total = len(files)
        for i, p in enumerate(files, 1):
            try:
                src = Image.open(p)
                use_bg = bool(bg_paths)

                if use_bg:
                    rgba = src.convert("RGBA")
                    img_np = np.array(rgba)
                    detect_np = np.array(flatten_alpha(src, bg=bg))
                else:
                    pil = flatten_alpha(src, bg=bg)
                    img_np = np.array(pil)
                    detect_np = img_np

                img_bgr = cv2.cvtColor(detect_np, cv2.COLOR_RGB2BGR)
                lm = detector.detect(img_bgr)
                if lm is None:
                    fail += 1
                    log(f"SKIP  {p.name}  (no face detected)")
                    progress(i, total)
                    continue

                metric = face_metric(lm, mode)
                if metric < 1:
                    fail += 1
                    log(f"SKIP  {p.name}  (bad face geometry)")
                    progress(i, total)
                    continue

                scale = target_metric / metric
                border = (0, 0, 0, 0) if use_bg else bg
                canvas = warp_to_canvas(
                    img_np, eye_midpoint(lm), scale, size, target_eye_xy,
                    border_value=border,
                )

                if use_bg:
                    bg_path = pick_background(bg_paths, p.name, cfg["stable_bg"])
                    base = cover_crop(Image.open(bg_path), size).convert("RGBA")
                    out = Image.alpha_composite(base, Image.fromarray(canvas))
                else:
                    out = Image.fromarray(canvas).convert("RGBA")
                out.putalpha(mask)
                out_path = out_dir / f"{p.stem}{suffix}.png"
                out.save(out_path, "PNG")
                ok += 1
                tag = f"  [bg: {bg_path.name}]" if use_bg else ""
                log(f"OK    {p.name}  ->  {out_path.name}{tag}")
            except Exception as e:
                fail += 1
                log(f"ERROR {p.name}  ({e})")
            progress(i, total)

        log(f"\nDone: {ok} processed, {fail} skipped.")
        log(f"Output folder: {out_dir}")
        done(ok, fail)
    except Exception as e:
        log(f"\nFATAL: {e}")
        done(0, 0)


# ===========================================================================
# GUI
# ===========================================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Headshot Circle Cropper")
        self.geometry("660x640")
        self.minsize(580, 520)

        self.msg_q = queue.Queue()
        self.worker = None
        self.bg_colour = (255, 255, 255)

        pad = {"padx": 8, "pady": 4}
        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=10, pady=10)
        frm.columnconfigure(1, weight=1)

        self.var_input = tk.StringVar()
        self.var_output = tk.StringVar()
        self.var_ref = tk.StringVar()
        self.var_bgdir = tk.StringVar()

        ttk.Label(frm, text="Input folder:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.var_input).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_input).grid(row=0, column=2, **pad)

        ttk.Label(frm, text="Output folder:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.var_output).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_output).grid(row=1, column=2, **pad)
        ttk.Label(frm, text="(leave blank for <input>\\circled)",
                  foreground="grey").grid(row=2, column=1, sticky="w", padx=8)

        ttk.Label(frm, text="Reference image:").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.var_ref).grid(row=3, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_ref).grid(row=3, column=2, **pad)
        ttk.Label(frm, text="(optional — sets target head size/position)",
                  foreground="grey").grid(row=4, column=1, sticky="w", padx=8)

        ttk.Label(frm, text="Backgrounds folder:").grid(row=5, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.var_bgdir).grid(row=5, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_bgdir).grid(row=5, column=2, **pad)

        self.var_stable = tk.BooleanVar(value=True)
        bg_row = ttk.Frame(frm)
        bg_row.grid(row=6, column=1, sticky="w", padx=8)
        ttk.Checkbutton(bg_row, text="Same background per person on re-runs",
                        variable=self.var_stable).pack(side="left")
        ttk.Label(frm, text="(optional — composited behind transparent photos)",
                  foreground="grey").grid(row=7, column=1, sticky="w", padx=8)

        opts = ttk.LabelFrame(frm, text="Options")
        opts.grid(row=8, column=0, columnspan=3, sticky="ew", padx=8, pady=8)

        ttk.Label(opts, text="Output size (px):").grid(row=0, column=0, sticky="w", **pad)
        self.var_size = tk.StringVar(value="600")
        ttk.Spinbox(opts, from_=100, to=4000, increment=50,
                    textvariable=self.var_size, width=8).grid(row=0, column=1, **pad)

        ttk.Label(opts, text="Suffix:").grid(row=0, column=2, sticky="w", **pad)
        self.var_suffix = tk.StringVar(value="_circle")
        ttk.Entry(opts, textvariable=self.var_suffix, width=12).grid(row=0, column=3, **pad)

        ttk.Label(opts, text="Fill colour:").grid(row=0, column=4, sticky="w", **pad)
        self.btn_colour = tk.Button(opts, text="      ", bg="#ffffff",
                                    relief="ridge", command=self.pick_colour)
        self.btn_colour.grid(row=0, column=5, **pad)

        ttk.Label(opts, text="Scale by:").grid(row=1, column=0, sticky="w", **pad)
        self.var_scale_mode = tk.StringVar(value=SCALE_MODES[0])
        ttk.Combobox(opts, textvariable=self.var_scale_mode,
                     values=SCALE_MODES, state="readonly",
                     width=24).grid(row=1, column=1, columnspan=2,
                                    sticky="w", **pad)

        self.var_center_x = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opts,
            text="Centre face horizontally (override reference X position)",
            variable=self.var_center_x,
        ).grid(row=2, column=0, columnspan=6, sticky="w", **pad)

        self.btn_run = ttk.Button(frm, text="Run", command=self.start)
        self.btn_run.grid(row=9, column=0, columnspan=3, sticky="ew", padx=8, pady=(8, 4))

        self.prog = ttk.Progressbar(frm, mode="determinate")
        self.prog.grid(row=10, column=0, columnspan=3, sticky="ew", padx=8, pady=4)

        self.log_box = tk.Text(frm, height=12, state="disabled",
                               font=("Consolas", 9))
        self.log_box.grid(row=11, column=0, columnspan=3, sticky="nsew", padx=8, pady=(4, 0))
        frm.rowconfigure(11, weight=1)

        scroll = ttk.Scrollbar(frm, command=self.log_box.yview)
        scroll.grid(row=11, column=3, sticky="ns", pady=(4, 0))
        self.log_box.configure(yscrollcommand=scroll.set)

        self.after(100, self.poll_queue)

    def pick_input(self):
        d = filedialog.askdirectory(title="Select input folder")
        if d:
            self.var_input.set(d)

    def pick_output(self):
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self.var_output.set(d)

    def pick_bgdir(self):
        d = filedialog.askdirectory(title="Select backgrounds folder")
        if d:
            self.var_bgdir.set(d)

    def pick_ref(self):
        f = filedialog.askopenfilename(
            title="Select reference image",
            filetypes=[("Images", "*.jpg *.jpeg *.png *.webp *.tif *.tiff *.bmp"),
                       ("All files", "*.*")],
        )
        if f:
            self.var_ref.set(f)

    def pick_colour(self):
        rgb, hexval = colorchooser.askcolor(color="#%02x%02x%02x" % self.bg_colour,
                                            title="Flatten / fill colour")
        if rgb:
            self.bg_colour = tuple(int(c) for c in rgb)
            self.btn_colour.configure(bg=hexval)

    def log(self, text):
        self.msg_q.put(("log", str(text)))

    def progress(self, i, total):
        self.msg_q.put(("prog", (i, total)))

    def done(self, ok, fail):
        self.msg_q.put(("done", (ok, fail)))

    def poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_q.get_nowait()
                if kind == "log":
                    self.log_box.configure(state="normal")
                    self.log_box.insert("end", payload + "\n")
                    self.log_box.see("end")
                    self.log_box.configure(state="disabled")
                elif kind == "prog":
                    i, total = payload
                    self.prog["maximum"] = total
                    self.prog["value"] = i
                elif kind == "done":
                    ok, fail = payload
                    self.btn_run.configure(state="normal", text="Run")
                    if ok or fail:
                        messagebox.showinfo(
                            "Finished",
                            f"{ok} image(s) processed, {fail} skipped."
                        )
        except queue.Empty:
            pass
        self.after(100, self.poll_queue)

    def start(self):
        in_dir = self.var_input.get().strip()
        if not in_dir or not Path(in_dir).is_dir():
            messagebox.showerror("Missing input", "Please select a valid input folder.")
            return
        ref = self.var_ref.get().strip()
        if ref and not Path(ref).is_file():
            messagebox.showerror("Bad reference", "Reference image not found.")
            return
        bgdir = self.var_bgdir.get().strip()
        if bgdir and not Path(bgdir).is_dir():
            messagebox.showerror("Bad backgrounds folder",
                                 "Backgrounds folder not found.")
            return
        try:
            size = int(self.var_size.get())
            if size < 50:
                raise ValueError
        except ValueError:
            messagebox.showerror("Bad size", "Output size must be a number ≥ 50.")
            return

        cfg = {
            "input": in_dir,
            "output": self.var_output.get().strip() or None,
            "reference": ref or None,
            "size": size,
            "suffix": self.var_suffix.get() or "_circle",
            "bg": self.bg_colour,
            "backgrounds": bgdir or None,
            "stable_bg": self.var_stable.get(),
            "center_x": self.var_center_x.get(),
            "scale_mode": self.var_scale_mode.get(),
        }

        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        self.prog["value"] = 0
        self.btn_run.configure(state="disabled", text="Processing…")

        self.worker = threading.Thread(
            target=run_batch, args=(cfg, self.log, self.progress, self.done),
            daemon=True,
        )
        self.worker.start()


if __name__ == "__main__":
    App().mainloop()
