#!/usr/bin/env python3
"""
headshot_circle_gui.py  (v7 — settings files, drag & drop, watched folders)
---------------------------------------------------------------------------
Windows GUI for batch headshot alignment + circle cropping.

New in v7:
    * Settings save/load — all options serialise to a small JSON preset
      file (Save/Load buttons).  The last-used settings are remembered
      automatically between sessions.
    * Drag & drop        — drop image files or folders anywhere on the
      window to process them with the current on-screen settings
      (requires the optional tkinterdnd2 package).
    * Watched folders    — a second tab lists hot folders, each holding a
      headcropper.json settings file.  While watching is on, new images
      appearing in a folder are processed per that folder's settings;
      originals are moved to processed\\ (or failed\\).  Reference and
      backgrounds paths in the settings file may be relative to the
      folder, so watched folders stay portable.
    * Drop targets      — a tab of named drop tiles.  Each tile links a
      settings preset (.json, re-read at every drop) to a required output
      folder.  Drop files/folders on a tile and they process with that
      preset — previews stay off for speed, and the tile shows its own
      progress bar.
    * Jobs queue up: drops and watch events wait politely behind any
      batch that is already running.

New in v6:
    * Live preview       — each result is shown in a window as it is produced.
    * Landmark overlay   — the preview (not the saved file) can show the
                           detected face landmarks in bright green, plus the
                           target eye position in magenta.  A second window
                           shows the reference image with its landmarks.
    * Match reference    — output canvas takes the reference image's exact
                           resolution and aspect ratio; faces are scaled and
                           positioned to match the reference 1:1.  Where a
                           source image falls short of the canvas the gap is
                           padded with fully transparent pixels.  The circular
                           crop only applies when the output is square.

Also fixed/improved in v6:
    * EXIF orientation is now honoured (phone photos no longer load sideways).
    * Backgrounds are scaled once and cached, not re-loaded per photo, and
      the scaling is selectable: Auto (fill + centre-crop, the classic
      behaviour), Fit (letterbox), Stretch, or a Custom percentage.
    * Optional halo fix: RGBA sources can be warped with premultiplied
      alpha to avoid dark edge fringing (on by default, user-selectable).
    * Transparent-source handling is explicit: pre-contoured PNGs can keep
      their transparency, be flattened onto the fill colour (classic), or
      be composited over background images from a folder.
    * Output format is selectable: PNG, or JPEG with a quality setting
      (JPEG cannot store alpha, so transparency is flattened onto the
      fill colour).

Carried over from v5: selectable face-size metric (eyes / eye-mouth /
combined), optional reference image for targets, optional backgrounds folder
(random or stable-per-filename), horizontal-centring override, flattening of
transparent sources, anti-aliased circular crop, transparent PNG output.

Face/eye detection uses OpenCV's built-in YuNet detector. The ~230 KB
model file is downloaded automatically on first run.

Dependencies:
    pip install opencv-python Pillow numpy
    pip install tkinterdnd2        (optional — enables drag & drop)

Run:
    python headshot_circle_gui_v7.pyw
"""

import hashlib
import json
import math
import queue
import random
import shutil
import threading
import tkinter as tk
import urllib.request
from collections import deque
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, ttk

import cv2
import numpy as np
from PIL import Image, ImageChops, ImageDraw, ImageOps, ImageTk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    BaseTk = TkinterDnD.Tk
    HAS_DND = True
except ImportError:
    BaseTk = tk.Tk
    HAS_DND = False

VALID_EXT = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp"}

# Canvas-relative defaults used when no reference image is supplied
DEFAULT_EYE_DIST_FRAC = 0.28    # inter-eye distance / canvas width
DEFAULT_EYE_MOUTH_FRAC = 0.30   # eye-midpoint to mouth-midpoint / canvas width
DEFAULT_EYE_Y_FRAC = 0.42       # eye line height from canvas top
DEFAULT_EYE_X_FRAC = 0.50       # eye midpoint horizontal position

SCALE_MODES = ("Combined (recommended)", "Eyes only", "Eye-mouth height")

# How transparent (pre-contoured PNG) sources are backed
BG_SOURCE_MODES = ("Fill colour", "Keep transparent",
                   "Background images (folder)")

OUTPUT_FORMATS = ("PNG", "JPEG")

SETTINGS_VERSION = 1
WATCH_SETTINGS_NAME = "headcropper.json"   # per watched-folder settings file
LAST_SETTINGS_NAME = "headcropper_last.json"
WATCH_POLL_MS = 3000                       # watched-folder poll interval
WATCH_DONE_DIR = "processed"
WATCH_FAIL_DIR = "failed"

MODEL_NAME = "face_detection_yunet_2023mar.onnx"
MODEL_URL = ("https://media.githubusercontent.com/media/opencv/opencv_zoo/"
             "main/models/face_detection_yunet/" + MODEL_NAME)
DETECT_MAX_DIM = 1024

PREVIEW_MAX = 520               # preview windows fit within this many pixels

MARK_GREEN = (0, 255, 0, 255)
MARK_MAGENTA = (255, 0, 255, 255)


# ===========================================================================
# Face detection (OpenCV YuNet)
# ===========================================================================

def ensure_model(log=print):
    model_path = Path(__file__).resolve().parent / MODEL_NAME
    if model_path.exists() and model_path.stat().st_size > 100_000:
        return model_path
    log(f"Downloading face detection model ({MODEL_NAME})…")
    tmp = model_path.with_suffix(".tmp")
    with urllib.request.urlopen(MODEL_URL, timeout=60) as resp, \
            open(tmp, "wb") as fh:
        shutil.copyfileobj(resp, fh)
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
    """Wraps cv2.FaceDetectorYN; returns eye, nose and mouth landmarks."""

    def __init__(self, model_path):
        self.det = cv2.FaceDetectorYN_create(
            str(model_path), "", (320, 320),
            score_threshold=0.6, nms_threshold=0.3, top_k=500,
        )

    def detect(self, img_bgr):
        """Return landmark dict in original-image pixels, or None.

        Keys: right_eye, left_eye, nose, mouth_right, mouth_left,
        mouth (midpoint of mouth corners).
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
            "nose": (best[8] * inv, best[9] * inv),
            "mouth_right": (best[10] * inv, best[11] * inv),
            "mouth_left": (best[12] * inv, best[13] * inv),
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


def transform_landmarks(lm, scale, anchor_xy, target_xy):
    """Map source-pixel landmarks into output-canvas coordinates."""
    tx = target_xy[0] - anchor_xy[0] * scale
    ty = target_xy[1] - anchor_xy[1] * scale
    return {k: (x * scale + tx, y * scale + ty) for k, (x, y) in lm.items()}


# ===========================================================================
# Image helpers
# ===========================================================================

def open_image(path):
    """Open an image with EXIF orientation applied.

    Fully loads and releases the file handle so the original can be
    moved afterwards (watched folders move processed files).
    """
    with Image.open(path) as img:
        img.load()
        return ImageOps.exif_transpose(img)


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


BG_MODES = ("Auto (fill, centre-crop)", "Fit (letterbox)",
            "Stretch to canvas", "Custom scale %")


def prepare_background(pil_img, cw, ch, bg_mode, pct, fill):
    """Scale a background to exactly cw x ch per the selected mode.

    Auto     — cover-fit and centre-crop (the classic behaviour).
    Fit      — letterbox inside the canvas; gaps take the fill colour.
    Stretch  — resize to the canvas, ignoring aspect ratio.
    Custom % — like Auto but scaled to pct% of the cover size, centred;
               any exposed canvas takes the fill colour.
    """
    img = pil_img.convert("RGB")
    w, h = img.size
    if bg_mode == "Stretch to canvas":
        return img.resize((cw, ch), Image.LANCZOS)
    if bg_mode == "Fit (letterbox)":
        scale = min(cw / w, ch / h)
    else:
        scale = max(cw / w, ch / h)
        if bg_mode == "Custom scale %":
            scale *= pct / 100.0
    nw = max(1, round(w * scale))
    nh = max(1, round(h * scale))
    img = img.resize((nw, nh), Image.LANCZOS)
    if nw == cw and nh == ch:
        return img
    if nw >= cw and nh >= ch:
        left = (nw - cw) // 2
        top = (nh - ch) // 2
        return img.crop((left, top, left + cw, top + ch))
    # Smaller in at least one dimension: centre over fill colour,
    # cropping any overhang in the other dimension.
    base = Image.new("RGB", (cw, ch), fill)
    if nw > cw or nh > ch:
        left = max(0, (nw - cw) // 2)
        top = max(0, (nh - ch) // 2)
        img = img.crop((left, top, left + min(nw, cw), top + min(nh, ch)))
        nw, nh = img.size
    base.paste(img, ((cw - nw) // 2, (ch - nh) // 2))
    return base


def warp_to_canvas(img_np, anchor_xy, scale, canvas_wh, target_xy,
                   border_value):
    """Uniformly scale an RGB image and place anchor_xy at target_xy."""
    tx = target_xy[0] - anchor_xy[0] * scale
    ty = target_xy[1] - anchor_xy[1] * scale
    M = np.array([[scale, 0, tx], [0, scale, ty]], dtype=np.float32)
    return cv2.warpAffine(
        img_np, M, canvas_wh,
        flags=cv2.INTER_LANCZOS4,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=border_value,
    )


def warp_rgba_premult(img_np, anchor_xy, scale, canvas_wh, target_xy):
    """Warp an RGBA image with premultiplied alpha to avoid edge halos."""
    f = img_np.astype(np.float32)
    alpha = f[..., 3:4] / 255.0
    f[..., :3] *= alpha
    tx = target_xy[0] - anchor_xy[0] * scale
    ty = target_xy[1] - anchor_xy[1] * scale
    M = np.array([[scale, 0, tx], [0, scale, ty]], dtype=np.float32)
    warped = cv2.warpAffine(
        f, M, canvas_wh,
        flags=cv2.INTER_LANCZOS4,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=(0, 0, 0, 0),
    )
    a = np.clip(warped[..., 3:4], 0.0, 255.0)
    rgb = warped[..., :3] * (255.0 / np.maximum(a, 1e-3))
    rgb = np.clip(rgb, 0.0, 255.0)
    return np.dstack((rgb, a)).astype(np.uint8)


def circle_mask(size, supersample=4):
    big = size * supersample
    mask = Image.new("L", (big, big), 0)
    ImageDraw.Draw(mask).ellipse((0, 0, big - 1, big - 1), fill=255)
    return mask.resize((size, size), Image.LANCZOS)


def analyze_reference(ref_path, detector, bg):
    """Load the reference image (flattened) and detect its landmarks."""
    ref = flatten_alpha(open_image(ref_path), bg=bg)
    ref_bgr = cv2.cvtColor(np.array(ref), cv2.COLOR_RGB2BGR)
    lm = detector.detect(ref_bgr)
    if lm is None:
        raise ValueError("No face detected in reference image.")
    return ref, lm


def square_reference_targets(ref, lm, size, mode):
    """Derive (target_metric, target_eye_xy) for square output.

    Maps the reference through the same scale + centre-crop it would get
    if it were cover-fitted onto the square canvas (v5 behaviour).
    """
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
# Landmark overlay (preview only — never baked into saved files)
# ===========================================================================

def draw_landmarks(pil_rgba, lm, target_xy=None):
    """Draw detection markers in bright green; optional target in magenta."""
    d = ImageDraw.Draw(pil_rgba)
    w, h = pil_rgba.size
    r = max(3, round(min(w, h) / 150))
    lw = max(1, r // 2)

    re, le = lm["right_eye"], lm["left_eye"]
    mx, my = eye_midpoint(lm)
    mo = lm["mouth"]

    # Structure lines: eye line, eye-midpoint to mouth
    d.line([re, le], fill=MARK_GREEN, width=lw)
    d.line([(mx, my), mo], fill=MARK_GREEN, width=lw)

    # Point markers
    for key in ("right_eye", "left_eye", "nose", "mouth_right", "mouth_left"):
        x, y = lm[key]
        d.ellipse((x - r, y - r, x + r, y + r), outline=MARK_GREEN, width=lw)

    # Detected eye-midpoint crosshair (green)
    c = r * 2
    d.line([(mx - c, my), (mx + c, my)], fill=MARK_GREEN, width=lw)
    d.line([(mx, my - c), (mx, my + c)], fill=MARK_GREEN, width=lw)

    # Target eye position crosshair (magenta) — shows detected-vs-target
    if target_xy is not None:
        tx, ty = target_xy
        c = r * 3
        d.line([(tx - c, ty), (tx + c, ty)], fill=MARK_MAGENTA, width=lw)
        d.line([(tx, ty - c), (tx, ty + c)], fill=MARK_MAGENTA, width=lw)


# ===========================================================================
# Settings (JSON presets — shared by Save/Load, last-used, watched folders)
# ===========================================================================

DEFAULT_SETTINGS = {
    "version": SETTINGS_VERSION,
    "input": "",
    "output": "",
    "reference": "",
    "backgrounds": "",
    "size": 600,
    "suffix": "_circle",
    "fill_colour": "#ffffff",
    "stable_bg": True,
    "center_x": True,
    "scale_mode": SCALE_MODES[0],
    "match_ref": False,
    "preview": True,
    "marks": True,
    "halo_fix": True,
    "bg_mode": BG_MODES[0],
    "bg_pct": 100.0,
    "bg_source": BG_SOURCE_MODES[0],
    "format": OUTPUT_FORMATS[0],
    "jpeg_q": 90,
}


def hex_to_colour(value, default=(255, 255, 255)):
    try:
        value = value.lstrip("#")
        return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return default


def colour_to_hex(rgb):
    return "#%02x%02x%02x" % tuple(rgb)


def clean_settings(raw):
    """Merge a loaded dict over the defaults, dropping bad values."""
    sd = dict(DEFAULT_SETTINGS)
    if not isinstance(raw, dict):
        return sd
    for key, dval in DEFAULT_SETTINGS.items():
        if key not in raw:
            continue
        val = raw[key]
        try:
            if isinstance(dval, bool):
                sd[key] = bool(val)
            elif isinstance(dval, int):
                sd[key] = int(val)
            elif isinstance(dval, float):
                sd[key] = float(val)
            else:
                sd[key] = str(val)
        except (TypeError, ValueError):
            pass
    if sd["scale_mode"] not in SCALE_MODES:
        sd["scale_mode"] = SCALE_MODES[0]
    if sd["bg_mode"] not in BG_MODES:
        sd["bg_mode"] = BG_MODES[0]
    if sd["bg_source"] not in BG_SOURCE_MODES:
        sd["bg_source"] = BG_SOURCE_MODES[0]
    if sd["format"] not in OUTPUT_FORMATS:
        sd["format"] = OUTPUT_FORMATS[0]
    sd["size"] = min(max(sd["size"], 50), 8000)
    sd["bg_pct"] = min(max(sd["bg_pct"], 10.0), 400.0)
    sd["jpeg_q"] = min(max(sd["jpeg_q"], 1), 100)
    sd["version"] = SETTINGS_VERSION
    return sd


def load_settings_file(path):
    with open(path, "r", encoding="utf-8") as fh:
        return clean_settings(json.load(fh))


def save_settings_file(path, sd, base_dir=None):
    """Write settings; paths under base_dir are stored relative (portable)."""
    out = dict(sd)
    if base_dir is not None:
        base = Path(base_dir).resolve()
        for key in ("input", "output", "reference", "backgrounds"):
            val = out.get(key, "")
            if val:
                try:
                    out[key] = str(Path(val).resolve().relative_to(base))
                except ValueError:
                    pass    # outside base_dir — keep absolute
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(out, fh, indent=2)


def resolve_path(value, base_dir):
    """Absolute path from a settings value, resolving relative to base_dir."""
    if not value:
        return None
    p = Path(value)
    if not p.is_absolute() and base_dir is not None:
        p = Path(base_dir) / p
    return str(p)


def settings_to_cfg(sd, base_dir, folder, log):
    """Build a run_batch cfg from a settings dict (watched-folder jobs).

    `folder` is the watched folder providing the images; paths in the
    settings resolve relative to `base_dir` (the settings file's folder).
    Invalid optional paths are dropped with a warning rather than failing.
    """
    ref = resolve_path(sd["reference"], base_dir)
    if ref and not Path(ref).is_file():
        log(f"Warning: reference not found ({ref}) — using defaults.")
        ref = None
    bgdir = resolve_path(sd["backgrounds"], base_dir)
    if bgdir and not Path(bgdir).is_dir():
        log(f"Warning: backgrounds folder not found ({bgdir}) — ignored.")
        bgdir = None
    output = resolve_path(sd["output"], base_dir)
    return {
        "input": str(folder),
        "output": output,
        "reference": ref,
        "size": sd["size"],
        "suffix": sd["suffix"] or "_circle",
        "bg": hex_to_colour(sd["fill_colour"]),
        "backgrounds": bgdir,
        "stable_bg": sd["stable_bg"],
        "center_x": sd["center_x"],
        "scale_mode": sd["scale_mode"],
        "match_ref": sd["match_ref"] and bool(ref),
        "preview": sd["preview"],
        "marks": sd["marks"],
        "halo_fix": sd["halo_fix"],
        "bg_mode": sd["bg_mode"],
        "bg_pct": sd["bg_pct"],
        "bg_source": sd["bg_source"],
        "format": sd["format"],
        "jpeg_q": sd["jpeg_q"],
    }


# ===========================================================================
# Batch worker
# ===========================================================================

def move_original(p, subname):
    """Move a source file into a sibling subfolder, avoiding collisions."""
    dest_dir = p.parent / subname
    dest_dir.mkdir(exist_ok=True)
    dest = dest_dir / p.name
    n = 1
    while dest.exists():
        dest = dest_dir / f"{p.stem}_{n}{p.suffix}"
        n += 1
    shutil.move(str(p), str(dest))


def run_batch(cfg, log, progress, done, preview):
    try:
        files_arg = cfg.get("files")
        in_dir = Path(cfg["input"]) if cfg["input"] else None
        move_originals = cfg.get("move_originals", False)
        if cfg["output"]:
            fixed_out = Path(cfg["output"])
        elif files_arg:
            fixed_out = None            # per-source-folder \circled
        else:
            fixed_out = in_dir / "circled"
        if fixed_out is not None:
            fixed_out.mkdir(parents=True, exist_ok=True)
        size = cfg["size"]
        bg = cfg["bg"]
        suffix = cfg["suffix"]
        mode = cfg["scale_mode"]
        match_ref = cfg["match_ref"]
        show_preview = cfg["preview"]
        show_marks = cfg["marks"]
        bg_source = cfg["bg_source"]
        fmt = cfg["format"]
        jpeg_q = cfg["jpeg_q"]

        bg_paths = []
        if bg_source == "Background images (folder)":
            if cfg["backgrounds"]:
                bg_paths = load_background_paths(cfg["backgrounds"])
            if not bg_paths:
                log("Warning: no images found in backgrounds folder — "
                    "falling back to solid fill colour.")
                bg_source = "Fill colour"
            else:
                assign = "stable per file" if cfg["stable_bg"] else "random"
                log(f"Loaded {len(bg_paths)} background(s), assignment: {assign}")
                bg_note = cfg["bg_mode"]
                if cfg["bg_mode"] == "Custom scale %":
                    bg_note += f" ({cfg['bg_pct']:.0f}%)"
                log(f"Background scaling: {bg_note}")
        log(f"Transparent-source handling: {bg_source}")
        if fmt == "JPEG":
            log(f"Output format: JPEG (quality {jpeg_q}) — any transparency "
                "is flattened onto the fill colour.")
        else:
            log("Output format: PNG")

        detector = FaceDetector(ensure_model(log))
        log(f"Scale metric: {mode}")

        ref_img = ref_lm = None
        if cfg["reference"]:
            ref_img, ref_lm = analyze_reference(cfg["reference"], detector, bg)

        if match_ref:
            canvas_w, canvas_h = ref_img.size
            target_metric = face_metric(ref_lm, mode)
            target_eye_xy = eye_midpoint(ref_lm)
            log(f"Match-reference mode: canvas {canvas_w}x{canvas_h} px "
                "(output size field ignored). Short canvas edges are padded "
                "with transparency.")
        else:
            canvas_w = canvas_h = size
            if ref_img is not None:
                target_metric, target_eye_xy = square_reference_targets(
                    ref_img, ref_lm, size, mode)
            else:
                target_metric = default_target_metric(size, mode)
                target_eye_xy = (size * DEFAULT_EYE_X_FRAC,
                                 size * DEFAULT_EYE_Y_FRAC)

        if ref_img is not None:
            log(f"Reference targets: metric {target_metric:.0f}px, "
                f"eye midpoint ({target_eye_xy[0]:.0f}, {target_eye_xy[1]:.0f})")

        if cfg["center_x"]:
            target_eye_xy = (canvas_w / 2.0, target_eye_xy[1])
            log("Horizontal centring ON — eye midpoint forced to canvas centre.")

        use_circle = canvas_w == canvas_h
        mask = circle_mask(canvas_w) if use_circle else None
        if not use_circle:
            log("Output is not square — circular crop disabled.")

        if show_preview and ref_img is not None:
            disp = ref_img.convert("RGBA")
            if show_marks:
                draw_landmarks(disp, ref_lm)
            preview("ref", disp, Path(cfg["reference"]).name)

        if files_arg:
            files = [Path(f) for f in files_arg]
        else:
            files = sorted(p for p in in_dir.iterdir()
                           if p.suffix.lower() in VALID_EXT and p.is_file())
        if not files:
            log("No images found in input folder.")
            done(0, 0)
            return

        use_bg = bool(bg_paths)
        # RGBA pipeline whenever source alpha must survive the warp:
        # backgrounds composited later, or transparency kept in the output.
        keep_alpha = use_bg or bg_source == "Keep transparent"
        ext = ".jpg" if fmt == "JPEG" else ".png"
        bg_cache = {}

        def finish_original(p, success):
            if not move_originals:
                return
            try:
                move_original(p, WATCH_DONE_DIR if success else WATCH_FAIL_DIR)
            except Exception as e:
                log(f"Warning: could not move {p.name} ({e})")

        ok = fail = 0
        total = len(files)
        for i, p in enumerate(files, 1):
            try:
                src = open_image(p)

                if keep_alpha:
                    img_np = np.array(src.convert("RGBA"))
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
                    finish_original(p, False)
                    progress(i, total)
                    continue

                metric = face_metric(lm, mode)
                if metric < 1:
                    fail += 1
                    log(f"SKIP  {p.name}  (bad face geometry)")
                    finish_original(p, False)
                    progress(i, total)
                    continue

                scale = target_metric / metric
                anchor = eye_midpoint(lm)
                if keep_alpha:
                    if cfg["halo_fix"]:
                        canvas = warp_rgba_premult(
                            img_np, anchor, scale, (canvas_w, canvas_h),
                            target_eye_xy,
                        )
                    else:
                        canvas = warp_to_canvas(
                            img_np, anchor, scale, (canvas_w, canvas_h),
                            target_eye_xy, border_value=(0, 0, 0, 0),
                        )
                else:
                    canvas = warp_to_canvas(
                        img_np, anchor, scale, (canvas_w, canvas_h),
                        target_eye_xy, border_value=bg,
                    )

                bg_path = None
                if use_bg:
                    bg_path = pick_background(bg_paths, p.name, cfg["stable_bg"])
                    base = bg_cache.get(bg_path)
                    if base is None:
                        base = prepare_background(
                            open_image(bg_path), canvas_w, canvas_h,
                            cfg["bg_mode"], cfg["bg_pct"], bg,
                        ).convert("RGBA")
                        bg_cache[bg_path] = base
                    out = Image.alpha_composite(base, Image.fromarray(canvas))
                else:
                    out = Image.fromarray(canvas).convert("RGBA")

                if mask is not None:
                    # Multiply so transparent padding stays transparent
                    out.putalpha(ImageChops.multiply(out.getchannel("A"), mask))

                out_dir = fixed_out
                if out_dir is None:
                    out_dir = p.parent / "circled"
                    out_dir.mkdir(exist_ok=True)
                out_path = out_dir / f"{p.stem}{suffix}{ext}"
                if fmt == "JPEG":
                    flat = Image.new("RGB", out.size, bg)
                    flat.paste(out, mask=out.getchannel("A"))
                    flat.save(out_path, "JPEG", quality=jpeg_q)
                    out = flat.convert("RGBA")
                else:
                    out.save(out_path, "PNG")
                ok += 1
                tag = f"  [bg: {bg_path.name}]" if bg_path else ""
                log(f"OK    {p.name}  ->  {out_path.name}{tag}")
                finish_original(p, True)

                if show_preview:
                    disp = out.copy()
                    if show_marks:
                        out_lm = transform_landmarks(lm, scale, anchor,
                                                     target_eye_xy)
                        draw_landmarks(disp, out_lm, target_xy=target_eye_xy)
                    preview("result", disp,
                            f"{p.name}  ({i}/{total})")
            except Exception as e:
                fail += 1
                log(f"ERROR {p.name}  ({e})")
                finish_original(p, False)
            progress(i, total)

        log(f"\nDone: {ok} processed, {fail} skipped.")
        log(f"Output folder: {fixed_out if fixed_out else 'per-source folders'}")
        done(ok, fail)
    except Exception as e:
        log(f"\nFATAL: {e}")
        done(0, 0)


# ===========================================================================
# GUI
# ===========================================================================

class PreviewWindow(tk.Toplevel):
    """Reusable image window; closing hides it (batch keeps running)."""

    def __init__(self, master, title):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.withdraw)
        self.caption = ttk.Label(self, anchor="center")
        self.caption.pack(fill="x", padx=4, pady=(4, 0))
        self.label = tk.Label(self)
        self.label.pack(padx=4, pady=4)
        self.photo = None

    def show_image(self, pil_img, caption):
        img = pil_img.copy()
        img.thumbnail((PREVIEW_MAX, PREVIEW_MAX), Image.LANCZOS)
        self.photo = ImageTk.PhotoImage(img)
        self.label.configure(image=self.photo)
        self.caption.configure(
            text=f"{caption}   ({pil_img.width}x{pil_img.height}px)")
        if self.state() == "withdrawn":
            self.deiconify()


TILE_COLOURS = ("#dbeafe", "#dcfce7", "#fef9c3", "#fde2e2",
                "#ede9fe", "#cffafe")


def shorten_path(text, limit=42):
    return text if len(text) <= limit else "…" + text[-(limit - 1):]


class DropTile(tk.Frame):
    """A named drop zone bound to a preset file and an output folder."""

    def __init__(self, master, target, colour, on_remove, on_drop):
        super().__init__(master, bd=2, relief="ridge", bg=colour)
        self.target = target
        top = tk.Frame(self, bg=colour)
        top.pack(fill="x", padx=8, pady=(6, 0))
        tk.Label(top, text=target["name"], bg=colour,
                 font=("Segoe UI", 10, "bold")).pack(side="left")
        tk.Button(top, text="✕", bg=colour, relief="flat", bd=0,
                  command=on_remove).pack(side="right")
        tk.Label(self, text=f"preset: {Path(target['preset']).name}",
                 bg=colour, fg="#555555").pack(anchor="w", padx=8)
        tk.Label(self, text="→ " + shorten_path(target["output"]),
                 bg=colour, fg="#555555").pack(anchor="w", padx=8)
        self.bar = ttk.Progressbar(self, mode="determinate")
        self.bar.pack(fill="x", padx=8, pady=(4, 8))
        if HAS_DND:
            handler = lambda e: on_drop(target, e)
            for w in self._all_widgets():
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", handler)

    def _all_widgets(self):
        stack, out = [self], []
        while stack:
            w = stack.pop()
            out.append(w)
            stack.extend(w.winfo_children())
        return out

    def set_progress(self, i, total):
        self.bar["maximum"] = total
        self.bar["value"] = i

    def reset_progress(self):
        self.bar["value"] = 0


class TargetDialog(tk.Toplevel):
    """Modal dialog to create a drop target — every field is required."""

    def __init__(self, master, existing_names):
        super().__init__(master)
        self.title("New drop target")
        self.resizable(False, False)
        self.result = None
        self.existing = existing_names
        pad = {"padx": 8, "pady": 4}

        self.var_name = tk.StringVar()
        self.var_preset = tk.StringVar()
        self.var_out = tk.StringVar()

        ttk.Label(self, text="Name:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(self, textvariable=self.var_name,
                  width=34).grid(row=0, column=1, sticky="ew", **pad)

        ttk.Label(self, text="Preset file:").grid(row=1, column=0,
                                                  sticky="w", **pad)
        ttk.Entry(self, textvariable=self.var_preset,
                  width=34).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(self, text="Browse…",
                   command=self._pick_preset).grid(row=1, column=2, **pad)

        ttk.Label(self, text="Output folder:").grid(row=2, column=0,
                                                    sticky="w", **pad)
        ttk.Entry(self, textvariable=self.var_out,
                  width=34).grid(row=2, column=1, sticky="ew", **pad)
        ttk.Button(self, text="Browse…",
                   command=self._pick_out).grid(row=2, column=2, **pad)

        btns = ttk.Frame(self)
        btns.grid(row=3, column=0, columnspan=3, pady=(8, 8))
        ttk.Button(btns, text="Create", command=self._ok).pack(side="left",
                                                               padx=4)
        ttk.Button(btns, text="Cancel",
                   command=self.destroy).pack(side="left", padx=4)
        self.grab_set()
        self.transient(master)

    def _pick_preset(self):
        f = filedialog.askopenfilename(
            parent=self, title="Select settings preset",
            filetypes=[("JSON settings", "*.json"), ("All files", "*.*")])
        if f:
            self.var_preset.set(f)

    def _pick_out(self):
        d = filedialog.askdirectory(parent=self, title="Select output folder")
        if d:
            self.var_out.set(d)

    def _ok(self):
        name = self.var_name.get().strip()
        preset = self.var_preset.get().strip()
        out = self.var_out.get().strip()
        if not name:
            messagebox.showerror("Name required",
                                 "Give the target a name.", parent=self)
            return
        if name in self.existing:
            messagebox.showerror("Duplicate name",
                                 "A target with that name already exists.",
                                 parent=self)
            return
        if not preset or not Path(preset).is_file():
            messagebox.showerror("Preset required",
                                 "Select an existing settings preset file "
                                 "(save one from the Batch tab first).",
                                 parent=self)
            return
        try:
            load_settings_file(preset)
        except (OSError, json.JSONDecodeError) as e:
            messagebox.showerror("Bad preset",
                                 f"That file is not a valid settings "
                                 f"preset:\n{e}", parent=self)
            return
        if not out or not Path(out).is_dir():
            messagebox.showerror("Output folder required",
                                 "Select an existing output folder — a drop "
                                 "target cannot be created without one.",
                                 parent=self)
            return
        self.result = {"name": name, "preset": preset, "output": out}
        self.destroy()


class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.title("Headshot Circle Cropper v7")
        self.geometry("680x780")
        self.minsize(600, 640)

        self.msg_q = queue.Queue()
        self.worker = None
        self.bg_colour = (255, 255, 255)
        self.preview_windows = {}

        # Job queue: manual runs, drops and watch events execute in order
        self.jobs = deque()
        self.busy = False
        self.current_job = None

        # Watched-folder state
        self.watching = False
        self.watch_seen = {}        # path -> last seen size (stability check)
        self.watch_inflight = set() # paths queued or being processed
        self.watch_bad_settings = set()

        pad = {"padx": 8, "pady": 4}
        nb = ttk.Notebook(self)
        nb.pack(fill="x", padx=10, pady=(10, 0))

        frm = ttk.Frame(nb)
        nb.add(frm, text="Batch")
        frm.columnconfigure(1, weight=1)

        targets_tab = ttk.Frame(nb)
        nb.add(targets_tab, text="Drop targets")

        watch_tab = ttk.Frame(nb)
        nb.add(watch_tab, text="Watched folders")

        # Drop-target state
        self.drop_targets = []      # [{name, preset, output}, ...]
        self.tiles = {}             # name -> DropTile

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

        self.var_match = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frm,
            text="Match reference resolution && aspect ratio "
                 "(pads shortfall with transparency)",
            variable=self.var_match,
        ).grid(row=5, column=0, columnspan=3, sticky="w", padx=8)

        ttk.Label(frm, text="Backgrounds folder:").grid(row=6, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.var_bgdir).grid(row=6, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_bgdir).grid(row=6, column=2, **pad)

        self.var_stable = tk.BooleanVar(value=True)
        bg_row = ttk.Frame(frm)
        bg_row.grid(row=7, column=1, sticky="w", padx=8)
        ttk.Checkbutton(bg_row, text="Same background per person on re-runs",
                        variable=self.var_stable).pack(side="left")
        ttk.Label(frm, text="(optional — composited behind transparent photos)",
                  foreground="grey").grid(row=8, column=1, sticky="w", padx=8)

        opts = ttk.LabelFrame(frm, text="Options")
        opts.grid(row=9, column=0, columnspan=3, sticky="ew", padx=8, pady=8)

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

        ttk.Label(opts, text="BG scaling:").grid(row=1, column=3, sticky="e", **pad)
        self.var_bg_mode = tk.StringVar(value=BG_MODES[0])
        self.cmb_bg_mode = ttk.Combobox(opts, textvariable=self.var_bg_mode,
                                        values=BG_MODES, state="readonly",
                                        width=20)
        self.cmb_bg_mode.grid(row=1, column=4, columnspan=2, sticky="w", **pad)
        self.cmb_bg_mode.bind("<<ComboboxSelected>>", self._bg_mode_changed)

        ttk.Label(opts, text="BG custom %:").grid(row=2, column=3, sticky="e", **pad)
        self.var_bg_pct = tk.StringVar(value="100")
        self.spn_bg_pct = ttk.Spinbox(opts, from_=10, to=400, increment=5,
                                      textvariable=self.var_bg_pct, width=8,
                                      state="disabled")
        self.spn_bg_pct.grid(row=2, column=4, sticky="w", **pad)

        self.var_center_x = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opts,
            text="Centre face horizontally (override reference X position)",
            variable=self.var_center_x,
        ).grid(row=2, column=0, columnspan=3, sticky="w", **pad)

        ttk.Label(opts, text="Transparent BG:").grid(row=3, column=0,
                                                     sticky="w", **pad)
        self.var_bg_source = tk.StringVar(value=BG_SOURCE_MODES[0])
        ttk.Combobox(opts, textvariable=self.var_bg_source,
                     values=BG_SOURCE_MODES, state="readonly",
                     width=24).grid(row=3, column=1, columnspan=2,
                                    sticky="w", **pad)

        ttk.Label(opts, text="Format:").grid(row=3, column=3, sticky="e", **pad)
        self.var_format = tk.StringVar(value=OUTPUT_FORMATS[0])
        self.cmb_format = ttk.Combobox(opts, textvariable=self.var_format,
                                       values=OUTPUT_FORMATS, state="readonly",
                                       width=8)
        self.cmb_format.grid(row=3, column=4, sticky="w", **pad)
        self.cmb_format.bind("<<ComboboxSelected>>", self._format_changed)

        self.var_halo_fix = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opts,
            text="Fix alpha edge halos (premultiplied warp)",
            variable=self.var_halo_fix,
        ).grid(row=4, column=0, columnspan=3, sticky="w", **pad)

        ttk.Label(opts, text="JPEG quality:").grid(row=4, column=3,
                                                   sticky="e", **pad)
        self.var_jpeg_q = tk.StringVar(value="90")
        self.spn_jpeg_q = ttk.Spinbox(opts, from_=1, to=100, increment=1,
                                      textvariable=self.var_jpeg_q, width=8,
                                      state="disabled")
        self.spn_jpeg_q.grid(row=4, column=4, sticky="w", **pad)

        self.var_preview = tk.BooleanVar(value=True)
        self.var_marks = tk.BooleanVar(value=True)
        prev_row = ttk.Frame(opts)
        prev_row.grid(row=5, column=0, columnspan=6, sticky="w", **pad)
        ttk.Checkbutton(prev_row, text="Live preview",
                        variable=self.var_preview).pack(side="left")
        ttk.Checkbutton(prev_row,
                        text="Show landmarks (green, preview only — "
                             "never saved)",
                        variable=self.var_marks).pack(side="left", padx=(16, 0))

        set_row = ttk.Frame(frm)
        set_row.grid(row=10, column=0, columnspan=3, sticky="w", padx=8)
        ttk.Button(set_row, text="Save settings…",
                   command=self.save_settings_as).pack(side="left")
        ttk.Button(set_row, text="Load settings…",
                   command=self.load_settings_from).pack(side="left",
                                                         padx=(8, 0))
        dnd_note = ("Drop images/folders anywhere to process them"
                    if HAS_DND else
                    "pip install tkinterdnd2 to enable drag && drop")
        ttk.Label(set_row, text=dnd_note,
                  foreground="grey").pack(side="left", padx=(16, 0))

        self.btn_run = ttk.Button(frm, text="Run", command=self.start)
        self.btn_run.grid(row=11, column=0, columnspan=3, sticky="ew",
                          padx=8, pady=(8, 8))

        # ---- Drop targets tab --------------------------------------------
        dnd_state = ("Drop files or folders on a tile to process them with "
                     "its preset (previews off for speed)."
                     if HAS_DND else
                     "pip install tkinterdnd2 to enable drop targets.")
        ttk.Label(
            targets_tab, justify="left",
            text=(f"{dnd_state}\nEach target links a settings preset "
                  "(re-read at every drop) to a required output folder."),
        ).pack(anchor="w", padx=8, pady=(8, 4))

        self.tiles_frame = ttk.Frame(targets_tab)
        self.tiles_frame.pack(fill="both", expand=True, padx=8, pady=4)
        self.tiles_frame.columnconfigure(0, weight=1)
        self.tiles_frame.columnconfigure(1, weight=1)

        tbtns = ttk.Frame(targets_tab)
        tbtns.pack(fill="x", padx=8, pady=(4, 8))
        ttk.Button(tbtns, text="Add target…",
                   command=self.add_drop_target).pack(side="left")

        # ---- Watched folders tab -----------------------------------------
        ttk.Label(
            watch_tab, justify="left",
            text=("Each watched folder needs a settings file "
                  f"({WATCH_SETTINGS_NAME}) — 'Add folder…' offers to create "
                  "one from the current Batch-tab settings.\n"
                  "While watching, new images are processed per that "
                  "folder's settings; originals move to "
                  f"{WATCH_DONE_DIR}\\ (or {WATCH_FAIL_DIR}\\).\n"
                  "Reference/backgrounds paths inside the settings file may "
                  "be relative to the folder."),
        ).pack(anchor="w", padx=8, pady=(8, 4))

        lst_frame = ttk.Frame(watch_tab)
        lst_frame.pack(fill="both", expand=True, padx=8, pady=4)
        self.lst_watch = tk.Listbox(lst_frame, height=6, activestyle="none")
        self.lst_watch.pack(side="left", fill="both", expand=True)
        wscroll = ttk.Scrollbar(lst_frame, command=self.lst_watch.yview)
        wscroll.pack(side="left", fill="y")
        self.lst_watch.configure(yscrollcommand=wscroll.set)

        wbtns = ttk.Frame(watch_tab)
        wbtns.pack(fill="x", padx=8, pady=(4, 8))
        ttk.Button(wbtns, text="Add folder…",
                   command=self.add_watch_folder).pack(side="left")
        ttk.Button(wbtns, text="Remove selected",
                   command=self.remove_watch_folder).pack(side="left",
                                                          padx=(8, 0))
        ttk.Button(wbtns, text="Write settings file",
                   command=self.write_watch_settings).pack(side="left",
                                                           padx=(8, 0))
        self.btn_watch = ttk.Button(wbtns, text="Start watching",
                                    command=self.toggle_watching)
        self.btn_watch.pack(side="right")
        self.lbl_watch = ttk.Label(watch_tab, text="Not watching.",
                                   foreground="grey")
        self.lbl_watch.pack(anchor="w", padx=8, pady=(0, 8))

        # ---- Shared progress + log below the tabs ------------------------
        bottom = ttk.Frame(self)
        bottom.pack(fill="both", expand=True, padx=10, pady=(4, 10))

        self.prog = ttk.Progressbar(bottom, mode="determinate")
        self.prog.pack(fill="x", pady=(0, 4))

        log_frame = ttk.Frame(bottom)
        log_frame.pack(fill="both", expand=True)
        self.log_box = tk.Text(log_frame, height=12, state="disabled",
                               font=("Consolas", 9))
        self.log_box.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(log_frame, command=self.log_box.yview)
        scroll.pack(side="left", fill="y")
        self.log_box.configure(yscrollcommand=scroll.set)

        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self.on_drop)

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.restore_last_settings()
        self.after(100, self.poll_queue)

    def _bg_mode_changed(self, _event=None):
        state = ("normal" if self.var_bg_mode.get() == "Custom scale %"
                 else "disabled")
        self.spn_bg_pct.configure(state=state)

    def _format_changed(self, _event=None):
        state = "normal" if self.var_format.get() == "JPEG" else "disabled"
        self.spn_jpeg_q.configure(state=state)

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
            # Picking a folder implies the user wants image backgrounds
            self.var_bg_source.set("Background images (folder)")

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

    def preview(self, kind, pil_img, caption):
        self.msg_q.put(("preview", (kind, pil_img, caption)))

    def show_preview(self, kind, pil_img, caption):
        titles = {"ref": "Reference (landmarks)", "result": "Result preview"}
        win = self.preview_windows.get(kind)
        if win is None or not win.winfo_exists():
            win = PreviewWindow(self, titles.get(kind, "Preview"))
            # Place preview windows to the right of the main window
            x = self.winfo_x() + self.winfo_width() + 12
            y = self.winfo_y() + (0 if kind == "result" else 80)
            win.geometry(f"+{x}+{y}")
            self.preview_windows[kind] = win
        win.show_image(pil_img, caption)

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
                    tile = self.tiles.get(
                        (self.current_job or {}).get("target_name"))
                    if tile is not None:
                        tile.set_progress(i, total)
                elif kind == "preview":
                    self.show_preview(*payload)
                elif kind == "done":
                    ok, fail = payload
                    job = self.current_job or {}
                    for f in job.get("files") or []:
                        self.watch_inflight.discard(str(f))
                    tile = self.tiles.get(job.get("target_name"))
                    if tile is not None:
                        tile.reset_progress()
                    self.current_job = None
                    self.busy = False
                    if not self.jobs:
                        self.btn_run.configure(state="normal", text="Run")
                    if (ok or fail) and not job.get("quiet"):
                        messagebox.showinfo(
                            "Finished",
                            f"{ok} image(s) processed, {fail} skipped."
                        )
                    self._next_job()
        except queue.Empty:
            pass
        self.after(100, self.poll_queue)

    # ------------------------------------------------------------------
    # Job queue
    # ------------------------------------------------------------------
    def enqueue_job(self, cfg, label):
        self.jobs.append(cfg)
        if self.busy:
            self.log(f"Queued: {label} ({len(self.jobs)} job(s) waiting)")
        self._next_job()

    def _next_job(self):
        if self.busy or not self.jobs:
            return
        cfg = self.jobs.popleft()
        self.busy = True
        self.current_job = cfg
        self.prog["value"] = 0
        self.btn_run.configure(state="disabled", text="Processing…")
        self.worker = threading.Thread(
            target=run_batch,
            args=(cfg, self.log, self.progress, self.done, self.preview),
            daemon=True,
        )
        self.worker.start()

    def start(self, files=None, quiet=False):
        cfg = self._build_cfg(files_mode=files is not None)
        if cfg is None:
            return
        if files is not None:
            cfg["files"] = [str(f) for f in files]
        cfg["quiet"] = quiet
        if not self.busy:
            self.log_box.configure(state="normal")
            self.log_box.delete("1.0", "end")
            self.log_box.configure(state="disabled")
        label = (f"{len(files)} dropped file(s)" if files is not None
                 else f"batch of {cfg['input']}")
        self.enqueue_job(cfg, label)

    def _build_cfg(self, files_mode=False):
        in_dir = self.var_input.get().strip()
        if not files_mode and (not in_dir or not Path(in_dir).is_dir()):
            messagebox.showerror("Missing input", "Please select a valid input folder.")
            return None
        ref = self.var_ref.get().strip()
        if ref and not Path(ref).is_file():
            messagebox.showerror("Bad reference", "Reference image not found.")
            return
        if self.var_match.get() and not ref:
            messagebox.showerror(
                "Reference required",
                "'Match reference resolution' needs a reference image — "
                "please select one (or untick the option).")
            return
        bgdir = self.var_bgdir.get().strip()
        if bgdir and not Path(bgdir).is_dir():
            messagebox.showerror("Bad backgrounds folder",
                                 "Backgrounds folder not found.")
            return
        if self.var_bg_source.get() == "Background images (folder)" and not bgdir:
            messagebox.showerror(
                "Backgrounds folder required",
                "'Background images' is selected as the transparent-source "
                "background — please pick a backgrounds folder (or choose "
                "Fill colour / Keep transparent).")
            return
        try:
            jpeg_q = int(self.var_jpeg_q.get())
            if not 1 <= jpeg_q <= 100:
                raise ValueError
        except ValueError:
            messagebox.showerror("Bad JPEG quality",
                                 "JPEG quality must be a number 1–100.")
            return
        try:
            size = int(self.var_size.get())
            if size < 50:
                raise ValueError
        except ValueError:
            messagebox.showerror("Bad size", "Output size must be a number ≥ 50.")
            return
        try:
            bg_pct = float(self.var_bg_pct.get())
            if not 10 <= bg_pct <= 400:
                raise ValueError
        except ValueError:
            messagebox.showerror("Bad BG scale",
                                 "BG custom % must be a number 10–400.")
            return

        return {
            "input": in_dir or None,
            "output": self.var_output.get().strip() or None,
            "reference": ref or None,
            "size": size,
            "suffix": self.var_suffix.get() or "_circle",
            "bg": self.bg_colour,
            "backgrounds": bgdir or None,
            "stable_bg": self.var_stable.get(),
            "center_x": self.var_center_x.get(),
            "scale_mode": self.var_scale_mode.get(),
            "match_ref": self.var_match.get(),
            "preview": self.var_preview.get(),
            "marks": self.var_marks.get(),
            "halo_fix": self.var_halo_fix.get(),
            "bg_mode": self.var_bg_mode.get(),
            "bg_pct": bg_pct,
            "bg_source": self.var_bg_source.get(),
            "format": self.var_format.get(),
            "jpeg_q": jpeg_q,
        }

    # ------------------------------------------------------------------
    # Settings persistence
    # ------------------------------------------------------------------
    def collect_settings(self):
        sd = dict(DEFAULT_SETTINGS)
        sd.update({
            "input": self.var_input.get().strip(),
            "output": self.var_output.get().strip(),
            "reference": self.var_ref.get().strip(),
            "backgrounds": self.var_bgdir.get().strip(),
            "suffix": self.var_suffix.get() or "_circle",
            "fill_colour": colour_to_hex(self.bg_colour),
            "stable_bg": self.var_stable.get(),
            "center_x": self.var_center_x.get(),
            "scale_mode": self.var_scale_mode.get(),
            "match_ref": self.var_match.get(),
            "preview": self.var_preview.get(),
            "marks": self.var_marks.get(),
            "halo_fix": self.var_halo_fix.get(),
            "bg_mode": self.var_bg_mode.get(),
            "bg_source": self.var_bg_source.get(),
            "format": self.var_format.get(),
        })
        for key, var, cast in (("size", self.var_size, int),
                               ("bg_pct", self.var_bg_pct, float),
                               ("jpeg_q", self.var_jpeg_q, int)):
            try:
                sd[key] = cast(var.get())
            except (TypeError, ValueError):
                pass
        return clean_settings(sd)

    def apply_settings(self, sd, base_dir=None):
        def rp(key):
            return resolve_path(sd[key], base_dir) or ""
        self.var_input.set(rp("input"))
        self.var_output.set(rp("output"))
        self.var_ref.set(rp("reference"))
        self.var_bgdir.set(rp("backgrounds"))
        self.var_size.set(str(sd["size"]))
        self.var_suffix.set(sd["suffix"])
        self.bg_colour = hex_to_colour(sd["fill_colour"])
        self.btn_colour.configure(bg=colour_to_hex(self.bg_colour))
        self.var_stable.set(sd["stable_bg"])
        self.var_center_x.set(sd["center_x"])
        self.var_scale_mode.set(sd["scale_mode"])
        self.var_match.set(sd["match_ref"])
        self.var_preview.set(sd["preview"])
        self.var_marks.set(sd["marks"])
        self.var_halo_fix.set(sd["halo_fix"])
        self.var_bg_mode.set(sd["bg_mode"])
        self.var_bg_pct.set(str(sd["bg_pct"]))
        self.var_bg_source.set(sd["bg_source"])
        self.var_format.set(sd["format"])
        self.var_jpeg_q.set(str(sd["jpeg_q"]))
        self._bg_mode_changed()
        self._format_changed()

    def save_settings_as(self):
        f = filedialog.asksaveasfilename(
            title="Save settings", defaultextension=".json",
            initialfile="headcropper_preset.json",
            filetypes=[("JSON settings", "*.json")])
        if not f:
            return
        try:
            save_settings_file(f, self.collect_settings(),
                               base_dir=Path(f).parent)
            self.log(f"Settings saved to {f}")
        except OSError as e:
            messagebox.showerror("Save failed", str(e))

    def load_settings_from(self):
        f = filedialog.askopenfilename(
            title="Load settings",
            filetypes=[("JSON settings", "*.json"), ("All files", "*.*")])
        if not f:
            return
        try:
            sd = load_settings_file(f)
        except (OSError, json.JSONDecodeError) as e:
            messagebox.showerror("Load failed", str(e))
            return
        self.apply_settings(sd, base_dir=Path(f).parent)
        self.log(f"Settings loaded from {f}")

    def _last_settings_path(self):
        return Path(__file__).resolve().parent / LAST_SETTINGS_NAME

    def restore_last_settings(self):
        path = self._last_settings_path()
        if not path.is_file():
            return
        try:
            with open(path, "r", encoding="utf-8") as fh:
                raw = json.load(fh)
        except (OSError, json.JSONDecodeError):
            return
        self.apply_settings(clean_settings(raw))
        for folder in raw.get("watched_folders", []):
            if isinstance(folder, str) and Path(folder).is_dir():
                self.lst_watch.insert("end", folder)
        for t in raw.get("drop_targets", []):
            if (isinstance(t, dict)
                    and all(isinstance(t.get(k), str) and t.get(k)
                            for k in ("name", "preset", "output"))
                    and t["name"] not in (x["name"]
                                          for x in self.drop_targets)):
                self.drop_targets.append(
                    {k: t[k] for k in ("name", "preset", "output")})
        if self.drop_targets:
            self._rebuild_tiles()

    def on_close(self):
        try:
            sd = self.collect_settings()
            sd["watched_folders"] = list(self.lst_watch.get(0, "end"))
            sd["drop_targets"] = self.drop_targets
            with open(self._last_settings_path(), "w", encoding="utf-8") as fh:
                json.dump(sd, fh, indent=2)
        except OSError:
            pass
        self.destroy()

    # ------------------------------------------------------------------
    # Drag & drop
    # ------------------------------------------------------------------
    def _files_from_drop(self, data):
        files = []
        # Tcl splitlist treats backslashes as escapes; tkdnd normally
        # sends forward slashes, but normalise so both survive.
        data = str(data).replace("\\", "/")
        for p in (Path(p) for p in self.tk.splitlist(data)):
            if p.is_dir():
                files.extend(sorted(
                    f for f in p.iterdir()
                    if f.suffix.lower() in VALID_EXT and f.is_file()))
            elif p.is_file() and p.suffix.lower() in VALID_EXT:
                files.append(p)
        return files

    def on_drop(self, event):
        files = self._files_from_drop(event.data)
        if not files:
            self.log("Drop ignored — no image files found.")
            return
        self.log(f"Dropped {len(files)} image(s) — processing with "
                 "current settings.")
        self.start(files=files, quiet=True)

    # ------------------------------------------------------------------
    # Drop targets
    # ------------------------------------------------------------------
    def add_drop_target(self):
        dlg = TargetDialog(self, [t["name"] for t in self.drop_targets])
        self.wait_window(dlg)
        if dlg.result:
            self.drop_targets.append(dlg.result)
            self._rebuild_tiles()
            self.log(f"Drop target '{dlg.result['name']}' created "
                     f"-> {dlg.result['output']}")

    def _remove_drop_target(self, target):
        self.drop_targets = [t for t in self.drop_targets if t is not target]
        self._rebuild_tiles()

    def _rebuild_tiles(self):
        for w in self.tiles_frame.winfo_children():
            w.destroy()
        self.tiles = {}
        for i, target in enumerate(self.drop_targets):
            tile = DropTile(
                self.tiles_frame, target,
                TILE_COLOURS[i % len(TILE_COLOURS)],
                on_remove=lambda t=target: self._remove_drop_target(t),
                on_drop=self.on_target_drop,
            )
            tile.grid(row=i // 2, column=i % 2, sticky="ew", padx=4, pady=4)
            self.tiles[target["name"]] = tile

    def on_target_drop(self, target, event):
        files = self._files_from_drop(event.data)
        if not files:
            self.log(f"[{target['name']}] drop ignored — no image files.")
            return
        try:
            sd = load_settings_file(target["preset"])
        except (OSError, json.JSONDecodeError) as e:
            self.log(f"[{target['name']}] preset unreadable "
                     f"({target['preset']}): {e}")
            return
        cfg = settings_to_cfg(sd, Path(target["preset"]).parent, "", self.log)
        cfg["input"] = None
        cfg["output"] = target["output"]
        cfg["files"] = [str(f) for f in files]
        cfg["preview"] = False      # drop targets are for speed — no preview
        cfg["quiet"] = True
        cfg["target_name"] = target["name"]
        self.log(f"[{target['name']}] {len(files)} image(s) dropped.")
        self.enqueue_job(cfg, f"drop target '{target['name']}'")

    # ------------------------------------------------------------------
    # Watched folders
    # ------------------------------------------------------------------
    def add_watch_folder(self):
        d = filedialog.askdirectory(title="Select folder to watch")
        if not d:
            return
        if d in self.lst_watch.get(0, "end"):
            messagebox.showinfo("Already watched",
                                "That folder is already in the list.")
            return
        settings_file = Path(d) / WATCH_SETTINGS_NAME
        if not settings_file.is_file():
            if messagebox.askyesno(
                    "Create settings file?",
                    f"No {WATCH_SETTINGS_NAME} in this folder.\n\n"
                    "Create one from the current Batch-tab settings?"):
                self._write_watch_settings_to(Path(d))
            else:
                messagebox.showinfo(
                    "Settings required",
                    "The folder will be listed, but it needs a "
                    f"{WATCH_SETTINGS_NAME} before it can be processed.")
        self.lst_watch.insert("end", d)

    def remove_watch_folder(self):
        sel = self.lst_watch.curselection()
        if sel:
            self.lst_watch.delete(sel[0])

    def write_watch_settings(self):
        sel = self.lst_watch.curselection()
        if not sel:
            messagebox.showinfo("No folder selected",
                                "Select a watched folder first.")
            return
        folder = Path(self.lst_watch.get(sel[0]))
        self._write_watch_settings_to(folder)

    def _write_watch_settings_to(self, folder):
        sd = self.collect_settings()
        sd["input"] = ""            # the watched folder itself provides input
        try:
            save_settings_file(folder / WATCH_SETTINGS_NAME, sd,
                               base_dir=folder)
            self.watch_bad_settings.discard(str(folder))
            self.log(f"Wrote {WATCH_SETTINGS_NAME} to {folder}")
        except OSError as e:
            messagebox.showerror("Write failed", str(e))

    def toggle_watching(self):
        if self.watching:
            self.watching = False
            self.btn_watch.configure(text="Start watching")
            self.lbl_watch.configure(text="Not watching.")
            self.log("Watching stopped.")
            return
        folders = list(self.lst_watch.get(0, "end"))
        if not folders:
            messagebox.showinfo("Nothing to watch",
                                "Add at least one folder first.")
            return
        self.watching = True
        self.watch_seen.clear()
        self.btn_watch.configure(text="Stop watching")
        self.lbl_watch.configure(
            text=f"Watching {len(folders)} folder(s), polling every "
                 f"{WATCH_POLL_MS // 1000}s…")
        self.log(f"Watching {len(folders)} folder(s).")
        self.after(WATCH_POLL_MS, self._watch_tick)

    def _watch_tick(self):
        if not self.watching:
            return
        for folder in self.lst_watch.get(0, "end"):
            try:
                self._scan_watch_folder(Path(folder))
            except OSError as e:
                self.log(f"Watch error in {folder}: {e}")
        # Forget entries for files that vanished (moved/renamed/deleted)
        self.watch_seen = {k: v for k, v in self.watch_seen.items()
                           if Path(k).exists()}
        self.after(WATCH_POLL_MS, self._watch_tick)

    def _scan_watch_folder(self, folder):
        if not folder.is_dir():
            return
        ready = []
        for f in sorted(folder.iterdir()):
            if not (f.is_file() and f.suffix.lower() in VALID_EXT):
                continue
            key = str(f)
            if key in self.watch_inflight:
                continue
            try:
                size = f.stat().st_size
            except OSError:
                continue
            if self.watch_seen.get(key) == size and size > 0:
                ready.append(f)        # size stable across two polls
            else:
                self.watch_seen[key] = size
        if not ready:
            return
        settings_file = folder / WATCH_SETTINGS_NAME
        if not settings_file.is_file():
            if str(folder) not in self.watch_bad_settings:
                self.watch_bad_settings.add(str(folder))
                self.log(f"Watch: {folder} has no {WATCH_SETTINGS_NAME} — "
                         "images are waiting but will not be processed.")
            return
        try:
            sd = load_settings_file(settings_file)
        except (OSError, json.JSONDecodeError) as e:
            if str(folder) not in self.watch_bad_settings:
                self.watch_bad_settings.add(str(folder))
                self.log(f"Watch: bad settings file in {folder} ({e})")
            return
        self.watch_bad_settings.discard(str(folder))
        cfg = settings_to_cfg(sd, folder, folder, self.log)
        cfg["files"] = [str(f) for f in ready]
        cfg["output"] = cfg["output"] or str(folder / "circled")
        cfg["move_originals"] = True
        cfg["quiet"] = True
        for f in ready:
            self.watch_inflight.add(str(f))
        self.log(f"Watch: {len(ready)} new image(s) in {folder.name}")
        self.enqueue_job(cfg, f"watch job ({folder.name})")


if __name__ == "__main__":
    App().mainloop()
