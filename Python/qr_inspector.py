#!/usr/bin/env python3
"""
QR Inspector – Load a QR image, decode contents, and display helpful metadata.

Features
- Opens an image containing 1+ QR codes
- Decodes and shows the contents in a copyable field
- Displays metadata (decoder, format, error correction, byte length, codec guess, etc.)
- Handles multiple codes (prev/next)
- Estimates QR min version and matrix size from payload + EC level

Author: You + ChatGPT
"""

from __future__ import annotations

import importlib
import subprocess
import sys
import traceback
from typing import Any, Dict, List, Optional, Tuple


# =========================
# Imports (manual deps; no auto-install)
# =========================

from PIL import Image, ImageTk  # pillow
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import chardet

try:
    import zxingcpp  # preferred decoder
    ZXING_AVAILABLE = True
except Exception as e:
    zxingcpp = None  # type: ignore
    ZXING_AVAILABLE = False
    ZXING_IMPORT_ERR = e

# =========================
# Utilities
# =========================

def safe_text_from_bytes(raw: Optional[bytes]) -> Tuple[str, str]:
    """
    Robustly turn raw bytes into text and report the codec used.
    Try UTF‑8 first, then fall back to chardet's best guess, then replace.
    """
    if raw is None:
        return "", "n/a"
    try:
        return raw.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        guess = chardet.detect(raw) or {}
        enc = guess.get("encoding") or "binary"
        try:
            return raw.decode(enc, errors="replace"), enc
        except Exception:
            return raw.decode("utf-8", errors="replace"), "utf-8 (replace)"


def guess_payload_kind(text: str) -> str:
    """
    Lightweight classifier to give a friendly sense of 'what' the QR holds.
    """
    t = (text or "").strip()
    tl = t.lower()

    if tl.startswith(("http://", "https://")):
        return "URL"
    if tl.startswith("mailto:"):
        return "Email link"
    if tl.startswith("tel:"):
        return "Telephone"
    if tl.startswith("sms:"):
        return "SMS link"
    if tl.startswith("geo:"):
        return "Geo coordinates"
    if tl.startswith("wifi:"):
        return "Wi‑Fi config"
    if tl.startswith(("bitcoin:", "ethereum:", "web+")):
        return "Payment / App link"
    if tl.startswith("begin:vcard"):
        return "vCard contact"
    if "BEGIN:VEVENT" in t or tl.startswith("begin:vevent"):
        return "Calendar event"
    if ("@" in t) and (" " not in t) and ("." in t):
        return "Likely email address"
    return "Plain text"


def ensure_rgb(img: Image.Image) -> Image.Image:
    """Ensure image is RGB/RGBA for Tk display and decoders."""
    if img.mode not in ("RGB", "RGBA"):
        return img.convert("RGB")
    return img


# =========================
# QR capacity tables & estimation (version/matrix)
# =========================

# Allowed chars for QR "alphanumeric" mode
_ALNUM_TABLE = set("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:")

# Capacity tables for versions 1..40 (index 0 -> version 1)
# Each entry is a tuple (L, M, Q, H)

# Numeric:
_CAP_NUM = [
    (41, 34, 27, 17), (77, 63, 48, 34), (127, 101, 77, 58), (187, 149, 111, 82),
    (255, 202, 144, 106), (322, 255, 178, 139), (370, 293, 207, 154), (461, 365, 259, 202),
    (552, 432, 312, 235), (652, 513, 364, 288), (772, 604, 427, 331), (883, 691, 489, 374),
    (1022, 796, 580, 427), (1101, 871, 621, 468), (1250, 991, 703, 530), (1408, 1082, 775, 602),
    (1548, 1212, 876, 674), (1725, 1346, 948, 746), (1903, 1500, 1063, 813), (2061, 1600, 1159, 919),
    (2232, 1708, 1224, 969), (2409, 1872, 1358, 1056), (2620, 2059, 1468, 1108), (2812, 2188, 1588, 1228),
    (3057, 2395, 1718, 1286), (3283, 2544, 1804, 1425), (3517, 2701, 1933, 1501), (3669, 2857, 2085, 1581),
    (3909, 3035, 2182, 1677), (4158, 3289, 2358, 1782), (4417, 3486, 2473, 1897), (4686, 3693, 2670, 2022),
    (4965, 3909, 2805, 2157), (5253, 4134, 2949, 2301), (5529, 4343, 3081, 2361), (5836, 4588, 3244, 2524),
    (6153, 4775, 3417, 2625), (6479, 5039, 3599, 2735), (6743, 5313, 3791, 2927), (7089, 5596, 3993, 3057),
]

# Alphanumeric:
_CAP_ALNUM = [
    (25, 20, 16, 10), (47, 38, 29, 20), (77, 61, 47, 35), (114, 90, 67, 50),
    (154, 122, 87, 64), (195, 154, 108, 84), (224, 178, 125, 93), (279, 221, 157, 122),
    (335, 262, 189, 143), (395, 311, 221, 174), (468, 366, 259, 200), (535, 419, 296, 227),
    (619, 483, 352, 259), (667, 528, 376, 283), (758, 600, 426, 321), (854, 656, 470, 365),
    (938, 734, 531, 408), (1046, 816, 574, 452), (1153, 909, 644, 493), (1249, 970, 702, 557),
    (1352, 1035, 742, 587), (1460, 1134, 823, 640), (1588, 1248, 890, 672), (1704, 1326, 963, 744),
    (1853, 1451, 1041, 779), (1990, 1542, 1094, 864), (2132, 1637, 1172, 910), (2223, 1732, 1263, 958),
    (2369, 1839, 1322, 1016), (2520, 1994, 1429, 1080), (2677, 2113, 1499, 1150), (2840, 2238, 1618, 1226),
    (3009, 2369, 1700, 1307), (3183, 2506, 1787, 1394), (3351, 2632, 1867, 1431), (3537, 2780, 1966, 1530),
    (3729, 2894, 2071, 1591), (3927, 3054, 2181, 1658), (4087, 3220, 2298, 1774), (4296, 3391, 2420, 1852),
]

# Byte:
_CAP_BYTE = [
    (17, 14, 11, 7), (32, 26, 20, 14), (53, 42, 32, 24), (78, 62, 46, 34),
    (106, 84, 60, 44), (134, 106, 74, 58), (154, 122, 86, 64), (192, 152, 108, 84),
    (230, 180, 130, 98), (271, 213, 151, 119), (321, 251, 177, 137), (367, 287, 203, 155),
    (425, 331, 241, 177), (458, 362, 258, 194), (520, 412, 292, 220), (586, 450, 322, 250),
    (644, 504, 364, 280), (718, 560, 394, 310), (792, 624, 442, 338), (858, 666, 482, 382),
    (929, 711, 509, 403), (1003, 779, 565, 439), (1091, 857, 611, 461), (1171, 911, 661, 511),
    (1273, 997, 715, 535), (1367, 1059, 751, 593), (1465, 1125, 805, 625), (1528, 1190, 868, 658),
    (1628, 1264, 908, 698), (1732, 1370, 982, 742), (1840, 1452, 1030, 790), (1952, 1538, 1112, 842),
    (2068, 1628, 1168, 898), (2188, 1722, 1228, 958), (2303, 1809, 1283, 983), (2431, 1911, 1351, 1051),
    (2563, 1989, 1423, 1093), (2699, 2099, 1499, 1139), (2809, 2213, 1579, 1219), (2953, 2331, 1663, 1273),
]


def _cap_for(version: int, ec: str, mode: str) -> int:
    idx = version - 1
    ec_index = {"L": 0, "M": 1, "Q": 2, "H": 3}.get(ec, 1)
    table = {"numeric": _CAP_NUM, "alphanumeric": _CAP_ALNUM, "byte": _CAP_BYTE}[mode]
    return table[idx][ec_index]


def estimate_version_lower_bound(payload_text: str, payload_bytes_len: int, ec: str) -> Tuple[int, str]:
    """
    Return (min_version, mode_used) based on the smallest version that can hold
    the payload under the given error-correction level.

    We choose the 'most optimistic' plausible mode (numeric > alnum > byte).
    """
    # Try best-case: numeric, then alphanumeric, else byte
    candidates: List[Tuple[int, str]] = []

    # numeric
    if payload_text.isdigit():
        need = len(payload_text)
        for v in range(1, 41):
            if _cap_for(v, ec, "numeric") >= need:
                candidates.append((v, "numeric"))
                break

    # alphanumeric
    if all(ch in _ALNUM_TABLE for ch in payload_text):
        need = len(payload_text)
        for v in range(1, 41):
            if _cap_for(v, ec, "alphanumeric") >= need:
                candidates.append((v, "alphanumeric"))
                break

    # byte (worst-case; capacity table is in bytes)
    need = payload_bytes_len
    for v in range(1, 41):
        if _cap_for(v, ec, "byte") >= need:
            candidates.append((v, "byte"))
            break

    if not candidates:
        # beyond max ⇒ 40+ (very unlikely for typical use)
        return 40, "byte"

    v, m = min(candidates, key=lambda t: t[0])
    return v, m


def matrix_size_for_version(version: int) -> int:
    """QR matrix size in modules: 21 + 4*(version - 1)."""
    return 21 + 4 * (version - 1)


# =========================
# Decoding backends
# =========================

def decode_with_zxing(img: Image.Image) -> List[Dict[str, Any]]:
    """
    Use zxing-cpp to decode. Returns a list of dicts with standardised fields.
    """
    if zxingcpp is None:
        return []

    results: List[Dict[str, Any]] = []
    pil_img = ensure_rgb(img)
    try:
        decoded = zxingcpp.read_barcodes(pil_img)  # type: ignore[attr-defined]
    except Exception:
        # If it fails, let fallback run.
        return []

    for r in decoded:
        text = getattr(r, "text", "") or ""
        raw_bytes = getattr(r, "bytes", b"") or b""
        format_name = str(getattr(r, "format", "QR_CODE"))
        ec_level = getattr(r, "ec_level", None)  # 'L', 'M', 'Q', 'H'
        symbology = getattr(r, "symbology_identifier", None)
        position = getattr(r, "position", None)

        # Prefer raw bytes if present; otherwise encode text to bytes for length & decoding report
        basis = raw_bytes if raw_bytes else text.encode("utf-8")
        decoded_text, codec = safe_text_from_bytes(basis)
        payload_kind = guess_payload_kind(decoded_text)

        meta: Dict[str, Any] = {
            "Decoder": "ZXing‑C++",
            "Format": format_name,
            "Error correction": ec_level or "unknown",
            "Symbology ID": symbology or "n/a",
            "Bytes (len)": len(basis),
            "Codec": codec,
            "Payload kind (guess)": payload_kind,
        }

        # Structured append (if available)
        sa_seq = getattr(r, "sequence_index", None)
        sa_total = getattr(r, "sequence_size", None)
        if sa_seq is not None and sa_total is not None and sa_total > 1:
            meta["Structured append"] = f"{sa_seq + 1} of {sa_total}"

        if position and hasattr(position, "top_left"):
            meta["Position points"] = "yes"

        # --- Version & matrix estimation ---
        try:
            ec = (meta.get("Error correction") or "M").strip()[0]  # 'L'/'M'/'Q'/'H'
            min_ver, mode_used = estimate_version_lower_bound(decoded_text, meta["Bytes (len)"], ec)
            meta["Estimated min version"] = f"v{min_ver} (mode: {mode_used})"
            meta["Matrix size (min)"] = f"{matrix_size_for_version(min_ver)}×{matrix_size_for_version(min_ver)} modules"
        except Exception:
            pass

        # Mask pattern (ZXing Python bindings may not expose; try getattr, else note)
        mask = getattr(r, "mask_pattern", None)
        if mask is not None:
            meta["Mask pattern"] = str(mask)
        else:
            meta.setdefault("Mask pattern", "not exposed by decoder")

        results.append({"text": decoded_text, "raw_bytes": raw_bytes, "metadata": meta})

    return results



def decode_image(img: Image.Image) -> List[Dict[str, Any]]:
    """
    Decode with ZXing only. If ZXing isn't available, raise an instructive error.
    """
    if not ZXING_AVAILABLE:
        raise RuntimeError(
            "ZXing backend not available. Install dependencies: 'pip install zxing-cpp pillow chardet'"
        )
    res = decode_with_zxing(img)
    if res:
        return res
    raise RuntimeError("No QR/barcodes detected, or decoding failed.")

    if zxingcpp is not None:
        try:
            res = decode_with_zxing(img)
            if res:
                return res
        except Exception as exc:  # pragma: no cover
            backend_errors.append(("ZXing", repr(exc)))

    if pyzbar is not None:
        try:
            res = decode_with_pyzbar(img)
            if res:
                return res
        except Exception as exc:  # pragma: no cover
            backend_errors.append(("pyzbar", repr(exc)))

    if (zxingcpp is None) and (pyzbar is None):
        raise RuntimeError(
            "No decoder backends available. Please ensure at least one of:\n"
            "  pip install zxing-cpp\n"
            "  pip install pyzbar   (may also require system ZBar)"
        )
    raise RuntimeError(
        "No QR/barcodes detected, or decoding failed."
        + (f" Details: {backend_errors}" if backend_errors else "")
    )


# =========================
# GUI
# =========================

class QRInspectorApp(tk.Tk):
    """Main window for QR Inspector."""

    def __init__(self) -> None:
        super().__init__()
        self.title("QR Inspector")
        self.minsize(820, 560)
        self.configure(padx=10, pady=10)

        # State
        self.current_image: Optional[Image.Image] = None
        self.current_photo: Optional[ImageTk.PhotoImage] = None
        self.results: List[Dict[str, Any]] = []
        self.current_index: int = 0

        # Layout grid
        self.columnconfigure(0, weight=1, uniform="cols")
        self.columnconfigure(1, weight=1, uniform="cols")
        self.rowconfigure(2, weight=1)

        # Controls row
        ctl = ttk.Frame(self)
        ctl.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ctl.columnconfigure(2, weight=1)

        ttk.Button(ctl, text="Open Image…", command=self.open_image_dialog).grid(row=0, column=0, padx=(0, 8))

        self.index_label_var = tk.StringVar(value="No codes detected")
        ttk.Label(ctl, textvariable=self.index_label_var).grid(row=0, column=1)

        self.backend_label_var = tk.StringVar(value="")
        ttk.Label(ctl, textvariable=self.backend_label_var, foreground="#666").grid(row=0, column=2, sticky="e")

        # Image preview
        img_frame = ttk.LabelFrame(self, text="Image preview")
        img_frame.grid(row=2, column=0, sticky="nsew", padx=(0, 8))
        img_frame.rowconfigure(0, weight=1)
        img_frame.columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(img_frame, background="#fafafa", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        # Right panel (tabs)
        right = ttk.Notebook(self)
        right.grid(row=2, column=1, sticky="nsew")

        # Tab: Contents
        self.content_tab = ttk.Frame(right)
        right.add(self.content_tab, text="Contents")
        self.content_tab.rowconfigure(2, weight=1)
        self.content_tab.columnconfigure(0, weight=1)

        self.content_text = tk.Text(self.content_tab, height=6, wrap="word")
        self.content_text.insert("1.0", "Decoded content will appear here…")
        self.content_text.configure(state="disabled")
        self.content_text.grid(row=0, column=0, sticky="nsew")

        btns = ttk.Frame(self.content_tab)
        btns.grid(row=1, column=0, sticky="ew", pady=6)
        ttk.Button(btns, text="Copy", command=self.copy_contents).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(btns, text="◀ Previous code", command=self.prev_code).grid(row=0, column=1, padx=6)
        ttk.Button(btns, text="Next code ▶", command=self.next_code).grid(row=0, column=2, padx=6)

        # Tab: Metadata
        self.meta_tab = ttk.Frame(right)
        right.add(self.meta_tab, text="Metadata")
        self.meta_tab.rowconfigure(0, weight=1)
        self.meta_tab.columnconfigure(0, weight=1)

        self.meta_tree = ttk.Treeview(self.meta_tab, columns=("key", "value"), show="headings", height=8)
        self.meta_tree.heading("key", text="Field")
        self.meta_tree.heading("value", text="Value")
        self.meta_tree.column("key", width=210, anchor="w")
        self.meta_tree.column("value", anchor="w")
        self.meta_tree.grid(row=0, column=0, sticky="nsew")

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status_var, anchor="w").grid(
            row=3, column=0, columnspan=2, sticky="ew", pady=(8, 0)
        )

        self._tweak_style()
        self.bind("<Configure>", lambda _e: self._update_preview())

        # Show which backend(s) are available in the status
        backends = []
        if ZXING_AVAILABLE:
            backends.append("ZXing‑C++")
        if not backends:
            backends = ["(none)"]
        self.status_var.set(f"Ready • Backends: {', '.join(backends)}")}")

    # ---------- UI niceties ----------

    def _tweak_style(self) -> None:
        try:
            style = ttk.Style(self)
            style.configure("TButton", padding=(8, 6))
            style.configure("Treeview", rowheight=22)
        except Exception:
            pass

    # ---------- Actions ----------

    def open_image_dialog(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose an image",
            filetypes=[
                ("Images", "*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.tif;*.tiff"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        self.load_and_decode(path)

    def load_and_decode(self, path: str) -> None:
        try:
            with Image.open(path) as im:
                im.load()
                self.current_image = im.copy()
        except Exception as exc:
            messagebox.showerror("Open image failed", f"Could not open image:\n{exc}")
            return

        self._update_preview()
        self.status_var.set("Decoding…")
        self.update_idletasks()

        try:
            self.results = decode_image(self.current_image)
            self.current_index = 0
            self._show_result()
            backend = self.results[0]["metadata"].get("Decoder", "Unknown")
            self.backend_label_var.set(f"Decoder: {backend}")
            self.status_var.set("Decoded successfully.")
        except Exception as exc:
            self.results = []
            self.current_index = 0
            self.index_label_var.set("No codes detected")
            self.backend_label_var.set("")
            self._set_content_text("No QR/barcode detected in this image.")
            self._populate_metadata({})
            self.status_var.set(f"Decode error: {exc}")

    def _update_preview(self) -> None:
        if not self.current_image:
            self._clear_canvas()
            return

        self._clear_canvas()
        cw = self.canvas.winfo_width() or 10
        ch = self.canvas.winfo_height() or 10
        iw, ih = self.current_image.size

        # Fit image within canvas while preserving aspect ratio
        scale = min((cw - 10) / iw, (ch - 10) / ih, 1.0)
        scale = 1.0 if scale <= 0 else scale
        disp = self.current_image if scale == 1.0 else self.current_image.resize(
            (max(1, int(iw * scale)), max(1, int(ih * scale))),
            Image.LANCZOS,
        )
        self.current_photo = ImageTk.PhotoImage(ensure_rgb(disp))
        self.canvas.create_image(cw // 2, ch // 2, image=self.current_photo, anchor="center")

    def _clear_canvas(self) -> None:
        self.canvas.delete("all")

    def _set_content_text(self, text: str) -> None:
        self.content_text.configure(state="normal")
        self.content_text.delete("1.0", "end")
        self.content_text.insert("1.0", text if text else "—")
        self.content_text.configure(state="disabled")

    def _populate_metadata(self, meta: Dict[str, Any]) -> None:
        for item in self.meta_tree.get_children():
            self.meta_tree.delete(item)
        # Keep a stable, readable order
        preferred_order = [
            "Decoder",
            "Format",
            "Error correction",
            "Symbology ID",
            "Bytes (len)",
            "Codec",
            "Payload kind (guess)",
            "Estimated min version",
            "Matrix size (min)",
            "Mask pattern",
            "Structured append",
            "Position points",
        ]
        used = set()
        for key in preferred_order:
            if key in meta:
                self.meta_tree.insert("", "end", values=(key, meta[key]))
                used.add(key)
        for key, value in meta.items():
            if key not in used:
                self.meta_tree.insert("", "end", values=(key, value))

    def _show_result(self) -> None:
        if not self.results:
            self.index_label_var.set("No codes detected")
            self._set_content_text("")
            self._populate_metadata({})
            return

        idx = self.current_index
        total = len(self.results)
        self.index_label_var.set(f"Code {idx + 1} of {total}")
        res = self.results[idx]
        self._set_content_text(res.get("text", ""))
        self._populate_metadata(res.get("metadata", {}))

    def copy_contents(self) -> None:
        if not self.results:
            return
        text = self.results[self.current_index].get("text", "")
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status_var.set("Copied to clipboard.")

    def next_code(self) -> None:
        if not self.results:
            return
        self.current_index = (self.current_index + 1) % len(self.results)
        self._show_result()

    def prev_code(self) -> None:
        if not self.results:
            return
        self.current_index = (self.current_index - 1) % len(self.results)
        self._show_result()


def main() -> None:
    try:
        app = QRInspectorApp()
        app.mainloop()
    except Exception as exc:
        print("Fatal error:", exc, file=sys.stderr)
        traceback.print_exc()


if __name__ == "__main__":
    main()
