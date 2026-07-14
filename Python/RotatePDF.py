#!/usr/bin/env python3
"""
RotatePDF.py — v1.0
True PDF rotation for print workflows (Windows GUI, tkinter + pikepdf).

WHY THIS EXISTS
    Acrobat's rotate sets the page's /Rotate flag — a VIEWING instruction.
    Many RIPs, imposition and placement tools ignore it, so the file arrives
    "unrotated". This tool bakes rotation into the page CONTENT:

        1. Wraps each page's content stream in  q <matrix> cm ... Q
        2. Transforms ALL page boxes (Media/Crop/Bleed/Trim/Art) through the
           same matrix, normalising the MediaBox origin to (0,0)
        3. Sets /Rotate 0

    Only geometry changes. Fonts, ICC profiles, spot colours, overprint and
    metadata are untouched — unlike an Illustrator open/rotate/re-save.

MODES
    90 CW / 180 / 90 CCW  — visual rotation, same direction as Acrobat's
    0 (normalise only)     — bakes any EXISTING /Rotate flags into content
                             and zeroes the flag. This alone fixes files
                             already rotated in Acrobat.
    "Bake existing /Rotate flags" (default ON) folds per-page flags into the
    chosen rotation, so mixed-flag documents come out visually identical to
    what Acrobat showed, with every flag at 0.

REQUIREMENTS
    Python 3.8+ and:   pip install pikepdf
    (tkinter ships with the standard Windows Python installer.)

NOTES FOR PRINT
    - Never overwrites the source; writes alongside with a suffix (or to a
      chosen folder). Optional overwrite of an existing OUTPUT file.
    - Annotations (form fields, some printer marks) live outside the content
      stream and are NOT rotated; files containing them are flagged in the
      report so you can check.
    - Encrypted PDFs are skipped and reported.
"""

import os
import sys
import queue
import threading
import traceback

try:
    import pikepdf
except ImportError:
    pikepdf = None

APP_TITLE = "Rotate PDF (true rotation) — v1.0"
SUFFIX_DEFAULT = "_rotated"

# Clockwise visual rotation -> content matrix linear part (a, b, c, d).
# PDF cm maps (x,y) -> (a*x + c*y + e,  b*x + d*y + f).
_CW_LINEAR = {
    90:  (0.0, -1.0, 1.0, 0.0),
    180: (-1.0, 0.0, 0.0, -1.0),
    270: (0.0, 1.0, -1.0, 0.0),
}

_BOX_KEYS = ("/MediaBox", "/CropBox", "/BleedBox", "/TrimBox", "/ArtBox")


# ---------------------------------------------------------------------------
# Engine
# ---------------------------------------------------------------------------

def _fmt(n):
    """Compact PDF number formatting."""
    s = ("%.4f" % n).rstrip("0").rstrip(".")
    return s if s not in ("", "-0") else "0"


def _box_floats(arr):
    x0, y0, x1, y1 = (float(v) for v in arr)
    # normalise ordering, some producers store boxes reversed
    return (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))


def _transform_box(box, a, b, c, d, e, f):
    x0, y0, x1, y1 = box
    pts = ((x0, y0), (x1, y0), (x0, y1), (x1, y1))
    xs = [a * x + c * y + e for (x, y) in pts]
    ys = [b * x + d * y + f for (x, y) in pts]
    return (min(xs), min(ys), max(xs), max(ys))


_INHERITABLE = ("/MediaBox", "/CropBox")


def _find_inherited(page, key):
    """Raw box if genuinely present on the page or an ancestor; else None.
    Never falls back to another box (pikepdf's .cropbox falls back to the
    MediaBox, which creates phantom CropBoxes — the v-test bug)."""
    node = page
    for _ in range(64):
        if node is None:
            return None
        try:
            if key in node:
                return node[key]
        except Exception:
            return None
        try:
            node = node.get("/Parent")
        except Exception:
            return None
    return None


def _bake_rotation(pdf, page, cw):
    """Rotate one page's content + boxes by cw degrees clockwise (90/180/270)."""
    a, b, c, d = _CW_LINEAR[cw]

    # Snapshot every genuinely-present box BEFORE touching anything —
    # writing the new MediaBox first and reading CropBox after was the
    # double-transform bug.
    originals = {}
    for key in _BOX_KEYS:
        raw = _find_inherited(page, key) if key in _INHERITABLE else page.get(key)
        if raw is None:
            continue
        try:
            originals[key] = _box_floats(raw)
        except Exception:
            pass

    media = originals.get("/MediaBox")
    if media is None:
        raise ValueError("page has no resolvable MediaBox")

    # translation so the rotated MediaBox lands with lower-left at (0,0)
    rx0, ry0, _, _ = _transform_box(media, a, b, c, d, 0.0, 0.0)
    e, f = -rx0, -ry0

    # content first (matrix references ORIGINAL coordinates)
    pre = ("q %s %s %s %s %s %s cm\n" % tuple(_fmt(v) for v in (a, b, c, d, e, f))).encode("ascii")
    _wrap_contents(pdf, page, pre, b"\nQ\n")

    # then every box that actually existed, through the same affine
    for key, box in originals.items():
        nb = _transform_box(box, a, b, c, d, e, f)
        page[key] = pikepdf.Array([nb[0], nb[1], nb[2], nb[3]])


def _wrap_contents(pdf, page, pre_bytes, post_bytes):
    """Prepend/append content streams around the page's existing content."""
    pre = pikepdf.Stream(pdf, pre_bytes)
    post = pikepdf.Stream(pdf, post_bytes)
    if hasattr(page, "contents_add"):
        page.contents_add(pre, prepend=True)
        page.contents_add(post, prepend=False)
        return
    # Fallback for older pikepdf: build a contents array by hand
    existing = page.get("/Contents")
    items = []
    if existing is None:
        pass
    elif isinstance(existing, pikepdf.Array):
        items = list(existing)
    else:
        items = [existing]
    page.Contents = pikepdf.Array([pre] + items + [post])




def rotate_pdf_file(src_path, dst_path, user_cw=90, bake_existing=True):
    """
    Rotate a PDF with true content baking.

    user_cw       0 / 90 / 180 / 270 (visual clockwise, Acrobat convention)
    bake_existing fold each page's existing /Rotate flag into the rotation

    Returns dict: pages, rotated_pages, flags_found (sorted list),
                  annot_pages (count of pages carrying annotations).
    Raises pikepdf.PasswordError for encrypted files, ValueError for others.
    """
    if pikepdf is None:
        raise RuntimeError("pikepdf is not installed.  pip install pikepdf")
    if user_cw % 90 != 0:
        raise ValueError("Rotation must be a multiple of 90.")
    user_cw %= 360

    flags_found = set()
    rotated_pages = 0
    annot_pages = 0
    n_pages = 0

    with pikepdf.open(src_path) as pdf:
        n_pages = len(pdf.pages)
        for page in pdf.pages:
            try:
                existing = int(page.get("/Rotate", 0)) % 360
            except Exception:
                existing = 0
            if existing:
                flags_found.add(existing)

            try:
                annots = page.get("/Annots")
                if annots is not None and len(annots) > 0:
                    annot_pages += 1
            except Exception:
                pass

            total = ((existing if bake_existing else 0) + user_cw) % 360
            if total:
                _bake_rotation(pdf, page, total)
                rotated_pages += 1
            # flag is baked (or intentionally untouched content): zero it out
            if bake_existing or user_cw:
                page.Rotate = 0

        pdf.save(dst_path)

    return {
        "pages": n_pages,
        "rotated_pages": rotated_pages,
        "flags_found": sorted(flags_found),
        "annot_pages": annot_pages,
    }


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

def main():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    if pikepdf is None:
        # a bare Tk error is friendlier than a console traceback on Windows
        root = tk.Tk(); root.withdraw()
        messagebox.showerror(APP_TITLE, "pikepdf is not installed.\n\nOpen a command prompt and run:\n\n    pip install pikepdf")
        return

    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("860x560")
    root.minsize(760, 480)

    files = []          # absolute paths, in display order
    msg_q = queue.Queue()
    working = {"flag": False}

    # ---------- top: file list ----------
    frmTop = ttk.Frame(root, padding=8)
    frmTop.pack(fill="both", expand=True)

    cols = ("pages", "flags", "status")
    tree = ttk.Treeview(frmTop, columns=cols, show="tree headings", selectmode="extended")
    tree.heading("#0", text="File")
    tree.heading("pages", text="Pages")
    tree.heading("flags", text="/Rotate flags")
    tree.heading("status", text="Status")
    tree.column("#0", width=380, anchor="w")
    tree.column("pages", width=60, anchor="center")
    tree.column("flags", width=110, anchor="center")
    tree.column("status", width=220, anchor="w")
    tree.pack(side="left", fill="both", expand=True)
    sb = ttk.Scrollbar(frmTop, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    sb.pack(side="left", fill="y")

    frmBtns = ttk.Frame(frmTop, padding=(8, 0, 0, 0))
    frmBtns.pack(side="left", fill="y")

    def probe(path):
        """pages + existing flags summary, without keeping the file open."""
        try:
            with pikepdf.open(path) as pdf:
                n = len(pdf.pages)
                flags = set()
                for pg in pdf.pages:
                    try:
                        r = int(pg.get("/Rotate", 0)) % 360
                    except Exception:
                        r = 0
                    if r:
                        flags.add(r)
                return n, ("none" if not flags else ",".join(str(x) for x in sorted(flags)))
        except pikepdf.PasswordError:
            return None, "ENCRYPTED"
        except Exception:
            return None, "unreadable"

    def add_files():
        paths = filedialog.askopenfilenames(title="Add PDF files", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        for p in paths:
            p = os.path.abspath(p)
            if p in files:
                continue
            n, flags = probe(p)
            files.append(p)
            tree.insert("", "end", iid=p, text=os.path.basename(p),
                        values=(n if n is not None else "?", flags, ""))

    def remove_sel():
        for iid in tree.selection():
            if iid in files:
                files.remove(iid)
            tree.delete(iid)

    def clear_all():
        files[:] = []
        for iid in tree.get_children():
            tree.delete(iid)

    ttk.Button(frmBtns, text="Add Files\u2026", command=add_files).pack(fill="x", pady=2)
    ttk.Button(frmBtns, text="Remove Selected", command=remove_sel).pack(fill="x", pady=2)
    ttk.Button(frmBtns, text="Clear", command=clear_all).pack(fill="x", pady=2)

    # ---------- middle: options ----------
    frmOpt = ttk.LabelFrame(root, text="Rotation", padding=8)
    frmOpt.pack(fill="x", padx=8, pady=(0, 4))

    rot = tk.IntVar(value=90)
    ttk.Radiobutton(frmOpt, text="90\u00b0 CW", variable=rot, value=90).grid(row=0, column=0, padx=6, sticky="w")
    ttk.Radiobutton(frmOpt, text="180\u00b0", variable=rot, value=180).grid(row=0, column=1, padx=6, sticky="w")
    ttk.Radiobutton(frmOpt, text="90\u00b0 CCW", variable=rot, value=270).grid(row=0, column=2, padx=6, sticky="w")
    ttk.Radiobutton(frmOpt, text="0\u00b0  (normalise only \u2014 bake existing Acrobat rotation)", variable=rot, value=0).grid(row=0, column=3, padx=6, sticky="w")

    bake = tk.BooleanVar(value=True)
    ttk.Checkbutton(frmOpt, text="Bake existing /Rotate flags into the content (recommended)", variable=bake).grid(row=1, column=0, columnspan=4, padx=6, pady=(6, 0), sticky="w")

    frmOut = ttk.LabelFrame(root, text="Output", padding=8)
    frmOut.pack(fill="x", padx=8, pady=(0, 4))

    same_folder = tk.BooleanVar(value=True)
    out_dir = tk.StringVar(value="")
    suffix = tk.StringVar(value=SUFFIX_DEFAULT)
    overwrite = tk.BooleanVar(value=False)

    ttk.Checkbutton(frmOut, text="Save next to source", variable=same_folder,
                    command=lambda: entDir.configure(state=("disabled" if same_folder.get() else "normal"))).grid(row=0, column=0, sticky="w", padx=4)
    entDir = ttk.Entry(frmOut, textvariable=out_dir, width=52, state="disabled")
    entDir.grid(row=0, column=1, sticky="we", padx=4)

    def browse_dir():
        d = filedialog.askdirectory(title="Output folder")
        if d:
            out_dir.set(d)
            same_folder.set(False)
            entDir.configure(state="normal")
    ttk.Button(frmOut, text="Browse\u2026", command=browse_dir).grid(row=0, column=2, padx=4)

    ttk.Label(frmOut, text="Filename suffix:").grid(row=1, column=0, sticky="e", padx=4, pady=(6, 0))
    ttk.Entry(frmOut, textvariable=suffix, width=18).grid(row=1, column=1, sticky="w", padx=4, pady=(6, 0))
    ttk.Checkbutton(frmOut, text="Overwrite existing output files", variable=overwrite).grid(row=1, column=2, sticky="w", padx=4, pady=(6, 0))
    frmOut.columnconfigure(1, weight=1)

    # ---------- bottom: run + log ----------
    frmRun = ttk.Frame(root, padding=8)
    frmRun.pack(fill="both", expand=False)

    prog = ttk.Progressbar(frmRun, mode="determinate")
    prog.pack(fill="x")
    btnGo = ttk.Button(frmRun, text="Rotate PDFs")
    btnGo.pack(pady=6)

    log = tk.Text(root, height=8, state="disabled", wrap="none")
    log.pack(fill="both", expand=False, padx=8, pady=(0, 8))

    def log_line(s):
        log.configure(state="normal")
        log.insert("end", s + "\n")
        log.see("end")
        log.configure(state="disabled")

    def out_path_for(src):
        base = os.path.splitext(os.path.basename(src))[0]
        folder = os.path.dirname(src) if same_folder.get() else (out_dir.get() or os.path.dirname(src))
        cand = os.path.join(folder, base + suffix.get() + ".pdf")
        if os.path.exists(cand) and not overwrite.get():
            n = 2
            while True:
                alt = os.path.join(folder, "%s%s (%d).pdf" % (base, suffix.get(), n))
                if not os.path.exists(alt):
                    return alt
                n += 1
        return cand

    def worker(job_files, cw, bake_flags):
        done = 0
        for src in job_files:
            try:
                dst = out_path_for(src)
                info = rotate_pdf_file(src, dst, user_cw=cw, bake_existing=bake_flags)
                warn = ""
                if info["annot_pages"]:
                    warn = "  \u26a0 %d page(s) carry annotations (not rotated \u2014 check)" % info["annot_pages"]
                msg_q.put(("status", src, "OK \u2192 " + os.path.basename(dst)))
                msg_q.put(("log", None, "OK   %s  \u2192  %s   (%d/%d pages rotated, flags in source: %s)%s" %
                           (os.path.basename(src), os.path.basename(dst),
                            info["rotated_pages"], info["pages"],
                            ",".join(str(x) for x in info["flags_found"]) or "none", warn)))
            except pikepdf.PasswordError:
                msg_q.put(("status", src, "SKIPPED \u2014 encrypted"))
                msg_q.put(("log", None, "SKIP %s  \u2014 password protected" % os.path.basename(src)))
            except Exception as ex:
                msg_q.put(("status", src, "FAILED"))
                msg_q.put(("log", None, "FAIL %s  \u2014 %s" % (os.path.basename(src), ex)))
                msg_q.put(("log", None, traceback.format_exc().strip().splitlines()[-1]))
            done += 1
            msg_q.put(("progress", None, done))
        msg_q.put(("done", None, None))

    def pump():
        try:
            while True:
                kind, key, val = msg_q.get_nowait()
                if kind == "status":
                    if tree.exists(key):
                        tree.set(key, "status", val)
                elif kind == "log":
                    log_line(val)
                elif kind == "progress":
                    prog["value"] = val
                elif kind == "done":
                    working["flag"] = False
                    btnGo.configure(state="normal", text="Rotate PDFs")
                    log_line("Done.")
        except queue.Empty:
            pass
        root.after(80, pump)

    def go():
        if working["flag"]:
            return
        if not files:
            messagebox.showinfo(APP_TITLE, "Add at least one PDF first.")
            return
        if rot.get() == 0 and not bake.get():
            messagebox.showinfo(APP_TITLE, "0\u00b0 with flag-baking off would do nothing.\nTick 'Bake existing /Rotate flags' or pick a rotation.")
            return
        if not same_folder.get() and out_dir.get() and not os.path.isdir(out_dir.get()):
            messagebox.showerror(APP_TITLE, "Output folder does not exist:\n" + out_dir.get())
            return
        working["flag"] = True
        btnGo.configure(state="disabled", text="Working\u2026")
        prog.configure(maximum=len(files), value=0)
        for iid in tree.get_children():
            tree.set(iid, "status", "queued")
        threading.Thread(target=worker, args=(list(files), rot.get(), bake.get()), daemon=True).start()

    btnGo.configure(command=go)
    pump()
    root.mainloop()


if __name__ == "__main__":
    main()
