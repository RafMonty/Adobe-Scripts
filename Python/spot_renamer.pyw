#!/usr/bin/env python3
"""
Spot Plate Renamer — prepress tool
==================================
Scans one or more PDFs for named spot colours (Separation / DeviceN
colorants, e.g. "PANTONE 185 C"), lists each unique plate with its total
occurrence count and an approximate colour swatch (taken from the plate's
tint transform / alternate colour space), and lets you rename plates (or
bulk-convert names to UPPERCASE) before saving — either in place or into an
output folder, leaving the originals untouched.

Requires:  pip install pikepdf
Tested with pikepdf 8/9.x on Python 3.9+
"""

import math
import operator
import os
import sys
import threading
import queue
import traceback
from collections import Counter, defaultdict
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    HAVE_TK = True
except ImportError:
    HAVE_TK = False

try:
    import pikepdf
    from pikepdf import Name, Array, Dictionary
except ImportError:
    sys.exit("pikepdf is required:  pip install pikepdf")

IGNORE_NAMES = {"All", "None", "Cyan", "Magenta", "Yellow", "Black"}

APP_TITLE = "Spot Plate Renamer"


# ----------------------------------------------------------------------
# Core PDF logic (no UI dependencies — reusable from the command line)
# ----------------------------------------------------------------------

def _name_str(name_obj) -> str:
    """Decoded PDF name without the leading slash ('/PANTONE#20185#20C' -> 'PANTONE 185 C')."""
    return str(name_obj)[1:]


def _walk(obj, visit, seen):
    """Recursively visit every Dictionary/Array reachable from obj.

    Indirect objects are tracked by (objnum, gen) to avoid cycles;
    direct objects are plain trees so recursion is safe.
    """
    if isinstance(obj, (Dictionary, Array)):
        if obj.is_indirect:
            key = obj.objgen
            if key in seen:
                return
            seen.add(key)
        visit(obj)
        if isinstance(obj, Dictionary):
            for v in obj.values():
                _walk(v, visit, seen)
        else:
            for v in obj:
                _walk(v, visit, seen)


def _for_each_spot(pdf, on_sep, on_devn, on_colorants):
    """Fire callbacks for every Separation array, DeviceN colorant list,
    and /Colorants dictionary in the document."""

    def visit(obj):
        if isinstance(obj, Array) and len(obj) >= 2:
            try:
                head = obj[0]
            except Exception:
                return
            if head == Name.Separation and isinstance(obj[1], Name):
                on_sep(obj)
            elif head == Name.DeviceN and isinstance(obj[1], Array):
                on_devn(obj)
        elif isinstance(obj, Dictionary) and Name.Colorants in obj:
            col = obj[Name.Colorants]
            if isinstance(col, Dictionary):
                on_colorants(col)

    seen = set()
    for indirect in pdf.objects:
        _walk(indirect, visit, seen)
    _walk(pdf.trailer, visit, seen)


def scan_pdf(path, colours: dict = None) -> Counter:
    """Return Counter{spot_name: occurrences} for one PDF.

    If `colours` (a dict) is given it is filled with {spot_name: (r, g, b)}
    approximate sRGB previews (0..1 floats) for every plate whose tint
    transform could be evaluated. Plates that fail are simply left out.
    """
    counts = Counter()

    def remember_colour(n, getter):
        if colours is None or n in colours:
            return
        try:
            colours[n] = getter()
        except Exception:
            pass

    def on_sep(arr):
        n = _name_str(arr[1])
        if n not in IGNORE_NAMES:
            counts[n] += 1
            remember_colour(n, lambda: _spot_rgb_from_sep(arr))

    def on_devn(arr):
        for i, item in enumerate(arr[1]):
            if isinstance(item, Name):
                n = _name_str(item)
                if n not in IGNORE_NAMES:
                    counts[n] += 1
                    remember_colour(n, lambda i=i: _spot_rgb_from_devn(arr, i))

    def on_colorants(d):
        # Keys of /Colorants duplicate the colorant names; counted via the
        # Separation values they point at, so don't double-count here.
        pass

    with pikepdf.open(path) as pdf:
        _for_each_spot(pdf, on_sep, on_devn, on_colorants)
    return counts


def rename_in_pdf(in_path, out_path, mapping: dict) -> int:
    """Apply {old_name: new_name} across a PDF. Returns number of edits.

    in_path may equal out_path (safe in-place save via a temp file).
    """
    edits = 0

    def maybe_new(name_obj):
        old = _name_str(name_obj)
        new = mapping.get(old)
        if new and new != old:
            return Name("/" + new)
        return None

    def on_sep(arr):
        nonlocal edits
        repl = maybe_new(arr[1])
        if repl is not None:
            arr[1] = repl
            edits += 1

    def on_devn(arr):
        nonlocal edits
        names = arr[1]
        for i, item in enumerate(names):
            if isinstance(item, Name):
                repl = maybe_new(item)
                if repl is not None:
                    names[i] = repl
                    edits += 1

    def on_colorants(d):
        nonlocal edits
        for key in list(d.keys()):
            old = str(key).lstrip("/")
            new = mapping.get(old)
            if new and new != old:
                d[Name("/" + new)] = d[key]
                del d[key]
                edits += 1

    with pikepdf.open(in_path, allow_overwriting_input=True) as pdf:
        _for_each_spot(pdf, on_sep, on_devn, on_colorants)
        if edits:
            pdf.save(out_path)
        elif str(in_path) != str(out_path):
            pdf.save(out_path)  # copy through unchanged so the out folder is complete
    return edits


# ----------------------------------------------------------------------
# Spot colour preview — evaluate a plate's tint transform at 100% and
# convert the alternate-space result to an approximate sRGB triple.
# Only used for on-screen swatches; the PDF itself is never modified by it.
# ----------------------------------------------------------------------

def _nums(obj):
    return [float(x) for x in obj]


def _clamp01(v):
    return max(0.0, min(1.0, float(v)))


def _interp(x, x0, x1, y0, y1):
    if x1 == x0:
        return y0
    return y0 + (x - x0) * (y1 - y0) / (x1 - x0)


# --- PDF functions (ISO 32000 §7.10) ---

def _read_bits(data, bitpos, nbits):
    byte, off = divmod(bitpos, 8)
    nbytes = (off + nbits + 7) // 8
    chunk = int.from_bytes(data[byte:byte + nbytes].ljust(nbytes, b"\0"), "big")
    return (chunk >> (nbytes * 8 - off - nbits)) & ((1 << nbits) - 1)


def _eval_sampled(fn, xs, domain):
    """Type 0: nearest-sample lookup (no interpolation — fine for a swatch)."""
    data = fn.read_bytes()
    size = [int(s) for s in fn["/Size"]]
    bps = int(fn["/BitsPerSample"])
    rng = _nums(fn["/Range"])
    n_out = len(rng) // 2
    encode = _nums(fn.get("/Encode", [v for s in size for v in (0, s - 1)]))
    decode = _nums(fn.get("/Decode", rng))
    maxval = (1 << bps) - 1
    idx, stride = 0, 1
    for i, s in enumerate(size):
        e = _interp(xs[i], domain[2 * i], domain[2 * i + 1], encode[2 * i], encode[2 * i + 1])
        idx += int(round(max(0, min(s - 1, e)))) * stride
        stride *= s
    out = []
    for j in range(n_out):
        v = _read_bits(data, (idx * n_out + j) * bps, bps)
        out.append(_interp(v, 0, maxval, decode[2 * j], decode[2 * j + 1]))
    return out


def _ps_bool_or_int(f, a, b):
    if isinstance(a, bool) and isinstance(b, bool):
        return bool(f(a, b))
    return float(f(int(a), int(b)))


_PS_BINARY = {
    "add": operator.add, "sub": operator.sub, "mul": operator.mul,
    "div": lambda a, b: a / b if b else 0.0,
    "idiv": lambda a, b: float(int(a / b)) if b else 0.0,
    "mod": lambda a, b: float(math.fmod(int(a), int(b))) if int(b) else 0.0,
    "exp": lambda a, b: a ** b,
    "atan": lambda a, b: math.degrees(math.atan2(a, b)) % 360.0,
    "eq": operator.eq, "ne": operator.ne, "gt": operator.gt,
    "ge": operator.ge, "lt": operator.lt, "le": operator.le,
    "and": lambda a, b: _ps_bool_or_int(operator.and_, a, b),
    "or": lambda a, b: _ps_bool_or_int(operator.or_, a, b),
    "xor": lambda a, b: _ps_bool_or_int(operator.xor, a, b),
    "bitshift": lambda a, b: float(int(a) << int(b)) if b >= 0 else float(int(a) >> int(-b)),
}
_PS_UNARY = {
    "neg": operator.neg, "abs": abs,
    "sqrt": lambda a: math.sqrt(max(a, 0.0)),
    "sin": lambda a: math.sin(math.radians(a)),
    "cos": lambda a: math.cos(math.radians(a)),
    "ln": lambda a: math.log(a) if a > 0 else 0.0,
    "log": lambda a: math.log10(a) if a > 0 else 0.0,
    "cvi": lambda a: float(int(a)), "cvr": float,
    "floor": lambda a: float(math.floor(a)),
    "ceiling": lambda a: float(math.ceil(a)),
    "round": lambda a: float(math.floor(a + 0.5)),
    "truncate": lambda a: float(int(a)),
    "not": lambda a: (not a) if isinstance(a, bool) else float(~int(a)),
}


def _ps_exec(prog, st, depth=0):
    if depth > 64:
        raise ValueError("PostScript calculator: nesting too deep")
    i = 0
    while i < len(prog):
        t = prog[i]
        i += 1
        if isinstance(t, list):                       # procedure body
            if i < len(prog) and prog[i] == "if":
                i += 1
                if st.pop():
                    _ps_exec(t, st, depth + 1)
            elif (i + 1 < len(prog) and isinstance(prog[i], list)
                  and prog[i + 1] == "ifelse"):
                other = prog[i]
                i += 2
                _ps_exec(t if st.pop() else other, st, depth + 1)
            else:
                raise ValueError("PostScript calculator: stray procedure")
            continue
        if t == "true":
            st.append(True)
        elif t == "false":
            st.append(False)
        elif t in _PS_BINARY:
            b = st.pop()
            a = st.pop()
            st.append(_PS_BINARY[t](a, b))
        elif t in _PS_UNARY:
            st.append(_PS_UNARY[t](st.pop()))
        elif t == "dup":
            st.append(st[-1])
        elif t == "pop":
            st.pop()
        elif t == "exch":
            st[-1], st[-2] = st[-2], st[-1]
        elif t == "copy":
            n = int(st.pop())
            if n > 0:
                st.extend(st[-n:])
        elif t == "index":
            n = int(st.pop())
            st.append(st[-1 - n])
        elif t == "roll":
            j = int(st.pop())
            n = int(st.pop())
            if n > 0 and j % n:
                j %= n
                seg = st[-n:]
                del st[-n:]
                st.extend(seg[-j:] + seg[:-j])
        else:
            st.append(float(t))                       # ValueError on junk


def _eval_postscript(fn, xs):
    """Type 4 (PostScript calculator) — small interpreter, enough for tint transforms."""
    tokens = fn.read_bytes().decode("latin-1").replace("{", " { ").replace("}", " } ").split()

    def parse(pos):
        prog = []
        while pos < len(tokens):
            tok = tokens[pos]
            pos += 1
            if tok == "{":
                sub, pos = parse(pos)
                prog.append(sub)
            elif tok == "}":
                return prog, pos
            else:
                prog.append(tok)
        return prog, pos

    prog, _ = parse(0)
    while len(prog) == 1 and isinstance(prog[0], list):   # strip outer { }
        prog = prog[0]
    st = list(xs)
    _ps_exec(prog, st)
    return [float(v) for v in st]


def _eval_function(fn, inputs):
    """Evaluate a PDF function (Type 0/2/3/4, or an array of them) -> list of floats."""
    if isinstance(fn, Array):
        out = []
        for sub in fn:
            out.extend(_eval_function(sub, inputs))
        return out
    ftype = int(fn.get("/FunctionType", -1))
    domain = _nums(fn.get("/Domain", [0, 1]))
    xs = [max(domain[2 * i], min(domain[2 * i + 1], x)) for i, x in enumerate(inputs)]
    if ftype == 2:
        c0 = _nums(fn.get("/C0", [0.0]))
        c1 = _nums(fn.get("/C1", [1.0]))
        n = float(fn.get("/N", 1))
        t = xs[0] ** n if xs[0] > 0 else 0.0
        out = [a + t * (b - a) for a, b in zip(c0, c1)]
    elif ftype == 3:
        funcs = fn["/Functions"]
        bounds = _nums(fn.get("/Bounds", []))
        encode = _nums(fn.get("/Encode", []))
        x = xs[0]
        k = 0
        while k < len(bounds) and x >= bounds[k]:
            k += 1
        k = min(k, len(funcs) - 1)
        lo = domain[0] if k == 0 else bounds[k - 1]
        hi = domain[1] if k >= len(bounds) else bounds[k]
        e0, e1 = (encode[2 * k], encode[2 * k + 1]) if len(encode) >= 2 * k + 2 else (0.0, 1.0)
        out = _eval_function(funcs[k], [_interp(x, lo, hi, e0, e1)])
    elif ftype == 0:
        out = _eval_sampled(fn, xs, domain)
    elif ftype == 4:
        out = _eval_postscript(fn, xs)
    else:
        raise ValueError("unsupported function type %r" % ftype)
    rng = fn.get("/Range")
    if rng is not None:
        rng = _nums(rng)
        out = [max(rng[2 * i], min(rng[2 * i + 1], v))
               for i, v in enumerate(out) if 2 * i + 1 < len(rng)]
    return out


# --- alternate colour space -> sRGB ---

# Approximate sRGB appearance of 100% process inks (roughly SWOP/ISO Coated).
# Mixed multiplicatively — closer to what Acrobat shows than naive 1-min(1,c+k).
_INK_RGB = {
    "c": (0.00, 0.68, 0.94),
    "m": (0.93, 0.00, 0.55),
    "y": (1.00, 0.95, 0.00),
    "k": (0.14, 0.12, 0.13),
}


def _cmyk_to_rgb(c, m, y, k):
    rgb = [1.0, 1.0, 1.0]
    for amt, ink in zip((c, m, y, k), ("c", "m", "y", "k")):
        amt = _clamp01(amt)
        for i in range(3):
            rgb[i] *= 1.0 - amt * (1.0 - _INK_RGB[ink][i])
    return tuple(rgb)


def _lab_to_rgb(L, a, b, wp):
    fy = (L + 16.0) / 116.0
    fx = fy + a / 500.0
    fz = fy - b / 200.0

    def finv(t):
        return t ** 3 if t > 6.0 / 29 else 3 * (6.0 / 29) ** 2 * (t - 4.0 / 29)

    X, Y, Z = wp[0] * finv(fx), wp[1] * finv(fy), wp[2] * finv(fz)
    # XYZ (D50) -> linear sRGB (Bradford-adapted matrix)
    r = 3.1339 * X - 1.6170 * Y - 0.4906 * Z
    g = -0.9785 * X + 1.9160 * Y + 0.0333 * Z
    bb = 0.0720 * X - 0.2290 * Y + 1.4057 * Z

    def gam(v):
        v = _clamp01(v)
        return 1.055 * v ** (1 / 2.4) - 0.055 if v > 0.0031308 else 12.92 * v

    return (gam(r), gam(g), gam(bb))


def _cs_to_rgb(cs, comps):
    """Convert components in colour space `cs` to an (r, g, b) tuple of 0..1 floats."""
    if isinstance(cs, Name):
        fam, params = str(cs), None
    elif isinstance(cs, Array) and len(cs) and isinstance(cs[0], Name):
        fam, params = str(cs[0]), cs
    else:
        raise ValueError("unrecognised colour space")

    if fam == "/ICCBased":
        n = int(params[1].get("/N", len(comps)))
        fam = {1: "/DeviceGray", 3: "/DeviceRGB", 4: "/DeviceCMYK"}.get(n, fam)
    if fam == "/DeviceCMYK":
        return _cmyk_to_rgb(*comps[:4])
    if fam in ("/DeviceRGB", "/CalRGB"):
        return tuple(_clamp01(v) for v in comps[:3])
    if fam in ("/DeviceGray", "/CalGray"):
        g = _clamp01(comps[0])
        return (g, g, g)
    if fam == "/Lab":
        wp = _nums(params[1].get("/WhitePoint", [0.9643, 1.0, 0.8251]))
        return _lab_to_rgb(comps[0], comps[1], comps[2], wp)
    if fam in ("/Separation", "/DeviceN"):
        # nested spot as alternate (rare): evaluate it at the given tint(s)
        return _cs_to_rgb(params[2], _eval_function(params[3], list(comps)))
    raise ValueError("unsupported alternate colour space %s" % fam)


def _spot_rgb_from_sep(arr):
    """[/Separation /Name alt tint] -> (r, g, b) at 100% tint."""
    return _cs_to_rgb(arr[2], _eval_function(arr[3], [1.0]))


def _spot_rgb_from_devn(arr, index):
    """[/DeviceN [names] alt tint attrs] -> (r, g, b) for colorant `index` at 100%.

    Prefers the per-colorant Separation in the /Colorants attribute (exact);
    otherwise solos that input of the shared tint transform.
    """
    names = arr[1]
    item = names[index]
    if len(arr) > 4 and isinstance(arr[4], Dictionary):
        colorants = arr[4].get("/Colorants")
        if isinstance(colorants, Dictionary):
            sep = colorants.get(str(item))
            if isinstance(sep, Array) and len(sep) >= 4:
                try:
                    return _spot_rgb_from_sep(sep)
                except Exception:
                    pass
    inputs = [1.0 if j == index else 0.0 for j in range(len(names))]
    return _cs_to_rgb(arr[2], _eval_function(arr[3], inputs))


def rgb_to_hex(rgb) -> str:
    return "#%02x%02x%02x" % tuple(int(round(_clamp01(v) * 255)) for v in rgb)


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------

if HAVE_TK:
  class SpotRenamerApp(tk.Tk):
      def __init__(self):
          super().__init__()
          self.title(APP_TITLE)
          self.geometry("880x640")
          self.minsize(760, 540)

          self.files: list[str] = []
          self.spot_counts: Counter = Counter()          # name -> total occurrences
          self.spot_files: defaultdict = defaultdict(set)  # name -> set(paths)
          self.spot_rgb: dict[str, tuple] = {}            # name -> (r, g, b) 0..1, if known
          self.renames: dict[str, str] = {}              # old -> new
          self._swatches: dict = {}                       # hex -> PhotoImage (kept alive)
          self._q = queue.Queue()
          self._busy = False

          self._build_ui()
          self.after(100, self._poll_queue)

      # ---------------- layout ----------------

      def _build_ui(self):
          pad = {"padx": 6, "pady": 4}

          # --- file section ---
          f_files = ttk.LabelFrame(self, text="PDF files")
          f_files.pack(fill="x", **pad)

          btns = ttk.Frame(f_files)
          btns.pack(side="left", fill="y", padx=4, pady=4)
          ttk.Button(btns, text="Add PDFs…", command=self.add_files).pack(fill="x", pady=2)
          ttk.Button(btns, text="Add Folder…", command=self.add_folder).pack(fill="x", pady=2)
          ttk.Button(btns, text="Remove Selected", command=self.remove_selected).pack(fill="x", pady=2)
          ttk.Button(btns, text="Clear List", command=self.clear_files).pack(fill="x", pady=2)

          self.lst = tk.Listbox(f_files, height=6, selectmode="extended")
          self.lst.pack(side="left", fill="both", expand=True, padx=4, pady=4)
          sb = ttk.Scrollbar(f_files, orient="vertical", command=self.lst.yview)
          sb.pack(side="left", fill="y")
          self.lst.configure(yscrollcommand=sb.set)

          # --- scan / plate table ---
          f_mid = ttk.LabelFrame(self, text="Spot plates  (double-click 'New Name' to edit)")
          f_mid.pack(fill="both", expand=True, **pad)

          bar = ttk.Frame(f_mid)
          bar.pack(fill="x", padx=4, pady=(4, 0))
          self.btn_scan = ttk.Button(bar, text="Scan PDFs", command=self.scan)
          self.btn_scan.pack(side="left", padx=2)
          ttk.Button(bar, text="UPPERCASE All", command=self.uppercase_all).pack(side="left", padx=2)
          ttk.Button(bar, text="Reset Renames", command=self.reset_renames).pack(side="left", padx=2)

          cols = ("plate", "occ", "nfiles", "new")
          self.tree = ttk.Treeview(f_mid, columns=cols, show="tree headings", height=10)
          # Column #0 (the "tree" column) carries the colour swatch + hex value
          self.tree.heading("#0", text="Swatch")
          self.tree.column("#0", width=120, minwidth=100, stretch=False)
          self.tree.heading("plate", text="Spot Colour")
          self.tree.heading("occ", text="Occurrences")
          self.tree.heading("nfiles", text="Files")
          self.tree.heading("new", text="New Name")
          self.tree.column("plate", width=280)
          self.tree.column("occ", width=90, anchor="center")
          self.tree.column("nfiles", width=60, anchor="center")
          self.tree.column("new", width=280)
          self.tree.pack(fill="both", expand=True, padx=4, pady=4)
          self.tree.bind("<Double-1>", self._edit_cell)

          # --- output options ---
          f_out = ttk.LabelFrame(self, text="Output")
          f_out.pack(fill="x", **pad)

          self.out_mode = tk.StringVar(value="folder")
          ttk.Radiobutton(f_out, text="Save into output folder (originals retained)",
                          variable=self.out_mode, value="folder").grid(row=0, column=0, sticky="w", padx=6)
          ttk.Radiobutton(f_out, text="Overwrite originals in place",
                          variable=self.out_mode, value="inplace").grid(row=1, column=0, sticky="w", padx=6)

          self.out_dir = tk.StringVar(value="")
          ttk.Entry(f_out, textvariable=self.out_dir, width=52).grid(row=0, column=1, sticky="ew", padx=6)
          ttk.Button(f_out, text="Browse…", command=self.pick_out_dir).grid(row=0, column=2, padx=6)
          f_out.columnconfigure(1, weight=1)

          self.btn_go = ttk.Button(f_out, text="Process PDFs", command=self.process)
          self.btn_go.grid(row=1, column=2, padx=6, pady=4, sticky="e")

          # --- log ---
          f_log = ttk.LabelFrame(self, text="Log")
          f_log.pack(fill="both", **pad)
          self.log = tk.Text(f_log, height=7, state="disabled", wrap="word")
          self.log.pack(fill="both", expand=True, padx=4, pady=4)

      # ---------------- file list ----------------

      def add_files(self):
          paths = filedialog.askopenfilenames(title="Select PDFs",
                                              filetypes=[("PDF files", "*.pdf")])
          self._add_paths(paths)

      def add_folder(self):
          d = filedialog.askdirectory(title="Select folder of PDFs")
          if d:
              self._add_paths(sorted(str(p) for p in Path(d).glob("*.pdf")))

      def _add_paths(self, paths):
          for p in paths:
              if p and p not in self.files:
                  self.files.append(p)
                  self.lst.insert("end", p)

      def remove_selected(self):
          for i in reversed(self.lst.curselection()):
              self.files.pop(i)
              self.lst.delete(i)

      def clear_files(self):
          self.files.clear()
          self.lst.delete(0, "end")

      def pick_out_dir(self):
          d = filedialog.askdirectory(title="Output folder")
          if d:
              self.out_dir.set(d)
              self.out_mode.set("folder")

      # ---------------- scanning ----------------

      def scan(self):
          if self._busy:
              return
          if not self.files:
              messagebox.showinfo(APP_TITLE, "Add some PDFs first.")
              return
          self._set_busy(True)
          self._log("Scanning %d file(s)…" % len(self.files))
          threading.Thread(target=self._scan_worker, args=(list(self.files),), daemon=True).start()

      def _scan_worker(self, files):
          counts, per_file, colours = Counter(), defaultdict(set), {}
          for f in files:
              try:
                  c = scan_pdf(f, colours)
                  counts.update(c)
                  for name in c:
                      per_file[name].add(f)
                  self._q.put(("log", f"  {os.path.basename(f)}: "
                                      f"{len(c)} spot plate(s), {sum(c.values())} occurrence(s)"))
              except Exception as e:
                  self._q.put(("log", f"  ERROR scanning {os.path.basename(f)}: {e}"))
          self._q.put(("scan_done", (counts, per_file, colours)))

      def _refresh_tree(self):
          self.tree.delete(*self.tree.get_children())
          for name in sorted(self.spot_counts):
              rgb = self.spot_rgb.get(name)
              self.tree.insert("", "end", iid=name,
                               image=self._swatch(rgb),
                               text=(" " + rgb_to_hex(rgb).upper()) if rgb else "  n/a",
                               values=(
                                   name,
                                   self.spot_counts[name],
                                   len(self.spot_files[name]),
                                   self.renames.get(name, ""),
                               ))

      def _swatch(self, rgb, w=30, h=16):
          """Solid-colour PhotoImage for the swatch column (cached per hex).
          `rgb` None -> white box with a diagonal grey slash (colour unknown)."""
          key = rgb_to_hex(rgb) if rgb else None
          img = self._swatches.get(key)
          if img is None:
              img = tk.PhotoImage(width=w, height=h)
              img.put("#7f7f7f", to=(0, 0, w, h))              # 1px border
              img.put(key or "#ffffff", to=(1, 1, w - 1, h - 1))
              if key is None:
                  for i in range(1, h - 1):                    # diagonal slash
                      x = 1 + int((i - 1) * (w - 3) / (h - 3))
                      img.put("#b0b0b0", to=(x, i, x + 2, i + 1))
              self._swatches[key] = img
          return img

      # ---------------- rename editing ----------------

      def _edit_cell(self, event):
          row = self.tree.identify_row(event.y)
          if not row or self.tree.identify_column(event.x) != "#4":
              return
          x, y, w, h = self.tree.bbox(row, "#4")
          entry = ttk.Entry(self.tree)
          entry.place(x=x, y=y, width=w, height=h)
          entry.insert(0, self.renames.get(row, row))
          entry.select_range(0, "end")
          entry.focus_set()

          def commit(_=None):
              val = entry.get().strip()
              if val and val != row:
                  self.renames[row] = val
              else:
                  self.renames.pop(row, None)
              entry.destroy()
              self._refresh_tree()

          entry.bind("<Return>", commit)
          entry.bind("<FocusOut>", commit)
          entry.bind("<Escape>", lambda e: entry.destroy())

      def uppercase_all(self):
          for name in self.spot_counts:
              up = name.upper()
              if up != name:
                  self.renames[name] = up
          self._refresh_tree()

      def reset_renames(self):
          self.renames.clear()
          self._refresh_tree()

      # ---------------- processing ----------------

      def process(self):
          if self._busy:
              return
          if not self.spot_counts:
              messagebox.showinfo(APP_TITLE, "Scan the PDFs first.")
              return
          mapping = dict(self.renames)
          if not mapping:
              messagebox.showinfo(APP_TITLE, "No renames set — nothing to do.")
              return

          # Warn if renames merge two plates into one
          targets = Counter(mapping.get(n, n) for n in self.spot_counts)
          merged = [t for t, k in targets.items() if k > 1]
          if merged and not messagebox.askyesno(
                  APP_TITLE,
                  "These renames will MERGE plates:\n  " + "\n  ".join(merged) +
                  "\n\nContinue?"):
              return

          if self.out_mode.get() == "folder":
              out_dir = self.out_dir.get().strip()
              if not out_dir:
                  out_dir = os.path.join(os.path.dirname(self.files[0]), "out")
                  self.out_dir.set(out_dir)
              os.makedirs(out_dir, exist_ok=True)
          else:
              out_dir = None
              if not messagebox.askyesno(APP_TITLE,
                                         "Overwrite the original PDFs in place?"):
                  return

          self._set_busy(True)
          self._log("Processing…")
          threading.Thread(target=self._process_worker,
                           args=(list(self.files), mapping, out_dir), daemon=True).start()

      def _process_worker(self, files, mapping, out_dir):
          total = 0
          for f in files:
              try:
                  dest = os.path.join(out_dir, os.path.basename(f)) if out_dir else f
                  n = rename_in_pdf(f, dest, mapping)
                  total += n
                  self._q.put(("log", f"  {os.path.basename(f)}: {n} rename(s) -> {dest}"))
              except Exception as e:
                  self._q.put(("log", f"  ERROR processing {os.path.basename(f)}: {e}"))
                  self._q.put(("log", traceback.format_exc(limit=2)))
          self._q.put(("log", f"Done — {total} rename(s) applied across {len(files)} file(s)."))
          self._q.put(("process_done", None))

      # ---------------- plumbing ----------------

      def _poll_queue(self):
          try:
              while True:
                  kind, payload = self._q.get_nowait()
                  if kind == "log":
                      self._log(payload)
                  elif kind == "scan_done":
                      self.spot_counts, self.spot_files, self.spot_rgb = payload
                      self.renames = {k: v for k, v in self.renames.items()
                                      if k in self.spot_counts}
                      self._refresh_tree()
                      missing = [n for n in self.spot_counts if n not in self.spot_rgb]
                      self._log(f"Scan complete — {len(self.spot_counts)} unique plate(s), "
                                f"{len(self.spot_counts) - len(missing)} with colour preview.")
                      if missing:
                          self._log("  No preview (tint transform not understood): "
                                    + ", ".join(missing))
                      self._set_busy(False)
                  elif kind == "process_done":
                      self._set_busy(False)
          except queue.Empty:
              pass
          self.after(100, self._poll_queue)

      def _set_busy(self, busy):
          self._busy = busy
          state = "disabled" if busy else "normal"
          self.btn_scan.configure(state=state)
          self.btn_go.configure(state=state)

      def _log(self, msg):
          self.log.configure(state="normal")
          self.log.insert("end", msg + "\n")
          self.log.see("end")
          self.log.configure(state="disabled")


if __name__ == "__main__":
    if not HAVE_TK:
        sys.exit("tkinter is required for the UI (install python3-tk).")
    SpotRenamerApp().mainloop()
