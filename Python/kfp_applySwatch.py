#!/usr/bin/env python3
"""
kfp_apply.py — apply Acrobat Preflight .kfp colour fixups to a PDF,
without Acrobat.

Implements the two fixup features found in colour-mapping profiles:

  RemapProcessToSpot   Find DeviceCMYK fills/strokes (k / K operators)
                       whose values match a target build within the
                       profile's tolerance, and replace them with a
                       Separation (spot) colourspace at 100% tint.

  MapSpotColors        Find existing Separation colourspaces by name,
                       rename them to the target spot, and rewrite the
                       alternate-space tint transform to the profile's
                       CMYK values.

Content streams of pages AND Form XObjects are processed; text and
vector both use the same fill/stroke operators, matching the profile's
VectorAndText scope. Images, shadings and DeviceN spaces are left
untouched (reported instead).

Usage
-----
  python kfp_apply.py profile.kfp input.pdf output.pdf   # apply
  python kfp_apply.py profile.kfp input.pdf --dry-run    # report only
  python kfp_apply.py                                    # GUI mode

Requires: pikepdf  (pip install pikepdf), PySide6 for GUI mode only.
"""

import argparse
import re
import sys
from collections import Counter
from pathlib import Path

import pikepdf
from pikepdf import Array, Dictionary, Name, Operator

# ------------------------------------------------------------------ #
#  .kfp parsing (same schema as kfp_tool / kfp_editor)               #
# ------------------------------------------------------------------ #

FCFG_RE = re.compile(
    r"<fcfg>\s*<ffeat>(?P<feat>[^<]*)</ffeat>\s*<fparams>(?P<params>.*?)"
    r"</fparams>\s*</fcfg>", re.S)
FPARAM_RE = re.compile(r"<fparam>(.*?)</fparam>", re.S)


def _unesc(s):
    return (s.replace("&lt;", "<").replace("&gt;", ">")
             .replace("&apos;", "'").replace("&quot;", '"')
             .replace("&amp;", "&"))


class KfpProfile:
    """Colour rules extracted from a .kfp export."""

    def __init__(self, path):
        text = Path(path).read_text(encoding="utf-8")
        # (c, m, y, k, tol) all in 0-100, spot name
        self.remaps = []              # [((c,m,y,k), tol, spot)]
        self.spot_renames = {}        # source -> (target, (c,m,y,k))
        for m in FCFG_RE.finditer(text):
            p = [_unesc(x) for x in FPARAM_RE.findall(m.group("params"))]
            if m.group("feat") == "RemapProcessToSpot" and len(p) >= 7:
                cmyk = tuple(float(v) for v in p[1:5])
                self.remaps.append((cmyk, float(p[5]), p[6]))
            elif m.group("feat") == "MapSpotColors" and len(p) >= 8:
                cmyk = tuple(float(v) for v in p[4:8])
                self.spot_renames[p[0]] = (p[2], cmyk)
        if not self.remaps and not self.spot_renames:
            raise ValueError("No colour-mapping fixups found in profile")

    def match_remap(self, cmyk_0_1):
        """cmyk as 0..1 floats -> (spot, target_cmyk_0_100) or None."""
        vals = [v * 100 for v in cmyk_0_1]
        for target, tol, spot in self.remaps:
            if all(abs(a - b) <= tol for a, b in zip(vals, target)):
                return spot, target
        return None


# ------------------------------------------------------------------ #
#  PDF surgery                                                       #
# ------------------------------------------------------------------ #

OP_FILL_CMYK = Operator("k")
OP_STROKE_CMYK = Operator("K")
OP_CS = Operator("cs")
OP_CS_STROKE = Operator("CS")
OP_SCN = Operator("scn")
OP_SCN_STROKE = Operator("SCN")


def _tint_function(pdf, cmyk_0_100):
    return pdf.make_indirect(Dictionary(
        FunctionType=2, Domain=[0, 1], N=1,
        C0=[0, 0, 0, 0],
        C1=[round(v / 100, 6) for v in cmyk_0_100]))


def _separation(pdf, spot_name, cmyk_0_100):
    return pdf.make_indirect(Array([
        Name("/Separation"), Name("/" + spot_name),
        Name("/DeviceCMYK"), _tint_function(pdf, cmyk_0_100)]))


class Applier:
    def __init__(self, pdf, profile):
        self.pdf = pdf
        self.profile = profile
        self.report = Counter()
        self.notes = []
        self._sep_cache = {}     # spot name -> indirect Separation array
        self._visited = set()    # objgen of processed XObject streams

    # ---- shared Separation objects -------------------------------- #
    def _sep_for(self, spot, cmyk):
        if spot not in self._sep_cache:
            self._sep_cache[spot] = _separation(self.pdf, spot, cmyk)
        return self._sep_cache[spot]

    # ---- fixup 2: MapSpotColors ----------------------------------- #
    def rename_spots(self):
        renames = self.profile.spot_renames
        if not renames:
            return
        for obj in self.pdf.objects:
            if not isinstance(obj, Array) or len(obj) < 4:
                continue
            if obj[0] != Name("/Separation"):
                if obj[0] == Name("/DeviceN"):
                    names = [str(n)[1:] for n in obj[1]
                             if isinstance(n, Name)]
                    hit = [n for n in names if n in renames]
                    if hit:
                        self.notes.append(
                            f"DeviceN space contains {hit} — left "
                            f"untouched (rename Separations only)")
                continue
            src = str(obj[1])[1:]
            if src not in renames:
                continue
            target, cmyk = renames[src]
            obj[1] = Name("/" + target)
            obj[2] = Name("/DeviceCMYK")
            obj[3] = _tint_function(self.pdf, cmyk)
            self.report[("rename", src, target, tuple(cmyk))] += 1

    # ---- fixup 1: RemapProcessToSpot ------------------------------ #
    def remap_pages(self):
        if not self.profile.remaps:
            return
        for page in self.pdf.pages:
            self._remap_container(page)

    def _remap_container(self, container):
        """container: Page or Form XObject with Resources + content."""
        res = container.get("/Resources", None)
        new_keys = {}
        instructions = []
        changed = False

        try:
            parsed = pikepdf.parse_content_stream(container)
        except pikepdf.PdfError as exc:
            self.notes.append(f"content stream skipped: {exc}")
            parsed = []

        for operands, op in parsed:
            hit = None
            if op in (OP_FILL_CMYK, OP_STROKE_CMYK) and len(operands) == 4:
                cmyk = [float(v) for v in operands]
                hit = self.profile.match_remap(cmyk)
            if hit:
                spot, target = hit
                key = new_keys.setdefault(
                    spot, f"KFPSpot{len(self._sep_cache)}"
                    if spot not in self._sep_cache
                    else self._existing_key(spot, new_keys))
                self._sep_for(spot, target)
                cs_op = OP_CS if op == OP_FILL_CMYK else OP_CS_STROKE
                sc_op = OP_SCN if op == OP_FILL_CMYK else OP_SCN_STROKE
                instructions.append(([Name("/" + key)], cs_op))
                instructions.append(([1], sc_op))
                kind = "fill" if op == OP_FILL_CMYK else "stroke"
                src = tuple(round(c * 100, 2) for c in cmyk)
                self.report[("remap", kind, src, spot, tuple(target))] += 1
                changed = True
            else:
                instructions.append((operands, op))

        if changed:
            self._install(container, res, new_keys, instructions)

        # recurse into Form XObjects
        if res is not None and "/XObject" in res:
            for xkey in list(res["/XObject"].keys()):
                xo = res["/XObject"][xkey]
                if xo.get("/Subtype") == Name("/Form"):
                    og = (xo.objgen if hasattr(xo, "objgen") else None)
                    if og and og in self._visited:
                        continue
                    if og:
                        self._visited.add(og)
                    self._remap_container(xo)

    def _existing_key(self, spot, new_keys):
        return new_keys.get(spot) or f"KFPSpot{len(self._sep_cache)}"

    def _install(self, container, res, new_keys, instructions):
        if res is None:
            res = container.Resources = self.pdf.make_indirect(Dictionary())
        if "/ColorSpace" not in res:
            res["/ColorSpace"] = Dictionary()
        for spot, key in new_keys.items():
            res["/ColorSpace"][Name("/" + key)] = self._sep_cache[spot]
        new_stream = pikepdf.unparse_content_stream(instructions)
        if isinstance(container, pikepdf.Page):
            container.Contents = self.pdf.make_stream(new_stream)
        else:                                   # Form XObject
            container.write(new_stream)

    # ---- run ------------------------------------------------------- #
    def run(self):
        self.rename_spots()
        self.remap_pages()
        return self.report, self.notes


def apply_profile(kfp_path, pdf_in, pdf_out=None):
    """Returns (report Counter, notes list). Writes pdf_out if given."""
    profile = KfpProfile(kfp_path)
    with pikepdf.open(pdf_in) as pdf:
        applier = Applier(pdf, profile)
        report, notes = applier.run()
        if pdf_out:
            pdf.save(pdf_out)
    return report, notes


def _rgb(cmyk_0_100):
    c, m, y, k = (min(max(float(v), 0), 100) / 100 for v in cmyk_0_100)
    return (round(255 * (1 - c) * (1 - k)),
            round(255 * (1 - m) * (1 - k)),
            round(255 * (1 - y) * (1 - k)))


def _cmyk_str(cmyk):
    return "/".join(str(v).rstrip("0").rstrip(".") if "." in str(v)
                    else str(v) for v in (round(float(x), 2) for x in cmyk))


def _describe(entry):
    """entry -> (text, src_cmyk|None, dst_cmyk). CMYK in 0-100."""
    if entry[0] == "remap":
        _, kind, src, spot, target = entry
        return (f"remap {kind}: {_cmyk_str(src)} -> {spot}", src, target)
    _, src_name, target_name, cmyk = entry
    return (f"spot renamed: {src_name} -> {target_name}", None, cmyk)


def _ansi_swatch(cmyk):
    r, g, b = _rgb(cmyk)
    return f"\x1b[48;2;{r};{g};{b}m  \x1b[0m"


def format_report(report, notes, applied, color=None):
    """Plain/ANSI text report. color: True/False/None (auto-detect)."""
    if color is None:
        color = sys.stdout.isatty()
    lines = []
    verb = "Applied" if applied else "Would apply (dry run)"
    if not report:
        lines.append("No matching colours found — nothing to change.")
    else:
        lines.append(f"{verb}:")
        for entry, n in sorted(report.items()):
            text, src, dst = _describe(entry)
            if color:
                sw = ((_ansi_swatch(src) + "\u2192" if src else "  \u2192")
                      + _ansi_swatch(dst))
                lines.append(f"  {n:4d} x {sw}  {text}")
            else:
                lines.append(f"  {n:4d} x {text}")
        lines.append(f"  total operations: {sum(report.values())}")
    for note in notes:
        lines.append(f"  note: {note}")
    return "\n".join(lines)


def format_report_html(report, notes, applied):
    """HTML report with inline colour chips, for the GUI."""
    def chip(cmyk):
        r, g, b = _rgb(cmyk)
        return (f'<span style="background:rgb({r},{g},{b});'
                'border:1px solid #888;">&nbsp;&nbsp;&nbsp;&nbsp;</span>')

    verb = "Applied" if applied else "Would apply (dry run)"
    if not report:
        body = ["No matching colours found — nothing to change."]
    else:
        body = [f"<b>{verb}:</b>"]
        for entry, n in sorted(report.items()):
            text, src, dst = _describe(entry)
            swatches = ((chip(src) + " \u2192 ") if src else "\u2192 ") \
                + chip(dst)
            body.append(f"&nbsp;&nbsp;{n} \u00d7 {swatches} "
                        f"&nbsp;{text}")
        body.append(f"&nbsp;&nbsp;total operations: "
                    f"{sum(report.values())}")
    for note in notes:
        body.append(f"&nbsp;&nbsp;note: {note}")
    return ('<div style="font-family:monospace; font-size:10pt; '
            'line-height:1.6;">' + "<br>".join(body) + "</div>")


# ------------------------------------------------------------------ #
#  CLI                                                               #
# ------------------------------------------------------------------ #

def cli(argv):
    ap = argparse.ArgumentParser(
        description="Apply Acrobat Preflight .kfp colour fixups to a PDF")
    ap.add_argument("kfp", help="Preflight profile (.kfp)")
    ap.add_argument("pdf_in", help="input PDF")
    ap.add_argument("pdf_out", nargs="?", help="output PDF")
    ap.add_argument("--dry-run", action="store_true",
                    help="report matches without writing")
    ap.add_argument("--color", choices=["auto", "always", "never"],
                    default="auto", help="colour swatches in the report")
    args = ap.parse_args(argv)
    if not args.dry_run and not args.pdf_out:
        ap.error("give an output path, or use --dry-run")
    out = None if args.dry_run else args.pdf_out
    report, notes = apply_profile(args.kfp, args.pdf_in, out)
    color = {"auto": None, "always": True, "never": False}[args.color]
    print(format_report(report, notes, applied=bool(out), color=color))
    if out:
        print(f"Wrote {out}")


# ------------------------------------------------------------------ #
#  GUI (optional, PySide6)                                           #
# ------------------------------------------------------------------ #

def gui():
    from PySide6.QtWidgets import (
        QApplication, QFileDialog, QGridLayout, QLabel, QLineEdit,
        QMainWindow, QMessageBox, QPushButton, QTextEdit, QWidget)

    class Win(QMainWindow):
        def __init__(self):
            super().__init__()
            self.setWindowTitle("Apply KFP colour fixups to PDF")
            self.resize(720, 480)
            w = QWidget(); grid = QGridLayout(w)
            self.kfp = QLineEdit(); self.pdf = QLineEdit()
            for row, (label, edit, filt) in enumerate((
                    ("Profile (.kfp):", self.kfp,
                     "Preflight profiles (*.kfp)"),
                    ("Input PDF:", self.pdf, "PDF files (*.pdf)"))):
                grid.addWidget(QLabel(label), row, 0)
                grid.addWidget(edit, row, 1)
                btn = QPushButton("Browse…")
                btn.clicked.connect(
                    lambda _=False, e=edit, f=filt: self._pick(e, f))
                grid.addWidget(btn, row, 2)
            self.dry = QPushButton("Dry run (report only)")
            self.go = QPushButton("Apply && Save As…")
            self.dry.clicked.connect(lambda: self._run(dry=True))
            self.go.clicked.connect(lambda: self._run(dry=False))
            grid.addWidget(self.dry, 2, 1)
            grid.addWidget(self.go, 2, 2)
            self.out = QTextEdit(readOnly=True)
            grid.addWidget(self.out, 3, 0, 1, 3)
            self.setCentralWidget(w)

        def _pick(self, edit, filt):
            path, _ = QFileDialog.getOpenFileName(self, "Choose", "", filt)
            if path:
                edit.setText(path)

        def _run(self, dry):
            kfp, pdf = self.kfp.text().strip(), self.pdf.text().strip()
            if not (Path(kfp).is_file() and Path(pdf).is_file()):
                QMessageBox.warning(self, "Missing file",
                                    "Choose a valid .kfp and PDF first.")
                return
            out = None
            if not dry:
                out, _ = QFileDialog.getSaveFileName(
                    self, "Save converted PDF",
                    str(Path(pdf).with_stem(Path(pdf).stem + "_spot")),
                    "PDF files (*.pdf)")
                if not out:
                    return
            try:
                report, notes = apply_profile(kfp, pdf, out)
            except Exception as exc:                   # noqa: BLE001
                QMessageBox.critical(self, "Failed", str(exc))
                return
            html = format_report_html(report, notes, applied=bool(out))
            if out:
                html += (f'<div style="font-family:monospace;">'
                         f'Wrote {out}</div>')
            self.out.setHtml(html)

    app = QApplication(sys.argv)
    win = Win(); win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    if len(sys.argv) > 1:
        cli(sys.argv[1:])
    else:
        gui()
