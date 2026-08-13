#!/usr/bin/env python3
"""
kfp_editor.py — PySide6 editor for Acrobat Preflight .kfp profiles
containing colour-mapping fixups (RemapProcessToSpot / MapSpotColors).

Two tables:
  * Process → Spot   (match CMYK build ± tolerance, convert to spot)
  * Spot Renames     (rename spot, set CMYK alternate values)

Add / delete / edit rows, live CMYK swatches, then Save As. On save the
tool regenerates only the <fcfgs> payload of each colour fixup inside
the ORIGINAL file text, so all of Acrobat's ids, formatting, CRLF line
endings and untouched sections survive byte-for-byte. The fixup
<modification> timestamps are bumped so Acrobat treats re-import as an
update to the same profile.

    pip install PySide6
    python kfp_editor.py [file.kfp]
"""

import datetime
import re
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QBrush, QColor, QKeySequence
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QHBoxLayout, QHeaderView, QLabel,
    QMainWindow, QMessageBox, QPushButton, QSplitter, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
)

# ------------------------------------------------------------------ #
#  .kfp engine (text-surgery, format preserving)                     #
# ------------------------------------------------------------------ #

FIXUP_RE = re.compile(r"<fixup>.*?</fixup>", re.S)
FCFGS_RE = re.compile(r"(<fcfgs>)(.*?)(</fcfgs>)", re.S)
FCFG_RE = re.compile(
    r"<fcfg>\s*<ffeat>(?P<feat>[^<]*)</ffeat>\s*<fparams>(?P<params>.*?)"
    r"</fparams>\s*</fcfg>", re.S)
FPARAM_RE = re.compile(r"<fparam>(.*?)</fparam>", re.S)

IND = "\t" * 5          # indentation of <fcfg> within <fcfgs>
NL = "\r\n"

REMAP_TPL = (
    "<fcfg>{nl}"
    "{i}\t<ffeat>RemapProcessToSpot</ffeat>{nl}"
    "{i}\t<fparams>{nl}"
    "{p}CMYKPercent</fparam>{nl}{p}{c}</fparam>{nl}{p}{m}</fparam>{nl}"
    "{p}{y}</fparam>{nl}{p}{k}</fparam>{nl}{p}{tol}</fparam>{nl}"
    "{p}{spot}</fparam>{nl}{extra}"
    "{i}\t</fparams>{nl}"
    "{i}</fcfg>"
)

SPOTMAP_TPL = (
    "<fcfg>{nl}"
    "{i}\t<ffeat>MapSpotColors</ffeat>{nl}"
    "{i}\t<fparams>{nl}"
    "{p}{src}</fparam>{nl}{p}MapOrRename</fparam>{nl}{p}{tgt}</fparam>{nl}"
    "{p}CMYKPercent</fparam>{nl}{p}{c}</fparam>{nl}{p}{m}</fparam>{nl}"
    "{p}{y}</fparam>{nl}{p}{k}</fparam>{nl}{extra}"
    "{i}\t</fparams>{nl}"
    "{i}</fcfg>"
)


REMAP_EXTRA_DEFAULT = [
    "Automatic", "0", "0", "0", "0", "100", "PRCCUSTOMCHECK"]
SPOTMAP_EXTRA_DEFAULT = [
    "Unchanged", "VectorAndText", "1", "0", "3", "100", "0"]


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def unesc(s):
    return (s.replace("&lt;", "<").replace("&gt;", ">")
             .replace("&apos;", "'").replace("&quot;", '"')
             .replace("&amp;", "&"))


def fmt_num(v):
    """Render 75.0 as 75 but keep 67.5 as-is."""
    try:
        f = float(v)
        return str(int(f)) if f.is_integer() else str(f)
    except (TypeError, ValueError):
        return str(v)


class KfpDocument:
    """Loads a .kfp, exposes the two colour-fixup tables, writes back."""

    def __init__(self, path):
        self.path = Path(path)
        with open(self.path, encoding="utf-8", newline="") as fh:
            self.text = fh.read()          # newline='' keeps CRLF intact
        self.remaps = []      # [ [c, m, y, k, tol, spot], ... ]
        self.spotmaps = []    # [ [source, target, c, m, y, k], ... ]
        self._parse()

    def _parse(self):
        for m in FCFG_RE.finditer(self.text):
            p = [unesc(x) for x in FPARAM_RE.findall(m.group("params"))]
            if m.group("feat") == "RemapProcessToSpot" and len(p) >= 7:
                self.remaps.append(
                    ([p[1], p[2], p[3], p[4], p[5], p[6]], p[7:]))
            elif m.group("feat") == "MapSpotColors" and len(p) >= 8:
                self.spotmaps.append(
                    ([p[0], p[2], p[4], p[5], p[6], p[7]], p[8:]))
        if not self.remaps and not self.spotmaps:
            raise ValueError(
                "No RemapProcessToSpot / MapSpotColors fixups found — "
                "this .kfp doesn't look like a colour-mapping profile.")

    # ---- serialisation ------------------------------------------- #

    @staticmethod
    def _stamp():
        now = datetime.datetime.now().astimezone()
        off = now.strftime("%z")
        return now.strftime("D:%Y%m%d%H%M%S") + (
            f"{off[:3]}&apos;{off[3:]}&apos;" if off else "")

    def _render(self, feature, rows):
        p = IND + "\t\t<fparam>"
        kw = dict(nl=NL, i=IND, p=p)
        blocks = []
        for values, extras in rows:
            extra = "".join(f"{p}{esc(x)}</fparam>{NL}" for x in extras)
            if feature == "RemapProcessToSpot":
                c, m, y, k, tol, spot = values
                blocks.append(REMAP_TPL.format(
                    c=fmt_num(c), m=fmt_num(m), y=fmt_num(y), k=fmt_num(k),
                    tol=fmt_num(tol), spot=esc(spot), extra=extra, **kw))
            else:
                src, tgt, c, m, y, k = values
                blocks.append(SPOTMAP_TPL.format(
                    src=esc(src), tgt=esc(tgt), c=fmt_num(c), m=fmt_num(m),
                    y=fmt_num(y), k=fmt_num(k), extra=extra, **kw))
        return NL + IND + (NL + IND).join(blocks) + NL + "\t" * 4

    def save(self, out_path):
        text = self.text
        # Walk each <fixup>; regenerate its <fcfgs> if it holds our feature.
        out, pos = [], 0
        for fx in FIXUP_RE.finditer(text):
            seg = fx.group(0)
            feature = None
            if "<ffeat>RemapProcessToSpot</ffeat>" in seg:
                feature, rows = "RemapProcessToSpot", self.remaps
            elif "<ffeat>MapSpotColors</ffeat>" in seg:
                feature, rows = "MapSpotColors", self.spotmaps
            if feature:
                seg = FCFGS_RE.sub(
                    lambda m: m.group(1) + self._render(feature, rows)
                    + m.group(3), seg, count=1)
                seg = re.sub(r"<modification>.*?</modification>",
                             f"<modification>{self._stamp()}</modification>",
                             seg, count=1)
            out.append(text[pos:fx.start()])
            out.append(seg)
            pos = fx.end()
        out.append(text[pos:])
        with open(out_path, "w", encoding="utf-8", newline="") as fh:
            fh.write("".join(out))


# ------------------------------------------------------------------ #
#  UI helpers                                                        #
# ------------------------------------------------------------------ #

def cmyk_qcolor(c, m, y, k):
    try:
        c, m, y, k = (max(0.0, min(100.0, float(v))) / 100.0
                      for v in (c, m, y, k))
    except (TypeError, ValueError):
        return QColor(255, 255, 255)
    r = round(255 * (1 - c) * (1 - k))
    g = round(255 * (1 - m) * (1 - k))
    b = round(255 * (1 - y) * (1 - k))
    return QColor(r, g, b)


class MappingTable(QTableWidget):
    """Editable table with a leading colour-swatch column."""

    def __init__(self, headers, cmyk_cols, parent=None):
        super().__init__(0, len(headers) + 1, parent)
        self.cmyk_cols = cmyk_cols          # data-column indices of C M Y K
        self.setHorizontalHeaderLabels([""] + headers)
        self.verticalHeader().setVisible(False)
        self.setAlternatingRowColors(True)
        hdr = self.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.setColumnWidth(0, 34)
        self.itemChanged.connect(self._on_item_changed)
        self._loading = False

    # data <-> table ------------------------------------------------ #
    def load_rows(self, rows):
        self._loading = True
        self.setRowCount(0)
        for values, extras in rows:
            self._append(values, extras)
        self._loading = False

    def _append(self, values, extras):
        row = self.rowCount()
        self.insertRow(row)
        swatch = QTableWidgetItem()
        swatch.setFlags(Qt.ItemFlag.ItemIsEnabled)
        swatch.setData(Qt.ItemDataRole.UserRole, list(extras))
        self.setItem(row, 0, swatch)
        for col, v in enumerate(values, start=1):
            self.setItem(row, col, QTableWidgetItem(str(v)))
        self._update_swatch(row)

    def add_blank(self, defaults, extras):
        self._loading = True
        self._append(defaults, list(extras))
        self._loading = False
        self.scrollToBottom()

    def delete_selected(self):
        rows = sorted({i.row() for i in self.selectedIndexes()}, reverse=True)
        for r in rows:
            self.removeRow(r)
        return bool(rows)

    def rows(self):
        out = []
        for r in range(self.rowCount()):
            values = [
                (self.item(r, c + 1).text().strip()
                 if self.item(r, c + 1) else "")
                for c in range(self.columnCount() - 1)]
            extras = self.item(r, 0).data(Qt.ItemDataRole.UserRole) or []
            out.append((values, extras))
        return out

    # swatch -------------------------------------------------------- #
    def _on_item_changed(self, item):
        if not self._loading and item.column() != 0:
            self._update_swatch(item.row())

    def _update_swatch(self, row):
        vals = []
        for c in self.cmyk_cols:
            it = self.item(row, c + 1)
            vals.append(it.text() if it else "0")
        sw = self.item(row, 0)
        if sw:
            sw.setBackground(QBrush(cmyk_qcolor(*vals)))


class Editor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.doc = None
        self.setWindowTitle("KFP Colour Mapping Editor")
        self.resize(940, 640)
        self._build_ui()
        self._build_menu()

    # UI construction ------------------------------------------------ #
    def _build_ui(self):
        splitter = QSplitter(Qt.Orientation.Vertical)

        self.remap_table = MappingTable(
            ["C", "M", "Y", "K", "Tol %", "Spot name"],
            cmyk_cols=[0, 1, 2, 3])
        self.spot_table = MappingTable(
            ["Source spot", "Target spot", "C", "M", "Y", "K"],
            cmyk_cols=[2, 3, 4, 5])

        splitter.addWidget(self._panel(
            "Process → Spot  (RemapProcessToSpot)", self.remap_table,
            lambda: self.remap_table.add_blank(
                ["0", "0", "0", "0", "5", "New Spot"],
                REMAP_EXTRA_DEFAULT),
            self.remap_table.delete_selected))
        splitter.addWidget(self._panel(
            "Spot Renames  (MapSpotColors)", self.spot_table,
            lambda: self.spot_table.add_blank(
                ["Old Name", "New Name", "0", "0", "0", "0"],
                SPOTMAP_EXTRA_DEFAULT),
            self.spot_table.delete_selected))

        self.setCentralWidget(splitter)
        self.statusBar().showMessage("Open a .kfp file to begin")

    def _panel(self, title, table, on_add, on_del):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(6, 6, 6, 2)
        top = QHBoxLayout()
        lbl = QLabel(title)
        lbl.setStyleSheet("font-weight: 600;")
        top.addWidget(lbl)
        top.addStretch(1)
        add = QPushButton("Add")
        rem = QPushButton("Delete selected")
        add.clicked.connect(on_add)
        rem.clicked.connect(on_del)
        top.addWidget(add)
        top.addWidget(rem)
        lay.addLayout(top)
        lay.addWidget(table)
        return w

    def _build_menu(self):
        m = self.menuBar().addMenu("&File")
        for label, keys, slot in (
                ("&Open…", QKeySequence.StandardKey.Open, self.open_file),
                ("&Save", QKeySequence.StandardKey.Save, self.save),
                ("Save &As…", QKeySequence.StandardKey.SaveAs, self.save_as)):
            act = QAction(label, self)
            act.setShortcut(keys)
            act.triggered.connect(slot)
            m.addAction(act)

    # file ops -------------------------------------------------------- #
    def open_file(self, _=False, path=None):
        if path is None:
            path, _f = QFileDialog.getOpenFileName(
                self, "Open Preflight profile", "",
                "Preflight profiles (*.kfp);;All files (*)")
            if not path:
                return
        try:
            self.doc = KfpDocument(path)
        except Exception as exc:                       # noqa: BLE001
            QMessageBox.critical(self, "Cannot open", str(exc))
            return
        self.remap_table.load_rows(self.doc.remaps)
        self.spot_table.load_rows(self.doc.spotmaps)
        self.setWindowTitle(f"KFP Colour Mapping Editor — {Path(path).name}")
        self.statusBar().showMessage(
            f"{len(self.doc.remaps)} remaps, "
            f"{len(self.doc.spotmaps)} spot renames loaded")

    def _validate(self):
        problems = []
        for label, tbl, num_cols, name_cols in (
                ("Process → Spot", self.remap_table, (0, 1, 2, 3, 4), (5,)),
                ("Spot Renames", self.spot_table, (2, 3, 4, 5), (0, 1))):
            for i, (row, _extras) in enumerate(tbl.rows(), start=1):
                for c in num_cols:
                    try:
                        float(row[c])
                    except ValueError:
                        problems.append(
                            f"{label} row {i}: '{row[c]}' is not a number")
                for c in name_cols:
                    if not row[c]:
                        problems.append(f"{label} row {i}: empty name")
        return problems

    def _commit(self):
        self.doc.remaps = self.remap_table.rows()
        self.doc.spotmaps = self.spot_table.rows()

    def save(self):
        if not self.doc:
            return
        if QMessageBox.question(
                self, "Overwrite?",
                f"Overwrite {self.doc.path.name} in place?"
        ) != QMessageBox.StandardButton.Yes:
            return
        self._write(self.doc.path)

    def save_as(self):
        if not self.doc:
            return
        path, _f = QFileDialog.getSaveFileName(
            self, "Save Preflight profile", str(self.doc.path),
            "Preflight profiles (*.kfp)")
        if path:
            self._write(path)

    def _write(self, path):
        problems = self._validate()
        if problems:
            QMessageBox.warning(self, "Fix these first",
                                "\n".join(problems[:12]))
            return
        self._commit()
        try:
            self.doc.save(path)
        except Exception as exc:                       # noqa: BLE001
            QMessageBox.critical(self, "Save failed", str(exc))
            return
        self.statusBar().showMessage(f"Saved {path}")


def main():
    app = QApplication(sys.argv)
    win = Editor()
    if len(sys.argv) > 1:
        win.open_file(path=sys.argv[1])
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
