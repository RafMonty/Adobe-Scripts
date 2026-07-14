"""
AgentPicCrawler v2 — High-Res Staff Image Extractor
====================================================
Recursively scans a directory tree for high-resolution images, copies (or moves)
them to a flat destination folder named after their parent folder.

v2 changes over v1:
  - Persistent settings via QSettings (all fields restored between sessions)
  - Fixed crash: exception handlers referenced `file.name` on a str
  - Whole-word folder exclusion matching ("old" no longer matches "Harold")
  - Resolution modes: Both / Either / Shortest-side >= threshold
  - EXIF orientation-aware dimension reads (rotated portraits handled)
  - Smarter folder-name cleaning (year stripping, separator cleanup)
  - Idempotent re-runs: identical files (size + partial hash) are skipped
  - Copy or Move mode, plus Dry Run preview
  - Safe cancel and safe window close (no QThread-destroyed crash)
  - Throttled progress updates, dest-inside-src guard

Requires:  pip install PyQt6 Pillow
"""

import sys
import os
import re
import shutil
import hashlib
from pathlib import Path

from PIL import Image, UnidentifiedImageError
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QLineEdit, QPushButton,
                             QFileDialog, QSpinBox, QProgressBar, QTextEdit,
                             QGroupBox, QFormLayout, QComboBox, QCheckBox,
                             QRadioButton, QButtonGroup, QMessageBox)
from PyQt6.QtCore import QThread, pyqtSignal, QSettings

ORIENTATION_TAG = 274  # EXIF orientation
SWAPPED_ORIENTATIONS = {5, 6, 7, 8}  # 90/270-degree rotations


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def clean_folder_name(name: str, strip_digits: bool) -> str:
    """Clean a parent folder name into a person/staff name.

    - Optionally removes standalone 4-digit years and stray digit runs
    - Collapses whitespace, trims leading/trailing separators
    """
    cleaned = name
    if strip_digits:
        # Remove standalone years first (e.g. "John Smith 2023")
        cleaned = re.sub(r'\b(19|20)\d{2}\b', '', cleaned)
        # Remove remaining standalone digit runs (e.g. "John Smith 2")
        cleaned = re.sub(r'\b\d+\b', '', cleaned)
    # Collapse whitespace, trim separator junk left behind
    cleaned = re.sub(r'\s+', ' ', cleaned)
    cleaned = cleaned.strip(' -_.,')
    return cleaned or "Unknown_Staff"


def folder_excluded(folder: str, exclusions: list[str]) -> bool:
    """Whole-word, case-insensitive exclusion match.

    "old" matches "Old Photos" and "old" but NOT "Harold" or "Goldsmith".
    Multi-word exclusions ("shop photo") match as phrases.
    """
    low = folder.lower()
    for excl in exclusions:
        if re.search(r'(?<!\w)' + re.escape(excl) + r'(?!\w)', low):
            return True
    return False


def read_dimensions(path: Path) -> tuple[int, int]:
    """Read pixel dimensions from the image header, honouring EXIF rotation."""
    with Image.open(path) as img:
        width, height = img.size
        try:
            orientation = img.getexif().get(ORIENTATION_TAG)
            if orientation in SWAPPED_ORIENTATIONS:
                width, height = height, width
        except Exception:
            pass  # No/broken EXIF: use raw dimensions
    return width, height


def quick_hash(path: Path, chunk: int = 65536) -> str:
    """Hash of first + last 64KB — fast near-identity check for large images."""
    h = hashlib.sha1()
    size = path.stat().st_size
    with open(path, 'rb') as f:
        h.update(f.read(chunk))
        if size > chunk * 2:
            f.seek(-chunk, os.SEEK_END)
            h.update(f.read(chunk))
    h.update(str(size).encode())
    return h.hexdigest()


def files_identical(a: Path, b: Path) -> bool:
    try:
        if a.stat().st_size != b.stat().st_size:
            return False
        return quick_hash(a) == quick_hash(b)
    except OSError:
        return False


# ---------------------------------------------------------------------------
# Worker thread
# ---------------------------------------------------------------------------

class ScannerThread(QThread):
    progress_update = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished_scan = pyqtSignal(int, int, int)  # copied, skipped_dupes, errors

    MODE_BOTH = 0      # width AND height >= thresholds
    MODE_EITHER = 1    # width OR height >= thresholds
    MODE_SHORTEST = 2  # min(width, height) >= min_w (min_h ignored)

    def __init__(self, src, dest, min_w, min_h, res_mode, exclusions,
                 extensions, strip_digits, skip_identical, move_files, dry_run):
        super().__init__()
        self.src = src
        self.dest = dest
        self.min_w = min_w
        self.min_h = min_h
        self.res_mode = res_mode
        self.exclusions = [e.strip().lower() for e in exclusions.split(',') if e.strip()]
        self.extensions = {self._norm_ext(e) for e in extensions.split(',') if e.strip()}
        self.strip_digits = strip_digits
        self.skip_identical = skip_identical
        self.move_files = move_files
        self.dry_run = dry_run

        self.is_running = True
        self.copied = 0
        self.dupes = 0
        self.errors = 0
        self._last_pct = -1

    @staticmethod
    def _norm_ext(ext: str) -> str:
        ext = ext.strip().lower()
        return ext if ext.startswith('.') else '.' + ext

    def stop(self):
        self.is_running = False

    def _emit_progress(self, processed, total):
        pct = int((processed / total) * 100)
        if pct != self._last_pct:  # Throttle: only emit on change
            self._last_pct = pct
            self.progress_update.emit(pct)

    def _passes_resolution(self, w, h) -> bool:
        if self.res_mode == self.MODE_BOTH:
            return w >= self.min_w and h >= self.min_h
        if self.res_mode == self.MODE_EITHER:
            return w >= self.min_w or h >= self.min_h
        return min(w, h) >= self.min_w  # MODE_SHORTEST

    def _walk(self):
        """os.walk with in-place whole-word directory pruning."""
        for root, dirs, files in os.walk(self.src):
            dirs[:] = [d for d in dirs if not folder_excluded(d, self.exclusions)]
            yield root, files

    def run(self):
        self.log_message.emit("Counting files..." if not self.dry_run
                              else "Counting files (DRY RUN — nothing will be written)...")

        total_files = sum(len(files) for _, files in self._walk())
        if total_files == 0:
            self.log_message.emit("No files found in the source directory.")
            self.progress_update.emit(100)
            self.finished_scan.emit(0, 0, 0)
            return

        self.log_message.emit(f"Scanning {total_files:,} files...")
        processed = 0
        dest_path = Path(self.dest)

        for root, files in self._walk():
            if not self.is_running:
                break
            for filename in files:
                if not self.is_running:
                    break
                processed += 1
                self._emit_progress(processed, total_files)

                file_path = Path(root) / filename
                if file_path.suffix.lower() not in self.extensions:
                    continue

                try:
                    width, height = read_dimensions(file_path)
                except (UnidentifiedImageError, OSError):
                    self.log_message.emit(f"Skipped {filename}: unreadable image.")
                    self.errors += 1
                    continue
                except Exception as e:
                    self.log_message.emit(f"Error reading {filename}: {e}")
                    self.errors += 1
                    continue

                if self._passes_resolution(width, height):
                    self.process_file(file_path, dest_path, width, height)

        self.progress_update.emit(100)
        self.finished_scan.emit(self.copied, self.dupes, self.errors)

    def process_file(self, src_path: Path, dest_dir: Path, width: int, height: int):
        clean_name = clean_folder_name(src_path.parent.name, self.strip_digits)
        ext = src_path.suffix.lower()

        final_dest = dest_dir / f"{clean_name}{ext}"
        counter = 1
        while final_dest.exists():
            if self.skip_identical and files_identical(src_path, final_dest):
                self.dupes += 1
                self.log_message.emit(f"Skipped (identical exists): {final_dest.name}")
                return
            final_dest = dest_dir / f"{clean_name}_{counter}{ext}"
            counter += 1

        verb = "Would move" if (self.dry_run and self.move_files) else \
               "Would copy" if self.dry_run else \
               "Moved" if self.move_files else "Copied"

        if self.dry_run:
            self.copied += 1
            self.log_message.emit(
                f"{verb}: {src_path.name} ({width}x{height}) -> {final_dest.name}")
            return

        try:
            if self.move_files:
                shutil.move(str(src_path), str(final_dest))
            else:
                shutil.copy2(src_path, final_dest)
            self.copied += 1
            self.log_message.emit(
                f"{verb}: {clean_name} ({width}x{height}) -> {final_dest.name}")
        except Exception as e:
            self.errors += 1
            self.log_message.emit(f"Failed on {src_path.name}: {e}")


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class ImageCopierApp(QMainWindow):
    ORG, APP = "Raf", "AgentPicCrawler"

    def __init__(self):
        super().__init__()
        self.setWindowTitle("High-Res Staff Image Extractor v2")
        self.resize(780, 640)
        self.thread = None
        self.settings = QSettings(self.ORG, self.APP)
        self.setup_ui()
        self.load_settings()

    # -- UI construction ----------------------------------------------------

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # Directories
        dir_group = QGroupBox("Directories")
        dir_layout = QVBoxLayout(dir_group)

        src_layout = QHBoxLayout()
        self.src_input = QLineEdit()
        self.src_input.setPlaceholderText("Root directory to scan (e.g. Staff Headshots)")
        src_btn = QPushButton("Browse...")
        src_btn.clicked.connect(lambda: self.browse_into(self.src_input, "Select Source Directory"))
        src_layout.addWidget(QLabel("Source:"))
        src_layout.addWidget(self.src_input)
        src_layout.addWidget(src_btn)

        dest_layout = QHBoxLayout()
        self.dest_input = QLineEdit()
        self.dest_input.setPlaceholderText("Flat output directory")
        dest_btn = QPushButton("Browse...")
        dest_btn.clicked.connect(lambda: self.browse_into(self.dest_input, "Select Destination Directory"))
        dest_layout.addWidget(QLabel("Destination:"))
        dest_layout.addWidget(self.dest_input)
        dest_layout.addWidget(dest_btn)

        dir_layout.addLayout(src_layout)
        dir_layout.addLayout(dest_layout)
        main_layout.addWidget(dir_group)

        # Parameters
        param_group = QGroupBox("Scan Parameters")
        param_layout = QFormLayout(param_group)

        res_layout = QHBoxLayout()
        self.min_width = QSpinBox()
        self.min_width.setRange(100, 20000)
        self.min_width.setValue(1500)
        self.min_height = QSpinBox()
        self.min_height.setRange(100, 20000)
        self.min_height.setValue(1500)
        self.res_mode = QComboBox()
        self.res_mode.addItems([
            "Both dimensions \u2265 min",
            "Either dimension \u2265 min",
            "Shortest side \u2265 min width",
        ])
        self.res_mode.currentIndexChanged.connect(
            lambda i: self.min_height.setEnabled(i != ScannerThread.MODE_SHORTEST))
        res_layout.addWidget(QLabel("W:"))
        res_layout.addWidget(self.min_width)
        res_layout.addWidget(QLabel("H:"))
        res_layout.addWidget(self.min_height)
        res_layout.addWidget(self.res_mode)
        param_layout.addRow("Resolution:", res_layout)

        self.ext_input = QLineEdit(".jpg, .jpeg, .png, .tif, .tiff")
        param_layout.addRow("Extensions:", self.ext_input)

        self.excl_input = QLineEdit("old, shop photo, web, lifestyle, temporary, superseded, archived")
        self.excl_input.setToolTip(
            "Comma-separated. Whole-word match: 'old' skips 'Old Photos' but not 'Harold'.")
        param_layout.addRow("Skip folders containing:", self.excl_input)

        main_layout.addWidget(param_group)

        # Options
        opt_group = QGroupBox("Options")
        opt_layout = QHBoxLayout(opt_group)

        self.strip_digits_cb = QCheckBox("Strip years/digits from names")
        self.strip_digits_cb.setChecked(True)
        self.skip_identical_cb = QCheckBox("Skip identical files (safe re-runs)")
        self.skip_identical_cb.setChecked(True)
        self.dry_run_cb = QCheckBox("Dry run (preview only)")

        self.copy_radio = QRadioButton("Copy")
        self.move_radio = QRadioButton("Move")
        self.copy_radio.setChecked(True)
        mode_group = QButtonGroup(self)
        mode_group.addButton(self.copy_radio)
        mode_group.addButton(self.move_radio)

        opt_layout.addWidget(self.copy_radio)
        opt_layout.addWidget(self.move_radio)
        opt_layout.addSpacing(20)
        opt_layout.addWidget(self.strip_digits_cb)
        opt_layout.addWidget(self.skip_identical_cb)
        opt_layout.addWidget(self.dry_run_cb)
        opt_layout.addStretch()
        main_layout.addWidget(opt_group)

        # Progress + log
        self.progress_bar = QProgressBar()
        main_layout.addWidget(self.progress_bar)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        main_layout.addWidget(self.log_output)

        # Buttons
        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("Start Scan")
        self.start_btn.setMinimumHeight(40)
        self.start_btn.clicked.connect(self.start_scan)
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setMinimumHeight(40)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self.cancel_scan)
        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(self.cancel_btn)
        main_layout.addLayout(btn_layout)

    # -- Settings persistence ------------------------------------------------

    def load_settings(self):
        s = self.settings
        self.src_input.setText(s.value("src", ""))
        self.dest_input.setText(s.value("dest", ""))
        self.min_width.setValue(int(s.value("min_w", 1500)))
        self.min_height.setValue(int(s.value("min_h", 1500)))
        self.res_mode.setCurrentIndex(int(s.value("res_mode", 0)))
        self.ext_input.setText(s.value("extensions", self.ext_input.text()))
        self.excl_input.setText(s.value("exclusions", self.excl_input.text()))
        self.strip_digits_cb.setChecked(s.value("strip_digits", "true") == "true")
        self.skip_identical_cb.setChecked(s.value("skip_identical", "true") == "true")
        self.move_radio.setChecked(s.value("move_files", "false") == "true")
        self.copy_radio.setChecked(not self.move_radio.isChecked())

    def save_settings(self):
        s = self.settings
        s.setValue("src", self.src_input.text())
        s.setValue("dest", self.dest_input.text())
        s.setValue("min_w", self.min_width.value())
        s.setValue("min_h", self.min_height.value())
        s.setValue("res_mode", self.res_mode.currentIndex())
        s.setValue("extensions", self.ext_input.text())
        s.setValue("exclusions", self.excl_input.text())
        s.setValue("strip_digits", "true" if self.strip_digits_cb.isChecked() else "false")
        s.setValue("skip_identical", "true" if self.skip_identical_cb.isChecked() else "false")
        s.setValue("move_files", "true" if self.move_radio.isChecked() else "false")

    # -- Actions --------------------------------------------------------------

    def browse_into(self, line_edit: QLineEdit, title: str):
        start = line_edit.text() or self.settings.value("last_browse", "")
        folder = QFileDialog.getExistingDirectory(self, title, start)
        if folder:
            line_edit.setText(folder)
            self.settings.setValue("last_browse", folder)

    def log(self, message: str):
        self.log_output.append(message)
        sb = self.log_output.verticalScrollBar()
        sb.setValue(sb.maximum())

    def start_scan(self):
        src = self.src_input.text().strip()
        dest = self.dest_input.text().strip()

        if not src or not dest:
            self.log("ERROR: Select both source and destination directories.")
            return
        if not os.path.isdir(src):
            self.log("ERROR: Source directory does not exist.")
            return

        # Guard: destination inside source would re-scan its own output
        try:
            src_res = Path(src).resolve()
            dest_res = Path(dest).resolve()
            if dest_res == src_res or src_res in dest_res.parents:
                reply = QMessageBox.warning(
                    self, "Destination inside source",
                    "The destination folder is inside the source tree. "
                    "The scanner may pick up its own output.\n\nContinue anyway?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No)
                if reply != QMessageBox.StandardButton.Yes:
                    return
        except OSError:
            pass

        if self.move_radio.isChecked() and not self.dry_run_cb.isChecked():
            reply = QMessageBox.question(
                self, "Confirm Move",
                "Move mode will remove originals from the source tree. Continue?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No)
            if reply != QMessageBox.StandardButton.Yes:
                return

        if not self.dry_run_cb.isChecked():
            os.makedirs(dest, exist_ok=True)

        self.save_settings()
        self.progress_bar.setValue(0)
        self.log_output.clear()
        self.log("Starting scan..." + (" [DRY RUN]" if self.dry_run_cb.isChecked() else ""))
        self.start_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)

        self.thread = ScannerThread(
            src=src,
            dest=dest,
            min_w=self.min_width.value(),
            min_h=self.min_height.value(),
            res_mode=self.res_mode.currentIndex(),
            exclusions=self.excl_input.text(),
            extensions=self.ext_input.text(),
            strip_digits=self.strip_digits_cb.isChecked(),
            skip_identical=self.skip_identical_cb.isChecked(),
            move_files=self.move_radio.isChecked(),
            dry_run=self.dry_run_cb.isChecked(),
        )
        self.thread.progress_update.connect(self.progress_bar.setValue)
        self.thread.log_message.connect(self.log)
        self.thread.finished_scan.connect(self.scan_complete)
        self.thread.start()

    def cancel_scan(self):
        if self.thread and self.thread.isRunning():
            self.log("Cancelling... finishing current file.")
            self.thread.stop()
            self.cancel_btn.setEnabled(False)

    def scan_complete(self, copied: int, dupes: int, errors: int):
        verb = "would be processed" if self.dry_run_cb.isChecked() else "processed"
        self.log("\n--- Scan Complete ---")
        self.log(f"{copied} file(s) {verb}, {dupes} identical duplicate(s) skipped, "
                 f"{errors} error(s).")
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)

    # -- Safe shutdown ----------------------------------------------------------

    def closeEvent(self, event):
        self.save_settings()
        if self.thread and self.thread.isRunning():
            self.thread.stop()
            self.thread.wait(5000)  # Give it up to 5s to exit cleanly
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    if sys.platform == "win32":
        app.setStyle("windowsvista")  # Native light styling on Windows
    window = ImageCopierApp()
    window.show()
    sys.exit(app.exec())
