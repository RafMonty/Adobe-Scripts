import os
import sys
import json
import datetime
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLineEdit, QVBoxLayout, QListWidget, QListWidgetItem,
    QLabel, QPushButton, QHBoxLayout, QFrame, QDialog, QCheckBox, QFileDialog,
    QMenu, QSystemTrayIcon, QAction, QMessageBox, QComboBox
)
from PyQt5.QtCore import Qt, QTimer, QSize
from PyQt5.QtGui import QFontMetrics, QIcon, QKeySequence, QColor, QCursor
from PyQt5.QtCore import QStandardPaths

APP_NAME = "Folder Finder"
ORG_NAME = "PrintsByRaf"
BASE_WIDTH = 500
COLLAPSED_HEIGHT = 50

def user_config_path() -> Path:
    base = QStandardPaths.writableLocation(QStandardPaths.AppConfigLocation)
    full = Path(base) / ORG_NAME / APP_NAME
    full.mkdir(parents=True, exist_ok=True)
    return full / "folder_finder_config.json"

class FloatingSearchBar(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.config_file = user_config_path()
        self.settings = {
            'search_folders': [],
            'single_click_open': True,
            'shortcut_keys': True,
            'geometry': None,
            'theme': 'light'
        }
        self.load_config()
        self.search_results = []
        self.debounce = QTimer(self)
        self.debounce.setSingleShot(True)
        self.debounce.setInterval(200)
        self.debounce.timeout.connect(self._perform_search)
        self.init_ui()
        self.install_shortcuts()
        self.setup_tray()
        self.apply_theme()
        self.time_timer = QTimer(self)
        self.time_timer.timeout.connect(self.update_time)
        self.time_timer.start(1000)
        if self.settings.get("geometry"):
            x, y, w, h = self.settings["geometry"]
            self.move(x, y)
            self.resize(BASE_WIDTH, COLLAPSED_HEIGHT)
        else:
            self.center_on_screen()
        self.show()

    def init_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        self.search_frame = QFrame(self)
        self.search_frame.setObjectName("searchFrame")
        top = QHBoxLayout(self.search_frame)
        top.setContentsMargins(10, 5, 10, 5)
        self.search_icon = QLabel("🔍")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search folders…")
        self.search_input.setStyleSheet("QLineEdit { border:none; background:transparent; font-size:14px; padding:8px 0; }")
        self.search_input.textChanged.connect(self.on_search_change)
        self.time_label = QLabel()
        self.update_time()
        self.settings_btn = QPushButton("⋮")
        self.settings_btn.setFixedSize(24, 24)
        self.settings_btn.clicked.connect(self.show_context_menu)
        self.close_btn = QPushButton("✕")
        self.close_btn.setFixedSize(24, 24)
        self.close_btn.clicked.connect(self.request_quit)
        top.addWidget(self.search_icon)
        top.addWidget(self.search_input, 1)
        top.addWidget(self.time_label)
        top.addWidget(self.settings_btn)
        top.addWidget(self.close_btn)
        root.addWidget(self.search_frame)
        self.results_list = QListWidget()
        self.results_list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.results_list.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.results_list.hide()
        if self.settings.get('single_click_open', True):
            self.results_list.itemClicked.connect(self.on_result_activated)
        else:
            self.results_list.itemDoubleClicked.connect(self.on_result_activated)
        root.addWidget(self.results_list)
        self.setFixedWidth(BASE_WIDTH)
        self.setMinimumHeight(COLLAPSED_HEIGHT)
        self.resize(BASE_WIDTH, COLLAPSED_HEIGHT)

    def install_shortcuts(self):
        quit_act = QAction(self)
        quit_act.setShortcut(QKeySequence.Quit)
        quit_act.triggered.connect(self.request_quit)
        self.addAction(quit_act)
        # Focus search (Ctrl+L)
        focus_act = QAction(self)
        focus_act.setShortcut("Ctrl+L")
        focus_act.triggered.connect(lambda: (self.search_input.setFocus(), self.search_input.selectAll()))
        self.addAction(focus_act)

    def setup_tray(self):
        # Guard for environments without a system tray
        if not QSystemTrayIcon.isSystemTrayAvailable():
            return
        self.tray = QSystemTrayIcon(self)
        # Create a simple icon if none exists
        icon = self.windowIcon()
        if icon.isNull():
            # Create a simple colored icon as fallback
            from PyQt5.QtGui import QPixmap, QPainter, QBrush
            pixmap = QPixmap(16, 16)
            pixmap.fill(Qt.transparent)
            painter = QPainter(pixmap)
            painter.setBrush(QBrush(QColor(70, 130, 180)))  # Steel blue color
            painter.drawEllipse(2, 2, 12, 12)
            painter.end()
            icon = QIcon(pixmap)
        self.tray.setIcon(icon)
        self.tray.setToolTip(APP_NAME)
        menu = QMenu()
        menu.addAction("Show/Hide", self.toggle_visibility)
        menu.addAction("Settings…", self.show_settings_dialog)
        menu.addSeparator()
        menu.addAction("Quit", self.request_quit)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(lambda r: self.toggle_visibility() if r == QSystemTrayIcon.Trigger else None)
        self.tray.show()

    def toggle_visibility(self):
        if self.isVisible():
            self.hide()
        else:
            self.show()
            self.raise_()
            self.activateWindow()

    def show_context_menu(self):
        menu = QMenu(self)
        menu.addAction("Settings…", self.show_settings_dialog)
        menu.addSeparator()
        menu.addAction("Quit", self.request_quit)
        # Show menu at cursor position for expected UX
        menu.exec_(QCursor.pos())

    def theme_colors(self):
        if self.settings.get('theme', 'light') == 'dark':
            return {'panel': '#323232', 'border': '#3c3c3c', 'text': '#e8e8e8', 'muted': '#a0a0a0', 'selection': '#40465a'}
        return {'panel': '#ffffff', 'border': '#e0e0e0', 'text': '#000000', 'muted': '#777777', 'selection': '#f0f0f5'}

    def apply_theme(self):
        c = self.theme_colors()
        self.search_frame.setStyleSheet(f"#searchFrame {{background-color: {c['panel']}; border: 1px solid {c['border']}; border-radius: 8px;}} QLineEdit {{color:{c['text']};}}")
        self.time_label.setStyleSheet(f"color:{c['muted']};")
        self.search_icon.setStyleSheet(f"color:{c['muted']};")
        self.results_list.setStyleSheet(f"QListWidget {{background-color:{c['panel']}; border:1px solid {c['border']}; color:{c['text']};}} QListWidget::item:selected {{background-color:{c['selection']};}}")

    def update_time(self):
        self.time_label.setText(datetime.datetime.now().strftime("%I:%M %p"))
        self.time_label.setVisible(not self.search_input.text())

    def on_search_change(self):
        if self.search_input.text():
            self.time_label.hide()
            self.debounce.start()
        else:
            self.results_list.clear()
            self.results_list.hide()
            self.resize(BASE_WIDTH, COLLAPSED_HEIGHT)
            self.time_label.show()

    def _perform_search(self):
        text = self.search_input.text()
        self.results_list.clear()
        self.search_results = []
        if not text or not self.settings.get('search_folders'):
            self.results_list.hide()
            self.resize(BASE_WIDTH, COLLAPSED_HEIGHT)
            return
        matches = []
        for base in self.settings.get('search_folders', []):
            try:
                for entry in os.scandir(base):
                    if entry.is_dir() and text.lower() in entry.name.lower():
                        matches.append(entry.path)
            except Exception as e:
                print(f"Error searching {base}: {e}")
        fm = QFontMetrics(self.font())
        for path in matches:
            self.add_search_result(path, fm)
        if self.results_list.count() > 0:
            self.results_list.show()
            num_results = min(6, self.results_list.count())
            self.resize(BASE_WIDTH, COLLAPSED_HEIGHT + num_results * 60)
            self.results_list.setCurrentRow(0)
        else:
            self.results_list.hide()
            self.resize(BASE_WIDTH, COLLAPSED_HEIGHT)

    def _elide_middle(self, text: str, fm: QFontMetrics, max_px: int) -> str:
        return fm.elidedText(text, Qt.ElideMiddle, max_px)

    def add_search_result(self, path, fm: QFontMetrics):
        result_widget = QWidget()
        v = QVBoxLayout(result_widget)
        v.setContentsMargins(5, 5, 5, 5)
        folder_name = os.path.basename(path)
        name_label = QLabel(folder_name)
        name_label.setStyleSheet("font-weight:bold;")
        parent_path = os.path.dirname(path)
        parent_elided = self._elide_middle(parent_path, fm, 420)
        path_label = QLabel(parent_elided)
        path_label.setStyleSheet("color:#777; font-size:12px;")
        v.addWidget(name_label)
        v.addWidget(path_label)
        item = QListWidgetItem()
        item.setSizeHint(QSize(self.results_list.width(), 60))
        self.results_list.addItem(item)
        self.results_list.setItemWidget(item, result_widget)
        self.search_results.append(path)

    def on_result_activated(self, item):
        index = self.results_list.row(item)
        if 0 <= index < len(self.search_results):
            self.open_folder(self.search_results[index])

    def open_folder(self, path):
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                from subprocess import Popen
                Popen(["open", path])
            else:
                from subprocess import Popen
                Popen(["xdg-open", path])
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Error opening folder: {e}")

    def center_on_screen(self):
        screen = QApplication.primaryScreen()
        if screen:
            geo = screen.availableGeometry()
            w, h = self.width(), self.height()
            x = geo.x() + (geo.width() - w) // 2
            y = geo.y() + (geo.height() - h) // 3
            self.move(x, y)

    def load_config(self):
        try:
            if self.config_file.exists():
                loaded = json.loads(self.config_file.read_text(encoding="utf-8"))
                loaded['search_folders'] = [f for f in loaded.get('search_folders', []) if os.path.exists(f)]
                if 'theme' in loaded:
                    loaded['theme'] = str(loaded['theme']).lower()
                self.settings.update(loaded)
        except Exception as e:
            print(f"Error loading configuration: {e}")

    def save_config(self):
        try:
            g = self.geometry()
            self.settings['geometry'] = [int(g.x()), int(g.y()), int(g.width()), COLLAPSED_HEIGHT]
            self.config_file.write_text(json.dumps(self.settings, indent=4), encoding="utf-8")
        except Exception as e:
            print(f"Error saving configuration: {e}")

    def show_settings_dialog(self):
        dlg = SettingsDialog(self.settings, self)
        if dlg.exec_():
            self.settings = dlg.get_settings()
            self.save_config()
            try:
                self.results_list.itemClicked.disconnect()
            except Exception:
                pass
            try:
                self.results_list.itemDoubleClicked.disconnect()
            except Exception:
                pass
            if self.settings.get('single_click_open', True):
                self.results_list.itemClicked.connect(self.on_result_activated)
            else:
                self.results_list.itemDoubleClicked.connect(self.on_result_activated)
            self.apply_theme()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            if self.search_input.text():
                self.search_input.clear()
            else:
                self.close()
            return
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            current_row = self.results_list.currentRow()
            if 0 <= current_row < len(self.search_results):
                self.open_folder(self.search_results[current_row])
            return
        if event.key() == Qt.Key_Down and self.results_list.count() > 0:
            r = self.results_list.currentRow()
            self.results_list.setCurrentRow((r + 1) % self.results_list.count())
            return
        if event.key() == Qt.Key_Up and self.results_list.count() > 0:
            r = self.results_list.currentRow()
            self.results_list.setCurrentRow((r - 1) % self.results_list.count())
            return
        if self.settings.get('shortcut_keys', True) and (event.modifiers() == Qt.AltModifier) and (Qt.Key_1 <= event.key() <= Qt.Key_9):
            idx = event.key() - Qt.Key_1
            if 0 <= idx < len(self.search_results):
                self.open_folder(self.search_results[idx])
            return
        super().keyPressEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_pos = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()
        elif event.button() == Qt.RightButton:
            self.show_context_menu()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton and hasattr(self, '_drag_pos'):
            self.move(event.globalPos() - self._drag_pos)
            event.accept()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and hasattr(self, '_drag_pos'):
            delattr(self, '_drag_pos')
            event.accept()

    def request_quit(self):
        self.close()

    def closeEvent(self, event):
        self.save_config()
        event.accept()

class SettingsDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Folder Finder Settings")
        self.setFixedSize(520, 380)
        self._settings = dict(settings)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Search Locations:"))
        self.folders_list = QListWidget()
        layout.addWidget(self.folders_list)
        for folder in self._settings.get('search_folders', []):
            self.folders_list.addItem(folder)
        buttons = QHBoxLayout()
        add_btn = QPushButton("Add Folder")
        rm_btn = QPushButton("Remove Selected")
        add_btn.clicked.connect(self.add_folder)
        rm_btn.clicked.connect(self.remove_folder)
        buttons.addWidget(add_btn)
        buttons.addWidget(rm_btn)
        layout.addLayout(buttons)
        layout.addWidget(QLabel("Options:"))
        theme_row = QHBoxLayout()
        theme_row.addWidget(QLabel("Theme:"))
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Light", "Dark"])
        idx = 1 if str(self._settings.get('theme', 'light')).lower() == 'dark' else 0
        self.theme_combo.setCurrentIndex(idx)
        theme_row.addWidget(self.theme_combo)
        layout.addLayout(theme_row)
        self.single_click_cb = QCheckBox("Single‑click to open folders")
        self.single_click_cb.setChecked(self._settings.get('single_click_open', True))
        self.shortcut_keys_cb = QCheckBox("Enable Alt+number shortcuts")
        self.shortcut_keys_cb.setChecked(self._settings.get('shortcut_keys', True))
        layout.addWidget(self.single_click_cb)
        layout.addWidget(self.shortcut_keys_cb)
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.accept)
        layout.addWidget(save_btn)

    def add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self.folders_list.addItem(folder)

    def remove_folder(self):
        for item in self.folders_list.selectedItems():
            self.folders_list.takeItem(self.folders_list.row(item))

    def get_settings(self):
        folders = [self.folders_list.item(i).text() for i in range(self.folders_list.count())]
        self._settings['search_folders'] = folders
        self._settings['single_click_open'] = self.single_click_cb.isChecked()
        self._settings['shortcut_keys'] = self.shortcut_keys_cb.isChecked()
        self._settings['theme'] = 'dark' if self.theme_combo.currentIndex() == 1 else 'light'
        return self._settings

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(ORG_NAME)
    app.setStyle("Fusion")
    w = FloatingSearchBar()
    sys.exit(app.exec_())