import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QLabel, QFileDialog, QFrame, QGridLayout)
from PyQt6.QtGui import QPixmap, QFont
from PyQt6.QtCore import Qt

class PrintCalcApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Print Quality & Resolution Guide")
        self.setMinimumSize(500, 600)
        
        # Modern "Material" Style
        self.setStyleSheet("""
            QMainWindow { background-color: #f5f5f5; }
            QLabel { color: #333; font-family: 'Segoe UI', Arial; }
            .Card { 
                background-color: white; 
                border-radius: 10px; 
                border: 1px solid #ddd;
            }
            .Header { font-size: 18px; font-weight: bold; color: #1a73e8; }
            .DimLabel { font-size: 14px; color: #666; }
            QPushButton {
                background-color: #1a73e8;
                color: white;
                border-radius: 5px;
                padding: 10px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover { background-color: #1557b0; }
            .ResultRow { border-bottom: 1px solid #eee; padding: 10px; }
            .High { color: #2e7d32; font-weight: bold; } /* Green */
            .Mid { color: #f57c00; font-weight: bold; }  /* Orange */
            .Low { color: #d32f2f; font-weight: bold; }   /* Red */
        """)

        # Main Layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.main_layout = QVBoxLayout(central_widget)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(15)

        # Header Section
        self.lbl_title = QLabel("Image Print Calculator")
        self.lbl_title.setProperty("class", "Header")
        self.main_layout.addWidget(self.lbl_title)

        # File Loader
        self.btn_load = QPushButton("Select Image for Audit")
        self.btn_load.clicked.connect(self.load_image)
        self.main_layout.addWidget(self.btn_load)

        # Image Info Card
        self.info_card = QFrame()
        self.info_card.setProperty("class", "Card")
        self.info_card_layout = QVBoxLayout(self.info_card)
        
        self.lbl_filename = QLabel("No image loaded")
        self.lbl_dims = QLabel("0 x 0 pixels")
        self.lbl_dims.setProperty("class", "DimLabel")
        
        self.info_card_layout.addWidget(self.lbl_filename)
        self.info_card_layout.addWidget(self.lbl_dims)
        self.main_layout.addWidget(self.info_card)

        # Results Grid
        self.results_frame = QFrame()
        self.results_frame.setProperty("class", "Card")
        self.grid = QGridLayout(self.results_frame)
        self.grid.setSpacing(10)
        
        # Table Headers
        headers = ["Quality", "DPI", "Max Size (Inches)", "Max Size (CM)"]
        for i, h in enumerate(headers):
            lbl = QLabel(h)
            lbl.setStyleSheet("font-weight: bold; color: #888; font-size: 12px;")
            self.grid.addWidget(lbl, 0, i)

        self.main_layout.addWidget(self.results_frame)
        self.main_layout.addStretch()

    def load_image(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Image", "", "Images (*.png *.jpg *.jpeg *.tif *.bmp)")
        if file_path:
            pixmap = QPixmap(file_path)
            w = pixmap.width()
            h = pixmap.height()
            
            self.lbl_filename.setText(file_path.split("/")[-1])
            self.lbl_dims.setText(f"Original Dimensions: {w}px x {h}px")
            
            self.update_results(w, h)

    def update_results(self, w, h):
        # Clear previous rows (except headers)
        for i in reversed(range(self.grid.count())): 
            if i > 3: # Keep headers
                self.grid.itemAt(i).widget().setParent(None)

        # Quality Levels: (Label, DPI, CSS Class)
        levels = [
            ("Professional", 300, "High"),
            ("Good Print", 150, "High"),
            ("Acceptable", 100, "Mid"),
            ("Signboard/Web", 72, "Low")
        ]

        for row, (label, dpi, css) in enumerate(levels, start=1):
            # Calculations
            w_in = w / dpi
            h_in = h / dpi
            w_cm = w_in * 2.54
            h_cm = h_in * 2.54

            q_lbl = QLabel(label)
            q_lbl.setProperty("class", css)
            
            self.grid.addWidget(q_lbl, row, 0)
            self.grid.addWidget(QLabel(f"{dpi}"), row, 1)
            self.grid.addWidget(QLabel(f"{w_in:.1f}\" x {h_in:.1f}\""), row, 2)
            self.grid.addWidget(QLabel(f"{w_cm:.1f} x {h_cm:.1f} cm"), row, 3)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PrintCalcApp()
    window.show()
    sys.exit(app.exec())
