"""
main.py — MediSort with Batch Processing
"""

import sys
import os
import csv
import subprocess
from pathlib import Path
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTextEdit, QFrame,
    QDialog, QScrollArea, QSizePolicy, QProgressBar,
    QMessageBox, QStatusBar, QTabWidget, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView,
    QLineEdit, QCheckBox,
)
from PySide6.QtGui import (
    QPixmap, QFont, QColor, QPainter, QLinearGradient,
)
from PySide6.QtCore import (
    Qt, QThread, Signal, QSize, QTimer,
)

from backend import DocumentClassifier, BatchWorker, CATEGORIES, SUPPORTED_EXTENSIONS

# ─────────────────────────────────────────────
# Constants & Theme
# ─────────────────────────────────────────────
APP_TITLE   = "MediSort — Medical Document Classifier"
WINDOW_W, WINDOW_H = 1150, 820

THEME = {
    "bg":          "#0F1923",
    "surface":     "#162030",
    "surface2":    "#1C2A3E",
    "border":      "#243347",
    "accent":      "#00C9A7",
    "accent2":     "#0082FF",
    "danger":      "#FF5C5C",
    "warning":     "#FFAA00",
    "text":        "#E8F0FE",
    "text_dim":    "#7A93B0",
    "success":     "#4CAF82",
}

CATEGORY_COLORS = {
    "prescription":      "#00C9A7",
    "lab_report":        "#0082FF",
    "discharge_summary": "#A855F7",
    "medical_bill":      "#FFAA00",
    "insurance_claim":   "#FF7043",
    "referral_letter":   "#26C6DA",
    "other":             "#7A93B0",
}

CATEGORY_ICONS = {
    "prescription":      "💊",
    "lab_report":        "🧪",
    "discharge_summary": "🏥",
    "medical_bill":      "💳",
    "insurance_claim":   "📋",
    "referral_letter":   "📨",
    "other":             "📄",
}

STYLESHEET = f"""
QMainWindow, QWidget {{
    background-color: {THEME['bg']};
    color: {THEME['text']};
    font-family: 'Segoe UI', 'SF Pro Display', sans-serif;
}}
QFrame#card {{
    background-color: {THEME['surface']};
    border: 1px solid {THEME['border']};
    border-radius: 12px;
}}
QFrame#preview_card {{
    background-color: {THEME['surface']};
    border: 2px dashed {THEME['border']};
    border-radius: 12px;
}}
QTabWidget::pane {{
    border: 1px solid {THEME['border']};
    border-radius: 10px;
    background: {THEME['surface']};
    top: -1px;
}}
QTabBar::tab {{
    background: {THEME['surface2']};
    color: {THEME['text_dim']};
    border: 1px solid {THEME['border']};
    border-bottom: none;
    border-radius: 8px 8px 0 0;
    padding: 10px 28px;
    font-size: 13px;
    font-weight: 600;
    margin-right: 4px;
}}
QTabBar::tab:selected {{
    background: {THEME['surface']};
    color: {THEME['accent']};
    border-color: {THEME['accent']};
}}
QTabBar::tab:hover:!selected {{
    color: {THEME['text']};
    background: {THEME['surface']};
}}
QPushButton#primary_btn {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 {THEME['accent']}, stop:1 {THEME['accent2']});
    color: #0F1923;
    border: none;
    border-radius: 10px;
    padding: 12px 28px;
    font-size: 14px;
    font-weight: 700;
}}
QPushButton#primary_btn:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #00e5be, stop:1 #3399FF);
}}
QPushButton#primary_btn:disabled {{
    background: {THEME['border']};
    color: {THEME['text_dim']};
}}
QPushButton#secondary_btn {{
    background-color: transparent;
    color: {THEME['accent']};
    border: 2px solid {THEME['accent']};
    border-radius: 10px;
    padding: 10px 24px;
    font-size: 13px;
    font-weight: 600;
}}
QPushButton#secondary_btn:hover {{
    background-color: rgba(0,201,167,0.1);
}}
QPushButton#secondary_btn:disabled {{
    color: {THEME['text_dim']};
    border-color: {THEME['border']};
}}
QPushButton#danger_btn {{
    background-color: transparent;
    color: {THEME['danger']};
    border: 2px solid {THEME['danger']};
    border-radius: 10px;
    padding: 10px 24px;
    font-size: 13px;
    font-weight: 600;
}}
QPushButton#danger_btn:hover {{
    background-color: rgba(255,92,92,0.1);
}}
QPushButton#ghost_btn {{
    background-color: transparent;
    color: {THEME['text_dim']};
    border: 1px solid {THEME['border']};
    border-radius: 8px;
    padding: 8px 18px;
    font-size: 13px;
}}
QPushButton#ghost_btn:hover {{
    color: {THEME['text']};
    border-color: {THEME['text_dim']};
    background: rgba(255,255,255,0.04);
}}
QPushButton#ghost_btn:disabled {{ color: {THEME['border']}; }}
QScrollArea {{ border: none; background: transparent; }}
QProgressBar {{
    border: none;
    background-color: {THEME['border']};
    border-radius: 4px;
    height: 8px;
    text-align: center;
    color: transparent;
}}
QProgressBar::chunk {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 {THEME['accent']}, stop:1 {THEME['accent2']});
    border-radius: 4px;
}}
QStatusBar {{
    background-color: {THEME['surface']};
    color: {THEME['text_dim']};
    font-size: 12px;
    border-top: 1px solid {THEME['border']};
}}
QTableWidget {{
    background-color: {THEME['surface']};
    color: {THEME['text']};
    border: none;
    gridline-color: {THEME['border']};
    font-size: 12px;
    alternate-background-color: {THEME['surface2']};
}}
QTableWidget::item {{ padding: 6px 10px; }}
QTableWidget::item:selected {{
    background-color: rgba(0,201,167,0.15);
    color: {THEME['text']};
}}
QHeaderView::section {{
    background-color: {THEME['surface2']};
    color: {THEME['text_dim']};
    border: none;
    border-bottom: 1px solid {THEME['border']};
    padding: 8px 10px;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.8px;
    text-transform: uppercase;
}}
QLineEdit {{
    background: {THEME['surface2']};
    border: 1px solid {THEME['border']};
    border-radius: 8px;
    color: {THEME['text']};
    padding: 8px 14px;
    font-size: 13px;
}}
QLineEdit:focus {{ border-color: {THEME['accent']}; }}
QTextEdit {{
    background-color: {THEME['surface2']};
    color: {THEME['text']};
    border: 1px solid {THEME['border']};
    border-radius: 8px;
    padding: 10px;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 12px;
}}
QLabel#title {{ font-size: 24px; font-weight: 800; color: {THEME['text']}; }}
QLabel#subtitle {{ font-size: 12px; color: {THEME['text_dim']}; }}
QLabel#section_label {{
    font-size: 10px; font-weight: 700;
    color: {THEME['text_dim']}; letter-spacing: 1.4px;
}}
"""


# ─────────────────────────────────────────────
# Worker: single-file
# ─────────────────────────────────────────────
class ProcessWorker(QThread):
    finished = Signal(dict)
    error    = Signal(str)

    def __init__(self, classifier, file_path):
        super().__init__()
        self.classifier = classifier
        self.file_path  = file_path

    def run(self):
        try:
            self.finished.emit(self.classifier.process_document(self.file_path))
        except Exception as e:
            self.error.emit(str(e))


# ─────────────────────────────────────────────
# OCR Text Dialog
# ─────────────────────────────────────────────
class OCRTextDialog(QDialog):
    def __init__(self, text, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Extracted OCR Text")
        self.setMinimumSize(600, 480)
        self.setStyleSheet(f"QDialog {{ background:{THEME['surface']}; color:{THEME['text']}; }}"
                           f"QPushButton {{ background:{THEME['accent']}; color:#0F1923; border:none;"
                           f" border-radius:6px; padding:8px 24px; font-weight:600; }}")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        header = QLabel("  OCR Extracted Text")
        header.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        layout.addWidget(header)
        te = QTextEdit()
        te.setReadOnly(True)
        te.setPlainText(text or "(No text extracted)")
        te.setMinimumHeight(340)
        layout.addWidget(te)
        btn = QPushButton("Close")
        btn.clicked.connect(self.accept)
        btn.setFixedWidth(100)
        row = QHBoxLayout()
        row.addStretch()
        row.addWidget(btn)
        layout.addLayout(row)


# ─────────────────────────────────────────────
# Document Preview (drag-and-drop)
# ─────────────────────────────────────────────
class DocumentPreview(QLabel):
    file_dropped = Signal(str)
    SUPPORTED = SUPPORTED_EXTENSIONS

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setMinimumHeight(300)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self._show_placeholder()

    def _show_placeholder(self):
        self.setText(
            f"<div style='text-align:center;'>"
            f"<p style='font-size:40px; margin:0;'>📄</p>"
            f"<p style='font-size:13px; color:{THEME['text_dim']}; margin:8px 0 0 0;'>"
            f"Drag &amp; drop a file here<br>or click <b>Upload Document</b></p></div>"
        )
        self.setStyleSheet("color: #E8F0FE; background: transparent;")

    def set_image(self, path):
        ext = Path(path).suffix.lower()
        if ext == ".pdf":
            self.setText(
                f"<div style='text-align:center;'>"
                f"<p style='font-size:48px; margin:0;'>📄</p>"
                f"<p style='font-size:13px; color:{THEME['text']}; margin:8px 0 0;'>"
                f"<b>{Path(path).name}</b></p>"
                f"<p style='font-size:11px; color:{THEME['text_dim']};'>PDF Document</p></div>"
            )
        else:
            pix = QPixmap(path)
            if not pix.isNull():
                self.setPixmap(pix.scaled(self.width()-40, self.height()-40,
                    Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
            else:
                self._show_placeholder()

    def reset(self):
        self.clear()
        self._show_placeholder()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            p = event.mimeData().urls()[0].toLocalFile()
            if Path(p).suffix.lower() in self.SUPPORTED:
                event.acceptProposedAction()
                self.setStyleSheet(
                    f"border:2px dashed {THEME['accent']}; border-radius:12px;"
                    f" background:rgba(0,201,167,0.06);")

    def dragLeaveEvent(self, event):
        self.setStyleSheet("background: transparent;")

    def dropEvent(self, event):
        self.setStyleSheet("background: transparent;")
        p = event.mimeData().urls()[0].toLocalFile()
        if Path(p).suffix.lower() in self.SUPPORTED:
            self.file_dropped.emit(p)


# ─────────────────────────────────────────────
# Category Badge
# ─────────────────────────────────────────────
class CategoryBadge(QFrame):
    def __init__(self):
        super().__init__()
        self.setObjectName("card")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(24, 18, 24, 18)
        layout.setSpacing(16)
        self.icon_label = QLabel("🩺")
        self.icon_label.setFont(QFont("Segoe UI", 30))
        self.icon_label.setFixedWidth(50)
        right = QVBoxLayout()
        right.setSpacing(3)
        self.cat_label = QLabel("Awaiting document")
        self.cat_label.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        self.cat_label.setStyleSheet(f"color:{THEME['text_dim']};")
        self.conf_label = QLabel("")
        self.conf_label.setStyleSheet(f"color:{THEME['text_dim']}; font-size:12px;")
        self.saved_label = QLabel("")
        self.saved_label.setStyleSheet(f"color:{THEME['text_dim']}; font-size:11px;")
        self.saved_label.setWordWrap(True)
        right.addWidget(self.cat_label)
        right.addWidget(self.conf_label)
        right.addWidget(self.saved_label)
        layout.addWidget(self.icon_label)
        layout.addLayout(right)
        layout.addStretch()

    def show_result(self, category, confidence, saved_path):
        color = CATEGORY_COLORS.get(category, THEME["text_dim"])
        self.icon_label.setText(CATEGORY_ICONS.get(category, "📄"))
        self.cat_label.setText(category.replace("_", " ").title())
        self.cat_label.setStyleSheet(f"color:{color}; font-size:18px; font-weight:700;")
        self.conf_label.setText(f"Confidence: {confidence:.1%}" if confidence else "")
        self.saved_label.setText(f"✓  Saved → {saved_path}" if saved_path else "")
        self.setStyleSheet(f"QFrame#card {{ background:{THEME['surface']}; border:2px solid {color}; border-radius:12px; }}")

    def show_error(self, msg):
        self.icon_label.setText("⚠️")
        self.cat_label.setText("Error")
        self.cat_label.setStyleSheet(f"color:{THEME['danger']};")
        self.conf_label.setText(msg)
        self.saved_label.setText("")
        self.setStyleSheet(f"QFrame#card {{ background:{THEME['surface']}; border:2px solid {THEME['danger']}; border-radius:12px; }}")

    def reset(self):
        self.icon_label.setText("🩺")
        self.cat_label.setText("Awaiting document")
        self.cat_label.setStyleSheet(f"color:{THEME['text_dim']};")
        self.conf_label.setText("")
        self.saved_label.setText("")
        self.setStyleSheet(f"QFrame#card {{ background:{THEME['surface']}; border:1px solid {THEME['border']}; border-radius:12px; }}")


# ─────────────────────────────────────────────
# SINGLE FILE TAB
# ─────────────────────────────────────────────
class SingleFileTab(QWidget):
    def __init__(self, classifier, status_bar, parent=None):
        super().__init__(parent)
        self.classifier  = classifier
        self.status_bar  = status_bar
        self._current_file = None
        self._ocr_text   = ""
        self._worker     = None
        self._build_ui()

    def _build_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(20)

        # ── Left: preview ──
        left = QVBoxLayout()
        lbl = QLabel("DOCUMENT PREVIEW")
        lbl.setObjectName("section_label")
        left.addWidget(lbl)

        preview_card = QFrame()
        preview_card.setObjectName("preview_card")
        pc_layout = QVBoxLayout(preview_card)
        pc_layout.setContentsMargins(12, 12, 12, 12)
        self.preview = DocumentPreview()
        self.preview.file_dropped.connect(self._on_file_selected)
        pc_layout.addWidget(self.preview)
        left.addWidget(preview_card, 1)

        self.file_info = QLabel("No file selected")
        self.file_info.setStyleSheet(f"color:{THEME['text_dim']}; font-size:12px;")
        self.file_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        left.addWidget(self.file_info)

        btn_row = QHBoxLayout()
        self.upload_btn = QPushButton("⬆  Upload Document")
        self.upload_btn.setObjectName("primary_btn")
        self.upload_btn.setFixedHeight(44)
        self.upload_btn.clicked.connect(self._browse)
        self.classify_btn = QPushButton("⚡  Classify")
        self.classify_btn.setObjectName("secondary_btn")
        self.classify_btn.setFixedHeight(44)
        self.classify_btn.setEnabled(False)
        self.classify_btn.clicked.connect(self._classify)
        btn_row.addWidget(self.upload_btn)
        btn_row.addWidget(self.classify_btn)
        left.addLayout(btn_row)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setFixedHeight(6)
        self.progress.setVisible(False)
        left.addWidget(self.progress)

        root.addLayout(left, 3)

        # ── Right: results ──
        right = QVBoxLayout()

        res_lbl = QLabel("CLASSIFICATION RESULT")
        res_lbl.setObjectName("section_label")
        right.addWidget(res_lbl)

        self.badge = CategoryBadge()
        right.addWidget(self.badge)

        legend_lbl = QLabel("CATEGORIES")
        legend_lbl.setObjectName("section_label")
        right.addWidget(legend_lbl)

        legend_card = QFrame()
        legend_card.setObjectName("card")
        lc = QVBoxLayout(legend_card)
        lc.setContentsMargins(16, 12, 16, 12)
        lc.setSpacing(6)
        for cat in CATEGORIES:
            color = CATEGORY_COLORS[cat]
            row = QLabel(f'{CATEGORY_ICONS[cat]}  <span style="color:{color}; font-weight:600;">'
                         f'{cat.replace("_"," ").title()}</span>')
            row.setTextFormat(Qt.TextFormat.RichText)
            row.setStyleSheet("font-size:13px;")
            lc.addWidget(row)
        right.addWidget(legend_card)

        util_row = QHBoxLayout()
        self.ocr_btn = QPushButton("📝  View OCR Text")
        self.ocr_btn.setObjectName("ghost_btn")
        self.ocr_btn.setEnabled(False)
        self.ocr_btn.clicked.connect(self._show_ocr)
        self.folder_btn = QPushButton("📁  Output Folder")
        self.folder_btn.setObjectName("ghost_btn")
        self.folder_btn.clicked.connect(self._open_folder)
        util_row.addWidget(self.ocr_btn)
        util_row.addWidget(self.folder_btn)
        right.addLayout(util_row)
        right.addStretch()

        root.addLayout(right, 2)

    def _browse(self):
        exts = "Documents (*.pdf *.png *.jpg *.jpeg *.tiff *.tif *.bmp *.webp);;All Files (*)"
        path, _ = QFileDialog.getOpenFileName(self, "Select Document", "", exts)
        if path:
            self._on_file_selected(path)

    def _on_file_selected(self, path):
        self._current_file = path
        fname = Path(path).name
        fsize = Path(path).stat().st_size / 1024
        self.file_info.setText(f"{fname}   •   {fsize:.1f} KB")
        self.preview.set_image(path)
        self.classify_btn.setEnabled(True)
        self.ocr_btn.setEnabled(False)
        self._ocr_text = ""
        self.badge.reset()
        self.status_bar.showMessage(f"Loaded: {fname}")

    def _classify(self):
        if not self._current_file:
            return
        self.upload_btn.setEnabled(False)
        self.classify_btn.setEnabled(False)
        self.classify_btn.setText("Processing…")
        self.progress.setVisible(True)
        self.status_bar.showMessage("Running OCR and classification…")
        self._worker = ProcessWorker(self.classifier, self._current_file)
        self._worker.finished.connect(self._on_done)
        self._worker.error.connect(self._on_error)
        self._worker.start()

    def _on_done(self, result):
        self._restore_ui()
        self._ocr_text = result.get("ocr_text", "")
        cat = result.get("category", "other")
        conf = result.get("confidence", 0.0)
        saved = result.get("saved_path", "")
        err = result.get("error", "")
        if err and not result.get("success"):
            self.badge.show_error(err)
            self.status_bar.showMessage(f"Error: {err}")
        else:
            self.badge.show_result(cat, conf, saved)
            self.status_bar.showMessage(
                f"✓ '{cat.replace('_',' ').title()}'"
                + (f" ({conf:.1%})" if conf else "") + f" — saved to '{cat}' folder.")
        self.ocr_btn.setEnabled(bool(self._ocr_text))

    def _on_error(self, msg):
        self._restore_ui()
        self.badge.show_error(msg)
        self.status_bar.showMessage(f"Error: {msg}")

    def _restore_ui(self):
        self.upload_btn.setEnabled(True)
        self.classify_btn.setEnabled(True)
        self.classify_btn.setText("⚡  Classify")
        self.progress.setVisible(False)

    def _show_ocr(self):
        OCRTextDialog(self._ocr_text, self).exec()

    def _open_folder(self):
        folder = self.classifier.get_output_directory()
        Path(folder).mkdir(parents=True, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(folder)
        elif sys.platform == "darwin":
            subprocess.run(["open", folder])
        else:
            subprocess.run(["xdg-open", folder])


# ─────────────────────────────────────────────
# BATCH TAB
# ─────────────────────────────────────────────
class BatchTab(QWidget):
    def __init__(self, classifier, status_bar, parent=None):
        super().__init__(parent)
        self.classifier = classifier
        self.status_bar = status_bar
        self._worker    = None
        self._files: list[str] = []
        self._results: list[dict] = []
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(14)

        # ── Top: source selection row ──
        src_card = QFrame()
        src_card.setObjectName("card")
        src_layout = QVBoxLayout(src_card)
        src_layout.setContentsMargins(20, 16, 20, 16)
        src_layout.setSpacing(10)

        src_hdr = QLabel("BATCH SOURCE")
        src_hdr.setObjectName("section_label")
        src_layout.addWidget(src_hdr)

        path_row = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("Select a folder containing medical documents…")
        self.path_edit.setReadOnly(True)

        self.browse_folder_btn = QPushButton("📁  Browse Folder")
        self.browse_folder_btn.setObjectName("primary_btn")
        self.browse_folder_btn.setFixedHeight(38)
        self.browse_folder_btn.clicked.connect(self._browse_folder)

        self.browse_files_btn = QPushButton("📄  Select Files")
        self.browse_files_btn.setObjectName("secondary_btn")
        self.browse_files_btn.setFixedHeight(38)
        self.browse_files_btn.clicked.connect(self._browse_files)

        path_row.addWidget(self.path_edit, 1)
        path_row.addWidget(self.browse_folder_btn)
        path_row.addWidget(self.browse_files_btn)
        src_layout.addLayout(path_row)

        info_row = QHBoxLayout()
        self.lbl_file_count = QLabel("No files selected")
        self.lbl_file_count.setStyleSheet(f"color:{THEME['text_dim']}; font-size:12px;")
        self.chk_subfolders = QCheckBox("Include subfolders")
        self.chk_subfolders.setStyleSheet(f"color:{THEME['text_dim']}; font-size:12px;")
        self.chk_subfolders.stateChanged.connect(self._refresh_folder_files)
        info_row.addWidget(self.lbl_file_count)
        info_row.addStretch()
        info_row.addWidget(self.chk_subfolders)
        src_layout.addLayout(info_row)

        root.addWidget(src_card)

        # ── Middle: progress + stats ──
        mid_row = QHBoxLayout()

        # Progress card
        prog_card = QFrame()
        prog_card.setObjectName("card")
        prog_layout = QVBoxLayout(prog_card)
        prog_layout.setContentsMargins(20, 14, 20, 14)
        prog_layout.setSpacing(8)
        prog_lbl = QLabel("PROGRESS")
        prog_lbl.setObjectName("section_label")
        prog_layout.addWidget(prog_lbl)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(10)
        self.progress_bar.setValue(0)
        prog_layout.addWidget(self.progress_bar)

        self.lbl_progress = QLabel("Ready")
        self.lbl_progress.setStyleSheet(f"color:{THEME['text_dim']}; font-size:12px;")
        prog_layout.addWidget(self.lbl_progress)

        # Stats row inside progress card
        stats_row = QHBoxLayout()
        self.stat_success = self._make_stat("✓ Classified", "0", THEME["success"])
        self.stat_failed  = self._make_stat("✗ Failed", "0", THEME["danger"])
        self.stat_review  = self._make_stat("⚠ Review", "0", THEME["warning"])
        stats_row.addWidget(self.stat_success)
        stats_row.addWidget(self.stat_failed)
        stats_row.addWidget(self.stat_review)
        prog_layout.addLayout(stats_row)

        mid_row.addWidget(prog_card, 2)

        # Action card
        act_card = QFrame()
        act_card.setObjectName("card")
        act_layout = QVBoxLayout(act_card)
        act_layout.setContentsMargins(20, 14, 20, 14)
        act_layout.setSpacing(10)
        act_lbl = QLabel("ACTIONS")
        act_lbl.setObjectName("section_label")
        act_layout.addWidget(act_lbl)

        self.start_btn = QPushButton("▶  Start Batch")
        self.start_btn.setObjectName("primary_btn")
        self.start_btn.setFixedHeight(42)
        self.start_btn.setEnabled(False)
        self.start_btn.clicked.connect(self._start_batch)

        self.cancel_btn = QPushButton("⏹  Cancel")
        self.cancel_btn.setObjectName("danger_btn")
        self.cancel_btn.setFixedHeight(42)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._cancel_batch)

        self.export_btn = QPushButton("⬇  Export CSV")
        self.export_btn.setObjectName("ghost_btn")
        self.export_btn.setFixedHeight(38)
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self._export_csv)

        self.clear_btn = QPushButton("🗑  Clear Results")
        self.clear_btn.setObjectName("ghost_btn")
        self.clear_btn.setFixedHeight(38)
        self.clear_btn.clicked.connect(self._clear)

        act_layout.addWidget(self.start_btn)
        act_layout.addWidget(self.cancel_btn)
        act_layout.addWidget(self.export_btn)
        act_layout.addWidget(self.clear_btn)
        act_layout.addStretch()

        mid_row.addWidget(act_card, 1)
        root.addLayout(mid_row)

        # ── Bottom: results table ──
        tbl_lbl = QLabel("RESULTS")
        tbl_lbl.setObjectName("section_label")
        root.addWidget(tbl_lbl)

        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Filename", "Category", "Confidence", "Status", "Saved To"])
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setShowGrid(False)
        self.table.verticalHeader().setVisible(False)
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        root.addWidget(self.table, 1)

        # Current file label
        self.lbl_current = QLabel("")
        self.lbl_current.setStyleSheet(f"color:{THEME['text_dim']}; font-size:11px;")
        root.addWidget(self.lbl_current)

    def _make_stat(self, label, value, color):
        """Small stat widget."""
        frame = QFrame()
        frame.setObjectName("card")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(2)
        val_lbl = QLabel(value)
        val_lbl.setStyleSheet(f"color:{color}; font-size:20px; font-weight:700;")
        lbl_lbl = QLabel(label)
        lbl_lbl.setStyleSheet(f"color:{THEME['text_dim']}; font-size:10px;")
        layout.addWidget(val_lbl)
        layout.addWidget(lbl_lbl)
        frame._val_lbl = val_lbl
        return frame

    def _update_stat(self, stat_frame, value):
        stat_frame._val_lbl.setText(str(value))

    # ── File selection ──────────────────────────────────────────────────────

    def _browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder with Medical Documents")
        if folder:
            self.path_edit.setText(folder)
            self._load_folder_files(folder)

    def _refresh_folder_files(self):
        folder = self.path_edit.text()
        if folder and Path(folder).is_dir():
            self._load_folder_files(folder)

    def _load_folder_files(self, folder: str):
        p = Path(folder)
        pattern = "**/*" if self.chk_subfolders.isChecked() else "*"
        self._files = [
            str(f) for f in p.glob(pattern)
            if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS
        ]
        self._files.sort()
        self._update_file_count()

    def _browse_files(self):
        exts = "Documents (*.pdf *.png *.jpg *.jpeg *.tiff *.tif *.bmp *.webp);;All Files (*)"
        paths, _ = QFileDialog.getOpenFileNames(self, "Select Medical Documents", "", exts)
        if paths:
            self._files = paths
            self.path_edit.setText(f"{len(paths)} files selected individually")
            self._update_file_count()

    def _update_file_count(self):
        n = len(self._files)
        self.lbl_file_count.setText(
            f"{n} document{'s' if n != 1 else ''} ready for processing"
            if n else "No supported files found in selected location"
        )
        self.start_btn.setEnabled(n > 0)

    # ── Batch control ───────────────────────────────────────────────────────

    def _start_batch(self):
        if not self._files:
            return
        self._results.clear()
        self.table.setRowCount(0)
        self.progress_bar.setMaximum(len(self._files))
        self.progress_bar.setValue(0)
        self._update_stat(self.stat_success, 0)
        self._update_stat(self.stat_failed, 0)
        self._update_stat(self.stat_review, 0)
        self.lbl_progress.setText(f"Starting — 0 / {len(self._files)}")

        self.start_btn.setEnabled(False)
        self.browse_folder_btn.setEnabled(False)
        self.browse_files_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.export_btn.setEnabled(False)

        self._worker = BatchWorker(self.classifier, self._files)
        self._worker.file_started.connect(self._on_file_started)
        self._worker.file_done.connect(self._on_file_done)
        self._worker.batch_finished.connect(self._on_batch_finished)
        self._worker.start()
        self.status_bar.showMessage(f"Batch processing {len(self._files)} documents…")

    def _cancel_batch(self):
        if self._worker:
            self._worker.cancel()
        self.cancel_btn.setEnabled(False)
        self.lbl_progress.setText("Cancelling…")
        self.status_bar.showMessage("Cancelling batch…")

    def _on_file_started(self, idx, total, filename):
        self.lbl_current.setText(f"Processing ({idx}/{total}): {filename}")
        self.lbl_progress.setText(f"{idx - 1} / {total} complete")
        self.progress_bar.setValue(idx - 1)

    def _on_file_done(self, idx, total, result):
        self._results.append(result)
        self.progress_bar.setValue(idx)
        self.lbl_progress.setText(f"{idx} / {total} complete")

        # Update stats
        success_count = sum(1 for r in self._results if r.get("success"))
        failed_count  = sum(1 for r in self._results if not r.get("success"))
        review_count  = sum(1 for r in self._results if r.get("needs_review"))
        self._update_stat(self.stat_success, success_count)
        self._update_stat(self.stat_failed, failed_count)
        self._update_stat(self.stat_review, review_count)

        # Add row to table
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setRowHeight(row, 36)

        filename = result.get("filename", "")
        cat      = result.get("category", "other")
        conf     = result.get("confidence", 0.0)
        saved    = result.get("saved_path", "")
        err      = result.get("error", "")
        success  = result.get("success", False)
        needs_rv = result.get("needs_review", False)

        color = CATEGORY_COLORS.get(cat, THEME["text_dim"])

        def cell(text, fg=None):
            item = QTableWidgetItem(text)
            if fg:
                item.setForeground(QColor(fg))
            return item

        self.table.setItem(row, 0, cell(filename))
        self.table.setItem(row, 1, cell(f"{CATEGORY_ICONS.get(cat,'📄')} {cat.replace('_',' ').title()}", color))
        self.table.setItem(row, 2, cell(f"{conf:.1%}" if conf else "—"))

        if err and not success:
            self.table.setItem(row, 3, cell("✗ Error", THEME["danger"]))
        elif needs_rv:
            self.table.setItem(row, 3, cell("⚠ Review", THEME["warning"]))
        else:
            self.table.setItem(row, 3, cell("✓ OK", THEME["success"]))

        self.table.setItem(row, 4, cell(saved if saved else err[:60] if err else "—", THEME["text_dim"]))
        self.table.scrollToBottom()

    def _on_batch_finished(self, summary):
        total   = summary["total"]
        success = summary["success"]
        failed  = summary["failed"]
        self.lbl_current.setText("")
        self.lbl_progress.setText(
            f"Done — {success}/{total} classified"
            + (f", {failed} failed" if failed else "")
        )
        self.progress_bar.setValue(total)
        self.start_btn.setEnabled(True)
        self.browse_folder_btn.setEnabled(True)
        self.browse_files_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.export_btn.setEnabled(bool(self._results))
        self.status_bar.showMessage(
            f"✓ Batch complete — {success} classified, {failed} failed"
        )

    # ── Export / Clear ──────────────────────────────────────────────────────

    def _export_csv(self):
        if not self._results:
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"batch_results_{ts}.csv"
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Results as CSV", default_name, "CSV Files (*.csv)"
        )
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=[
                "filename", "category", "confidence", "needs_review",
                "saved_path", "error", "file_path"
            ])
            writer.writeheader()
            for r in self._results:
                writer.writerow({
                    "filename":     r.get("filename", ""),
                    "category":     r.get("category", ""),
                    "confidence":   f"{r.get('confidence', 0.0):.4f}",
                    "needs_review": r.get("needs_review", False),
                    "saved_path":   r.get("saved_path", ""),
                    "error":        r.get("error", ""),
                    "file_path":    r.get("file_path", ""),
                })
        self.status_bar.showMessage(f"✓ Exported {len(self._results)} rows → {path}")

    def _clear(self):
        self._results.clear()
        self._files.clear()
        self.table.setRowCount(0)
        self.path_edit.clear()
        self.lbl_file_count.setText("No files selected")
        self.lbl_progress.setText("Ready")
        self.progress_bar.setValue(0)
        self._update_stat(self.stat_success, 0)
        self._update_stat(self.stat_failed, 0)
        self._update_stat(self.stat_review, 0)
        self.start_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.status_bar.showMessage("Cleared.")


# ─────────────────────────────────────────────
# Main Window
# ─────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.classifier = DocumentClassifier()
        self.setWindowTitle(APP_TITLE)
        self.resize(WINDOW_W, WINDOW_H)
        self.setMinimumSize(900, 650)
        self.setStyleSheet(STYLESHEET)
        self._build_ui()
        QTimer.singleShot(200, self._load_model)

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # ── Header ──
        header = QFrame()
        header.setFixedHeight(68)
        header.setStyleSheet(f"QFrame {{ background:{THEME['surface']}; border-bottom:1px solid {THEME['border']}; }}")
        h_layout = QHBoxLayout(header)
        h_layout.setContentsMargins(28, 0, 28, 0)

        logo_row = QHBoxLayout()
        logo_row.setSpacing(10)
        logo = QLabel("🩺")
        logo.setFont(QFont("Segoe UI", 24))
        logo_row.addWidget(logo)
        title_stack = QVBoxLayout()
        title_stack.setSpacing(1)
        title = QLabel("MediSort")
        title.setObjectName("title")
        title.setFont(QFont("Segoe UI", 17, QFont.Weight.Bold))
        subtitle = QLabel("Medical Document Classifier")
        subtitle.setObjectName("subtitle")
        title_stack.addWidget(title)
        title_stack.addWidget(subtitle)
        logo_row.addLayout(title_stack)
        h_layout.addLayout(logo_row)
        h_layout.addStretch()

        self.model_status_lbl = QLabel("⚙ Loading…")
        self.model_status_lbl.setStyleSheet(f"color:{THEME['warning']}; font-size:12px;")
        h_layout.addWidget(self.model_status_lbl)
        layout.addWidget(header)

        # ── Tabs ──
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready — upload a document or start a batch.")

        tabs = QTabWidget()
        tabs.setDocumentMode(False)

        self.single_tab = SingleFileTab(self.classifier, self.status_bar)
        self.batch_tab  = BatchTab(self.classifier, self.status_bar)

        tabs.addTab(self.single_tab, "  📄  Single Document  ")
        tabs.addTab(self.batch_tab,  "  📦  Batch Processing  ")

        layout.addWidget(tabs)

    def _load_model(self):
        success, msg = self.classifier.initialize()
        if success:
            self.model_status_lbl.setText("✓ Classifier ready")
            self.model_status_lbl.setStyleSheet(f"color:{THEME['accent']}; font-size:12px;")
            self.status_bar.showMessage("Classifier ready. Upload a document or start a batch.")
        else:
            self.model_status_lbl.setText("⚠ Model not found")
            self.model_status_lbl.setStyleSheet(f"color:{THEME['warning']}; font-size:12px;")
            self.status_bar.showMessage(f"Warning: {msg}")


# ─────────────────────────────────────────────
# Entry Point
# ─────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setApplicationName("MediSort")
    app.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
