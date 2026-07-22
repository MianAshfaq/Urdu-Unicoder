
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import traceback
import unicodedata
import urllib.request
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF
from PySide6.QtCore import Qt, QThread, Signal, QSize, QTimer, QMarginsF, QUrl
from PySide6.QtGui import (
    QAction, QDesktopServices, QFont, QIcon, QImage, QPixmap, QTextDocument,
    QPageSize, QPageLayout, QTextCursor,
)
from PySide6.QtPrintSupport import QPrinter
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QMessageBox, QWidget, QHBoxLayout,
    QVBoxLayout, QListWidget, QListWidgetItem, QLabel, QPushButton, QSpinBox,
    QDoubleSpinBox, QComboBox, QLineEdit, QCheckBox, QFormLayout, QGroupBox,
    QProgressBar, QSplitter, QPlainTextEdit, QTextBrowser, QTabWidget,
    QToolBar, QScrollArea, QDialog, QDialogButtonBox, QInputDialog
)

APP_NAME = "Urdu Unicoder"
APP_VERSION = "1.3.0"
APP_AUTHOR = "Muhammad Ashfaq"
AUTHOR_GITHUB = "https://github.com/MianAshfaq"
AUTHOR_WEBSITE = "https://cyberoly.com/"
AUTHOR_FACEBOOK = "https://www.facebook.com/MianAshfaq012"
PROJECT_EXT = ".ubp"
WINDOWS_APP_ID = "MuhammadAshfaq.UrduUnicoder"
UPDATE_MANIFEST_URL = "https://raw.githubusercontent.com/MianAshfaq/Urdu-Unicoder/main/version.json"
UPDATE_PAGE_URL = "https://github.com/MianAshfaq/Urdu-Unicoder/releases/latest"


def resource_path(relative_path: str) -> Path:
    """Resolve bundled assets in source checkouts and PyInstaller builds."""
    bundle_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return bundle_root / relative_path


LOGO_PATH = resource_path("assets/urdu-unicoder-logo-final.png")
ICON_PATH = resource_path("assets/urdu-unicoder.ico")


def version_tuple(version: str) -> tuple[int, ...]:
    numbers = re.findall(r"\d+", version)
    return tuple(int(number) for number in numbers[:4]) or (0,)


def set_windows_app_identity():
    """Make Windows group the app under Urdu Unicoder rather than pythonw.exe."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(WINDOWS_APP_ID)
    except Exception:
        pass


def extract_docx_text(path: str) -> str:
    """Extract paragraphs and tables from a DOCX in their original body order."""
    try:
        from docx import Document as WordDocument
        from docx.oxml.table import CT_Tbl
        from docx.oxml.text.paragraph import CT_P
        from docx.table import Table
        from docx.text.paragraph import Paragraph
    except ImportError as exc:
        raise RuntimeError(
            "Word import requires python-docx. Run setup_windows.bat to install it."
        ) from exc

    document = WordDocument(path)
    lines: list[str] = []

    def add_blank():
        if lines and lines[-1] != "":
            lines.append("")

    for element in document.element.body.iterchildren():
        if isinstance(element, CT_P):
            paragraph = Paragraph(element, document)
            text = paragraph.text.strip()
            style_name = paragraph.style.name if paragraph.style is not None else ""
            if not text:
                add_blank()
                continue
            if style_name.startswith("Heading") or style_name in {"Title", "Subtitle"}:
                add_blank()
                lines.append(text)
                add_blank()
            elif style_name.startswith("List Bullet"):
                lines.append(f"• {text}")
            elif style_name.startswith("List Number"):
                lines.append(text)
            else:
                lines.append(text)
                add_blank()
        elif isinstance(element, CT_Tbl):
            table = Table(element, document)
            add_blank()
            for row in table.rows:
                cells = [re.sub(r"\s+", " ", cell.text).strip() for cell in row.cells]
                if any(cells):
                    lines.append(" | ".join(cells))
            add_blank()

    while lines and lines[-1] == "":
        lines.pop()
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()


@dataclass
class LayoutSettings:
    font_family: str = "Noto Nastaliq Urdu"
    font_size: float = 16.0
    line_height: float = 2.0
    paragraph_spacing: float = 8.0
    page_size: str = "A5"
    margin_top_mm: float = 18.0
    margin_bottom_mm: float = 18.0
    margin_inner_mm: float = 20.0
    margin_outer_mm: float = 15.0
    remove_jari_hai: bool = True
    join_wrapped_lines: bool = True
    preserve_blank_lines: bool = True
    page_numbers: bool = True
    running_header: str = ""
    recover_visual_order: bool = True
    normalize_presentation_forms: bool = True
    continuous_paragraph_mode: bool = True
    first_line_indent_mm: float = 0.0
    word_spacing_px: float = 0.0
    letter_spacing_px: float = 0.0
    text_align: str = "Justify"
    auto_paragraph_lines: int = 5


class ExtractWorker(QThread):
    progress = Signal(int, int, str)
    finished_ok = Signal(str, list)
    failed = Signal(str)

    def __init__(self, pdf_path: str, start_page: int, end_page: int, settings: LayoutSettings):
        super().__init__()
        self.pdf_path = pdf_path
        self.start_page = start_page
        self.end_page = end_page
        self.settings = settings

    @staticmethod
    def contains_presentation_forms(text: str) -> bool:
        return any(
            0xFB50 <= ord(ch) <= 0xFDFF or 0xFE70 <= ord(ch) <= 0xFEFF
            for ch in text
        )

    @staticmethod
    def normalize_glyphs(text: str) -> str:
        return unicodedata.normalize("NFKC", text)

    @staticmethod
    def _tokenize_visual_line(line: str) -> list[str]:
        line = re.sub(r"(?:\.\s*){2,}", " … ", line)
        line = re.sub(r"(?:…\s*){2,}", " … ", line)
        return re.findall(r"…|[^\s]+", line)

    @classmethod
    def recover_visual_order_line(cls, line: str) -> str:
        had_forms = cls.contains_presentation_forms(line)
        normalized = cls.normalize_glyphs(line)
        normalized = normalized.replace("\u200f", "").replace("\u200e", "")
        normalized = re.sub(r"[ \t]+", " ", normalized).strip()

        if had_forms:
            tokens = cls._tokenize_visual_line(normalized)
            if len(tokens) > 1:
                tokens.reverse()
                normalized = " ".join(tokens)

        normalized = re.sub(r"\s+([،؛؟!۔,:;])", r"\1", normalized)
        normalized = re.sub(r"([“‘])\s+", r"\1", normalized)
        normalized = re.sub(r"\s+([”’])", r"\1", normalized)
        normalized = re.sub(r"\s*…\s*", "… ", normalized).strip()
        return normalized

    def clean_line(self, line: str) -> str:
        if self.settings.normalize_presentation_forms or self.settings.recover_visual_order:
            if self.settings.recover_visual_order:
                return self.recover_visual_order_line(line)
            line = self.normalize_glyphs(line)

        line = line.replace("\u200f", "").replace("\u200e", "")
        line = re.sub(r"[ \t]+", " ", line).strip()
        line = re.sub(r"\s+([،؛؟!۔,:;])", r"\1", line)
        return line

    @staticmethod
    def looks_like_heading(line: str) -> bool:
        s = line.strip()
        if not s or len(s) > 100:
            return False
        explicit = (
            "باب ", "باب نمبر", "حصہ ", "فصل ", "عنوان ",
            "مقدمہ", "دیباچہ", "انتساب", "پیش لفظ"
        )
        if s in {"مقدمہ", "دیباچہ", "انتساب", "پیش لفظ"}:
            return True
        return any(s.startswith(prefix) for prefix in explicit)

    @staticmethod
    def is_dialogue_start(line: str) -> bool:
        s = line.strip()
        return s.startswith(("“", "\"", "’", "‘", "—", "-", "؎"))

    def reconstruct(self, lines: list[str]) -> str:
        out: list[str] = []
        buffer: list[str] = []

        def flush():
            nonlocal buffer
            if buffer:
                para = " ".join(buffer)
                para = re.sub(r"\s{2,}", " ", para).strip()
                if para:
                    out.append(para)
                buffer = []

        for raw in lines:
            line = self.clean_line(raw)

            if self.settings.remove_jari_hai and re.fullmatch(r"جاری\s+ہے[۔.!…]*", line):
                continue

            if not line:
                # Legacy PDFs often insert a blank line after every visual text line.
                # In continuous paragraph mode these are layout artefacts, not real paragraphs.
                if self.settings.continuous_paragraph_mode:
                    continue
                flush()
                if self.settings.preserve_blank_lines and (not out or out[-1] != ""):
                    out.append("")
                continue

            if not self.settings.join_wrapped_lines:
                flush()
                out.append(line)
                continue

            if self.looks_like_heading(line):
                flush()
                out.append(line)
                out.append("")
                continue

            if self.is_dialogue_start(line):
                flush()
                out.append(line)
                continue

            if buffer and not self.settings.continuous_paragraph_mode:
                prev = buffer[-1]
                if re.search(r"[۔؟!…][”\"']?$", prev) and len(" ".join(buffer)) > 120:
                    flush()

            buffer.append(line)

        flush()
        while out and out[-1] == "":
            out.pop()
        return "\n".join(out)

    def run(self):
        try:
            doc = fitz.open(self.pdf_path)
            page_texts = []
            all_lines = []
            total = self.end_page - self.start_page + 1
            for idx, pno in enumerate(range(self.start_page - 1, self.end_page), start=1):
                page = doc.load_page(pno)
                blocks = page.get_text("blocks", sort=True)
                text = "\n".join((b[4] or "") for b in blocks)
                lines = text.splitlines()
                page_texts.append(text)
                all_lines.extend(lines)
                all_lines.append("")
                self.progress.emit(idx, total, f"Extracting page {pno + 1}")
            doc.close()
            rebuilt = self.reconstruct(all_lines)
            self.finished_ok.emit(rebuilt, page_texts)
        except Exception:
            self.failed.emit(traceback.format_exc())


class PdfPageWorker(QThread):
    done = Signal(int, QImage)
    failed = Signal(str)

    def __init__(self, pdf_path: str, page_no: int, zoom: float = 1.25):
        super().__init__()
        self.pdf_path = pdf_path
        self.page_no = page_no
        self.zoom = zoom

    def run(self):
        try:
            doc = fitz.open(self.pdf_path)
            page = doc.load_page(self.page_no)
            pix = page.get_pixmap(matrix=fitz.Matrix(self.zoom, self.zoom), alpha=False)
            image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
            doc.close()
            self.done.emit(self.page_no, image)
        except Exception:
            self.failed.emit(traceback.format_exc())


class UpdateCheckWorker(QThread):
    update_info = Signal(dict)
    failed = Signal(str)

    def run(self):
        try:
            request = urllib.request.Request(
                UPDATE_MANIFEST_URL,
                headers={"User-Agent": f"{APP_NAME}/{APP_VERSION}"},
            )
            with urllib.request.urlopen(request, timeout=8) as response:
                data = json.loads(response.read().decode("utf-8"))
            if not isinstance(data, dict) or not data.get("version"):
                raise ValueError("The update manifest is invalid.")
            self.update_info.emit(data)
        except Exception as exc:
            self.failed.emit(str(exc))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        if ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(ICON_PATH)))
        self.resize(1450, 900)

        self.pdf_path: Optional[str] = None
        self.project_path: Optional[str] = None
        self.page_count = 0
        self.extracted_text = ""
        self.original_page_texts: list[str] = []
        self.worker: Optional[ExtractWorker] = None
        self.page_worker: Optional[PdfPageWorker] = None
        self.update_worker: Optional[UpdateCheckWorker] = None
        self.update_check_manual = False
        self.settings = LayoutSettings()

        self._build_ui()
        self._build_actions()
        self._apply_dark_theme()
        self.statusBar().showMessage("Ready")
        if not os.environ.get("URDU_UNICODER_DISABLE_UPDATE_CHECK"):
            QTimer.singleShot(2500, self.check_for_updates)

    def _build_actions(self):
        tb = QToolBar("Main")
        tb.setIconSize(QSize(22, 22))
        self.addToolBar(tb)

        if LOGO_PATH.exists():
            brand = QLabel()
            brand.setPixmap(QPixmap(str(LOGO_PATH)).scaled(
                30, 30, Qt.KeepAspectRatio, Qt.SmoothTransformation
            ))
            brand.setToolTip(f"{APP_NAME} {APP_VERSION}")
            brand.setContentsMargins(4, 0, 8, 0)
            tb.addWidget(brand)

        open_action = QAction("Open PDF", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_pdf)
        tb.addAction(open_action)
        tb.addAction(self.import_word_action)

        open_project_action = QAction("Open Project", self)
        open_project_action.setShortcut("Ctrl+Shift+O")
        open_project_action.triggered.connect(self.open_project)

        new_action = QAction("New Project", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self.new_project)
        tb.addAction(new_action)

        save_action = QAction("Save Project", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_project)
        tb.addAction(save_action)

        tb.addSeparator()

        extract_action = QAction("Reconstruct", self)
        extract_action.setShortcut("Ctrl+R")
        extract_action.triggered.connect(self.start_extract)
        tb.addAction(extract_action)

        join_selection_action = QAction("Join Selection", self)
        join_selection_action.triggered.connect(self.join_selected_lines_as_paragraph)
        tb.addAction(join_selection_action)

        split_selection_action = QAction("Split Selection", self)
        split_selection_action.triggered.connect(self.split_selected_text_at_sentence_end)
        tb.addAction(split_selection_action)

        export_action = QAction("Export PDF", self)
        export_action.setShortcut("Ctrl+Shift+P")
        export_action.triggered.connect(self.export_pdf)
        tb.addAction(export_action)

        html_action = QAction("Export HTML", self)
        html_action.triggered.connect(self.export_html)
        tb.addAction(html_action)

        file_menu = self.menuBar().addMenu("&File")
        file_menu.addAction(new_action)
        file_menu.addAction(open_action)
        file_menu.addAction(open_project_action)
        file_menu.addAction(save_action)
        file_menu.addSeparator()
        file_menu.addAction(self.import_word_action)
        file_menu.addAction(self.import_text_action)
        file_menu.addAction(self.save_text_action)
        file_menu.addSeparator()
        file_menu.addAction(export_action)
        file_menu.addAction(html_action)
        file_menu.addSeparator()
        exit_action = QAction("Exit", self)
        exit_action.setShortcut("Alt+F4")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        edit_menu = self.menuBar().addMenu("&Edit")
        for action in [
            self.undo_action, self.redo_action, None,
            self.cut_action, self.copy_action, self.paste_action,
            self.paste_unicode_action, self.select_all_action, None, self.find_action,
            self.goto_action, None, self.normalize_action,
            self.stats_action,
        ]:
            if action is None:
                edit_menu.addSeparator()
            else:
                edit_menu.addAction(action)

        view_menu = self.menuBar().addMenu("&View")
        view_menu.addAction(self.zoom_in_action)
        view_menu.addAction(self.zoom_out_action)
        view_menu.addAction(self.zoom_reset_action)
        view_menu.addSeparator()
        view_menu.addAction(self.rtl_action)
        view_menu.addAction(self.ltr_action)

        tools_menu = self.menuBar().addMenu("&Text Tools")
        tools_menu.addAction(self.paste_unicode_action)
        tools_menu.addAction(self.normalize_action)
        tools_menu.addAction(self.recover_order_action)
        tools_menu.addAction(self.clean_text_action)
        tools_menu.addSeparator()
        tools_menu.addAction(self.duplicate_line_action)
        tools_menu.addAction(self.delete_line_action)

        help_menu = self.menuBar().addMenu("&Help")
        update_action = QAction("Check for Updates…", self)
        update_action.triggered.connect(lambda: self.check_for_updates(manual=True))
        help_menu.addAction(update_action)
        help_menu.addSeparator()
        guide_action = QAction("Complete User Guide", self)
        guide_action.setShortcut("F1")
        guide_action.triggered.connect(self.show_user_guide)
        help_menu.addAction(guide_action)
        about_action = QAction(f"About {APP_NAME}", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def _build_ui(self):
        splitter = QSplitter(Qt.Horizontal)
        self.setCentralWidget(splitter)

        # Left: pages
        left = QWidget()
        left.setMinimumWidth(210)
        left_layout = QVBoxLayout(left)
        self.file_label = QLabel("No PDF loaded")
        self.file_label.setWordWrap(True)
        left_layout.addWidget(self.file_label)

        range_box = QGroupBox("Page Range")
        range_form = QFormLayout(range_box)
        self.start_page = QSpinBox()
        self.start_page.setRange(1, 1)
        self.end_page = QSpinBox()
        self.end_page.setRange(1, 1)
        range_form.addRow("Start:", self.start_page)
        range_form.addRow("End:", self.end_page)
        left_layout.addWidget(range_box)

        self.page_list = QListWidget()
        self.page_list.currentRowChanged.connect(self.show_original_page)
        left_layout.addWidget(self.page_list, 1)

        self.extract_btn = QPushButton("Extract and Reconstruct")
        self.extract_btn.clicked.connect(self.start_extract)
        left_layout.addWidget(self.extract_btn)

        self.progress = QProgressBar()
        left_layout.addWidget(self.progress)
        splitter.addWidget(left)

        # Center tabs
        self.tabs = QTabWidget()

        original_tab = QWidget()
        original_layout = QVBoxLayout(original_tab)
        self.original_scroll = QScrollArea()
        self.original_scroll.setWidgetResizable(True)
        self.original_image = QLabel("Open a PDF to preview pages")
        self.original_image.setAlignment(Qt.AlignCenter)
        self.original_scroll.setWidget(self.original_image)
        original_layout.addWidget(self.original_scroll)
        self.tabs.addTab(original_tab, "Original PDF")

        compare_tab = QWidget()
        compare_layout = QVBoxLayout(compare_tab)
        compare_layout.setContentsMargins(8, 8, 8, 8)
        compare_layout.setSpacing(8)

        editor_toolbar = QToolBar("Editor")
        editor_toolbar.setObjectName("editorToolbar")
        editor_toolbar.setMovable(False)
        editor_toolbar.setToolButtonStyle(Qt.ToolButtonTextOnly)

        self.undo_action = QAction("Undo", self)
        self.undo_action.setShortcut("Ctrl+Z")
        self.redo_action = QAction("Redo", self)
        self.redo_action.setShortcut("Ctrl+Y")
        self.cut_action = QAction("Cut", self)
        self.cut_action.setShortcut("Ctrl+X")
        self.copy_action = QAction("Copy", self)
        self.copy_action.setShortcut("Ctrl+C")
        self.paste_action = QAction("Paste", self)
        self.paste_action.setShortcut("Ctrl+V")
        self.paste_unicode_action = QAction("Paste + Unicode", self)
        self.paste_unicode_action.setShortcut("Ctrl+Shift+V")
        self.paste_unicode_action.setToolTip("Paste clipboard text after converting legacy Urdu glyphs to standard Unicode")
        self.import_word_action = QAction("Import Word Document…", self)
        self.import_word_action.setShortcut("Ctrl+Alt+W")
        self.import_text_action = QAction("Import Text File…", self)
        self.import_text_action.setShortcut("Ctrl+Alt+O")
        self.save_text_action = QAction("Save Editor Text…", self)
        self.save_text_action.setShortcut("Ctrl+Alt+S")
        self.select_all_action = QAction("Select All", self)
        self.select_all_action.setShortcut("Ctrl+A")
        self.find_action = QAction("Find / Replace", self)
        self.find_action.setShortcut("Ctrl+F")
        self.goto_action = QAction("Go to Line", self)
        self.goto_action.setShortcut("Ctrl+G")
        self.normalize_action = QAction("Normalize Unicode", self)
        self.normalize_action.setShortcut("Ctrl+Shift+U")
        self.recover_order_action = QAction("Recover Legacy Visual Order", self)
        self.recover_order_action.setToolTip("Convert presentation forms and correct reversed visual-order Urdu lines")
        self.clean_text_action = QAction("Clean Whitespace", self)
        self.clean_text_action.setShortcut("Ctrl+Shift+Space")
        self.duplicate_line_action = QAction("Duplicate Line", self)
        self.duplicate_line_action.setShortcut("Ctrl+D")
        self.delete_line_action = QAction("Delete Line", self)
        self.delete_line_action.setShortcut("Ctrl+Shift+K")
        self.stats_action = QAction("Text Statistics", self)
        self.stats_action.setShortcut("Ctrl+Shift+I")
        self.zoom_in_action = QAction("Zoom In", self)
        self.zoom_in_action.setShortcut("Ctrl++")
        self.zoom_out_action = QAction("Zoom Out", self)
        self.zoom_out_action.setShortcut("Ctrl+-")
        self.zoom_reset_action = QAction("Reset Zoom", self)
        self.zoom_reset_action.setShortcut("Ctrl+0")
        self.rtl_action = QAction("Right-to-Left", self)
        self.rtl_action.setShortcut("Ctrl+Shift+R")
        self.ltr_action = QAction("Left-to-Right", self)
        self.ltr_action.setShortcut("Ctrl+Shift+L")

        for action in [
            self.undo_action, self.redo_action, self.cut_action, self.copy_action,
            self.paste_action, self.paste_unicode_action, self.find_action, self.goto_action,
            self.normalize_action, self.clean_text_action, self.stats_action, self.zoom_out_action,
            self.zoom_reset_action, self.zoom_in_action,
        ]:
            editor_toolbar.addAction(action)
        compare_layout.addWidget(editor_toolbar)

        self.search_panel = QWidget()
        self.search_panel.setObjectName("searchPanel")
        search_layout = QHBoxLayout(self.search_panel)
        search_layout.setContentsMargins(8, 5, 8, 5)
        self.find_input = QLineEdit()
        self.find_input.setPlaceholderText("Find Unicode text…")
        self.find_input.setClearButtonEnabled(True)
        self.find_input.setMinimumWidth(160)
        self.replace_input = QLineEdit()
        self.replace_input.setPlaceholderText("Replace with…")
        self.replace_input.setClearButtonEnabled(True)
        self.match_case = QCheckBox("Match case")
        self.whole_words = QCheckBox("Whole words")
        find_previous_btn = QPushButton("Previous")
        find_next_btn = QPushButton("Next")
        replace_btn = QPushButton("Replace")
        replace_all_btn = QPushButton("Replace All")
        close_search_btn = QPushButton("Close")
        for button in [find_previous_btn, find_next_btn, replace_btn, replace_all_btn, close_search_btn]:
            button.setMinimumHeight(0)
        search_layout.addWidget(self.find_input, 2)
        search_layout.addWidget(self.replace_input, 2)
        search_layout.addWidget(self.match_case)
        search_layout.addWidget(self.whole_words)
        search_layout.addWidget(find_previous_btn)
        search_layout.addWidget(find_next_btn)
        search_layout.addWidget(replace_btn)
        search_layout.addWidget(replace_all_btn)
        search_layout.addWidget(close_search_btn)
        self.search_panel.hide()
        compare_layout.addWidget(self.search_panel)

        compare_splitter = QSplitter(Qt.Horizontal)
        self.source_text = QPlainTextEdit()
        self.source_text.setReadOnly(True)
        self.source_text.setPlaceholderText("Original extracted text")
        self.editor = QPlainTextEdit()
        self.editor.setLayoutDirection(Qt.RightToLeft)
        self.editor_base_font = QFont(self.editor.font())
        self.editor.setPlaceholderText("Reconstructed Urdu text will appear here")
        self.editor.textChanged.connect(self.schedule_preview)
        self.editor.setTabStopDistance(40)
        self.editor.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        compare_splitter.addWidget(self.source_text)
        compare_splitter.addWidget(self.editor)
        compare_splitter.setSizes([420, 580])
        compare_layout.addWidget(compare_splitter, 1)

        editor_status = QWidget()
        editor_status_layout = QHBoxLayout(editor_status)
        editor_status_layout.setContentsMargins(4, 0, 4, 0)
        self.editor_position_label = QLabel("Line 1, Column 1")
        self.editor_count_label = QLabel("0 words · 0 characters")
        editor_status_layout.addWidget(self.editor_position_label)
        editor_status_layout.addStretch(1)
        editor_status_layout.addWidget(self.editor_count_label)
        compare_layout.addWidget(editor_status)
        self.tabs.addTab(compare_tab, "Text Editor")

        self.undo_action.triggered.connect(self.editor.undo)
        self.redo_action.triggered.connect(self.editor.redo)
        self.cut_action.triggered.connect(self.editor.cut)
        self.copy_action.triggered.connect(self.editor.copy)
        self.paste_action.triggered.connect(self.editor.paste)
        self.paste_unicode_action.triggered.connect(self.paste_and_convert_unicode)
        self.import_word_action.triggered.connect(self.import_word_document)
        self.import_text_action.triggered.connect(self.import_text_file)
        self.save_text_action.triggered.connect(self.save_editor_text)
        self.select_all_action.triggered.connect(self.editor.selectAll)
        self.find_action.triggered.connect(self.show_find_replace)
        self.goto_action.triggered.connect(self.goto_line)
        self.normalize_action.triggered.connect(self.normalize_editor_unicode)
        self.recover_order_action.triggered.connect(self.recover_legacy_visual_order)
        self.clean_text_action.triggered.connect(self.clean_editor_text)
        self.duplicate_line_action.triggered.connect(self.duplicate_current_line)
        self.delete_line_action.triggered.connect(self.delete_current_line)
        self.stats_action.triggered.connect(self.show_text_statistics)
        self.zoom_in_action.triggered.connect(lambda: self.editor.zoomIn(1))
        self.zoom_out_action.triggered.connect(lambda: self.editor.zoomOut(1))
        self.zoom_reset_action.triggered.connect(self.reset_editor_zoom)
        self.rtl_action.triggered.connect(lambda: self.editor.setLayoutDirection(Qt.RightToLeft))
        self.ltr_action.triggered.connect(lambda: self.editor.setLayoutDirection(Qt.LeftToRight))
        find_previous_btn.clicked.connect(lambda: self.find_text(backward=True))
        find_next_btn.clicked.connect(self.find_text)
        replace_btn.clicked.connect(self.replace_one)
        replace_all_btn.clicked.connect(self.replace_all)
        close_search_btn.clicked.connect(self.search_panel.hide)
        self.find_input.returnPressed.connect(self.find_text)
        self.editor.cursorPositionChanged.connect(self.update_editor_status)
        self.editor.textChanged.connect(self.update_editor_status)

        preview_tab = QWidget()
        preview_layout = QVBoxLayout(preview_tab)
        self.preview = QTextBrowser()
        self.preview.setLayoutDirection(Qt.RightToLeft)
        preview_layout.addWidget(self.preview)
        self.tabs.addTab(preview_tab, "Book Preview")

        splitter.addWidget(self.tabs)

        # Right: settings
        right = QWidget()
        right.setMinimumWidth(360)
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(12, 12, 12, 12)
        right_layout.setSpacing(12)

        layout_box = QGroupBox("Book Layout")
        form = QFormLayout(layout_box)
        form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(9)

        self.font_combo = QComboBox()
        self.font_combo.setEditable(True)
        self.font_combo.addItems([
            "Noto Nastaliq Urdu", "Jameel Noori Nastaleeq",
            "Mehr Nastaliq Web", "Noto Naskh Arabic", "Arial"
        ])
        self.font_combo.setCurrentText(self.settings.font_family)
        self.font_combo.setToolTip("Choose an installed Unicode Urdu font. Nastaleeq fonts give the most traditional book appearance.")

        self.font_size = QDoubleSpinBox()
        self.font_size.setRange(8, 40)
        self.font_size.setValue(self.settings.font_size)
        self.font_size.setSuffix(" pt")
        self.font_size.setToolTip("Printed text size in points.")

        self.line_height = QDoubleSpinBox()
        self.line_height.setRange(1.0, 3.5)
        self.line_height.setSingleStep(0.1)
        self.line_height.setValue(self.settings.line_height)
        self.line_height.setToolTip("Space between text lines. Nastaleeq usually works best between 1.8 and 2.3.")

        self.para_space = QDoubleSpinBox()
        self.para_space.setRange(0, 30)
        self.para_space.setValue(self.settings.paragraph_spacing)
        self.para_space.setSuffix(" pt")
        self.para_space.setToolTip("Extra vertical space after each paragraph.")

        self.page_size = QComboBox()
        self.page_size.addItems(["A5", "B5", "A4"])
        self.page_size.setCurrentText(self.settings.page_size)
        self.page_size.setToolTip("Select the physical paper size used by preview and PDF export.")

        self.header_text = QLineEdit()
        self.header_text.setPlaceholderText("Optional book title")
        self.header_text.setToolTip("Optional title printed at the top of each exported page.")

        self.remove_jari = QCheckBox("Remove repeated 'جاری ہے'")
        self.remove_jari.setChecked(True)
        self.remove_jari.setToolTip("Remove repeated 'جاری ہے' markers commonly found at the bottom of source pages.")
        self.join_lines = QCheckBox("Join wrapped lines")
        self.join_lines.setChecked(True)
        self.join_lines.setToolTip("Treat source PDF line wrapping as continuous paragraph text.")
        self.preserve_blanks = QCheckBox("Preserve blank lines")
        self.preserve_blanks.setChecked(True)
        self.preserve_blanks.setToolTip("Keep meaningful blank lines when continuous paragraph mode is disabled.")
        self.page_numbers = QCheckBox("Page numbers")
        self.page_numbers.setChecked(True)
        self.page_numbers.setToolTip("Print an automatic page number in the exported PDF footer.")

        self.first_indent = QDoubleSpinBox()
        self.first_indent.setRange(0, 25)
        self.first_indent.setSingleStep(0.5)
        self.first_indent.setValue(0)
        self.first_indent.setSuffix(" mm")
        self.first_indent.setToolTip("Indent the first line of each normal paragraph.")

        self.word_spacing = QDoubleSpinBox()
        self.word_spacing.setRange(-2, 12)
        self.word_spacing.setSingleStep(0.25)
        self.word_spacing.setValue(0)
        self.word_spacing.setSuffix(" px")
        self.word_spacing.setToolTip("Fine-tune space between words. Use small values to avoid damaging Urdu readability.")

        self.letter_spacing = QDoubleSpinBox()
        self.letter_spacing.setRange(-1, 5)
        self.letter_spacing.setSingleStep(0.1)
        self.letter_spacing.setValue(0)
        self.letter_spacing.setSuffix(" px")
        self.letter_spacing.setToolTip("Fine-tune character spacing. Zero is recommended for joining Urdu scripts.")

        self.text_align = QComboBox()
        self.text_align.addItems(["Justify", "Right", "Center"])
        self.text_align.setToolTip("Choose paragraph alignment. Justify is recommended for books.")

        self.normalize_forms = QCheckBox("Convert legacy Urdu glyphs to Unicode")
        self.normalize_forms.setChecked(True)
        self.normalize_forms.setToolTip("Convert Arabic Presentation Form glyphs into standard searchable Unicode text.")
        self.recover_visual = QCheckBox("Recover reversed visual-order lines")
        self.recover_visual.setChecked(True)
        self.recover_visual.setToolTip("Reverse visually stored legacy PDF word order into normal logical Urdu reading order.")
        self.continuous_para = QCheckBox("Join all wrapped lines into paragraphs")
        self.continuous_para.setChecked(True)
        self.continuous_para.setToolTip("Ignore artificial blank rows inserted by legacy PDFs and rebuild flowing paragraphs.")

        form.addRow("Urdu font:", self.font_combo)
        form.addRow("Font size:", self.font_size)
        form.addRow("Line height:", self.line_height)
        form.addRow("Paragraph spacing:", self.para_space)
        form.addRow("Page size:", self.page_size)
        form.addRow("Running header:", self.header_text)
        form.addRow("First-line indent:", self.first_indent)
        form.addRow("Word spacing:", self.word_spacing)
        form.addRow("Letter spacing:", self.letter_spacing)
        form.addRow("Text alignment:", self.text_align)
        form.addRow(self.remove_jari)
        form.addRow(self.join_lines)
        form.addRow(self.preserve_blanks)
        form.addRow(self.normalize_forms)
        form.addRow(self.recover_visual)
        form.addRow(self.continuous_para)
        form.addRow(self.page_numbers)

        for widget in [self.font_combo, self.font_size, self.line_height, self.para_space,
                       self.page_size, self.header_text, self.page_numbers,
                       self.normalize_forms, self.recover_visual,
                       self.continuous_para, self.first_indent,
                       self.word_spacing, self.letter_spacing, self.text_align]:
            if hasattr(widget, "currentTextChanged"):
                widget.currentTextChanged.connect(self.schedule_preview)
            elif hasattr(widget, "valueChanged"):
                widget.valueChanged.connect(self.schedule_preview)
            elif hasattr(widget, "textChanged"):
                widget.textChanged.connect(self.schedule_preview)
            elif hasattr(widget, "toggled"):
                widget.toggled.connect(self.schedule_preview)

        right_layout.addWidget(layout_box)

        margins_box = QGroupBox("Margins (mm)")
        margins_form = QFormLayout(margins_box)
        margins_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        margins_form.setVerticalSpacing(9)
        self.margin_top = self._make_mm_spin(self.settings.margin_top_mm)
        self.margin_bottom = self._make_mm_spin(self.settings.margin_bottom_mm)
        self.margin_inner = self._make_mm_spin(self.settings.margin_inner_mm)
        self.margin_outer = self._make_mm_spin(self.settings.margin_outer_mm)
        margins_form.addRow("Top:", self.margin_top)
        margins_form.addRow("Bottom:", self.margin_bottom)
        margins_form.addRow("Inner:", self.margin_inner)
        margins_form.addRow("Outer:", self.margin_outer)
        right_layout.addWidget(margins_box)

        paragraph_box = QGroupBox("Paragraph Tools")
        paragraph_form = QFormLayout(paragraph_box)
        paragraph_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        paragraph_form.setVerticalSpacing(9)

        self.auto_para_lines = QSpinBox()
        self.auto_para_lines.setRange(2, 20)
        self.auto_para_lines.setValue(5)
        self.auto_para_lines.setToolTip(
            "After this many source lines, the program looks for a sentence ending "
            "(۔ . ؟ !) and starts a new paragraph."
        )

        self.join_selected_btn = QPushButton("Join Selected Lines as Paragraph")
        self.join_selected_btn.setToolTip("Join selected editor lines while preserving blank-line paragraph boundaries. With no selection, processes all text.")
        self.join_selected_btn.clicked.connect(self.join_selected_lines_as_paragraph)

        self.split_selected_btn = QPushButton("Split Selected Text at ۔ or .")
        self.split_selected_btn.setToolTip("Insert paragraph breaks after Urdu or Latin sentence-ending punctuation.")
        self.split_selected_btn.clicked.connect(self.split_selected_text_at_sentence_end)

        self.auto_para_btn = QPushButton("Auto Make Paragraphs")
        self.auto_para_btn.setToolTip("Build paragraphs automatically, waiting for sentence punctuation after the selected minimum number of lines.")
        self.auto_para_btn.clicked.connect(self.auto_make_paragraphs)

        self.remove_extra_breaks_btn = QPushButton("Remove Extra Line Breaks")
        self.remove_extra_breaks_btn.setToolTip("Remove hard wraps inside paragraphs while keeping blank lines as paragraph separators.")
        self.remove_extra_breaks_btn.clicked.connect(self.remove_extra_line_breaks)

        paragraph_form.addRow("Minimum source lines:", self.auto_para_lines)
        paragraph_form.addRow(self.join_selected_btn)
        paragraph_form.addRow(self.split_selected_btn)
        paragraph_form.addRow(self.auto_para_btn)
        paragraph_form.addRow(self.remove_extra_breaks_btn)
        right_layout.addWidget(paragraph_box)

        self.preview_btn = QPushButton("Refresh Preview")
        self.preview_btn.setToolTip("Rebuild the Book Preview tab using the current text and layout settings.")
        self.preview_btn.clicked.connect(self.update_preview)
        right_layout.addWidget(self.preview_btn)

        self.export_pdf_btn = QPushButton("Export Print PDF")
        self.export_pdf_btn.setToolTip("Create a print-ready PDF using the current page, margin, font, and paragraph settings.")
        self.export_pdf_btn.clicked.connect(self.export_pdf)
        right_layout.addWidget(self.export_pdf_btn)

        self.export_html_btn = QPushButton("Export HTML")
        self.export_html_btn.setToolTip("Create a portable UTF-8 HTML document with RTL styling.")
        self.export_html_btn.clicked.connect(self.export_html)
        right_layout.addWidget(self.export_html_btn)

        right_layout.addStretch(1)
        settings_scroll = QScrollArea()
        settings_scroll.setObjectName("settingsScroll")
        settings_scroll.setWidgetResizable(True)
        settings_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        settings_scroll.setMinimumWidth(380)
        settings_scroll.setWidget(right)
        splitter.addWidget(settings_scroll)

        splitter.setChildrenCollapsible(False)
        splitter.setSizes([240, 850, 400])

        self.preview_timer = QTimer(self)
        self.preview_timer.setSingleShot(True)
        self.preview_timer.timeout.connect(self.update_preview)

    def _make_mm_spin(self, value: float) -> QDoubleSpinBox:
        s = QDoubleSpinBox()
        s.setRange(5, 50)
        s.setValue(value)
        s.setSuffix(" mm")
        s.setToolTip("Printed distance from the page edge, in millimetres.")
        s.valueChanged.connect(self.schedule_preview)
        return s

    def _apply_dark_theme(self):
        self.setStyleSheet("""
        QMainWindow, QWidget { background: #15171a; color: #e8e8e8; font-size: 10pt; }
        QScrollArea#settingsScroll { border: none; background: #15171a; }
        QWidget#searchPanel { background: #20252b; border: 1px solid #3d4854; border-radius: 7px; }
        QToolBar#editorToolbar { background: #20252b; border: 1px solid #343b44; border-radius: 7px; spacing: 3px; padding: 3px; }
        QToolBar#editorToolbar QToolButton { background: transparent; color: #e9f1f8; padding: 6px 8px; border-radius: 4px; }
        QToolBar#editorToolbar QToolButton:hover { background: #2f78bd; }
        QMenuBar, QMenu, QToolBar { background: #202328; color: #f2f2f2; }
        QGroupBox { border: 1px solid #3b3f45; border-radius: 8px; margin-top: 12px; padding: 14px 10px 10px 10px; font-weight: 600; }
        QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
        QPushButton { background: #2f78bd; border: none; min-height: 20px; padding: 9px 12px; border-radius: 6px; font-weight: 600; }
        QPushButton:hover { background: #347bc4; }
        QLineEdit, QPlainTextEdit, QTextBrowser, QComboBox, QSpinBox, QDoubleSpinBox, QListWidget {
            background: #0f1113; color: #f4f4f4; border: 1px solid #3c4148; border-radius: 5px;
            selection-background-color: #2b6cb0;
        }
        QTabWidget::pane { border: 1px solid #3c4148; }
        QTabBar::tab { background: #24272c; padding: 8px 14px; }
        QTabBar::tab:selected { background: #2b6cb0; }
        QProgressBar { border: 1px solid #444; border-radius: 5px; text-align: center; }
        QProgressBar::chunk { background: #2b6cb0; }
        QToolTip { background: #f5f7fa; color: #15171a; border: 1px solid #9aa4b2; padding: 6px; }
        """)

    def current_settings(self) -> LayoutSettings:
        return LayoutSettings(
            font_family=self.font_combo.currentText().strip() or "Noto Nastaliq Urdu",
            font_size=self.font_size.value(),
            line_height=self.line_height.value(),
            paragraph_spacing=self.para_space.value(),
            page_size=self.page_size.currentText(),
            margin_top_mm=self.margin_top.value(),
            margin_bottom_mm=self.margin_bottom.value(),
            margin_inner_mm=self.margin_inner.value(),
            margin_outer_mm=self.margin_outer.value(),
            remove_jari_hai=self.remove_jari.isChecked(),
            join_wrapped_lines=self.join_lines.isChecked(),
            preserve_blank_lines=self.preserve_blanks.isChecked(),
            page_numbers=self.page_numbers.isChecked(),
            running_header=self.header_text.text().strip(),
            recover_visual_order=self.recover_visual.isChecked(),
            normalize_presentation_forms=self.normalize_forms.isChecked(),
            continuous_paragraph_mode=self.continuous_para.isChecked(),
            first_line_indent_mm=self.first_indent.value(),
            word_spacing_px=self.word_spacing.value(),
            letter_spacing_px=self.letter_spacing.value(),
            text_align=self.text_align.currentText(),
            auto_paragraph_lines=self.auto_para_lines.value(),
        )

    def new_project(self):
        self.pdf_path = None
        self.project_path = None
        self.page_count = 0
        self.extracted_text = ""
        self.original_page_texts = []
        self.page_list.clear()
        self.editor.clear()
        self.source_text.clear()
        self.preview.clear()
        self.original_image.setText("Open a PDF to preview pages")
        self.file_label.setText("No PDF loaded")
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.statusBar().showMessage("New project")

    def apply_settings(self, settings: LayoutSettings):
        """Apply persisted settings to the controls without depending on field order."""
        self.font_combo.setCurrentText(settings.font_family)
        self.font_size.setValue(settings.font_size)
        self.line_height.setValue(settings.line_height)
        self.para_space.setValue(settings.paragraph_spacing)
        self.page_size.setCurrentText(settings.page_size)
        self.margin_top.setValue(settings.margin_top_mm)
        self.margin_bottom.setValue(settings.margin_bottom_mm)
        self.margin_inner.setValue(settings.margin_inner_mm)
        self.margin_outer.setValue(settings.margin_outer_mm)
        self.remove_jari.setChecked(settings.remove_jari_hai)
        self.join_lines.setChecked(settings.join_wrapped_lines)
        self.preserve_blanks.setChecked(settings.preserve_blank_lines)
        self.page_numbers.setChecked(settings.page_numbers)
        self.header_text.setText(settings.running_header)
        self.recover_visual.setChecked(settings.recover_visual_order)
        self.normalize_forms.setChecked(settings.normalize_presentation_forms)
        self.continuous_para.setChecked(settings.continuous_paragraph_mode)
        self.first_indent.setValue(settings.first_line_indent_mm)
        self.word_spacing.setValue(settings.word_spacing_px)
        self.letter_spacing.setValue(settings.letter_spacing_px)
        self.text_align.setCurrentText(settings.text_align)
        self.auto_para_lines.setValue(settings.auto_paragraph_lines)
        self.settings = settings
        self.update_preview()

    def open_project(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Urdu Unicoder Project", "", "Urdu Unicoder Project (*.ubp)"
        )
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            defaults = asdict(LayoutSettings())
            saved_settings = data.get("settings", {})
            if not isinstance(saved_settings, dict):
                raise ValueError("The project settings are not valid.")
            defaults.update({key: value for key, value in saved_settings.items() if key in defaults})
            self.apply_settings(LayoutSettings(**defaults))
            self.project_path = path
            self.pdf_path = data.get("pdf_path")
            self.page_count = int(data.get("page_count", 0))
            self.editor.setPlainText(data.get("reconstructed_text", ""))
            self.start_page.setRange(1, max(1, self.page_count))
            self.end_page.setRange(1, max(1, self.page_count))
            self.start_page.setValue(max(1, int(data.get("start_page", 1))))
            self.end_page.setValue(max(1, int(data.get("end_page", self.page_count or 1))))
            self.page_list.clear()
            if self.pdf_path and Path(self.pdf_path).exists():
                self.file_label.setText(f"{Path(self.pdf_path).name}\n{self.page_count:,} pages")
                for i in range(self.page_count):
                    self.page_list.addItem(QListWidgetItem(f"Page {i + 1}"))
            else:
                self.pdf_path = None
                self.file_label.setText("Source PDF is unavailable; reconstructed text is still editable")
            self.setWindowTitle(f"{APP_NAME} — {Path(path).name}")
            self.statusBar().showMessage(f"Project opened: {path}")
        except Exception as exc:
            QMessageBox.critical(self, "Cannot open project", f"The project could not be opened.\n\n{exc}")

    def open_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Editable Urdu PDF", "", "PDF Files (*.pdf)")
        if not path:
            return
        try:
            doc = fitz.open(path)
            self.page_count = doc.page_count
            doc.close()
            self.pdf_path = path
            self.file_label.setText(f"{Path(path).name}\n{self.page_count:,} pages")
            self.start_page.setRange(1, self.page_count)
            self.end_page.setRange(1, self.page_count)
            self.start_page.setValue(1)
            self.end_page.setValue(self.page_count)

            self.page_list.clear()
            # Avoid creating thousands of expensive thumbnails; lightweight page entries only.
            for i in range(self.page_count):
                self.page_list.addItem(QListWidgetItem(f"Page {i + 1}"))
            if self.page_count:
                self.page_list.setCurrentRow(0)

            self.statusBar().showMessage(f"Loaded {self.page_count:,} pages")
        except Exception as exc:
            QMessageBox.critical(self, "Cannot open PDF", str(exc))

    def show_original_page(self, row: int):
        if not self.pdf_path or row < 0:
            return
        self.original_image.setText(f"Rendering page {row + 1}…")
        self.page_worker = PdfPageWorker(self.pdf_path, row)
        self.page_worker.done.connect(self.on_page_rendered)
        self.page_worker.failed.connect(lambda e: self.original_image.setText(e))
        self.page_worker.start()

    def on_page_rendered(self, page_no: int, image: QImage):
        pix = QPixmap.fromImage(image)
        self.original_image.setPixmap(pix)
        self.original_image.adjustSize()

    def start_extract(self):
        if not self.pdf_path:
            QMessageBox.information(self, "Open PDF", "Please open an editable PDF first.")
            return
        start = self.start_page.value()
        end = self.end_page.value()
        if end < start:
            QMessageBox.warning(self, "Invalid range", "End page must be greater than or equal to start page.")
            return

        self.extract_btn.setEnabled(False)
        self.progress.setRange(0, end - start + 1)
        self.progress.setValue(0)
        self.settings = self.current_settings()
        self.worker = ExtractWorker(self.pdf_path, start, end, self.settings)
        self.worker.progress.connect(self.on_extract_progress)
        self.worker.finished_ok.connect(self.on_extract_done)
        self.worker.failed.connect(self.on_extract_failed)
        self.worker.start()

    def on_extract_progress(self, current: int, total: int, message: str):
        self.progress.setMaximum(total)
        self.progress.setValue(current)
        self.statusBar().showMessage(message)

    def on_extract_done(self, rebuilt: str, page_texts: list):
        self.extract_btn.setEnabled(True)
        self.extracted_text = rebuilt
        self.original_page_texts = page_texts
        self.source_text.setPlainText("\n\n--- PAGE BREAK ---\n\n".join(page_texts))
        self.editor.setPlainText(rebuilt)
        self.update_preview()
        self.tabs.setCurrentIndex(1)
        self.statusBar().showMessage("Text reconstruction completed")

    def on_extract_failed(self, error: str):
        self.extract_btn.setEnabled(True)
        QMessageBox.critical(self, "Extraction failed", error)
        self.statusBar().showMessage("Extraction failed")

    def _selected_or_all_text(self) -> tuple[QTextCursor, str, bool]:
        cursor = self.editor.textCursor()
        has_selection = cursor.hasSelection()
        if has_selection:
            text = cursor.selectedText().replace("\u2029", "\n").replace("\u2028", "\n")
        else:
            text = self.editor.toPlainText()
        return cursor, text, has_selection

    def _load_text_into_editor(self, text: str, source_name: str):
        if not text.strip():
            QMessageBox.information(self, "No text found", f"No readable text was found in {source_name}.")
            return
        if self.editor.toPlainText().strip():
            choice = QMessageBox.question(
                self,
                "Import text",
                "Replace the current editor text?\n\n"
                "Choose No to append the imported text, or Cancel to stop.",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
            )
            if choice == QMessageBox.Cancel:
                return
            if choice == QMessageBox.No:
                cursor = self.editor.textCursor()
                cursor.movePosition(QTextCursor.End)
                cursor.insertText("\n\n" + text)
            else:
                self.editor.setPlainText(text)
        else:
            self.editor.setPlainText(text)
        self.tabs.setCurrentIndex(1)
        self.editor.setFocus()
        self.update_preview()
        self.statusBar().showMessage(f"Imported {source_name}", 5000)

    def import_word_document(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Import Microsoft Word Document", "", "Word Documents (*.docx)"
        )
        if not path:
            return
        try:
            text = extract_docx_text(path)
            self._load_text_into_editor(text, Path(path).name)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Word import failed",
                f"The Word document could not be imported.\n\n{exc}\n\n"
                "Legacy .doc files must first be saved as .docx in Microsoft Word or LibreOffice.",
            )

    def import_text_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Import Unicode Text", "", "Text Files (*.txt *.md);;All Files (*.*)"
        )
        if not path:
            return
        try:
            raw = Path(path).read_bytes()
            text = None
            for encoding in ("utf-8-sig", "utf-16", "cp1256"):
                try:
                    text = raw.decode(encoding)
                    break
                except UnicodeDecodeError:
                    continue
            if text is None:
                raise UnicodeError("The text encoding was not recognized.")
            self._load_text_into_editor(text, Path(path).name)
        except Exception as exc:
            QMessageBox.critical(self, "Text import failed", str(exc))

    def save_editor_text(self):
        if not self.editor.toPlainText():
            QMessageBox.information(self, "Nothing to save", "The editor is empty.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Unicode Text", "Urdu_Unicoder_Text.txt", "Text Files (*.txt)"
        )
        if not path:
            return
        if not path.lower().endswith(".txt"):
            path += ".txt"
        Path(path).write_text(self.editor.toPlainText(), encoding="utf-8-sig")
        self.statusBar().showMessage(f"Unicode text saved: {path}", 5000)

    def paste_and_convert_unicode(self):
        text = QApplication.clipboard().text()
        if not text:
            self.statusBar().showMessage("The clipboard does not contain text", 3500)
            return
        converted = unicodedata.normalize("NFKC", text)
        converted = converted.replace("\u200e", "").replace("\u200f", "")
        converted = converted.replace("\r\n", "\n").replace("\r", "\n")
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        cursor.insertText(converted)
        cursor.endEditBlock()
        self.editor.setTextCursor(cursor)
        self.tabs.setCurrentIndex(1)
        self.statusBar().showMessage("Clipboard text pasted and converted to Unicode", 4000)

    def recover_legacy_visual_order(self):
        cursor, text, had_selection = self._selected_or_all_text()
        if not text:
            return
        recovered = "\n".join(
            ExtractWorker.recover_visual_order_line(line) if line.strip() else ""
            for line in text.splitlines()
        )
        self._replace_selection_or_all(cursor, recovered, had_selection)
        self.statusBar().showMessage("Legacy Urdu visual order recovered", 4000)

    def clean_editor_text(self):
        cursor, text, had_selection = self._selected_or_all_text()
        if not text:
            return
        lines = [re.sub(r"[ \t]+", " ", line).strip() for line in text.splitlines()]
        cleaned = re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()
        self._replace_selection_or_all(cursor, cleaned, had_selection)
        self.statusBar().showMessage("Whitespace and extra blank lines cleaned", 4000)

    def duplicate_current_line(self):
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        cursor.select(QTextCursor.BlockUnderCursor)
        line = cursor.selectedText()
        cursor.movePosition(QTextCursor.EndOfBlock)
        cursor.insertText("\n" + line)
        cursor.endEditBlock()
        self.editor.setTextCursor(cursor)

    def delete_current_line(self):
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        cursor.select(QTextCursor.BlockUnderCursor)
        cursor.removeSelectedText()
        if not cursor.atEnd():
            cursor.deleteChar()
        elif cursor.position() > 0:
            cursor.deletePreviousChar()
        cursor.endEditBlock()
        self.editor.setTextCursor(cursor)

    def show_find_replace(self):
        cursor = self.editor.textCursor()
        selected = cursor.selectedText().replace("\u2029", "\n")
        if selected and "\n" not in selected:
            self.find_input.setText(selected)
        self.search_panel.show()
        self.find_input.setFocus()
        self.find_input.selectAll()

    def find_text(self, backward: bool = False):
        query = self.find_input.text()
        if not query:
            self.show_find_replace()
            return
        flags = QTextDocument.FindFlags()
        if backward:
            flags |= QTextDocument.FindBackward
        if self.match_case.isChecked():
            flags |= QTextDocument.FindCaseSensitively
        if self.whole_words.isChecked():
            flags |= QTextDocument.FindWholeWords
        if not self.editor.find(query, flags):
            cursor = self.editor.textCursor()
            cursor.movePosition(QTextCursor.End if backward else QTextCursor.Start)
            self.editor.setTextCursor(cursor)
            if not self.editor.find(query, flags):
                self.statusBar().showMessage(f"Text not found: {query}", 3500)

    def replace_one(self):
        query = self.find_input.text()
        if not query:
            return
        cursor = self.editor.textCursor()
        selected = cursor.selectedText()
        matches = selected == query if self.match_case.isChecked() else selected.casefold() == query.casefold()
        if matches:
            cursor.insertText(self.replace_input.text())
            self.editor.setTextCursor(cursor)
        self.find_text()

    def replace_all(self):
        query = self.find_input.text()
        if not query:
            return
        text = self.editor.toPlainText()
        pattern = re.escape(query)
        if self.whole_words.isChecked():
            pattern = rf"(?<!\w){pattern}(?!\w)"
        flags = 0 if self.match_case.isChecked() else re.IGNORECASE
        replaced, count = re.subn(pattern, lambda _match: self.replace_input.text(), text, flags=flags)
        if count:
            cursor_position = self.editor.textCursor().position()
            self.editor.setPlainText(replaced)
            cursor = self.editor.textCursor()
            cursor.setPosition(min(cursor_position, len(replaced)))
            self.editor.setTextCursor(cursor)
        self.statusBar().showMessage(f"Replaced {count:,} occurrence(s)", 4000)

    def goto_line(self):
        maximum = max(1, self.editor.blockCount())
        current = self.editor.textCursor().blockNumber() + 1
        line, accepted = QInputDialog.getInt(
            self, "Go to Line", f"Line number (1–{maximum}):", current, 1, maximum
        )
        if not accepted:
            return
        block = self.editor.document().findBlockByNumber(line - 1)
        if block.isValid():
            cursor = QTextCursor(block)
            self.editor.setTextCursor(cursor)
            self.editor.centerCursor()
            self.editor.setFocus()

    def normalize_editor_unicode(self):
        cursor, text, had_selection = self._selected_or_all_text()
        if not text:
            return
        normalized = unicodedata.normalize("NFKC", text)
        self._replace_selection_or_all(cursor, normalized, had_selection)
        scope = "selection" if had_selection else "document"
        self.statusBar().showMessage(f"Unicode normalized in {scope}", 3500)

    def reset_editor_zoom(self):
        self.editor.setFont(QFont(self.editor_base_font))
        self.statusBar().showMessage("Editor zoom reset", 2500)

    def update_editor_status(self):
        cursor = self.editor.textCursor()
        text = self.editor.toPlainText()
        words = len(re.findall(r"\S+", text))
        self.editor_position_label.setText(
            f"Line {cursor.blockNumber() + 1:,}, Column {cursor.positionInBlock() + 1:,}"
        )
        self.editor_count_label.setText(f"{words:,} words · {len(text):,} characters")

    def show_text_statistics(self):
        text = self.editor.toPlainText()
        words = len(re.findall(r"\S+", text))
        non_space = len(re.sub(r"\s", "", text))
        paragraphs = len([part for part in re.split(r"\n\s*\n", text) if part.strip()])
        QMessageBox.information(
            self,
            "Text Statistics",
            f"Words: {words:,}\n"
            f"Characters: {len(text):,}\n"
            f"Characters without spaces: {non_space:,}\n"
            f"Lines: {self.editor.blockCount():,}\n"
            f"Paragraphs: {paragraphs:,}",
        )

    @staticmethod
    def _join_wrapped_lines(text: str) -> str:
        lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in text.splitlines()]
        chunks = []
        current = []
        for line in lines:
            if not line:
                if current:
                    chunks.append(" ".join(current))
                    current = []
                continue
            current.append(line)
        if current:
            chunks.append(" ".join(current))
        return "\n\n".join(chunks)

    def _replace_selection_or_all(self, cursor: QTextCursor, text: str, had_selection: bool):
        if had_selection:
            cursor.beginEditBlock()
            cursor.insertText(text)
            cursor.endEditBlock()
            self.editor.setTextCursor(cursor)
        else:
            self.editor.setPlainText(text)
        self.update_preview()

    def join_selected_lines_as_paragraph(self):
        cursor, text, had_selection = self._selected_or_all_text()
        if not text.strip():
            return
        joined = self._join_wrapped_lines(text)
        # A blank line after the selection marks the next paragraph.
        if had_selection:
            joined = joined.strip() + "\n\n"
        self._replace_selection_or_all(cursor, joined, had_selection)
        self.statusBar().showMessage("Selected lines joined into a paragraph")

    @staticmethod
    def _split_sentences_into_paragraphs(text: str) -> str:
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\s*\n\s*", " ", text).strip()
        # Keep punctuation with the sentence and create a real blank line after it.
        text = re.sub(r"([۔؟!]|(?<!\.)\.(?!\.))\s+", r"\1\n\n", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()

    def split_selected_text_at_sentence_end(self):
        cursor, text, had_selection = self._selected_or_all_text()
        if not text.strip():
            return
        split_text = self._split_sentences_into_paragraphs(text)
        self._replace_selection_or_all(cursor, split_text, had_selection)
        self.statusBar().showMessage("Paragraph breaks inserted after sentence endings")

    def auto_make_paragraphs(self):
        cursor, text, had_selection = self._selected_or_all_text()
        if not text.strip():
            return

        min_lines = self.auto_para_lines.value()
        source_lines = [
            re.sub(r"[ \t]+", " ", ln).strip()
            for ln in text.splitlines()
            if ln.strip()
        ]

        paragraphs = []
        current = []
        for line in source_lines:
            current.append(line)
            joined = " ".join(current)
            ends_sentence = bool(re.search(r"[۔؟!]|(?<!\.)\.(?!\.)\s*$", line))
            if len(current) >= min_lines and ends_sentence:
                paragraphs.append(joined)
                current = []

        if current:
            paragraphs.append(" ".join(current))

        result = "\n\n".join(paragraphs)
        self._replace_selection_or_all(cursor, result, had_selection)
        self.statusBar().showMessage(
            f"Automatic paragraphs created after at least {min_lines} source lines"
        )

    def remove_extra_line_breaks(self):
        cursor, text, had_selection = self._selected_or_all_text()
        if not text.strip():
            return
        result = self._join_wrapped_lines(text)
        self._replace_selection_or_all(cursor, result, had_selection)
        self.statusBar().showMessage("Extra line breaks removed")

    def schedule_preview(self, *args):
        self.preview_timer.start(350)

    @staticmethod
    def esc(text: str) -> str:
        return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

    def build_html(self, full_document: bool = True) -> str:
        s = self.current_settings()
        text = self.editor.toPlainText().strip()
        paragraphs = []
        # A single newline is usually a wrapped source line. Only a blank line
        # starts a new paragraph. This produces normal book-style flowing text.
        for block in re.split(r"\n\s*\n", text):
            block = re.sub(r"\s*\n\s*", " ", block).strip()
            block = re.sub(r"[ \t]{2,}", " ", block)
            if not block:
                continue
            css_class = "heading" if ExtractWorker.looks_like_heading(block) else "para"
            paragraphs.append(f'<p class="{css_class}">{self.esc(block)}</p>')

        size_map = {
            "A5": ("148mm", "210mm"),
            "B5": ("176mm", "250mm"),
            "A4": ("210mm", "297mm"),
        }
        width, height = size_map.get(s.page_size, size_map["A5"])
        align_map = {
            "Justify": ("justify", "right"),
            "Right": ("right", "right"),
            "Center": ("center", "center"),
        }
        css_align, css_last_align = align_map.get(s.text_align, ("justify", "right"))

        footer = """
        @bottom-center {
            content: counter(page);
            font-size: 9pt;
        }
        """ if s.page_numbers else ""

        header = ""
        if s.running_header:
            safe_header = self.esc(s.running_header).replace('"', '\\"')
            header = f"""
            @top-center {{
                content: "{safe_header}";
                font-size: 9pt;
            }}
            """

        html = f"""<!doctype html>
<html lang="ur" dir="rtl">
<head>
<meta charset="utf-8">
<title>{self.esc(s.running_header or "Urdu Book")}</title>
<style>
@page {{
    size: {width} {height};
    margin: {s.margin_top_mm}mm {s.margin_outer_mm}mm {s.margin_bottom_mm}mm {s.margin_inner_mm}mm;
    {footer}
    {header}
}}
html, body {{
    direction: rtl;
    background: white;
    color: black;
}}
body {{
    font-family: "{s.font_family}", "Noto Nastaliq Urdu", "Noto Naskh Arabic", serif;
    font-size: {s.font_size}pt;
    line-height: {s.line_height};
    text-align: {css_align};
    text-align-last: {css_last_align};
    word-spacing: {s.word_spacing_px}px;
    letter-spacing: {s.letter_spacing_px}px;
    white-space: normal;
    margin: 0;
}}
p.para {{
    margin: 0 0 {s.paragraph_spacing}pt 0;
    text-align: {css_align};
    text-align-last: {css_last_align};
    word-spacing: {s.word_spacing_px}px;
    letter-spacing: {s.letter_spacing_px}px;
    text-indent: {s.first_line_indent_mm}mm;
    orphans: 2;
    widows: 2;
}}
p.heading {{
    text-align: center;
    font-weight: bold;
    font-size: {s.font_size + 3}pt;
    margin: {s.paragraph_spacing + 8}pt 0 {s.paragraph_spacing + 6}pt 0;
    page-break-after: avoid;
}}
</style>
</head>
<body>
{''.join(paragraphs)}
</body>
</html>"""
        return html

    def update_preview(self):
        self.preview.setHtml(self.build_html())
        self.extracted_text = self.editor.toPlainText()

    def export_html(self):
        if not self.editor.toPlainText().strip():
            QMessageBox.information(self, "Nothing to export", "Extract or enter text first.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Export HTML", "Urdu_Unicoder_Book.html", "HTML Files (*.html)")
        if not path:
            return
        Path(path).write_text(self.build_html(), encoding="utf-8")
        self.statusBar().showMessage(f"HTML exported: {path}")
        QMessageBox.information(self, "Export complete", f"HTML saved successfully:\n{path}")

    def export_pdf(self):
        if not self.editor.toPlainText().strip():
            QMessageBox.information(self, "Nothing to export", "Extract or enter text first.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Export Print PDF", "Urdu_Unicoder_Book.pdf", "PDF Files (*.pdf)")
        if not path:
            return
        try:
            s = self.current_settings()
            printer = QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(path)

            page_size_map = {
                "A5": QPageSize(QPageSize.A5),
                "B5": QPageSize(QPageSize.B5),
                "A4": QPageSize(QPageSize.A4),
            }
            printer.setPageSize(page_size_map.get(s.page_size, QPageSize(QPageSize.A5)))
            # PySide6 6.7+ accepts one QMarginsF object, in left/top/right/bottom order.
            printer.setPageMargins(
                QMarginsF(s.margin_outer_mm, s.margin_top_mm, s.margin_inner_mm, s.margin_bottom_mm),
                QPageLayout.Millimeter,
            )

            doc = QTextDocument()
            doc.setDefaultFont(QFont(s.font_family, int(round(s.font_size))))
            doc.setHtml(self.build_html())
            doc.print_(printer)

            self.statusBar().showMessage(f"PDF exported: {path}")
            QMessageBox.information(
                self, "Export complete",
                f"PDF saved successfully:\n{path}\n\n"
                "For the best Nastaleeq result, install your preferred Unicode Urdu font in Windows."
            )
        except Exception as exc:
            QMessageBox.critical(self, "PDF export failed", str(exc))

    def check_for_updates(self, manual: bool = False):
        if self.update_worker is not None and self.update_worker.isRunning():
            if manual:
                self.statusBar().showMessage("An update check is already running", 3000)
            return
        self.update_check_manual = manual
        if manual:
            self.statusBar().showMessage("Checking GitHub for updates…")
        self.update_worker = UpdateCheckWorker(self)
        self.update_worker.update_info.connect(self.on_update_info)
        self.update_worker.failed.connect(self.on_update_check_failed)
        self.update_worker.start()

    def on_update_info(self, info: dict):
        latest = str(info.get("version", "0"))
        if version_tuple(latest) <= version_tuple(APP_VERSION):
            if self.update_check_manual:
                QMessageBox.information(
                    self, "No updates available",
                    f"{APP_NAME} {APP_VERSION} is the latest version."
                )
            self.statusBar().showMessage(f"{APP_NAME} is up to date", 3500)
            return

        notes = str(info.get("notes", "Performance improvements and new features."))
        reply = QMessageBox.question(
            self,
            "Urdu Unicoder update available",
            f"Version {latest} is available. You are using {APP_VERSION}.\n\n"
            f"What is new:\n{notes}\n\nDownload and install this update now?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            self.install_update(info)

    def on_update_check_failed(self, error: str):
        if self.update_check_manual:
            QMessageBox.warning(
                self, "Update check failed",
                f"Urdu Unicoder could not contact GitHub.\n\n{error}"
            )
        self.statusBar().showMessage("Update check unavailable", 3500)

    def install_update(self, info: dict):
        if getattr(sys, "frozen", False):
            release_url = str(info.get("release_url") or UPDATE_PAGE_URL)
            QDesktopServices.openUrl(QUrl(release_url))
            QMessageBox.information(
                self, "Download opened",
                "Download the latest Urdu Unicoder installer from the GitHub release page, "
                "close this version, and run the installer. Your project files are not removed."
            )
            return

        if self.editor.toPlainText().strip():
            save_reply = QMessageBox.question(
                self,
                "Save before updating?",
                "Save the current project before Urdu Unicoder closes for the update?",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
            )
            if save_reply == QMessageBox.Cancel:
                return
            if save_reply == QMessageBox.Yes and not self.save_project():
                return

        project_root = Path(__file__).resolve().parent.parent
        updater = project_root / "app" / "updater.py"
        if not updater.exists():
            QDesktopServices.openUrl(QUrl(str(info.get("release_url") or UPDATE_PAGE_URL)))
            return
        try:
            command = [sys.executable, str(updater), str(project_root), str(os.getpid())]
            kwargs = {"cwd": str(project_root)}
            if sys.platform == "win32":
                kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
            subprocess.Popen(command, **kwargs)
            self.statusBar().showMessage("Installing update and restarting…")
            QApplication.quit()
        except Exception as exc:
            QMessageBox.critical(self, "Update could not start", str(exc))

    def save_project(self) -> bool:
        if not self.project_path:
            path, _ = QFileDialog.getSaveFileName(self, "Save Project", "My_Urdu_Book.ubp", "Urdu Unicoder Project (*.ubp)")
            if not path:
                return False
            if not path.lower().endswith(PROJECT_EXT):
                path += PROJECT_EXT
            self.project_path = path

        data = {
            "app_version": APP_VERSION,
            "pdf_path": self.pdf_path,
            "page_count": self.page_count,
            "start_page": self.start_page.value(),
            "end_page": self.end_page.value(),
            "settings": asdict(self.current_settings()),
            "reconstructed_text": self.editor.toPlainText(),
        }
        Path(self.project_path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        self.setWindowTitle(f"{APP_NAME} — {Path(self.project_path).name}")
        self.statusBar().showMessage("Project saved")
        return True

    def show_user_guide(self):
        dialog = QDialog(self)
        dialog.setWindowTitle(f"{APP_NAME} — Complete User Guide")
        dialog.resize(820, 680)
        layout = QVBoxLayout(dialog)
        guide = QTextBrowser()
        guide.setOpenExternalLinks(True)
        guide.setHtml("""
        <h1>Urdu Unicoder user guide</h1>
        <p><b>Recommended workflow:</b> Open a text-based PDF, choose a small page range,
        reconstruct it, correct the result in Text Editor, check Book Preview, save the
        project, and export PDF or HTML.</p>
        <h2>Source and reconstruction</h2>
        <p><b>Open PDF</b> loads an editable/text-based PDF. Scanned image PDFs require OCR
        before use. <b>Start / End</b> limits processing to selected pages. Test 10–30 pages
        first for a large book. <b>Extract and Reconstruct</b> reads PDF text in the background.</p>
        <p><b>Convert legacy Urdu glyphs to Unicode</b> changes Arabic presentation glyphs to
        standard, searchable Unicode. <b>Recover reversed visual-order lines</b> corrects older
        PDFs that store words in display order. Disable it if an already-correct PDF becomes
        reversed. <b>Join wrapped lines</b> merges visual PDF rows. <b>Join all wrapped lines into
        paragraphs</b> also ignores artificial blank rows. <b>Preserve blank lines</b> matters
        when continuous paragraph mode is off. <b>Remove repeated 'جاری ہے'</b> removes page-end
        continuation markers.</p>
        <h2>Book layout</h2>
        <p><b>Urdu font</b> selects an installed typeface; Noto Nastaliq Urdu is recommended.
        <b>Font size</b> controls printed type size. <b>Line height</b> controls baseline spacing;
        1.8–2.3 is a useful Nastaleeq range. <b>Paragraph spacing</b> adds space after paragraphs.
        <b>Page size</b> chooses A5, B5, or A4. <b>Running header</b> prints an optional title.
        <b>First-line indent</b> indents normal paragraphs. <b>Word spacing</b> and
        <b>letter spacing</b> are fine adjustments; keep letter spacing at zero unless required.
        <b>Text alignment</b> supports justified book text, right alignment, or centered text.
        <b>Page numbers</b> enables automatic PDF footers.</p>
        <h2>Margins</h2>
        <p><b>Top / Bottom</b> reserve vertical page space. <b>Inner</b> is the binding-side
        margin and should normally be wider. <b>Outer</b> is the outside edge.</p>
        <h2>Paragraph tools</h2>
        <p>Select text in the reconstructed editor before using a tool; with no selection the
        full document is processed. <b>Join Selected Lines</b> removes hard line wraps.
        <b>Split Selected Text</b> inserts paragraph breaks at sentence punctuation.
        <b>Auto Make Paragraphs</b> waits for sentence punctuation after the chosen minimum
        source-line count. <b>Remove Extra Line Breaks</b> keeps blank-line paragraph boundaries
        while removing line wrapping inside them.</p>
        <h2>Advanced text editor</h2>
        <p><b>Undo, Redo, Cut, Copy, Paste, and Select All</b> use standard Windows shortcuts.
        <b>Find / Replace</b> supports previous/next navigation, matching case, whole words,
        one replacement, or replacing every occurrence. <b>Go to Line</b> jumps through long
        manuscripts. <b>Normalize Unicode</b> applies safe NFKC normalization to the selection,
        or the whole document when nothing is selected. <b>Text Statistics</b> reports word,
        character, line, and paragraph totals. <b>Zoom</b> changes only the editing view and does
        not affect the exported font size. The View menu can switch between right-to-left Urdu
        editing and left-to-right mixed-language editing. The editor footer continuously shows
        cursor line/column and document counts.</p>
        <h2>Clipboard, Word, and text import</h2>
        <p><b>Paste + Unicode (Ctrl+Shift+V)</b> reads Windows clipboard text, converts legacy
        presentation glyphs to standard Unicode, removes directional control artifacts, and
        inserts the result at the cursor. Use <b>Recover Legacy Visual Order</b> afterward only
        when the words themselves still appear reversed. <b>Import Word Document</b> reads DOCX
        paragraphs, headings, lists, and table rows in their original order. Visual Word styling
        becomes reliable editable text; old binary .doc files must first be saved as .docx.
        <b>Import Text File</b> supports UTF-8, UTF-16, and common Arabic Windows text.
        <b>Save Editor Text</b> produces portable UTF-8 text. <b>Clean Whitespace</b> removes
        repeated spaces and excessive blank lines, while <b>Duplicate Line</b> and
        <b>Delete Line</b> speed up manuscript editing.</p>
        <h2>Updates</h2>
        <p>Urdu Unicoder checks its public GitHub version manifest shortly after startup and
        displays a confirmation before installing anything. Source installations update the
        application files and dependencies, then restart automatically; saved .ubp projects and
        exported books are preserved. Packaged EXE installations open the official GitHub release
        download. Choose <b>Help → Check for Updates</b> to check manually at any time.</p>
        <h2>Preview, projects, and export</h2>
        <p><b>Refresh Preview</b> applies all current settings. <b>Save Project</b> stores text,
        layout, source path, and page range in a .ubp file. <b>Open Project</b> restores it.
        <b>Export Print PDF</b> creates a print-ready PDF. <b>Export HTML</b> creates portable
        UTF-8 RTL HTML. Always review a sample PDF before processing a complete book.</p>
        """)
        layout.addWidget(guide)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        dialog.exec()

    def show_about(self):
        dialog = QDialog(self)
        dialog.setWindowTitle(f"About {APP_NAME}")
        dialog.resize(540, 390)
        layout = QVBoxLayout(dialog)
        if LOGO_PATH.exists():
            logo = QLabel()
            logo.setAlignment(Qt.AlignCenter)
            logo.setPixmap(QPixmap(str(LOGO_PATH)).scaled(
                150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation
            ))
            layout.addWidget(logo)
        about = QTextBrowser()
        about.setOpenExternalLinks(True)
        about.setHtml(
            f"<div style='text-align:center'>"
            f"<h1>{APP_NAME}</h1><h3>Version {APP_VERSION}</h3>"
            f"<p>Advanced Unicode Urdu PDF reconstruction, RTL editing, "
            f"and professional book-layout software.</p>"
            f"<p><b>Created and maintained by {APP_AUTHOR}</b></p>"
            f"<p><a href='{AUTHOR_WEBSITE}'>Website: CyberOly.com</a><br>"
            f"<a href='{AUTHOR_GITHUB}'>GitHub: @MianAshfaq</a><br>"
            f"<a href='{AUTHOR_FACEBOOK}'>Facebook: @MianAshfaq012</a></p>"
            f"<p>Open-source software released under the MIT License.</p>"
            f"</div>"
        )
        layout.addWidget(about)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        dialog.exec()

    def closeEvent(self, event):
        if self.editor.toPlainText().strip():
            reply = QMessageBox.question(
                self, "Save project?",
                "Do you want to save the current project before closing?",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
            )
            if reply == QMessageBox.Cancel:
                event.ignore()
                return
            if reply == QMessageBox.Yes:
                if not self.save_project():
                    event.ignore()
                    return
        event.accept()


def main():
    set_windows_app_identity()
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName("Muhammad Ashfaq")
    if ICON_PATH.exists():
        app.setWindowIcon(QIcon(str(ICON_PATH)))
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
