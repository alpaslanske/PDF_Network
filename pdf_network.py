#!/usr/bin/env python3
import datetime
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import pymupdf
from PySide6.QtCore import Qt, QUrl, QEvent, QStandardPaths, QThread, Signal, QThreadPool
from PySide6.QtGui import QIcon, QDesktopServices, QPainter, QImage, QPixmap
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QHBoxLayout, QInputDialog, QLabel,
    QLineEdit, QListWidget, QListWidgetItem, QMainWindow, QMenu,
    QMessageBox, QPushButton, QSplitter, QStackedWidget, QTextBrowser,
    QTreeWidgetItem, QVBoxLayout, QWidget, QComboBox, QTabWidget, QProgressDialog,
    QScrollArea, QGridLayout
)
from PySide6.QtPrintSupport import QPrinter, QPrintDialog

from database import Database
from widgets.bookmark_tree import BookmarkTree
from widgets.pdf_viewer import PDFViewer
from lang import t, set_lang
from widgets.text_doc_viewer import TextDocViewer
from widgets.recent_file_card import RecentFileCard
from widgets.pdf_workers import TableExportWorker, PDFPrintWorker, ThumbnailWorker


class OfficeConvertWorker(QThread):
    finished = Signal(str, str, int, float, float)
    error = Signal(str)

    def __init__(self, input_path, page, x, y):
        super().__init__()
        self.input_path = input_path
        self.page = page
        self.x = x
        self.y = y

    def run(self):
        try:
            input_p = Path(self.input_path).resolve()

            cache_dir = Path(tempfile.gettempdir()) / "PDF_Network_Cache"
            cache_dir.mkdir(parents=True, exist_ok=True)

            file_stat = input_p.stat()
            unique_id = f"{input_p.stem}_{file_stat.st_size}_{int(file_stat.st_mtime)}"
            cached_pdf = cache_dir / f"{unique_id}.pdf"

            if cached_pdf.exists() and cached_pdf.stat().st_size > 0:
                self.finished.emit(str(cached_pdf), self.input_path, self.page, self.x, self.y)
                return

            temp_working_dir = Path(tempfile.mkdtemp(dir=cache_dir))

            # KESİN ÇÖZÜM: Orijinal dosyanın kilitlenmesini engellemek için kopya oluştur
            temp_input_file = temp_working_dir / input_p.name
            shutil.copy2(str(input_p), str(temp_input_file))

            expected_temp_pdf = temp_working_dir / f"{input_p.stem}.pdf"
            converted = False

            if sys.platform == "win32" and input_p.suffix.lower() == ".docx":
                try:
                    import win32com.client
                    import pythoncom
                    pythoncom.CoInitialize()
                    word = win32com.client.DispatchEx("Word.Application")
                    word.Visible = False

                    word_doc = word.Documents.Open(str(temp_input_file))
                    word_doc.SaveAs(str(expected_temp_pdf), FileFormat=17)
                    word_doc.Close()
                    word.Quit()
                    pythoncom.CoUninitialize()

                    if expected_temp_pdf.exists() and expected_temp_pdf.stat().st_size > 0:
                        converted = True
                except Exception:
                    pass

            if not converted:
                lo_profile = temp_working_dir / "lo_profile"
                env_arg = f"-env:UserInstallation=file://{lo_profile.as_posix()}"

                # ÇÖZÜM: Calc (Excel) dosyaları için "Fit to Page" parametresi eklendi.
                pdf_export_filter = "writer_pdf_Export"
                filter_data = ""

                if input_p.suffix.lower() in [".xlsx", ".xls", ".ods"]:
                    pdf_export_filter = "calc_pdf_Export"
                    filter_data = '{"FitWidth":{v:"1",type:"long"},"IsFitToPage":{v:"true",type:"boolean"}}'

                cmd = [
                    "soffice", env_arg, "--headless",
                    "--convert-to", f"pdf:{pdf_export_filter}:{filter_data}",
                    "--outdir", str(temp_working_dir),
                    str(temp_input_file)
                ]
                subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=25)

            # --- AKILLI FİLİGRAN VE ÇÖP TEMİZLEME ALGORİTMASI ---
            if expected_temp_pdf.exists() and expected_temp_pdf.stat().st_size > 0:
                try:
                    cleanup_doc = pymupdf.open(str(expected_temp_pdf))
                    for p_idx in range(len(cleanup_doc)):
                        p = cleanup_doc[p_idx]

                        found_watermarks = []
                        for pattern in ["Sayfa", "Sayfa ", "Page", "Page ", "Page 1", "Sayfa 1", "Sayfa 2"]:
                            found_watermarks.extend(p.search_for(pattern))

                        for watermark_rect in found_watermarks:
                            p.add_redact_annot(watermark_rect, fill=(1,1,1))

                        page_width = p.rect.width
                        page_height = p.rect.height

                        p.add_redact_annot(pymupdf.Rect(0, 0, 20, page_height), fill=(1,1,1))
                        p.add_redact_annot(pymupdf.Rect(page_width-20, 0, page_width, page_height), fill=(1,1,1))
                        p.add_redact_annot(pymupdf.Rect(0, 0, page_width, 20), fill=(1,1,1))
                        p.add_redact_annot(pymupdf.Rect(0, page_height-20, page_width, page_height), fill=(1,1,1))

                        p.apply_redactions()

                    cleanup_doc.save(str(expected_temp_pdf), garbage=4, deflate=True)
                    cleanup_doc.close()
                except Exception as clean_err:
                    print(f"Filigran temizleme hatası (atlandı): {clean_err}")

                shutil.move(str(expected_temp_pdf), str(cached_pdf))
                self.finished.emit(str(cached_pdf), self.input_path, self.page, self.x, self.y)
            else:
                self.error.emit("Dönüştürülen PDF boş veya oluşturulamadı.")

            shutil.rmtree(temp_working_dir, ignore_errors=True)

        except Exception as e:
            self.error.emit(str(e))


class NoteBrowser(QTextBrowser):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = None

    def mouseDoubleClickEvent(self, event):
        if self.main_window:
            self.main_window.enable_note_editing()
        super().mouseDoubleClickEvent(event)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = Database()

        self.thumbnail_pool = QThreadPool.globalInstance()
        self.thumbnail_pool.setMaxThreadCount(3)

        if "APPIMAGE" in os.environ:
            PROJE_ANA_DIZINI = Path(os.environ["APPIMAGE"]).parent
        else:
            PROJE_ANA_DIZINI = Path(os.path.realpath(__file__)).parent

        icon_path = PROJE_ANA_DIZINI / "icon.svg"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        self.tree = BookmarkTree()

        self.page_label = QLabel()
        self.zoom_label = QLabel()
        self.btn_open = QPushButton()
        self.btn_group = QPushButton()
        self.btn_prev = QPushButton()
        self.btn_next = QPushButton()
        self.btn_zoom_in = QPushButton()
        self.btn_zoom_out = QPushButton()
        self.btn_toggle = QPushButton()
        self.btn_print = QPushButton()
        self.btn_export = QPushButton()
        self.note_label = QLabel()
        self.btn_save_note = QPushButton()
        self.btn_search = QPushButton()
        self.dash_title = QLabel()
        self.dash_recent_lbl = QLabel()
        self.btn_dashboard = QPushButton()

        self.settings_file = PROJE_ANA_DIZINI / "data" / "settings.json"
        self.selected_group_id = None
        self.active_bookmark_id = None
        self.pdf_history = []
        self.current_pdf = None

        self.setup_ui()
        self.load_settings()
        self.metinleri_guncelle()
        self.load_groups()

        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.show_tree_menu)
        self.tree.itemClicked.connect(self.on_tree_item_clicked)

    def get_current_viewer(self):
        if self.tab_widget.count() > 0:
            return self.tab_widget.currentWidget()
        return None

    def metinleri_guncelle(self):
        base_title = t("app_title")
        if self.current_pdf:
            self.setWindowTitle(f"{base_title} - {Path(self.current_pdf).name}")
        else:
            self.setWindowTitle(base_title)

        self.update_page_label()
        self.update_zoom_label()

        self.btn_dashboard.setText(t("btn_dashboard"))
        self.update_dashboard_button()
        self.page_edit.setPlaceholderText(t("ph_page"))
        self.btn_open.setText(t("btn_open"))
        self.btn_group.setText(t("btn_group"))
        self.btn_prev.setText(t("btn_prev"))
        self.btn_next.setText(t("btn_next"))
        self.btn_zoom_in.setText(t("btn_zoom_in"))
        self.btn_zoom_out.setText(t("btn_zoom_out"))
        self.btn_toggle.setText(t("btn_toggle"))
        self.btn_print.setText(t("btn_print"))
        self.btn_export.setText(t("btn_export"))

        self.search_box.setPlaceholderText(t("ph_search_tree"))
        self.note_label.setText(t("lbl_note"))
        self.note_edit.setPlaceholderText(t("ph_note"))
        self.btn_save_note.setText(t("btn_save_note"))

        self.pdf_search.setPlaceholderText(t("ph_search_pdf"))
        self.btn_search.setText(t("btn_search_find"))

        self.dash_title.setText(t("dash_title"))
        self.dash_recent_lbl.setText(t("dash_recent_title"))
        self.tree.setHeaderLabel(t("tree_header"))

        self.update_dashboard_ui(self.pdf_history)
        self.load_groups()

    def lang_degistir(self, index):
        if index == 0: set_lang("tr")
        else: set_lang("en")
        self.metinleri_guncelle()
        self.save_settings()

    def setup_ui(self):
        central = QWidget()
        root = QVBoxLayout()
        toolbar = QHBoxLayout()

        self.page_edit = QLineEdit()
        self.page_edit.setFixedWidth(80)
        self.page_edit.returnPressed.connect(self.goto_page)

        toolbar.setContentsMargins(0, 0, 0, 0)
        toolbar.setSpacing(4)

        self.btn_dashboard.clicked.connect(self.toggle_dashboard)
        self.btn_open.clicked.connect(self.open_pdf)
        self.btn_prev.clicked.connect(self.previous_page)
        self.btn_next.clicked.connect(self.next_page)
        self.btn_zoom_in.clicked.connect(self.zoom_in)
        self.btn_zoom_out.clicked.connect(self.zoom_out)
        self.btn_group.clicked.connect(self.add_group)
        self.btn_toggle.clicked.connect(self.toggle_bookmarks)
        self.btn_print.clicked.connect(self.print_pdf)
        self.btn_export.clicked.connect(self.export_bookmarks_to_markdown)

        self.combo_lang = QComboBox()
        self.combo_lang.addItems(["🇹🇷 TR", "🇬🇧 EN"])
        self.combo_lang.setMinimumWidth(85)
        self.combo_lang.currentIndexChanged.connect(self.lang_degistir)

        for btn in (self.btn_dashboard, self.btn_print, self.btn_open, self.btn_prev, self.btn_next,
                    self.btn_zoom_in, self.btn_zoom_out, self.btn_group,
                    self.btn_toggle, self.btn_export, self.page_label,
                    self.zoom_label, self.combo_lang):
            btn.setMaximumHeight(28)

        toolbar.addWidget(self.btn_dashboard)
        toolbar.addWidget(self.btn_open)
        toolbar.addWidget(self.btn_print)
        toolbar.addWidget(self.btn_export)
        toolbar.addWidget(self.btn_group)
        toolbar.addWidget(self.btn_prev)
        toolbar.addWidget(self.btn_next)
        toolbar.addWidget(self.page_edit)
        toolbar.addWidget(self.btn_zoom_in)
        toolbar.addWidget(self.btn_zoom_out)
        toolbar.addWidget(self.btn_toggle)
        toolbar.addStretch()
        toolbar.addWidget(self.zoom_label)
        toolbar.addWidget(self.page_label)
        toolbar.addWidget(self.combo_lang)

        self.search_box = QLineEdit()
        self.search_box.textChanged.connect(self.filter_tree)

        self.left_panel = QWidget()
        self.left_panel.setMinimumWidth(200)
        left_layout = QVBoxLayout(self.left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(4)

        self.note_label.setStyleSheet("font-weight: bold; margin-top: 5px;")

        self.note_edit = NoteBrowser()
        self.note_edit.main_window = self
        self.note_edit.setReadOnly(True)
        self.note_edit.setOpenLinks(False)
        self.note_edit.setMaximumHeight(150)
        self.note_edit.setEnabled(False)
        self.note_edit.anchorClicked.connect(self.handle_knowledge_link)

        self.btn_save_note.setMaximumHeight(28)
        self.btn_save_note.setEnabled(False)
        self.btn_save_note.clicked.connect(self.save_current_bookmark_note)

        left_layout.addWidget(self.search_box)
        left_layout.addWidget(self.tree, 1)
        left_layout.addWidget(self.note_label)
        left_layout.addWidget(self.note_edit)
        left_layout.addWidget(self.btn_save_note)

        self.pdf_search = QLineEdit()
        self.pdf_search.returnPressed.connect(self.search_in_pdf)
        self.btn_search.clicked.connect(self.search_in_pdf)

        self.search_results = QListWidget()
        self.search_results.itemDoubleClicked.connect(self.goto_search_result)

        search_panel = QWidget()
        search_layout = QVBoxLayout(search_panel)
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.addWidget(self.pdf_search)
        search_layout.addWidget(self.btn_search)
        search_layout.addWidget(self.search_results)

        self.splitter = QSplitter()
        self.splitter.addWidget(self.left_panel)

        self.right_stack = QStackedWidget()

        self.dashboard_widget = QWidget()
        dash_layout = QVBoxLayout(self.dashboard_widget)
        dash_layout.setContentsMargins(20, 20, 20, 20)
        dash_layout.setSpacing(15)

        self.dash_title.setStyleSheet("font-size: 20px; font-weight: bold; color: #2c3e50;")
        dash_layout.addWidget(self.dash_title)

        self.stats_label = QLabel()
        self.stats_label.setStyleSheet("font-size: 13px; color: #7f8c8d; font-style: italic;")
        dash_layout.addWidget(self.stats_label)

        dash_layout.addWidget(self.dash_recent_lbl)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("background-color: transparent; border: none;")

        self.grid_widget = QWidget()
        self.grid_widget.setStyleSheet("background-color: transparent;")
        self.grid_layout = QGridLayout(self.grid_widget)
        self.grid_layout.setSpacing(15)
        self.grid_layout.setAlignment(Qt.AlignLeft | Qt.AlignTop)

        self.scroll_area.setWidget(self.grid_widget)
        dash_layout.addWidget(self.scroll_area, 1)

        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.tabCloseRequested.connect(self.close_tab)
        self.tab_widget.currentChanged.connect(self.on_tab_changed)

        self.right_stack.addWidget(self.dashboard_widget)
        self.right_stack.addWidget(self.tab_widget)
        self.right_stack.currentChanged.connect(self.update_dashboard_button)

        self.splitter.addWidget(self.right_stack)
        self.splitter.addWidget(search_panel)

        self.splitter.setSizes([250, 900, 300])
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 1)

        root.addLayout(toolbar)
        root.addWidget(self.splitter, 1)

        central.setLayout(root)
        self.setCentralWidget(central)
        self.resize(1400, 900)

    def on_tab_changed(self, index):
        viewer = self.get_current_viewer()
        if viewer:
            self.current_pdf = getattr(viewer, "original_file_path", viewer.current_file_path)
            self.setWindowTitle(f"{t('app_title')} - {Path(self.current_pdf).name}")
            self.update_page_label()
            self.update_zoom_label()
            self.load_pdf_bookmarks()
            self.btn_zoom_in.setEnabled(True)
            self.btn_zoom_out.setEnabled(True)
        else:
            self.current_pdf = None
            self.setWindowTitle(t("app_title"))
            self.page_label.setText(t("lbl_page_empty"))
            self.page_edit.clear()
        self.update_dashboard_button()

    def on_tree_item_clicked(self, item):
        group_id = item.data(0, Qt.UserRole)
        if group_id is not None:
            self.selected_group_id = group_id
        self.open_bookmark(item)

    def toggle_dashboard(self):
        if self.right_stack.currentIndex() == 1:
            self.right_stack.setCurrentIndex(0)
            self.setWindowTitle(t("app_title"))
        else:
            if self.tab_widget.count() > 0:
                self.right_stack.setCurrentIndex(1)
                if self.current_pdf:
                    self.setWindowTitle(f"{t('app_title')} - {Path(self.current_pdf).name}")

    def update_dashboard_button(self, index=None):
        if self.right_stack.currentIndex() == 0:
            if self.tab_widget.count() > 0:
                self.btn_dashboard.setText(t("btn_back_to_doc"))
                self.btn_dashboard.setEnabled(True)
            else:
                self.btn_dashboard.setText(t("btn_dashboard"))
                self.btn_dashboard.setEnabled(False)
        else:
            self.btn_dashboard.setText(t("btn_dashboard"))
            self.btn_dashboard.setEnabled(True)

    def add_to_history(self, pdf_path, page):
        history = getattr(self, "pdf_history", [])
        history = [item for item in history if (item.get("path") if isinstance(item, dict) else item) != pdf_path]
        history.insert(0, {
            "path": pdf_path,
            "page": page,
            "date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        })
        self.pdf_history = history[:5]
        self.update_dashboard_ui(self.pdf_history)

    def close_tab(self, index):
        viewer = self.tab_widget.widget(index)
        if viewer:
            if hasattr(viewer, "doc") and viewer.doc:
                viewer.doc.close()
            viewer.deleteLater()
        self.tab_widget.removeTab(index)
        if self.tab_widget.count() == 0:
            self.right_stack.setCurrentIndex(0)

    def open_in_tab(self, file_path, page=0, x=0, y=0):
        for i in range(self.tab_widget.count()):
            viewer = self.tab_widget.widget(i)
            if getattr(viewer, "original_file_path", viewer.current_file_path) == file_path:
                self.tab_widget.setCurrentIndex(i)
                viewer.goto_position(page, x, y)
                self.right_stack.setCurrentIndex(1)
                self.add_to_history(file_path, page)
                return i

        suffix = Path(file_path).suffix.lower()
        is_office_doc = suffix in [".docx", ".odt", ".xlsx", ".ods", ".pptx", ".ppt", ".odp", ".odg", ".odf", ".rtf", ".csv"]

        if is_office_doc:
            self.convert_progress = QProgressDialog(t("lbl_loading"), None, 0, 0, self)
            self.convert_progress.setWindowTitle(t("lbl_loading"))
            self.convert_progress.setWindowModality(Qt.WindowModal)
            self.convert_progress.show()

            self.office_worker = OfficeConvertWorker(file_path, page, x, y)
            self.office_worker.finished.connect(self._on_office_converted)
            self.office_worker.error.connect(self._on_office_error)
            self.office_worker.start()
            return -1

        self._load_pdf_into_tab(file_path, file_path, False, page, x, y)
        return self.tab_widget.currentIndex()

    def _on_office_converted(self, pdf_path, original_path, page, x, y):
        if hasattr(self, "convert_progress"): self.convert_progress.close()
        self._load_pdf_into_tab(pdf_path, original_path, True, page, x, y)

    def _on_office_error(self, err_msg):
        if hasattr(self, "convert_progress"): self.convert_progress.close()
        QMessageBox.critical(self, t("err_convert_title"), t("err_convert_msg") + f"\n\nSistem Detayı: {err_msg}")

    def _load_pdf_into_tab(self, active_open_path, original_path, is_office_doc, page, x, y):
        new_viewer = PDFViewer()
        new_viewer.original_file_path = original_path
        new_viewer.is_office_doc = is_office_doc

        new_viewer.zoom_changed.connect(self.update_zoom_label)
        new_viewer.bookmarkRequested.connect(self.add_bookmark_from_selection)
        new_viewer.highlightRequested.connect(self.create_highlight)
        new_viewer.deleteHighlightRequested.connect(self.delete_highlight)
        new_viewer.exportTableRequested.connect(self.export_table_to_excel)
        new_viewer.addNoteRequested.connect(self.add_page_note)
        new_viewer.page_changed.connect(self.update_page_label)

        new_viewer.open_pdf(active_open_path)
        new_viewer.goto_position(page, x, y)

        tab_title = Path(original_path).name
        index = self.tab_widget.addTab(new_viewer, tab_title)
        self.tab_widget.setCurrentIndex(index)
        self.right_stack.setCurrentIndex(1)
        self.add_to_history(original_path, page)
        return index

    def open_pdf(self):
        default_dir = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)

        file_filter = (
            "Tüm Desteklenen Belgeler (*.pdf *.xlsx *.ods *.docx *.odt *.pptx *.odp *.odg *.odf *.epub *.mobi *.fb2 *.cbz *.txt *.rtf *.csv);;"
            "PDF ve E-Kitaplar (*.pdf *.epub *.mobi *.fb2 *.cbz);;"
            "Desteklenen Ofis ve PDF Belgeleri (*.pdf *.ods *.odt *.odp *.odg *.odf *.xlsx *.docx *.pptx);;"
            "Tüm Dosyalar (*.*)"
        )
        
# LibreOffice
# *.ods *.odt *.odp *.odg *.odf
# MS OFfice
# *.xlsx *.docx *.pptx
# e-Kitap
# *.epub *.mobi *.fb2 *.cbz
# Rich Text
# *.txt *.rtf *.csv

        filename, _ = QFileDialog.getOpenFileName(self, t("btn_open"), default_dir, file_filter)
        if not filename:
            return

        self.open_in_tab(filename)

    def load_pdf_bookmarks(self):
        pdf_root_name = t("pdf_bookmarks_root")
        from lang import get_all_translations_for_key
        possible_roots = get_all_translations_for_key("pdf_bookmarks_root")

        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            if item.text(0) in possible_roots:
                self.tree.takeTopLevelItem(i)
                break

        viewer = self.get_current_viewer()
        if not viewer or not viewer.doc: return

        toc = viewer.doc.get_toc()
        self.pdf_root = QTreeWidgetItem([pdf_root_name])
        self.tree.addTopLevelItem(self.pdf_root)
        parents = {0: self.pdf_root}

        for level, title, page in toc:
            item = QTreeWidgetItem([title])
            item.setData(0, Qt.UserRole + 100, page - 1)
            parent_item = parents.get(level - 1, self.pdf_root)
            parent_item.addChild(item)
            parents[level] = item
        self.pdf_root.setExpanded(True)

    def zoom_in(self):
        viewer = self.get_current_viewer()
        if viewer: viewer.zoom_in(); self.update_zoom_label()

    def zoom_out(self):
        viewer = self.get_current_viewer()
        if viewer: viewer.zoom_out(); self.update_zoom_label()

    def update_zoom_label(self):
        viewer = self.get_current_viewer()
        if viewer: self.zoom_label.setText(t("lbl_zoom").format(int(viewer.zoom * 100)))
        else: self.zoom_label.setText("")

    def next_page(self):
        viewer = self.get_current_viewer()
        if viewer: viewer.next_page(); self.update_page_label()

    def previous_page(self):
        viewer = self.get_current_viewer()
        if viewer: viewer.previous_page(); self.update_page_label()

    def update_page_label(self):
        viewer = self.get_current_viewer()
        if not viewer or not viewer.doc:
            self.page_label.setText(t("lbl_page_empty"))
            self.page_edit.clear()
            return
        if self.sender() and isinstance(self.sender(), PDFViewer) and self.sender() != viewer: return
        current = viewer.current_page + 1
        total = len(viewer.doc)
        self.page_label.setText(t("lbl_page").format(current, total))
        self.page_edit.setText(str(current))

    def add_group(self):
        name, ok = QInputDialog.getText(self, t("dlg_new_group"), t("dlg_group_name"))
        if ok and name.strip(): self.db.add_group(name.strip()); self.load_groups()

    def load_groups(self):
        self.tree.clear()
        groups = self.db.get_groups()
        nodes = {}
        for gid, name, parent_id in groups:
            item = QTreeWidgetItem([name])
            item.setData(0, Qt.UserRole, gid)
            nodes[gid] = item
        for gid, name, parent_id in groups:
            item = nodes[gid]
            if parent_id and parent_id in nodes: nodes[parent_id].addChild(item)
            else: self.tree.addTopLevelItem(item)
        bookmarks = self.db.get_bookmarks()
        for bookmark in bookmarks:
            bid, title, pdf_path, page, group_id = bookmark[0], bookmark[1], bookmark[2], bookmark[3], bookmark[4]
            if group_id in nodes:
                bookmark_item = QTreeWidgetItem([title])
                bookmark_item.setData(0, Qt.UserRole + 1, bid)
                nodes[group_id].addChild(bookmark_item)
        self.load_pdf_bookmarks()

    def toggle_bookmarks(self):
        sizes = self.splitter.sizes()
        if sizes[0] > 0: self._last_left_width = sizes[0]; sizes[0] = 0
        else: sizes[0] = getattr(self, "_last_left_width", 250)
        self.splitter.setSizes(sizes)

    def save_settings(self):
        sizes = self.splitter.sizes()
        is_visible = sizes[0] > 0
        if not is_visible and hasattr(self, "_last_left_width"): sizes[0] = self._last_left_width
        data = {
            "window_width": self.width(), "window_height": self.height(),
            "splitter_sizes": sizes, "bookmarks_visible": is_visible,
            "pdf_history": getattr(self, "pdf_history", []),
            "last_active_index": self.right_stack.currentIndex(),
            "language": "tr" if self.combo_lang.currentIndex() == 0 else "en"
        }
        with open(self.settings_file, "w", encoding="utf-8") as f: json.dump(data, f, indent=4)

    def load_settings(self):
        if not self.settings_file.exists(): self.update_dashboard_ui([]); return
        try:
            with open(self.settings_file, "r", encoding="utf-8") as f: data = json.load(f)
            set_lang(data.get("language", "tr"))
            self.combo_lang.blockSignals(True)
            self.combo_lang.setCurrentIndex(0 if data.get("language") == "tr" else 1)
            self.combo_lang.blockSignals(False)
            self.resize(data.get("window_width", 1400), data.get("window_height", 900))
            self.pdf_history = data.get("pdf_history", [])
            self.update_dashboard_ui(self.pdf_history)
        except: self.update_dashboard_ui([])

    def update_dashboard_ui(self, history):
        while self.grid_layout.count():
            item = self.grid_layout.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        try: self.stats_label.setText(t("dash_stats").format(self.db.get_bookmark_count(), self.db.get_group_count()))
        except: pass
        self.thumbnail_pool.clear()
        if not history:
            lbl = QLabel(t("dash_empty"))
            lbl.setStyleSheet("color: #7f8c8d; font-style: italic;")
            self.grid_layout.addWidget(lbl, 0, 0)
            return
        for idx, d in enumerate(history):
            if not isinstance(d, dict): continue
            card = RecentFileCard(d)
            card.clicked.connect(self.open_from_dashboard_card)
            card.delete_requested.connect(self.delete_from_dashboard_card)
            self.grid_layout.addWidget(card, idx // 4, idx % 4)
            self.thumbnail_pool.start(ThumbnailWorker(d.get("path", ""), card))

    def _on_pencere_thumbnail_ready(self, samples, width, height, stride, status, card_ref):
        if card_ref and status == "success" and samples:
            img = QImage(samples, width, height, stride, QImage.Format_RGB888)
            card_ref.set_pixmap(QPixmap.fromImage(img).scaled(256, 130, Qt.KeepAspectRatio, Qt.SmoothTransformation), "success")

    def open_from_dashboard_card(self, d):
        if d and Path(d.get("path", "")).exists(): self.open_in_tab(d["path"], d.get("page", 0))

    def delete_from_dashboard_card(self, d):
        self.pdf_history = [i for i in self.pdf_history if i.get("path") != d.get("path")]
        self.update_dashboard_ui(self.pdf_history); self.save_settings()

    def closeEvent(self, event): self.save_settings(); super().closeEvent(event)

    def search_in_pdf(self):
        text_val = self.pdf_search.text().strip()
        if not text_val or not self.get_current_viewer(): return
        pages = self.get_current_viewer().search_text(text_val)
        self.search_results.clear()
        self.search_results.addItem(f"Toplam Eşleşme: {len(pages)}")
        for p, r in pages:
            itm = QListWidgetItem(f"Sayfa {p + 1}")
            itm.setData(Qt.UserRole, (p, [r]))
            self.search_results.addItem(itm)

    def goto_search_result(self, item):
        v = self.get_current_viewer()
        if v and item.data(Qt.UserRole):
            p, r = item.data(Qt.UserRole); v.goto_page(p); v.set_highlight(r); v.render_page()

    def filter_tree(self, text):
        for i in range(self.tree.topLevelItemCount()): self.filter_item(self.tree.topLevelItem(i), text.lower().strip())

    def filter_item(self, item, text):
        vis = text in item.text(0).lower()
        c_vis = any(self.filter_item(item.child(i), text) for i in range(item.childCount()))
        item.setHidden(not (vis or c_vis or text == ""))
        return vis or c_vis

    def show_tree_menu(self, pos):
        item = self.tree.itemAt(pos)
        
        # --- BOŞ ALANA SAĞ TIKLAMA KONTROLÜ ---
        if not item:
            m = QMenu()
            a_new_gr = m.addAction(t("btn_group")) # Dil dosyasından "Grup Ekle" metnini alır
            action = m.exec(self.tree.mapToGlobal(pos))
            
            if action == a_new_gr:
                self.add_group() # Mevcut ana grup ekleme fonksiyonunu tetikler
            return

        m = QMenu()
        group_id = item.data(0, Qt.UserRole)
        bookmark_id = item.data(0, Qt.UserRole + 1)
        pdf_page = item.data(0, Qt.UserRole + 100)

        if pdf_page is not None:
            return

        if group_id is not None:
            # --- GRUP İÇİN SAĞ TIK MENÜSÜ ---
            a_bm = m.addAction(t("menu_add_bookmark"))
            a_sub = m.addAction(t("menu_add_subgroup"))
            r_gr = m.addAction(t("dlg_rename_group"))   
            d_gr = m.addAction(t("msg_delete_group_confirm").format(item.text(0))) 
            
            action = m.exec(self.tree.mapToGlobal(pos))
            
            if action == a_bm: 
                self.add_bookmark(item)
            elif action == a_sub:
                self.add_subgroup_action(item)
            elif action == r_gr: 
                self.rename_group_action(item)
            elif action == d_gr: 
                self.delete_group_action(item)

        elif bookmark_id is not None:
            r_bm = m.addAction(t("menu_rename_bookmark"))
            l_bm = m.addAction("🔗 Link Ekle (Çapraz PDF)") # YENİ EKLENEN SATIR
            d_bm = m.addAction(t("menu_del_bookmark"))
            
            action = m.exec(self.tree.mapToGlobal(pos))
            
            if action == r_bm: 
                self.rename_bookmark(item)
            elif action == l_bm:                   # YENİ EKLENEN SATIR
                self.add_cross_pdf_link(item)      # YENİ EKLENEN SATIR
            elif action == d_bm: 
                self.delete_bookmark(item)


    # ---------------- EKSİK OLAN VE GERİ GETİRİLEN FONKSİYONLAR ----------------
    def add_bookmark(self, item):
        viewer = self.get_current_viewer()
        if not viewer or not item: return
        current_page = viewer.current_page
        group_id = item.data(0, Qt.UserRole)
        title, ok = QInputDialog.getText(self, t("dlg_new_bookmark"), t("dlg_bookmark_name"))
        if not ok or not title.strip():
            return

        kaydedilecek_yol = getattr(viewer, "original_file_path", viewer.current_file_path)

        self.db.add_bookmark(title.strip(), kaydedilecek_yol, current_page, group_id)
        self.load_groups()
        viewer.goto_page(current_page)
        self.update_page_label()

    def delete_bookmark(self, item):
        if not item: return
        bookmark_id = item.data(0, Qt.UserRole + 1)
        if not bookmark_id: return
        self.db.delete_bookmark(bookmark_id)
        self.load_groups()

    def rename_bookmark(self, item):
        if not item: return
        bookmark_id = item.data(0, Qt.UserRole + 1)
        if not bookmark_id: return
        name, ok = QInputDialog.getText(self, t("dlg_rename_bookmark"), t("dlg_new_name"), text=item.text(0))
        if ok and name.strip():
            self.db.rename_bookmark(bookmark_id, name.strip())
            self.load_groups()

    def open_bookmark(self, item):
        if QApplication.mouseButtons() == Qt.RightButton:
            return

        pdf_page = item.data(0, Qt.UserRole + 100)
        if pdf_page is not None:
            viewer = self.get_current_viewer()
            if viewer:
                viewer.goto_page(pdf_page)
                self.update_page_label()
                self.right_stack.setCurrentIndex(1)
            return

        bookmark_id = item.data(0, Qt.UserRole + 1)
        if not bookmark_id:
            self.active_bookmark_id = None
            self.note_edit.clear()
            self.note_edit.setEnabled(False)
            self.btn_save_note.setEnabled(False)
            return

        bookmark = self.db.get_bookmark(bookmark_id)
        if not bookmark:
            return

        self.active_bookmark_id = bookmark_id
        pdf_path = bookmark[2]
        page = bookmark[3]
        x = bookmark[5] if len(bookmark) > 5 else 0
        y = bookmark[6] if len(bookmark) > 6 else 0
        current_note = bookmark[7] if len(bookmark) > 7 and bookmark[7] else ""

        if "[[" in current_note and "]]" in current_note and "#" in current_note:
            try:
                link_match = re.search(r'\[\[([^\]]+)\]\]', current_note)
                if link_match:
                    _, page_str = link_match.group(1).split("#", 1)
                    clean_page_str = "".join(filter(str.isdigit, page_str))
                    if clean_page_str:
                        page = int(clean_page_str) - 1
            except Exception as e:
                print("Link çözümleme hatası:", e)

        path_check = Path(pdf_path)
        if not path_check.exists():
            reply = QMessageBox.question(
                self, t("msg_error"), t("err_file_relocate").format(pdf_path),
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
            )
            if reply == QMessageBox.Yes:
                new_filename, _ = QFileDialog.getOpenFileName(self, t("dlg_file_relocate"), QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation), "PDF Files (*.pdf)")
                if not new_filename: return
                self.db.update_pdf_path(pdf_path, new_filename)
                pdf_path = new_filename
            else:
                return

        self.open_in_tab(pdf_path, page, x, y)

        self.note_edit.setReadOnly(True)
        html_note = current_note
        matches = re.findall(r'(\[\[([^\]]+)\]\])', current_note)
        for full_match, link_content in matches:
            html_link = f'<a href="{link_content}" style="color: #0078ff; font-weight: bold; text-decoration: none;">{full_match}</a>'
            html_note = html_note.replace(full_match, html_link)

        self.note_edit.setHtml(html_note.replace("\n", "<br>"))
        self.note_edit.setEnabled(True)
        self.btn_save_note.setEnabled(True)

    def rename_group_action(self, item):
        if not item: 
            return
        group_id = item.data(0, Qt.UserRole)
        if group_id is None: 
            return
            
        # text=item.text(0) parametresi ile mevcut grup adı kutucuğa otomatik doldurulur
        name, ok = QInputDialog.getText(
            self, 
            t("dlg_rename_group"), 
            t("dlg_group_new_name"), 
            text=item.text(0)
        )
        if ok and name.strip():
            self.db.rename_group(group_id, name.strip())
            self.load_groups()


    def delete_group_action(self, item):
        if not item: return
        group_id = item.data(0, Qt.UserRole)
        if group_id is None: return
        msg_text = t("msg_delete_group_confirm").format(item.text(0))
        reply = QMessageBox.question(self, t("msg_warning"), msg_text, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            bookmarks = self.db.get_bookmarks()
            for b in bookmarks:
                if len(b) > 4 and b[4] == group_id:
                    self.db.delete_bookmark(b[0])
            self.db.delete_group(group_id)
            self.load_groups()

    def enable_note_editing(self):
        if not self.active_bookmark_id or not self.note_edit.isReadOnly(): return
        bookmark_data = self.db.get_bookmark(self.active_bookmark_id)
        if bookmark_data and len(bookmark_data) > 7:
            current_note = bookmark_data[7] or ""
            self.note_edit.setReadOnly(False)
            self.note_edit.setTextInteractionFlags(Qt.TextEditable | Qt.TextSelectableByMouse)
            self.note_edit.setPlainText(current_note)

    def save_current_bookmark_note(self):
        if not self.active_bookmark_id: return
        try:
            notes_text = self.note_edit.toPlainText().strip()
            self.db.update_bookmark_notes(self.active_bookmark_id, notes_text)
            self.note_edit.setReadOnly(True)
            self.note_edit.setTextInteractionFlags(Qt.LinksAccessibleByMouse)

            html_note = notes_text
            matches = re.findall(r'(\[\[([^\]]+)\]\])', notes_text)
            for full_match, link_content in matches:
                html_link = f'<a href="{link_content}" style="color: #0078ff; font-weight: bold; text-decoration: none;">{full_match}</a>'
                html_note = html_note.replace(full_match, html_link)
            self.note_edit.setHtml(html_note.replace("\n", "<br>"))
        except Exception as e:
            QMessageBox.warning(self, t("msg_error"), t("err_save_note").format(e))

    def handle_knowledge_link(self, url):
        try:
            raw_str = url.toString().replace("link://", "").strip().rstrip("/")
            if not raw_str and url.isValid():
                raw_str = url.toDisplayString().replace("link://", "").strip().rstrip("/")
            if not raw_str: return

            if "#" not in raw_str:
                target_file, target_page = raw_str, 0
            else:
                target_file, page_str = raw_str.split("#", 1)
                target_file = target_file.strip()
                clean_page_str = "".join(filter(str.isdigit, page_str))
                target_page = int(clean_page_str) - 1 if clean_page_str else 0

            viewer = self.get_current_viewer()
            if viewer and (Path(viewer.current_file_path).name == target_file or viewer.current_file_path == target_file):
                viewer.goto_page(target_page)
                self.update_page_label()
                self.right_stack.setCurrentIndex(1)
                return

            full_pdf_path = None
            if viewer:
                sibling_path = Path(viewer.current_file_path).parent / target_file
                if sibling_path.exists() and sibling_path.is_file():
                    full_pdf_path = str(sibling_path)

            if not full_pdf_path:
                for item in getattr(self, "pdf_history", []):
                    p_str = item.get("path", "") if isinstance(item, dict) else item
                    if p_str and (p_str.endswith(target_file) or Path(p_str).name == target_file):
                        if Path(p_str).exists() and Path(p_str).is_file():
                            full_pdf_path = p_str
                            break

            if not full_pdf_path:
                QMessageBox.warning(self, t("msg_error"), t("err_target_pdf").format(target_file))
                return

            self.open_in_tab(full_pdf_path, target_page)
        except Exception as e:
            print("Knowledge link hatası:", e)

    def add_bookmark_from_selection(self, page, text, x, y):
        viewer = self.sender()
        if not viewer: return
        if self.selected_group_id is None:
            QMessageBox.warning(self, t("msg_warning"), t("warn_select_group"))
            return
        kaydedilecek_yol = getattr(viewer, "original_file_path", viewer.current_file_path)
        self.db.add_bookmark(text, kaydedilecek_yol, page, self.selected_group_id, x, y)
        self.load_groups()
        viewer.goto_page(page)
        self.update_page_label()

    def _save_and_reload_pdf(self, viewer):
        try:
            if getattr(viewer.doc, "is_repaired", False):
                temp_path = viewer.current_file_path + ".tmp.pdf"
                viewer.doc.save(temp_path, garbage=4, deflate=True)
                viewer.doc.close()
                import os
                os.replace(temp_path, viewer.current_file_path)
                viewer.doc = pymupdf.open(viewer.current_file_path)
            else:
                viewer.doc.saveIncr()

            if hasattr(viewer.doc, "reload"):
                viewer.doc.reload()

            viewer.render_page()
        except Exception as e:
            QMessageBox.warning(self, t("msg_error"), t("err_save_failed").format(e))

    def create_highlight(self, page_no, x0, y0, x1, y1):
        viewer = self.sender()
        if not viewer or not viewer.doc: return
        try:
            page = viewer.doc[page_no]
            rect = pymupdf.Rect(x0, y0, x1, y1)
            annot = page.add_rect_annot(rect)
            annot.set_colors(stroke=(1, 1, 0), fill=(1, 1, 0))
            annot.set_opacity(0.4)
            annot.update()
            self._save_and_reload_pdf(viewer)
        except Exception as e:
            QMessageBox.warning(self, t("msg_error"), t("err_highlight_readonly").format(e))

    def delete_highlight(self, page_no, x, y):
        viewer = self.sender()
        if not viewer or not viewer.doc: return
        try:
            page = viewer.doc[page_no]
            point = pymupdf.Point(x, y)
            annot = page.first_annot
            while annot:
                next_annot = annot.next
                if annot.rect.contains(point):
                    page.delete_annot(annot)
                    self._save_and_reload_pdf(viewer)
                    return
                annot = next_annot
        except Exception as e:
            QMessageBox.warning(self, t("msg_error"), t("err_delete_failed").format(e))

    def add_page_note(self, page_no, x, y):
        viewer = self.sender()
        if not viewer or not viewer.doc: return

        text, ok = QInputDialog.getMultiLineText(self, t("dlg_add_note_title"), t("dlg_add_note_msg"))
        if not ok or not text.strip(): return

        try:
            page = viewer.doc[page_no]
            lines = text.strip().split('\n')
            max_len = max([len(line) for line in lines] + [1])
            total_lines = sum([max(1, len(line) // 35) for line in lines])
            width = min(250, max(100, max_len * 7))
            height = max(40, total_lines * 16 + 15)
            rect = pymupdf.Rect(x, y, x + width, y + height)
            rect.normalize()
            annot = page.add_freetext_annot(
                rect, text.strip(), fontsize=11, fontname="helv",
                text_color=(1, 0, 0), fill_color=(1, 1, 0.8)
            )
            annot.set_border(width=1)
            annot.update()
            self._save_and_reload_pdf(viewer)
        except Exception as e:
            QMessageBox.warning(self, t("msg_error"), t("err_add_note_failed").format(e))

    def print_pdf(self):
        viewer = self.get_current_viewer()
        if not viewer or not viewer.doc: return
        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer, self)
        dialog.setWindowTitle(t("print_title"))
        if dialog.exec() == QPrintDialog.Accepted:
            from_page = printer.fromPage() - 1
            to_page = printer.toPage() - 1
            if from_page < 0: from_page = 0
            if to_page < 0 or to_page >= len(viewer.doc): to_page = len(viewer.doc) - 1
            self.progress_dialog = QProgressDialog(t("prog_print_wait"), None, 0, 0, self)
            self.progress_dialog.setWindowTitle(t("print_title"))
            self.progress_dialog.setWindowModality(Qt.WindowModal)
            self.progress_dialog.show()
            self.print_worker = PDFPrintWorker(viewer.current_file_path, printer, from_page, to_page)
            self.print_worker.finished.connect(self._on_print_finished)
            self.print_worker.error.connect(self._on_print_error)
            self.print_worker.start()

    def _on_print_finished(self, printer_name):
        if hasattr(self, "progress_dialog"): self.progress_dialog.close()
        QMessageBox.information(self, t("print_title"), t("print_all_success").format(printer_name))

    def _on_print_error(self, err_msg):
        if hasattr(self, "progress_dialog"): self.progress_dialog.close()
        QMessageBox.warning(self, t("print_err"), err_msg)

# ---- BURADAN İTİBAREN AŞAĞIDAKİ BLOĞU EKLİYORUZ ----
    def goto_page(self):
        viewer = self.get_current_viewer()
        if not viewer or not viewer.doc: return
        try:
            page_no = int(self.page_edit.text())
            if 1 <= page_no <= viewer.doc.page_count:
                viewer.current_page = page_no - 1
                viewer.render_page()
        except ValueError: pass
        finally: self.update_page_label()
    # ----------------------------------------------------

    def export_bookmarks_to_markdown(self):
        viewer = self.get_current_viewer()
        if not viewer:
            QMessageBox.warning(self, t("msg_warning"), t("export_warn_open"))
            return
        cpdf = getattr(viewer, "original_file_path", viewer.current_file_path)
        file_name = Path(cpdf).stem
        default_save = PROJE_ANA_DIZINI / f"{file_name}_Calisma_Notlari.md"
        save_filename, _ = QFileDialog.getSaveFileName(self, t("export_save_title"), str(default_save), "Markdown Files (*.md)")
        if not save_filename: return
        try:
            bookmarks = self.db.get_bookmarks()
            current_bookmarks = [b for b in bookmarks if b[2] == cpdf]
            if not current_bookmarks:
                QMessageBox.information(self, t("msg_info"), t("export_no_bookmarks"))
                return
            groups = {g[0]: g[1] for g in self.db.get_groups()}
            current_bookmarks.sort(key=lambda x: x[3])
            md = t("export_md_title")
            md += t("export_md_doc").format(Path(cpdf).name)
            md += t("export_md_path").format(cpdf)
            md += t("export_md_date").format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M'))
            md += t("export_md_header")
            for bookmark in current_bookmarks:
                title, page, group_id = bookmark[1], bookmark[3], bookmark[4]
                group_name = groups.get(group_id, t("export_general"))
                md += f"### 📘 {title}\n"
                md += t("export_md_group").format(group_name, page + 1)
                note_text = bookmark[7].strip() if len(bookmark) > 7 and bookmark[7] else None
                if note_text:
                    md += t("export_md_note_title")
                    indented_note = "\n".join([f"  > {line}" for line in note_text.split("\n")])
                    md += f"{indented_note}\n"
                else:
                    md += t("export_md_no_note")
                md += "\n"
            with open(save_filename, "w", encoding="utf-8") as f:
                f.write(md)
            QMessageBox.information(self, t("msg_success"), t("export_success").format(Path(save_filename).name))
        except Exception as e:
            QMessageBox.warning(self, t("msg_error"), t("export_err").format(e))

    def export_table_to_excel(self, page_no, x0, y0, x1, y1):
        viewer = self.sender()
        if not viewer or not viewer.doc: return
        try:
            import pandas as pd
        except ImportError:
            QMessageBox.warning(self, t("exc_lib_err"), t("exc_lib_msg"))
            return
        docs_dir = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
        default_save = str(Path(docs_dir) / f"{Path(viewer.current_file_path).stem}_Tablo_Sayfa{page_no + 1}.xlsx")
        save_filename, _ = QFileDialog.getSaveFileName(self, t("exc_save_title"), default_save, "Excel Files (*.xlsx)")
        if not save_filename: return
        self.table_progress = QProgressDialog(t("prog_table_wait"), None, 0, 0, self)
        self.table_progress.setWindowTitle(t("exc_save_title"))
        self.table_progress.setWindowModality(Qt.WindowModal)
        self.table_progress.show()
        self.table_worker = TableExportWorker(viewer.current_file_path, page_no, x0, y0, x1, y1, save_filename)
        self.table_worker.finished.connect(self._on_table_finished)
        self.table_worker.error.connect(self._on_table_error)
        self.table_worker.no_table.connect(self._on_table_missing)
        self.table_worker.start()

    def _on_table_finished(self, save_filename):
        if hasattr(self, "table_progress"): self.table_progress.close()
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(t("msg_success"))
        msg_box.setText(t("exc_success_msg").format(Path(save_filename).name))
        msg_box.setIcon(QMessageBox.Information)
        btn_open = msg_box.addButton(t("exc_btn_open"), QMessageBox.ActionRole)
        msg_box.addButton(t("exc_btn_ok"), QMessageBox.AcceptRole)
        msg_box.exec()
        if msg_box.clickedButton() == btn_open:
            QDesktopServices.openUrl(QUrl.fromLocalFile(save_filename))

    def _on_table_error(self, err_msg):
        if hasattr(self, "table_progress"): self.table_progress.close()
        QMessageBox.warning(self, t("msg_error"), t("exc_err").format(err_msg))

    def _on_table_missing(self):
        if hasattr(self, "table_progress"): self.table_progress.close()
        QMessageBox.warning(self, t("exc_notable_title"), t("exc_notable_msg"))

    def add_cross_pdf_link(self, item):
        bid = item.data(0, Qt.UserRole + 1)
        if not bid: return
        pdf_options = [t("link_wiz_newfile")]
        history = getattr(self, "pdf_history", [])
        pdf_options.extend([Path(d.get("path")).name for d in history if isinstance(d, dict) and d.get("path")])
        target_file, ok1 = QInputDialog.getItem(self, t("link_wiz_title"), t("link_wiz_q1"), pdf_options, 0, False)
        if not ok1 or not target_file: return
        final_file_name = target_file
        if target_file == t("link_wiz_newfile"):
            filename, _ = QFileDialog.getOpenFileName(self, t("link_wiz_open"), QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation), "PDF Files (*.pdf)")
            if not filename: return
            final_file_name = Path(filename).name
            self.add_to_history(filename, 0)
        target_page, ok2 = QInputDialog.getInt(self, t("link_wiz_title"), t("link_wiz_q2").format(final_file_name), 1, 1, 9999, 1)
        if not ok2: return
        try:
            self.note_edit.setReadOnly(False)
            updated_text = self.note_edit.toPlainText() + f" [[{final_file_name}#{target_page}]]"
            self.db.update_bookmark_notes(bid, updated_text)
            self.note_edit.setReadOnly(True)
            html_note = updated_text
            for full_match, link_content in re.findall(r'(\[\[([^\]]+)\]\])', updated_text):
                html_link = f'<a href="{link_content}" style="color: #0078ff; font-weight: bold; text-decoration: none;">{full_match}</a>'
                html_note = html_note.replace(full_match, html_link)
            self.note_edit.setHtml(html_note.replace("\n", "<br>"))
        except Exception as e:
            QMessageBox.warning(self, t("msg_error"), str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # WAYLAND VE KDE 6 İÇİN SİHRİ GERÇEKLEŞTİREN İKİ KRİTİK MÜHÜR:
    # Görev çubuğunun icon.svg dosyasını yakalamasını sağlar.
    app.setApplicationName("PDF_Network")
    app.setDesktopFileName("PDF_Network")

    if "APPIMAGE" in os.environ:
        PROJE_ANA_DIZINI = Path(os.environ["APPIMAGE"]).parent
    else:
        PROJE_ANA_DIZINI = Path(os.path.realpath(__file__)).parent

    window = MainWindow()

    # Eğer terminalden parametre olarak bir dosya yolu gönderildiyse otomatik aç
    if len(sys.argv) > 1 and Path(sys.argv[1]).exists():
        window.open_in_tab(sys.argv[1])

    window.show()
    sys.exit(app.exec())

