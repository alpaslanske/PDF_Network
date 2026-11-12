#!/usr/bin/env python3
import datetime
from pathlib import Path
from PySide6.QtCore import QUrl, QStandardPaths
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QFileDialog, QMessageBox, QProgressDialog
from lang import t
from widgets.pdf_workers import TableExportWorker


class ExportManager:
    def __init__(self, main_window):
        self.mw = main_window

    def export_bookmarks_to_markdown(self, base_dir):
        viewer = self.mw.get_current_viewer()
        if not viewer:
            QMessageBox.warning(self.mw, t("msg_warning"), t("export_warn_open"))
            return

        cpdf = getattr(viewer, "original_file_path", viewer.current_file_path)
        file_name = Path(cpdf).stem
        default_save = base_dir / f"{file_name}_Calisma_Notlari.md"
        save_filename, _ = QFileDialog.getSaveFileName(
            self.mw, t("export_save_title"), str(default_save), "Markdown Files (*.md)"
        )
        if not save_filename:
            return

        try:
            bookmarks = self.mw.db.get_bookmarks()
            current_bookmarks = [b for b in bookmarks if b[2] == cpdf]
            if not current_bookmarks:
                QMessageBox.information(self.mw, t("msg_info"), t("export_no_bookmarks"))
                return

            groups = {g[0]: g[1] for g in self.mw.db.get_groups()}
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
            QMessageBox.information(self.mw, t("msg_success"), t("export_success").format(Path(save_filename).name))
        except Exception as e:
            QMessageBox.warning(self.mw, t("msg_error"), t("export_err").format(e))

    def export_table_to_excel(self, viewer, page_no, x0, y0, x1, y1):
        if not viewer or not viewer.doc:
            return
        try:
            import pandas as pd
        except ImportError:
            QMessageBox.warning(self.mw, t("exc_lib_err"), t("exc_lib_msg"))
            return

        docs_dir = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
        default_save = str(Path(docs_dir) / f"{Path(viewer.current_file_path).stem}_Tablo_Sayfa{page_no + 1}.xlsx")
        save_filename, _ = QFileDialog.getSaveFileName(self.mw, t("exc_save_title"), default_save, "Excel Files (*.xlsx)")
        if not save_filename:
            return

        self.table_progress = QProgressDialog(t("prog_table_wait"), None, 0, 0, self.mw)
        self.table_progress.setWindowTitle(t("exc_save_title"))
        self.table_progress.setWindowModality(self.mw.windowModality())
        self.table_progress.show()

        self.table_worker = TableExportWorker(viewer.current_file_path, page_no, x0, y0, x1, y1, save_filename)
        self.table_worker.finished.connect(self._on_table_finished)
        self.table_worker.error.connect(self._on_table_error)
        self.table_worker.no_table.connect(self._on_table_missing)
        self.table_worker.start()

    def _on_table_finished(self, save_filename):
        if hasattr(self, "table_progress"):
            self.table_progress.close()
        msg_box = QMessageBox(self.mw)
        msg_box.setWindowTitle(t("msg_success"))
        msg_box.setText(t("exc_success_msg").format(Path(save_filename).name))
        msg_box.setIcon(QMessageBox.Information)
        btn_open = msg_box.addButton(t("exc_btn_open"), QMessageBox.ActionRole)
        msg_box.addButton(t("exc_btn_ok"), QMessageBox.AcceptRole)
        msg_box.exec()
        if msg_box.clickedButton() == btn_open:
            QDesktopServices.openUrl(QUrl.fromLocalFile(save_filename))

    def _on_table_error(self, err_msg):
        if hasattr(self, "table_progress"):
            self.table_progress.close()
        QMessageBox.warning(self.mw, t("msg_error"), t("exc_err").format(err_msg))

    def _on_table_missing(self):
        if hasattr(self, "table_progress"):
            self.table_progress.close()
        QMessageBox.warning(self.mw, t("exc_notable_title"), t("exc_notable_msg"))
