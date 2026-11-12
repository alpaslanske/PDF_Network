#!/usr/bin/env python3
from PySide6.QtCore import QThread, Signal
import pymupdf
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from PySide6.QtGui import QImage, QPainter
from PySide6.QtPrintSupport import QPrinter


class PDFSearchWorker(QThread):
    finished = Signal(list)
    error = Signal(str)

    def __init__(self, file_path, search_text):
        super().__init__()
        self.file_path = file_path
        self.search_text = search_text

    def run(self):
        try:
            doc = pymupdf.open(self.file_path)
            results = []
            for page_no in range(len(doc)):
                if self.isInterruptionRequested():
                    break
                page = doc[page_no]
                for rect in page.search_for(self.search_text):
                    results.append((page_no, rect))
            doc.close()
            self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))


class TableExportWorker(QThread):
    finished = Signal(str)
    error = Signal(str)
    no_table = Signal()

    def __init__(self, file_path, page_no, x0, y0, x1, y1, save_filename):
        super().__init__()
        self.file_path = file_path
        self.page_no = page_no
        self.x0 = x0
        self.y0 = y0
        self.x1 = x1
        self.y1 = y1
        self.save_filename = save_filename

    def run(self):
        try:
            doc = pymupdf.open(self.file_path)
            page = doc[self.page_no]
            rect = pymupdf.Rect(self.x0, self.y0, self.x1, self.y1)
            tables = page.find_tables(clip=rect)

            if not tables or not tables.tables:
                self.no_table.emit()
                doc.close()
                return

            data = tables.tables[0].extract()
            if not data:
                self.no_table.emit()
                doc.close()
                return

            df = pd.DataFrame(data).fillna('').astype(str).replace(['NaN', 'None'], '', regex=True)
            cleaned_rows = df.apply(
                lambda row: [str(val).replace('\n', ' ').strip() for val in row if str(val).strip() != ''],
                axis=1
            ).tolist()

            if not cleaned_rows:
                self.no_table.emit()
                doc.close()
                return

            max_cols = max(len(row) for row in cleaned_rows if row)
            final_data = []
            for i, row in enumerate(cleaned_rows):
                if not row:
                    final_data.append([""] * max_cols)
                    continue
                if len(row) < max_cols:
                    if i == 0:
                        row = [""] * (max_cols - len(row)) + row
                    else:
                        row = row + [""] * (max_cols - len(row))
                final_data.append(row)

            df_clean = pd.DataFrame(final_data)
            df_clean.to_excel(self.save_filename, index=False, header=False)

            try:
                wb = load_workbook(self.save_filename)
                ws = wb.active
                serif_font = Font(name="Times New Roman", size=11)
                for row in ws.iter_rows():
                    for cell in row:
                        cell.font = serif_font
                wb.save(self.save_filename)
            except Exception:
                pass

            doc.close()
            self.finished.emit(self.save_filename)
        except Exception as e:
            self.error.emit(str(e))


class PDFPrintWorker(QThread):
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, file_path, printer, from_page, to_page):
        super().__init__()
        self.file_path = file_path
        self.printer = printer
        self.from_page = from_page
        self.to_page = to_page

    def run(self):
        try:
            doc = pymupdf.open(self.file_path)
            painter = QPainter(self.printer)

            for i in range(self.from_page, self.to_page + 1):
                if i > self.from_page:
                    self.printer.newPage()
                page = doc[i]
                matrix = pymupdf.Matrix(3, 3)
                pix = page.get_pixmap(matrix=matrix)
                image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)

                rect = self.printer.pageRect(QPrinter.DevicePixel)
                painter.drawImage(rect, image)

            painter.end()
            doc.close()
            self.finished.emit(self.printer.printerName())
        except Exception as e:
            self.error.emit(str(e))
