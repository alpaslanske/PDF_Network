#!/usr/bin/env python3
from PySide6.QtCore import QThread, Signal
import pymupdf
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from PySide6.QtGui import QImage, QPainter
from PySide6.QtPrintSupport import QPrinter
from pathlib import Path

class TableExportWorker(QThread):
    finished = Signal(str)  # Başarılı ise kaydedilen dosya yolunu döndürür
    error = Signal(str)     # Hata oluşursa mesajı döndürür
    no_table = Signal()     # Tablo bulunamazsa tetiklenir

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
            cleaned_rows = df.apply(lambda row: [str(val).replace('\n', ' ').strip() for val in row if str(val).strip() != ''], axis=1).tolist()

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
    finished = Signal(str)  # Yazıcı adını döndürür
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


from PySide6.QtCore import QRunnable, QObject, Signal
import pymupdf
from pathlib import Path

class ThumbnailSignals(QObject):
    finished_data = Signal(bytes, int, int, int, str, object)

class ThumbnailWorker(QRunnable):
    """Kapak resimlerini global havuzda (QThreadPool) sıfır çökme riskiyle üreten akıllı işçi."""
    def __init__(self, file_path, card_ref):
        super().__init__()
        self.file_path = file_path
        self.card_ref = card_ref
        self.signals = ThumbnailSignals()

    def run(self):
        import os
        import pymupdf
        import tempfile
        from pathlib import Path

        if not self.file_path:
            self.signals.finished_data.emit(b"", 0, 0, 0, "error", self.card_ref)
            return

        path = Path(self.file_path)
        suffix = path.suffix.lower()

        try:
            hedef_okuma_yolu = str(path)

            # EĞER HEDEF BİR EXCEL VEYA WORD BELGESİYSE:
            # Silinen geçici klasör yerine, ana uygulamanın asla silmediği kararlı taze önbellek PDF'ini buluyoruz!
            if suffix in [".docx", ".odt", ".xlsx", ".ods"]:
                try:
                    file_stat = path.stat()
                    # Ana uygulamanın kaydederken kullandığı benzersiz ID algoritmasının birebir aynısı
                    unique_id = f"{path.stem}_{file_stat.st_size}_{int(file_stat.st_mtime)}"

                    # Ortak ve kalıcı olan PDF konumu
                    final_cache_pdf = Path(tempfile.gettempdir()) / "PDF_Network_Cache" / f"{unique_id}.pdf"

                    if final_cache_pdf.exists() and final_cache_pdf.stat().st_size > 0:
                        hedef_okuma_yolu = str(final_cache_pdf)
                    else:
                        # Eğer henüz dönüştürme bitmediyse veya dosya yoksa yükleniyor durumunda beklet
                        self.signals.finished_data.emit(b"", 0, 0, 0, "office_doc", self.card_ref)
                        return
                except:
                    self.signals.finished_data.emit(b"", 0, 0, 0, "error", self.card_ref)
                    return

            # Artık elimizde kesinlikle fiziksel olarak var olan bir PDF/Kitap dökümanı var!
            doc = pymupdf.open(hedef_okuma_yolu)
            if doc and len(doc) > 0:
                page = doc[0] # İlk sayfayı kapak resmi (thumbnail) olarak alıyoruz

                # Kapak kalitesini pikselleşmeden net almak için hafif bir matris ölçeği veriyoruz
                pix = page.get_pixmap(matrix=pymupdf.Matrix(0.35, 0.35))

                # Ham bayt verilerini güvenle ana thread'e (RecentFileCard'a) fırlatıyoruz
                self.signals.finished_data.emit(pix.samples, pix.width, pix.height, pix.stride, "success", self.card_ref)
                doc.close()
            else:
                self.signals.finished_data.emit(b"", 0, 0, 0, "empty", self.card_ref)

        except Exception as e:
            print(f"Thumbnail İşleyici Hatası: {e}")
            self.signals.finished_data.emit(b"", 0, 0, 0, "error", self.card_ref)

