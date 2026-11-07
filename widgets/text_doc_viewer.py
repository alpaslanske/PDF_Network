#!/usr/bin/env python3
from pathlib import Path
from PySide6.QtCore import Signal, Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QTextEdit, QMessageBox, QHBoxLayout, QPushButton, QFrame
)

import docx
from odf import opendocument, text, style

class TextDocViewer(QWidget):
    document_saved = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_file_path = ""
        self.file_type = ""
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        # Üst basit araç çubuğu
        toolbar = QHBoxLayout()
        self.btn_save = QPushButton("💾 Kaydet (Ctrl+S)")
        self.btn_save.clicked.connect(self.save_document)
        self.btn_save.setMaximumHeight(28)
        self.btn_save.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                padding: 4px 12px;
                background-color: #2ecc71;
                color: white;
                border-radius: 4px;
                border: none;
            }
            QPushButton:hover { background-color: #27ae60; }
        """)
        toolbar.addWidget(self.btn_save)
        toolbar.addStretch()
        layout.addLayout(toolbar)

        # Kelime işlemci arka plan çerçevesi (Gri alan)
        self.container_frame = QFrame()
        self.container_frame.setStyleSheet("background-color: #eef2f5; border: 1px solid #dcdde1; border-radius: 4px;")
        container_layout = QHBoxLayout(self.container_frame)
        container_layout.setContentsMargins(30, 20, 30, 20)
        container_layout.setAlignment(Qt.AlignHCenter)

        # Zengin metin editörü (Beyaz Sayfa)
        # NOT: Kalın/italik formatları kaybetmemek için QTextEdit'i varsayılan bırakıyoruz
        self.editor = QTextEdit()
        self.editor.setFont(QFont("Arial", 11))
        self.editor.setFixedWidth(800)
        self.editor.setStyleSheet("""
            QTextEdit {
                background-color: #ffffff;
                color: #2f3640;
                border: 1px solid #b2bec3;
                border-radius: 2px;
                padding: 45px 55px;
                line-height: 1.6;
            }
        """)

        container_layout.addWidget(self.editor)
        layout.addWidget(self.container_frame, 1)

    def open_document(self, file_path):
        self.current_file_path = file_path
        path = Path(file_path)
        self.file_type = path.suffix.lower().lstrip(".")

        try:
            if self.file_type == "docx":
                self.load_docx(path)
            elif self.file_type == "odt":
                self.load_odt(path)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Belge yüklenirken bir hata oldu:\n{e}")

    def load_docx(self, path):
        doc = docx.Document(str(path))
        html_content = "<div style='font-family: Arial, sans-serif; font-size: 14px; color: #2f3640;'>"

        for paragraph in doc.paragraphs:
            p_html = ""
            for run in paragraph.runs:
                text_clean = run.text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

                # Biçimlendirme etiketlerini inline CSS olarak topluyoruz
                styles = []
                if run.bold: styles.append("font-weight: bold;")
                if run.italic: styles.append("font-style: italic;")
                if run.underline: styles.append("text-decoration: underline;")

                if styles:
                    style_str = " ".join(styles)
                    p_html += f'<span style="{style_str}">{text_clean}</span>'
                else:
                    p_html += text_clean

            align = ""
            if paragraph.alignment == docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER: align = ' text-align: center;'
            elif paragraph.alignment == docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT: align = ' text-align: right;'

            html_content += f'<p style="margin-bottom: 12px;{align}">{p_html if p_html else "&nbsp;"}</p>'

        html_content += "</div>"

        # Kesin olarak HTML olarak yüklüyoruz
        self.editor.setHtml(html_content)

    def load_odt(self, path):
        doc = opendocument.load(str(path))

        # 1. Dökümandaki tüm stilleri tek bir güvenli havuzda toplayıp haritalandırıyoruz
        style_map = {}

        # Tüm olası stil kaynaklarını birleştir
        all_styles = []
        if doc.styles:
            all_styles.extend(doc.styles.getElementsByType(style.Style))
        if doc.automaticstyles:
            all_styles.extend(doc.automaticstyles.getElementsByType(style.Style))

        for s in all_styles:
            name = s.getAttribute("name")
            if not name: continue

            styles = []
            for text_prop in s.getElementsByType(style.TextProperties):
                fw = text_prop.getAttribute("fontweight")
                fs = text_prop.getAttribute("fontstyle")
                td = text_prop.getAttribute("textunderlinestyle")

                if fw == "bold": styles.append("font-weight: bold;")
                if fs == "italic": styles.append("font-style: italic;")
                if td == "solid" or text_prop.getAttribute("textunderlinecolor"): styles.append("text-decoration: underline;")

            if styles:
                style_map[name] = " ".join(styles)

        # 2. Metni ayrıştırıp HTML blokları inşa ediyoruz
        html_content = "<div style='font-family: Arial, sans-serif; font-size: 14px; color: #2f3640;'>"
        paragraphs = doc.getElementsByType(text.P)

        for p in paragraphs:
            p_html = ""
            for node in p.childNodes:
                # Düz metin ise direkt ekle
                if node.nodeType == node.TEXT_NODE:
                    text_clean = node.data.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                    p_html += text_clean

                # Biçimlendirilmiş öğe (Span) ise haritadan stilini bulup inline CSS uygula
                elif node.nodeType == node.ELEMENT_NODE and node.tagName == "text:span":
                    span_style_name = node.getAttribute("stylename")
                    span_text = "".join([c.data for c in node.childNodes if c.nodeType == c.TEXT_NODE])
                    span_text = span_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

                    inline_css = style_map.get(span_style_name, "")
                    if inline_css:
                        p_html += f'<span style="{inline_css}">{span_text}</span>'
                    else:
                        p_html += span_text

            html_content += f'<p style="margin-bottom: 12px;">{p_html if p_html else "&nbsp;"}</p>'

        html_content += "</div>"

        # İçeriği QTextEdit'e HTML (Zengin Metin) olarak aktar
        self.editor.setHtml(html_content)

    def save_document(self):
        if not self.current_file_path: return
        try:
            # Düzenlenen metni düz yazı olarak geri alıp kaydediyoruz
            plain_text = self.editor.toPlainText()
            lines = plain_text.split("\n")

            if self.file_type == "docx":
                doc = docx.Document()
                for line in lines: doc.add_paragraph(line)
                doc.save(self.current_file_path)
            elif self.file_type == "odt":
                doc = opendocument.OpenDocumentText()
                for line in lines: doc.text.addElement(text.P(text=line))
                doc.save(self.current_file_path)

            QMessageBox.information(self, "Başarılı", "Değişiklikler başarıyla dosyaya kaydedildi.")
            self.document_saved.emit(self.current_file_path)
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dosya kaydedilirken hata oluştu:\n{e}")
