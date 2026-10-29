import pymupdf
import pytesseract
from PIL import Image

from PySide6.QtCore import Signal, Qt, QRect, QUrl, QMimeData
from PySide6.QtGui import (
    QPixmap,
    QImage,
    QPainter,
    QColor,
    QPen,
    QShortcut,
    QKeySequence,
    QDesktopServices,
)
from PySide6.QtWidgets import (
    QLabel,
    QScrollArea,
    QApplication,
    QMenu,
    QRubberBand
)
from lang import t

class PDFViewer(QScrollArea):
    zoom_changed = Signal()
    bookmarkRequested = Signal(int, str, float, float)
    highlightRequested = Signal(int, float, float, float, float)
    deleteHighlightRequested = Signal(int, float, float)
    page_changed = Signal(int)
    exportTableRequested = Signal(int, float, float, float, float)

    def __init__(self):
        super().__init__()
        self.doc = None
        self.current_page = 0
        self.zoom = 1.5
        self.current_file_path = ""

        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.mousePressEvent = self.mousePressEvent
        self.image_label.mouseMoveEvent = self.mouseMoveEvent
        self.image_label.mouseReleaseEvent = self.mouseReleaseEvent

        self.image_label.setMouseTracking(True)
        self.setMouseTracking(True)

        self.setWidget(self.image_label)
        self.setWidgetResizable(False)
        self.setAlignment(Qt.AlignCenter)

        self.rubberBand = QRubberBand(QRubberBand.Rectangle, self.image_label)

        self.highlight_rects = []
        self.selecting = False
        self.selection_start = None
        self.selection_end = None
        self.selection_page = None

        self.pan_mode = False
        self.pan_start_pos = None
        self.dragged_enough = False
        self.active_search_rect = None

        self.copy_shortcut = QShortcut(QKeySequence.Copy, self)
        self.copy_shortcut.activated.connect(self.copy_selected_text)

    def open_pdf(self, pdf_path):
        self.current_file_path = pdf_path
        self.doc = pymupdf.open(pdf_path)
        self.current_page = 0
        self.highlight_rects = []
        self.render_page()

    def render_page(self):
        if not self.doc:
            return

        page = self.doc[self.current_page]
        matrix = pymupdf.Matrix(self.zoom, self.zoom)
        pix = page.get_pixmap(matrix=matrix)

        image = QImage(
            pix.samples,
            pix.width,
            pix.height,
            pix.stride,
            QImage.Format_RGB888,
        )

        painter = QPainter(image)

        if self.highlight_rects:
            for page_no, rect in self.highlight_rects:
                if page_no != self.current_page:
                    continue

                x = int(rect.x0 * self.zoom)
                y = int(rect.y0 * self.zoom)
                w = int((rect.x1 - rect.x0) * self.zoom)
                h = int((rect.y1 - rect.y0) * self.zoom)

                is_active = (
                    self.active_search_rect is not None
                    and abs(rect.x0 - self.active_search_rect.x0) < 0.1
                    and abs(rect.y0 - self.active_search_rect.y0) < 0.1
                )

                if is_active:
                    painter.fillRect(x, y, w, h, QColor(0, 255, 150, 120))
                    painter.setPen(QPen(QColor(0, 50, 255), 3))
                else:
                    painter.fillRect(x, y, w, h, QColor(255, 255, 0, 80))
                    painter.setPen(QPen(QColor(255, 0, 0), 1.5))

                painter.drawRect(x, y, w, h)

        if (
            self.selection_page == self.current_page
            and self.selection_start is not None
            and self.selection_end is not None
        ):
            rect = QRect(self.selection_start, self.selection_end).normalized()
            painter.fillRect(rect, QColor(0, 120, 255, 60))
            painter.setPen(QPen(QColor(0, 120, 255), 2))
            painter.drawRect(rect)

        painter.end()
        self.image_label.setPixmap(QPixmap.fromImage(image))
        self.image_label.resize(pix.width, pix.height)
        pix = None

    def set_highlight(self, rects):
        self.highlight_rects = rects
        self.render_page()

    def next_page(self):
        if self.doc and self.current_page < len(self.doc) - 1:
            self.highlight_rects = []
            self.active_search_rect = None
            self.current_page += 1
            self.render_page()
            self.page_changed.emit(self.current_page)

    def previous_page(self):
        if self.doc and self.current_page > 0:
            self.highlight_rects = []
            self.active_search_rect = None
            self.current_page -= 1
            self.render_page()
            self.page_changed.emit(self.current_page)

    def goto_page(self, page):
        self.selection_start = None
        self.selection_end = None
        self.highlight_rects = []
        self.active_search_rect = None

        if not self.doc or page < 0 or page >= len(self.doc):
            return

        self.current_page = page
        self.render_page()
        self.image_label.repaint()
        self.page_changed.emit(self.current_page)

    def goto_position(self, page, x, y):
        self.goto_page(page)
        QApplication.processEvents()
        self.horizontalScrollBar().setValue(int(x * self.zoom))
        self.verticalScrollBar().setValue(int(y * self.zoom))

    def search_text(self, text):
        if not self.doc:
            return []
        results = []
        for page_no in range(len(self.doc)):
            page = self.doc[page_no]
            for rect in page.search_for(text):
                results.append((page_no, rect))
        return results

    def get_pdf_point(self, event):
        local_pos = self.image_label.mapFromGlobal(event.globalPosition().toPoint())
        if local_pos.x() < 0 or local_pos.y() < 0:
            return None
        return pymupdf.Point(local_pos.x() / self.zoom, local_pos.y() / self.zoom)

    def mousePressEvent(self, event):
        if event.button() == Qt.RightButton:
            self.pan_mode = True
            self.dragged_enough = False
            self.pan_start_pos = event.globalPosition().toPoint()
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()
            return

        if event.button() == Qt.LeftButton:
            self.selecting = False
            self.mouse_press_pos = event.position().toPoint()
            self.is_link_click = False

            pdf_point = self.get_pdf_point(event)
            if pdf_point and self.check_and_follow_link(pdf_point):
                self.is_link_click = True
                self.selection_start = None
                self.selection_end = None
                event.accept()
                return

            self.rubberBand.hide()
            event.accept()

    def mouseMoveEvent(self, event):
        if self.pan_mode and self.pan_start_pos is not None:
            current_pos = event.globalPosition().toPoint()
            delta = self.pan_start_pos - current_pos
            if delta.manhattanLength() > 3:
                self.dragged_enough = True

            self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() + delta.x())
            self.verticalScrollBar().setValue(self.verticalScrollBar().value() + delta.y())
            self.pan_start_pos = current_pos
            event.accept()
            return

        if event.buttons() & Qt.LeftButton:
            if getattr(self, "is_link_click", False):
                event.accept()
                return

            current_pos = event.position().toPoint()
            press_pos = getattr(self, "mouse_press_pos", current_pos)

            if not self.selecting and (current_pos - press_pos).manhattanLength() > 5:
                self.selecting = True
                self.selection_start = press_pos
                self.rubberBand.setGeometry(QRect(self.selection_start, self.selection_start).normalized())
                self.rubberBand.show()

            if self.selecting:
                self.selection_end = current_pos
                self.rubberBand.setGeometry(QRect(self.selection_start, self.selection_end).normalized())
                event.accept()
                return

        if self.doc and 0 <= self.current_page < len(self.doc):
            pdf_point = self.get_pdf_point(event)
            if pdf_point:
                page = self.doc[self.current_page]
                is_over_link = any(
                    link.get("from") and link.get("from").contains(pdf_point)
                    for link in page.get_links()
                )
                self.setCursor(Qt.PointingHandCursor if is_over_link else Qt.ArrowCursor)
            else:
                self.setCursor(Qt.ArrowCursor)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.RightButton and self.pan_mode:
            self.pan_mode = False
            self.pan_start_pos = None
            self.setCursor(Qt.ArrowCursor)
            if not self.dragged_enough:
                self.show_context_menu(event.position().toPoint())
            event.accept()
            return

        if event.button() == Qt.LeftButton:
            if getattr(self, "is_link_click", False):
                self.is_link_click = False
                self.selecting = False
                event.accept()
                return

            if self.selecting:
                self.selecting = False
                self.selection_page = self.current_page
                self.selection_end = event.position().toPoint()
                self.rubberBand.hide()
                self.render_page()
            else:
                self.selection_start = None
                self.selection_end = None
                self.rubberBand.hide()
                self.render_page()

            event.accept()

    def check_and_follow_link(self, point):
        if not self.doc or not (0 <= self.current_page < len(self.doc)):
            return False

        page = self.doc[self.current_page]
        links = page.get_links()

        for link in links:
            rect = link.get("from")
            if not rect or not rect.contains(point):
                continue

            uri = link.get("uri")
            if uri:
                QDesktopServices.openUrl(QUrl(uri))
                return True

            name_action = link.get("name")
            if name_action:
                if name_action == "NextPage":
                    self.next_page()
                    return True
                elif name_action == "PrevPage":
                    self.previous_page()
                    return True
                elif name_action == "FirstPage":
                    self.goto_page(0)
                    return True
                elif name_action == "LastPage":
                    self.goto_page(len(self.doc) - 1)
                    return True

            file_spec = link.get("file")
            if file_spec:
                return True

            target_page = link.get("page")
            to_point = link.get("to")

            if target_page is None or target_page < 0:
                named_dest = link.get("nameddest") or link.get("dest")
                if named_dest:
                    try:
                        names_dict = self.doc.resolve_names()
                        if named_dest in names_dict:
                            target_info = names_dict[named_dest]
                            target_page = target_info[0]
                    except Exception:
                        pass

                if target_page is None and named_dest:
                    try:
                        resolved = self.doc.resolve_link(link)
                        if isinstance(resolved, (tuple, list)) and len(resolved) >= 1:
                            target_page = resolved[0]
                            if len(resolved) > 1 and resolved[1] is not None:
                                to_point = resolved[1]
                    except Exception:
                        pass

            if target_page is not None and 0 <= target_page < len(self.doc):
                if target_page != self.current_page:
                    self.goto_page(target_page)
                    QApplication.processEvents()

                if to_point:
                    try:
                        px = getattr(to_point, "x", to_point[0] if isinstance(to_point, (list, tuple)) else 0)
                        py = getattr(to_point, "y", to_point[1] if isinstance(to_point, (list, tuple)) else 0)

                        target_y = int(py * self.zoom)
                        target_x = int(px * self.zoom)

                        v_bar = self.verticalScrollBar()
                        h_bar = self.horizontalScrollBar()

                        v_bar.setValue(min(target_y, v_bar.maximum()))
                        h_bar.setValue(min(target_x, h_bar.maximum()))

                    except Exception as e:
                        print(f"[Kaydırma Hatası]: {e}")
                else:
                    if target_page != self.current_page:
                        self.verticalScrollBar().setValue(0)
                        self.horizontalScrollBar().setValue(0)

                return True

        return False

    def zoom_in(self):
        self.zoom += 0.2
        self.render_page()

    def zoom_out(self):
        self.zoom = max(0.4, self.zoom - 0.2)
        self.render_page()

    def get_selected_text(self):
        if self.selection_start is None or self.selection_end is None or not self.doc:
            return ""

        rect = pymupdf.Rect(
            min(self.selection_start.x(), self.selection_end.x()) / self.zoom,
            min(self.selection_start.y(), self.selection_end.y()) / self.zoom,
            max(self.selection_start.x(), self.selection_end.x()) / self.zoom,
            max(self.selection_start.y(), self.selection_end.y()) / self.zoom,
        )
        return self.doc[self.current_page].get_text("text", clip=rect).strip()

    def copy_selected_text(self, force_ocr=False):
        if self.selection_start is None or self.selection_end is None or not self.doc:
            return

        rect = pymupdf.Rect(
            min(self.selection_start.x(), self.selection_end.x()) / self.zoom,
            min(self.selection_start.y(), self.selection_end.y()) / self.zoom,
            max(self.selection_start.x(), self.selection_end.x()) / self.zoom,
            max(self.selection_start.y(), self.selection_end.y()) / self.zoom,
        )

        page = self.doc[self.current_page]
        plain_text = page.get_text("text", clip=rect, sort=True).strip()
        clean_html = ""
        is_ocr_used = False

        if force_ocr or len(plain_text) < 3:
            try:
                matrix = pymupdf.Matrix(3, 3)
                pix = page.get_pixmap(matrix=matrix, clip=rect)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                ocr_text = pytesseract.image_to_string(img, lang="tur+pol+eng", config="--psm 6")

                if ocr_text.strip():
                    plain_text = ocr_text.strip()
                    is_ocr_used = True
                pix = None
            except Exception as e:
                print(f"[HATA] OCR Hatası: {e}")

        if not is_ocr_used:
            text_dict = page.get_text("dict", clip=rect, sort=True)
            html_blocks = []
            for block in text_dict.get("blocks", []):
                if block.get("type") == 0:
                    lines = block.get("lines", [])
                    if not lines:
                        continue

                    max_x1 = max(line.get("bbox", (0,0,0,0))[2] for line in lines)
                    assembled_html = ""
                    for i, line in enumerate(lines):
                        line_html = ""
                        for span in line.get("spans", []):
                            text = span.get("text", "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                            color_int = span.get("color", 0)
                            hex_color = f"#{color_int:06x}" if isinstance(color_int, int) else "#000000"
                            flags = span.get("flags", 0)

                            style = f"color: {hex_color};"
                            if bool(flags & 16):
                                style += " font-weight: bold;"
                            if bool(flags & 2):
                                style += " font-style: italic;"

                            line_html += f'<span style="{style}">{text}</span>'

                        if not line_html:
                            continue

                        if i == 0:
                            assembled_html = line_html
                        else:
                            prev_line = lines[i-1]
                            prev_x1 = prev_line.get("bbox", (0,0,0,0))[2]

                            if max_x1 - prev_x1 < 30:
                                assembled_html += " " + line_html
                            else:
                                assembled_html += f"</p><p style='margin: 0; text-align: left;'>{line_html}"

                    if assembled_html:
                        html_blocks.append(f"<p style='margin: 0 0 8px 0; text-align: left;'>{assembled_html}</p>")

            clean_html = "\n".join(html_blocks)

        mime_data = QMimeData()

        if is_ocr_used:
            normalized_ocr = plain_text.replace('\r\n', '\n')
            paragraphs = normalized_ocr.split('\n\n')

            html_lines = []
            plain_lines = []
            for p in paragraphs:
                lines = [l.strip() for l in p.split('\n') if l.strip()]
                if not lines:
                    continue

                max_len = max(len(l) for l in lines)
                assembled_p = lines[0]
                assembled_plain = lines[0]

                for i in range(1, len(lines)):
                    prev_len = len(lines[i-1])

                    if max_len > 40 and (max_len - prev_len) < 15:
                        assembled_p += " " + lines[i]
                        assembled_plain += " " + lines[i]
                    else:
                        assembled_p += f"</p><p style='margin: 0; text-align: left;'>{lines[i]}"
                        assembled_plain += "\r\n" + lines[i]

                html_lines.append(f"<p style='margin: 0 0 8px 0; text-align: left;'>{assembled_p}</p>")
                plain_lines.append(assembled_plain)

            html_formatted = "\n".join(html_lines)
            full_html = f"<html><body><div style='font-family: sans-serif;'>{html_formatted}</div></body></html>"
            safe_plain_text = "\r\n\r\n".join(plain_lines)
            mime_data.setText(safe_plain_text)
        else:
            safe_plain_text = plain_text.replace('\n', ' \r\n').replace('  \r\n', ' \r\n')
            mime_data.setText(safe_plain_text)
            full_html = f"<html><body><div style='text-align: left;'>{clean_html}</div></body></html>"

        mime_data.setHtml(full_html)
        QApplication.clipboard().setMimeData(mime_data)

        self.selection_start = None
        self.selection_end = None
        self.render_page()

    def create_bookmark_from_selection(self):
        text = self.get_selected_text()
        if not text:
            return
        title = " ".join(text.split())
        if len(title) > 60:
            title = title[:60] + "..."
        pos = self.get_selection_position()
        if pos is None:
            return
        self.bookmarkRequested.emit(self.current_page, title, pos[0], pos[1])

    def create_highlight(self):
        if not getattr(self, "current_file_path", "").lower().endswith(".pdf"):
            return
        if self.selection_start is None or self.selection_end is None:
            return
        rect = pymupdf.Rect(
            min(self.selection_start.x(), self.selection_end.x()) / self.zoom,
            min(self.selection_start.y(), self.selection_end.y()) / self.zoom,
            max(self.selection_start.x(), self.selection_end.x()) / self.zoom,
            max(self.selection_start.y(), self.selection_end.y()) / self.zoom,
        )
        self.highlightRequested.emit(self.current_page, rect.x0, rect.y0, rect.x1, rect.y1)

    def delete_highlight(self):
        if self.selection_start is None or self.selection_end is None:
            return
        x = min(self.selection_start.x(), self.selection_end.x()) / self.zoom
        y = min(self.selection_start.y(), self.selection_end.y()) / self.zoom
        self.deleteHighlightRequested.emit(self.current_page, x, y)

    def export_table_from_selection(self):
        if self.selection_start is None or self.selection_end is None:
            return

        x0 = min(self.selection_start.x(), self.selection_end.x()) / self.zoom
        y0 = min(self.selection_start.y(), self.selection_end.y()) / self.zoom
        x1 = max(self.selection_start.x(), self.selection_end.x()) / self.zoom
        y1 = max(self.selection_start.y(), self.selection_end.y()) / self.zoom

        self.exportTableRequested.emit(self.current_page, x0, y0, x1, y1)

    def get_selection_position(self):
        if self.selection_start is None or self.selection_end is None:
            return None
        return (
            min(self.selection_start.x(), self.selection_end.x()) / self.zoom,
            min(self.selection_start.y(), self.selection_end.y()) / self.zoom,
        )

    def show_context_menu(self, pos):
        menu = QMenu(self)
        copy_action = menu.addAction(t("view_copy"))
        copy_ocr_action = menu.addAction(t("view_copy_ocr"))
        bookmark_action = menu.addAction(t("view_add_bookmark"))
        highlight_action = menu.addAction(t("view_add_highlight"))
        delete_highlight_action = menu.addAction(t("view_del_highlight"))
        export_table_action = menu.addAction(t("view_export_excel"))

        is_pdf = getattr(self, "current_file_path", "").lower().endswith(".pdf")
        highlight_action.setEnabled(is_pdf)
        delete_highlight_action.setEnabled(is_pdf)

        selected = menu.exec(self.mapToGlobal(pos))
        if selected == copy_action:
            self.copy_selected_text(force_ocr=False)
        elif selected == copy_ocr_action:
            self.copy_selected_text(force_ocr=True)
        elif selected == bookmark_action:
            self.create_bookmark_from_selection()
        elif selected == highlight_action:
            self.create_highlight()
        elif selected == delete_highlight_action:
            self.delete_highlight()
        elif selected == export_table_action:
            self.export_table_from_selection()

    def wheelEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            if event.angleDelta().y() > 0:
                self.zoom_in()
            else:
                self.zoom_out()
            self.zoom_changed.emit()
            event.accept()
            return

        v_scrollbar = self.verticalScrollBar()
        delta = event.angleDelta().y()

        if delta < 0:
            if v_scrollbar.value() < v_scrollbar.maximum():
                super().wheelEvent(event)
            else:
                if self.doc and self.current_page < len(self.doc) - 1:
                    self.next_page()
                    v_scrollbar.setValue(0)
                event.accept()
        elif delta > 0:
            if v_scrollbar.value() > 0:
                super().wheelEvent(event)
            else:
                if self.current_page > 0:
                    self.previous_page()
                    v_scrollbar.setValue(v_scrollbar.maximum())
                event.accept()
