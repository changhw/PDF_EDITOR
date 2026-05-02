import sys
from dataclasses import dataclass
from typing import Optional

import math

import fitz
from PySide6.QtCore import QPoint, QRect, Qt, Signal
from PySide6.QtGui import QAction, QColor, QImage, QMouseEvent, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QColorDialog,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QSplitter,
    QTabWidget,
    QTextEdit,
    QToolBar,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)


PAGE_ROLE = Qt.ItemDataRole.UserRole + 1


def qcolor_to_fitz(color: QColor) -> tuple[float, float, float]:
    return (color.redF(), color.greenF(), color.blueF())


def fitz_color_to_qcolor(color_value: object) -> QColor:
    if isinstance(color_value, int):
        return QColor((color_value >> 16) & 0xFF, (color_value >> 8) & 0xFF, color_value & 0xFF)
    if isinstance(color_value, (list, tuple)) and len(color_value) >= 3:
        red, green, blue = color_value[:3]
        if max(red, green, blue) <= 1:
            return QColor.fromRgbF(float(red), float(green), float(blue))
        return QColor(int(red), int(green), int(blue))
    return QColor("#111827")


class ColorButton(QPushButton):
    color_changed = Signal(QColor)

    def __init__(self, initial: QColor, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._color = initial
        self.clicked.connect(self.choose_color)
        self._update_style()

    @property
    def color(self) -> QColor:
        return self._color

    def set_color(self, color: QColor) -> None:
        self._color = color
        self._update_style()

    def choose_color(self) -> None:
        color = QColorDialog.getColor(self._color, self.window(), "Choose color")
        if color.isValid():
            self._color = color
            self._update_style()
            self.color_changed.emit(color)

    def _update_style(self) -> None:
        self.setText(self._color.name())
        self.setStyleSheet(
            f"background-color: {self._color.name()};"
            f"color: {'#000000' if self._color.lightness() > 128 else '#ffffff'};"
        )


class SymbolDialog(QDialog):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Add symbol")
        self.symbol_input = QLineEdit("✓")
        self.font_size = QSpinBox()
        self.font_size.setRange(6, 144)
        self.font_size.setValue(24)
        self.font_combo = QComboBox()
        self.font_combo.setEditable(False)
        self._fonts = [
            ("Built-in (Helvetica)", "helv", None),
            ("Built-in (Times)", "tiro", None),
            ("Built-in (Courier)", "cour", None),
        ]
        for label, _name, _path in self._fonts:
            self.font_combo.addItem(label)
        self.color_button = ColorButton(QColor("#d97706"))

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        form = QFormLayout(self)
        form.addRow("Symbol", self.symbol_input)
        form.addRow("Font", self.font_combo)
        form.addRow("Font size", self.font_size)
        form.addRow("Color", self.color_button)
        form.addRow(buttons)

    def get_data(self) -> Optional[dict]:
        if self.exec() != QDialog.DialogCode.Accepted:
            return None
        symbol = self.symbol_input.text().strip()
        if not symbol:
            QMessageBox.warning(self, "Missing symbol", "Enter at least one symbol character.")
            return None
        _label, font_name, font_file = self._fonts[self.font_combo.currentIndex()]
        return {
            "symbol": symbol,
            "font_size": self.font_size.value(),
            "font_name": font_name,
            "font_file": font_file,
            "color": self.color_button.color,
        }


class TextEditDialog(QDialog):
    def __init__(self, title: str, parent: Optional[QWidget] = None, *, initial_text: str = "") -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.text_edit = QTextEdit()
        self.text_edit.setPlainText(initial_text)
        self.font_size = QSpinBox()
        self.font_size.setRange(6, 144)
        self.font_size.setValue(12)
        self.alignment = QComboBox()
        self.alignment.addItems(["Left", "Center", "Right", "Justify"])
        self.color_button = ColorButton(QColor("#111827"))

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        form = QFormLayout()
        form.addRow("Text", self.text_edit)
        form.addRow("Font size", self.font_size)
        form.addRow("Alignment", self.alignment)
        form.addRow("Color", self.color_button)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)

    def get_data(self) -> Optional[dict]:
        if self.exec() != QDialog.DialogCode.Accepted:
            return None
        text = self.text_edit.toPlainText()
        if not text.strip():
            QMessageBox.warning(self, "Missing text", "Enter some text to place in the selected region.")
            return None
        alignments = {
            "Left": fitz.TEXT_ALIGN_LEFT,
            "Center": fitz.TEXT_ALIGN_CENTER,
            "Right": fitz.TEXT_ALIGN_RIGHT,
            "Justify": fitz.TEXT_ALIGN_JUSTIFY,
        }
        return {
            "text": text,
            "font_size": self.font_size.value(),
            "alignment": alignments[self.alignment.currentText()],
            "color": self.color_button.color,
        }


class ShapeDialog(QDialog):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Add figure/object")
        self.shape_type = QComboBox()
        self.shape_type.addItems(["Rectangle", "Ellipse", "Arrow"])
        self.stroke_button = ColorButton(QColor("#2563eb"))
        self.fill_button = ColorButton(QColor("#93c5fd"))
        self.line_width = QSpinBox()
        self.line_width.setRange(1, 24)
        self.line_width.setValue(2)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        form = QFormLayout(self)
        form.addRow("Shape", self.shape_type)
        form.addRow("Border color", self.stroke_button)
        form.addRow("Fill color", self.fill_button)
        form.addRow("Border width", self.line_width)
        form.addRow(buttons)

    def get_data(self) -> Optional[dict]:
        if self.exec() != QDialog.DialogCode.Accepted:
            return None
        return {
            "shape": self.shape_type.currentText(),
            "stroke": self.stroke_button.color,
            "fill": self.fill_button.color,
            "width": self.line_width.value(),
        }


class BookmarkDialog(QDialog):
    def __init__(self, parent: Optional[QWidget] = None, *, title: str = "", page_number: int = 1) -> None:
        super().__init__(parent)
        self.setWindowTitle("Bookmark")
        self.title_input = QLineEdit(title)
        self.page_input = QSpinBox()
        self.page_input.setRange(1, 99999)
        self.page_input.setValue(page_number)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        form = QFormLayout(self)
        form.addRow("Title", self.title_input)
        form.addRow("Page", self.page_input)
        form.addRow(buttons)

    def get_data(self) -> Optional[dict]:
        if self.exec() != QDialog.DialogCode.Accepted:
            return None
        title = self.title_input.text().strip()
        if not title:
            QMessageBox.warning(self, "Missing title", "Enter a bookmark title.")
            return None
        return {"title": title, "page_number": self.page_input.value()}


@dataclass
class PendingAction:
    name: str
    payload: Optional[dict] = None


@dataclass
class TextParagraph:
    rect: fitz.Rect
    text: str
    font_size: float
    color: QColor
    alignment: int


class PageView(QLabel):
    point_selected = Signal(QPoint)
    point_hovered = Signal(QPoint)
    rect_selected = Signal(QRect)

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setMouseTracking(True)
        self._mode: Optional[str] = None
        self._drag_start: Optional[QPoint] = None
        self._current_rect: Optional[QRect] = None
        self._overlay_rect: Optional[QRect] = None

    def set_mode(self, mode: Optional[str]) -> None:
        self._mode = mode
        self._drag_start = None
        self._current_rect = None
        self._overlay_rect = None
        self.setCursor(Qt.CursorShape.CrossCursor if mode else Qt.CursorShape.ArrowCursor)
        self.update()

    def set_overlay_rect(self, rect: Optional[QRect]) -> None:
        self._overlay_rect = rect.normalized() if rect is not None else None
        self.update()

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if not self.pixmap():
            return
        if self._mode == "point" and event.button() == Qt.MouseButton.LeftButton:
            self.point_selected.emit(event.position().toPoint())
            return
        if self._mode == "rect" and event.button() == Qt.MouseButton.LeftButton:
            self._drag_start = event.position().toPoint()
            self._current_rect = QRect(self._drag_start, self._drag_start)
            self.update()

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        if self._mode == "point":
            self.point_hovered.emit(event.position().toPoint())
        if self._mode == "rect" and self._drag_start is not None:
            self._current_rect = QRect(self._drag_start, event.position().toPoint()).normalized()
            self.update()

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:
        if self._mode == "rect" and self._drag_start is not None:
            rect = QRect(self._drag_start, event.position().toPoint()).normalized()
            self._drag_start = None
            self._current_rect = None
            self.update()
            if rect.width() > 4 and rect.height() > 4:
                self.rect_selected.emit(rect)

    def leaveEvent(self, event) -> None:
        self.point_hovered.emit(QPoint(-1, -1))
        super().leaveEvent(event)

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        painter = QPainter(self)
        if self._overlay_rect is not None:
            painter.setPen(QPen(QColor("#f59e0b"), 2, Qt.PenStyle.SolidLine))
            painter.fillRect(self._overlay_rect, QColor(245, 158, 11, 40))
            painter.drawRect(self._overlay_rect)
        if self._current_rect is not None:
            painter.setPen(QPen(QColor("#2563eb"), 2, Qt.PenStyle.DashLine))
            painter.drawRect(self._current_rect)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Windows PDF Editor")
        self.resize(1400, 900)

        self.doc: Optional[fitz.Document] = None
        self.current_path: Optional[str] = None
        self.current_page_index = 0
        self.zoom = 1.2
        self.dirty = False
        self.annotations_dirty = True
        self.pending_action: Optional[PendingAction] = None
        self.hovered_paragraph: Optional[TextParagraph] = None
        self.rendered_width = 1
        self.rendered_height = 1
        self.page_rect = fitz.Rect(0, 0, 1, 1)

        self.page_view = PageView()
        self.page_view.point_selected.connect(self.handle_point_selection)
        self.page_view.point_hovered.connect(self.handle_point_hover)
        self.page_view.rect_selected.connect(self.handle_rect_selection)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(False)
        scroll_area.setAlignment(Qt.AlignmentFlag.AlignCenter)
        scroll_area.setWidget(self.page_view)
        self.scroll_area = scroll_area

        self.bookmark_tree = QTreeWidget()
        self.bookmark_tree.setHeaderLabels(["Bookmark", "Page"])
        self.bookmark_tree.itemDoubleClicked.connect(lambda item, _column: self.go_to_bookmark(item))

        bookmark_panel = QWidget()
        bookmark_layout = QVBoxLayout(bookmark_panel)
        bookmark_layout.setContentsMargins(0, 0, 0, 0)
        bookmark_layout.addWidget(self.bookmark_tree)

        bookmark_actions = QHBoxLayout()
        add_bm = QPushButton("Add")
        add_child_bm = QPushButton("Add Child")
        rename_bm = QPushButton("Rename")
        delete_bm = QPushButton("Delete")
        add_bm.clicked.connect(lambda: self.add_bookmark(child=False))
        add_child_bm.clicked.connect(lambda: self.add_bookmark(child=True))
        rename_bm.clicked.connect(self.rename_bookmark)
        delete_bm.clicked.connect(self.delete_bookmark)
        for button in (add_bm, add_child_bm, rename_bm, delete_bm):
            bookmark_actions.addWidget(button)
        bookmark_layout.addLayout(bookmark_actions)

        self.annotation_tree = QTreeWidget()
        self.annotation_tree.setHeaderLabels(["Annotation", "Page"])
        self.annotation_tree.itemDoubleClicked.connect(lambda item, _column: self.go_to_annotation(item))
        self.annotation_tree.setRootIsDecorated(False)

        annotation_panel = QWidget()
        annotation_layout = QVBoxLayout(annotation_panel)
        annotation_layout.setContentsMargins(0, 0, 0, 0)
        annotation_layout.addWidget(self.annotation_tree)

        annotation_actions = QHBoxLayout()
        show_content_btn = QPushButton("Show Content")
        delete_annot_btn = QPushButton("Delete")
        refresh_annot_btn = QPushButton("Refresh")
        show_content_btn.clicked.connect(self.show_annotation_content)
        delete_annot_btn.clicked.connect(self.delete_annotation)
        refresh_annot_btn.clicked.connect(self.load_annotations)
        for button in (show_content_btn, delete_annot_btn, refresh_annot_btn):
            annotation_actions.addWidget(button)
        annotation_layout.addLayout(annotation_actions)

        sidebar_tabs = QTabWidget()
        sidebar_tabs.addTab(bookmark_panel, "Bookmarks")
        sidebar_tabs.addTab(annotation_panel, "Annotations")

        splitter = QSplitter()
        splitter.addWidget(sidebar_tabs)
        splitter.addWidget(scroll_area)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([280, 1020])
        self.setCentralWidget(splitter)

        self.page_spinner = QSpinBox()
        self.page_spinner.setMinimum(1)
        self.page_spinner.valueChanged.connect(self.on_page_spin_changed)
        self.page_total_label = QLabel("/ 0")
        self.page_total_label.setMinimumWidth(60)

        self._build_toolbar()
        self.statusBar().showMessage("Open a PDF to start editing.")

    def _build_toolbar(self) -> None:
        toolbar = QToolBar("Main")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        open_action = QAction("Open", self)
        open_action.triggered.connect(self.open_pdf)
        save_action = QAction("Save As", self)
        save_action.triggered.connect(self.save_pdf_as)
        prev_action = QAction("Previous", self)
        prev_action.triggered.connect(lambda: self.change_page(-1))
        next_action = QAction("Next", self)
        next_action.triggered.connect(lambda: self.change_page(1))
        zoom_in_action = QAction("Zoom In", self)
        zoom_in_action.triggered.connect(lambda: self.adjust_zoom(1.15))
        zoom_out_action = QAction("Zoom Out", self)
        zoom_out_action.triggered.connect(lambda: self.adjust_zoom(1 / 1.15))
        fit_action = QAction("Fit Width", self)
        fit_action.triggered.connect(self.fit_to_width)

        comment_action = QAction("Add Comment", self)
        comment_action.triggered.connect(self.prepare_comment)
        highlight_comment_action = QAction("Highlight Comment", self)
        highlight_comment_action.triggered.connect(self.prepare_highlight_comment)
        box_comment_action = QAction("Box Comment", self)
        box_comment_action.triggered.connect(self.prepare_box_comment)
        symbol_action = QAction("Add Symbol", self)
        symbol_action.triggered.connect(self.prepare_symbol)
        text_action = QAction("Edit Text", self)
        text_action.triggered.connect(self.prepare_text_edit)
        erase_action = QAction("Erase Region", self)
        erase_action.triggered.connect(self.prepare_erase_region)
        image_action = QAction("Replace/Add Image", self)
        image_action.triggered.connect(self.prepare_replace_image)
        figure_action = QAction("Add Figure", self)
        figure_action.triggered.connect(self.prepare_draw_shape)
        cancel_action = QAction("Cancel Tool", self)
        cancel_action.triggered.connect(self.clear_pending_action)

        for action in (
            open_action,
            save_action,
            prev_action,
            next_action,
            zoom_in_action,
            zoom_out_action,
            fit_action,
            comment_action,
            highlight_comment_action,
            box_comment_action,
            symbol_action,
            text_action,
            erase_action,
            image_action,
            figure_action,
            cancel_action,
        ):
            toolbar.addAction(action)

        toolbar.addSeparator()
        toolbar.addWidget(QLabel("Page"))
        toolbar.addWidget(self.page_spinner)
        toolbar.addWidget(self.page_total_label)

    def ensure_document(self) -> bool:
        if self.doc is None:
            QMessageBox.information(self, "No PDF open", "Open a PDF before using this feature.")
            return False
        return True

    def open_pdf(self) -> None:
        if self.doc is not None and self.dirty:
            answer = QMessageBox.question(
                self,
                "Unsaved changes",
                "Opening another PDF will discard unsaved edits in the current document. Continue?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        path, _ = QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF files (*.pdf)")
        if not path:
            return
        try:
            doc = fitz.open(path)
        except Exception as exc:
            QMessageBox.critical(self, "Open failed", f"Could not open the PDF.\n\n{exc}")
            return

        if self.doc is not None:
            self.doc.close()

        self.doc = doc
        self.current_path = path
        self.current_page_index = 0
        self.zoom = 1.2
        self.dirty = False
        self.clear_pending_action()
        self.page_spinner.blockSignals(True)
        self.page_spinner.setMaximum(max(1, doc.page_count))
        self.page_spinner.setValue(1)
        self.page_spinner.blockSignals(False)
        self.page_total_label.setText(f"/ {doc.page_count}")
        self.load_bookmarks()
        self.annotations_dirty = True
        self.render_current_page()
        self.update_title()

    def save_pdf_as(self) -> None:
        if not self.ensure_document():
            return
        assert self.doc is not None

        default_name = "edited.pdf"
        if self.current_path:
            default_name = self.current_path.rsplit("\\", 1)[-1].replace(".pdf", "_edited.pdf")
        path, _ = QFileDialog.getSaveFileName(self, "Save PDF As", default_name, "PDF files (*.pdf)")
        if not path:
            return
        try:
            self.save_bookmarks()
            self.doc.save(path, garbage=4, deflate=True)
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", f"Could not save the PDF.\n\n{exc}")
            return
        self.dirty = False
        self.current_path = path
        self.update_title()
        self.statusBar().showMessage(f"Saved {path}", 5000)

    def update_title(self) -> None:
        title = "Windows PDF Editor"
        if self.current_path:
            title = f"{self.current_path} - Windows PDF Editor"
        if self.dirty:
            title = f"* {title}"
        self.setWindowTitle(title)

    def render_current_page(self) -> None:
        if not self.ensure_document():
            self.page_view.clear()
            self.bookmark_tree.clear()
            return
        assert self.doc is not None

        page = self.doc.load_page(self.current_page_index)
        self.page_rect = page.rect
        pix = page.get_pixmap(matrix=fitz.Matrix(self.zoom, self.zoom), alpha=False)
        image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888).copy()
        pixmap = QPixmap.fromImage(image)
        self.rendered_width = pix.width
        self.rendered_height = pix.height
        self.page_view.setPixmap(pixmap)
        self.page_view.resize(pixmap.size())
        self.page_spinner.blockSignals(True)
        self.page_spinner.setValue(self.current_page_index + 1)
        self.page_spinner.blockSignals(False)
        if self.annotations_dirty:
            self.load_annotations()
            self.annotations_dirty = False
        self.statusBar().showMessage(
            f"Page {self.current_page_index + 1} of {self.doc.page_count} | Zoom {int(self.zoom * 100)}%"
        )

    def on_page_spin_changed(self, value: int) -> None:
        if self.doc is None:
            return
        page_index = value - 1
        if 0 <= page_index < self.doc.page_count and page_index != self.current_page_index:
            self.current_page_index = page_index
            self.render_current_page()

    def change_page(self, delta: int) -> None:
        if not self.ensure_document():
            return
        assert self.doc is not None
        new_page = self.current_page_index + delta
        if 0 <= new_page < self.doc.page_count:
            self.current_page_index = new_page
            self.render_current_page()

    def adjust_zoom(self, factor: float) -> None:
        if not self.ensure_document():
            return
        self.zoom = max(0.2, min(6.0, self.zoom * factor))
        self.render_current_page()

    def fit_to_width(self) -> None:
        if not self.ensure_document():
            return
        assert self.doc is not None
        page = self.doc.load_page(self.current_page_index)
        available_width = max(200, self.scroll_area.viewport().width() - 40)
        self.zoom = max(0.2, min(6.0, available_width / page.rect.width))
        self.render_current_page()

    def mark_dirty(self) -> None:
        self.dirty = True
        self.update_title()

    def clear_pending_action(self) -> None:
        self.pending_action = None
        self.hovered_paragraph = None
        self.page_view.set_mode(None)
        if self.doc is not None:
            self.statusBar().showMessage("Ready.")

    def set_pending_point_action(self, name: str, payload: Optional[dict] = None, *, message: str) -> None:
        self.pending_action = PendingAction(name, payload)
        self.page_view.set_mode("point")
        self.statusBar().showMessage(message)

    def set_pending_rect_action(self, name: str, payload: Optional[dict] = None, *, message: str) -> None:
        self.pending_action = PendingAction(name, payload)
        self.page_view.set_mode("rect")
        self.statusBar().showMessage(message)

    def widget_point_to_page(self, point: QPoint) -> fitz.Point:
        x = max(0.0, min(point.x(), self.rendered_width))
        y = max(0.0, min(point.y(), self.rendered_height))
        page_x = self.page_rect.width * (x / self.rendered_width)
        page_y = self.page_rect.height * (y / self.rendered_height)
        return fitz.Point(page_x, page_y)

    def widget_rect_to_page(self, rect: QRect) -> fitz.Rect:
        top_left = self.widget_point_to_page(rect.topLeft())
        bottom_right = self.widget_point_to_page(rect.bottomRight())
        return fitz.Rect(top_left, bottom_right).normalize()

    def page_rect_to_widget(self, rect: fitz.Rect) -> QRect:
        x0 = round((rect.x0 / self.page_rect.width) * self.rendered_width)
        y0 = round((rect.y0 / self.page_rect.height) * self.rendered_height)
        x1 = round((rect.x1 / self.page_rect.width) * self.rendered_width)
        y1 = round((rect.y1 / self.page_rect.height) * self.rendered_height)
        return QRect(QPoint(x0, y0), QPoint(x1, y1)).normalized()

    def detect_alignment(self, block_rect: fitz.Rect, line_rects: list[fitz.Rect]) -> int:
        if not line_rects:
            return fitz.TEXT_ALIGN_LEFT

        first_line = line_rects[0]
        left_offset = abs(first_line.x0 - block_rect.x0)
        right_offset = abs(block_rect.x1 - first_line.x1)
        center_offset = abs((first_line.x0 + first_line.x1) / 2 - (block_rect.x0 + block_rect.x1) / 2)

        if center_offset <= max(2, block_rect.width * 0.05):
            return fitz.TEXT_ALIGN_CENTER
        if right_offset < left_offset * 0.7:
            return fitz.TEXT_ALIGN_RIGHT
        if left_offset < 2 and right_offset < max(4, block_rect.width * 0.05):
            return fitz.TEXT_ALIGN_JUSTIFY
        return fitz.TEXT_ALIGN_LEFT

    def get_text_paragraph_at_point(self, page: fitz.Page, page_point: fitz.Point) -> Optional[TextParagraph]:
        text_dict = page.get_text("dict")
        tolerance = 6 / max(self.zoom, 0.2)
        best_match: Optional[TextParagraph] = None
        best_area: Optional[float] = None

        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue

            block_rect = fitz.Rect(block["bbox"]).normalize()
            expanded_rect = fitz.Rect(
                block_rect.x0 - tolerance,
                block_rect.y0 - tolerance,
                block_rect.x1 + tolerance,
                block_rect.y1 + tolerance,
            )
            if not expanded_rect.contains(page_point):
                continue

            lines = block.get("lines", [])
            line_rects: list[fitz.Rect] = []
            paragraph_lines: list[str] = []
            font_sizes: list[float] = []
            colors: list[object] = []

            for line in lines:
                spans = line.get("spans", [])
                line_text = "".join(span.get("text", "") for span in spans).rstrip()
                if line_text:
                    paragraph_lines.append(line_text)
                    line_rects.append(fitz.Rect(line["bbox"]).normalize())
                for span in spans:
                    if span.get("text", "").strip():
                        font_sizes.append(float(span.get("size", 12)))
                        colors.append(span.get("color", 0))

            if not paragraph_lines:
                continue

            area = block_rect.get_area()
            if best_area is not None and area >= best_area:
                continue

            best_area = area
            best_match = TextParagraph(
                rect=block_rect,
                text="\n".join(paragraph_lines),
                font_size=sum(font_sizes) / len(font_sizes) if font_sizes else 12,
                color=fitz_color_to_qcolor(colors[0] if colors else 0),
                alignment=self.detect_alignment(block_rect, line_rects),
            )

        return best_match

    def detect_background_fill(self, page: fitz.Page, rect: fitz.Rect) -> tuple[float, float, float]:
        sample_w = rect.x0 + 1.0
        sample_h = rect.y0 - 3.0
        sample_x = min(max(sample_w, 0), page.rect.width)
        sample_y = min(max(sample_h, 0), page.rect.height)
        pix = page.get_pixmap(
            matrix=fitz.Matrix(0.5, 0.5),
            clip=fitz.Rect(sample_x - 2, sample_y - 2, sample_x + 2, sample_y + 2),
            alpha=False,
        )
        if pix.samples and len(pix.samples) >= 3:
            r_sum = g_sum = b_sum = 0
            count = 0
            for i in range(0, len(pix.samples), 3):
                r_sum += pix.samples[i]
                g_sum += pix.samples[i + 1]
                b_sum += pix.samples[i + 2]
                count += 1
            if count > 0:
                return (r_sum / count / 255.0, g_sum / count / 255.0, b_sum / count / 255.0)
        return (1.0, 1.0, 1.0)

    def get_fitting_font_size(
        self,
        page: fitz.Page,
        rect: fitz.Rect,
        text: str,
        alignment: int,
        font_size: float,
    ) -> float:
        working_size = max(6.0, float(font_size))
        while working_size >= 6.0:
            trial_shape = page.new_shape()
            result = trial_shape.insert_textbox(
                rect,
                text,
                fontsize=working_size,
                align=alignment,
            )
            if result >= 0:
                return working_size
            working_size -= 0.5
        raise ValueError("Replacement text does not fit inside the selected paragraph.")

    def prompt_for_comment_text(self, title: str) -> Optional[str]:
        text, ok = QInputDialog.getMultiLineText(self, title, "Comment text")
        if not ok:
            self.clear_pending_action()
            return None
        if not text.strip():
            QMessageBox.warning(self, "Missing comment", "Enter some comment text.")
            self.clear_pending_action()
            return None
        return text.strip()

    def prepare_comment(self) -> None:
        if not self.ensure_document():
            return
        self.set_pending_point_action("comment", message="Click the page where you want to add the comment.")

    def prepare_highlight_comment(self) -> None:
        if not self.ensure_document():
            return
        self.set_pending_rect_action(
            "highlight_comment",
            message="Drag a rectangle over the area you want to highlight and comment on.",
        )

    def prepare_box_comment(self) -> None:
        if not self.ensure_document():
            return
        self.set_pending_rect_action(
            "box_comment",
            message="Drag a rectangle over the area you want to box and comment on.",
        )

    def prepare_symbol(self) -> None:
        if not self.ensure_document():
            return
        dialog = SymbolDialog(self)
        data = dialog.get_data()
        if data is None:
            return
        self.set_pending_point_action("symbol", data, message="Click the page where the symbol should be inserted.")

    def prepare_text_edit(self) -> None:
        if not self.ensure_document():
            return
        self.set_pending_point_action(
            "edit_text",
            message="Move over text to preview the paragraph, then click to edit it.",
        )

    def handle_point_hover(self, point: QPoint) -> None:
        if self.pending_action is None or self.pending_action.name != "edit_text" or self.doc is None:
            self.page_view.set_overlay_rect(None)
            self.hovered_paragraph = None
            return
        if point.x() < 0 or point.y() < 0:
            self.page_view.set_overlay_rect(None)
            self.hovered_paragraph = None
            self.statusBar().showMessage("Move over text to preview the paragraph, then click to edit it.")
            return

        page = self.doc.load_page(self.current_page_index)
        paragraph = self.get_text_paragraph_at_point(page, self.widget_point_to_page(point))
        if paragraph is None:
            self.page_view.set_overlay_rect(None)
            self.hovered_paragraph = None
            self.statusBar().showMessage("Move over text to preview the paragraph, then click to edit it.")
            return

        self.hovered_paragraph = paragraph
        self.page_view.set_overlay_rect(self.page_rect_to_widget(paragraph.rect))
        preview = paragraph.text.replace("\n", " ").strip()
        if len(preview) > 80:
            preview = f"{preview[:77]}..."
        self.statusBar().showMessage(f"Selected paragraph: {preview}")

    def prepare_erase_region(self) -> None:
        if not self.ensure_document():
            return
        self.set_pending_rect_action("erase_region", message="Drag a rectangle over the content to remove.")

    def prepare_replace_image(self) -> None:
        if not self.ensure_document():
            return
        self.set_pending_rect_action(
            "replace_image",
            message="Drag a rectangle where the new or replacement image should be placed.",
        )

    def prepare_draw_shape(self) -> None:
        if not self.ensure_document():
            return
        dialog = ShapeDialog(self)
        data = dialog.get_data()
        if data is None:
            return
        self.set_pending_rect_action(
            "draw_shape",
            data,
            message="Drag a rectangle to place the new figure/object.",
        )

    def handle_point_selection(self, point: QPoint) -> None:
        if not self.ensure_document() or self.pending_action is None:
            return
        assert self.doc is not None
        page = self.doc.load_page(self.current_page_index)
        page_point = self.widget_point_to_page(point)

        try:
            if self.pending_action.name == "comment":
                text = self.prompt_for_comment_text("Add comment")
                if text is None:
                    return
                annot = page.add_text_annot(page_point, text, icon="Note")
                annot.set_info(title="Windows PDF Editor", content=text)
                annot.set_opacity(1.0)
                annot.update()
                self.annotations_dirty = True
            elif self.pending_action.name == "symbol":
                payload = self.pending_action.payload or {}
                page.insert_text(
                    page_point,
                    payload["symbol"],
                    fontsize=payload["font_size"],
                    fontname=payload.get("font_name", "helv"),
                    fontfile=payload.get("font_file"),
                    color=qcolor_to_fitz(payload["color"]),
                )
            elif self.pending_action.name == "edit_text":
                paragraph = self.hovered_paragraph
                if paragraph is None or not paragraph.rect.contains(page_point):
                    paragraph = self.get_text_paragraph_at_point(page, page_point)
                if paragraph is None:
                    self.page_view.set_overlay_rect(None)
                    self.hovered_paragraph = None
                    self.statusBar().showMessage("No paragraph found there. Move over text and click again.")
                    return

                dialog = TextEditDialog("Replace text", self, initial_text=paragraph.text)
                dialog.font_size.setValue(max(6, round(paragraph.font_size)))
                dialog.color_button.set_color(paragraph.color)
                alignment_names = {
                    fitz.TEXT_ALIGN_LEFT: "Left",
                    fitz.TEXT_ALIGN_CENTER: "Center",
                    fitz.TEXT_ALIGN_RIGHT: "Right",
                    fitz.TEXT_ALIGN_JUSTIFY: "Justify",
                }
                dialog.alignment.setCurrentText(alignment_names.get(paragraph.alignment, "Left"))
                data = dialog.get_data()
                if data is None:
                    self.clear_pending_action()
                    return

                fitted_font_size = self.get_fitting_font_size(
                    page,
                    paragraph.rect,
                    data["text"],
                    data["alignment"],
                    data["font_size"],
                )
                bg_fill = self.detect_background_fill(page, paragraph.rect)
                page.add_redact_annot(paragraph.rect, fill=bg_fill)
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
                page.insert_textbox(
                    paragraph.rect,
                    data["text"],
                    fontsize=fitted_font_size,
                    color=qcolor_to_fitz(data["color"]),
                    align=data["alignment"],
                )
            else:
                return
        except Exception as exc:
            QMessageBox.critical(self, "Edit failed", f"Could not update the PDF.\n\n{exc}")
            self.clear_pending_action()
            return

        self.mark_dirty()
        self.clear_pending_action()
        self.render_current_page()

    def handle_rect_selection(self, rect: QRect) -> None:
        if not self.ensure_document() or self.pending_action is None:
            return
        assert self.doc is not None
        page = self.doc.load_page(self.current_page_index)
        pdf_rect = self.widget_rect_to_page(rect)

        try:
            if self.pending_action.name == "erase_region":
                page.add_redact_annot(pdf_rect, fill=(1, 1, 1))
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_REMOVE)
            elif self.pending_action.name == "highlight_comment":
                text = self.prompt_for_comment_text("Add highlight comment")
                if text is None:
                    return
                annot = page.add_highlight_annot(pdf_rect)
                annot.set_colors(stroke=(1.0, 0.92, 0.23))
                annot.set_info(title="Windows PDF Editor", content=text)
                annot.update(opacity=0.35)
                self.annotations_dirty = True
            elif self.pending_action.name == "box_comment":
                text = self.prompt_for_comment_text("Add box comment")
                if text is None:
                    return
                annot = page.add_rect_annot(pdf_rect)
                annot.set_colors(stroke=(0.86, 0.16, 0.16))
                annot.set_border(width=2)
                annot.set_info(title="Windows PDF Editor", content=text)
                annot.update()
                self.annotations_dirty = True
            elif self.pending_action.name == "replace_image":
                image_path, _ = QFileDialog.getOpenFileName(
                    self,
                    "Choose image",
                    "",
                    "Image files (*.png *.jpg *.jpeg *.bmp *.gif)",
                )
                if not image_path:
                    self.clear_pending_action()
                    return
                page.add_redact_annot(pdf_rect, fill=(1, 1, 1))
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_REMOVE)
                page.insert_image(pdf_rect, filename=image_path, keep_proportion=True)
            elif self.pending_action.name == "draw_shape":
                payload = self.pending_action.payload or {}
                shape = page.new_shape()
                stroke = qcolor_to_fitz(payload["stroke"])
                fill = qcolor_to_fitz(payload["fill"])
                width = payload["width"]
                match payload["shape"]:
                    case "Rectangle":
                        shape.draw_rect(pdf_rect)
                    case "Ellipse":
                        shape.draw_oval(pdf_rect)
                    case "Arrow":
                        start = pdf_rect.tl
                        end = pdf_rect.br
                        shape.draw_line(start, end)
                        arrow_size = min(pdf_rect.width, pdf_rect.height) * 0.2
                        angle = math.atan2(end.y - start.y, end.x - start.x)
                        barb_angle = math.radians(25)
                        shape.draw_line(end, fitz.Point(
                            end.x - arrow_size * math.cos(angle - barb_angle),
                            end.y - arrow_size * math.sin(angle - barb_angle),
                        ))
                        shape.draw_line(end, fitz.Point(
                            end.x - arrow_size * math.cos(angle + barb_angle),
                            end.y - arrow_size * math.sin(angle + barb_angle),
                        ))
                        fill = None
                    case _:
                        raise ValueError("Unsupported figure type.")
                shape.finish(color=stroke, fill=fill, width=width)
                shape.commit()
            else:
                return
        except Exception as exc:
            QMessageBox.critical(self, "Edit failed", f"Could not update the PDF.\n\n{exc}")
            self.clear_pending_action()
            return

        self.mark_dirty()
        self.clear_pending_action()
        self.render_current_page()

    def load_bookmarks(self) -> None:
        self.bookmark_tree.clear()
        if self.doc is None:
            return

        parents: dict[int, QTreeWidgetItem] = {}
        for level, title, page_number, *_ in self.doc.get_toc(simple=False):
            item = QTreeWidgetItem([title, str(page_number)])
            item.setData(0, PAGE_ROLE, max(0, page_number - 1))
            if level == 1:
                self.bookmark_tree.addTopLevelItem(item)
            else:
                parent = parents.get(level - 1)
                if parent is None:
                    self.bookmark_tree.addTopLevelItem(item)
                else:
                    parent.addChild(item)
            parents[level] = item

        self.bookmark_tree.expandAll()

    def save_bookmarks(self) -> None:
        if self.doc is None:
            return
        toc: list[list] = []

        def visit(item: QTreeWidgetItem, level: int) -> None:
            toc.append([level, item.text(0), int(item.data(0, PAGE_ROLE)) + 1])
            for index in range(item.childCount()):
                visit(item.child(index), level + 1)

        for index in range(self.bookmark_tree.topLevelItemCount()):
            visit(self.bookmark_tree.topLevelItem(index), 1)

        self.doc.set_toc(toc)

    def go_to_bookmark(self, item: QTreeWidgetItem) -> None:
        if self.doc is None:
            return
        page_index = int(item.data(0, PAGE_ROLE))
        if 0 <= page_index < self.doc.page_count:
            self.current_page_index = page_index
            self.render_current_page()

    def add_bookmark(self, *, child: bool) -> None:
        if not self.ensure_document():
            return
        selected = self.bookmark_tree.currentItem()
        dialog = BookmarkDialog(self, page_number=self.current_page_index + 1)
        if self.doc is not None:
            dialog.page_input.setMaximum(self.doc.page_count)
        data = dialog.get_data()
        if data is None:
            return

        item = QTreeWidgetItem([data["title"], str(data["page_number"])])
        item.setData(0, PAGE_ROLE, data["page_number"] - 1)

        if child and selected is not None:
            selected.addChild(item)
            selected.setExpanded(True)
        elif selected is not None and selected.parent() is not None:
            parent = selected.parent()
            parent.insertChild(parent.indexOfChild(selected) + 1, item)
        elif selected is not None:
            row = self.bookmark_tree.indexOfTopLevelItem(selected)
            self.bookmark_tree.insertTopLevelItem(row + 1, item)
        else:
            self.bookmark_tree.addTopLevelItem(item)

        self.mark_dirty()
        self.save_bookmarks()

    def rename_bookmark(self) -> None:
        if not self.ensure_document():
            return
        item = self.bookmark_tree.currentItem()
        if item is None:
            QMessageBox.information(self, "No bookmark selected", "Select a bookmark to edit it.")
            return
        dialog = BookmarkDialog(
            self,
            title=item.text(0),
            page_number=int(item.data(0, PAGE_ROLE)) + 1,
        )
        if self.doc is not None:
            dialog.page_input.setMaximum(self.doc.page_count)
        data = dialog.get_data()
        if data is None:
            return
        item.setText(0, data["title"])
        item.setText(1, str(data["page_number"]))
        item.setData(0, PAGE_ROLE, data["page_number"] - 1)
        self.mark_dirty()
        self.save_bookmarks()

    def delete_bookmark(self) -> None:
        if not self.ensure_document():
            return
        item = self.bookmark_tree.currentItem()
        if item is None:
            QMessageBox.information(self, "No bookmark selected", "Select a bookmark to delete it.")
            return
        if item.parent() is not None:
            item.parent().removeChild(item)
        else:
            row = self.bookmark_tree.indexOfTopLevelItem(item)
            self.bookmark_tree.takeTopLevelItem(row)
        self.mark_dirty()
        self.save_bookmarks()

    ANNOT_KEY_ROLE = Qt.ItemDataRole.UserRole + 2

    def _annot_key(self, annot) -> str:
        rect = annot.rect
        return f"{rect.x0:.1f},{rect.y0:.1f},{rect.x1:.1f},{rect.y1:.1f}"

    def _annot_display(self, annot) -> str:
        annot_type = annot.type[1] if isinstance(annot.type, tuple) else str(annot.type)
        info = annot.info
        content = info.get("content", "") if info else ""
        preview = content.replace("\n", " ") if content else ""
        if len(preview) > 60:
            preview = f"{preview[:57]}..."
        return f"[{annot_type}] {preview}" if preview else f"[{annot_type}]"

    def load_annotations(self) -> None:
        self.annotation_tree.clear()
        if self.doc is None:
            return
        for page_idx in range(self.doc.page_count):
            page = self.doc.load_page(page_idx)
            for annot in page.annots():
                item = QTreeWidgetItem([self._annot_display(annot), str(page_idx + 1)])
                item.setData(0, PAGE_ROLE, page_idx)
                item.setData(0, self.ANNOT_KEY_ROLE, self._annot_key(annot))
                self.annotation_tree.addTopLevelItem(item)

    def _find_annot_by_item(self, item: QTreeWidgetItem) -> tuple[Optional[object], Optional[object]]:
        page_index = int(item.data(0, PAGE_ROLE))
        target_key = item.data(0, self.ANNOT_KEY_ROLE)
        page = self.doc.load_page(page_index)
        for annot in page.annots():
            if self._annot_key(annot) == target_key:
                return page, annot
        return page, None

    def go_to_annotation(self, item: QTreeWidgetItem) -> None:
        if self.doc is None:
            return
        page_index = int(item.data(0, PAGE_ROLE))
        if 0 <= page_index < self.doc.page_count:
            self.current_page_index = page_index
            self.render_current_page()

    def show_annotation_content(self) -> None:
        if self.doc is None:
            return
        item = self.annotation_tree.currentItem()
        if item is None:
            QMessageBox.information(self, "No annotation selected", "Select an annotation to view its content.")
            return
        _page, annot = self._find_annot_by_item(item)
        if annot is None:
            QMessageBox.information(self, "Not found", "Annotation no longer exists.")
            return
        info = annot.info
        content = info.get("content", "") if info else ""
        if content:
            QMessageBox.information(self, "Annotation content", content)
        else:
            QMessageBox.information(self, "Annotation content", "(No text content in this annotation.)")

    def delete_annotation(self) -> None:
        if self.doc is None:
            return
        item = self.annotation_tree.currentItem()
        if item is None:
            QMessageBox.information(self, "No annotation selected", "Select an annotation to delete.")
            return
        answer = QMessageBox.question(
            self,
            "Delete annotation",
            "Delete the selected annotation? This cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        page, annot = self._find_annot_by_item(item)
        if annot is None:
            QMessageBox.information(self, "Not found", "Annotation no longer exists.")
            self.load_annotations()
            return
        page.delete_annot(annot)
        self.mark_dirty()
        self.annotations_dirty = True
        self.render_current_page()

    def closeEvent(self, event) -> None:
        if self.doc is None or not self.dirty:
            event.accept()
            return
        answer = QMessageBox.question(
            self,
            "Unsaved changes",
            "You have unsaved PDF edits. Close without saving?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if answer == QMessageBox.StandardButton.Yes:
            event.accept()
        else:
            event.ignore()


def main() -> int:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
