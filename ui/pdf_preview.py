from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import QPointF, Qt
from PyQt6.QtWidgets import QLabel, QStackedLayout, QVBoxLayout, QWidget
from PyQt6.QtPdf import QPdfDocument
from PyQt6.QtPdfWidgets import QPdfView

class PdfPreviewPanel(QWidget):
	"""Panneau de prévisualisation PDF avec fallback texte si QtPdf n'est pas dispo."""

	def __init__(self, parent: QWidget | None = None):
		super().__init__(parent)
		self._pdf_path: str | None = None
		self._pdf_available = False

		self._placeholder = QLabel()
		self._ensure_valid_font(self._placeholder)
		self._placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
		self._placeholder.setWordWrap(True)
		self._placeholder.setStyleSheet(
			"""
			QLabel {
				border: 1px dashed #DEDEDE;
				border-radius: 8px;
				background: #9E9E9E;
				color: #1C2844;
				font-size: 13px;
				padding: 24px;
			}
			"""
		)

		self._stack = QStackedLayout()
		self._stack.addWidget(self._placeholder)

		root = QVBoxLayout(self)
		root.setContentsMargins(0, 0, 0, 0)
		root.addLayout(self._stack)

		self._pdf_view = None
		self._pdf_doc = None
		self._init_pdf_backend()
		self.show_placeholder("Aucun PDF chargé.")

	def _init_pdf_backend(self) -> None:
		self._pdf_doc = QPdfDocument(self)
		self._pdf_view = QPdfView(self)
		self._ensure_valid_font(self._pdf_view)
		self._pdf_view.setDocument(self._pdf_doc)
		self._pdf_view.setPageMode(QPdfView.PageMode.SinglePage)
		self._pdf_view.setZoomMode(QPdfView.ZoomMode.FitInView)
		self._stack.addWidget(self._pdf_view)
		self._pdf_available = True

	def _ensure_valid_font(self, widget: QWidget) -> None:
		font = widget.font()
		if font.pointSize() <= 0 and font.pixelSize() <= 0:
			print(f"Widget problématique : {widget.objectName()} (Type: {type(widget).__name__})")
			font.setPointSize(10)
			widget.setFont(font)

	def show_placeholder(self, text: str) -> None:
		self._placeholder.setText(text)
		self._stack.setCurrentWidget(self._placeholder)

	def load_pdf(self, path: str) -> None:
		self._pdf_path = path
		if not path or not Path(path).exists():
			self.show_placeholder("Fichier PDF introuvable.")
			return

		if not self._pdf_available or self._pdf_doc is None or self._pdf_view is None:
			self.show_placeholder(
				"Prévisualisation PDF indisponible (module QtPdf manquant).\n"
				f"Fichier généré :\n{path}"
			)
			return

		status = self._pdf_doc.load(path)
		if status == self._pdf_doc.Error.None_:
			# show first page by default
			self._stack.setCurrentWidget(self._pdf_view)
			self.show_page(0)
		else:
			self.show_placeholder(f"Impossible de charger le PDF.\nFichier :\n{path}")

	def page_count(self) -> int:
		if not self._pdf_available or self._pdf_doc is None:
			return 0
		return max(0, self._pdf_doc.pageCount())

	def current_page(self) -> int:
		if not self._pdf_available or self._pdf_view is None:
			return 0
		navigator = self._pdf_view.pageNavigator()
		if navigator is None:
			return 0
		return max(0, navigator.currentPage())

	def show_page(self, index: int) -> None:
		"""Display a single page by index (0-based). Clamped to valid range."""
		if not self._pdf_available or self._pdf_doc is None or self._pdf_view is None:
			return

		count = self.page_count()
		if count == 0:
			return

		idx = max(0, min(index, count - 1))
		navigator = self._pdf_view.pageNavigator()
		if navigator is None:
			return
		navigator.jump(idx, QPointF(0.0, 0.0), self._pdf_view.zoomFactor())

	def set_page(self, index: int) -> None:
		if not self._pdf_available or self._pdf_doc is None or self._pdf_view is None:
			return
		navigator = self._pdf_view.pageNavigator()
		if navigator:
			navigator.jump(index - 1, QPointF(0.0, 0.0), self._pdf_view.zoomFactor())  # index est 0-based

	def set_zoom(self, zoom: int) -> None:
		if not self._pdf_available or self._pdf_view is None:
			return
		if zoom == 0:
			self._pdf_view.setZoomMode(QPdfView.ZoomMode.FitInView)
		else:
			self._pdf_view.setZoomMode(QPdfView.ZoomMode.Custom)
			self._pdf_view.setZoomFactor((zoom / 100)+0.75)