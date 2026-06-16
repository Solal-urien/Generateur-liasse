"""
Fenêtre principale (main_window): onglets Document / Style / Page / Graphiques
+ panneau de prévisualisation PDF.
"""

import tempfile
from pathlib import Path

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QSplitter, QTabWidget, QStatusBar, QToolBar,
    QFileDialog, QMessageBox, QPushButton, QLabel,
    QSizePolicy, QSpinBox,
)
from PyQt6.QtCore import Qt, QThread, QObject, pyqtSignal, QTimer
from PyQt6.QtGui import QAction

from state_manager import state_manager
from pdf_generator import generate_pdf
from ui.tab_document import DocumentTab
from ui.tab_style import StyleTab
from ui.tab_page import PageTab
from ui.tab_line import LineTab
from ui.tab_charts import ChartsTab
from ui.pdf_preview import PdfPreviewPanel

from util import NUANCIER_COULEURS

# Dossier de l'application
dossier_courant = Path(__file__).resolve().parent


# ── Worker thread pour génération PDF ──────────────────────────────────────

class PdfWorker(QObject):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, output_path: str):
        super().__init__()
        self._output = output_path

    def run(self):
        try:
            sm = state_manager
            generate_pdf(
                model=sm.workbook,
                output_path=self._output,
                gparams=sm.gparams,
                charts=sm.charts,
            )
            self.finished.emit(self._output)
        except Exception as e:
            self.error.emit(str(e))


# ── Fenêtre principale ─────────────────────────────────────────────────────

class MainWindow(QMainWindow):

    # Couleurs du thème
    PRIMARY = NUANCIER_COULEURS["Palette principale"][0] # Bleu foncé
    ACCENT = NUANCIER_COULEURS["Palette principale"][3]
    OTHER_ACCENT = NUANCIER_COULEURS["Palette principale"][4]
    BG = NUANCIER_COULEURS["Palette principale"][5]
    ECRITURE = NUANCIER_COULEURS["Palette principale"][6]
    PANEL_BG = NUANCIER_COULEURS["Palette principale"][11]

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Générateur de Liasse PDF")
        self.setMinimumSize(1280, 800)
        self._pdf_thread: QThread | None = None
        self._preview_path: str | None = None
        self._preview_page: int = 1
        self._zoom: int = 0
        self._preview_timer = QTimer()
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(1200)
        self._preview_timer.timeout.connect(self._auto_preview)

        self._apply_theme()
        self._build_ui()
        self._connect_state()

    # ── Thème ──────────────────────────────────────────────────────────────

    def _apply_theme(self):
        self.setStyleSheet(f"""
            QMainWindow, QWidget {{
                background-color: {self.BG};
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 13px;
                color: {self.ECRITURE};
            }}
            QTabWidget::pane {{
                border: 1px solid {NUANCIER_COULEURS["Palette principale"][4]};
                background: {NUANCIER_COULEURS["Arrière-plans"][3]};
                border-radius: 6px;
            }}
            QTabBar::tab {{
                background: {NUANCIER_COULEURS["Arrière-plans"][0]};
                color: {NUANCIER_COULEURS["Palette principale"][6]};
                padding: 8px 20px;
                border: none;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                margin-right: 2px;
                font-weight: 500;
            }}
            QTabBar::tab:selected {{
                background: {self.ACCENT};
                color: white;
                font-weight: 600;
            }}
            QTabBar::tab:hover:!selected {{
                background: {NUANCIER_COULEURS["Palette principale"][4]};
            }}
            QPushButton {{
                background-color: {self.ACCENT};
                color: white;
                border: none;
                padding: 7px 18px;
                border-radius: 5px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                background-color: {NUANCIER_COULEURS["Palette principale"][2]};
            }}
            QPushButton:disabled {{
                background-color: {NUANCIER_COULEURS["Arrière-plans"][0]};
            }}
            QPushButton#secondary {{
                background-color: {NUANCIER_COULEURS["Arrière-plans"][0]};
                color: {self.PRIMARY};
                border: 1px solid {NUANCIER_COULEURS["Palette principale"][4]};
            }}
            QPushButton#secondary:hover {{
                background-color: {NUANCIER_COULEURS["Arrière-plans"][1]};
            }}
            QPushButton#danger {{
                background-color: {NUANCIER_COULEURS["Palette principale"][12]};
            }}
            QPushButton#danger:hover {{
                background-color: {NUANCIER_COULEURS["Palette principale"][13]};
            }}
            QStatusBar {{
                background: {self.PRIMARY};
                color: white;
                padding: 3px 8px;
                font-size: 11px;
            }}
            QToolBar {{
                background: {self.PRIMARY};
                border: none;
                spacing: 6px;
                padding: 4px 8px;
            }}
            QToolBar QToolButton {{
                background: transparent;
                color: white;
                padding: 5px 10px;
                border-radius: 4px;
                font-size: 12px;
            }}
            QToolBar QToolButton:hover {{
                background: rgba(255,255,255,0.15);
            }}
            QSplitter::handle {{
                background: {NUANCIER_COULEURS["Palette principale"][4]};
                width: 3px;
            }}
            QLabel#sectionTitle {{
                font-size: 15px;
                font-weight: 700;
                color: {self.PRIMARY};
                padding: 4px 0;
            }}
            QGroupBox {{
                border: 1px solid {NUANCIER_COULEURS["Palette principale"][4]};
                border-radius: 6px;
                margin-top: 10px;
                padding: 8px;
                font-weight: 600;
                color: {self.PRIMARY};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
            }}
            QScrollArea {{
                border: none;
                background: transparent;
            }}
            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox {{
                border: 1px solid {NUANCIER_COULEURS["Palette principale"][4]};
                border-radius: 4px;
                padding: 4px 8px;
                background: white;
                selection-background-color: {self.ACCENT};
            }}
            QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus {{
                border-color: {self.ACCENT};
            }}
            /* FIX: curseur arrow sur les flèches des spinbox */
            QSpinBox::up-button, QDoubleSpinBox::up-button,
            QSpinBox::down-button, QDoubleSpinBox::down-button {{
                subcontrol-origin: border;
                width: 18px;
            }}
            QSpinBox::up-button, QDoubleSpinBox::up-button {{
                subcontrol-position: top right;
            }}
            QSpinBox::down-button, QDoubleSpinBox::down-button {{
                subcontrol-position: bottom right;
            }}
            QCheckBox::indicator {{
                width: 16px; height: 16px;
                border: 2px solid {NUANCIER_COULEURS["Palette principale"][4]};
                border-radius: 3px;
            }}
            QCheckBox::indicator:checked {{
                background-color: {self.ACCENT};
                border-color: {self.ACCENT};
            }}
            QListWidget {{
                border: 1px solid {NUANCIER_COULEURS["Palette principale"][4]};
                border-radius: 6px;
                background: white;
            }}
            QListWidget::item {{
                padding: 6px 10px;
                border-bottom: 1px solid {NUANCIER_COULEURS["Arrière-plans"][0]};
            }}
            QListWidget::item:selected {{
                background: {NUANCIER_COULEURS["Palette principale"][4]};
                color: {self.PRIMARY};
            }}
        """)

    # ── Construction UI ────────────────────────────────────────────────────

    def _build_ui(self):
        toolbar = self._build_toolbar()
        self.addToolBar(toolbar)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(0)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(4)

        # Panneau gauche : onglets de paramétrage
        left_panel = QWidget()
        left_panel.setMaximumWidth(550)
        left_panel.setMinimumWidth(380)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)

        self.tabs = QTabWidget()
        self.tab_document = DocumentTab()
        self.tab_style = StyleTab()
        self.tab_page = PageTab()
        self.tab_line = LineTab()
        self.tab_charts = ChartsTab()

        self.tabs.addTab(self.tab_document, "📄 Document")
        self.tabs.addTab(self.tab_style, "🎨 Style")
        self.tabs.addTab(self.tab_page, "📋 Page")
        self.tabs.addTab(self.tab_line, "📑 Lignes")
        self.tabs.addTab(self.tab_charts, "📊 Graphiques")

        left_layout.addWidget(self.tabs)

        action_bar = self._build_action_bar()
        left_layout.addWidget(action_bar)

        # Panneau droit : prévisualisation PDF (une seule instance)
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(8, 0, 0, 0)

        preview_label = QLabel("Prévisualisation PDF")
        preview_label.setObjectName("sectionTitle")
        right_layout.addWidget(preview_label)

        # Contrôles de navigation de page
        preview_controls = QWidget()
        preview_controls_layout = QHBoxLayout(preview_controls)
        preview_controls_layout.setContentsMargins(0, 0, 0, 0)
        preview_controls_layout.setSpacing(8)

        preview_controls_layout.addWidget(QLabel("Page"))
        self._page_selector = QSpinBox()
        self._page_selector.setMinimum(1)
        self._page_selector.setMaximum(9999)  # sera mis à jour après chargement
        self._page_selector.setValue(1)
        # La connexion est faite APRÈS la création du preview pour éviter les appels prématurés
        preview_controls_layout.addWidget(self._page_selector)

        self._lbl_total_pages = QLabel("/ —")
        preview_controls_layout.addWidget(self._lbl_total_pages)
        preview_controls_layout.addSpacing(16)
        preview_controls_layout.addWidget(QLabel("Zoom"))
        self._zoom_selector = QSpinBox()
        self._zoom_selector.setMinimum(0)
        self._zoom_selector.setMaximum(200)
        self._zoom_selector.setSingleStep(10)
        self._zoom_selector.setValue(0)
        self._zoom_selector.setSuffix(" %")
        preview_controls_layout.addWidget(self._zoom_selector)
        preview_controls_layout.addStretch()

        right_layout.addWidget(preview_controls)

        # Unique instance du panneau de prévisualisation
        self.preview = PdfPreviewPanel()
        right_layout.addWidget(self.preview)

        # Connexion du sélecteur de page (après création de self.preview)
        self._page_selector.valueChanged.connect(self._on_preview_page_changed)
        self._zoom_selector.valueChanged.connect(self._on_preview_zoom_changed)

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([460, 800])

        main_layout.addWidget(splitter)

        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("Prêt — déposez un fichier Excel pour commencer.")

        self.preview.show_placeholder("Chargez un fichier Excel\npuis cliquez sur Prévisualiser.")

    def _build_toolbar(self) -> QToolBar:
        tb = QToolBar("Principal")
        tb.setMovable(False)

        act_save_profile = QAction("💾 Sauvegarder profil", self)
        act_save_profile.triggered.connect(self._save_profile)
        tb.addAction(act_save_profile)

        act_load_profile = QAction("📥 Charger profil", self)
        act_load_profile.triggered.connect(self._load_profile)
        tb.addAction(act_load_profile)

        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        spacer.setStyleSheet(f"background: {self.PRIMARY};")
        tb.addWidget(spacer)

        from init import reinitialize
        act_reinitialize = QAction("🔄 Réinitialiser", self)
        act_reinitialize.triggered.connect(reinitialize)
        tb.addAction(act_reinitialize)

        return tb

    def _build_action_bar(self) -> QWidget:
        bar = QWidget()
        bar.setStyleSheet(f"background: {self.PANEL_BG}; border-top: 1px solid #d0d9e8;")
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(10, 8, 10, 8)

        btn_preview = QPushButton("🔄 Prévisualiser")
        btn_preview.setObjectName("secondary")
        btn_preview.clicked.connect(self._trigger_preview)

        btn_generate = QPushButton("⬇️  Générer PDF")
        btn_generate.setStyleSheet("color: black;")
        btn_generate.clicked.connect(self._generate_pdf)

        layout.addWidget(btn_preview)
        layout.addStretch()
        layout.addWidget(btn_generate)

        self._btn_generate = btn_generate
        self._btn_preview = btn_preview
        return bar

    # ── State listeners ───────────────────────────────────────────────────

    def _connect_state(self):
        sm = state_manager
        sm.subscribe("workbook_loaded", self._on_workbook_loaded)
        sm.subscribe("pdf_generated", self._on_pdf_generated)
        sm.subscribe("error", self._on_error)
        sm.subscribe("sheet_order_changed", lambda _: self._schedule_preview())
        sm.subscribe("sheet_visibility_changed", lambda _: self._schedule_preview())
        sm.subscribe("page_params_changed", lambda _: self._schedule_preview())
        sm.subscribe("charts_changed", lambda _: self._schedule_preview())

    def _on_workbook_loaded(self, wb):
        n = len(wb.sheets)
        self.status.showMessage(f"✅ Fichier chargé — {n} feuille(s) détectées.")
        self._schedule_preview()

    def _on_pdf_generated(self, path):
        self._load_preview(path, page=1)
        self.status.showMessage(f"✅ PDF généré : {path}")
        self._btn_generate.setEnabled(True)
        self._btn_preview.setEnabled(True)

    def _on_error(self, msg):
        QMessageBox.critical(self, "Erreur", str(msg))
        self.status.showMessage(f"❌ Erreur : {msg}")
        self._btn_generate.setEnabled(True)
        self._btn_preview.setEnabled(True)

    # ── Actions ────────────────────────────────────────────────────────────

    def _open_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Ouvrir un fichier Excel",
            str(dossier_courant), "Fichiers Excel (*.xlsx *.xlsm *.xls)"
        )
        if path:
            self.status.showMessage(f"Chargement de {Path(path).name}…")
            state_manager.load_excel(path)

    def _save_profile(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Sauvegarder le profil",
            str(dossier_courant)+"/profils", "Profils JSON (*.json)"
        )
        if path:
            state_manager.save_profile(path)
            self.status.showMessage(f"Profil sauvegardé : {path}")

    def _load_profile(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Charger un profil",
            str(dossier_courant)+"/profils", "Profils JSON (*.json)"
        )
        if path:
            state_manager.load_profile(path)
            self.status.showMessage(f"Profil chargé : {path}")
            self._schedule_preview()

    def _schedule_preview(self):
        if state_manager.workbook:
            self._preview_timer.start()

    def _auto_preview(self):
        self._trigger_preview(auto=True)

    def _trigger_preview(self, auto: bool = False):
        if not state_manager.workbook:
            if not auto:
                QMessageBox.information(self, "Info", "Chargez d'abord un fichier Excel.")
            return
        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp.close()
        self._run_pdf_generation(tmp.name, preview_only=True)

    def _generate_pdf(self):
        if not state_manager.workbook:
            QMessageBox.information(self, "Info", "Chargez d'abord un fichier Excel.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Enregistrer le PDF",
            str(dossier_courant.parent), "PDF (*.pdf)"
        )
        if path:
            self._run_pdf_generation(path, preview_only=False)

    def _run_pdf_generation(self, output_path: str, preview_only: bool = False):
        if self._pdf_thread:
            try:
                if self._pdf_thread.isRunning():
                    return
            except RuntimeError:
                self._pdf_thread = None
        self._btn_generate.setEnabled(False)
        self._btn_preview.setEnabled(False)
        self.status.showMessage("Génération du PDF en cours…")

        self._pdf_thread = QThread()
        self._worker = PdfWorker(output_path)
        self._worker.moveToThread(self._pdf_thread)
        self._pdf_thread.started.connect(self._worker.run)
        self._worker.finished.connect(lambda p: self._pdf_done(p, preview_only))
        self._worker.error.connect(self._on_error)
        self._worker.finished.connect(self._pdf_thread.quit)
        self._worker.error.connect(self._pdf_thread.quit)
        self._pdf_thread.finished.connect(self._pdf_thread.deleteLater)
        self._pdf_thread.finished.connect(lambda: setattr(self, "_pdf_thread", None))
        self._pdf_thread.finished.connect(lambda: setattr(self, "_worker", None))
        self._pdf_thread.start()

    def _pdf_done(self, path: str, preview_only: bool):
        self._btn_generate.setEnabled(True)
        self._btn_preview.setEnabled(True)
        if not preview_only:
            self._load_preview(path, page=1)
            self.status.showMessage(f"✅ PDF généré : {path}")
        else:
            page_to_restore = self._preview_page
            self._load_preview(path, page=page_to_restore) #0 Conserve la page actuelle si possible
            self.status.showMessage("Prévisualisation mise à jour.")

    def _load_preview(self, path: str, page: int = 1):
        """Charge un PDF dans le panneau de prévisualisation et met à jour les contrôles."""
        self._preview_path = path
        self.preview.load_pdf(path)

        total = self.preview.page_count()
        self._lbl_total_pages.setText(f"/ {total}" if total > 0 else "/ —")

        # Mettre à jour le max AVANT de changer la valeur pour éviter le clamp
        self._page_selector.blockSignals(True)
        self._page_selector.setMaximum(max(1, total))
        self._page_selector.setValue(page)
        self._page_selector.blockSignals(False)

        self._preview_page = page
        self.preview.set_page(page)

    def _on_preview_zoom_changed(self, zoom: int):
        """Appelé quand l'utilisateur change le zoom dans le sélecteur."""
        self._zoom = zoom
        self.preview.set_zoom(zoom)

    def _on_preview_page_changed(self, page: int):
        """Appelé quand l'utilisateur change la page dans le sélecteur."""
        self._preview_page = page
        self.preview.set_page(page)