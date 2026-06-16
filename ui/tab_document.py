"""
Onglet Document :
  - Chargement du fichier Excel
  - Réorganisation des feuilles par drag & drop
  - Inclusion / exclusion sélective
  - Modification du titre et de la date de page de garde
"""
from pathlib import Path

from PyQt6.QtWidgets import (
    QSpinBox, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QCheckBox, QGroupBox,
    QLineEdit, QFileDialog, QAbstractItemView, QSizePolicy,
)
from PyQt6.QtCore import Qt, pyqtSignal

from state_manager import state_manager

# Dossier parent de l'application
dossier_courant = Path(__file__).resolve().parent.parent.parent

class SheetListItem(QWidget):
    """Widget représentant une feuille dans la liste (checkbox + nom + indicateur)."""

    toggled = pyqtSignal(str, bool)

    def __init__(self, sheet_name: str, included: bool, has_chart: bool, parent=None):
        super().__init__(parent)
        self.sheet_name = sheet_name
        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 3, 6, 3)
        layout.setSpacing(8)

        # Poignée drag
        handle = QLabel("⠿")
        handle.setStyleSheet("color: #aabbcc; font-size: 16px;")
        handle.setFixedWidth(16)
        layout.addWidget(handle)

        # Checkbox inclusion
        self.chk = QCheckBox()
        self.chk.setChecked(included)
        self.chk.stateChanged.connect(self._on_toggle)
        layout.addWidget(self.chk)

        # Nom
        name_lbl = QLabel(sheet_name)
        name_lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        layout.addWidget(name_lbl)

        # Badge graphique
        if has_chart:
            badge = QLabel("📊")
            badge.setToolTip("Contient un graphique")
            layout.addWidget(badge)

    def _on_toggle(self, state: int):
        self.toggled.emit(self.sheet_name, state == Qt.CheckState.Checked.value)


class DraggableList(QListWidget):
    """QListWidget avec drag & drop interne pour réordonner."""

    order_changed = pyqtSignal(list)  # émet la nouvelle liste de noms

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.setSpacing(2)
        self.model().rowsMoved.connect(self._on_rows_moved)

    def _on_rows_moved(self, *_):
        names = []
        for i in range(self.count()):
            item = self.item(i)
            w = self.itemWidget(item)
            if w:
                names.append(w.sheet_name)
        self.order_changed.emit(names)


class DocumentTab(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self._building = False
        self._build_ui()
        state_manager.subscribe("workbook_loaded", self._on_workbook_loaded)
        state_manager.subscribe("profile_loaded", self._refresh)

    # ── Construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(12)

        # ── Section : fichier source ──
        grp_file = QGroupBox("Fichier source")
        file_layout = QHBoxLayout(grp_file)
        self._lbl_file = QLabel("Aucun fichier chargé")
        self._lbl_file.setStyleSheet("color: #667788; font-style: italic;")
        self._lbl_file.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        btn_open = QPushButton("📂 Ouvrir")
        btn_open.clicked.connect(self._open_excel)
        file_layout.addWidget(self._lbl_file)
        file_layout.addWidget(btn_open)
        root.addWidget(grp_file)

        # ── Section : années considérées ──
        grp_years = QGroupBox("Années considérées")
        years_layout = QVBoxLayout(grp_years)
        row_fs = QHBoxLayout()
        row_fs.addWidget(QLabel("Année début affichage (ré-ouvrir l'excel pour appliquer) :"))
        self._spin_font = QSpinBox()
        self._spin_font.setRange(2020, 2026)
        self._spin_font.setValue(2025)
        self._spin_font.setSingleStep(1)
        self._spin_font.setFixedWidth(90)
        self._spin_font.valueChanged.connect(self._on_years_changed)
        # row_fs.addStretch()
        row_fs.addWidget(self._spin_font)
        years_layout.addLayout(row_fs)
        root.addWidget(grp_years)

        # ── Section : page de garde ──
        grp_cover = QGroupBox("Page de garde")
        cover_layout = QVBoxLayout(grp_cover)
        cover_layout.setSpacing(6)
        row_title = QHBoxLayout()
        row_title.addWidget(QLabel("Titre :"))
        self._title_edit = QLineEdit()
        self._title_edit.setPlaceholderText("Titre de la liasse")
        self._title_edit.textChanged.connect(self._on_title_changed)
        row_title.addWidget(self._title_edit)
        cover_layout.addLayout(row_title)
        row_author = QHBoxLayout()
        row_author.addWidget(QLabel("Sous-titre :"))
        self._author_edit = QLineEdit()
        self._author_edit.setPlaceholderText("BEE, BEH...")
        self._author_edit.textChanged.connect(self._on_author_changed)
        row_author.addWidget(self._author_edit)
        cover_layout.addLayout(row_author)
        root.addWidget(grp_cover)

        # ── Section : feuilles ──
        grp_sheets = QGroupBox("Feuilles (glisser pour réordonner)")
        sheets_layout = QVBoxLayout(grp_sheets)

        # Boutons sélection rapide
        sel_bar = QHBoxLayout()
        btn_all = QPushButton("Tout inclure")
        btn_all.setObjectName("secondary")
        btn_all.clicked.connect(self._select_all)
        btn_none = QPushButton("Tout exclure")
        btn_none.setObjectName("secondary")
        btn_none.clicked.connect(self._select_none)
        sel_bar.addWidget(btn_all)
        sel_bar.addWidget(btn_none)
        sel_bar.addStretch()
        sheets_layout.addLayout(sel_bar)
        self._list = DraggableList()
        self._list.setMinimumHeight(200)
        self._list.order_changed.connect(self._on_order_changed)
        sheets_layout.addWidget(self._list)
        self._lbl_empty = QLabel("Chargez un fichier Excel pour\nafficher les feuilles.")
        self._lbl_empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._lbl_empty.setStyleSheet("color: #2F2F2F; font-style: italic;")
        sheets_layout.addWidget(self._lbl_empty)
        root.addWidget(grp_sheets)
        root.addStretch()

    # ── Callbacks état ───────────────────────────────────────────────────────

    def _on_workbook_loaded(self, wb):
        path_exc = Path(state_manager.state.current_excel_path).name if state_manager.state.current_excel_path else None
        self._lbl_file.setText(path_exc if path_exc else "Chargé")
        self._lbl_file.setStyleSheet("color: #7891C7; font-weight: 600;")
        self._refresh(None)

    def _refresh(self, _=None):
        wb = state_manager.workbook
        if not wb:
            return
        self._building = True
        self._list.clear()

        sheets_sorted = sorted(wb.sheets, key=lambda s: s.page_order)
        has_any = len(sheets_sorted) > 0
        self._lbl_empty.setVisible(not has_any)

        for sheet in sheets_sorted:
            has_chart = any(
                getattr(t, 'source', '') == 'chart'
                for t in sheet.tables
            )
            item = QListWidgetItem(self._list)
            item.setData(Qt.ItemDataRole.UserRole, sheet.name)

            widget = SheetListItem(sheet.name, sheet.include, has_chart)
            widget.toggled.connect(self._on_include_toggled)

            item.setSizeHint(widget.sizeHint())
            self._list.setItemWidget(item, widget)

        # Pré-remplir titre/auteur depuis gparams
        gp = state_manager.gparams
        self._title_edit.setText(getattr(gp, 'doc_title', ''))
        self._author_edit.setText(getattr(gp, 'doc_author', ''))
        self._building = False

    # ── Actions ─────────────────────────────────────────────────────────────
    def _on_years_changed(self, value: int):
        if not self._building:
            state_manager.state.gparams.__dict__['start_year'] = value

    def _open_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Ouvrir un fichier Excel",
            str(dossier_courant), "Fichiers Excel (*.xlsx *.xlsm *.xls)"
        )
        if path:
            state_manager.load_excel(path)

    def _on_include_toggled(self, sheet_name: str, included: bool):
        if not self._building:
            state_manager.toggle_sheet_include(sheet_name, included)

    def _on_order_changed(self, names: list[str]):
        if not self._building:
            state_manager.update_sheet_order(names)

    def _on_title_changed(self, text: str):
        if not self._building:
            state_manager.state.gparams.__dict__.setdefault('doc_title', text)
            state_manager.state.gparams.doc_title = text

    def _on_author_changed(self, text: str):
        if not self._building:
            state_manager.state.gparams.__dict__['doc_author'] = text

    def _select_all(self):
        wb = state_manager.workbook
        if not wb:
            return
        for sheet in wb.sheets:
            state_manager.toggle_sheet_include(sheet.name, True)
        self._refresh(None)

    def _select_none(self):
        wb = state_manager.workbook
        if not wb:
            return
        for sheet in wb.sheets:
            state_manager.toggle_sheet_include(sheet.name, False)
        self._refresh(None)