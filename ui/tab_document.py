"""
Onglet Document :
  - Chargement du fichier Excel
  - Réorganisation des feuilles par drag & drop
  - Inclusion / exclusion sélective
  - Modification du titre et du sous-titre de page de garde
  - Renommage de l'affichage de chaque page (display_name)
  - Note de bas de page par feuille (footer_note)
"""
from pathlib import Path

from PyQt6.QtWidgets import (
    QSpinBox, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QCheckBox, QGroupBox,
    QLineEdit, QFileDialog, QAbstractItemView, QSizePolicy,
    QFrame, QScrollArea,
)
from PyQt6.QtCore import Qt, pyqtSignal

from state_manager import state_manager

dossier_courant = Path(__file__).resolve().parent.parent.parent


# ── Widget d'une feuille dans la liste ────────────────────────────────────

class SheetListItem(QWidget):
    """Widget représentant une feuille dans la liste (checkbox + nom + indicateur)."""

    toggled = pyqtSignal(str, bool)

    def __init__(self, sheet_name: str, display_name: str | None,
                 included: bool, has_chart: bool, parent=None):
        super().__init__(parent)
        self.sheet_name = sheet_name
        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 3, 6, 3)
        layout.setSpacing(8)

        handle = QLabel("⠿")
        handle.setStyleSheet("color: #aabbcc; font-size: 16px;")
        handle.setFixedWidth(16)
        layout.addWidget(handle)

        self.chk = QCheckBox()
        self.chk.setChecked(included)
        self.chk.stateChanged.connect(self._on_toggle)
        layout.addWidget(self.chk)

        # Nom affiché : display_name si défini, sinon nom original
        shown = display_name if display_name else sheet_name
        self._name_lbl = QLabel(shown)
        if display_name:
            self._name_lbl.setStyleSheet("color: #385188; font-style: italic;")
        self._name_lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self._name_lbl.setToolTip(f"Nom Excel : {sheet_name}" if display_name else sheet_name)
        layout.addWidget(self._name_lbl)

        if has_chart:
            badge = QLabel("📊")
            badge.setToolTip("Contient un graphique")
            layout.addWidget(badge)

    def _on_toggle(self, state: int):
        self.toggled.emit(self.sheet_name, state == Qt.CheckState.Checked.value)


class DraggableList(QListWidget):
    """QListWidget avec drag & drop interne pour réordonner."""

    order_changed = pyqtSignal(list)

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


# ── Onglet Document ────────────────────────────────────────────────────────

class DocumentTab(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self._building = False
        self._current_sheet: str | None = None
        self._build_ui()
        state_manager.subscribe("workbook_loaded",       self._on_workbook_loaded)
        state_manager.subscribe("profile_loaded",        self._refresh)
        state_manager.subscribe("sheet_settings_changed", self._on_sheet_settings_changed)

    # ── Construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        inner = QWidget()
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        # ── Fichier source ──
        grp_file = QGroupBox("Fichier source")
        file_layout = QHBoxLayout(grp_file)
        self._lbl_file = QLabel("Aucun fichier chargé")
        self._lbl_file.setStyleSheet("color: #667788; font-style: italic;")
        self._lbl_file.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        btn_open = QPushButton("📂 Ouvrir")
        btn_open.clicked.connect(self._open_excel)
        file_layout.addWidget(self._lbl_file)
        file_layout.addWidget(btn_open)
        layout.addWidget(grp_file)

        # ── Années ──
        grp_years = QGroupBox("Années considérées")
        years_layout = QHBoxLayout(grp_years)
        years_layout.addWidget(QLabel("Année début affichage (ré-ouvrir l'excel pour appliquer) :"))
        self._spin_year = QSpinBox()
        self._spin_year.setRange(2020, 2030)
        self._spin_year.setValue(2025)
        self._spin_year.setSingleStep(1)
        self._spin_year.setFixedWidth(90)
        self._spin_year.valueChanged.connect(self._on_years_changed)
        years_layout.addWidget(self._spin_year)
        layout.addWidget(grp_years)

        # ── Liste des feuilles ──
        grp_sheets = QGroupBox("Feuilles (glisser pour réordonner)")
        sheets_layout = QVBoxLayout(grp_sheets)

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
        self._list.setMinimumHeight(180)
        self._list.order_changed.connect(self._on_order_changed)
        self._list.currentItemChanged.connect(self._on_list_selection_changed)
        sheets_layout.addWidget(self._list)

        self._lbl_empty = QLabel("Chargez un fichier Excel pour\nafficher les feuilles.")
        self._lbl_empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._lbl_empty.setStyleSheet("color: #2F2F2F; font-style: italic;")
        sheets_layout.addWidget(self._lbl_empty)
        layout.addWidget(grp_sheets)

        # ── Paramètres de la page sélectionnée ──
        grp_page = QGroupBox("Paramètres de la page sélectionnée")
        page_layout = QVBoxLayout(grp_page)
        page_layout.setSpacing(8)

        # Indicateur de page sélectionnée
        self._lbl_selected = QLabel("Aucune page sélectionnée")
        self._lbl_selected.setStyleSheet("color: #385188; font-weight: 600; font-size: 11px;")
        page_layout.addWidget(self._lbl_selected)

        # Nom affiché
        row_disp = QHBoxLayout()
        row_disp.addWidget(QLabel("Nom affiché (titre + sommaire) :"))
        self._edit_display_name = QLineEdit()
        self._edit_display_name.setPlaceholderText("Laisser vide pour utiliser le nom Excel")
        self._edit_display_name.textChanged.connect(self._on_display_name_changed)
        row_disp.addWidget(self._edit_display_name)
        page_layout.addLayout(row_disp)

        # Note de bas de page
        row_note = QHBoxLayout()
        row_note.addWidget(QLabel("Note de bas de page :"))
        self._edit_footer_note = QLineEdit()
        self._edit_footer_note.setPlaceholderText("Laissez vide pour aucune note")
        self._edit_footer_note.textChanged.connect(self._on_footer_note_changed)
        row_note.addWidget(self._edit_footer_note)
        page_layout.addLayout(row_note)

        self._set_page_controls_enabled(False)
        layout.addWidget(grp_page)

        layout.addStretch()
        scroll.setWidget(inner)
        root.addWidget(scroll)

    # ── Helpers ─────────────────────────────────────────────────────────────

    def _set_page_controls_enabled(self, enabled: bool):
        self._edit_display_name.setEnabled(enabled)
        self._edit_footer_note.setEnabled(enabled)

    # ── Callbacks état ───────────────────────────────────────────────────────

    def _on_workbook_loaded(self, wb):
        path_exc = Path(state_manager.state.current_excel_path).name \
            if state_manager.state.current_excel_path else None
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
                getattr(t, "source", "") == "chart"
                for t in sheet.tables
            )
            ss = state_manager.get_sheet_settings(sheet.name)

            item = QListWidgetItem(self._list)
            item.setData(Qt.ItemDataRole.UserRole, sheet.name)

            widget = SheetListItem(
                sheet.name,
                ss.display_name,
                ss.include,
                has_chart,
            )
            widget.toggled.connect(self._on_include_toggled)
            item.setSizeHint(widget.sizeHint())
            self._list.setItemWidget(item, widget)

        self._building = False

        # Restaurer la sélection courante si elle existe encore
        if self._current_sheet:
            self._select_sheet_in_list(self._current_sheet)
        else:
            self._set_page_controls_enabled(False)

    def _on_sheet_settings_changed(self, sheet_name: str):
        """Met à jour le label affiché dans la liste pour la feuille modifiée."""
        for i in range(self._list.count()):
            item = self._list.item(i)
            w = self._list.itemWidget(item)
            if w and w.sheet_name == sheet_name:
                ss = state_manager.get_sheet_settings(sheet_name)
                shown = ss.display_name if ss.display_name else sheet_name
                w._name_lbl.setText(shown)
                if ss.display_name:
                    w._name_lbl.setStyleSheet("color: #385188; font-style: italic;")
                    w._name_lbl.setToolTip(f"Nom Excel : {sheet_name}")
                else:
                    w._name_lbl.setStyleSheet("")
                    w._name_lbl.setToolTip(sheet_name)
                break

    def _on_list_selection_changed(self, current, _previous):
        if current is None:
            self._current_sheet = None
            self._set_page_controls_enabled(False)
            self._lbl_selected.setText("Aucune page sélectionnée")
            return

        sheet_name = current.data(Qt.ItemDataRole.UserRole)
        if not sheet_name:
            return

        self._current_sheet = sheet_name
        self._lbl_selected.setText(f"Page : {sheet_name}")
        ss = state_manager.get_sheet_settings(sheet_name)

        self._building = True
        self._edit_display_name.setText(ss.display_name or "")
        self._edit_footer_note.setText(ss.footer_note or "")
        self._building = False

        self._set_page_controls_enabled(True)

    def _select_sheet_in_list(self, sheet_name: str):
        for i in range(self._list.count()):
            item = self._list.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == sheet_name:
                self._list.setCurrentItem(item)
                return

    # ── Actions ─────────────────────────────────────────────────────────────

    def _on_years_changed(self, value: int):
        if not self._building:
            state_manager.set_start_year(value)

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

    def _on_display_name_changed(self, text: str):
        if not self._building and self._current_sheet:
            state_manager.set_display_name(self._current_sheet, text.strip() or None)

    def _on_footer_note_changed(self, text: str):
        if not self._building and self._current_sheet:
            state_manager.set_footer_note(self._current_sheet, text.strip() or None)

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