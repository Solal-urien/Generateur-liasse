"""
Onglet Lignes (tab_line) :
  Trois menus déroulants : Feuille → Tableau → Ligne
  Panneau d'édition de la ligne sélectionnée :
    - Visible / masquée
    - Couleur de fond (avec adaptation automatique de la couleur du texte)
    - Style texte : normal / bold / italic
    - Taille de police
    - Format nombre : normal / pourcentage
    - Décimales (0..3)

  Priorité : le RowStyle défini ici est appliqué EN PREMIER dans pdf_table_maker,
  devant tout paramètre par défaut de la page.
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QCheckBox, QDoubleSpinBox, QSpinBox, QComboBox,
    QSizePolicy, QScrollArea, QFrame, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QBrush

from state_manager import state_manager
from pdf_generator import RowStyle
# ColorButton est défini dans tab_style ; on l'importe directement
from ui.tab_style import ColorButton

_DEFAULT_BG = "#FFFFFF"


class LineTab(QWidget):
    """Onglet permettant de styliser chaque ligne d'un tableau individuellement."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._loading         = False
        self._current_sheet:  str | None = None
        self._current_table:  str | None = None
        self._current_row_idx: int | None = None   # index dans TableZone.data
        self._build_ui()
        state_manager.subscribe("workbook_loaded", self._on_workbook_loaded)
        state_manager.subscribe("profile_loaded",  self._on_workbook_loaded)

    # ── Construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        inner  = QWidget()
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # ── Sélecteurs en cascade ─────────────────────────────────────────
        grp_sel = QGroupBox("Sélection")
        sel_layout = QVBoxLayout(grp_sel)
        sel_layout.setSpacing(6)

        def _combo_row(label: str, combo_attr: str, slot) -> QComboBox:
            row = QHBoxLayout()
            row.addWidget(QLabel(label))
            combo = QComboBox()
            combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            combo.currentTextChanged.connect(slot)
            setattr(self, combo_attr, combo)
            row.addWidget(combo)
            sel_layout.addLayout(row)
            return combo

        _combo_row("Feuille :",  "_combo_sheet", self._on_sheet_selected)
        _combo_row("Tableau :",  "_combo_table", self._on_table_selected)
        layout.addWidget(grp_sel)

        # ── Tableau de résumé des lignes ──────────────────────────────────
        grp_rows = QGroupBox("Lignes du tableau")
        rows_layout = QVBoxLayout(grp_rows)

        self._row_table = QTableWidget()
        self._row_table.setColumnCount(6)
        self._row_table.setHorizontalHeaderLabels(
            ["Libellé", "Visible", "Fond", "Style", "Taille", "Format"]
        )
        hh = self._row_table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for c in range(1, 6):
            hh.setSectionResizeMode(c, QHeaderView.ResizeMode.ResizeToContents)
        self._row_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._row_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._row_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._row_table.setMinimumHeight(160)
        self._row_table.itemChanged.connect(self._on_vis_checkbox_changed)
        self._row_table.selectionModel().selectionChanged.connect(self._on_row_selected)
        rows_layout.addWidget(self._row_table)
        layout.addWidget(grp_rows)

        # ── Panneau d'édition ─────────────────────────────────────────────
        grp_edit = QGroupBox("Paramètres de la ligne sélectionnée")
        edit_layout = QVBoxLayout(grp_edit)
        edit_layout.setSpacing(8)

        # Visible
        self._chk_visible = QCheckBox("Ligne visible dans le PDF")
        self._chk_visible.setChecked(True)
        self._chk_visible.stateChanged.connect(self._push)
        edit_layout.addWidget(self._chk_visible)

        # Couleur de fond
        row_bg = QHBoxLayout()
        self._chk_bg = QCheckBox("Couleur de fond :")
        self._btn_bg = ColorButton(_DEFAULT_BG)
        self._btn_bg.setEnabled(False)
        self._chk_bg.stateChanged.connect(lambda s: (
            self._btn_bg.setEnabled(bool(s)), self._push()
        ))
        self._btn_bg.clicked.connect(self._push)
        row_bg.addWidget(self._chk_bg)
        row_bg.addStretch()
        row_bg.addWidget(self._btn_bg)
        edit_layout.addLayout(row_bg)

        info_bg = QLabel("(le texte passe automatiquement en blanc sur fond sombre)")
        info_bg.setStyleSheet("color: #5E5E5E; font-size: 10px; font-style: italic;")
        edit_layout.addWidget(info_bg)

        # Style texte
        row_ts = QHBoxLayout()
        row_ts.addWidget(QLabel("Style du texte :"))
        self._combo_text_style = QComboBox()
        self._combo_text_style.addItems(["normal", "bold", "italic"])
        self._combo_text_style.currentIndexChanged.connect(self._push)
        row_ts.addStretch()
        row_ts.addWidget(self._combo_text_style)
        edit_layout.addLayout(row_ts)

        # Taille police
        row_fs = QHBoxLayout()
        self._chk_font = QCheckBox("Taille de police :")
        self._spin_font = QDoubleSpinBox()
        self._spin_font.setRange(5, 28)
        self._spin_font.setValue(9.0)
        self._spin_font.setSuffix(" pt")
        self._spin_font.setSingleStep(0.5)
        self._spin_font.setEnabled(False)
        self._chk_font.stateChanged.connect(lambda s: (
            self._spin_font.setEnabled(bool(s)), self._push()
        ))
        self._spin_font.valueChanged.connect(self._push)
        row_fs.addWidget(self._chk_font)
        row_fs.addStretch()
        row_fs.addWidget(self._spin_font)
        edit_layout.addLayout(row_fs)

        # Format nombre
        row_fmt = QHBoxLayout()
        row_fmt.addWidget(QLabel("Format des nombres :"))
        self._combo_fmt = QComboBox()
        self._combo_fmt.addItems(["normal", "percent"])
        self._combo_fmt.currentIndexChanged.connect(self._push)
        row_fmt.addStretch()
        row_fmt.addWidget(self._combo_fmt)
        edit_layout.addLayout(row_fmt)

        # Décimales
        row_dec = QHBoxLayout()
        row_dec.addWidget(QLabel("Décimales (0–3) :"))
        self._spin_dec = QSpinBox()
        self._spin_dec.setRange(0, 3)
        self._spin_dec.setValue(1)
        self._spin_dec.valueChanged.connect(self._push)
        row_dec.addStretch()
        row_dec.addWidget(self._spin_dec)
        edit_layout.addLayout(row_dec)

        # Boutons reset
        btn_bar = QHBoxLayout()
        btn_apply_to_all = QPushButton("Appliquer à toutes les lignes")
        btn_apply_to_all.setObjectName("secondary")
        btn_apply_to_all.clicked.connect(self._apply_to_all_row)
        btn_reset_all = QPushButton("Réinitialiser toutes les lignes")
        btn_reset_all.setObjectName("secondary")
        btn_reset_all.clicked.connect(self._reset_all)
        btn_bar.addWidget(btn_apply_to_all)
        btn_bar.addWidget(btn_reset_all)
        edit_layout.addLayout(btn_bar)

        layout.addWidget(grp_edit)
        layout.addStretch()

        self._set_edit_enabled(False)

        scroll.setWidget(inner)
        root.addWidget(scroll)

    # ── Helpers ─────────────────────────────────────────────────────────────

    def _set_edit_enabled(self, enabled: bool):
        for w in [self._chk_visible, self._chk_bg, self._combo_text_style,
                  self._chk_font, self._combo_fmt, self._spin_dec]:
            w.setEnabled(enabled)
        # Les boutons conditionnels suivent leur checkbox
        self._btn_bg.setEnabled(enabled and self._chk_bg.isChecked())
        self._spin_font.setEnabled(enabled and self._chk_font.isChecked())

    def _get_tsp(self):
        """Retourne le TableStyleParams courant (crée si absent)."""
        if not self._current_sheet or not self._current_table:
            return None
        return state_manager.get_table_style(self._current_sheet, self._current_table)

    def _get_data(self) -> list[list] | None:
        wb = state_manager.workbook
        if not wb or not self._current_sheet or not self._current_table:
            return None
        sheet = next((s for s in wb.sheets if s.name == self._current_sheet), None)
        if not sheet:
            return None
        tz = next((t for t in sheet.tables if t.name == self._current_table), None)
        return tz.data if tz else None

    # ── Sync : workbook chargé ────────────────────────────────────────────

    def _on_workbook_loaded(self, _=None):
        wb = state_manager.workbook
        if not wb:
            return
        self._loading = True
        cur_sheet = self._combo_sheet.currentText()
        self._combo_sheet.blockSignals(True)
        self._combo_sheet.clear()
        for s in sorted(wb.sheets, key=lambda s: s.page_order):
            self._combo_sheet.addItem(s.name)
        idx = self._combo_sheet.findText(cur_sheet)
        self._combo_sheet.setCurrentIndex(max(0, idx))
        self._combo_sheet.blockSignals(False)
        self._loading = False
        self._on_sheet_selected(self._combo_sheet.currentText())

    # ── Cascade de sélection ──────────────────────────────────────────────

    def _on_sheet_selected(self, sheet_name: str):
        if not sheet_name or self._loading:
            return
        self._current_sheet = sheet_name
        self._loading = True
        cur_table = self._combo_table.currentText()
        self._combo_table.blockSignals(True)
        self._combo_table.clear()
        wb = state_manager.workbook
        if wb:
            sheet = next((s for s in wb.sheets if s.name == sheet_name), None)
            if sheet:
                for t in sheet.tables:
                    self._combo_table.addItem(t.name)
        idx = self._combo_table.findText(cur_table)
        self._combo_table.setCurrentIndex(max(0, idx))
        self._combo_table.blockSignals(False)
        self._loading = False
        self._on_table_selected(self._combo_table.currentText())

    def _on_table_selected(self, table_name: str):
        if self._loading:
            return
        self._current_table   = table_name if table_name else None
        self._current_row_idx = None
        self._refresh_row_table()
        self._set_edit_enabled(False)

    # ── Tableau de résumé ─────────────────────────────────────────────────

    def _refresh_row_table(self):
        self._loading = True
        self._row_table.blockSignals(True)
        self._row_table.clearContents()

        data = self._get_data()
        if data is None:
            self._row_table.setRowCount(0)
            self._row_table.blockSignals(False)
            self._loading = False
            return

        tsp = self._get_tsp()
        self._row_table.setRowCount(len(data))

        for ri, row in enumerate(data):
            rs = tsp.row_styles.get(ri, RowStyle()) if tsp else RowStyle()

            # Col 0 : libellé
            label = str(row[0]).strip() if row and row[0] is not None else ""
            if not label or label.lower() == "none":
                label = f"Ligne {ri + 1}"
            item_lbl = QTableWidgetItem(label)
            item_lbl.setFlags(item_lbl.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self._row_table.setItem(ri, 0, item_lbl)

            # Col 1 : visible (checkbox)
            item_vis = QTableWidgetItem()
            item_vis.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            item_vis.setCheckState(
                Qt.CheckState.Checked if rs.visible else Qt.CheckState.Unchecked
            )
            self._row_table.setItem(ri, 1, item_vis)

            # Col 2 : couleur de fond
            bg_txt  = rs.background_color or "—"
            item_bg = QTableWidgetItem(bg_txt)
            item_bg.setFlags(item_bg.flags() & ~Qt.ItemFlag.ItemIsEditable)
            if rs.background_color:
                item_bg.setBackground(QBrush(QColor(rs.background_color)))
            self._row_table.setItem(ri, 2, item_bg)

            # Col 3 : style texte
            item_ts = QTableWidgetItem(rs.text_style)
            item_ts.setFlags(item_ts.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self._row_table.setItem(ri, 3, item_ts)

            # Col 4 : taille
            fs_txt  = f"{rs.font_size} pt" if rs.font_size is not None else "—"
            item_fs = QTableWidgetItem(fs_txt)
            item_fs.setFlags(item_fs.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self._row_table.setItem(ri, 4, item_fs)

            # Col 5 : format
            fmt_txt = rs.number_format
            item_fmt = QTableWidgetItem(fmt_txt)
            item_fmt.setFlags(item_fmt.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self._row_table.setItem(ri, 5, item_fmt)

        self._row_table.blockSignals(False)
        self._loading = False

    def _on_vis_checkbox_changed(self, item: QTableWidgetItem):
        """Modification directe de la checkbox 'visible' dans le tableau."""
        if self._loading or item.column() != 1:
            return
        ri  = item.row()
        vis = item.checkState() == Qt.CheckState.Checked
        tsp = self._get_tsp()
        if tsp is None:
            return
        rs = tsp.row_styles.get(ri, RowStyle())
        rs.visible = vis
        tsp.row_styles[ri] = rs
        state_manager.update_table_style(self._current_sheet, self._current_table, tsp)

    def _on_row_selected(self, *_):
        """Charge les paramètres de la ligne sélectionnée dans le panneau d'édition."""
        items = self._row_table.selectedItems()
        if not items:
            self._current_row_idx = None
            self._set_edit_enabled(False)
            return
        ri  = self._row_table.currentRow()
        self._current_row_idx = ri

        tsp = self._get_tsp()
        rs  = tsp.row_styles.get(ri, RowStyle()) if tsp else RowStyle()

        self._loading = True

        self._chk_visible.setChecked(rs.visible)

        has_bg = bool(rs.background_color)
        self._chk_bg.setChecked(has_bg)
        self._btn_bg.setEnabled(has_bg)
        self._btn_bg.set_color(rs.background_color if has_bg else _DEFAULT_BG)

        idx_ts = self._combo_text_style.findText(rs.text_style)
        self._combo_text_style.setCurrentIndex(max(0, idx_ts))

        has_fs = rs.font_size is not None
        self._chk_font.setChecked(has_fs)
        self._spin_font.setEnabled(has_fs)
        self._spin_font.setValue(rs.font_size if has_fs else 9.0)

        idx_fmt = self._combo_fmt.findText(rs.number_format)
        self._combo_fmt.setCurrentIndex(max(0, idx_fmt))

        self._spin_dec.setValue(rs.decimal_places)

        self._loading = False
        self._set_edit_enabled(True)

    # ── Push vers l'état ──────────────────────────────────────────────────

    def _push(self, *_):
        if self._loading or self._current_row_idx is None:
            return
        tsp = self._get_tsp()
        if tsp is None:
            return
        ri = self._current_row_idx

        rs = RowStyle(
            visible          = self._chk_visible.isChecked(),
            background_color = self._btn_bg.color() if self._chk_bg.isChecked() else None,
            text_style       = self._combo_text_style.currentText(),
            font_size        = self._spin_font.value() if self._chk_font.isChecked() else None,
            number_format    = self._combo_fmt.currentText(),
            decimal_places   = self._spin_dec.value(),
        )
        tsp.row_styles[ri] = rs
        state_manager.update_table_style(self._current_sheet, self._current_table, tsp)

        # Mise à jour de la ligne dans le tableau de résumé
        self._loading = True
        self._update_summary_row(ri, rs)
        self._loading = False

    def _update_summary_row(self, ri: int, rs: RowStyle):
        if ri >= self._row_table.rowCount():
            return
        # Col 1 : visible
        item = self._row_table.item(ri, 1)
        if item:
            item.setCheckState(Qt.CheckState.Checked if rs.visible else Qt.CheckState.Unchecked)
        # Col 2 : fond
        item = self._row_table.item(ri, 2)
        if item:
            item.setText(rs.background_color or "—")
            if rs.background_color:
                item.setBackground(QBrush(QColor(rs.background_color)))
            else:
                item.setBackground(QBrush())
        # Col 3 : style
        item = self._row_table.item(ri, 3)
        if item:
            item.setText(rs.text_style)
        # Col 4 : taille
        item = self._row_table.item(ri, 4)
        if item:
            item.setText(f"{rs.font_size} pt" if rs.font_size is not None else "—")
        # Col 5 : format
        item = self._row_table.item(ri, 5)
        if item:
            fmt = rs.number_format + (f" {rs.decimal_places}d" if rs.number_format == "percent" else "")
            item.setText(fmt)

    # ── Reset ─────────────────────────────────────────────────────────────

    def _reset_row(self):
        ri = self._current_row_idx
        if ri is None:
            return
        tsp = self._get_tsp()
        if tsp and ri in tsp.row_styles:
            del tsp.row_styles[ri]
            state_manager.update_table_style(self._current_sheet, self._current_table, tsp)
        # Recharge le panneau avec les valeurs par défaut
        self._on_row_selected()
        self._loading = True
        self._update_summary_row(ri, RowStyle())
        self._loading = False

    def _reset_all(self):
        tsp = self._get_tsp()
        if tsp:
            tsp.row_styles.clear()
            state_manager.update_table_style(self._current_sheet, self._current_table, tsp)
        self._refresh_row_table()
        self._set_edit_enabled(False)

    def _apply_to_all_row(self):
        ri = self._current_row_idx
        if ri is None:
            return
        tsp = self._get_tsp()
        if tsp is None:
            return

        rs = RowStyle(
            visible          = self._chk_visible.isChecked(),
            background_color = self._btn_bg.color() if self._chk_bg.isChecked() else None,
            text_style       = self._combo_text_style.currentText(),
            font_size        = self._spin_font.value() if self._chk_font.isChecked() else None,
            number_format    = self._combo_fmt.currentText(),
            decimal_places   = self._spin_dec.value(),
        )

        for row_idx in range(self._row_table.rowCount()):
            tsp.row_styles[row_idx] = rs

        state_manager.update_table_style(self._current_sheet, self._current_table, tsp)

        self._loading = True
        for row_idx in range(self._row_table.rowCount()):
            self._update_summary_row(row_idx, rs)
        self._loading = False