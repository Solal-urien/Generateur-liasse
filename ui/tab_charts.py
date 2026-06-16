"""
Onglet Graphiques :
  - Liste des graphiques configurés
  - Formulaire d'ajout / édition (type, feuille, colonnes, palette, titre…)
  - Suppression
"""
from __future__ import annotations
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QListWidget, QListWidgetItem, QComboBox,
    QLineEdit, QSpinBox, QCheckBox,
    QScrollArea, QSizePolicy, QFrame, QMessageBox,
)
from PyQt6.QtCore import Qt

from state_manager import state_manager
from pdf_generator import ChartSpec

import util

CHART_TYPES = ["bar", "line", "area", "pie", "scatter"]
DEFAULT_PALETTE = util.NUANCIER_COULEURS["Graphiques"]

class ChartsTab(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self._edit_index: int | None = None  # None = nouveau, int = édition
        self._build_ui()
        state_manager.subscribe("workbook_loaded", self._on_workbook_loaded)
        state_manager.subscribe("charts_changed", self._refresh_list)
        state_manager.subscribe("profile_loaded", self._on_workbook_loaded)

    # ── Construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # ── Liste des graphiques ──
        grp_list = QGroupBox("Graphiques configurés")
        list_layout = QVBoxLayout(grp_list)

        self._chart_list = QListWidget()
        self._chart_list.setMaximumHeight(130)
        self._chart_list.itemSelectionChanged.connect(self._on_chart_selected)
        list_layout.addWidget(self._chart_list)

        btn_row = QHBoxLayout()
        self._btn_edit = QPushButton("✏ Modifier")
        self._btn_edit.setObjectName("secondary")
        self._btn_edit.setEnabled(False)
        self._btn_edit.clicked.connect(self._load_selected_for_edit)
        self._btn_del = QPushButton("🗑 Supprimer")
        self._btn_del.setObjectName("danger")
        self._btn_del.setEnabled(False)
        self._btn_del.clicked.connect(self._delete_selected)
        btn_row.addWidget(self._btn_edit)
        btn_row.addWidget(self._btn_del)
        btn_row.addStretch()
        list_layout.addLayout(btn_row)
        root.addWidget(grp_list)

        # ── Formulaire ──
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        form_widget = QWidget()
        form_layout = QVBoxLayout(form_widget)
        form_layout.setContentsMargins(0, 0, 0, 0)
        form_layout.setSpacing(10)

        # Titre formulaire
        self._form_title_lbl = QLabel("➕ Nouveau graphique")
        self._form_title_lbl.setStyleSheet(
            "font-size: 13px; font-weight: 700; color: #1a3a5c; padding: 4px 0;"
        )
        form_layout.addWidget(self._form_title_lbl)

        grp_source = QGroupBox("Source de données")
        src_layout = QVBoxLayout(grp_source)
        src_layout.setSpacing(8)

        # Feuille
        row_sheet = QHBoxLayout()
        row_sheet.addWidget(QLabel("Feuille :"))
        self._combo_sheet = QComboBox()
        self._combo_sheet.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._combo_sheet.currentTextChanged.connect(self._on_sheet_selected)
        row_sheet.addWidget(self._combo_sheet)
        src_layout.addLayout(row_sheet)

        # Zone / tableau
        row_table = QHBoxLayout()
        row_table.addWidget(QLabel("Zone :"))
        self._combo_table = QComboBox()
        self._combo_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        row_table.addWidget(self._combo_table)
        src_layout.addLayout(row_table)

        # Colonne X
        row_x = QHBoxLayout()
        row_x.addWidget(QLabel("Colonne X (index) :"))
        self._spin_x = QSpinBox()
        self._spin_x.setRange(0, 20) # TODO : limiter au nombre de colonnes de la table sélectionnée
        self._spin_x.setValue(0)
        row_x.addStretch()
        row_x.addWidget(self._spin_x)
        src_layout.addLayout(row_x)

        # Colonnes Y
        row_y = QHBoxLayout()
        row_y.addWidget(QLabel("Colonnes Y (ex: 1,2,3) :"))
        self._edit_y = QLineEdit("1")
        self._edit_y.setPlaceholderText("1")
        row_y.addWidget(self._edit_y)
        src_layout.addLayout(row_y)

        form_layout.addWidget(grp_source)

        # ── Type & apparence ──
        grp_appearance = QGroupBox("Type & apparence")
        app_layout = QVBoxLayout(grp_appearance)
        app_layout.setSpacing(8)

        row_type = QHBoxLayout()
        row_type.addWidget(QLabel("Type :"))
        self._combo_type = QComboBox()
        self._combo_type.addItems(CHART_TYPES)
        row_type.addStretch()
        row_type.addWidget(self._combo_type)
        app_layout.addLayout(row_type)

        row_title = QHBoxLayout()
        row_title.addWidget(QLabel("Titre :"))
        self._edit_title = QLineEdit()
        self._edit_title.setPlaceholderText("Titre du graphique (optionnel)")
        row_title.addWidget(self._edit_title)
        app_layout.addLayout(row_title)

        row_xlabel = QHBoxLayout()
        row_xlabel.addWidget(QLabel("Axe X :"))
        self._edit_xlabel = QLineEdit()
        self._edit_xlabel.setPlaceholderText("Label axe X")
        row_xlabel.addWidget(self._edit_xlabel)
        app_layout.addLayout(row_xlabel)

        row_ylabel = QHBoxLayout()
        row_ylabel.addWidget(QLabel("Axe Y :"))
        self._edit_ylabel = QLineEdit()
        self._edit_ylabel.setPlaceholderText("Label axe Y")
        row_ylabel.addWidget(self._edit_ylabel)
        app_layout.addLayout(row_ylabel)

        row_palette = QHBoxLayout()
        row_palette.addWidget(QLabel("Palette (hex, virgule) :"))
        self._edit_palette = QLineEdit(",".join(DEFAULT_PALETTE[:5]))
        self._edit_palette.setPlaceholderText("#2e7bc4,#e05c1a,…")
        row_palette.addWidget(self._edit_palette)
        app_layout.addLayout(row_palette)

        row_legend = QHBoxLayout()
        self._chk_legend = QCheckBox("Afficher la légende")
        self._chk_legend.setChecked(True)
        row_legend.addWidget(self._chk_legend)
        app_layout.addLayout(row_legend)

        form_layout.addWidget(grp_appearance)

        # ── Boutons formulaire ──
        btn_form_row = QHBoxLayout()
        self._btn_cancel = QPushButton("Annuler")
        self._btn_cancel.setObjectName("secondary")
        self._btn_cancel.clicked.connect(self._reset_form)
        self._btn_save = QPushButton("✅ Enregistrer")
        self._btn_save.clicked.connect(self._save_chart)
        btn_form_row.addWidget(self._btn_cancel)
        btn_form_row.addStretch()
        btn_form_row.addWidget(self._btn_save)
        form_layout.addLayout(btn_form_row)
        form_layout.addStretch()

        scroll.setWidget(form_widget)
        root.addWidget(scroll)

    # ── Helpers ─────────────────────────────────────────────────────────────

    def _refresh_list(self, charts=None):
        charts = charts or state_manager.charts
        self._chart_list.clear()
        for i, cs in enumerate(charts):
            label = f"{cs.chart_type.upper()} — {cs.sheet_name}/{cs.table_name}"
            if cs.title:
                label += f" « {cs.title} »"
            item = QListWidgetItem(label)
            item.setData(Qt.ItemDataRole.UserRole, i)
            self._chart_list.addItem(item)

    def _on_workbook_loaded(self, _=None):
        wb = state_manager.workbook
        if not wb:
            return
        current_sheet = self._combo_sheet.currentText()
        self._combo_sheet.blockSignals(True)
        self._combo_sheet.clear()
        for s in sorted(wb.sheets, key=lambda x: x.page_order):
            self._combo_sheet.addItem(s.name)
        idx = self._combo_sheet.findText(current_sheet)
        self._combo_sheet.setCurrentIndex(max(0, idx))
        self._combo_sheet.blockSignals(False)
        self._on_sheet_selected(self._combo_sheet.currentText())
        self._refresh_list()

    def _on_sheet_selected(self, sheet_name: str):
        wb = state_manager.workbook
        if not wb or not sheet_name:
            return
        sheet = next((s for s in wb.sheets if s.name == sheet_name), None)
        self._combo_table.clear()
        if sheet:
            for t in sheet.tables:
                self._combo_table.addItem(t.name)

    def _on_chart_selected(self):
        has_sel = bool(self._chart_list.selectedItems())
        self._btn_edit.setEnabled(has_sel)
        self._btn_del.setEnabled(has_sel)

    def _load_selected_for_edit(self):
        items = self._chart_list.selectedItems()
        if not items:
            return
        idx = items[0].data(Qt.ItemDataRole.UserRole)
        charts = state_manager.charts
        if idx >= len(charts):
            return
        cs = charts[idx]
        self._edit_index = idx
        self._form_title_lbl.setText("✏ Modifier le graphique")

        # Remplir le formulaire
        sheet_idx = self._combo_sheet.findText(cs.sheet_name)
        if sheet_idx >= 0:
            self._combo_sheet.setCurrentIndex(sheet_idx)
        table_idx = self._combo_table.findText(cs.table_name)
        if table_idx >= 0:
            self._combo_table.setCurrentIndex(table_idx)
        type_idx = self._combo_type.findText(cs.chart_type)
        if type_idx >= 0:
            self._combo_type.setCurrentIndex(type_idx)
        self._spin_x.setValue(cs.x_col)
        self._edit_y.setText(",".join(str(y) for y in cs.y_cols))
        self._edit_title.setText(cs.title)
        self._edit_xlabel.setText(cs.xlabel)
        self._edit_ylabel.setText(cs.ylabel)
        self._edit_palette.setText(",".join(cs.palette) if cs.palette else "")
        self._chk_legend.setChecked(cs.legend)

    def _delete_selected(self):
        items = self._chart_list.selectedItems()
        if not items:
            return
        idx = items[0].data(Qt.ItemDataRole.UserRole)
        reply = QMessageBox.question(
            self, "Supprimer", "Supprimer ce graphique ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            state_manager.remove_chart(idx)
            self._reset_form()

    def _save_chart(self):
        sheet_name = self._combo_sheet.currentText()
        table_name = self._combo_table.currentText()
        if not sheet_name or not table_name:
            QMessageBox.warning(self, "Incomplet", "Sélectionnez une feuille et une zone.")
            return

        y_cols_raw = self._edit_y.text().strip()
        try:
            y_cols = [int(x.strip()) for x in y_cols_raw.split(",") if x.strip()]
        except ValueError:
            QMessageBox.warning(self, "Erreur", "Les colonnes Y doivent être des entiers séparés par des virgules.")
            return

        palette_raw = self._edit_palette.text().strip()
        palette = [p.strip() for p in palette_raw.split(",") if p.strip()] if palette_raw else []

        cs = ChartSpec(
            sheet_name=sheet_name,
            table_name=table_name,
            chart_type=self._combo_type.currentText(),
            x_col=self._spin_x.value(),
            y_cols=y_cols or [1],
            title=self._edit_title.text().strip(),
            xlabel=self._edit_xlabel.text().strip(),
            ylabel=self._edit_ylabel.text().strip(),
            palette=palette,
            legend=self._chk_legend.isChecked(),
        )

        if self._edit_index is not None:
            state_manager.update_chart(self._edit_index, cs)
        else:
            state_manager.add_chart(cs)

        self._reset_form()

    def _reset_form(self):
        self._edit_index = None
        self._form_title_lbl.setText("➕ Nouveau graphique")
        self._edit_title.clear()
        self._edit_xlabel.clear()
        self._edit_ylabel.clear()
        self._edit_y.setText("1")
        self._spin_x.setValue(0)
        self._combo_type.setCurrentIndex(0)
        self._edit_palette.setText(",".join(DEFAULT_PALETTE[:5]))
        self._chk_legend.setChecked(True)
        self._chart_list.clearSelection()
        self._btn_edit.setEnabled(False)
        self._btn_del.setEnabled(False)