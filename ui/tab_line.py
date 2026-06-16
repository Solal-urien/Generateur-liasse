"""
Onglet Lignes (tab_line) :
  Sélecteur feuille → sélecteur tableau → liste des lignes de données éditables.

  Pour chaque ligne on peut configurer :
    • Visible / masquée
    • Couleur de fond
    • Style texte : normal / gras / italique
    • Taille de police (None = hérite de PageParams)
    • Format des nombres : normal / pourcentage
    • Nombre de décimales (0 à 3)

  Règles spéciales :
    • La ligne "années" (subheader, détectée via detect_year_row_index) est
      affichée séparément, en lecture seule, avec un badge "Années" pour
      que l'utilisateur comprenne son rôle. Elle n'est PAS incluse dans
      "Appliquer à toutes les lignes".
    • "Appliquer à toutes les lignes" ne touche jamais au subheader années.
"""
from __future__ import annotations

import copy

from PyQt6.QtCore    import Qt, pyqtSignal
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QCheckBox, QDoubleSpinBox, QComboBox, QSpinBox,
    QSizePolicy, QScrollArea, QFrame)

from state_manager import state_manager
from pdf_generator  import RowStyle
from pdf_table_maker import detect_year_row_index
from util import ColorButton

WIDTH_COLUMN_WIDGET = [60, 35, 35, 70, 70, 45, 60]


# ── Widget d'une ligne éditable ────────────────────────────────────────────

class RowStyleWidget(QWidget):
    """
    Widget représentant les paramètres d'une ligne de données.
    Émet changed() à chaque modification.
    """
    changed = pyqtSignal()

    def __init__(self, row_index: int, label: str, rs: RowStyle | None = None, parent=None):
        super().__init__(parent)
        self.row_index = row_index
        self._loading  = False

        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 2, 4, 2)
        layout.setSpacing(6)

        # Libellé de la ligne (on tronque si trop long)
        name_lbl = QLabel(label)
        name_lbl.setFixedWidth(WIDTH_COLUMN_WIDGET[0])
        name_lbl.setToolTip(label)
        fm = name_lbl.fontMetrics()
        elided = fm.elidedText(label, Qt.TextElideMode.ElideRight, WIDTH_COLUMN_WIDGET[0]) 
        name_lbl.setText(elided)
        layout.addWidget(name_lbl)

        # Visible
        self._chk_visible = QCheckBox("")
        self._chk_visible.setChecked(True)
        self._chk_visible.setFixedWidth(WIDTH_COLUMN_WIDGET[1])
        layout.addWidget(self._chk_visible)

        # Couleur de fond
        self._btn_color = ColorButton("#FFFFFF")
        self._btn_color.setToolTip("Couleur de fond (blanc = défaut)")
        self._btn_color.setFixedWidth(WIDTH_COLUMN_WIDGET[2])
        layout.addWidget(self._btn_color)

        # Style texte
        self._combo_style = QComboBox()
        self._combo_style.addItems(["Normal", "Gras", "Italique"])
        self._combo_style.setFixedWidth(WIDTH_COLUMN_WIDGET[3])
        layout.addWidget(self._combo_style)

        # Taille police
        self._spin_size = QDoubleSpinBox()
        self._spin_size.setRange(5, 20)
        self._spin_size.setValue(9)
        self._spin_size.setSuffix("pt")
        self._spin_size.setSpecialValueText("N.A") # On hérite du global
        self._spin_size.setMinimum(0)          
        self._spin_size.setFixedWidth(WIDTH_COLUMN_WIDGET[4])
        layout.addWidget(self._spin_size)

        # Format numérique
        self._chk_percent = QCheckBox("")
        self._chk_percent.setToolTip("Afficher au format %")
        self._chk_percent.setFixedWidth(WIDTH_COLUMN_WIDGET[5])
        layout.addWidget(self._chk_percent)

        # Décimales
        self._spin_dec = QSpinBox()
        self._spin_dec.setRange(0, 3)
        self._spin_dec.setValue(1)
        self._spin_dec.setFixedWidth(WIDTH_COLUMN_WIDGET[6])
        self._spin_dec.setToolTip("Nombre de décimales")
        layout.addWidget(self._spin_dec)

        # Connexions
        self._chk_visible.stateChanged.connect(self._emit)
        self._btn_color.clicked.connect(self._emit)
        self._combo_style.currentIndexChanged.connect(self._emit)
        self._spin_size.valueChanged.connect(self._emit)
        self._chk_percent.stateChanged.connect(self._emit)
        self._spin_dec.valueChanged.connect(self._emit)

        if rs is not None:
            self.load(rs)

    # ── Chargement / lecture ───────────────────────────────────────────────

    def load(self, rs: RowStyle):
        self._loading = True
        self._chk_visible.setChecked(rs.visible)
        self._btn_color.set_color(rs.background_color or "#FFFFFF")
        style_map = {"normal": 0, "bold": 1, "italic": 2}
        self._combo_style.setCurrentIndex(style_map.get(rs.text_style, 0))
        self._spin_size.setValue(rs.font_size if rs.font_size is not None else 0)
        self._chk_percent.setChecked(rs.number_format == "percent")
        self._spin_dec.setValue(rs.decimal_places)
        self._loading = False

    def get_row_style(self) -> RowStyle:
        style_map = {0: "normal", 1: "bold", 2: "italic"}
        bg = self._btn_color.color()
        bg = None if bg.upper() == "#FFFFFF" else bg
        sz = self._spin_size.value()
        return RowStyle(
            visible          = self._chk_visible.isChecked(),
            background_color = bg,
            text_style       = style_map[self._combo_style.currentIndex()],
            font_size        = sz if sz > 0 else None,
            number_format    = "percent" if self._chk_percent.isChecked() else "normal",
            decimal_places   = self._spin_dec.value(),
        )

    def _emit(self, *_):
        if not self._loading:
            self.changed.emit()


# ── Onglet Lignes ──────────────────────────────────────────────────────────

class LineTab(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_sheet: str | None = None
        self._current_table: str | None = None
        self._year_row_index: int | None = None   # index dans data du subheader années
        self._row_widgets: list[tuple[int, RowStyleWidget]] = []  # (data_idx, widget)
        self._loading = False
        self._build_ui()
        state_manager.subscribe("workbook_loaded", self._on_workbook_loaded)
        state_manager.subscribe("profile_loaded",  self._on_workbook_loaded)

    # ── Construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        inner = QWidget()
        self._inner_layout = QVBoxLayout(inner)
        self._inner_layout.setContentsMargins(12, 12, 12, 12)
        self._inner_layout.setSpacing(10)

        # ── Sélecteurs ──
        grp_sel = QGroupBox("Sélection")
        sel_layout = QVBoxLayout(grp_sel)

        row_s = QHBoxLayout()
        row_s.addWidget(QLabel("Feuille :"))
        self._combo_sheet = QComboBox()
        self._combo_sheet.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._combo_sheet.currentTextChanged.connect(self._on_sheet_changed)
        row_s.addWidget(self._combo_sheet)
        sel_layout.addLayout(row_s)

        row_t = QHBoxLayout()
        row_t.addWidget(QLabel("Tableau :"))
        self._combo_table = QComboBox()
        self._combo_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._combo_table.currentTextChanged.connect(self._on_table_changed)
        row_t.addWidget(self._combo_table)
        sel_layout.addLayout(row_t)
        self._inner_layout.addWidget(grp_sel)

        # ── Info subheader années ──
        self._lbl_year_row = QLabel()
        self._lbl_year_row.setWordWrap(True)
        self._lbl_year_row.setStyleSheet(
            "background: #D2DAEC; color: #1C2844; border-radius: 4px; "
            "padding: 4px 8px; font-size: 11px; font-style: italic;"
        )
        self._lbl_year_row.setVisible(False)
        self._inner_layout.addWidget(self._lbl_year_row)

        # ── En-tête des colonnes ──
        hdr = QWidget()
        hdr_layout = QHBoxLayout(hdr)
        hdr_layout.setContentsMargins(4, 0, 4, 0)
        hdr_layout.setSpacing(6)
        for lbl, w in [("Libellé", WIDTH_COLUMN_WIDGET[0]), ("Affiché", WIDTH_COLUMN_WIDGET[1]), ("Fond", WIDTH_COLUMN_WIDGET[2]),
                ("Style", WIDTH_COLUMN_WIDGET[3]), ("Taille", WIDTH_COLUMN_WIDGET[4]), ("Format %", WIDTH_COLUMN_WIDGET[5]), ("Décimales", WIDTH_COLUMN_WIDGET[6])]:
            l = QLabel(lbl)
            l.setFixedWidth(w) if w > 0 else l.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
            l.setStyleSheet("color: #385188; font-size: 10px; font-weight: 600;")
            hdr_layout.addWidget(l)
        self._inner_layout.addWidget(hdr)

        # ── Conteneur des lignes ──
        self._rows_container = QWidget()
        self._rows_layout    = QVBoxLayout(self._rows_container)
        self._rows_layout.setContentsMargins(0, 0, 0, 0)
        self._rows_layout.setSpacing(2)
        self._inner_layout.addWidget(self._rows_container)

        # ── Encadré des paramètres globaux ──
        self._global_params_group = QGroupBox("Paramètres à appliquer")
        global_params_layout = QVBoxLayout(self._global_params_group)

        # Ligne de paramètres globaux
        global_row_widget = RowStyleWidget(-1, "Paramètres globaux", RowStyle())
        global_row_widget.setToolTip("Configurez les paramètres à appliquer aux lignes.")
        self._global_row_widget = global_row_widget
        global_params_layout.addWidget(global_row_widget)

        # ── Boutons ──
        btn_bar = QHBoxLayout()
        btn_apply_all = QPushButton("↕  Appliquer à toutes les lignes (hors années)")
        btn_apply_all.setObjectName("secondary")
        btn_apply_all.clicked.connect(self._apply_to_all)

        btn_apply_every_other = QPushButton("↕  Appliquer à une ligne sur deux")
        btn_apply_every_other.setObjectName("secondary")
        btn_apply_every_other.clicked.connect(self._apply_to_every_other)

        btn_bar.addWidget(btn_apply_all)
        btn_bar.addWidget(btn_apply_every_other)
        global_params_layout.addLayout(btn_bar)

        self._inner_layout.addWidget(self._global_params_group)

        self._inner_layout.addStretch()
        scroll.setWidget(inner)
        root.addWidget(scroll)

    # ── Callbacks état ────────────────────────────────────────────────────

    def _on_workbook_loaded(self, _=None):
        wb = state_manager.workbook
        if not wb:
            return
        self._loading = True
        current_sheet = self._combo_sheet.currentText()
        self._combo_sheet.clear()
        for s in sorted(wb.sheets, key=lambda s: s.page_order):
            self._combo_sheet.addItem(s.name)
        idx = self._combo_sheet.findText(current_sheet)
        self._combo_sheet.setCurrentIndex(idx if idx >= 0 else 0)
        self._loading = False
        self._on_sheet_changed(self._combo_sheet.currentText())

    def _on_sheet_changed(self, sheet_name: str):
        if not sheet_name or self._loading:
            return
        self._current_sheet = sheet_name
        self._loading = True

        current_table = self._combo_table.currentText()
        self._combo_table.clear()
        wb = state_manager.workbook
        if wb:
            sheet = next((s for s in wb.sheets if s.name == sheet_name), None)
            if sheet:
                for tz in sheet.tables:
                    self._combo_table.addItem(tz.name)
        idx = self._combo_table.findText(current_table)
        self._combo_table.setCurrentIndex(idx if idx >= 0 else 0)
        self._loading = False
        self._on_table_changed(self._combo_table.currentText())

    def _on_table_changed(self, table_name: str):
        if not table_name or self._loading or not self._current_sheet:
            return
        self._current_table = table_name
        self._rebuild_rows()

    # ── Construction de la liste des lignes ──────────────────────────────

    def _get_table_zone(self):
        wb = state_manager.workbook
        if not wb or not self._current_sheet or not self._current_table:
            return None
        sheet = next((s for s in wb.sheets if s.name == self._current_sheet), None)
        if not sheet:
            return None
        return next((t for t in sheet.tables if t.name == self._current_table), None)

    def _rebuild_rows(self):
        """Reconstruit la liste des widgets de lignes pour le tableau sélectionné."""
        # Nettoyer
        for _, w in self._row_widgets:
            w.setParent(None)
        self._row_widgets.clear()

        tz = self._get_table_zone()
        if tz is None:
            self._lbl_year_row.setVisible(False)
            return

        # Détecter le subheader années
        self._year_row_index = detect_year_row_index(tz)

        if self._year_row_index is not None:
            self._lbl_year_row.setText(
                f"ℹ  La ligne {self._year_row_index + 1} est reconnue comme sous-en-tête "
                f"d'années. Son style est géré automatiquement et elle est exclue "
                f"du bouton « Appliquer à toutes les lignes »."
            )
            self._lbl_year_row.setVisible(True)
        else:
            self._lbl_year_row.setVisible(False)

        # Récupérer les styles existants
        tsp = state_manager.get_table_style(self._current_sheet, self._current_table)

        for data_idx, row in enumerate(tz.data):
            # Le subheader années est affiché mais en lecture seule (désactivé)
            is_year_row = (data_idx == self._year_row_index)

            # Libellé = première cellule non-vide
            label = str(row[0]).strip() if row and row[0] is not None else f"Ligne {data_idx + 1}"
            if not label:
                label = f"Ligne {data_idx + 1}"

            rs = tsp.row_styles.get(data_idx, RowStyle())
            widget = RowStyleWidget(data_idx, label, rs)

            if is_year_row:
                widget.setEnabled(False)
                widget.setToolTip("Sous-en-tête années — style automatique")
                # Badge visuel
                widget.setStyleSheet("background: #EEF2FA; border-radius: 3px;")

            widget.changed.connect(lambda di=data_idx: self._on_row_changed(di))
            self._rows_layout.addWidget(widget)
            self._row_widgets.append((data_idx, widget))

    # ── Sync état ─────────────────────────────────────────────────────────

    def _on_row_changed(self, data_idx: int):
        if not self._current_sheet or not self._current_table:
            return
        widget = next((w for di, w in self._row_widgets if di == data_idx), None)
        if widget is None:
            return
        tsp = state_manager.get_table_style(self._current_sheet, self._current_table)
        rs = widget.get_row_style()
        # Toujours conserver les modifications explicites
        tsp.row_styles[data_idx] = rs
        state_manager.update_table_style(self._current_sheet, self._current_table, tsp)

    def _apply_to_all(self):
        """
        Applique les paramètres globaux à toutes les lignes éditables
        (hors subheader années) du tableau courant.
        """
        if not self._current_sheet or not self._current_table:
            return

        global_style = self._global_row_widget.get_row_style()
        tsp = state_manager.get_table_style(self._current_sheet, self._current_table)

        for data_idx, _ in self._row_widgets:
            if data_idx != self._year_row_index:  # Exclure le subheader années
                tsp.row_styles[data_idx] = copy.copy(global_style)

        state_manager.update_table_style(self._current_sheet, self._current_table, tsp)
        self._rebuild_rows()

    def _apply_to_every_other(self):
        """
        Applique les paramètres globaux à une ligne sur deux (hors subheader années)
        dans le tableau courant.
        """
        if not self._current_sheet or not self._current_table:
            return

        global_style = self._global_row_widget.get_row_style()
        tsp = state_manager.get_table_style(self._current_sheet, self._current_table)

        for data_idx, _ in self._row_widgets:
            if data_idx != self._year_row_index and data_idx % 2 == 0:  # Une ligne sur deux
                tsp.row_styles[data_idx] = copy.copy(global_style)

        state_manager.update_table_style(self._current_sheet, self._current_table, tsp)
        self._rebuild_rows()

    def _reset_all(self):
        """Supprime tous les RowStyle du tableau courant (hors subheader années)."""
        if not self._current_sheet or not self._current_table:
            return
        tsp = state_manager.get_table_style(self._current_sheet, self._current_table)
        yr_idx = self._year_row_index

        keys_to_remove = [di for di in list(tsp.row_styles.keys()) if di != yr_idx]
        for di in keys_to_remove:
            tsp.row_styles.pop(di, None)

        state_manager.update_table_style(self._current_sheet, self._current_table, tsp)
        self._rebuild_rows()