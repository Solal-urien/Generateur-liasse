"""
Onglet Style :
  - Couleurs primaire / accent / complémentaire
  - Tailles de police globales
  - Options : numérotation, sommaire, date en pied de page
  - Tableau de page de garde : référencé par (sheet_name, table_name)
    et non plus par l'objet TableZone, afin que la mise en forme du tableau
    (TableStyleParams) soit correctement appliquée à l'affichage en page de garde.
"""
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QCheckBox, QDoubleSpinBox, QComboBox,
    QColorDialog, QMenu, QWidgetAction,
)
from PyQt6.QtGui import QColor, QIcon, QPixmap

from state_manager import state_manager
from util import NUANCIER_COULEURS

try:
    from reportlab.platypus import Table
except Exception:
    Table = None


# ── ColorButton ────────────────────────────────────────────────────────────

class ColorButton(QPushButton):
    """Bouton qui affiche et permet de choisir une couleur hex parmi un nuancier ou via un sélecteur."""

    def __init__(self, hex_color: str = "#1a3a5c", parent=None):
        super().__init__(parent)
        self._color = hex_color
        self._refresh_style()
        self.setFixedSize(36, 28)
        self.clicked.connect(self._show_color_menu)

    def _refresh_style(self):
        r, g, b = self._hex_to_rgb(self._color)
        luminance = 0.299 * r + 0.587 * g + 0.114 * b
        text_color = "#ffffff" if luminance < 128 else "#000000"
        self.setStyleSheet(
            f"background-color: {self._color}; color: {text_color}; "
            f"border: 1px solid #aabbcc; border-radius: 4px; font-size: 9px;"
        )
        self.setText(self._color.upper())

    def _show_color_menu(self):
        menu = QMenu(self)
        menu.setStyleSheet(
            "QMenu { background-color: #1C2844; color: #ffffff; border: 1px solid #4a5d7a; }"
            "QMenu::item { background-color: transparent; }"
            "QMenu::item:selected { background-color: #2e7bc4; }"
            "QMenu::separator { height: 1px; background: #4a5d7a; margin: 4px 8px; }"
        )
        for category, cat_colors in NUANCIER_COULEURS.items():
            cat_menu = menu.addMenu(category)
            cat_menu.setStyleSheet(menu.styleSheet())
            for color in cat_colors:
                action = QWidgetAction(cat_menu)
                pixmap = QPixmap(20, 20)
                pixmap.fill(QColor(color))
                action.setIcon(QIcon(pixmap))
                action.setData(color)
                action.triggered.connect(lambda _, c=color: self._set_color(c))
                cat_menu.addAction(action)

        custom = menu.addAction("Couleur personnalisée…")
        custom.triggered.connect(self._pick_custom_color)
        menu.exec(self.mapToGlobal(self.rect().bottomLeft()))

    def _pick_custom_color(self):
        dlg = QColorDialog(QColor(self._color), self)
        if dlg.exec():
            self._set_color(dlg.selectedColor().name())

    def _set_color(self, hex_color: str):
        self._color = hex_color
        self._refresh_style()
        self.clicked.emit()

    def color(self) -> str:
        return self._color

    def set_color(self, hex_color: str):
        self._color = hex_color
        self._refresh_style()

    @staticmethod
    def _hex_to_rgb(h: str):
        h = h.lstrip("#")
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


# ── StyleTab ───────────────────────────────────────────────────────────────

class StyleTab(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self._loading = False
        self._build_ui()
        state_manager.subscribe("workbook_loaded", self._on_workbook_loaded)
        state_manager.subscribe("profile_loaded",  self._load_from_state)

    # ── Construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(12)

        # ── Couleurs ──
        grp_colors = QGroupBox("Couleurs du document")
        colors_layout = QVBoxLayout(grp_colors)
        colors_layout.setSpacing(8)

        def _color_row(label: str, default: str, attr: str):
            row = QHBoxLayout()
            row.addWidget(QLabel(label))
            btn = ColorButton(default)
            btn.clicked.connect(lambda: self._on_color_changed(attr))
            row.addStretch()
            row.addWidget(btn)
            colors_layout.addLayout(row)
            setattr(self, f"_btn_{attr}", btn)

        _color_row("Couleur principale :",      "#1C2844", "primary")
        _color_row("Couleur accent :",           "#7891C7", "accent")
        _color_row("Couleur complémentaire :",   "#2A3D66", "complement")
        root.addWidget(grp_colors)

        # ── Typographie ──
        grp_typo = QGroupBox("Typographie")
        typo_layout = QVBoxLayout(grp_typo)
        typo_layout.setSpacing(8)

        row_ts = QHBoxLayout()
        row_ts.addWidget(QLabel("Taille titre :"))
        self._spin_title = QDoubleSpinBox()
        self._spin_title.setRange(8, 36)
        self._spin_title.setValue(14)
        self._spin_title.setSuffix(" pt")
        self._spin_title.setSingleStep(0.5)
        self._spin_title.valueChanged.connect(self._push_to_state)
        row_ts.addStretch()
        row_ts.addWidget(self._spin_title)
        typo_layout.addLayout(row_ts)

        row_hs = QHBoxLayout()
        row_hs.addWidget(QLabel("Taille en-tête section :"))
        self._spin_header = QDoubleSpinBox()
        self._spin_header.setRange(7, 24)
        self._spin_header.setValue(10)
        self._spin_header.setSuffix(" pt")
        self._spin_header.setSingleStep(0.5)
        self._spin_header.valueChanged.connect(self._push_to_state)
        row_hs.addStretch()
        row_hs.addWidget(self._spin_header)
        typo_layout.addLayout(row_hs)
        root.addWidget(grp_typo)

        # ── Options ──
        grp_opts = QGroupBox("Options document")
        opts_layout = QVBoxLayout(grp_opts)
        opts_layout.setSpacing(6)

        self._chk_toc = QCheckBox("Inclure un sommaire")
        self._chk_toc.setChecked(True)
        self._chk_toc.stateChanged.connect(self._push_to_state)
        opts_layout.addWidget(self._chk_toc)

        self._chk_page_num = QCheckBox("Numérotation des pages")
        self._chk_page_num.setChecked(True)
        self._chk_page_num.stateChanged.connect(self._push_to_state)
        opts_layout.addWidget(self._chk_page_num)

        self._chk_date = QCheckBox("Date de génération en pied de page")
        self._chk_date.setChecked(True)
        self._chk_date.stateChanged.connect(self._push_to_state)
        opts_layout.addWidget(self._chk_date)

        # Tableau de page de garde
        self._chk_intro_table = QCheckBox("Afficher un tableau sur la page de garde")
        self._chk_intro_table.setChecked(False)
        self._chk_intro_table.stateChanged.connect(self._push_to_state)
        opts_layout.addWidget(self._chk_intro_table)

        row_intro = QHBoxLayout()
        row_intro.addWidget(QLabel("Tableau de page de garde :"))
        self._combo_intro_table = QComboBox()
        self._combo_intro_table.setMaxVisibleItems(12)
        self._combo_intro_table.currentIndexChanged.connect(self._push_to_state)
        row_intro.addStretch()
        row_intro.addWidget(self._combo_intro_table)
        opts_layout.addLayout(row_intro)

        self._refresh_intro_tables()
        root.addWidget(grp_opts)

        info = QLabel(
            "Ces paramètres s'appliquent à l'ensemble du document.\n"
            "Pour un réglage par page, utilisez l'onglet Page."
        )
        info.setStyleSheet("color: #464646; font-size: 11px; font-style: italic;")
        info.setWordWrap(True)
        root.addWidget(info)
        root.addStretch()

    # ── Sync état ────────────────────────────────────────────────────────────

    def _on_workbook_loaded(self, _):
        self._refresh_intro_tables()
        self._load_from_state()

    def _refresh_intro_tables(self):
        """
        Peuple le combo avec les tableaux disponibles dans le workbook.
        Chaque item stocke (sheet_name, table_name) comme data, ce qui est
        la référence stable indépendante de l'objet TableZone.
        """
        self._combo_intro_table.blockSignals(True)
        self._combo_intro_table.clear()
        self._combo_intro_table.addItem("Aucun tableau", None)

        wb = getattr(state_manager, "workbook", None)
        if wb:
            for sheet in wb.sheets:
                for t in sheet.tables:
                    name = getattr(t, "name", None)
                    if name and name.startswith("Zone"):
                        continue
                    label = f"{sheet.name} / {name}" if name else f"{sheet.name} / Tableau"
                    self._combo_intro_table.addItem(label, (sheet.name, name))

        self._combo_intro_table.blockSignals(False)

    def _load_from_state(self, _=None):
        self._loading = True
        gp = state_manager.gparams
        self._btn_primary.set_color(gp.primary_color)
        self._btn_accent.set_color(gp.accent_color)
        self._btn_complement.set_color(gp.complement_color)
        self._spin_title.setValue(gp.title_font_size)
        self._spin_header.setValue(gp.header_font_size)
        self._chk_toc.setChecked(gp.show_toc)
        self._chk_page_num.setChecked(gp.show_page_numbers)
        self._chk_date.setChecked(gp.date_in_footer)

        ref = getattr(gp, "table_intro_ref", None)
        has_intro = ref is not None
        self._chk_intro_table.setChecked(has_intro)

        if ref:
            for i in range(self._combo_intro_table.count()):
                if self._combo_intro_table.itemData(i) == ref \
                        or (isinstance(self._combo_intro_table.itemData(i), (list, tuple))
                            and tuple(self._combo_intro_table.itemData(i)) == tuple(ref)):
                    self._combo_intro_table.setCurrentIndex(i)
                    break
        else:
            self._combo_intro_table.setCurrentIndex(0)

        self._loading = False

    def _on_color_changed(self, _which: str):
        self._push_to_state()

    def _push_to_state(self, *_):
        if self._loading:
            return
        gp = state_manager.gparams
        gp.primary_color    = self._btn_primary.color()
        gp.accent_color     = self._btn_accent.color()
        gp.complement_color = self._btn_complement.color()
        gp.title_font_size  = self._spin_title.value()
        gp.header_font_size = self._spin_header.value()
        gp.show_toc          = self._chk_toc.isChecked()
        gp.show_page_numbers = self._chk_page_num.isChecked()
        gp.date_in_footer    = self._chk_date.isChecked()

        if self._chk_intro_table.isChecked():
            raw = self._combo_intro_table.currentData()
            if isinstance(raw, (list, tuple)) and len(raw) == 2:
                gp.table_intro_ref = (raw[0], raw[1])
            else:
                gp.table_intro_ref = None
        else:
            gp.table_intro_ref = None

        state_manager.emit("global_params_changed", None)