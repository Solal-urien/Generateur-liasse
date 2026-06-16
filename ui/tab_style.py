"""
Onglet Style :
  Couleurs globales (primaire / accent / complémentaire),
  typographie (taille titre, taille en-tête de section),
  options document (sommaire, numérotation, date, tableau de garde).

  Ces paramètres s'appliquent à TOUT le document.
  Les couleurs par tableau sont configurables dans l'onglet Page.
  Les styles de lignes individuels sont configurables dans l'onglet Lignes.
"""
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QGroupBox, QCheckBox, QDoubleSpinBox, QComboBox,
    QColorDialog, QMenu, QWidgetAction, QPushButton,
)
from PyQt6.QtGui import QColor, QIcon, QPixmap

from state_manager import state_manager
from util import NUANCIER_COULEURS

try:
    from reportlab.platypus import Table as RLTable
except Exception:
    RLTable = None


# ── ColorButton ───────────────────────────────────────────────────────────

class ColorButton(QPushButton):
    """Bouton affichant une couleur hex et permettant de la modifier via un nuancier."""

    def __init__(self, hex_color: str = "#1a3a5c", parent=None):
        super().__init__(parent)
        self._color = hex_color
        self._refresh_style()
        self.setFixedSize(36, 28)
        self.clicked.connect(self._show_menu)

    # ── Style visuel ──────────────────────────────────────────────────────

    def _refresh_style(self):
        r, g, b   = self._hex_to_rgb(self._color)
        luminance = 0.299 * r + 0.587 * g + 0.114 * b
        txt       = "#ffffff" if luminance < 128 else "#000000"
        self.setStyleSheet(
            f"background-color:{self._color}; color:{txt}; "
            f"border:1px solid #aabbcc; border-radius:4px; font-size:9px;"
        )
        self.setText(self._color.upper())

    # ── Menu nuancier ─────────────────────────────────────────────────────

    def _show_menu(self):
        menu = QMenu(self)
        menu.setStyleSheet(
            "QMenu{background:#1C2844;color:#fff;border:1px solid #4a5d7a;}"
            "QMenu::item{background:transparent;}"
            "QMenu::item:selected{background:#2e7bc4;}"
            "QMenu::separator{height:1px;background:#4a5d7a;margin:4px 8px;}"
        )
        for category, clrs in NUANCIER_COULEURS.items():
            sub = menu.addMenu(category)
            sub.setStyleSheet(menu.styleSheet())
            for color in clrs:
                act = QWidgetAction(sub)
                px  = QPixmap(20, 20)
                px.fill(QColor(color))
                act.setIcon(QIcon(px))
                act.setData(color)
                act.triggered.connect(lambda _, c=color: self._set_color(c))
                sub.addAction(act)
        custom = menu.addAction("Couleur personnalisée…")
        custom.triggered.connect(self._pick_custom)
        menu.exec(self.mapToGlobal(self.rect().bottomLeft()))

    def _pick_custom(self):
        dlg = QColorDialog(QColor(self._color), self)
        if dlg.exec():
            self._set_color(dlg.selectedColor().name())

    def _set_color(self, hex_color: str):
        self._color = hex_color
        self._refresh_style()
        self.clicked.emit()

    # ── API publique ──────────────────────────────────────────────────────

    def color(self) -> str:
        return self._color

    def set_color(self, hex_color: str):
        self._color = hex_color
        self._refresh_style()

    @staticmethod
    def _hex_to_rgb(h: str) -> tuple[int, int, int]:
        h = h.lstrip("#")
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


# ── StyleTab ──────────────────────────────────────────────────────────────

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

        # ── Couleurs globales ────────────────────────────────────────────
        grp_colors = QGroupBox("Couleurs globales du document")
        cl = QVBoxLayout(grp_colors)
        cl.setSpacing(8)

        hint = QLabel(
            "Ces couleurs s'appliquent à tous les titres, sous-titres et tableaux.\n"
            "Pour surcharger tableau par tableau, utilisez l'onglet Page."
        )
        hint.setStyleSheet("color:#5E5E5E; font-size:11px; font-style:italic;")
        hint.setWordWrap(True)
        cl.addWidget(hint)

        def _color_row(label: str, default: str, btn_attr: str, slot):
            row = QHBoxLayout()
            row.addWidget(QLabel(label))
            btn = ColorButton(default)
            btn.clicked.connect(slot)
            setattr(self, btn_attr, btn)
            row.addStretch()
            row.addWidget(btn)
            cl.addLayout(row)

        _color_row("Couleur principale :",     "#1C2844", "_btn_primary",    lambda: self._push())
        _color_row("Couleur accent :",          "#7891C7", "_btn_accent",     lambda: self._push())
        _color_row("Couleur complémentaire :", "#2A3D66", "_btn_complement", lambda: self._push())
        root.addWidget(grp_colors)

        # ── Typographie ──────────────────────────────────────────────────
        grp_typo = QGroupBox("Typographie")
        tl = QVBoxLayout(grp_typo)
        tl.setSpacing(8)

        def _spin_row(label: str, spin_attr: str, vmin: float, vmax: float, default: float):
            row = QHBoxLayout()
            row.addWidget(QLabel(label))
            spin = QDoubleSpinBox()
            spin.setRange(vmin, vmax)
            spin.setValue(default)
            spin.setSuffix(" pt")
            spin.setSingleStep(0.5)
            spin.valueChanged.connect(self._push)
            setattr(self, spin_attr, spin)
            row.addStretch()
            row.addWidget(spin)
            tl.addLayout(row)

        _spin_row("Taille titre :",             "_spin_title",  8,  36, 14.0)
        _spin_row("Taille en-tête section :",   "_spin_header", 7,  24, 10.0)
        root.addWidget(grp_typo)

        # ── Options document ─────────────────────────────────────────────
        grp_opts = QGroupBox("Options document")
        ol = QVBoxLayout(grp_opts)
        ol.setSpacing(6)

        def _chk(label: str, attr: str, default: bool):
            c = QCheckBox(label)
            c.setChecked(default)
            c.stateChanged.connect(self._push)
            setattr(self, attr, c)
            ol.addWidget(c)

        _chk("Inclure un sommaire",                  "_chk_toc",      True)
        _chk("Numérotation des pages",               "_chk_page_num", True)
        _chk("Date de génération en pied de page",   "_chk_date",     True)
        _chk("Afficher un tableau sur la page de garde", "_chk_intro", False)

        row_intro = QHBoxLayout()
        row_intro.addWidget(QLabel("Tableau de page de garde :"))
        self._combo_intro = QComboBox()
        self._combo_intro.setMaxVisibleItems(12)
        self._combo_intro.currentIndexChanged.connect(self._push)
        row_intro.addStretch()
        row_intro.addWidget(self._combo_intro)
        ol.addLayout(row_intro)
        self._refresh_intro_tables()

        root.addWidget(grp_opts)
        root.addStretch()

    # ── Sync état ────────────────────────────────────────────────────────────

    def _on_workbook_loaded(self, _):
        self._refresh_intro_tables()
        self._load_from_state()

    def _refresh_intro_tables(self):
        self._combo_intro.blockSignals(True)
        self._combo_intro.clear()
        self._combo_intro.addItem("Aucun tableau", None)
        wb = getattr(state_manager, "workbook", None)
        if wb:
            for sheet in getattr(wb, "sheets", []):
                for t in getattr(sheet, "tables", []):
                    name = getattr(t, "name", None)
                    if name and name.startswith("Zone"):
                        continue
                    label = getattr(t, "title", None) or name or "Tableau"
                    self._combo_intro.addItem(label, t)
        self._combo_intro.blockSignals(False)

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
        self._chk_intro.setChecked(bool(getattr(gp, "show_intro_table", False)))
        table_intro = getattr(gp, "table_intro", None)
        if table_intro is not None:
            for i in range(self._combo_intro.count()):
                if self._combo_intro.itemData(i) is table_intro:
                    self._combo_intro.setCurrentIndex(i)
                    break
        self._loading = False

    def _push(self, *_):
        if self._loading:
            return
        gp = state_manager.gparams
        gp.primary_color     = self._btn_primary.color()
        gp.accent_color      = self._btn_accent.color()
        gp.complement_color  = self._btn_complement.color()
        gp.title_font_size   = self._spin_title.value()
        gp.header_font_size  = self._spin_header.value()
        gp.show_toc          = self._chk_toc.isChecked()
        gp.show_page_numbers = self._chk_page_num.isChecked()
        gp.date_in_footer    = self._chk_date.isChecked()
        gp.show_intro_table  = self._chk_intro.isChecked()
        sel = self._combo_intro.currentData()
        gp.table_intro = sel if gp.show_intro_table else None
        state_manager.emit("global_params_changed", None)