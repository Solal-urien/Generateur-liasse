"""
Onglet Page :
  Paramètres de mise en page par feuille (marges, titre, options)
  + couleurs par tableau (en-tête, alternance, texte).

  Les couleurs de titre de page et de sous-titre de section sont
  gérées dans l'onglet Style (GlobalParams) — elles ne sont pas
  modifiables ici pour garantir la cohérence du document.
"""
from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QCheckBox, QDoubleSpinBox, QComboBox,
    QSizePolicy, QLineEdit, QScrollArea, QFrame,
)
from PyQt6.QtCore import Qt

from state_manager import state_manager
from pdf_generator import PageParams
from ui.tab_style import ColorButton

_DEFAULT_PRIMARY  = "#1C2844"
_DEFAULT_COMP     = "#2A3D66"
_DEFAULT_ALT      = "#EEF2FA"
_DEFAULT_HDR_TEXT = "#FFFFFF"
_DEFAULT_BODY     = "#2F2F2F"


class PageTab(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_sheet: str | None  = None
        self._current_table: str | None  = None
        self._loading = False
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
        layout.setSpacing(12)

        # ── Feuille ──────────────────────────────────────────────────────
        grp_sheet = QGroupBox("Feuille")
        sh_layout = QHBoxLayout(grp_sheet)
        self._combo_sheet = QComboBox()
        self._combo_sheet.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._combo_sheet.currentTextChanged.connect(self._on_sheet_selected)
        sh_layout.addWidget(self._combo_sheet)
        layout.addWidget(grp_sheet)

        # ── Titre de la page ──────────────────────────────────────────────
        grp_title = QGroupBox("Titre affiché dans le PDF")
        title_layout = QHBoxLayout(grp_title)
        self._edit_title = QLineEdit()
        self._edit_title.setPlaceholderText("Laisser vide pour utiliser le nom de la feuille")
        self._edit_title.textChanged.connect(self._push_page)
        title_layout.addWidget(self._edit_title)
        layout.addWidget(grp_title)

        # ── Options de page ───────────────────────────────────────────────
        grp_opts = QGroupBox("Marges et taille de police pour la page")
        opts_layout = QVBoxLayout(grp_opts)

        self._chk_show_title = QCheckBox("Afficher le titre de la feuille")
        self._chk_show_title.setChecked(True)
        self._chk_show_title.stateChanged.connect(self._push_page)
        opts_layout.addWidget(self._chk_show_title)

        # Marges compactes
        mg_row = QHBoxLayout()
        mg_row.setSpacing(8)
        self._spins_margin: dict[str, QDoubleSpinBox] = {}
        # for sym, key in [("↑", "top"), ("↓", "bottom"), ("←", "left"), ("→", "right")]:
        for sym, key in [("Gauche", "left"), ("Droite", "right")]:
            col = QVBoxLayout()
            col.setSpacing(2)
            col.addWidget(QLabel(sym, alignment=Qt.AlignmentFlag.AlignHCenter))
            spin = QDoubleSpinBox()
            spin.setRange(0, 80)
            spin.setValue(15)
            spin.setSuffix(" mm")
            spin.setSingleStep(1)
            spin.setFixedWidth(82)
            spin.valueChanged.connect(self._push_page)
            self._spins_margin[key] = spin
            col.addWidget(spin)
            mg_row.addLayout(col)
        opts_layout.addLayout(mg_row)

        # Taille de police par page
        fs_row = QHBoxLayout()
        fs_row.addWidget(QLabel("Taille de police :"))
        self._spin_font = QDoubleSpinBox()
        self._spin_font.setRange(1, 28)
        self._spin_font.setValue(9)
        self._spin_font.setSuffix(" pt")
        self._spin_font.setSingleStep(0.5)
        self._spin_font.valueChanged.connect(self._push_page)
        fs_row.addStretch()
        fs_row.addWidget(self._spin_font)
        opts_layout.addLayout(fs_row)

        btn_apply_all = QPushButton("↕  Appliquer les paramètres à toutes les feuilles")
        btn_apply_all.setObjectName("secondary")
        btn_apply_all.clicked.connect(self._apply_to_all)
        opts_layout.addWidget(btn_apply_all)
        layout.addWidget(grp_opts)

        # ── Tableau ───────────────────────────────────────────────────────
        grp_tbl = QGroupBox("Tableau")
        tbl_layout = QHBoxLayout(grp_tbl)
        self._combo_table = QComboBox()
        self._combo_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._combo_table.currentTextChanged.connect(self._on_table_selected)
        tbl_layout.addWidget(self._combo_table)
        layout.addWidget(grp_tbl)

        # ── Couleurs du tableau ───────────────────────────────────────────
        grp_colors = QGroupBox("Couleurs du tableau")
        col_layout = QVBoxLayout(grp_colors)
        col_layout.setSpacing(6)

        hint = QLabel(
            "Laissez désactivé pour utiliser la couleur globale (onglet Style).\n"
            "Les couleurs de titre de page et sous-titre sont toujours globales."
        )
        hint.setStyleSheet("color: #5E5E5E; font-size: 11px; font-style: italic;")
        hint.setWordWrap(True)
        col_layout.addWidget(hint)

        def _color_row(label_text: str, default_hex: str, attr: str):
            row = QHBoxLayout()
            chk = QCheckBox(label_text)
            btn = ColorButton(default_hex)
            btn.setEnabled(False)
            chk.stateChanged.connect(lambda s, b=btn: (b.setEnabled(bool(s)), self._push_table()))
            btn.clicked.connect(self._push_table)
            row.addWidget(chk)
            row.addStretch()
            row.addWidget(btn)
            col_layout.addLayout(row)
            setattr(self, f"_chk_{attr}", chk)
            setattr(self, f"_btn_{attr}", btn)

        _color_row("Couleur principale (fond en-tête)",   _DEFAULT_PRIMARY,  "primary")
        _color_row("Couleur complémentaire (ligne années)", _DEFAULT_COMP,   "comp")
        _color_row("Lignes alternées",                      _DEFAULT_ALT,     "alt")
        _color_row("Texte de l'en-tête",                   _DEFAULT_HDR_TEXT, "hdr")
        _color_row("Texte du corps",                        _DEFAULT_BODY,    "body")

        btn_reset = QPushButton("↺  Couleurs globales pour ce tableau")
        btn_reset.setObjectName("secondary")
        btn_reset.clicked.connect(self._reset_colors)
        col_layout.addWidget(btn_reset)
        layout.addWidget(grp_colors)

        layout.addStretch()
        scroll.setWidget(inner)
        root.addWidget(scroll)

        self._set_enabled(False)

    # ── Helpers ─────────────────────────────────────────────────────────────

    def _set_enabled(self, on: bool):
        for w in [self._edit_title, self._chk_show_title, self._spin_font, self._combo_table]:
            w.setEnabled(on)
        for s in self._spins_margin.values():
            s.setEnabled(on)
        for attr in ("primary", "comp", "alt", "hdr", "body"):
            getattr(self, f"_chk_{attr}").setEnabled(on)
            # Les boutons suivent leur checkbox
            chk = getattr(self, f"_chk_{attr}")
            btn = getattr(self, f"_btn_{attr}")
            btn.setEnabled(on and chk.isChecked())

    # ── Sync état ────────────────────────────────────────────────────────────

    def _on_workbook_loaded(self, _=None):
        wb = state_manager.workbook
        if not wb:
            return
        self._loading = True
        cur = self._combo_sheet.currentText()
        self._combo_sheet.clear()
        for s in sorted(wb.sheets, key=lambda s: s.page_order):
            self._combo_sheet.addItem(s.name)
        idx = self._combo_sheet.findText(cur)
        self._combo_sheet.setCurrentIndex(max(0, idx))
        self._loading = False
        self._on_sheet_selected(self._combo_sheet.currentText())

    def _on_sheet_selected(self, sheet_name: str):
        if not sheet_name or self._loading:
            return
        self._current_sheet = sheet_name
        self._loading = True

        pp = state_manager.gparams.page_params.get(sheet_name, PageParams())
        self._edit_title.setText(pp.sheet_title or "")
        self._chk_show_title.setChecked(pp.show_sheet_title)
        self._spin_font.setValue(pp.font_size)
        # self._spins_margin["top"].setValue(pp.margin_top)
        # self._spins_margin["bottom"].setValue(pp.margin_bottom)
        self._spins_margin["left"].setValue(pp.margin_left)
        self._spins_margin["right"].setValue(pp.margin_right)

        # Recharger la liste des tableaux
        cur_tbl = self._combo_table.currentText()
        self._combo_table.blockSignals(True)
        self._combo_table.clear()
        wb = state_manager.workbook
        if wb:
            sheet = next((s for s in wb.sheets if s.name == sheet_name), None)
            if sheet:
                for t in sheet.tables:
                    self._combo_table.addItem(t.name)
        idx = self._combo_table.findText(cur_tbl)
        self._combo_table.setCurrentIndex(max(0, idx))
        self._combo_table.blockSignals(False)

        self._set_enabled(True)
        self._loading = False
        self._on_table_selected(self._combo_table.currentText())

    def _on_table_selected(self, table_name: str):
        if not table_name or self._loading:
            return
        self._current_table = table_name
        self._loading = True

        tsp = state_manager.get_table_style(self._current_sheet, table_name)
        for attr, field_name, default in [
            ("primary", "primary_color",    _DEFAULT_PRIMARY),
            ("comp",    "complement_color", _DEFAULT_COMP),
            ("alt",     "row_alt_color",    _DEFAULT_ALT),
            ("hdr",     "header_text_color", _DEFAULT_HDR_TEXT),
            ("body",    "body_text_color",  _DEFAULT_BODY),
        ]:
            chk = getattr(self, f"_chk_{attr}")
            btn = getattr(self, f"_btn_{attr}")
            val = getattr(tsp, field_name, None)
            has = bool(val)
            chk.setChecked(has)
            btn.setEnabled(has)
            btn.set_color(val if has else default)

        self._loading = False

    # ── Push vers l'état ─────────────────────────────────────────────────────

    def _push_page(self, *_):
        if self._loading or not self._current_sheet:
            return
        pp = state_manager.gparams.page_params.get(self._current_sheet, PageParams())
        pp.sheet_title      = self._edit_title.text().strip() or None
        pp.show_sheet_title = self._chk_show_title.isChecked()
        pp.font_size        = self._spin_font.value()
        # pp.margin_top       = self._spins_margin["top"].value()
        # pp.margin_bottom    = self._spins_margin["bottom"].value()
        pp.margin_left      = self._spins_margin["left"].value()
        pp.margin_right     = self._spins_margin["right"].value()
        state_manager.update_page_params(self._current_sheet, pp)

    def _push_table(self, *_):
        if self._loading or not self._current_sheet or not self._current_table:
            return

        def _opt(attr: str) -> str | None:
            chk = getattr(self, f"_chk_{attr}")
            btn = getattr(self, f"_btn_{attr}")
            return btn.color() if chk.isChecked() else None

        tsp = state_manager.get_table_style(self._current_sheet, self._current_table)
        tsp.primary_color     = _opt("primary")
        tsp.complement_color  = _opt("comp")
        tsp.row_alt_color     = _opt("alt")
        tsp.header_text_color = _opt("hdr")
        tsp.body_text_color   = _opt("body")
        state_manager.update_table_style(self._current_sheet, self._current_table, tsp)

    def _reset_colors(self):
        self._loading = True
        for attr in ("primary", "comp", "alt", "hdr", "body"):
            getattr(self, f"_chk_{attr}").setChecked(False)
            getattr(self, f"_btn_{attr}").setEnabled(False)
        self._loading = False
        self._push_table()

    def _apply_to_all(self):
        if not self._current_sheet:
            return
        src = state_manager.gparams.page_params.get(self._current_sheet, PageParams())
        wb  = state_manager.workbook
        if not wb:
            return
        for sheet in wb.sheets:
            if sheet.name == self._current_sheet:
                continue
            dst = state_manager.gparams.page_params.get(sheet.name, PageParams())
            dst.font_size       = src.font_size
            # dst.margin_top      = src.margin_top
            # dst.margin_bottom   = src.margin_bottom
            dst.margin_left     = src.margin_left
            dst.margin_right    = src.margin_right
            dst.show_sheet_title = src.show_sheet_title
            state_manager.update_page_params(sheet.name, dst)