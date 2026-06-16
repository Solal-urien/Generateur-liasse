"""
Onglet Page :
  - Selecteur de feuille
  - Titre personnalise de la feuille
  - Couleurs par page (principale, accent, complémentaires, texte en-tete, texte corps)
  - Taille de police par page
  - Marges (haut / bas / gauche / droite) par page
  - Affichage du titre de feuille
  - Bouton "Appliquer a toutes"
"""
from __future__ import annotations

from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QCheckBox, QDoubleSpinBox, QComboBox,
    QSizePolicy, QLineEdit, QScrollArea, QFrame)

from state_manager import state_manager
from pdf_generator import PageParams, TableStyleParams
from util import ColorButton


# Valeurs par defaut visuelles (correspondant aux GlobalParams par defaut)
_DEFAULT_PRIMARY     = "#1C2844"
_DEFAULT_ACCENT      = "#7891C7"
_DEFAULT_COMP        = "#2A3D66"
_DEFAULT_ALT         = "#EEF2FA"


class PageTab(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_sheet: str | None = None
        self._loading = False
        self._build_ui()
        state_manager.subscribe("workbook_loaded", self._on_workbook_loaded)
        state_manager.subscribe("profile_loaded", self._on_workbook_loaded)

    # ── Construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Zone scrollable pour tout le contenu
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        inner = QWidget()
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        # ── Selecteur de feuille ──
        grp_select = QGroupBox("Selectionner une feuille")
        sel_layout = QHBoxLayout(grp_select)
        self._combo = QComboBox()
        self._combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._combo.currentTextChanged.connect(self._on_sheet_selected)
        sel_layout.addWidget(self._combo)
        layout.addWidget(grp_select)

        # ── Titre personnalise ──
        grp_title = QGroupBox("Titre affiché dans le PDF")
        title_layout = QHBoxLayout(grp_title)
        self._edit_sheet_title = QLineEdit()
        self._edit_sheet_title.setPlaceholderText("Laisser vide pour utiliser le nom de la feuille")
        self._edit_sheet_title.textChanged.connect(self._push_to_state)
        title_layout.addWidget(self._edit_sheet_title)
        layout.addWidget(grp_title)

        # ── Couleurs par page ──
        grp_colors = QGroupBox("Couleurs de la page")
        colors_layout = QVBoxLayout(grp_colors)
        colors_layout.setSpacing(8)

        # Ligne : use global / reset
        hint = QLabel("Laisser la couleur globale (onglet Style) en désactivant la surcharge.")
        hint.setStyleSheet("color: #2F2F2F; font-size: 11px; font-style: italic;")
        hint.setWordWrap(True)
        colors_layout.addWidget(hint)

        def _color_row(label_text: str, default_hex: str, attr_name: str):
            """Cree une ligne checkbox + label + bouton couleur."""
            row = QHBoxLayout()
            chk = QCheckBox(label_text)
            btn = ColorButton(default_hex)
            btn.setEnabled(False)

            def _on_chk(state):
                btn.setEnabled(bool(state))
                self._push_to_state()

            def _on_btn_clicked():
                self._push_to_state()

            chk.stateChanged.connect(_on_chk)
            btn.clicked.connect(_on_btn_clicked)

            row.addWidget(chk)
            row.addStretch()
            row.addWidget(btn)
            colors_layout.addLayout(row)
            setattr(self, f"_chk_{attr_name}", chk)
            setattr(self, f"_btn_{attr_name}", btn)
        _color_row("Surcharger la couleur principale",   _DEFAULT_PRIMARY,  "primary")
        #_color_row("Surcharger la couleur des sous-titres",        _DEFAULT_ACCENT,   "accent")
        _color_row("Surcharger la couleur complémentaire", _DEFAULT_COMP,      "comp")
        _color_row("Surcharger la couleur lignes alternées", _DEFAULT_ALT,   "alt")
        # _color_row("Surcharger le texte en-tête tableau", _DEFAULT_HDR_TEXT, "hdr_text")
        # _color_row("Surcharger le texte corps tableau",   _DEFAULT_BODY_TEXT,"body_text")

        btn_reset = QPushButton("↺  Revenir aux couleurs globales pour cette page")
        btn_reset.setObjectName("secondary")
        btn_reset.clicked.connect(self._reset_colors)
        colors_layout.addWidget(btn_reset)

        layout.addWidget(grp_colors)

        # ── Typographie ──
        grp_typo = QGroupBox("Typographie")
        typo_layout = QVBoxLayout(grp_typo)
        typo_layout.setSpacing(8)

        row_fs = QHBoxLayout()
        row_fs.addWidget(QLabel("Taille de police :"))
        self._spin_font = QDoubleSpinBox()
        self._spin_font.setRange(5, 28)
        self._spin_font.setValue(9)
        self._spin_font.setSuffix(" pt")
        self._spin_font.setSingleStep(0.25)
        self._spin_font.valueChanged.connect(self._push_to_state)
        row_fs.addStretch()
        row_fs.addWidget(self._spin_font)
        typo_layout.addLayout(row_fs)

        layout.addWidget(grp_typo)

        # ── Marges ──
        grp_margins = QGroupBox("Marges (mm)")
        mg_layout = QVBoxLayout(grp_margins)
        mg_layout.setSpacing(8)

        self._spins_margin: dict[str, QDoubleSpinBox] = {}
        for label, key in [("Gauche", "left"), ("Droite", "right")]:
            row = QHBoxLayout()
            row.addWidget(QLabel(f"{label} :"))
            spin = QDoubleSpinBox()
            spin.setRange(0, 80)
            spin.setValue(15)
            spin.setSuffix(" mm")
            spin.setSingleStep(1)
            spin.valueChanged.connect(self._push_to_state)
            self._spins_margin[key] = spin
            row.addStretch()
            row.addWidget(spin)
            mg_layout.addLayout(row)

        layout.addWidget(grp_margins)

        # ── Options ──
        grp_opts = QGroupBox("Options")
        opts_layout = QVBoxLayout(grp_opts)
        self._chk_title = QCheckBox("Afficher le titre de la feuille")
        self._chk_title.setChecked(True)
        self._chk_title.stateChanged.connect(self._push_to_state)
        opts_layout.addWidget(self._chk_title)
        layout.addWidget(grp_opts)

        # ── Bouton appliquer a toutes ──
        btn_apply_all = QPushButton("↕  Appliquer ces reglages a toutes les feuilles")
        btn_apply_all.setObjectName("secondary")
        btn_apply_all.clicked.connect(self._apply_to_all)
        layout.addWidget(btn_apply_all)

        layout.addStretch()

        scroll.setWidget(inner)
        root.addWidget(scroll)

        # etat initial
        self._set_controls_enabled(False)

    # ── Helpers ─────────────────────────────────────────────────────────────

    def _set_controls_enabled(self, enabled: bool):
        for w in [self._spin_font, self._chk_title,
                  self._edit_sheet_title,
                  self._chk_primary, self._chk_alt, self._chk_comp,
                # self._chk_hdr_text, self._chk_body_text
                  ]:
            w.setEnabled(enabled)
        for spin in self._spins_margin.values():
            spin.setEnabled(enabled)
        # Les boutons couleur suivent l'etat de leur checkbox associee
        for attr in ("primary", "comp", "alt"):
            chk: QCheckBox = getattr(self, f"_chk_{attr}")
            btn: ColorButton = getattr(self, f"_btn_{attr}")
            btn.setEnabled(enabled and chk.isChecked())

    def _reset_colors(self):
        """Desactive toutes les surcharges de couleur pour la feuille courante."""
        self._loading = True
        for attr in ("primary", "comp", "alt"):
            chk: QCheckBox = getattr(self, f"_chk_{attr}")
            btn: ColorButton = getattr(self, f"_btn_{attr}")
            chk.setChecked(False)
            btn.setEnabled(False)
        self._loading = False
        self._push_to_state()

    # ── Sync etat ────────────────────────────────────────────────────────────

    def _on_workbook_loaded(self, _=None):
        wb = state_manager.workbook
        if not wb:
            return
        self._loading = True
        current = self._combo.currentText()
        self._combo.clear()
        sheets_sorted = sorted(wb.sheets, key=lambda s: s.page_order)
        for s in sheets_sorted:
            self._combo.addItem(s.name)
        idx = self._combo.findText(current)
        if idx >= 0:
            self._combo.setCurrentIndex(idx)
        elif self._combo.count() > 0:
            self._combo.setCurrentIndex(0)
        self._loading = False
        self._on_sheet_selected(self._combo.currentText())

    def _on_sheet_selected(self, sheet_name: str):
        if not sheet_name:
            return
        self._current_sheet = sheet_name
        self._loading = True

        pp = state_manager.gparams.page_params.get(sheet_name, PageParams())

        # Titre personnalisé
        self._edit_sheet_title.setText(pp.sheet_title or "")

        # Couleurs par page (directement depuis PageParams)
        for attr, field_name, default in [
            ("primary",   "primary_color",     _DEFAULT_PRIMARY),
            # ("accent",    "accent_color",      _DEFAULT_ACCENT),
            ("comp",      "complement_color",  _DEFAULT_COMP),
            ("alt",       "row_alt_color",     _DEFAULT_ALT),
        ]:
            chk: QCheckBox = getattr(self, f"_chk_{attr}")
            btn: ColorButton = getattr(self, f"_btn_{attr}")
            val = getattr(pp, field_name, None)
            has_override = bool(val)
            chk.setChecked(has_override)
            btn.setEnabled(has_override)
            btn.set_color(val if has_override else default)

        # Typographie
        self._spin_font.setValue(pp.font_size)

        # Marges
        self._spins_margin["left"].setValue(pp.margin_left)
        self._spins_margin["right"].setValue(pp.margin_right)

        # Option titre
        self._chk_title.setChecked(pp.show_sheet_title)

        self._set_controls_enabled(True)
        self._loading = False

    def _push_to_state(self, *_):
        if self._loading or not self._current_sheet:
            return

        def _opt_color(attr: str) -> str | None:
            chk: QCheckBox = getattr(self, f"_chk_{attr}")
            btn: ColorButton = getattr(self, f"_btn_{attr}")
            return btn.color() if chk.isChecked() else None

        pp = PageParams(
            font_size=self._spin_font.value(),
            margin_left=self._spins_margin["left"].value(),
            margin_right=self._spins_margin["right"].value(),
            show_sheet_title=self._chk_title.isChecked(),
            sheet_title=self._edit_sheet_title.text().strip() or None,
            primary_color=_opt_color("primary"),
            # accent_color=_opt_color("accent"),
            complement_color=_opt_color("comp"),
            row_alt_color=_opt_color("alt"),
            table_styles=state_manager.gparams.page_params.get(self._current_sheet, PageParams()).table_styles,
        )
        state_manager.update_page_params(self._current_sheet, pp)

    def _apply_to_all(self):
        if not self._current_sheet:
            return
        pp_src = state_manager.gparams.page_params.get(self._current_sheet, PageParams())
        wb = state_manager.workbook
        if not wb:
            return
        import copy
        for sheet in wb.sheets:
            state_manager.update_page_params(sheet.name, copy.copy(pp_src))