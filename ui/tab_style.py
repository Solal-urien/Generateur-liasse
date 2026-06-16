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
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QGroupBox, QCheckBox, QDoubleSpinBox, QComboBox,
)
from state_manager import state_manager


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

        # ── Page de garde ──
        grp_cover = QGroupBox("Page de garde")
        cover_layout = QVBoxLayout(grp_cover)
        cover_layout.setSpacing(8)

        row_title = QHBoxLayout()
        row_title.addWidget(QLabel("Titre :"))
        self._edit_doc_title = QLineEdit()
        self._edit_doc_title.setPlaceholderText("Titre de la liasse")
        self._edit_doc_title.textChanged.connect(self._push_to_state)
        row_title.addWidget(self._edit_doc_title)
        cover_layout.addLayout(row_title)

        row_subtitle = QHBoxLayout()
        row_subtitle.addWidget(QLabel("Sous-titre :"))
        self._edit_doc_author = QLineEdit()
        self._edit_doc_author.setPlaceholderText("BEE, BEH…")
        self._edit_doc_author.textChanged.connect(self._push_to_state)
        row_subtitle.addWidget(self._edit_doc_author)
        cover_layout.addLayout(row_subtitle)

        root.addWidget(grp_cover)

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
        self._edit_doc_title.setText(getattr(gp, "doc_title", ""))
        self._edit_doc_author.setText(getattr(gp, "doc_author", ""))
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
        gp.doc_title        = self._edit_doc_title.text().strip()
        gp.doc_author       = self._edit_doc_author.text().strip()
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