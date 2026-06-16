""" Utilitaire pour l'application de création de liasse, contenant notamment la palette de couleurs et les mappings de zones Excel. """

from typing import Any

from PyQt6.QtWidgets import QPushButton, QColorDialog, QMenu, QWidgetAction
from PyQt6.QtGui import QColor, QIcon, QPixmap

def is_year_like(value: Any) -> bool:
    """ Détecte si une valeur ressemble à une année (1900-2100), pour aider à filtrer les colonnes d'années."""
    if isinstance(value, bool):
        return False
    if isinstance(value, int):
        return 1900 <= value <= 2100
    if isinstance(value, float) and value.is_integer():
        return 1900 <= int(value) <= 2100
    if isinstance(value, str):
        s = value.strip()
        return s.isdigit() and len(s) == 4 and 1900 <= int(s) <= 2100
    return False

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

# Palette de couleurs inspirée du nuancier
NUANCIER_COULEURS = {
    "Palette principale": [
        "#1C2844",  # Bleu foncé
        "#2A3D66",  # Bleu moyen
        "#385188",  # Bleu-gris
        "#7891C7",  # Bleu clair
        "#A5B5DA",  # Bleu très clair
        "#D2DAEC",  # Bleu-gris clair
        "#2F2F2F",  # Gris foncé
        "#464646",  # Gris moyen
        "#5E5E5E",  # Gris
        "#9E9E9E",  # Gris foncé
        "#BFBFBF",  # Gris-noir
        "#DEDEDE",  # Noir-gris
        "#030303",  # Rouge foncé
        "#CF3D3D",  # Rouge
        "#E28B8B",  # Rose
        "#ECB1B1",  # Rose pâle
        "#F5D8D8"   # Rose très pâle
        "#FFFFFF"   # Blanc

    ],
    "Couleurs complémentaires": [
        "#EBA039",  # Orange clair
        "#F3C688",  # Jaune clair
        "#F7D9B0",  # Chair
        "#FBECD7",  # Crème
        "#DD6115"   # Orange
    ],
    "Couleurs d'appoint": [ # 
        "#5D4171",  # Violet foncé
        "#7C5797",  # Violet
        "#B197C3",  # Lilas
        "#2C6435",  # Vert foncé
        "#3A8647",  # Vert moyen
        "#7AC587",  # Vert clair
        "#A6D9AF"   # Vert très clair
    ],
    "Arrière-plans": [
        "#DEDEDE",  # Gris très clair
        "#BFBFBF",  # Gris clair
        "#A5B5DA",  # Bleu très clair
        "#D2DAEC",  # Bleu-gris clair
        "#ECB1B1",  # Rose pâle
        "#F5D8D8",  # Rose très pâle
        "#F7D9B0",  # Chair
        "#FBECD7"   # Crème
    ],
    "Graphiques": [
        "#2A3D66",  # Bleu moyen
        "#385188",  # Bleu-gris
        "#7891C7",  # Bleu clair
        "#A12828",  # Rouge foncé
        "#CF3D3D",  # Rouge
        "#E28B8B",  # Rose
        "#EBA039",  # Orange clair
        "#F3C688",  # Jaune clair
        "#DD6115",  # Orange
        "#5D4171",  # Violet foncé
        "#7C5797",  # Violet
        "#B197C3",  # Lilas
        "#2C6435",  # Vert foncé
        "#3A8647",  # Vert moyen
        "#7AC587"   # Vert clair
    ]
}

### Mapping des zones d'intérêt dans le fichier Excel, par feuille : dict{sheet_name : {zones : [{name, range}]}} ###
mapping_fipu2 = {"Intro": { "zones": [{"name": "Tableau introductif", "range": "D17:G25"}]},
                 "0. Environnement macro" : { "zones": [{"name": "Environnement macroéconomique", "range": "B5:M35"}]},
                "TDP simplifié" : { "zones": [{"name": "Prélèvements obligatoires", "range": "B2:D10"},
                                              {"name": "Recettes non fiscales", "range": "B12:D19"},]},
                "1.3. Tableau de passage N" : { "zones": [{"name": "Tableau de passage du BEH26N2 au BEH26N3 - année 2026", "range": "B6:G21"}]},
                "1.3. Tableau de passage PROV" : { "zones": []}, # Pas à ne pas prendre en compte
                "2.1. Tableau principal PO RFN" : { "zones": [{"name": "Niveaux", "range": "D5:AF12"},
                                                              {"name": "Elasticités", "range": "D15:AF21"},
                                                              {"name": "Généralités", "range": "D23:Q27"},
                                                              {"name": "MN", "range": "D29:AF36"},
                                                              {"name": "% spontanée", "range": "D38:AF45"},
                                                              {"name": "Ecart à l'unité", "range": "D47:AF54"},
                                                              ]},
                "2.2 Evolution du taux de PO" : { "zones": [{"name": "2.2.1 D'une année à l'autre", "range": "D5:M10"},
                                                            {"name": "2.2.2 En cumul depuis 2019", "range": "D14:M18"},
                                                            {"name": "2.2.3 Entre deux comptes", "range": "D22:J29"}]},
                "3.1. Graphiques type d'impôts" : { "zones": [{"name": "Contribution des différents impôts à l'écart à l'élasticité unitaire", "range": "B62:N77"}]},  
                "3.2. Miyazaki" : { "zones": [{"name": "Contribution à l'écart à l'élasticité unitaire", "range": "B56:S76"},
                                    {"name": "Composition du résidu par type d'impôt (écart pour chaque impôt entre sa croissance spontanée réelle et sa croissance spontanée théorique)", "range": "B81:S91"}]},
                "3.3. Format HCFP" : { "zones": [{"name": "Tableau des principaux impôts", "range": "B4:Z83"}]},
                "4.1. Ecart en PO" : { "zones": [{"name": "Ecart entre deux comptes : vision par principaux PO et RFN", "range": "D5:P51"}]},
                "4.2. Ecart en RFN" : { "zones": [{"name": "Ecarts en RFN", "range": "D3:AQ61"}]},
                "4.3. Tableau de passage des RNF" : { "zones": [{"name": "Présentation des RFN", "range": "E3:Q31"},
                                                                {"name": "Tableau de passage des RNF", "range": "E34:Q100"}]},
                "5. Ecart - ss-secteurs CN pt1" : { "zones": [{"name": "Ecart entre deux comptes : vision par sous-secteurs et code de CN", "range": "D4:O48"}]},
                "5. Ecart - ss-secteurs CN pt2" : { "zones": [{"name": "Ecart entre deux comptes : vision par sous-secteurs et code de CN", "range": "D1:O45"}]},
                "6. Crédit d'impôt" : { "zones": [{"name": "Crédits d'impôts en dépenses (en effet solde)", "range": "D5:O29"},
                                                  {"name": "Crédits d'impôts en cash", "range": "D32:O59"}]},
                "8. Recap des contentieux" : { "zones": [{"name": "Dépenses CN", "range": "D3:I9"},
                                                         {"name": "Recettes CN", "range": "D11:I16"},
                                                         {"name": "Recettes fiscales nettes", "range": "D18:I24"},
                                                         {"name": "Recettes non fiscales", "range": "D26:I27"},
                                                         {"name": "Clés en dépense", "range": "D29:I34"},
                                                         {"name": "Clés en recette", "range": "D36:I42"}]},
                "9. Principales MN" : { "zones": [{"name": "Principales mesures nouvelles en prélèvements obligatoires", "range": "D4:I70"}]},   
                 }