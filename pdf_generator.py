"""
Générateur PDF : prend un WorkbookModel + paramètres de mise en page
et produit un PDF multi-pages (1 feuille = 1 page par défaut).
"""
import io
from dataclasses import dataclass, field
from datetime import date
import math
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    Paragraph, Spacer, HRFlowable,
    PageBreak, Image, KeepTogether, BaseDocTemplate, Frame, PageTemplate,
    NextPageTemplate, ActionFlowable
)
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from excel_parser import WorkbookModel, TableZone
from pdf_table_maker import make_rl_table
from util import NUANCIER_COULEURS

pdfmetrics.registerFont(TTFont("Marianne",        "C:/Windows/Fonts/Marianne-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Marianne-Bold",   "C:/Windows/Fonts/Marianne-Bold.ttf"))
pdfmetrics.registerFont(TTFont("Marianne-Italic", "C:/Windows/Fonts/Marianne-RegularItalic.ttf"))
pdfmetrics.registerFontFamily("Marianne",
    normal="Marianne", bold="Marianne-Bold", italic="Marianne-Italic")
font_name = "Marianne"


# ── RowStyle : style d'une ligne individuelle ──────────────────────────────

@dataclass
class RowStyle:
    """
    Paramètres de style appliqués à une ligne de données d'un tableau.
    L'index correspond à la position dans TableZone.data (0 = première ligne hors en-tête).
    """
    visible:          bool       = True     # ligne affichée ou masquée
    background_color: str | None = None     # None = couleur alternée par défaut
    text_style:       str        = "normal" # "normal" | "bold" | "italic"
    font_size:        float | None = None   # None = hériter de PageParams.font_size
    number_format:    str        = "normal" # "normal" | "percent"
    decimal_places:   int        = 1        # 0..3


# ── TableStyleParams : style d'un tableau complet ─────────────────────────

@dataclass
class TableStyleParams:
    """
    Couleurs propres à un tableau.  None = utiliser la valeur globale (GlobalParams).
    Contient également les RowStyle ligne par ligne.
    """
    primary_color:      str | None = None   # fond de l'en-tête
    complement_color:   str | None = None   # ligne "années"
    row_alt_color:      str | None = None   # lignes alternées
    header_text_color:  str | None = None   # texte de l'en-tête
    body_text_color:    str | None = None   # texte du corps
    # clé = index dans TableZone.data (int)
    row_styles: dict[int, RowStyle] = field(default_factory=dict)


# ── PageParams : paramètres d'une page (feuille) ─────────────────────────

@dataclass
class PageParams:
    font_size:        float = 9.0
    margin_top:       float = 15.0   # mm
    margin_bottom:    float = 15.0
    margin_left:      float = 15.0
    margin_right:     float = 15.0
    show_sheet_title: bool  = True
    orientation:      str   = "portrait"
    sheet_title:      str | None = None   # None → nom de la feuille (déprécié, voir SheetSettings.display_name)
    # styles par tableau : clé = TableZone.name
    table_styles: dict[str, TableStyleParams] = field(default_factory=dict)


# ── SheetSettings : réglages par page liés à l'onglet Document ───────────

@dataclass
class SheetSettings:
    """
    Réglages d'une page (feuille), persistés dans le profil JSON.
    La clé de référencement dans GlobalParams.sheet_settings est SheetModel.name
    (le nom ORIGINAL de la feuille Excel), qui ne change jamais — ce qui permet
    de conserver le lien avec l'Excel même si l'ordre des pages est modifié
    ou si l'utilisateur renomme l'affichage de la page.
    """
    include:      bool = True
    page_order:   int  = 0
    display_name: str | None = None   # nom affiché (titre de page + sommaire) ; None → nom Excel
    footer_note:  str | None = None   # note affichée dans la marge inférieure de la page


# ── GlobalParams : paramètres globaux du document ─────────────────────────

@dataclass
class GlobalParams:
    title_font_size:  float = 14.0
    header_font_size: float = 10.0
    show_page_numbers: bool = True
    show_toc:          bool = True
    date_in_footer:    bool = True
    start_year:        int  = 2025
    primary_color:    str   = NUANCIER_COULEURS["Palette principale"][0]
    accent_color:     str   = NUANCIER_COULEURS["Palette principale"][3]
    complement_color: str   = NUANCIER_COULEURS["Palette principale"][1]
    font_family:      str   = font_name
    page_params: dict[str, PageParams] = field(default_factory=dict)
    # Réglages par page (inclusion, ordre, nom affiché, note de bas de page)
    # clé = SheetModel.name (nom Excel original, stable)
    sheet_settings: dict[str, SheetSettings] = field(default_factory=dict)
    # Référence du tableau utilisé en page de garde : (sheet_name, table_name) ou None
    table_intro_ref: tuple[str, str] | None = None


# ── ChartSpec ──────────────────────────────────────────────────────────────

@dataclass
class ChartSpec:
    sheet_name:  str
    table_name:  str
    chart_type:  str   = "bar"
    x_col:       int   = 0
    y_cols:      list[int] = field(default_factory=lambda: [1])
    title:       str   = ""
    xlabel:      str   = ""
    ylabel:      str   = ""
    palette:     list[str] = field(default_factory=list)
    figsize:     tuple[float, float] = (16, 6)
    legend:      bool  = True


# ── PageCollector ──────────────────────────────────────────────────────────

class PageCollector(ActionFlowable):
    def __init__(self, sheet_name: str, page_map: dict):
        super().__init__()
        self.sheet_name = sheet_name
        self.page_map   = page_map

    def apply(self, doc):
        self.page_map[self.sheet_name] = doc.page


# ── Utilitaires couleur ────────────────────────────────────────────────────

def _hex_to_rl(hex_color: str):
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return colors.Color(r / 255, g / 255, b / 255)


# ── Rendu graphique matplotlib ────────────────────────────────────────────

def _render_chart(spec: ChartSpec, table: TableZone) -> bytes | None:
    data    = table.data
    headers = table.headers
    if not data:
        return None

    palette = spec.palette or NUANCIER_COULEURS["Graphiques"]
    fig, ax = plt.subplots(figsize=spec.figsize, dpi=120)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("#DEDEDE")
    ax.grid(True, linestyle="--", alpha=0.5, color="#BFBFBF")
    for spine in ax.spines.values():
        spine.set_visible(False)

    try:
        x_vals   = [row[spec.x_col] if len(row) > spec.x_col else "" for row in data]
        y_series = []
        for yc in spec.y_cols:
            y_series.append([
                float(row[yc]) if len(row) > yc and row[yc] is not None else 0
                for row in data
            ])
        ct    = spec.chart_type
        x_pos = range(len(x_vals))

        if ct == "bar":
            width = 0.8 / max(len(y_series), 1)
            for i, (ys, col) in enumerate(zip(y_series, palette)):
                offset = (i - len(y_series) / 2 + 0.5) * width
                label  = headers[spec.y_cols[i]] if spec.y_cols[i] < len(headers) else f"S{i+1}"
                ax.bar([x + offset for x in x_pos], ys, width=width, label=label,
                       color=col, alpha=0.85, zorder=2)
            ax.set_xticks(list(x_pos))
            ax.set_xticklabels([str(v) for v in x_vals], rotation=30, ha="right")
        elif ct == "line":
            for i, (ys, col) in enumerate(zip(y_series, palette)):
                label = headers[spec.y_cols[i]] if spec.y_cols[i] < len(headers) else f"S{i+1}"
                ax.plot(list(x_pos), ys, marker="o", label=label, color=col, linewidth=2, zorder=2)
            ax.set_xticks(list(x_pos))
            ax.set_xticklabels([str(v) for v in x_vals], rotation=30, ha="right")
        elif ct == "area":
            for i, (ys, col) in enumerate(zip(y_series, palette)):
                label = headers[spec.y_cols[i]] if spec.y_cols[i] < len(headers) else f"S{i+1}"
                ax.fill_between(list(x_pos), ys, alpha=0.4, color=col)
                ax.plot(list(x_pos), ys, color=col, linewidth=1.5, label=label)
            ax.set_xticks(list(x_pos))
            ax.set_xticklabels([str(v) for v in x_vals], rotation=30, ha="right")
        elif ct == "pie":
            ys   = y_series[0] if y_series else []
            lbls = [str(v) for v in x_vals]
            ax.pie(ys, labels=lbls, colors=palette[:len(ys)], autopct="%1.1f%%", startangle=140)
            ax.axis("equal")
        elif ct == "scatter":
            for i, (ys, col) in enumerate(zip(y_series, palette)):
                label = headers[spec.y_cols[i]] if spec.y_cols[i] < len(headers) else f"S{i+1}"
                x_num = []
                for v in x_vals:
                    try:
                        x_num.append(float(v))
                    except (ValueError, TypeError):
                        x_num.append(i)
                ax.scatter(x_num, ys, color=col, label=label, alpha=0.7, s=60, zorder=2)

        if spec.title:
            ax.set_title(spec.title, fontsize=11, fontweight="bold", pad=10)
        if spec.xlabel:
            ax.set_xlabel(spec.xlabel, fontsize=9)
        if spec.ylabel:
            ax.set_ylabel(spec.ylabel, fontsize=9)
        if spec.legend and len(y_series) > 1:
            ax.legend(fontsize=8, framealpha=0.8)

        plt.tight_layout(pad=1.5)
        buf = io.BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight")
        buf.seek(0)
        result = buf.read()
        plt.close(fig)
        return result

    except Exception as e:
        plt.close(fig)
        print(f"[chart error] {e}")
        return None


# ── Footer ─────────────────────────────────────────────────────────────────

def _footer_fn(canvas, doc, gparams: GlobalParams, footer_map: dict):
    """
    Dessine le pied de page : numéro de page (droite), date de génération (gauche, page > 1),
    et la note de bas de page propre à la feuille en cours (centrée), si définie.

    footer_map : dict {page_number(int) -> footer_note(str)} rempli pendant la construction
    du story (via PageCollector), pour savoir quelle note afficher sur quelle page.
    Comme une feuille peut s'étendre sur plusieurs pages, on propage la note à toutes
    les pages de la feuille (cf _build_sheet_story).
    """
    canvas.saveState()
    canvas.setFont("Marianne", 7)
    canvas.setFillColor(colors.Color(0.5, 0.5, 0.5))
    if gparams.show_page_numbers:
        canvas.drawRightString(doc.pagesize[0] - 10 * mm, 8 * mm, f"Page {doc.page}")
    if gparams.date_in_footer and doc.page > 1:
        canvas.drawString(10 * mm, 8 * mm, f"Généré le {date.today().strftime('%d/%m/%Y')}")

    note = footer_map.get(doc.page)
    if note:
        canvas.setFont("Marianne", 7)
        canvas.drawCentredString(doc.pagesize[0] / 2, 8 * mm, note)
    canvas.restoreState()


# ── Templates de page ──────────────────────────────────────────────────────

def _build_page_templates(gparams: GlobalParams, footer_map: dict) -> list:
    templates = []
    for orient in ["portrait", "landscape"]:
        pp = PageParams(orientation=orient)
        ps = A4
        frame = Frame(
            pp.margin_left * mm, pp.margin_bottom * mm,
            ps[0] - (pp.margin_left + pp.margin_right) * mm,
            ps[1] - (pp.margin_top  + pp.margin_bottom) * mm,
            id=orient,
        )
        pt = PageTemplate(
            id=orient, frames=[frame], pagesize=ps,
            onPage=lambda c, d, gp=gparams, fm=footer_map: _footer_fn(c, d, gp, fm),
        )
        templates.append(pt)
    return templates


# ── Résolution du tableau d'introduction ──────────────────────────────────

def _resolve_table_intro(model: WorkbookModel, gparams: GlobalParams):
    """
    Retourne (TableZone, TableStyleParams) pour le tableau référencé par
    gparams.table_intro_ref = (sheet_name, table_name), ou (None, None) si
    la référence est absente ou ne pointe plus vers un tableau existant.
    """
    ref = getattr(gparams, "table_intro_ref", None)
    if not ref or not model:
        return None, None
    sheet_name, table_name = ref
    for sheet in model.sheets:
        if sheet.name == sheet_name:
            for t in sheet.tables:
                if t.name == table_name:
                    pp  = gparams.page_params.get(sheet_name, PageParams())
                    tsp = pp.table_styles.get(table_name, TableStyleParams())
                    return t, tsp
            break
    return None, None


# ── Page de garde ──────────────────────────────────────────────────────────

def _make_cover_page(model, gparams, primary_color, accent_color, complement_color) -> list:
    doc_title  = getattr(gparams, "doc_title",  "Titre du document")
    doc_author = getattr(gparams, "doc_author", "Auteur non spécifié")

    cover_title_style = ParagraphStyle("CoverTitle",
        fontName="Marianne-Bold", fontSize=24,
        textColor=primary_color, alignment=1, spaceAfter=20)
    cover_author_style = ParagraphStyle("CoverAuthor",
        fontName="Marianne", fontSize=14,
        textColor=accent_color, alignment=1, spaceAfter=10)
    cover_date_style = ParagraphStyle("CoverDate",
        fontName="Marianne", fontSize=12,
        textColor=colors.Color(0.3, 0.3, 0.3), alignment=1)

    content = [
        Spacer(1, 100 * mm),
        Paragraph(doc_title,  cover_title_style),
        Paragraph(doc_author, cover_author_style),
        Paragraph(f"Date : {date.today().strftime('%d/%m/%Y')}", cover_date_style),
        Spacer(1, 10 * mm),
    ]

    table_intro, tsp_intro = _resolve_table_intro(model, gparams)
    if table_intro is not None:
        try:
            default_pp = PageParams()
            avail_w = A4[0] - (default_pp.margin_left + default_pp.margin_right) * mm

            tbl_primary    = _hex_to_rl(tsp_intro.primary_color)    if tsp_intro and tsp_intro.primary_color    else primary_color
            tbl_complement = _hex_to_rl(tsp_intro.complement_color) if tsp_intro and tsp_intro.complement_color else complement_color
            tbl_alt        = _hex_to_rl(tsp_intro.row_alt_color)    if tsp_intro and tsp_intro.row_alt_color    else None
            tbl_hdr_txt    = _hex_to_rl(tsp_intro.header_text_color) if tsp_intro and tsp_intro.header_text_color else None
            tbl_body_txt   = _hex_to_rl(tsp_intro.body_text_color)   if tsp_intro and tsp_intro.body_text_color   else None

            rl_tbl = make_rl_table(
                table_intro, default_pp,
                tbl_primary, accent_color, tbl_complement,
                avail_w,
                row_alt_color=tbl_alt,
                header_text_color=tbl_hdr_txt,
                body_text_color=tbl_body_txt,
                table_style_params=tsp_intro,
            )
            if rl_tbl:
                content += [Spacer(1, 8 * mm), rl_tbl]
        except Exception as e:
            print(f"[cover] table_intro render failed: {e}")

    content.append(PageBreak())
    return content


# ── Sommaire ───────────────────────────────────────────────────────────────

def _make_sommaire(toc_entries, gparams, title_style, accent_color) -> list:
    if not gparams.show_toc or len(toc_entries) <= 1:
        return []

    entries = [
        NextPageTemplate("portrait"),
        Paragraph("Table des matières", title_style),
        HRFlowable(width="100%", thickness=1, color=accent_color, spaceAfter=6),
    ]
    fontSize  = 10
    toc_style = ParagraphStyle("TOCEntry", fontName=font_name, fontSize=fontSize,
                               spaceAfter=4, textColor=colors.black, leading=12)
    max_width = 170 * mm
    dot_width = stringWidth(".", font_name, fontSize)
    for num, name in toc_entries:
        num_str  = str(num)
        name_w   = stringWidth(name,    font_name, fontSize)
        num_w    = stringWidth(num_str, font_name, fontSize)
        num_dots = max(2, math.floor((max_width - name_w - num_w) / dot_width))
        line     = f"{name}{'.' * num_dots}{num_str}"
        while num_dots > 2 and stringWidth(line, font_name, fontSize) > max_width:
            num_dots -= 1
            line = f"{name}{'.' * num_dots} {num_str}"
        entries.append(Paragraph(line, toc_style))

    entries.append(PageBreak())
    return entries


# ── Génération principale ──────────────────────────────────────────────────

def generate_pdf(
    model:       WorkbookModel,
    output_path: str,
    gparams:     GlobalParams | None = None,
    charts:      list[ChartSpec] | None = None,
) -> str:
    if gparams is None:
        gparams = GlobalParams()
    if charts is None:
        charts = []

    primary_color    = _hex_to_rl(gparams.primary_color)
    accent_color     = _hex_to_rl(gparams.accent_color)
    complement_color = _hex_to_rl(gparams.complement_color)
    font_family      = getattr(gparams, "font_family", font_name)

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("LiasseTitle",
        parent=styles["Heading1"], fontSize=gparams.title_font_size,
        fontName=font_name, textColor=primary_color, spaceAfter=6, spaceBefore=0)
    section_style = ParagraphStyle("LiasseSection",
        parent=styles["Heading2"], fontSize=gparams.header_font_size,
        fontName=font_name, textColor=accent_color, spaceAfter=4, spaceBefore=8)

    # ── Application de sheet_settings (include / page_order / display_name) ──
    # sheet_settings est la source de vérité persistée dans le profil ; on la
    # réplique sur les SheetModel pour piloter tri / filtre / titres.
    settings_map = gparams.sheet_settings
    for sheet in model.sheets:
        ss = settings_map.get(sheet.name)
        if ss is not None:
            sheet.include      = ss.include
            sheet.page_order   = ss.page_order
            sheet.display_name = ss.display_name

    ordered_sheets = sorted(
        [s for s in model.sheets if s.include],
        key=lambda s: s.page_order,
    )
    page_map: dict[str, int] = {}

    charts_by_sheet: dict[str, list[ChartSpec]] = {}
    for cs in charts:
        charts_by_sheet.setdefault(cs.sheet_name, []).append(cs)

    # footer_map est rempli pendant _build_sheet_story (passe 1) puis réutilisé
    # tel quel en passe 2 (la pagination est identique entre les deux passes).
    footer_map: dict[int, str] = {}

    def _build_sheet_story() -> list:
        story = []
        for sheet in ordered_sheets:
            pp      = gparams.page_params.get(sheet.name, PageParams())
            avail_w = A4[0] - (pp.margin_left + pp.margin_right) * mm

            ss = settings_map.get(sheet.name, SheetSettings())
            displayed_title = sheet.get_display_name()

            story.append(NextPageTemplate(pp.orientation))

            # PageCollector : enregistre le numéro de page de DÉBUT de la feuille,
            # et propage la note de bas de page à toutes les pages couvertes par
            # cette feuille (mis à jour lors de chaque appel à apply()).
            story.append(_SheetFooterCollector(sheet.name, page_map, footer_map, ss.footer_note))

            if pp.show_sheet_title:
                story.append(Paragraph(displayed_title, title_style))
                story.append(HRFlowable(width="100%", thickness=0.8,
                                        color=accent_color, spaceAfter=4))

            for tz in sheet.tables:
                if not tz.data and not tz.headers:
                    continue

                tsp = pp.table_styles.get(tz.name, TableStyleParams())

                tbl_primary    = _hex_to_rl(tsp.primary_color)    if tsp.primary_color    else primary_color
                tbl_complement = _hex_to_rl(tsp.complement_color) if tsp.complement_color else complement_color
                tbl_alt        = _hex_to_rl(tsp.row_alt_color)    if tsp.row_alt_color    else None
                tbl_hdr_txt    = _hex_to_rl(tsp.header_text_color) if tsp.header_text_color else None
                tbl_body_txt   = _hex_to_rl(tsp.body_text_color)   if tsp.body_text_color   else None

                story.append(Paragraph(tz.name, section_style))
                tbl = make_rl_table(
                    tz, pp,
                    tbl_primary, accent_color, tbl_complement,
                    avail_w, font_family,
                    row_alt_color=tbl_alt,
                    header_text_color=tbl_hdr_txt,
                    body_text_color=tbl_body_txt,
                    table_style_params=tsp,
                )
                if tbl:
                    story.append(KeepTogether([tbl]))
                story.append(Spacer(1, 4 * mm))

            for cs in charts_by_sheet.get(sheet.name, []):
                target = next((t for t in sheet.tables if t.name == cs.table_name), None)
                if target is None:
                    print(f"[pdf] Table '{cs.table_name}' introuvable dans '{sheet.name}'")
                    continue
                img_bytes = _render_chart(cs, target)
                if img_bytes:
                    if cs.title:
                        story.append(Paragraph(cs.title, section_style))
                    story.append(Image(io.BytesIO(img_bytes), width=avail_w, height=avail_w * 0.38))
                    story.append(Spacer(1, 4 * mm))

            story.append(PageBreak())
        return story

    def build_full_story() -> list:
        toc_entries = [(page_map.get(s.name, 0), s.get_display_name()) for s in ordered_sheets]
        return (
            _make_cover_page(model, gparams, primary_color, accent_color, complement_color)
            + _make_sommaire(toc_entries, gparams, title_style, accent_color)
            + _build_sheet_story()
        )

    # Passe 1 : collecte des numéros de page (et de la footer_map)
    doc1 = BaseDocTemplate(io.BytesIO(), pageTemplates=_build_page_templates(gparams, footer_map), title="Liasse")
    doc1.multiBuild(build_full_story())

    # Passe 2 : génération finale (footer_map déjà connue, pagination identique)
    doc2 = BaseDocTemplate(output_path, pageTemplates=_build_page_templates(gparams, footer_map), title="Liasse")
    doc2.multiBuild(build_full_story())

    return output_path


# ── Collecteur de page + note de bas de page ───────────────────────────────

class _SheetFooterCollector(ActionFlowable):
    """
    Variante de PageCollector qui, en plus d'enregistrer le numéro de la première
    page de la feuille dans page_map (pour le sommaire), propage la footer_note
    de la feuille à TOUTES les pages qu'elle occupe.

    ReportLab appelle apply() pour chaque page traversée par ce flowable lors
    du re-flow (en réalité une seule fois au moment où le flowable est placé,
    mais doc.page reflète la page courante à cet instant). Pour couvrir les
    feuilles qui s'étendent sur plusieurs pages, on s'appuie sur le fait que
    multiBuild() effectue un afterFlowable callback à chaque saut de page ;
    ici on adopte une approche simple et robuste : on enregistre la page de
    départ, et on remplit rétroactivement la plage [start, doc.page] à chaque
    nouvel appel d'apply() pour une feuille différente (i.e. dès qu'on quitte
    la feuille courante, on connaît sa dernière page = page précédente).
    """
    def __init__(self, sheet_name: str, page_map: dict, footer_map: dict, footer_note: str | None):
        super().__init__()
        self.sheet_name  = sheet_name
        self.page_map    = page_map
        self.footer_map  = footer_map
        self.footer_note = footer_note

    def apply(self, doc):
        start_page = doc.page
        self.page_map[self.sheet_name] = start_page

        if self.footer_note:
            # Marque au minimum la première page ; les pages suivantes de la même
            # feuille seront couvertes via _extend_previous_sheet_footer ci-dessous.
            self.footer_map[start_page] = self.footer_note

        # Étend la note de la feuille précédente jusqu'à la page précédant celle-ci.
        prev = getattr(doc, "_last_sheet_footer", None)
        if prev is not None:
            prev_note, prev_start = prev
            if prev_note:
                for p in range(prev_start, start_page):
                    self.footer_map[p] = prev_note
        doc._last_sheet_footer = (self.footer_note, start_page)