"""
Construction des tableaux ReportLab depuis les TableZone.

Règles de wrapping :
  - Les cellules de texte pur (non-numériques) sont wrappées si elles débordent.
  - Les cellules numériques (contenant chiffres, virgule, espace, %, signe)
    ne sont JAMAIS wrappées : on réduit la police si nécessaire, pas la largeur.

Formatage des nombres via RowStyle :
  - number_format == "normal"  → valeur telle que produite par excel_parser
  - number_format == "percent" → réinterprétation en pourcentage
  - decimal_places              → précision (0..3)
"""
from __future__ import annotations

import re

from reportlab.lib import colors
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import Table, TableStyle

from util import is_year_like

# Regex : une chaîne est numérique si elle ne contient que chiffres,
# signe ±, virgule, point, espace et %
_RE_NUMERIC = re.compile(r'^[+\-]?[\d \s.,]+%?$')


def _is_numeric_string(s: str) -> bool:
    """True si la chaîne représente un nombre (avec séparateur de milliers, %, …)."""
    s = s.strip()
    if not s or s in ("-", "+"):
        return False
    return bool(_RE_NUMERIC.match(s))


# ── Formatage pour affichage selon RowStyle ────────────────────────────────

def _apply_row_format(v, row_style) -> str:
    """
    Convertit une valeur en chaîne d'affichage selon le RowStyle de sa ligne.
    SI number_format == "normal"  → chaîne telle quelle (déjà formatée par excel_parser).
    SI number_format == "percent" → interprète la valeur brute comme un ratio (0-1) si elle est dans [-1, 1], sinon comme un pourcentage direct.
    Affichage : "xx,x %" avec decimal_places décimales.
    """
    if v is None:
        return ""
    s = str(v)
    if s.lower() == "none" or s == "":
        return ""

    # If no RowStyle or not numeric, return as-is
    if row_style is None:
        return s

    dec = max(0, min(3, getattr(row_style, "decimal_places", 0)))

    # Try to parse the value (French format possible: "1,23").
    s_parse = s.replace("\xa0", "").replace(" ", "")
    # If contains comma as decimal separator and no dot, normalize to dot for float()
    if "," in s_parse and "." not in s_parse:
        s_parse = s_parse.replace(",", ".")

    # Percent handling: strip trailing % and interpret ratio vs direct percent
    if row_style.number_format == "percent":
        s_parse = s_parse.rstrip("%").strip()
        try:
            f = float(s_parse)
        except (ValueError, TypeError):
            return s

        # Ratio → multiplier by 100
        if -1.0 <= f <= 1.0:
            value = f * 100.0
        else:
            value = f

        rounded = round(value, dec)
        fmt = f"{rounded:.{dec}f}"
        # Use space as thousands sep and comma as decimal sep
        fmt = fmt.replace(",", " ").replace(".", ",")
        return f"{fmt}%"

    # Normal number formatting
    # Do not alter non-numeric strings
    if not _is_numeric_string(s):
        return s

    try:
        f = float(s_parse)
    except (ValueError, TypeError):
        return s

    rounded = round(f, dec)
    fmt = f"{rounded:.{dec}f}"
    fmt = fmt.replace(",", " ").replace(".", ",")
    return fmt


# ── Wrapping texte (jamais sur les numériques) ─────────────────────────────

def _wrap_text(text: str, font: str, font_size: float, max_width: float) -> str:
    """
    Insère des \\n dans `text` si nécessaire, en respectant max_width.
    Ne wrape JAMAIS une chaîne numérique.
    """
    if not text or _is_numeric_string(text):
        return text
    if stringWidth(text, font, font_size) <= max_width:
        return text

    words   = text.split()
    lines   = []
    current = ""
    for word in words:
        test = f"{current} {word}".strip() if current else word
        if stringWidth(test, font, font_size) <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return "\n".join(lines) if lines else text


# ── Construction du tableau ReportLab ─────────────────────────────────────

def make_rl_table(
    table_zone,
    params,
    primary_color,
    accent_color,
    complement_color,
    available_width: float,
    font_family:     str  = "Marianne",
    row_alt_color    = None,
    header_text_color = None,
    body_text_color   = None,
    table_style_params = None,
):
    """
    Construit un tableau ReportLab stylisé depuis un TableZone.

    Paramètres
    ----------
    table_zone         : TableZone  — données (headers + data)
    params             : PageParams — font_size, marges …
    primary_color      : RL Color   — fond de l'en-tête (déjà résolu)
    accent_color       : RL Color   — utilisé pour LINEBELOW sur cellule top-left
    complement_color   : RL Color   — ligne "années"
    available_width    : float      — largeur disponible en points
    font_family        : str
    row_alt_color      : RL Color | None
    header_text_color  : RL Color | None
    body_text_color    : RL Color | None
    table_style_params : TableStyleParams | None — contient row_styles
    """
    bold_font   = f"{font_family}-Bold"
    italic_font = f"{font_family}-Italic"

    eff_hdr_text  = header_text_color if header_text_color is not None else colors.white
    eff_body_text = body_text_color   if body_text_color   is not None else colors.black
    eff_alt       = row_alt_color     if row_alt_color     is not None else colors.Color(0.95, 0.97, 1.0)

    row_styles_map: dict[int, object] = {}
    if table_style_params is not None:
        row_styles_map = table_style_params.row_styles or {}

    # ── Filtrage des lignes invisibles + formatage ─────────────────────────
    visible_rows: list[tuple[int, list]] = []   # (data_index, formatted_row)
    for data_idx, row in enumerate(table_zone.data):
        rs = row_styles_map.get(data_idx)
        if rs is not None and not rs.visible:
            continue
        formatted = [_apply_row_format(v, rs) for v in row]
        visible_rows.append((data_idx, formatted))

    # ── Construction de all_rows ───────────────────────────────────────────
    all_rows: list[list] = []
    if table_zone.headers:
        all_rows.append(list(table_zone.headers))
    for _, row in visible_rows:
        all_rows.append(row)

    if not all_rows:
        return None

    ncols = max(len(r) for r in all_rows)

    # ── Calcul des largeurs de colonne ─────────────────────────────────────
    MAX_COL_RATIO = 0.50   # une colonne ne peut pas dépasser 50 % de la largeur totale

    raw_widths = [0.0] * ncols
    for ri, row in enumerate(all_rows):
        is_hdr = ri == 0 and bool(table_zone.headers)
        fnt    = bold_font if is_hdr else font_family
        for ci in range(ncols):
            cell = str(row[ci]) if ci < len(row) and row[ci] is not None else ""
            # Pour le calcul, on prend la ligne la plus longue en cas de \n existant
            longest = max(cell.split("\n"), key=lambda l: stringWidth(l, fnt, params.font_size), default=cell)
            w = stringWidth(longest, fnt, params.font_size) + 12
            if w > raw_widths[ci]:
                raw_widths[ci] = w

    total = sum(raw_widths)
    if total <= 0:
        col_widths = [available_width / ncols] * ncols
    elif total <= available_width:
        col_widths = raw_widths[:]
    else:
        max_col_w = available_width * MAX_COL_RATIO
        scale     = available_width / total
        col_widths = [min(max_col_w, max(18.0, w * scale)) for w in raw_widths]
        delta      = available_width - sum(col_widths)
        if col_widths:
            col_widths[-1] = max(18.0, col_widths[-1] + delta)

    # ── Wrapping des cellules texte ────────────────────────────────────────
    for ri, row in enumerate(all_rows):
        is_hdr = ri == 0 and bool(table_zone.headers)
        fnt    = bold_font if is_hdr else font_family
        for ci in range(ncols):
            if ci >= len(row):
                continue
            cell = str(row[ci]) if row[ci] is not None else ""
            cw   = col_widths[ci] if ci < len(col_widths) else col_widths[-1]
            wrapped = _wrap_text(cell, fnt, params.font_size, cw - 12)
            all_rows[ri][ci] = wrapped

    # ── Commandes de style de base ─────────────────────────────────────────
    style_cmds = [
        ("FONTNAME",       (0, 1), (-1, -1), font_family),
        ("FONTSIZE",       (0, 1), (-1, -1), params.font_size),
        ("LEADING",        (0, 1), (-1, -1), params.font_size * 1.2),
        ("TEXTCOLOR",      (0, 1), (-1, -1), eff_body_text),
        ("ALIGN",          (0, 0), (-1, -1), "LEFT"),
        ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, eff_alt]),
        ("GRID",           (0, 1), (-1, -1), 0.3, colors.Color(0.7, 0.7, 0.7)),
        ("TOPPADDING",     (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), 3),
        ("LEFTPADDING",    (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",   (0, 0), (-1, -1), 5),
    ]

    # ── Fusion de cellules d'en-tête (préfixe "BE" / "Ecarts") ───────────
    if table_zone.headers and len(all_rows[0]) >= 3:
        i = 0
        while i < len(all_rows[0]):
            val = str(all_rows[0][i]).strip() if all_rows[0][i] is not None else ""
            if any(val.lower().startswith(p.lower()) for p in ("BE", "Ecarts", "Partage de la valeur ajoutée", "Inflation et choc de prix importés", "Assiette retardées", "Assiettes peu")):
                c1 = all_rows[0][i+1] if len(all_rows[0]) > i+1 else None
                c2 = all_rows[0][i+2] if len(all_rows[0]) > i+2 else None
                e1 = c1 is None or str(c1).strip() in ("", "None")
                e2 = c2 is None or str(c2).strip() in ("", "None")
                if e1 and e2:
                    style_cmds.append(("SPAN",  (i, 0), (i+2, 0)))
                    style_cmds.append(("ALIGN", (i, 0), (i+2, 0), "CENTER"))
                    # Bordure droite identique à la couleur de l'en-tête
                    if i + 2 < len(all_rows[0]) - 3:
                        style_cmds.append(("LINEAFTER", (i+2, 1), (i+2, -1), 1, primary_color))
                    i += 3
                    continue
            i += 1

    # ── RowStyle : application ligne par ligne ─────────────────────────────
    for vis_idx, (data_idx, _) in enumerate(visible_rows):
        tbl_row = vis_idx + (1 if table_zone.headers else 0)
        rs = row_styles_map.get(data_idx)
        if rs is None:
            continue

        # Couleur de fond
        if rs.background_color:
            try:
                h = rs.background_color.lstrip("#")
                r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
                bg = colors.Color(r/255, g/255, b/255)
                style_cmds.append(("BACKGROUND", (0, tbl_row), (-1, tbl_row), bg))
                # Texte blanc si fond sombre
                luminance = 0.299*r + 0.587*g + 0.114*b
                txt_color = colors.white if luminance < 128 else eff_body_text
                style_cmds.append(("TEXTCOLOR", (0, tbl_row), (-1, tbl_row), txt_color))
            except Exception:
                pass

        # Style texte
        if rs.text_style == "bold":
            style_cmds.append(("FONTNAME", (0, tbl_row), (-1, tbl_row), bold_font))
        elif rs.text_style == "italic":
            style_cmds.append(("FONTNAME", (0, tbl_row), (-1, tbl_row), italic_font))

        # Taille de police
        if rs.font_size is not None:
            style_cmds.append(("FONTSIZE", (0, tbl_row), (-1, tbl_row), rs.font_size))

    # ── En-tête ────────────────────────────────────────────────────────────
    style_cmds.extend([
        ("BACKGROUND", (1, 0), (-1, 0), primary_color),
        ("TEXTCOLOR",  (1, 0), (-1, 0), eff_hdr_text),
        ("FONTNAME",   (1, 0), (-1, 0), bold_font),
        ("FONTSIZE",   (1, 0), (-1, 0), params.font_size),
        ("LEADING",    (1, 0), (-1, 0), params.font_size * 1.2),
    ])

    # Cellule top-left : transparente si vide, sinon stylisée comme le reste
    top_left = str(all_rows[0][0]).strip() if all_rows and all_rows[0] else ""
    if top_left == "":
        style_cmds.extend([
            ("BACKGROUND", (0, 0), (0, 0), colors.Color(1, 1, 1, 0)),
            ("TEXTCOLOR",  (0, 0), (0, 0), colors.Color(1, 1, 1, 0)),
            ("LINEABOVE",  (0, 0), (0, 0), 0, colors.Color(1, 1, 1, 0)),
            ("LINEBELOW",  (0, 0), (0, 0), 0, colors.Color(1, 1, 1, 0)),
            ("LINEBEFORE", (0, 0), (0, 0), 0, colors.Color(1, 1, 1, 0)),
            ("LINEAFTER",  (0, 0), (0, 0), 0, colors.Color(1, 1, 1, 0)),
        ])
    else:
        style_cmds.extend([
            ("BACKGROUND", (0, 0), (0, 0), primary_color),
            ("TEXTCOLOR",  (0, 0), (0, 0), eff_hdr_text),
            ("FONTNAME",   (0, 0), (0, 0), bold_font),
            ("FONTSIZE",   (0, 0), (0, 0), params.font_size),
            ("LEADING",    (0, 0), (0, 0), params.font_size * 1.2),
            ("LINEBELOW",  (0, 0), (0, 0), 1.2, accent_color),
        ])

    # Deuxième ligne : cellule (0,1) transparente si vide
    if len(all_rows) > 1:
        sec_first = str(all_rows[1][0]).strip() if all_rows[1] and all_rows[1][0] is not None else ""
        if sec_first == "":
            style_cmds.extend([
                ("BACKGROUND", (0, 1), (0, 1), colors.Color(1, 1, 1, 0)),
                ("TEXTCOLOR",  (0, 1), (0, 1), colors.Color(1, 1, 1, 0)),
                ("LINEABOVE",  (0, 1), (0, 1), 0, colors.Color(1, 1, 1, 0)),
                ("LINEBELOW",  (0, 1), (0, 1), 0, colors.Color(1, 1, 1, 0)),
                ("LINEBEFORE", (0, 1), (0, 1), 0, colors.Color(1, 1, 1, 0)),
                ("LINEAFTER",  (0, 1), (0, 1), 0, colors.Color(1, 1, 1, 0)),
            ])

    # Ligne "années" : fond complement_color si >1 valeur ressemble à une année
    if len(all_rows) > 1:
        nb_yr = sum(1 for v in all_rows[1] if is_year_like(v))
        if nb_yr > 1:
            style_cmds.extend([
                ("FONTNAME",   (0, 1), (-1, 1), bold_font),
                ("LEADING",    (0, 1), (-1, 1), params.font_size * 1.2),
                ("BACKGROUND", (0, 1), (-1, 1), complement_color),
                ("TEXTCOLOR",  (0, 1), (-1, 1), eff_hdr_text),
            ])

    tbl = Table(all_rows, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle(style_cmds))
    return tbl