"""
Construction des tableaux ReportLab depuis les TableZone.

Stratégie de largeur (ordre de priorité) :
  1. Si le tableau tient dans la largeur disponible → on l'affiche tel quel.
  2. Sinon, on essaie d'abord de wrapper le texte dans les colonnes « larges »
     (non-numériques, généralement la colonne de libellé à gauche) pour récupérer
     de la place, sans toucher aux colonnes numériques.
  3. Si le tableau est encore trop large après wrapping, on scale toutes les colonnes
     proportionnellement, avec un plancher de 14 pt par colonne numérique et 40 pt
     pour les colonnes texte.
  4. La police n'est JAMAIS réduite automatiquement ici ; c'est l'auto-fit de page
     (tab_page) qui s'en charge éventuellement.

Règles de wrapping :
  - Les cellules de texte pur (non-numériques) sont wrappées si elles débordent.
  - Les cellules numériques (chiffres, virgule, espace, %, signe) ne sont JAMAIS
    wrappées : on réduit la colonne mais pas la police.
"""
from __future__ import annotations

import re

from reportlab.lib import colors
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import Table, TableStyle

from util import is_year_like

# Regex : chaîne numérique (avec séparateur de milliers, %, signe, virgule, espace)
_RE_NUMERIC = re.compile(r'^[+\-]?[\d\u202f \s.,]+%?$')

# Largeurs plancher (pt)
_MIN_NUMERIC_COL = 14.0   # colonnes de chiffres
_MIN_TEXT_COL    = 40.0   # colonnes de libellés


def _is_numeric_string(s: str) -> bool:
    s = s.strip()
    if not s or s in ("-", "+"):
        return False
    return bool(_RE_NUMERIC.match(s))


def _col_is_numeric(all_rows: list[list], col_idx: int) -> bool:
    """
    True si la colonne est principalement numérique (hors ligne d'en-tête).
    On teste les lignes de données (à partir de l'index 1).
    """
    non_empty = 0
    numeric   = 0
    for row in all_rows[1:]:
        if col_idx >= len(row):
            continue
        cell = str(row[col_idx]).strip() if row[col_idx] is not None else ""
        if not cell:
            continue
        non_empty += 1
        if _is_numeric_string(cell) or is_year_like(row[col_idx]):
            numeric += 1
    if non_empty == 0:
        return False
    return numeric / non_empty >= 0.6   # ≥ 60 % de valeurs numériques → colonne numérique


# ── Formatage pour affichage selon RowStyle ────────────────────────────────

def _apply_row_format(v, row_style) -> str:
    if v is None:
        return ""
    s = str(v)
    if s.lower() == "none" or s == "":
        return ""
    if row_style is None:
        return s

    dec = max(0, min(3, getattr(row_style, "decimal_places", 0)))

    s_parse = s.replace("\u202f", "").replace("\xa0", "").replace(" ", "")
    if "," in s_parse and "." not in s_parse:
        s_parse = s_parse.replace(",", ".")

    if row_style.number_format == "percent":
        s_parse = s_parse.rstrip("%").strip()
        try:
            f = float(s_parse)
        except (ValueError, TypeError):
            return s
        value = f * 100.0 if -1.0 <= f <= 1.0 else f
        rounded = round(value, dec)
        fmt = f"{rounded:.{dec}f}".replace(".", ",")
        return f"{fmt} %"

    # Format normal : ne reformater que si c'est bien un nombre
    if not _is_numeric_string(s):
        return s
    try:
        f = float(s_parse)
    except (ValueError, TypeError):
        return s
    rounded = round(f, dec)
    fmt = f"{rounded:.{dec}f}".replace(".", ",")
    return fmt


# ── Wrapping texte ─────────────────────────────────────────────────────────

def _wrap_text(text: str, font: str, font_size: float, max_width: float) -> str:
    """Insère des \\n si nécessaire. Ne wrape JAMAIS une chaîne numérique."""
    if not text or _is_numeric_string(text):
        return text
    if stringWidth(text, font, font_size) <= max_width:
        return text
    words   = text.split()
    lines, current = [], ""
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


# ── Calcul des largeurs de colonnes ───────────────────────────────────────

def _compute_col_widths(
    all_rows: list[list],
    has_header: bool,
    font_family: str,
    bold_font: str,
    font_size: float,
    available_width: float,
) -> list[float]:
    """
    Calcule les largeurs optimales des colonnes.

    Stratégie (sans toucher à la taille de police) :
    1. Calculer la largeur "naturelle" de chaque colonne.
    2. Si ça tient → ok.
    3. Sinon, réduire les colonnes texte en les wrappant (jusqu'à leur plancher).
    4. Si encore trop large → scale uniforme avec planchers différenciés
       (numérique vs texte).
    """
    ncols = max(len(r) for r in all_rows)
    is_numeric = [_col_is_numeric(all_rows, ci) for ci in range(ncols)]

    # ── Largeurs naturelles ──
    natural = [0.0] * ncols
    for ri, row in enumerate(all_rows):
        is_hdr = ri == 0 and has_header
        fnt    = bold_font if is_hdr else font_family
        for ci in range(ncols):
            cell = str(row[ci]) if ci < len(row) and row[ci] is not None else ""
            longest = max(cell.split("\n"), key=lambda l: stringWidth(l, fnt, font_size), default=cell)
            w = stringWidth(longest, fnt, font_size) + 12
            if w > natural[ci]:
                natural[ci] = w

    total = sum(natural)
    if total <= available_width:
        return natural

    # ── Étape 2 : wrapper les colonnes texte ──
    # On calcule pour chaque colonne texte la largeur minimale après wrapping.
    # Pour les colonnes numériques on garde la largeur naturelle.
    wrapped = natural[:]
    for ci in range(ncols):
        if not is_numeric[ci]:
            # On peut réduire jusqu'à _MIN_TEXT_COL
            wrapped[ci] = max(_MIN_TEXT_COL, natural[ci])

    # Distribuer l'espace excédentaire des colonnes numériques aux colonnes texte
    # en laissant les numériques intactes et en réduisant les texte proportionnellement.
    numeric_total = sum(natural[ci] for ci in range(ncols) if is_numeric[ci])
    text_budget   = available_width - numeric_total
    text_natural  = sum(natural[ci] for ci in range(ncols) if not is_numeric[ci])

    if text_budget > 0 and text_natural > 0:
        scale = min(1.0, text_budget / text_natural)
        for ci in range(ncols):
            if not is_numeric[ci]:
                wrapped[ci] = max(_MIN_TEXT_COL, natural[ci] * scale)

    total_wrapped = sum(wrapped)
    if total_wrapped <= available_width:
        # Petite correction pour remplir exactement
        delta = available_width - total_wrapped
        if wrapped:
            wrapped[-1] += delta
        return wrapped

    # ── Étape 3 : scale uniforme avec planchers ──
    plancher = [_MIN_NUMERIC_COL if is_numeric[ci] else _MIN_TEXT_COL for ci in range(ncols)]
    floored  = [max(plancher[ci], w) for ci, w in enumerate(wrapped)]
    floor_total = sum(floored)

    if floor_total >= available_width:
        # Même les planchers dépassent : on scale tous à partir des planchers
        scale = available_width / floor_total
        result = [max(p, f * scale) for p, f in zip(plancher, floored)]
    else:
        remaining = available_width - floor_total
        surplus   = sum(max(0, w - plancher[ci]) for ci, w in enumerate(wrapped))
        result = []
        for ci in range(ncols):
            extra = max(0, wrapped[ci] - plancher[ci])
            share = (extra / surplus * remaining) if surplus > 0 else 0
            result.append(plancher[ci] + share)

    delta = available_width - sum(result)
    if result:
        result[-1] = max(plancher[-1], result[-1] + delta)
    return result


# ── Construction du tableau ReportLab ─────────────────────────────────────

def make_rl_table(
    table_zone,
    params,
    primary_color,
    accent_color,
    complement_color,
    available_width: float,
    font_family:      str  = "Marianne",
    row_alt_color     = None,
    header_text_color = None,
    body_text_color   = None,
    table_style_params = None,
):
    bold_font   = f"{font_family}-Bold"
    italic_font = f"{font_family}-Italic"

    eff_hdr_text  = header_text_color if header_text_color is not None else colors.white
    eff_body_text = body_text_color   if body_text_color   is not None else colors.black
    eff_alt       = row_alt_color     if row_alt_color     is not None else colors.Color(0.95, 0.97, 1.0)

    row_styles_map: dict[int, object] = {}
    if table_style_params is not None:
        row_styles_map = table_style_params.row_styles or {}

    # ── Détection de la ligne "années" (subheader) ────────────────────────
    # La ligne d'index 0 dans data (2ème ligne du tableau) est un subheader
    # si elle contient majoritairement des valeurs ressemblant à des années.
    year_row_data_index: int | None = None
    if table_zone.data and len(table_zone.data) > 0:
        first_data_row = table_zone.data[0]
        nb_yr = sum(1 for v in first_data_row if is_year_like(v))
        if nb_yr > 1:
            year_row_data_index = 0

    # ── Filtrage des lignes invisibles + formatage ─────────────────────────
    visible_rows: list[tuple[int, list]] = []
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

    # ── Largeurs ───────────────────────────────────────────────────────────
    col_widths = _compute_col_widths(
        all_rows,
        bool(table_zone.headers),
        font_family, bold_font,
        params.font_size,
        available_width,
    )

    # ── Wrapping des cellules texte avec les largeurs finales ──────────────
    is_numeric_col = [_col_is_numeric(all_rows, ci) for ci in range(ncols)]
    for ri, row in enumerate(all_rows):
        is_hdr = ri == 0 and bool(table_zone.headers)
        fnt    = bold_font if is_hdr else font_family
        for ci in range(ncols):
            if ci >= len(row):
                continue
            if is_numeric_col[ci]:
                continue   # jamais de wrapping sur les colonnes numériques
            cell    = str(row[ci]) if row[ci] is not None else ""
            cw      = col_widths[ci] if ci < len(col_widths) else col_widths[-1]
            wrapped = _wrap_text(cell, fnt, params.font_size, cw - 12)
            all_rows[ri][ci] = wrapped

    # ── Commandes de style de base ─────────────────────────────────────────
    style_cmds = [
        ("FONTNAME",       (0, 1), (-1, -1), font_family),
        ("FONTSIZE",       (0, 1), (-1, -1), params.font_size),
        ("LEADING",        (0, 1), (-1, -1), params.font_size * 1.3),
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

    # ── Fusion de cellules d'en-tête (préfixe "BE" / "Ecarts" / …) ───────
    if table_zone.headers and len(all_rows[0]) >= 3:
        i = 0
        while i < len(all_rows[0]):
            val = str(all_rows[0][i]).strip() if all_rows[0][i] is not None else ""
            prefixes = ("BE", "Ecarts", "Partage de la valeur ajoutée",
                        "Inflation et choc de prix importés",
                        "Assiette retardées", "Assiettes peu")
            if any(val.lower().startswith(p.lower()) for p in prefixes):
                c1 = all_rows[0][i+1] if len(all_rows[0]) > i+1 else None
                c2 = all_rows[0][i+2] if len(all_rows[0]) > i+2 else None
                e1 = c1 is None or str(c1).strip() in ("", "None")
                e2 = c2 is None or str(c2).strip() in ("", "None")
                if e1 and e2:
                    style_cmds.append(("SPAN",  (i, 0), (i+2, 0)))
                    style_cmds.append(("ALIGN", (i, 0), (i+2, 0), "CENTER"))
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

        if rs.background_color:
            try:
                h = rs.background_color.lstrip("#")
                r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
                bg = colors.Color(r/255, g/255, b/255)
                style_cmds.append(("BACKGROUND", (0, tbl_row), (-1, tbl_row), bg))
                luminance = 0.299*r + 0.587*g + 0.114*b
                txt_color = colors.white if luminance < 128 else eff_body_text
                style_cmds.append(("TEXTCOLOR", (0, tbl_row), (-1, tbl_row), txt_color))
            except Exception:
                pass

        if rs.text_style == "bold":
            style_cmds.append(("FONTNAME", (0, tbl_row), (-1, tbl_row), bold_font))
        elif rs.text_style == "italic":
            style_cmds.append(("FONTNAME", (0, tbl_row), (-1, tbl_row), italic_font))

        if rs.font_size is not None:
            style_cmds.append(("FONTSIZE", (0, tbl_row), (-1, tbl_row), rs.font_size))
            style_cmds.append(("LEADING",  (0, tbl_row), (-1, tbl_row), rs.font_size * 1.3))

    # ── En-tête ────────────────────────────────────────────────────────────
    style_cmds.extend([
        ("BACKGROUND", (1, 0), (-1, 0), primary_color),
        ("TEXTCOLOR",  (1, 0), (-1, 0), eff_hdr_text),
        ("FONTNAME",   (1, 0), (-1, 0), bold_font),
        ("FONTSIZE",   (1, 0), (-1, 0), params.font_size),
        ("LEADING",    (1, 0), (-1, 0), params.font_size * 1.3),
    ])

    top_left = str(all_rows[0][0]).strip() if all_rows and all_rows[0] else ""
    transparent = [
        ("BACKGROUND", (0, 0), (0, 0), colors.Color(1, 1, 1, 0)),
        ("TEXTCOLOR",  (0, 0), (0, 0), colors.Color(1, 1, 1, 0)),
        ("LINEABOVE",  (0, 0), (0, 0), 0, colors.Color(1, 1, 1, 0)),
        ("LINEBELOW",  (0, 0), (0, 0), 0, colors.Color(1, 1, 1, 0)),
        ("LINEBEFORE", (0, 0), (0, 0), 0, colors.Color(1, 1, 1, 0)),
        ("LINEAFTER",  (0, 0), (0, 0), 0, colors.Color(1, 1, 1, 0)),
    ]
    if top_left == "":
        style_cmds.extend(transparent)
    else:
        style_cmds.extend([
            ("BACKGROUND", (0, 0), (0, 0), primary_color),
            ("TEXTCOLOR",  (0, 0), (0, 0), eff_hdr_text),
            ("FONTNAME",   (0, 0), (0, 0), bold_font),
            ("FONTSIZE",   (0, 0), (0, 0), params.font_size),
            ("LEADING",    (0, 0), (0, 0), params.font_size * 1.3),
            ("LINEBELOW",  (0, 0), (0, 0), 1.2, accent_color),
        ])

    # Deuxième ligne (index 1 dans all_rows) : cellule (0,1) transparente si vide
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

    # ── Ligne "années" (subheader) ─────────────────────────────────────────
    # Identifiée via year_row_data_index ; on cherche sa position dans all_rows
    if year_row_data_index is not None:
        # Position dans all_rows = décalage dû à l'en-tête + position dans visible_rows
        yr_tbl_row = None
        for vis_idx, (data_idx, _) in enumerate(visible_rows):
            if data_idx == year_row_data_index:
                yr_tbl_row = vis_idx + (1 if table_zone.headers else 0)
                break
        if yr_tbl_row is not None:
            style_cmds.extend([
                ("FONTNAME",   (0, yr_tbl_row), (-1, yr_tbl_row), bold_font),
                ("LEADING",    (0, yr_tbl_row), (-1, yr_tbl_row), params.font_size * 1.3),
                ("BACKGROUND", (0, yr_tbl_row), (-1, yr_tbl_row), complement_color),
                ("TEXTCOLOR",  (0, yr_tbl_row), (-1, yr_tbl_row), eff_hdr_text),
            ])

    tbl = Table(all_rows, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle(style_cmds))
    return tbl


# ── Fonction exportée pour détecter le subheader depuis tab_line ──────────

def detect_year_row_index(table_zone) -> int | None:
    """
    Retourne l'index dans table_zone.data de la ligne "années" (subheader),
    ou None si elle n'existe pas.
    Utilisé par tab_line.py pour l'exclure de la liste de lignes éditables
    et du "appliquer à toutes".
    """
    if not table_zone.data:
        return None
    first_row = table_zone.data[0]
    nb_yr = sum(1 for v in first_row if is_year_like(v))
    return 0 if nb_yr > 1 else None