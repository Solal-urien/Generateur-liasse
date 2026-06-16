"""
Gestionnaire d'état central — pattern Observer.
Gère la sérialisation / désérialisation complète des RowStyle, TableStyleParams,
SheetSettings et de la référence du tableau d'introduction.
"""
import json
import os
import dataclasses
from dataclasses import dataclass, field
from typing import Callable, Any

from excel_parser import WorkbookModel, parse_workbook, load_mapping
from pdf_generator import (
    GlobalParams, PageParams, ChartSpec,
    TableStyleParams, RowStyle, SheetSettings,
)

import util


@dataclass
class AppState:
    workbook:            WorkbookModel | None = None
    gparams:             GlobalParams          = field(default_factory=GlobalParams)
    charts:              list[ChartSpec]       = field(default_factory=list)
    current_excel_path:  str | None            = None
    mapping:             dict                  = field(default_factory=dict)


class StateManager:
    def __init__(self):
        self._state     = AppState()
        self._listeners: dict[str, list[Callable]] = {}

    # ── Accesseurs ─────────────────────────────────────────────────────────

    @property
    def state(self)    -> AppState:            return self._state
    @property
    def workbook(self) -> WorkbookModel | None: return self._state.workbook
    @property
    def gparams(self)  -> GlobalParams:         return self._state.gparams
    @property
    def charts(self)   -> list[ChartSpec]:      return self._state.charts

    # ── Observer ───────────────────────────────────────────────────────────

    def subscribe(self, event: str, callback: Callable):
        self._listeners.setdefault(event, []).append(callback)

    def emit(self, event: str, data: Any = None):
        for cb in self._listeners.get(event, []):
            try:
                cb(data)
            except Exception as e:
                print(f"[StateManager] listener error ({event}): {e}")

    # ── Chargement Excel ───────────────────────────────────────────────────

    def load_excel(self, path: str, mapping_path: str | None = None) -> bool:
        try:
            mapping = {}
            if "FIPU2" in path:
                mapping = util.mapping_fipu2
            if mapping_path and os.path.exists(mapping_path):
                mapping = load_mapping(mapping_path)
                self._state.mapping = mapping

            self._state.workbook           = parse_workbook(path, mapping or None)
            self._state.current_excel_path = path

            gp = self._state.gparams

            for sheet in self._state.workbook.sheets:
                # PageParams (style)
                if sheet.name not in gp.page_params:
                    gp.page_params[sheet.name] = PageParams()

                # SheetSettings (document) — ne pas écraser si déjà présent (profil chargé avant Excel)
                if sheet.name not in gp.sheet_settings:
                    gp.sheet_settings[sheet.name] = SheetSettings(
                        include=sheet.include,
                        page_order=sheet.page_order,
                        display_name=None,
                        footer_note=None,
                    )
                else:
                    # Appliquer les réglages du profil sur le SheetModel
                    ss = gp.sheet_settings[sheet.name]
                    sheet.include      = ss.include
                    sheet.page_order   = ss.page_order
                    sheet.display_name = ss.display_name

            self.emit("workbook_loaded", self._state.workbook)
            return True
        except Exception as e:
            self.emit("error", str(e))
            return False

    # ── Ordre / inclusion des feuilles ─────────────────────────────────────

    def update_sheet_order(self, new_order: list[str]):
        if not self._state.workbook:
            return
        idx_map = {name: i for i, name in enumerate(new_order)}
        for sheet in self._state.workbook.sheets:
            new_idx = idx_map.get(sheet.name, sheet.page_order)
            sheet.page_order = new_idx
            self._state.gparams.sheet_settings.setdefault(
                sheet.name, SheetSettings()
            ).page_order = new_idx
        self.emit("sheet_order_changed", new_order)

    def toggle_sheet_include(self, sheet_name: str, included: bool):
        if not self._state.workbook:
            return
        for sheet in self._state.workbook.sheets:
            if sheet.name == sheet_name:
                sheet.include = included
                self._state.gparams.sheet_settings.setdefault(
                    sheet_name, SheetSettings()
                ).include = included
        self.emit("sheet_visibility_changed", sheet_name)

    # ── Nom affiché / note de bas de page ─────────────────────────────────

    def set_display_name(self, sheet_name: str, display_name: str | None):
        """Modifie le nom affiché d'une page (titre PDF + sommaire)."""
        ss = self._state.gparams.sheet_settings.setdefault(sheet_name, SheetSettings())
        ss.display_name = display_name or None
        if self._state.workbook:
            for sheet in self._state.workbook.sheets:
                if sheet.name == sheet_name:
                    sheet.display_name = ss.display_name
                    break
        self.emit("sheet_settings_changed", sheet_name)

    def set_footer_note(self, sheet_name: str, note: str | None):
        """Modifie la note de bas de page d'une page."""
        ss = self._state.gparams.sheet_settings.setdefault(sheet_name, SheetSettings())
        ss.footer_note = note or None
        self.emit("sheet_settings_changed", sheet_name)

    def get_sheet_settings(self, sheet_name: str) -> SheetSettings:
        return self._state.gparams.sheet_settings.get(sheet_name, SheetSettings())

    # ── PageParams (style) ─────────────────────────────────────────────────

    def update_page_params(self, sheet_name: str, pp: PageParams):
        self._state.gparams.page_params[sheet_name] = pp
        self.emit("page_params_changed", sheet_name)

    def update_global_params(self, gp: GlobalParams):
        self._state.gparams = gp
        self.emit("global_params_changed", None)

    # ── Année de départ ────────────────────────────────────────────────────

    def set_start_year(self, year: int):
        self._state.gparams.start_year = year

    def get_start_year(self) -> int:
        return getattr(self._state.gparams, "start_year", 2025)

    # ── Titre / auteur ─────────────────────────────────────────────────────

    def set_doc_title(self, title: str):   self._state.gparams.doc_title  = title
    def set_doc_author(self, author: str): self._state.gparams.doc_author = author
    def get_doc_title(self)  -> str: return getattr(self._state.gparams, "doc_title",  "")
    def get_doc_author(self) -> str: return getattr(self._state.gparams, "doc_author", "")

    # ── Tableau d'introduction (référence par nom) ─────────────────────────

    def set_table_intro_ref(self, sheet_name: str | None, table_name: str | None):
        """Enregistre la référence (sheet_name, table_name) du tableau d'intro."""
        if sheet_name and table_name:
            self._state.gparams.table_intro_ref = (sheet_name, table_name)
        else:
            self._state.gparams.table_intro_ref = None
        self.emit("global_params_changed", None)

    def get_table_intro_ref(self) -> tuple[str, str] | None:
        return getattr(self._state.gparams, "table_intro_ref", None)

    # ── TableStyleParams ───────────────────────────────────────────────────

    def get_table_style(self, sheet_name: str, table_name: str) -> TableStyleParams:
        pp = self._state.gparams.page_params.setdefault(sheet_name, PageParams())
        return pp.table_styles.setdefault(table_name, TableStyleParams())

    def update_table_style(self, sheet_name: str, table_name: str, tsp: TableStyleParams):
        pp = self._state.gparams.page_params.setdefault(sheet_name, PageParams())
        pp.table_styles[table_name] = tsp
        self.emit("page_params_changed", sheet_name)

    # ── Charts ─────────────────────────────────────────────────────────────

    def add_chart(self, spec: ChartSpec):
        self._state.charts.append(spec)
        self.emit("charts_changed", self._state.charts)

    def remove_chart(self, index: int):
        if 0 <= index < len(self._state.charts):
            self._state.charts.pop(index)
            self.emit("charts_changed", self._state.charts)

    def update_chart(self, index: int, spec: ChartSpec):
        if 0 <= index < len(self._state.charts):
            self._state.charts[index] = spec
            self.emit("charts_changed", self._state.charts)

    # ── Persistance JSON ───────────────────────────────────────────────────

    def save_profile(self, path: str):
        gp = self._state.gparams

        # ── GlobalParams (champs dataclass uniquement) ──
        gp_fields = {f.name for f in dataclasses.fields(GlobalParams)}
        gp_d = {f: getattr(gp, f) for f in gp_fields}

        # Attributs dynamiques (non-dataclass)
        gp_d["doc_title"]  = getattr(gp, "doc_title",  "")
        gp_d["doc_author"] = getattr(gp, "doc_author", "")

        # table_intro_ref : tuple → list (JSON-compatible)
        ref = getattr(gp, "table_intro_ref", None)
        gp_d["table_intro_ref"] = list(ref) if ref else None

        # ── page_params : sérialisation TableStyleParams → RowStyle ──
        pp_serial = {}
        for sname, pp in gp.page_params.items():
            pp_fields = {f.name for f in dataclasses.fields(PageParams)}
            pp_d = {f: getattr(pp, f) for f in pp_fields if f != "table_styles"}
            ts_serial = {}
            for tname, tsp in pp.table_styles.items():
                tsp_fields = {f.name for f in dataclasses.fields(TableStyleParams)}
                tsp_d = {f: getattr(tsp, f) for f in tsp_fields if f != "row_styles"}
                tsp_d["row_styles"] = {
                    str(k): dataclasses.asdict(rs)
                    for k, rs in tsp.row_styles.items()
                }
                ts_serial[tname] = tsp_d
            pp_d["table_styles"] = ts_serial
            pp_serial[sname] = pp_d
        gp_d["page_params"] = pp_serial

        # ── sheet_settings ──
        ss_serial = {}
        for sname, ss in gp.sheet_settings.items():
            ss_serial[sname] = dataclasses.asdict(ss)
        gp_d["sheet_settings"] = ss_serial

        profile = {
            "gparams": gp_d,
            "charts":  [dataclasses.asdict(c) for c in self._state.charts],
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(profile, f, ensure_ascii=False, indent=2)

    def load_profile(self, path: str):
        with open(path, "r", encoding="utf-8") as f:
            profile = json.load(f)

        gp_data = profile.get("gparams", {})

        # Extraire les sous-structures avant de construire GlobalParams
        pp_data          = gp_data.pop("page_params",    {})
        ss_data          = gp_data.pop("sheet_settings", {})
        doc_title        = gp_data.pop("doc_title",      "")
        doc_author       = gp_data.pop("doc_author",     "")
        table_intro_ref  = gp_data.pop("table_intro_ref", None)

        gp_fields = {f.name for f in dataclasses.fields(GlobalParams)}
        gp = GlobalParams(**{k: v for k, v in gp_data.items() if k in gp_fields})
        gp.doc_title  = doc_title
        gp.doc_author = doc_author

        # table_intro_ref : list → tuple
        gp.table_intro_ref = tuple(table_intro_ref) if table_intro_ref else None

        # ── Reconstruction PageParams → TableStyleParams → RowStyle ──
        tsp_fields = {f.name for f in dataclasses.fields(TableStyleParams)}
        rs_fields  = {f.name for f in dataclasses.fields(RowStyle)}
        pp_fields  = {f.name for f in dataclasses.fields(PageParams) if f != "table_styles"}

        for sname, pp_dict in pp_data.items():
            ts_data = pp_dict.pop("table_styles", {})
            pp = PageParams(**{k: v for k, v in pp_dict.items() if k in pp_fields})
            for tname, tsp_dict in ts_data.items():
                rs_data = tsp_dict.pop("row_styles", {})
                tsp = TableStyleParams(**{k: v for k, v in tsp_dict.items() if k in tsp_fields})
                for row_key, rs_dict in rs_data.items():
                    tsp.row_styles[int(row_key)] = RowStyle(
                        **{k: v for k, v in rs_dict.items() if k in rs_fields}
                    )
                pp.table_styles[tname] = tsp
            gp.page_params[sname] = pp

        # ── Reconstruction SheetSettings ──
        ss_fields = {f.name for f in dataclasses.fields(SheetSettings)}
        for sname, ss_dict in ss_data.items():
            gp.sheet_settings[sname] = SheetSettings(
                **{k: v for k, v in ss_dict.items() if k in ss_fields}
            )

        self._state.gparams = gp

        # Appliquer immédiatement sheet_settings aux SheetModel si le workbook est déjà chargé
        if self._state.workbook:
            for sheet in self._state.workbook.sheets:
                ss = gp.sheet_settings.get(sheet.name)
                if ss:
                    sheet.include      = ss.include
                    sheet.page_order   = ss.page_order
                    sheet.display_name = ss.display_name

        # ── Charts ──
        cf = {f.name for f in dataclasses.fields(ChartSpec)}
        self._state.charts = [
            ChartSpec(**{k: v for k, v in c.items() if k in cf})
            for c in profile.get("charts", [])
        ]
        self.emit("profile_loaded", None)


# Singleton global
state_manager = StateManager()