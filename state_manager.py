"""
Gestionnaire d'état central — pattern Observer.
Gère la sérialisation / désérialisation complète des RowStyle et TableStyleParams.
"""
import json
import os
import dataclasses
from dataclasses import dataclass, field
from typing import Callable, Any

from excel_parser import WorkbookModel, parse_workbook, load_mapping
from pdf_generator import GlobalParams, PageParams, ChartSpec, TableStyleParams, RowStyle

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

    # ── Actions ────────────────────────────────────────────────────────────

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

            for sheet in self._state.workbook.sheets:
                if sheet.name not in self._state.gparams.page_params:
                    self._state.gparams.page_params[sheet.name] = PageParams()

            self.emit("workbook_loaded", self._state.workbook)
            return True
        except Exception as e:
            self.emit("error", str(e))
            return False

    def update_sheet_order(self, new_order: list[str]):
        if not self._state.workbook:
            return
        idx_map = {name: i for i, name in enumerate(new_order)}
        for sheet in self._state.workbook.sheets:
            sheet.page_order = idx_map.get(sheet.name, sheet.page_order)
        self.emit("sheet_order_changed", new_order)

    def toggle_sheet_include(self, sheet_name: str, included: bool):
        if not self._state.workbook:
            return
        for sheet in self._state.workbook.sheets:
            if sheet.name == sheet_name:
                sheet.include = included
        self.emit("sheet_visibility_changed", sheet_name)

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

    def set_doc_title(self, title: str):  self._state.gparams.doc_title  = title
    def set_doc_author(self, author: str): self._state.gparams.doc_author = author
    def get_doc_title(self)  -> str: return getattr(self._state.gparams, "doc_title",  "")
    def get_doc_author(self) -> str: return getattr(self._state.gparams, "doc_author", "")

    # ── TableStyleParams ───────────────────────────────────────────────────

    def get_table_style(self, sheet_name: str, table_name: str) -> TableStyleParams:
        """Retourne le TableStyleParams (crée les objets intermédiaires si absents)."""
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
        gp     = self._state.gparams
        gp_d   = dataclasses.asdict(gp)

        # Attributs dynamiques (non-dataclass)
        gp_d["doc_title"]  = getattr(gp, "doc_title",  "")
        gp_d["doc_author"] = getattr(gp, "doc_author", "")

        # table_intro → référence (sheet, table) au lieu de l'objet
        table_intro = getattr(gp, "table_intro", None)
        intro_ref   = None
        if table_intro is not None and self._state.workbook:
            for sheet in self._state.workbook.sheets:
                for t in sheet.tables:
                    if t is table_intro:
                        intro_ref = {"sheet": sheet.name, "table": t.name}
                        break
                if intro_ref:
                    break
        gp_d["table_intro"] = intro_ref

        # page_params : sérialisation complète (table_styles + row_styles)
        pp_serial = {}
        for sname, pp in gp.page_params.items():
            pp_d = dataclasses.asdict(pp)
            # row_styles : clés int → str (JSON n'accepte que des str comme clés)
            for tname, tsp_d in pp_d.get("table_styles", {}).items():
                tsp_d["row_styles"] = {
                    str(k): v for k, v in tsp_d.get("row_styles", {}).items()
                }
            pp_serial[sname] = pp_d
        gp_d["page_params"] = pp_serial

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
        pp_data        = gp_data.pop("page_params", {})
        doc_title      = gp_data.pop("doc_title",   "")
        doc_author     = gp_data.pop("doc_author",  "")
        table_intro_ref = gp_data.pop("table_intro", None)

        gp_fields = {f.name for f in dataclasses.fields(GlobalParams)}
        gp = GlobalParams(**{k: v for k, v in gp_data.items() if k in gp_fields})
        gp.doc_title  = doc_title
        gp.doc_author = doc_author

        # Reconstruction PageParams → TableStyleParams → RowStyle
        pp_fields  = {f.name for f in dataclasses.fields(PageParams)}
        tsp_fields = {f.name for f in dataclasses.fields(TableStyleParams)}
        rs_fields  = {f.name for f in dataclasses.fields(RowStyle)}

        for sname, pp_dict in pp_data.items():
            ts_data  = pp_dict.pop("table_styles", {})
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

        # Résolution de table_intro
        gp.table_intro = None
        if table_intro_ref and self._state.workbook:
            sn = table_intro_ref.get("sheet")
            tn = table_intro_ref.get("table")
            for sheet in self._state.workbook.sheets:
                if sheet.name == sn:
                    for t in sheet.tables:
                        if t.name == tn:
                            gp.table_intro = t
                            break
                    break

        self._state.gparams = gp

        # Charts
        cf = {f.name for f in dataclasses.fields(ChartSpec)}
        self._state.charts = [
            ChartSpec(**{k: v for k, v in c.items() if k in cf})
            for c in profile.get("charts", [])
        ]
        self.emit("profile_loaded", None)


# Singleton global
state_manager = StateManager()