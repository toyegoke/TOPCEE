# -*- coding: utf-8 -*-
"""
Equipment Cost Database Module for TOPCEE v2

Provides:
- Reading a Capital-Equipment-Cost Excel database (auto-discovers sheets/columns)
- Equipment cost estimation with capacity scaling & cost-index adjustment
- A project equipment list with running total
- Export of the equipment list to CSV/Excel

Usage:
    from equipment_cost_db import EquipmentCostDB, EquipmentItem, cost_index_ratio
"""

from __future__ import annotations

import os
import csv
import math
import re
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from dataclasses import dataclass, field, asdict
from typing import Optional, List, Dict, Any, Tuple

try:
    import openpyxl
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False

try:
    import pandas as pd
    HAS_PANDAS = True
except Exception:
    HAS_PANDAS = False


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class EquipmentItem:
    """A single piece of equipment in the project cost estimate."""
    name: str = ""
    process_id: str = ""  # Identifier used for this equipment on the process PFD
    category: str = ""
    equipment_type: str = ""
    reference_capacity: float = 0.0
    reference_cost: float = 0.0
    capacity: float = 0.0          # required / design capacity
    scaling_exponent: float = 0.6  # capacity scaling exponent
    cost_year: int = 0
    target_year: int = 0
    reference_source: str = ""
    reference_cepci: float = 0.0
    target_cepci: float = 0.0
    cost_index_ratio: float = 1.0  # CEPCI or similar index ratio (target/base)
    material_factor: float = 1.0   # material of construction factor
    pressure_factor: float = 1.0   # design pressure factor
    quantity: int = 1
    unit: str = ""
    notes: str = ""

    def scaled_cost(self) -> float:
        """Cost for ONE unit after capacity scaling, index & factors."""
        if self.reference_cost <= 0 or self.reference_capacity <= 0:
            return 0.0
        ratio = self.capacity / self.reference_capacity if self.capacity > 0 else 1.0
        if ratio <= 0:
            ratio = 1.0
        scaled = self.reference_cost * (ratio ** self.scaling_exponent)
        scaled *= self.cost_index_ratio
        scaled *= self.material_factor
        scaled *= self.pressure_factor
        return scaled

    def total_cost(self) -> float:
        return self.scaled_cost() * max(1, self.quantity)

    def to_row(self) -> List[Any]:
        return [
            self.name, self.process_id, self.category, self.equipment_type,
            self.reference_source, self.cost_year,
            f"{self.reference_capacity:.4g}", f"{self.reference_cost:.4g}",
            f"{self.capacity:.4g}", f"{self.scaling_exponent:.2f}",
            f"{self.reference_cepci:.3f}", self.target_year, f"{self.target_cepci:.3f}",
            f"{self.cost_index_ratio:.4f}",
            f"{self.material_factor:.4f}", f"{self.pressure_factor:.4f}",
            self.quantity, self.unit, f"{self.total_cost():.4g}", self.notes,
        ]

    @staticmethod
    def header() -> List[str]:
        return [
            "Name", "Process ID", "Category", "Type", "Reference Source", "Ref Year",
            "Ref Capacity", "Ref Cost ($)", "Design Capacity", "Exponent", "Ref CEPCI",
            "Target Year", "Target CEPCI", "Index Ratio", "Material Factor", "Pressure Factor",
            "Qty", "Unit", "Total Cost ($)", "Notes",
        ]


# ---------------------------------------------------------------------------
# Excel database reader
# ---------------------------------------------------------------------------

class EquipmentCostDB:
    """
    Loads a Capital-Equipment-Cost Excel workbook and exposes equipment records.
    Column mapping is discovered automatically from common header names.
    """

    DEFAULT_COLUMN_ALIASES = {
        "name": [
            "equipment", "equipment name", "item", "description",
            "name", "equipment description", "asset name", "equipment_description",
        ],
        "category": [
            "category", "type", "class", "equipment type", "service",
            "equipment class", "group", "equipment_category",
        ],
        "reference_capacity": [
            "min scale", "minimum scale", "scale min", "min capacity",
            "capacity", "size", "ref capacity", "base capacity",
            "rated capacity", "design capacity", "capacity (ref)",
            "capacity ref", "capacity_base", "ref_capacity", "max scale",
        ],
        "reference_cost": [
            "cost", "base cost", "ref cost", "purchase cost",
            "equipment cost", "price", "cost ($)", "cost usd",
            "cost $", "base price", "ref_cost", "purchase_cost",
        ],
        "cost_year": [
            "year", "cost year", "base year", "data year", "year of cost",
            "cost_year", "base_year",
        ],
        "source": ["source", "reference", "citation", "cost source", "reference source"],
        "scaling_exponent": [
            "scaling exponent", "capacity exponent", "cost exponent",
            "scale exponent", "exponent", "scaling factor", "scale factor",
            "scaling_exponent",
        ],
        "unit": [
            "scaling unit", "unit", "units", "uom", "unit of measure",
        ],
    }

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.workbook = None
        self.sheets: List[str] = []
        self.records: List[Dict[str, Any]] = []
        self.active_sheet: Optional[str] = None
        self.columns: List[str] = []
        self.mapping: Dict[str, str] = {}
        self.loaded = False

    def load(self) -> bool:
        if not HAS_OPENPYXL:
            raise RuntimeError(
                "openpyxl is required to read Excel files. "
                "Install it: pip install openpyxl"
            )
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"Excel file not found: {self.file_path}")
        self.workbook = openpyxl.load_workbook(self.file_path, data_only=True)
        self.sheets = self.workbook.sheetnames
        self.loaded = True
        return True

    def discover(self, sheet_name: Optional[str] = None,
                 mapping: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
        """Pick a sheet, read its header row, and build records."""
        if not self.loaded:
            self.load()

        if sheet_name is None:
            sheet_name = self.sheets[0]
        if sheet_name not in self.sheets:
            raise ValueError(f"Sheet '{sheet_name}' not found. Available: {self.sheets}")

        ws = self.workbook[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            raise ValueError(f"Sheet '{sheet_name}' is empty.")

        header = [str(c) if c is not None else "" for c in rows[0]]
        while header and header[-1].strip() == "":
            header.pop()
        self.columns = header
        self.active_sheet = sheet_name

        if mapping:
            self.mapping = mapping
        else:
            self.mapping = self._guess_mapping(header)

        self.records = []
        for r in rows[1:]:
            if not any(v is not None and str(v).strip() != "" for v in r):
                continue
            rec: Dict[str, Any] = {}
            for idx, col in enumerate(header):
                rec[col] = r[idx] if idx < len(r) else None
            for sem, col in self.mapping.items():
                if col and col in rec:
                    rec[f"__{sem}__"] = rec[col]
            self.records.append(rec)

        preview = [self._record_summary(r) for r in self.records[:8]]
        return {
            "sheet": sheet_name,
            "columns": self.columns,
            "mapping": self.mapping,
            "n_records": len(self.records),
            "preview": preview,
        }

    def _guess_mapping(self, header: List[str]) -> Dict[str, str]:
        mapping = {}
        lowered = [h.strip().lower() for h in header]
        for sem, aliases in self.DEFAULT_COLUMN_ALIASES.items():
            found = None
            if sem == "scaling_exponent":
                for i, h in enumerate(lowered):
                    if h in {"n", "m"}:
                        found = header[i]
                        break
            for alias in aliases:
                if found:
                    break
                for i, h in enumerate(lowered):
                    if alias == h or alias in h:
                        found = header[i]
                        break
                if found:
                    break
            if found:
                mapping[sem] = found
        return mapping

    def _record_summary(self, rec: Dict[str, Any]) -> Dict[str, Any]:
        out = {}
        for sem in ["name", "category", "reference_capacity", "reference_cost", "cost_year"]:
            col = self.mapping.get(sem)
            out[sem] = rec.get(col) if col else None
        return out

    def list_equipment(self, category_filter: Optional[str] = None,
                       search: Optional[str] = None) -> List[Dict[str, Any]]:
        results = []
        for rec in self.records:
            name = str(rec.get(self.mapping.get("name", ""), ""))
            category = str(rec.get(self.mapping.get("category", ""), ""))
            if category_filter and category_filter.lower() not in category.lower():
                continue
            if search and search.lower() not in name.lower() and search.lower() not in category.lower():
                continue
            results.append(rec)
        return results

    def get_categories(self) -> List[str]:
        cats = set()
        col = self.mapping.get("category", "")
        for rec in self.records:
            v = rec.get(col)
            if v:
                cats.add(str(v).strip())
        return sorted(cats)

    def record_to_item(self, rec: Dict[str, Any],
                       design_capacity: Optional[float] = None,
                       quantity: int = 1,
                       target_year: Optional[int] = None,
                       cost_index_ratio: float = 1.0,
                       material_factor: float = 1.0,
                       pressure_factor: float = 1.0) -> EquipmentItem:
        col = lambda sem: self.mapping.get(sem)

        def _num(key, default=0.0):
            c = col(key)
            v = rec.get(c) if c else None
            try:
                return float(v) if v is not None and str(v).strip() != "" else default
            except (ValueError, TypeError):
                return default

        def _int(key, default=0):
            c = col(key)
            v = rec.get(c) if c else None
            try:
                return int(v) if v is not None else default
            except (ValueError, TypeError):
                return default

        def _str(key, default=""):
            c = col(key)
            v = rec.get(c) if c else None
            return str(v).strip() if v is not None else default

        ref_cap = _num("reference_capacity")
        ref_cost = _num("reference_cost")
        exp = _num("scaling_exponent", 0.6)
        source = _str("source")
        year = parse_reference_year(source) or 0
        ref_cepci = cepci_for_year(year)[0] if year else 0.0

        cap = design_capacity if design_capacity is not None else ref_cap
        tgt_year = target_year if target_year is not None else year
        target_cepci = cepci_for_year(tgt_year)[0] if tgt_year else 0.0
        ratio = (target_cepci / ref_cepci) if ref_cepci > 0 and target_cepci > 0 else cost_index_ratio

        return EquipmentItem(
            name=_str("name"),
            category=_str("category"),
            equipment_type=_str("category"),
            reference_source=source,
            reference_capacity=ref_cap,
            reference_cost=ref_cost,
            capacity=cap,
            scaling_exponent=exp,
            cost_year=year,
            target_year=tgt_year,
            reference_cepci=ref_cepci,
            target_cepci=target_cepci,
            cost_index_ratio=ratio,
            material_factor=material_factor,
            pressure_factor=pressure_factor,
            quantity=quantity,
            unit=_str("unit"),
            notes="Ref CEPCI derived from Ref Year using TOPCEE's reported CEPCI data/regression; equipment-database CEPCI column not used.",
        )


# ---------------------------------------------------------------------------
# Cost index helpers (CEPCI)
# ---------------------------------------------------------------------------

# Reported annual values supplied for TOPCEE's regression. 2014 uses 573.6 as
# specified by the project author (some sources report 576.1).
DEFAULT_CEPCI = {
    2007: 525.4, 2008: 575.4, 2009: 521.9, 2010: 550.8,
    2011: 585.7, 2012: 584.6, 2013: 567.3, 2014: 573.6,
    2015: 556.8, 2016: 541.7, 2017: 567.5, 2018: 603.1,
    2019: 607.5, 2020: 596.2, 2021: 708.8, 2022: 816.0,
    2023: 797.9,
}


def parse_reference_year(source: str) -> Optional[int]:
    """Extract a plausible publication year from an equipment source citation."""
    years = [int(v) for v in re.findall(r"\b(?:19|20)\d{2}\b", str(source or ""))]
    return years[-1] if years else None


def cepci_regression() -> Tuple[float, float, float]:
    """Return (slope, intercept, R²) for OLS CEPCI = slope*year + intercept."""
    xs = list(DEFAULT_CEPCI)
    ys = [DEFAULT_CEPCI[x] for x in xs]
    xbar, ybar = sum(xs) / len(xs), sum(ys) / len(ys)
    denominator = sum((x - xbar) ** 2 for x in xs)
    slope = sum((x - xbar) * (y - ybar) for x, y in zip(xs, ys)) / denominator
    intercept = ybar - slope * xbar
    total = sum((y - ybar) ** 2 for y in ys)
    residual = sum((y - (slope * x + intercept)) ** 2 for x, y in zip(xs, ys))
    return slope, intercept, 1.0 - residual / total if total else 1.0


def cepci_for_year(year: int) -> Tuple[float, bool, float]:
    """Return (index, is_regression_estimate, R²) for a selected calendar year."""
    year = int(year)
    if year in DEFAULT_CEPCI:
        return DEFAULT_CEPCI[year], False, cepci_regression()[2]
    slope, intercept, r_squared = cepci_regression()
    estimate = slope * year + intercept
    if estimate <= 0:
        raise ValueError(f"Linear CEPCI regression returned a non-positive estimate for {year}.")
    return estimate, True, r_squared


def cost_index_ratio(base_year: int, target_year: int,
                     index_table: Optional[Dict[int, float]] = None) -> float:
    """Get target/base CEPCI ratio, using OLS regression outside reported years."""
    table = index_table or DEFAULT_CEPCI
    base = table.get(int(base_year))
    if base is None:
        base = cepci_for_year(int(base_year))[0]
    target = table.get(int(target_year))
    if target is None:
        target = cepci_for_year(int(target_year))[0]
    return target / base


# ---------------------------------------------------------------------------
# Standalone equipment costing editor (can be used independently)
# ---------------------------------------------------------------------------

class EquipmentCostingWindow:
    """Toplevel window for selecting equipment from the DB and building a cost list."""

    def __init__(self, parent: tk.Tk, db: EquipmentCostDB,
                 on_complete=None):
        self.parent = parent
        self.db = db
        self.on_complete = on_complete
        self.items: List[EquipmentItem] = []

        self.top = tk.Toplevel(parent)
        self.top.title("Equipment Costing - TOPCEE v3")
        self.top.geometry("1200x720")
        self.top.transient(parent)
        self.top.grab_set()

        self._build_ui()
        self._refresh_sheet_list()

    def _build_ui(self):
        nb = ttk.Notebook(self.top)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # --- Tab 1: Database browser & estimator ---
        db_tab = ttk.Frame(nb)
        nb.add(db_tab, text="1. Equipment Database")

        ctrl = tk.Frame(db_tab)
        ctrl.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        tk.Label(ctrl, text="Excel file:").pack(side=tk.LEFT)
        self.file_var = tk.StringVar(value=self.db.file_path)
        tk.Entry(ctrl, textvariable=self.file_var, width=52).pack(side=tk.LEFT, padx=4)
        tk.Button(ctrl, text="Browse", command=self._browse_file).pack(side=tk.LEFT, padx=2)
        tk.Button(ctrl, text="Reload", command=self._reload).pack(side=tk.LEFT, padx=2)

        tk.Label(ctrl, text="Sheet:").pack(side=tk.LEFT, padx=(10, 0))
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(ctrl, textvariable=self.sheet_var,
                                        state="readonly", width=26)
        self.sheet_combo.pack(side=tk.LEFT, padx=2)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda e: self._load_sheet())

        # Column mapping
        map_frame = tk.LabelFrame(db_tab, text="Column Mapping (auto-detected - adjust if needed)")
        map_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        self.map_entries: Dict[str, tk.Entry] = {}
        aliases = ["name", "category", "reference_capacity", "reference_cost",
                   "cost_year", "scaling_exponent", "unit"]
        for i, sem in enumerate(aliases):
            tk.Label(map_frame, text=sem.replace("_", " ").title() + ":").grid(
                row=0, column=i * 2, sticky="e", padx=3, pady=2)
            ent = tk.Entry(map_frame, width=18)
            ent.grid(row=0, column=i * 2 + 1, sticky="w", padx=3, pady=2)
            self.map_entries[sem] = ent

        tk.Button(map_frame, text="Apply Mapping", command=self._apply_mapping).grid(
            row=0, column=len(aliases) * 2 + 1, padx=6, pady=2)

        # Equipment list (database side)
        list_frame = tk.LabelFrame(db_tab, text="Available Equipment")
        list_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)

        cols = ("name", "category", "reference_capacity", "reference_cost",
                "cost_year", "unit")
        self.db_tree = ttk.Treeview(list_frame, columns=cols, show="headings", height=9)
        for c in cols:
            self.db_tree.heading(c, text=c.replace("_", " ").title())
            self.db_tree.column(c, width=150)
        self.db_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.db_tree.yview)
        self.db_tree.configure(yscroll=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Search / filter
        sf = tk.Frame(db_tab)
        sf.pack(side=tk.TOP, fill=tk.X, padx=5, pady=3)
        tk.Label(sf, text="Search:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self._populate_db_tree())
        tk.Entry(sf, textvariable=self.search_var, width=30).pack(side=tk.LEFT, padx=4)
        tk.Label(sf, text="Category filter:").pack(side=tk.LEFT, padx=(10, 0))
        self.cat_var = tk.StringVar()
        self.cat_combo = ttk.Combobox(sf, textvariable=self.cat_var,
                                      state="readonly", width=22)
        self.cat_combo.pack(side=tk.LEFT, padx=2)
        self.cat_combo.bind("<<ComboboxSelected>>", lambda e: self._populate_db_tree())

        # Estimation controls
        est_frame = tk.LabelFrame(db_tab, text="Specification & Cost Estimation for Selected Equipment")
        est_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        labels = [
            ("Design capacity:", "cap_var", ""),
            ("Scaling exponent:", "exp_var", "0.6"),
            ("Target year:", "tyear_var", str(datetime.datetime.now().year)),
            ("Cost-index ratio:", "cidx_var", "1.0"),
            ("Material factor:", "mat_var", "1.0"),
            ("Pressure factor:", "pres_var", "1.0"),
            ("Quantity:", "qty_var", "1"),
        ]
        self.est_vars: Dict[str, tk.StringVar] = {}
        for i, (label, varname, default) in enumerate(labels):
            tk.Label(est_frame, text=label).grid(row=i // 4, column=(i % 4) * 2,
                                                 sticky="e", padx=4, pady=2)
            var = tk.StringVar(value=default)
            self.est_vars[varname] = var
            tk.Entry(est_frame, textvariable=var, width=12).grid(
                row=i // 4, column=(i % 4) * 2 + 1, sticky="w", padx=4, pady=2)

        btn_frame = tk.Frame(est_frame)
        btn_frame.grid(row=2, column=0, columnspan=8, pady=6)
        tk.Button(btn_frame, text="Add to Project List", bg="#4CAF50", fg="white",
                  command=self._add_selected_to_project).pack(side=tk.LEFT, padx=4)
        tk.Button(btn_frame, text="Quick Add (use DB capacity)",
                  command=self._add_selected_quick).pack(side=tk.LEFT, padx=4)

        # --- Tab 2: Project list ---
        proj_tab = ttk.Frame(nb)
        nb.add(proj_tab, text="2. Project Equipment List")

        plist_frame = tk.LabelFrame(proj_tab, text="Equipment in Project")
        plist_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)

        pcols = EquipmentItem.header()
        self.proj_tree = ttk.Treeview(plist_frame, columns=pcols, show="headings", height=14)
        for c in pcols:
            self.proj_tree.heading(c, text=c)
            self.proj_tree.column(c, width=105 if c != "Notes" else 180)
        self.proj_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        proj_vsb = ttk.Scrollbar(plist_frame, orient="vertical", command=self.proj_tree.yview)
        self.proj_tree.configure(yscroll=proj_vsb.set)
        proj_vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Totals
        self.total_var = tk.StringVar(value="Total Equipment Cost: $ 0.00")
        tk.Label(proj_tab, textvariable=self.total_var, font=("TkDefaultFont", 12, "bold")).pack(
            side=tk.TOP, anchor="e", padx=5, pady=4)

        # Actions
        pbtn = tk.Frame(proj_tab)
        pbtn.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        tk.Button(pbtn, text="Remove Selected", command=self._remove_selected).pack(side=tk.LEFT, padx=4)
        tk.Button(pbtn, text="Clear All", command=self._clear_project).pack(side=tk.LEFT, padx=4)
        tk.Button(pbtn, text="Export CSV", command=self._export_project_csv).pack(side=tk.LEFT, padx=4)
        tk.Button(pbtn, text="Use Total in TOPCEE", bg="#2196F3", fg="white",
                  command=self._use_in_topcee).pack(side=tk.RIGHT, padx=4)
        tk.Button(pbtn, text="Close", command=self.top.destroy).pack(side=tk.RIGHT, padx=4)

    # ---- Database tab actions ----------------------------------------------

    def _refresh_sheet_list(self):
        try:
            if not self.db.loaded:
                self.db.load()
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open Excel file:\n{e}")
            return
        self.sheet_combo["values"] = self.db.sheets
        if self.db.sheets:
            self.sheet_var.set(self.db.sheets[0])
            self._load_sheet()

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Equipment Cost Database",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.file_var.set(path)
            self.db.file_path = path
            try:
                self.db.load()
                self._refresh_sheet_list()
                messagebox.showinfo("Loaded", f"Loaded {len(self.db.sheets)} sheet(s).")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load workbook:\n{e}")

    def _reload(self):
        try:
            self.db.load()
            self._refresh_sheet_list()
            messagebox.showinfo("Reloaded", f"Sheets: {', '.join(self.db.sheets)}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _load_sheet(self):
        sheet = self.sheet_var.get()
        if not sheet:
            return
        try:
            info = self.db.discover(sheet)
            for sem, ent in self.map_entries.items():
                ent.delete(0, tk.END)
                ent.insert(0, self.db.mapping.get(sem, ""))
            cats = self.db.get_categories()
            self.cat_combo["values"] = [""] + cats
            self.cat_var.set("")
            self._populate_db_tree()
            messagebox.showinfo(
                "Sheet Loaded",
                f"Sheet: {info['sheet']}\nRecords: {info['n_records']}\n"
                f"Columns: {', '.join(info['columns'])}",
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheet:\n{e}")

    def _apply_mapping(self):
        mapping = {}
        for sem, ent in self.map_entries.items():
            v = ent.get().strip()
            if v:
                mapping[sem] = v
        self.db.mapping = mapping
        self._populate_db_tree()
        messagebox.showinfo("Mapping", "Column mapping applied.")

    def _populate_db_tree(self):
        self.db_tree.delete(*self.db_tree.get_children())
        search = self.search_var.get().lower()
        cat = self.cat_var.get()
        for rec in self.db.list_equipment(category_filter=cat or None, search=search or None):
            col = lambda sem: self.db.mapping.get(sem, "")
            vals = [
                rec.get(col("name"), ""),
                rec.get(col("category"), ""),
                rec.get(col("reference_capacity"), ""),
                rec.get(col("reference_cost"), ""),
                rec.get(col("cost_year"), ""),
                rec.get(col("unit"), ""),
            ]
            self.db_tree.insert("", tk.END, values=vals)

    def _selected_db_record(self) -> Optional[Dict[str, Any]]:
        sel = self.db_tree.selection()
        if not sel:
            return None
        idx = self.db_tree.index(sel[0])
        search = self.search_var.get().lower()
        cat = self.cat_var.get()
        filtered = self.db.list_equipment(
            category_filter=cat or None, search=search or None)
        if 0 <= idx < len(filtered):
            return filtered[idx]
        return None

    # ---- Estimation / add to project ---------------------------------------

    def _add_selected_to_project(self):
        rec = self._selected_db_record()
        if rec is None:
            messagebox.showwarning("Select", "Please select an equipment row first.")
            return
        try:
            cap = float(self.est_vars["cap_var"].get()) if self.est_vars["cap_var"].get().strip() else None
            exp = float(self.est_vars["exp_var"].get())
            tyear = int(self.est_vars["tyear_var"].get()) if self.est_vars["tyear_var"].get().strip() else None
            cidx = float(self.est_vars["cidx_var"].get())
            mat = float(self.est_vars["mat_var"].get())
            pres = float(self.est_vars["pres_var"].get())
            qty = int(self.est_vars["qty_var"].get())
        except ValueError:
            messagebox.showerror("Input", "Please enter valid numbers for estimation fields.")
            return

        if tyear is not None:
            rec_year = rec.get(self.db.mapping.get("cost_year", ""), None)
            try:
                rec_year = int(rec_year)
                cidx = cost_index_ratio(rec_year, tyear)
            except (ValueError, TypeError):
                pass

        item = self.db.record_to_item(
            rec, design_capacity=cap, quantity=qty,
            target_year=tyear, cost_index_ratio=cidx,
            material_factor=mat, pressure_factor=pres,
        )
        self.items.append(item)
        self._refresh_project_tree()
        self.est_vars["exp_var"].set("0.6")
        self.est_vars["cidx_var"].set("1.0")
        self.est_vars["mat_var"].set("1.0")
        self.est_vars["pres_var"].set("1.0")

    def _add_selected_quick(self):
        rec = self._selected_db_record()
        if rec is None:
            messagebox.showwarning("Select", "Please select an equipment row first.")
            return
        try:
            qty = int(self.est_vars["qty_var"].get())
        except ValueError:
            qty = 1
        item = self.db.record_to_item(rec, quantity=qty)
        self.items.append(item)
        self._refresh_project_tree()

    # ---- Project list tab actions ------------------------------------------

    def _refresh_project_tree(self):
        self.proj_tree.delete(*self.proj_tree.get_children())
        total = 0.0
        for item in self.items:
            self.proj_tree.insert("", tk.END, values=item.to_row())
            total += item.total_cost()
        self.total_var.set(f"Total Equipment Cost: $ {total:,.2f}")

    def _remove_selected(self):
        sel = self.proj_tree.selection()
        if not sel:
            return
        idx = self.proj_tree.index(sel[0])
        if 0 <= idx < len(self.items):
            self.items.pop(idx)
            self._refresh_project_tree()

    def _clear_project(self):
        if messagebox.askyesno("Confirm", "Clear the entire project equipment list?"):
            self.items.clear()
            self._refresh_project_tree()

    def _export_project_csv(self):
        if not self.items:
            messagebox.showinfo("Info", "Project list is empty.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV files", "*.csv")],
            title="Export Equipment Cost List",
        )
        if not path:
            return
        try:
            with open(path, "w", newline="") as f:
                w = csv.writer(f)
                w.writerow(EquipmentItem.header())
                for it in self.items:
                    w.writerow(it.to_row())
                w.writerow([])
                w.writerow(["TOTAL"] + [""] * 17 +
                            [f"{sum(i.total_cost() for i in self.items):.2f}", ""])
            messagebox.showinfo("Saved", f"Equipment list exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _use_in_topcee(self):
        if not self.items:
            messagebox.showinfo("Info", "Project list is empty.")
            return
        total = sum(i.total_cost() for i in self.items)
        if self.on_complete:
            self.on_complete(total, self.items)
        messagebox.showinfo(
            "Equipment Cost Ready",
            f"Total equipment cost = $ {total:,.2f}\n"
            f"This value will be sent back to TOPCEE.",
        )
        self.top.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    default_path = r"C:\Users\Dr_Oyegoke\Downloads\Capital-Equipment-Cost-Database.xlsx"
    db = EquipmentCostDB(default_path)
    try:
        db.load()
    except Exception as e:
        print("Could not auto-load DB:", e)

    def on_complete(total, items):
        print("Total equipment cost:", total)
        for it in items:
            print(" ", it.name, it.total_cost())

    EquipmentCostingWindow(root, db, on_complete=on_complete)
    root.mainloop()
