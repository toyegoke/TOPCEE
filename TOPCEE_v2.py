# -*- coding: utf-8 -*-
"""
TOPCEE v2.0 - Toyese Oyegoke Process Cost & Economic Evaluator
==============================================================

Complete process economic analysis application with:
  1. Equipment selection from Capital-Equipment-Cost database with auto-costing
  2. Process economic input data (CAPEX/OPEX parameters)
  3. Profitability analysis (NPV, IRR, ROI, payback, B/C ratio, cash flows)
  4. Integrated report generation: Equipment cost report + CAPEX breakdown +
     OPEX breakdown + Profitability analysis
  5. Save/Load inputs, remembers last output folder

Requires:
    pip install openpyxl matplotlib numpy

Author: Engr. Dr. T. Oyegoke
"""

from __future__ import annotations

import os
import sys
import json
import csv
import html
import base64
import io
import datetime
import math
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional, List, Dict, Any, Tuple

import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

# Local modules
from equipment_cost_db import (
    EquipmentCostDB, EquipmentItem, EquipmentCostingWindow, cost_index_ratio,
    DEFAULT_CEPCI, cepci_for_year, parse_reference_year
)
from cost_report import CostReportBuilder

# ---------------------------------------------------------------------------
# Configuration (remembers last output folder)
# ---------------------------------------------------------------------------

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "topcee_config.json")
DEFAULT_DB_PATH = r"C:\Users\Dr_Oyegoke\Downloads\Capital-Equipment-Cost-Database.xlsx"


def load_config() -> Dict[str, Any]:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_config(cfg: Dict[str, Any]) -> None:
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Economic calculation engine (adapted from TOPCEE v2.9)
# ---------------------------------------------------------------------------

FLUID = [1.00, .47, .36, .68, .11, .18, .10, .70, 3.60, .33, .41, 4.34, .26, .44, 5.04, .89, 5.93]
MIXED = [1.00, .39, .26, .31, .10, .29, .12, .55, 3.02, .32, .34, 3.68, .23, .37, 4.28, .75, 5.03]
SOLID = [1.00, .45, .18, .16, .10, .25, .15, .40, 2.69, .33, .39, 3.41, .21, .35, 3.97, .70, 4.67]
DEPRECIABLE_LIFE_YEARS = 10
OPERATOR_WORK_HOURS_PER_YEAR = 49 * 5 * 8  # 49 work weeks × 5 shifts/week × 8 h/shift


def calculate_profitability(inputs: Dict[str, str], show_errors: bool = True) -> Optional[Dict[str, Any]]:
    """
    Run the full CAPEX / OPEX / profitability calculation.

    inputs: dict of field_name -> string value (as entered by user)
    Returns a dict with all results, or None on error.
    """
    def show_calc_error(title, message):
        if show_errors:
            messagebox.showerror(title, message)

    try:
        # ---- Parse inputs ----
        def safe_float(val, default=0.0):
            """Safely convert any value to float, returning default on failure."""
            try:
                if val is None:
                    return default
                # Handle tkinter _NoValueType and Variable objects - check type FIRST
                val_type_name = type(val).__name__
                if 'Variable' in val_type_name or 'NoValue' in val_type_name or 'Tcl' in val_type_name:
                    return default
                val_str = str(val).strip()
                if val_str == "":
                    return default
                number = float(val_str.replace(",", ""))
                if not math.isfinite(number):
                    raise ValueError(f"Numeric input must be finite: {val_str!r}.")
                return number
            except (TypeError, AttributeError):
                return default
            except ValueError as exc:
                if str(exc).startswith("Numeric input must be finite:"):
                    raise
                raise ValueError(f"Enter a valid number instead of {str(val).strip()!r}.") from exc

        X = safe_float(inputs.get('Exchange rate (local currency per $)',
                                  inputs.get('Exchange rate (N/$)', '0')))
        if X <= 0:
            raise ValueError("Exchange rate must be greater than zero.")
        currency_code = str(inputs.get('Local currency code', 'NGN')).strip().upper() or 'LOCAL'
        depreciation_life = safe_float(inputs.get('Depreciation period in years',
                                                   str(DEPRECIABLE_LIFE_YEARS)))
        startup_value = safe_float(inputs.get(
            'Construction/Startup period (production starts in year)', '3'), 3)
        if not float(startup_value).is_integer():
            raise ValueError("Construction/Startup period must be a whole year number.")
        startup_year = int(startup_value)
        if not float(depreciation_life).is_integer():
            raise ValueError("Depreciation period must be a whole number of years.")
        if depreciation_life <= 0:
            raise ValueError("Depreciation period must be greater than zero years.")
        if startup_year < 1:
            raise ValueError("Construction/Startup period must be year 1 or later.")
        if 'Operating labor cost (local/hour)' in inputs:
            operator_hourly_wage = safe_float(inputs.get('Operating labor cost (local/hour)', '0'))
        else:
            # Convert legacy monthly wage entries to the new hourly basis using
            # the assumed 1,960 paid operating hours per employee-year.
            legacy_monthly_wage = safe_float(inputs.get('Wage (local/month)',
                                                        inputs.get('Wage (N/month)', '0')))
            operator_hourly_wage = legacy_monthly_wage * 12 / OPERATOR_WORK_HOURS_PER_YEAR
        D = safe_float(inputs.get('Subsidy (%)', '0'))
        C = safe_float(inputs.get('Tax (%)', '0'))
        h = safe_float(inputs.get('Working hour in h/day', '24'))
        d = safe_float(inputs.get('Operating days per year',
                                  inputs.get('Working days in d/year', '330')))
        shifts_per_day_value = safe_float(inputs.get('Number of shifts per day', '3'), 3)
        if not float(shifts_per_day_value).is_integer() or shifts_per_day_value < 1:
            raise ValueError("Number of shifts per day must be a positive whole number.")
        shifts_per_day = int(shifts_per_day_value)
        if h <= 0 or h > 24:
            raise ValueError("Operating hours per day must be greater than zero and no more than 24.")
        if d < 0 or d > 366:
            raise ValueError("Operating days per year must be between 0 and 366.")
        if operator_hourly_wage < 0:
            raise ValueError("Operating labor cost per hour cannot be negative.")
        if not 0 <= D <= 100:
            raise ValueError("Subsidy must be between 0 and 100%.")
        if not 0 <= C <= 100:
            raise ValueError("Tax must be between 0 and 100%.")
        if 'Selling price of your product (local/L)' in inputs:
            SPv = safe_float(inputs.get('Selling price of your product (local/L)', '0')) / X
        else:
            # Older saved projects stored the product selling price directly in USD/L.
            SPv = safe_float(inputs.get('Selling price of your product in $/L', '0'))
        PCE = safe_float(inputs.get('Total equipment cost ($)', '0'))
        if PCE < 0:
            raise ValueError("Total equipment cost cannot be negative.")
        ProcessType = str(inputs.get('Process Type: F-Fluid, S-Solid, M-Mixed', 'F')).strip().upper()

        raw_materials = []
        for i in range(1, 11):
            qty_key = f'Raw material {i} quantity (unit/h)'
            cost_key = f'Raw material {i} cost (local/unit)'
            if qty_key not in inputs:
                qty_key = ('Raw material quantity (kg/h)' if i == 1
                           else f'Raw material {i} quantity (kg/h)')
            if cost_key not in inputs:
                cost_key = ('Raw material cost (N/kg)' if i == 1
                            else f'Raw material {i} cost (local/kg)')
            name = str(inputs.get(f'Raw material {i} name', f'Raw material {i}')).strip() or f'Raw material {i}'
            qty = safe_float(inputs.get(qty_key, '0'))
            cost = safe_float(inputs.get(cost_key, '0'))
            if qty < 0 or cost < 0:
                raise ValueError(f"Flow rate and price cannot be negative for {name}.")
            if qty or cost:
                rate_unit = str(inputs.get(f'Raw material {i} rate unit', 'kg/h')).strip().lower()
                price_unit = str(inputs.get(f'Raw material {i} price basis', 'kg')).strip().lower()
                rate_unit = rate_unit.replace('³', '3').replace(' ', '')
                price_unit = price_unit.replace('³', '3').replace(' ', '').replace('local/', '')
                density = safe_float(inputs.get(f'Raw material {i} density (kg/L)', '0'))
                mass_rate_units = {'kg/h', 't/h'}
                volume_rate_units = {'l/h', 'm3/h'}
                mass_price_units = {'kg', 't'}
                volume_price_units = {'l', 'm3'}
                if rate_unit not in mass_rate_units | volume_rate_units:
                    raise ValueError(f"Unsupported flow unit '{rate_unit}' for {name}.")
                if price_unit not in mass_price_units | volume_price_units:
                    raise ValueError(f"Unsupported price basis '{price_unit}' for {name}.")

                if rate_unit in mass_rate_units:
                    mass_kg_h = qty * (1000.0 if rate_unit == 't/h' else 1.0)
                    if price_unit in mass_price_units:
                        local_cost_per_h = mass_kg_h * cost / (1000.0 if price_unit == 't' else 1.0)
                    else:
                        if qty and cost and density <= 0:
                            raise ValueError(
                                f"Enter density in kg/L for {name} to convert a mass flow to a volume price basis."
                            )
                        volume_l_h = mass_kg_h / density if density > 0 else 0.0
                        local_cost_per_h = volume_l_h * cost / (1000.0 if price_unit == 'm3' else 1.0)
                else:
                    volume_l_h = qty * (1000.0 if rate_unit == 'm3/h' else 1.0)
                    if price_unit in volume_price_units:
                        local_cost_per_h = volume_l_h * cost / (1000.0 if price_unit == 'm3' else 1.0)
                    else:
                        if qty and cost and density <= 0:
                            raise ValueError(
                                f"Enter density in kg/L for {name} to convert a volume flow to a mass price basis."
                            )
                        mass_kg_h = volume_l_h * density if density > 0 else 0.0
                        local_cost_per_h = mass_kg_h * cost / (1000.0 if price_unit == 't' else 1.0)
                raw_materials.append((name, local_cost_per_h))

        process_consumables = []
        has_catalyst_consumable = False
        annual_operating_hours = h * d
        for i in range(1, 11):
            name = str(inputs.get(f'Process consumable {i} name', '')).strip() or f'Consumable {i}'
            mode = str(inputs.get(f'Process consumable {i} mode', 'Continuous')).strip().lower()
            quantity = safe_float(inputs.get(f'Process consumable {i} quantity', '0'))
            unit = str(inputs.get(f'Process consumable {i} unit', '')).strip()
            unit_cost = safe_float(inputs.get(f'Process consumable {i} cost (local/unit)', '0'))
            interval = safe_float(inputs.get(
                f'Process consumable {i} replacement interval (operating h)', '0'))
            if quantity < 0 or unit_cost < 0 or interval < 0:
                raise ValueError(f"Quantity, price, and replacement interval cannot be negative for {name}.")
            if quantity or unit_cost:
                if 'periodic' in mode:
                    if quantity and unit_cost and interval <= 0:
                        raise ValueError(
                            f"Enter a replacement interval in operating hours for {name}."
                        )
                    annual_cost = (quantity * unit_cost / X * annual_operating_hours / interval
                                   if interval > 0 else 0.0)
                else:
                    annual_cost = quantity * unit_cost / X * annual_operating_hours
                process_consumables.append((name, unit, annual_cost))
                if 'catalyst' in name.lower():
                    has_catalyst_consumable = True

        # Accept catalyst fields from projects saved before the consumables
        # table was introduced, treating their rate as continuous make-up use.
        legacy_catalyst_qty = safe_float(inputs.get(
            'Catalyst quantity (kg/h)', inputs.get('Catalyst consumption (kg/h)', '0')))
        legacy_catalyst_cost = safe_float(inputs.get(
            'Catalyst cost (local/kg)', inputs.get('Catalyst cost (N/kg)', '0')))
        if (legacy_catalyst_qty or legacy_catalyst_cost) and not has_catalyst_consumable:
            legacy_catalyst_annual = (legacy_catalyst_qty * legacy_catalyst_cost / X
                                      * annual_operating_hours)
            process_consumables.append(('Catalyst', 'kg', legacy_catalyst_annual))
            has_catalyst_consumable = True

        utility_items = []
        utility_catalyst = None
        for i in range(1, 11):
            name = str(inputs.get(f'Utility {i} name', '')).strip() or f'Utility {i}'
            qty = safe_float(inputs.get(f'Utility {i} quantity (unit/h)', '0'))
            cost = safe_float(inputs.get(f'Utility {i} cost (local/unit)', '0'))
            if qty < 0 or cost < 0:
                raise ValueError(f"Utility quantity and price cannot be negative for {name}.")
            # Older app versions placed catalyst in a utility row. Migrate it
            # to the process consumables list and never count it twice.
            if 'catalyst' in name.lower():
                if (qty or cost) and not has_catalyst_consumable and utility_catalyst is None:
                    utility_catalyst = (qty, cost)
                continue
            if qty or cost:
                utility_items.append((name, qty, cost))
        # Support input files saved by earlier versions of TOPCEE.
        if not utility_items:
            legacy_utilities = [
                ('Cooling water', 'Cooling water quantity (kg/h)', 'Cooling water cost (N/kg)'),
                ('Process water', 'Process water quantity (kg/h)', 'Process water cost (N/kg)'),
                ('Power', 'Power demand quantity (kWh)', 'Power supply cost (N/kWh)'),
                ('Other utilities', 'Other utilities quantity (kg/h)', 'Other utilities cost (N/kg)'),
            ]
            utility_items = []
            for name, qty_key, cost_key in legacy_utilities:
                qty = safe_float(inputs.get(qty_key, '0'))
                cost = safe_float(inputs.get(cost_key, '0'))
                if qty or cost:
                    utility_items.append((name, qty, cost))
        if utility_catalyst and not has_catalyst_consumable:
            qty, cost = utility_catalyst
            process_consumables.append((
                'Catalyst', 'kg', qty * cost / X * annual_operating_hours
            ))
            has_catalyst_consumable = True
        prod_qty = safe_float(inputs.get('Quantity of product manufactured in L/h', '0'))
        selling_price_local = safe_float(inputs.get('Selling price of your product (local/L)', '0'))
        if prod_qty < 0 or selling_price_local < 0:
            raise ValueError("Product production rate and selling price cannot be negative.")
        nPS_value = safe_float(inputs.get('Particulate processing steps (P)', '0'))
        nNPS_value = safe_float(inputs.get(
            'Non-particulate processing steps (Nnp)',
            inputs.get('Number of processing steps/units', '0')))
        if not nPS_value.is_integer() or not nNPS_value.is_integer():
            raise ValueError("Particulate and non-particulate process-step counts must be whole numbers.")
        nPS = int(nPS_value)
        nNPS = int(nNPS_value)
        if nPS < 0 or nNPS < 0:
            raise ValueError("Particulate and non-particulate process-step counts cannot be negative.")
        nOL = (6.29 + 31.7 * nPS ** 2 + 0.23 * nNPS) ** 0.5
        operators_per_shift = int(math.ceil(nOL))
        shift_hours = h / shifts_per_day
        annual_shifts_per_operator = OPERATOR_WORK_HOURS_PER_YEAR / shift_hours
        total_plant_shifts = shifts_per_day * d
        total_operators = (int(math.ceil(operators_per_shift * total_plant_shifts /
                                         annual_shifts_per_operator))
                           if total_plant_shifts > 0 else 0)
        annual_labor_cost = (total_operators * OPERATOR_WORK_HOURS_PER_YEAR *
                             operator_hourly_wage / X)
        project_life_value = safe_float(inputs.get('Project life in years', '0'))
        if not project_life_value.is_integer():
            raise ValueError("Project life must be a whole number of years.")
        p = int(project_life_value)
        r = safe_float(inputs.get('Discount rate in %', '0')) / 100.0
        if r <= -1.0:
            raise ValueError("Discount rate must be greater than -100%.")

        # Ensure project life is at least 1 year for numpy array operations
        if p < 1:
            p = 1

        if ProcessType not in ('F', 'M', 'S'):
            raise ValueError("Process Type must be F, S, or M")

        # ---- Select Lang factor list ----
        if ProcessType == 'F':
            lang = FLUID
        elif ProcessType == 'M':
            lang = MIXED
        else:
            lang = SOLID

        # ---- CAPEX estimation ----
        f = [0.0] * 17
        f[0] = PCE * lang[0]  # Purchase cost major equipment
        f[1] = PCE * lang[1]  # Installations
        f[2] = PCE * lang[2]  # Instrumentation and control
        f[3] = PCE * lang[3]  # Piping
        f[4] = PCE * lang[4]  # Electrical
        f[5] = PCE * lang[5]  # Buildings and structures
        f[6] = PCE * lang[6]  # Yard Improvement
        f[7] = PCE * lang[7]  # Service facilities
        f[8] = PCE * lang[8]; DPC = f[8]  # Total direct plant cost
        f[9] = PCE * lang[9]  # Engineering and supervision
        f[10] = PCE * lang[10]  # Construction expenses
        f[11] = PCE * lang[11]; TPC = f[11]  # Total Direct & Indirect Plant Cost
        f[12] = PCE * lang[12]  # Contractor's fee and legal expenses
        f[13] = PCE * lang[13]  # Contingency
        f[14] = PCE * lang[14]; FCI = f[14]; TDC = FCI  # Total fixed-capital investment
        f[15] = PCE * lang[15]; WC = f[15]  # Working Capital & Land
        f[16] = PCE * lang[16]; TCI = f[16]  # FCI + WC
        TPI = 1.10 * FCI  # Total permanent investment allowance used by this estimate
        f_OSBL = sum(lang[5:7])
        DPI = (1 + f_OSBL) * DPC
        # Lang-factor capital basis: FCI is the fixed/depreciable capital;
        # working capital is separate and included in TCI.
        TDC = FCI

        # ---- Maintenance supply ----
        if ProcessType == 'F':
            MS = 0.035 * TDC
        elif ProcessType == 'M':
            MS = 0.045 * TDC
        else:
            MS = 0.05 * TDC

        # ---- Cost of Manufacturing (iterative) ----
        tol = 0.0001
        km = 1000
        COMo = 0.5 * FCI
        err = 0.02
        k = 0
        m = [0.0] * 21
        COM = 0.0
        raw_material_cost_items = [
            (name, local_cost_per_h / X * h * d)
            for name, local_cost_per_h in raw_materials
        ]
        utility_cost_items = [
            (name, qty * cost / X * h * d)
            for name, qty, cost in utility_items if qty or cost
        ]
        process_consumable_cost_items = [
            (name, amount) for name, _unit, amount in process_consumables
        ]
        process_consumable_cost_total = sum(amount for _, amount in process_consumable_cost_items)
        total_utility_cost = sum(amount for _, amount in utility_cost_items)
        power_annual = sum(amount for name, amount in utility_cost_items
                           if 'power' in name.lower() or 'electric' in name.lower())
        other_utility_cost = total_utility_cost - power_annual
        while err > tol and k < km:
            COM = COMo
            # Raw material
            m[0] = sum(local_cost_per_h for _, local_cost_per_h in raw_materials) / X * h * d
            RM = m[0]
            # Utilities
            m[1] = other_utility_cost
            m[2] = 0.0
            m[3] = 0.0
            m[4] = power_annual
            m[5] = 0.0
            UC = sum(m[1:6])
            # Annual direct operating labor cost from the shift-coverage model.
            m[6] = annual_labor_cost
            OL = m[6]
            # Operating supervision
            m[7] = 0.15 * OL
            # Maintenance supplies
            m[8] = MS
            # Operating supplies
            m[9] = 0.15 * MS
            # Maintenance labor & supervision
            m[10] = 1.20 * MS
            # Laboratory charges
            m[11] = 0.15 * OL
            # Patents & royalties
            m[12] = 0.04 * COM
            # Direct Manufacturing Cost / Operating Cost
            DMC = sum(m[0:13]) + process_consumable_cost_total; TOC = DMC
            # Fixed manufacturing cost
            m[13] = 0.22 * (OL + MS)
            # Straight-line depreciation over the user-selected useful life.
            # It must not vary with the financial discount rate.
            m[14] = TDC / depreciation_life; DP = m[14]
            m[15] = 0.02 * TPI; LT = m[15]
            m[16] = 0.01 * TDC; IS = m[16]
            m[17] = sum(m[13:17]); FMC = m[17]
            # General expenses
            m[18] = 0.02 * COM; AC = m[18]
            m[19] = 0.10 * COM; DC = m[19]
            m[20] = 0.05 * COM; RD = m[20]
            GE = sum(m[18:21])
            COMF = DMC + FMC + GE
            err = abs((COMF - COM) / COMF) if COMF != 0 else 0.0
            k += 1
            COMo = COMF

        # Reconcile the converged cost model so every reported subtotal equals
        # the sum of its displayed components (including patents and insurance).
        variable_cost_base = sum(m[0:12]) + process_consumable_cost_total
        COM = (variable_cost_base + FMC) / 0.79
        m[12] = 0.04 * COM
        DMC = sum(m[0:13]) + process_consumable_cost_total; TOC = DMC
        m[18] = 0.02 * COM; AC = m[18]
        m[19] = 0.10 * COM; DC = m[19]
        m[20] = 0.05 * COM; RD = m[20]
        GE = sum(m[18:21])
        COM = DMC + FMC + GE
        COMd = COM - DP
        nV = prod_qty * h * d  # L/year
        CoPv = COM / nV if nV != 0 else 0.0

        # Annual power cost
        power_cost_annual = power_annual

        Sub = D / 100
        NCv = SPv * X

        # ---- Profitability ----
        SubPay = Sub * SPv * nV
        R = SPv * nV + SubPay
        GI = R - COM
        TR = C / 100
        NP = GI * (1 - TR)

        # ---- Convert ALL variables to Python primitives (bulletproof vs tkinter _NoValueType) ----
        try:
            p_int = int(float(p))
        except (ValueError, TypeError):
            p_int = 1
        if p_int <= startup_year:
            raise ValueError(
                f"Project life must extend beyond the construction/startup period "
                f"(production starts in year {startup_year})."
            )

        try:
            r_float = float(r)
        except (ValueError, TypeError):
            r_float = 0.0

        try:
            WC_float = float(WC)
        except (ValueError, TypeError):
            WC_float = 0.0

        try:
            TDC_float = float(TDC)
        except (ValueError, TypeError):
            TDC_float = 0.0

        try:
            DP_float = float(DP)
        except (ValueError, TypeError):
            DP_float = 0.0

        try:
            TPI_float = float(TPI)
        except (ValueError, TypeError):
            TPI_float = 0.0

        try:
            COMd_float = float(COMd)
        except (ValueError, TypeError):
            COMd_float = 0.0

        try:
            TR_float = float(TR)
        except (ValueError, TypeError):
            TR_float = 0.0

        try:
            R_float = float(R)
        except (ValueError, TypeError):
            R_float = 0.0

        try:
            NP_float = float(NP)
        except (ValueError, TypeError):
            NP_float = 0.0

        try:
            TCI_float = float(TCI)
        except (ValueError, TypeError):
            TCI_float = 0.0

        try:
            DNPW_float = float(DNPW) if 'DNPW' in dir() else 0.0
        except (ValueError, TypeError):
            DNPW_float = 0.0

        try:
            NNPW_float = float(NNPW) if 'NNPW' in dir() else 0.0
        except (ValueError, TypeError):
            NNPW_float = 0.0

        try:
            y_array = np.arange(0, p_int)
        except Exception:
            y_array = np.array([0])

        try:
            # For the default three-year startup, TDC is funded 60/40 in years
            # 0-1. If production starts sooner, fund TDC in year 0. Working
            # capital is added in the year immediately before production.
            I = np.zeros(p_int, dtype=float)
            if startup_year >= 3:
                I[0] = 0.60 * TDC_float
                I[1] = 0.40 * TDC_float
            else:
                I[0] = TDC_float
            I[startup_year - 1] += WC_float
            I[-1] -= WC_float  # negative expenditure means cash recovery
        except Exception as np_err:
            show_calc_error("Calculation Error",
                                 f"Failed to create cash flow arrays.\n"
                                 f"p_int = {p_int}\n"
                                 f"Error: {np_err}\n\n"
                                 f"Please ensure 'Project life in years' is a valid number >= 4.")
            return

        try:
            # Straight-line depreciation starts at production and is capped at
            # the depreciable basis if the project ends before its full life.
            dk = np.zeros(p_int, dtype=float)
            remaining_basis = max(TDC_float, 0.0)
            for year in range(startup_year, min(p_int, startup_year + int(depreciation_life))):
                dk[year] = min(max(DP_float, 0.0), remaining_basis)
                remaining_basis -= dk[year]
        except Exception as np_err:
            show_calc_error("Calculation Error",
                                 f"Failed to create depreciation array.\n"
                                 f"Error: {np_err}\n\n"
                                 f"Please check all input values.")
            return

        try:
            if p_int > startup_year:
                Rv = R_float * np.concatenate((np.zeros(startup_year), np.ones(p_int - startup_year)))
            else:
                Rv = np.zeros(p_int)
        except Exception as np_err:
            show_calc_error("Calculation Error",
                                 f"Failed to create revenue array.\n"
                                 f"Error: {np_err}\n\n"
                                 f"Please check all input values.")
            return

        try:
            if p_int > startup_year:
                revenue_factor = np.concatenate((np.zeros(startup_year), np.ones(p_int - startup_year)))
            else:
                revenue_factor = np.zeros(p_int)
            # Remove the modeled annual depreciation from COM to recover cash
            # operating expenses. Apply the selected year's tax depreciation
            # allowance separately when calculating taxable income.
            taxable_income = Rv - COMd_float * revenue_factor - dk
            tax_expense = taxable_income * TR_float
            Income = taxable_income - tax_expense + dk
            NCF = Income - I
        except Exception as np_err:
            show_calc_error("Calculation Error",
                                 f"Failed to compute cash flow.\n"
                                 f"Error: {np_err}\n\n"
                                 f"Please check all input values.")
            return

        try:
            CNCF = np.zeros(p_int)
            if p_int > 0:
                CNCF[0] = NCF[0]
                for i in range(1, p_int):
                    CNCF[i] = CNCF[i - 1] + NCF[i]
        except Exception as np_err:
            show_calc_error("Calculation Error",
                                 f"Failed to create cumulative cash flow.\n"
                                 f"Error: {np_err}\n\n"
                                 f"Please check all input values.")
            return

        try:
            dr = (1 + r_float) * np.ones(p_int)
            DCF = NCF / dr ** y_array
            CDCF = np.zeros(p_int)
            if p_int > 0:
                CDCF[0] = DCF[0]
                for i in range(1, p_int):
                    CDCF[i] = CDCF[i - 1] + DCF[i]
        except Exception as np_err:
            show_calc_error("Calculation Error",
                                 f"Failed to create discounted cash flow.\n"
                                 f"Error: {np_err}\n\n"
                                 f"Please check all input values.")
            return

        try:
            def payback_period(period_flows, cumulative_flows):
                if cumulative_flows[0] >= 0:
                    return 0.0
                for year in range(1, len(cumulative_flows)):
                    if cumulative_flows[year] >= 0 and period_flows[year] > 0:
                        return (year - 1) + (-cumulative_flows[year - 1] / period_flows[year])
                return None

            NPBP = payback_period(NCF, CNCF)
            DPBP = payback_period(DCF, CDCF)

            NNPW = CNCF[-1] if p_int > 0 else 0.0
            DNPW = CDCF[-1] if p_int > 0 else 0.0
            ROI = NP_float / TCI_float * 100 if TCI_float != 0 else 0.0
            def npv_at_rate(rate):
                log_one_plus_rate = math.log1p(rate)
                terms = []
                for year, flow in enumerate(NCF):
                    flow = float(flow)
                    if flow:
                        terms.append((math.log(abs(flow)) - year * log_one_plus_rate,
                                      math.copysign(1.0, flow)))
                if not terms:
                    return 0.0
                scale = max(log_value for log_value, _ in terms)
                scaled_sum = sum(sign * math.exp(log_value - scale)
                                 for log_value, sign in terms)
                if scaled_sum == 0:
                    return 0.0
                log_result = scale + math.log(abs(scaled_sum))
                if log_result > 709:
                    return math.copysign(float("inf"), scaled_sum)
                if log_result < -745:
                    return 0.0
                return math.copysign(math.exp(log_result), scaled_sum)

            # A conventional investment has one sign change: an initial net
            # outflow followed by net inflows. Only then is a unique project
            # IRR guaranteed; nonconventional streams are reported undefined.
            signs = [1 if flow > 0 else -1 for flow in NCF if flow != 0]
            sign_changes = sum(signs[i] != signs[i - 1] for i in range(1, len(signs)))
            if len(signs) >= 2 and signs[0] < 0 and sign_changes == 1:
                npv_zero = npv_at_rate(0.0)
                bracketed = False
                if abs(npv_zero) <= max(1.0, abs(TCI_float)) * 1e-12:
                    IRR = 0.0
                elif npv_zero < 0:
                    low, high = -0.999999, 0.0
                    bracketed = npv_at_rate(low) > 0
                else:
                    low, high = 0.0, 1.0
                    high_npv = npv_at_rate(high)
                    while high_npv > 0 and high < 1_000_000:
                        high *= 2.0
                        high_npv = npv_at_rate(high)
                    bracketed = high_npv < 0
                if npv_zero != 0 and bracketed:
                    for _ in range(200):
                        mid = (low + high) / 2.0
                        mid_npv = npv_at_rate(mid)
                        if mid_npv > 0:
                            low, low_npv = mid, mid_npv
                        else:
                            high = mid
                        if abs(high - low) <= 1e-12 * max(1.0, abs(mid)):
                            break
                    IRR = (low + high) / 2.0 * 100.0
                elif npv_zero != 0:
                    IRR = None
            else:
                IRR = None

            # Conventional benefit-cost ratios compare present worth of gross
            # benefits with present worth of costs, not positive/negative NCF.
            cash_operating_costs = np.full(p_int, max(COMd_float, 0.0)) * revenue_factor
            taxes = tax_expense
            benefit_flows = Rv + np.maximum(-taxes, 0.0)
            cost_flows = I + cash_operating_costs + np.maximum(taxes, 0.0)

            def benefit_cost_ratio(discount_rate):
                factors = (1.0 + discount_rate) ** y_array
                pv_benefits = float(np.sum(benefit_flows / factors))
                pv_costs = float(np.sum(cost_flows / factors))
                return pv_benefits / pv_costs if pv_costs > 0 else None

            B_C_ratio_NCF = benefit_cost_ratio(0.0)
            B_C_ratio_DCF = benefit_cost_ratio(r_float)
        except Exception as np_err:
            show_calc_error("Calculation Error",
                                 f"Failed to compute profitability metrics.\n"
                                 f"Error: {np_err}\n\n"
                                 f"Please check all input values are valid numbers.")
            return

        return {
            "depreciation_life_years": depreciation_life,
            "startup_year": startup_year,
            "operators_per_shift": operators_per_shift,
            "operators_per_shift_estimate": nOL,
            "total_operators": total_operators,
            "operator_hourly_wage_local": operator_hourly_wage,
            "operator_shift_hours": shift_hours,
            # CAPEX
            "PCE": PCE, "DPC": DPC, "TPC": TPC,
            "CAPEX_installations": f[1], "CAPEX_instrumentation": f[2],
            "CAPEX_piping": f[3], "CAPEX_electrical": f[4],
            "CAPEX_buildings": f[5], "CAPEX_yard": f[6], "CAPEX_services": f[7],
            "IPC": TPC - DPC, "CAPEX_engineering": f[9],
            "CAPEX_construction": f[10],
            "contractor_fee": f[12], "contingency": f[13],
            "FCI": FCI, "TDC": TDC,
            "WC": WC, "TCI": TCI, "TPI": TPI, "DPI": DPI,
            # OPEX / COM
            "COM": COM, "COMd": COMd, "DMC": DMC, "TOC": TOC, "FMC": FMC,
            "GE": GE, "RM": RM, "UC": UC, "PC": power_cost_annual,
            "process_consumable_cost_total": process_consumable_cost_total,
            "process_consumable_cost_items": process_consumable_cost_items,
            "OL": OL, "OS": m[7], "MS": MS, "OSUP": m[9], "MLS": m[10],
            "LAB": m[11], "PR": m[12], "maintenance_overhead": m[13],
            "depreciation": m[14], "local_taxes": m[15], "insurance": m[16],
            "admin_costs": m[18], "distribution_costs": m[19],
            "research_costs": m[20], "raw_material_cost_items": raw_material_cost_items,
            "utility_cost_items": utility_cost_items, "CoPv": CoPv,
            # Profitability
            "NNPW": NNPW, "DNPW": DNPW, "NPBP": NPBP, "DPBP": DPBP,
            "ROI": ROI, "IRR": IRR,
            "B_C_ratio_NCF": B_C_ratio_NCF, "B_C_ratio_DCF": B_C_ratio_DCF,
            "R": R, "GI": GI, "NP": NP, "currency_code": currency_code,
            "exchange_rate": X,
            # Cash flows
            "NCF": NCF, "CNCF": CNCF, "DCF": DCF, "CDCF": CDCF,
            "y": y_array, "years": [int(v) for v in y_array],
            # Inputs echo
            "inputs": inputs,
        }

    except Exception as e:
        show_calc_error("Calculation Error", f"An unexpected error occurred:\n{str(e)}\n\nPlease check all input values and try again.")
        return None


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------

class TOPCEEApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("TOPCEE v2.0 - Process Cost & Economic Evaluator")
        self.root.geometry("1300x850")
        self.root.minsize(1000, 700)
        self.root.configure(bg="#f3f6fa")

        # State
        self.equipment_items: List[EquipmentItem] = []
        self.equipment_total = 0.0
        self.equipment_db: Optional[EquipmentCostDB] = None
        self.calc_result: Optional[Dict[str, Any]] = None
        self.last_folder = load_config().get("last_folder", os.path.expanduser("~"))
        self.db_path = load_config().get("db_path", DEFAULT_DB_PATH)

        self._build_ui()
        self._init_equipment_db()

    # ------------------------------------------------------------------ UI --
    def _build_ui(self):
        self.root.option_add("*Font", ("Segoe UI", 9))
        self.root.option_add("*Label.Background", "#f3f6fa")
        self.root.option_add("*Label.Foreground", "#24364b")
        self.root.option_add("*LabelFrame.Background", "#f3f6fa")
        self.root.option_add("*LabelFrame.Foreground", "#164e78")
        self.root.option_add("*Button.Relief", "flat")
        # Style
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background="#f3f6fa")
        style.configure("TLabel", background="#f3f6fa", foreground="#24364b", font=("Segoe UI", 9))
        style.configure("TLabelframe", background="#f3f6fa", bordercolor="#d5dfeb")
        style.configure("TLabelframe.Label", background="#f3f6fa", foreground="#164e78",
                        font=("Segoe UI", 10, "bold"))
        style.configure("TNotebook", background="#e7edf5", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(16, 9), font=("Segoe UI", 9, "bold"))
        style.map("TNotebook.Tab", background=[("selected", "#ffffff")],
                  foreground=[("selected", "#164e78")])
        style.configure("TEntry", padding=6, fieldbackground="#ffffff")
        style.configure("TButton", padding=(10, 7), font=("Segoe UI", 9))
        style.configure("Accent.TButton", background="#176b87", foreground="#ffffff",
                        padding=(12, 8), font=("Segoe UI", 9, "bold"))
        style.map("Accent.TButton", background=[("active", "#10546c")],
                  foreground=[("active", "#ffffff")])
        style.configure("SectionHint.TLabel", foreground="#52667a", font=("Segoe UI", 10))
        style.configure("WelcomeCard.TLabelframe", background="#ffffff", bordercolor="#d7e1ec")
        style.configure("WelcomeCard.TLabelframe.Label", background="#ffffff",
                        foreground="#164e78", font=("Segoe UI", 11, "bold"))
        style.configure("WelcomeBody.TLabel", background="#ffffff", foreground="#52667a",
                        font=("Segoe UI", 9), wraplength=340)
        style.configure("Hero.TFrame", background="#12304a")
        style.configure("GuideCard.TLabelframe", background="#ffffff", bordercolor="#d7e1ec")
        style.configure("GuideCard.TLabelframe.Label", background="#ffffff",
                        foreground="#164e78", font=("Segoe UI", 11, "bold"))
        style.configure("GuideBody.TLabel", background="#ffffff", foreground="#34495e",
                        font=("Segoe UI", 10))
        style.configure("GuideTitle.TLabel", background="#f3f6fa", foreground="#12304a",
                        font=("Segoe UI", 22, "bold"))
        style.configure("GuideSubtitle.TLabel", background="#f3f6fa", foreground="#52667a",
                        font=("Segoe UI", 10))

        # Notebook
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Tab 1: Welcome
        self.tab_welcome = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_welcome, text="Welcome")
        self._build_welcome_tab()

        # Tab 2: Equipment Selection & Costing
        self.tab_equip = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_equip, text="Equipment Costing")
        self._build_equipment_tab()

        # Tab 3: Process Inputs, split into two focused sub-tabs
        self.tab_inputs = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_inputs, text="Process Inputs")
        self._build_inputs_tab()

        # Tab 4: Results
        self.tab_results = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_results, text="Results")
        self._build_results_tab()

        # Tab 5: Sensitivity Analysis (explore scenarios before final reporting)
        self.tab_sensitivity = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_sensitivity, text="Sensitivity Analysis")
        self._build_sensitivity_tab()

        # Tab 6: Report Generation
        self.tab_report = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_report, text="Generate Report")
        self._build_report_tab()

        # Tab 7: User Guide
        self.tab_guide = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_guide, text="User Guide")
        self._build_user_guide_tab()

        # Tab 8: About TOPCEE
        self.tab_about = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_about, text="About TOPCEE")
        self._build_about_tab()

        self.notebook.bind("<<NotebookTabChanged>>", self._on_notebook_tab_changed, add="+")

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN,
                          anchor=tk.W, font=("TkDefaultFont", 9))
        status.pack(side=tk.BOTTOM, fill=tk.X)

        # Bottom nav
        nav = ttk.Frame(self.root)
        nav.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)
        tk.Button(nav, text="Previous Tab", command=self._prev_tab).pack(side=tk.LEFT, padx=5)
        tk.Button(nav, text="Next Tab", command=self._next_tab).pack(side=tk.RIGHT, padx=5)

    def _prev_tab(self):
        idx = self.notebook.index(self.notebook.select())
        self.notebook.select((idx - 1) % self.notebook.index("end"))

    def _next_tab(self):
        idx = self.notebook.index(self.notebook.select())
        self.notebook.select((idx + 1) % self.notebook.index("end"))

    # ------------------------------------------------- Welcome Tab ----------
    def _build_welcome_tab(self):
        frame = self.tab_welcome
        hero = tk.Frame(frame, bg="#12304a", padx=34, pady=28)
        hero.pack(fill=tk.X, padx=18, pady=(18, 12))
        tk.Label(hero, text="TOPCEE", bg="#12304a", fg="#69d3c0",
                 font=("Segoe UI", 12, "bold")).pack(anchor="w")
        tk.Label(hero, text="Dr. Oyegoke welcomes you", bg="#12304a", fg="#d8e5ef",
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(7, 0))
        tk.Label(hero, text="Process economics, made clear.", bg="#12304a", fg="#ffffff",
                 font=("Segoe UI", 24, "bold")).pack(anchor="w", pady=(3, 4))
        tk.Label(hero,
                 text="Estimate equipment investment, model operating costs and cash flows, test key assumptions, and export a complete process economics report.",
                 bg="#12304a", fg="#d8e5ef", font=("Segoe UI", 11),
                 wraplength=1000, justify=tk.LEFT).pack(anchor="w")
        hero_actions = ttk.Frame(hero, style="Hero.TFrame")
        hero_actions.pack(anchor="w", pady=(16, 0))
        ttk.Button(hero_actions, text="Start with Equipment Costing",
                   command=lambda: self.notebook.select(self.tab_equip),
                   style="Accent.TButton").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(hero_actions, text="Review Input Details",
                   command=self._open_input_details).pack(side=tk.LEFT)

        cards_frame = ttk.Frame(frame, padding=(18, 4, 18, 8))
        cards_frame.pack(fill=tk.BOTH, expand=True)
        cards = [
            ("01  Equipment Costing",
             "Browse the equipment cost database, then scale selected equipment to your design capacity. Reference capacity and scaling exponent are filled from the database where available; keep each equipment item linked to its PFD Process ID.",
             "Open Equipment Costing", lambda: self.notebook.select(self.tab_equip)),
            ("02  Project & Process Inputs",
             "Describe the project, choose a local currency, and set project life, discount and tax assumptions, depreciation years, startup year, operating shifts and days, and hourly labour cost. Save or reload input files as JSON.",
             "Open Input Details", self._open_input_details),
            ("03  Operating Cost Parameters",
             "Enter up to 10 raw materials, 10 utilities, and 10 process consumables. Specify flow and price units, local-currency prices, and continuous-use or replacement costing; unused zero-value rows are omitted.",
             "Open Operating Costs", lambda: self._open_input_details_tab(2)),
            ("04  TEA Results & Cash Flow",
             "Review CAPEX / Total Capital Investment, OPEX / cost of manufacture, profitability indicators, and annual nominal and discounted cash flows, including the cash-flow chart.",
             "View Results", lambda: self.notebook.select(self.tab_results)),
            ("05  Sensitivity Analysis",
             "Vary a selected numeric economic or operating-cost input from −100% to +100% in 10% steps. Choose one, two, or three profitability responses and view the results as a table and figure.",
             "Open Sensitivity Analysis", lambda: self.notebook.select(self.tab_sensitivity)),
            ("06  Reports & Exports",
             "Generate the full analysis as HTML, PDF, CSV, and Excel, with project details, cost breakdowns, equipment and input schedules, cash-flow data, and charts in supported formats.",
             "Generate Reports", lambda: self.notebook.select(self.tab_report)),
        ]
        for index, (title, description, action, command) in enumerate(cards):
            card = ttk.LabelFrame(cards_frame, text=title, style="WelcomeCard.TLabelframe", padding=16)
            card.grid(row=index // 3, column=index % 3, sticky="nsew",
                      padx=8, pady=8, ipadx=4, ipady=4)
            ttk.Label(card, text=description, style="WelcomeBody.TLabel",
                      justify=tk.LEFT).pack(anchor="w", fill=tk.X, pady=(0, 14))
            ttk.Button(card, text=action, command=command).pack(anchor="w")
        for col in range(3):
            cards_frame.columnconfigure(col, weight=1, uniform="welcome_cards")
        for row in range(2):
            cards_frame.rowconfigure(row, weight=1, uniform="welcome_cards")

        tk.Label(frame,
                 text=("TOPCEE  •  Chemical Engineering, ABU Zaria  •  "
                       "Pencil Team - CAD Engineering of Processes & Reactive Interface (PT-CEPRIs) Group"),
                 bg="#f3f6fa", fg="#708195", font=("Segoe UI", 9)).pack(pady=(0, 12))

    def _open_input_details(self):
        self.notebook.select(self.tab_inputs)
        self.inputs_notebook.select(self.inputs_details_tab)

    def _open_input_details_tab(self, tab_index: int):
        self.notebook.select(self.tab_inputs)
        self.inputs_notebook.select(tab_index)

    # ---------------------------------------------- Equipment Costing Tab ---
    def _build_equipment_tab(self):
        frame = self.tab_equip

        # Top controls
        ctrl = tk.Frame(frame)
        ctrl.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)

        tk.Label(ctrl, text="Database:").pack(side=tk.LEFT)
        self.eq_db_var = tk.StringVar(value=self.db_path)
        tk.Entry(ctrl, textvariable=self.eq_db_var, width=50).pack(side=tk.LEFT, padx=4)
        tk.Button(ctrl, text="Browse", command=self._eq_browse_db).pack(side=tk.LEFT, padx=2)
        tk.Button(ctrl, text="Load Sheet", command=self._eq_load_sheet).pack(side=tk.LEFT, padx=2)

        tk.Label(ctrl, text="Sheet:").pack(side=tk.LEFT, padx=(10, 0))
        self.eq_sheet_var = tk.StringVar()
        self.eq_sheet_combo = ttk.Combobox(ctrl, textvariable=self.eq_sheet_var,
                                           state="readonly", width=22)
        self.eq_sheet_combo.pack(side=tk.LEFT, padx=2)
        self.eq_sheet_combo.bind("<<ComboboxSelected>>", lambda e: self._eq_display_sheet())

        tk.Label(ctrl, text="Search:").pack(side=tk.LEFT, padx=(10, 0))
        self.eq_search_var = tk.StringVar()
        self.eq_search_var.trace_add("write", lambda *a: self._eq_filter_tree())
        tk.Entry(ctrl, textvariable=self.eq_search_var, width=25).pack(side=tk.LEFT, padx=4)

        # Main split: left = DB browser, right = spec & project
        main = tk.PanedWindow(frame, orient=tk.HORIZONTAL, sashrelief=tk.RAISED)
        main.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        # --- Left: Equipment database (scrollable, all columns) ---
        left = ttk.LabelFrame(main, text="Equipment Database — click a row to select")
        main.add(left, width=500)

        self.eq_tree = ttk.Treeview(left, show="headings", height=15)
        self.eq_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        eq_vsb = ttk.Scrollbar(left, orient="vertical", command=self.eq_tree.yview)
        self.eq_tree.configure(yscroll=eq_vsb.set)
        eq_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.eq_tree.bind("<<TreeviewSelect>>", lambda e: self._eq_on_row_select())

        # --- Right: Selected equipment details + specification + project list ---
        right = ttk.Frame(main)
        main.add(right, width=620)

        # Selected equipment details
        det = ttk.LabelFrame(right, text="Selected Equipment (from database)")
        det.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        self.eq_detail_text = tk.Text(det, height=6, wrap=tk.WORD, font=("Consolas", 9))
        self.eq_detail_text.pack(fill=tk.X, padx=5, pady=5)
        self.eq_detail_text.config(state='disabled')

        # Specification entry
        spec = ttk.LabelFrame(right, text="Your Design Specification")
        spec.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        spec_labels = [
            ("Process ID (PFD):", "process_id_var", ""),
            ("Ref Capacity (from DB):", "refcap_var", ""),
            ("Ref Cost $ (from DB):", "refcost_var", ""),
            ("Design Capacity:", "cap_var", ""),
            ("Scaling Exponent (auto-filled):", "exp_var", "0.6"),
            ("Quantity:", "qty_var", "1"),
            ("Material Factor:", "mat_var", "1.0"),
            ("Pressure Factor:", "pres_var", "1.0"),
            ("Ref Year (from Source):", "ref_year_var", ""),
            ("Ref CEPCI:", "ref_cepci_var", ""),
            ("Target Year:", "target_year_var", str(datetime.datetime.now().year)),
            ("Target CEPCI (auto):", "target_cepci_var", ""),
        ]
        self.eq_spec_vars: Dict[str, tk.StringVar] = {}
        self.eq_spec_entries: Dict[str, tk.Entry] = {}
        grid = tk.Frame(spec)
        grid.pack(fill=tk.X, padx=5, pady=5)
        for i, (label, varname, default) in enumerate(spec_labels):
            tk.Label(grid, text=label, anchor="w").grid(row=i // 2, column=(i % 2) * 2,
                                            sticky="e", padx=4, pady=3)
            var = tk.StringVar(value=default)
            self.eq_spec_vars[varname] = var
            entry = tk.Entry(grid, textvariable=var, width=14)
            entry.grid(
                row=i // 2, column=(i % 2) * 2 + 1, sticky="ew", padx=4, pady=3)
            self.eq_spec_entries[varname] = entry
        for col in (1, 3):
            grid.columnconfigure(col, weight=1)
        self.eq_spec_vars["target_year_var"].trace_add(
            "write", lambda *_: self._eq_update_target_cepci()
        )
        self.eq_spec_vars["ref_year_var"].trace_add(
            "write", lambda *_: self._eq_update_reference_cepci()
        )
        self.eq_exp_source_var = tk.StringVar(
            value="Select a database row to auto-fill its exponent (0.6 used when none is listed)."
        )
        self.eq_refcap_source_var = tk.StringVar(value="Reference capacity is read from the selected database row.")
        self.eq_scale_unit_var = tk.StringVar(value="")
        self.eq_cost_index_source_var = tk.StringVar(
            value="Target CEPCI is estimated by linear regression of reported annual CEPCI values (2007–2023)."
        )
        self.eq_spec_hint_var = tk.StringVar()
        for hint_var in (
            self.eq_exp_source_var, self.eq_refcap_source_var,
            self.eq_scale_unit_var, self.eq_cost_index_source_var,
        ):
            hint_var.trace_add("write", lambda *_: self._eq_refresh_spec_hint())
        tk.Label(spec, textvariable=self.eq_spec_hint_var, fg="#52667a",
                 anchor="w", justify=tk.LEFT, wraplength=560).pack(fill=tk.X, padx=8, pady=(0, 4))
        self._eq_refresh_spec_hint()
        self._eq_update_target_cepci()

        # Auto-calculated unit cost display
        self.eq_unitcost_var = tk.StringVar(value="Unit Cost: $ 0.00")
        tk.Label(spec, textvariable=self.eq_unitcost_var,
                 font=("TkDefaultFont", 10, "bold"), fg="#1a3c5a").pack(pady=3)

        btns = tk.Frame(spec)
        btns.pack(fill=tk.X, padx=5, pady=4)
        tk.Button(btns, text="Calculate Cost", command=self._eq_calc_cost,
                  bg="#2196F3", fg="white").pack(side=tk.LEFT, padx=3)
        tk.Button(btns, text="Add to Project List", bg="#4CAF50", fg="white",
                  command=self._eq_add_to_project).pack(side=tk.LEFT, padx=3)

        # Keep the equipment actions pinned to the bottom of the pane so they
        # remain reachable even when the specification and project list are tall.
        footer = tk.Frame(right)
        footer.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=4)

        tot_frame = tk.Frame(footer)
        tot_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 4))
        self.eq_total_var = tk.StringVar(value="Total Equipment Cost: $ 0.00")
        tk.Label(tot_frame, textvariable=self.eq_total_var,
                 font=("TkDefaultFont", 11, "bold")).pack(side=tk.LEFT)
        tk.Button(tot_frame, text="Remove Selected",
                  command=self._eq_remove_selected).pack(side=tk.RIGHT, padx=3)
        tk.Button(tot_frame, text="Clear All",
                  command=self._eq_clear_project).pack(side=tk.RIGHT, padx=3)

        act = tk.Frame(footer)
        act.pack(side=tk.TOP, fill=tk.X)
        tk.Button(act, text="Use Total as PCE in Process Inputs",
                  bg="#2196F3", fg="white", font=("TkDefaultFont", 10, "bold"),
                  command=self._eq_use_total_as_pce).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        tk.Button(act, text="Export Equipment List (CSV)",
                  bg="#4CAF50", fg="white", activebackground="#3d9140",
                  activeforeground="white", command=self._eq_export_csv).pack(
                      side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

        # Project list
        plist = ttk.LabelFrame(right, text="Project Equipment List")
        plist.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)

        pcols = ["Name", "Process ID", "Category", "Ref Year", "Ref CEPCI", "Target Year",
                 "Target CEPCI", "Index Ratio", "Ref Cap", "Ref Cost $", "Design Cap",
                 "Exponent", "Qty", "Unit Cost $", "Total Cost $"]
        self.eq_proj_tree = ttk.Treeview(plist, columns=pcols, show="headings", height=10)
        widths = [115, 95, 80, 75, 75, 75, 85, 75, 70, 75, 75, 60, 45, 85, 90]
        for c, w in zip(pcols, widths):
            self.eq_proj_tree.heading(c, text=c)
            self.eq_proj_tree.column(c, width=w, anchor='e' if c not in ("Name", "Category") else 'w')
        self.eq_proj_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=4)
        p_vsb = ttk.Scrollbar(plist, orient="vertical", command=self.eq_proj_tree.yview)
        self.eq_proj_tree.configure(yscroll=p_vsb.set)
        p_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        p_hsb = ttk.Scrollbar(plist, orient="horizontal", command=self.eq_proj_tree.xview)
        self.eq_proj_tree.configure(xscrollcommand=p_hsb.set)
        p_hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.eq_proj_tree.bind("<Double-1>", self._eq_edit_process_id)

    # ---------------------------------------------- Equipment Tab Helpers ---
    def _init_equipment_db(self):
        try:
            self.equipment_db = EquipmentCostDB(self.db_path)
            self.equipment_db.load()
            self.eq_sheet_combo["values"] = self.equipment_db.sheets
            if self.equipment_db.sheets:
                self.eq_sheet_var.set(self.equipment_db.sheets[0])
            self.eq_db_var.set(self.db_path)
            self.status_var.set(f"Database loaded: {len(self.equipment_db.sheets)} sheet(s).")
        except Exception as e:
            self.status_var.set(f"Equipment DB not loaded: {e}")

    def _eq_browse_db(self):
        path = filedialog.askopenfilename(
            title="Select Equipment Cost Database",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.eq_db_var.set(path)
            self.db_path = path
            self._eq_load_sheet()

    def _eq_load_sheet(self):
        try:
            self.equipment_db = EquipmentCostDB(self.eq_db_var.get())
            self.equipment_db.load()
            self.eq_sheet_combo["values"] = self.equipment_db.sheets
            if self.equipment_db.sheets:
                self.eq_sheet_var.set(self.equipment_db.sheets[0])
                self._eq_display_sheet()
            messagebox.showinfo("Loaded",
                                f"Database loaded: {len(self.equipment_db.sheets)} sheet(s).\n"
                                f"Select a sheet from the dropdown.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load database:\n{e}")

    def _eq_display_sheet(self):
        """Display ALL columns from the selected sheet in a scrollable treeview."""
        if not self.equipment_db or not self.equipment_db.loaded:
            return
        sheet = self.eq_sheet_var.get()
        if not sheet:
            return
        try:
            # Load raw sheet data without any column mapping
            ws = self.equipment_db.workbook[sheet]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                messagebox.showwarning("Empty", "Sheet is empty.")
                return

            header = [str(c) if c is not None else f"Col{idx+1}" for idx, c in enumerate(rows[0])]
            # Drop trailing empty columns
            while header and header[-1].strip() == "":
                header.pop()

            # Store raw data for reference
            self.eq_sheet_header = header
            self.eq_sheet_rows = [list(r) for r in rows[1:] if any(v is not None and str(v).strip() != "" for v in r)]

            # Configure treeview columns
            self.eq_tree["columns"] = header
            self.eq_tree["show"] = "headings"
            for c in header:
                self.eq_tree.heading(c, text=c)
                self.eq_tree.column(c, width=110, anchor='w')

            # Clear and repopulate
            self.eq_tree.delete(*self.eq_tree.get_children())
            for r in self.eq_sheet_rows:
                vals = [str(v) if v is not None else "" for v in r]
                # Pad to header length
                while len(vals) < len(header):
                    vals.append("")
                self.eq_tree.insert("", tk.END, values=vals[:len(header)])

            # Clear detail and spec
            self.eq_detail_text.config(state='normal')
            self.eq_detail_text.delete("1.0", tk.END)
            self.eq_detail_text.insert("1.0", "Click a row to see its details.")
            self.eq_detail_text.config(state='disabled')
            self.status_var.set(f"Sheet '{sheet}': {len(self.eq_sheet_rows)} rows, {len(header)} columns.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to display sheet:\n{e}")

    def _eq_filter_tree(self):
        search = self.eq_search_var.get().lower()
        self.eq_tree.delete(*self.eq_tree.get_children())
        if not hasattr(self, 'eq_sheet_rows'):
            return
        for r in self.eq_sheet_rows:
            row_str = " ".join(str(v) if v is not None else "" for v in r).lower()
            if search in row_str:
                vals = [str(v) if v is not None else "" for v in r]
                while len(vals) < len(self.eq_sheet_header):
                    vals.append("")
                self.eq_tree.insert("", tk.END, values=vals[:len(self.eq_sheet_header)])

    def _eq_on_row_select(self):
        """Show the selected row's details in the text panel."""
        sel = self.eq_tree.selection()
        if not sel or not hasattr(self, 'eq_sheet_header'):
            return
        idx = self.eq_tree.index(sel[0])
        # Find the corresponding raw row (accounting for filtering)
        search = self.eq_search_var.get().lower()
        filtered = [r for r in self.eq_sheet_rows
                    if search in " ".join(str(v) if v is not None else "" for v in r).lower()]
        if not (0 <= idx < len(filtered)):
            return
        row = filtered[idx]
        self.eq_calculated = None
        self.eq_unitcost_var.set("Equipment selected — calculate its cost before adding it.")
        lines = []
        for h, v in zip(self.eq_sheet_header, row):
            lines.append(f"{h}: {v if v is not None else ''}")
        self.eq_selected_row = row
        self.eq_selected_header = self.eq_sheet_header

        self.eq_detail_text.config(state='normal')
        self.eq_detail_text.delete("1.0", tk.END)
        self.eq_detail_text.insert("1.0", "\n".join(lines))
        self.eq_detail_text.config(state='disabled')

        # Try to auto-fill ref capacity and ref cost from detected columns
        self._eq_autofill_from_row(row)

    def _eq_autofill_from_row(self, row):
        """Read reference capacity, cost, and exponent from the selected DB record."""
        if not hasattr(self, 'eq_selected_header'):
            return
        headers = [" ".join(str(h).lower().replace("_", " ").replace("-", " ").split())
                   for h in self.eq_selected_header]

        def number_at(index):
            if index >= len(row) or row[index] is None:
                return None
            try:
                return float(str(row[index]).replace(",", "").strip())
            except (ValueError, TypeError):
                return None

        def pick_numeric(exact=(), contains=()):
            for matcher in (lambda h: h in exact,
                            lambda h: any(token in h for token in contains)):
                for index, header in enumerate(headers):
                    value = number_at(index)
                    if value is not None and matcher(header):
                        return value, self.eq_selected_header[index]
            return None, None

        # Prefer a named reference/base value. This workbook uses Min Scale and
        # Min Cost, which are the paired reference point for its scaling rule.
        refcap, refcap_header = pick_numeric(
            exact=("reference capacity", "ref capacity", "base capacity", "rated capacity",
                   "capacity ref", "capacity base", "min scale", "minimum scale",
                   "scale min", "min capacity", "minimum capacity"),
            contains=("reference capacity", "ref capacity", "base capacity", "min scale", "minimum scale"),
        )
        if refcap is None:
            refcap, refcap_header = pick_numeric(
                contains=("capacity", "size", "rating", "rated"),
            )
        if refcap is None:
            refcap, refcap_header = pick_numeric(
                exact=("max scale", "maximum scale", "scale max", "max capacity", "maximum capacity"),
                contains=("max scale", "maximum scale"),
            )

        refcost, refcost_header = pick_numeric(
            exact=("reference cost", "ref cost", "base cost", "purchase cost", "min cost", "minimum cost"),
            contains=("reference cost", "ref cost", "base cost", "min cost", "minimum cost"),
        )
        if refcost is None:
            refcost, refcost_header = pick_numeric(contains=("cost", "price", "usd"))
        if refcost is None:
            refcost, refcost_header = pick_numeric(
                exact=("max cost", "maximum cost"), contains=("max cost", "maximum cost"))

        exponent_val, exponent_header = pick_numeric(
            exact=("n", "m", "scaling factor", "scale factor"),
            contains=("scaling exponent", "scale exponent", "capacity exponent",
                      "cost exponent", "exponent", "scaling factor", "scale factor"),
        )

        source_index = next(
            (i for i, h in enumerate(headers)
             if h in ("source", "reference", "citation", "cost source", "reference source")),
            None,
        )
        source_text = ""
        if source_index is not None and source_index < len(row) and row[source_index] is not None:
            source_text = str(row[source_index]).strip()
        reference_year = parse_reference_year(source_text)

        def update_reference_field(varname, value, found_header):
            entry = self.eq_spec_entries[varname]
            entry.configure(state="normal")
            self.eq_spec_vars[varname].set("" if value is None else f"{value:.4g}")
            entry.configure(state="readonly" if value is not None else "normal")

        update_reference_field("refcap_var", refcap, refcap_header)
        update_reference_field("refcost_var", refcost, refcost_header)
        update_reference_field("ref_year_var", reference_year, "Source" if reference_year else None)
        # Allow correction if the Source citation contains multiple years or the
        # detected year is not the year associated with the equipment cost.
        self.eq_spec_entries["ref_year_var"].configure(state="normal")
        if reference_year is None:
            self.eq_cost_index_source_var.set(
                "No year was found in the Source citation. Enter Ref Year; Ref CEPCI will be derived from TOPCEE's reported data or regression."
            )
        else:
            self.eq_cost_index_source_var.set(
                f"Ref Year {reference_year} parsed from the Source citation. Ref CEPCI is derived from TOPCEE's CEPCI data/regression; the equipment database CEPCI column is ignored."
            )
        if refcap is None:
            self.eq_refcap_source_var.set(
                "No reference scale column was found in this row; enter Ref Capacity manually."
            )
        else:
            self.eq_refcap_source_var.set(
                f"Ref Capacity auto-filled from '{refcap_header}'."
            )

        unit_value, _unit_header = next(
            ((row[i], self.eq_selected_header[i]) for i, h in enumerate(headers)
             if i < len(row) and (h == "scaling unit" or h == "scale unit") and row[i] not in (None, "")),
            (None, None),
        )
        self.eq_scale_unit_var.set(
            f"Database scale unit: {unit_value}. Enter Design Capacity in the same unit."
            if unit_value is not None else "Enter Design Capacity in the same unit as Ref Capacity."
        )

        if exponent_val is None:
            self.eq_spec_vars["exp_var"].set("0.6")
            self.eq_exp_source_var.set(
                "No scaling exponent found in the selected database row; using 0.6. You can edit it."
            )
        else:
            self.eq_spec_vars["exp_var"].set(f"{exponent_val:g}")
            self.eq_exp_source_var.set(
                f"Auto-filled from database column '{exponent_header}'. You can edit it if needed."
            )

    def _eq_refresh_spec_hint(self):
        """Render the equipment guidance messages as one flowing paragraph."""
        hints = (
            self.eq_exp_source_var.get(),
            self.eq_refcap_source_var.get(),
            self.eq_scale_unit_var.get(),
            self.eq_cost_index_source_var.get(),
        )
        self.eq_spec_hint_var.set(" ".join(text.strip() for text in hints if text.strip()))

    def _eq_update_target_cepci(self):
        """Update displayed target CEPCI from reported values or OLS regression."""
        try:
            year = int(float(self.eq_spec_vars["target_year_var"].get()))
            value, estimated, r_squared = cepci_for_year(year)
        except (ValueError, KeyError):
            return
        entry = self.eq_spec_entries.get("target_cepci_var")
        if entry is None:
            return
        self.eq_spec_vars["target_cepci_var"].set(f"{value:.2f}")
        if estimated:
            self.eq_cost_index_source_var.set(
                f"Target CEPCI {value:.2f} is an OLS regression estimate for {year} "
                f"(2007–2023 data; R²={r_squared:.3f})."
            )
        else:
            self.eq_cost_index_source_var.set(
                f"Target CEPCI {value:.2f} is the reported annual value for {year}."
            )

    def _eq_update_reference_cepci(self):
        """Derive Ref CEPCI solely from the source citation year and TOPCEE's model."""
        entry = self.eq_spec_entries.get("ref_cepci_var")
        if entry is None:
            return
        year_text = self.eq_spec_vars["ref_year_var"].get().strip()
        try:
            year = int(float(year_text))
            value, estimated, r_squared = cepci_for_year(year)
        except (ValueError, KeyError):
            entry.configure(state="normal")
            self.eq_spec_vars["ref_cepci_var"].set("")
            if hasattr(self, "eq_cost_index_source_var"):
                self.eq_cost_index_source_var.set(
                    "Enter Ref Year from the Source citation; TOPCEE will derive Ref CEPCI from its own CEPCI data/regression."
                )
            return
        entry.configure(state="normal")
        self.eq_spec_vars["ref_cepci_var"].set(f"{value:.2f}")
        entry.configure(state="readonly")
        method = "regression estimate" if estimated else "reported annual value"
        self.eq_cost_index_source_var.set(
            f"Ref CEPCI {value:.2f} derived from Ref Year {year} using TOPCEE's {method}"
            + (f" (OLS R²={r_squared:.3f}); database CEPCI is not used."
               if estimated else "; database CEPCI is not used.")
        )

    def _eq_calc_cost(self):
        """Calculate unit cost from specs and display it."""
        try:
            refcap = float(self.eq_spec_vars["refcap_var"].get() or 0)
            refcost = float(self.eq_spec_vars["refcost_var"].get() or 0)
            cap = float(self.eq_spec_vars["cap_var"].get() or 0)
            exp = float(self.eq_spec_vars["exp_var"].get() or 0.6)
            qty = int(self.eq_spec_vars["qty_var"].get() or 1)
            mat = float(self.eq_spec_vars["mat_var"].get() or 1.0)
            pres = float(self.eq_spec_vars["pres_var"].get() or 1.0)
            source_year_text = self.eq_spec_vars["ref_year_var"].get().strip()
            source_year = int(float(source_year_text)) if source_year_text else 0
            reference_cepci = float(self.eq_spec_vars["ref_cepci_var"].get() or 0)
            target_year = int(float(self.eq_spec_vars["target_year_var"].get()))
            target_cepci = float(self.eq_spec_vars["target_cepci_var"].get())
        except ValueError:
            messagebox.showerror("Input", "Enter valid numbers for capacities, factors, CEPCI, and target year.")
            return

        if refcost <= 0 or refcap <= 0 or cap <= 0:
            self.eq_unitcost_var.set("Unit Cost: $ 0.00 (enter ref cost, ref cap, design cap)")
            return
        if reference_cepci <= 0 or target_cepci <= 0:
            messagebox.showerror("Cost Index", "A positive Ref CEPCI and Target CEPCI are required.")
            return

        ratio = cap / refcap
        index_ratio = target_cepci / reference_cepci
        unit_cost = refcost * (ratio ** exp) * index_ratio * mat * pres
        total = unit_cost * qty
        self.eq_unitcost_var.set(
            f"Unit Cost: $ {unit_cost:,.2f}   |   Total (x{qty}): $ {total:,.2f}   |   CI: {index_ratio:.4f}"
        )
        generated_target, target_estimated, model_r2 = cepci_for_year(target_year)
        if abs(target_cepci - generated_target) > 0.005:
            method_note = f"User-entered Target CEPCI {target_cepci:.2f} used for {target_year}."
        elif target_estimated:
            method_note = (
                f"Target CEPCI {target_cepci:.2f} estimated by OLS linear regression of 2007–2023 data "
                f"(R²={model_r2:.3f})."
            )
        else:
            method_note = f"Target CEPCI {target_cepci:.2f} is the reported annual value for {target_year}."
        self.eq_calculated = {"unit_cost": unit_cost, "total_cost": total,
                              "refcap": refcap, "refcost": refcost,
                              "cap": cap, "exp": exp, "qty": qty,
                              "mat": mat, "pres": pres, "source_year": source_year,
                              "reference_cepci": reference_cepci, "target_year": target_year,
                              "target_cepci": target_cepci, "index_ratio": index_ratio,
                              "index_note": method_note}

    def _eq_add_to_project(self):
        if not hasattr(self, 'eq_calculated') or not self.eq_calculated:
            messagebox.showwarning("Calculate", "Please click 'Calculate Cost' first.")
            return
        if not hasattr(self, 'eq_selected_row'):
            messagebox.showwarning("Select", "Please select an equipment row from the database.")
            return

        calc = self.eq_calculated
        row = self.eq_selected_row
        hdr = getattr(self, 'eq_selected_header', [])
        # Get equipment name from first column or 'Unknown'
        name = str(row[0]) if row and len(row) > 0 else "Equipment"
        category = str(row[1]) if row and len(row) > 1 else ""
        source_text = ""
        source_index = next(
            (i for i, h in enumerate(hdr)
             if " ".join(str(h).lower().replace("_", " ").split()) in
             ("source", "reference", "citation", "cost source", "reference source")),
            None,
        )
        if source_index is not None and source_index < len(row) and row[source_index] is not None:
            source_text = str(row[source_index]).strip()

        item = EquipmentItem(
            name=name,
            process_id=self.eq_spec_vars["process_id_var"].get().strip(),
            category=category,
            equipment_type=category,
            reference_source=source_text,
            reference_capacity=calc["refcap"],
            reference_cost=calc["refcost"],
            capacity=calc["cap"],
            scaling_exponent=calc["exp"],
            cost_year=calc["source_year"],
            target_year=calc["target_year"],
            reference_cepci=calc["reference_cepci"],
            target_cepci=calc["target_cepci"],
            cost_index_ratio=calc["index_ratio"],
            quantity=calc["qty"],
            material_factor=calc["mat"],
            pressure_factor=calc["pres"],
            unit="",
            notes=(f"Ref CEPCI {calc['reference_cepci']:.2f} derived from Ref Year {calc['source_year']} using TOPCEE's CEPCI data/regression; "
                   f"{calc['index_note']}"
                   + (f" Source citation year: {calc['source_year']}." if calc["source_year"] else "")
                   + (f" Source: {source_text}." if source_text else "")),
        )
        self.equipment_items.append(item)
        self._eq_refresh_project_tree()
        messagebox.showinfo("Added", f"'{name}' added to project list.")

    def _eq_refresh_project_tree(self):
        self.eq_proj_tree.delete(*self.eq_proj_tree.get_children())
        total = 0.0
        for it in self.equipment_items:
            self.eq_proj_tree.insert("", tk.END, values=[
                it.name, it.process_id, it.category, it.cost_year or "—",
                f"{it.reference_cepci:.2f}", it.target_year, f"{it.target_cepci:.2f}",
                f"{it.cost_index_ratio:.4f}",
                f"{it.reference_capacity:.4g}", f"{it.reference_cost:.4g}",
                f"{it.capacity:.4g}", f"{it.scaling_exponent:.2f}",
                it.quantity, f"{it.scaled_cost():.2f}", f"{it.total_cost():.2f}",
            ])
            total += it.total_cost()
        self.equipment_total = total
        self.eq_total_var.set(f"Total Equipment Cost: $ {total:,.2f}")

    def _eq_edit_process_id(self, event):
        """Allow the PFD Process ID to be edited directly in the project list."""
        region = self.eq_proj_tree.identify("region", event.x, event.y)
        column = self.eq_proj_tree.identify_column(event.x)
        row_id = self.eq_proj_tree.identify_row(event.y)
        if region != "cell" or column != "#2" or not row_id:
            return
        bbox = self.eq_proj_tree.bbox(row_id, column)
        if not bbox:
            return
        x, y, width, height = bbox
        old_value = self.eq_proj_tree.set(row_id, "Process ID")
        editor = ttk.Entry(self.eq_proj_tree)
        editor.insert(0, old_value)
        editor.select_range(0, tk.END)
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()

        def save(_event=None):
            if not editor.winfo_exists():
                return
            value = editor.get().strip()
            item_index = self.eq_proj_tree.index(row_id)
            if 0 <= item_index < len(self.equipment_items):
                self.equipment_items[item_index].process_id = value
                self._eq_refresh_project_tree()
            editor.destroy()

        editor.bind("<Return>", save)
        editor.bind("<FocusOut>", save)
        editor.bind("<Escape>", lambda _event: editor.destroy())

    def _eq_remove_selected(self):
        sel = self.eq_proj_tree.selection()
        if not sel:
            return
        idx = self.eq_proj_tree.index(sel[0])
        if 0 <= idx < len(self.equipment_items):
            self.equipment_items.pop(idx)
            self._eq_refresh_project_tree()

    def _eq_clear_project(self):
        if messagebox.askyesno("Confirm", "Clear the entire project equipment list?"):
            self.equipment_items.clear()
            self._eq_refresh_project_tree()

    def _eq_export_csv(self):
        if not self.equipment_items:
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
                for it in self.equipment_items:
                    w.writerow(it.to_row())
                w.writerow([])
                w.writerow(["TOTAL"] + [""] * 17 + [f"{self.equipment_total:.2f}", ""])
            messagebox.showinfo("Saved", f"Exported to: {path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _eq_use_total_as_pce(self):
        if self.equipment_total <= 0:
            messagebox.showinfo("Info", "Equipment list is empty. Add equipment first.")
            return
        ent = self.input_entries.get('Total equipment cost ($)')
        if ent:
            ent.config(state='normal')
            ent.delete(0, tk.END)
            ent.insert(0, f"{self.equipment_total:,.2f}")
            ent.config(state='readonly')
        self.notebook.select(self.tab_inputs)  # switch to Process Inputs tab
        self.status_var.set(f"PCE set to ${self.equipment_total:,.2f}")

    # ---------------------------------------------------- Inputs Tab -------
    def _build_inputs_tab(self):
        """Build the split economic and operating-material input workspace."""
        sub_tabs = ttk.Notebook(self.tab_inputs)
        sub_tabs.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.inputs_notebook = sub_tabs
        details_tab = ttk.Frame(sub_tabs, padding=14)
        economic_tab = ttk.Frame(sub_tabs, padding=14)
        materials_tab = ttk.Frame(sub_tabs, padding=14)
        self.inputs_details_tab = details_tab
        sub_tabs.add(details_tab, text="Input Details")
        sub_tabs.add(economic_tab, text="Economic Parameters")
        sub_tabs.add(materials_tab, text="Operating Cost Parameters")
        self.input_entries: Dict[str, tk.Entry] = {}
        self.currency_var = tk.StringVar(value="NGN")

        ttk.Label(details_tab, text="Input Details", style="SectionHint.TLabel",
                  font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 10))
        project_card = ttk.LabelFrame(details_tab, text="Project Description", padding=14)
        project_card.pack(fill=tk.X, pady=(0, 12))
        for key, default in (
            ("Project title", "Chemical Process Plant TEA Study"),
            ("Project summary", "Preliminary techno-economic analysis of a modeled chemical process plant."),
        ):
            ttk.Label(project_card, text=key).pack(anchor="w", pady=(4, 2))
            entry = ttk.Entry(project_card, width=100)
            entry.insert(0, default)
            entry.pack(fill=tk.X, pady=(0, 5))
            self.input_entries[key] = entry
        detail_card = ttk.LabelFrame(details_tab, text="What to enter", padding=14)
        detail_card.pack(fill=tk.X, pady=(0, 12))
        detail_text = ttk.Label(
            detail_card, justify=tk.LEFT, wraplength=1000,
            text=("Economic Parameters: enter the local currency code, the exchange rate as local currency per USD, "
                  "wage, production assumptions, project life, depreciation period, construction/startup period, and process type.\n\n"
                  "Process Consumables: choose Continuous for a usage or make-up rate per operating hour, or Periodic replacement for a quantity per event and the operating hours between events. TOPCEE annualizes the cost and converts the local-currency amount to USD.\n\n"
                  "Operating Cost Parameters: enter each raw material's flow-rate unit and price basis. Choose mass or volume units; if flow and price use different bases, supply density in kg/L for conversion. "
                  "List each utility with its name, hourly use, unit, and local price per unit. Up to 10 entries are available "
                        "in each list. The utility names are suggestions you can edit or erase. A row with both quantity and price "
                        "at zero is ignored, even if its name is filled in.\n\n"
                  "TOPCEE converts local-currency values to USD for the TEA calculation. Equipment costs are already in USD.")
        )
        detail_text.pack(anchor="w")
        old_file_card = ttk.LabelFrame(details_tab, text="Existing input files", padding=14)
        old_file_card.pack(fill=tk.X)
        old_file_text = ttk.Label(
            old_file_card, justify=tk.LEFT, wraplength=1000,
            text=("Use Load Inputs to open a previously saved TOPCEE JSON input file. Older TOPCEE input files are "
                  "recognized and their existing values are mapped into the updated fields. After loading, review "
                  "the currency code, exchange rate, and material and utility rows before calculating.")
        )
        old_file_text.pack(anchor="w")

        ttk.Label(economic_tab,
                  text="Enter prices in your local currency. TOPCEE converts them to USD using the exchange rate below.",
                  style="SectionHint.TLabel").pack(anchor="w", pady=(0, 10))
        econ_card = ttk.LabelFrame(economic_tab, text="Currency and Economics", padding=12)
        econ_card.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(econ_card, text="Hourly labor wage defaults to 3,850 local currency/hour (77,000 ÷ 20).",
                  style="SectionHint.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=(2, 6))
        econ_fields = [
            ("Local currency code", "NGN"),
            ("Exchange rate (local currency per $)", ""),
            ("Operating labor cost (local/hour)", "3850"),
            ("Subsidy (%)", "0"), ("Tax (%)", "0"),
            ("Working hour in h/day", "24"),
            ("Number of shifts per day", "3"),
            ("Operating days per year", "330"),
            ("Selling price of your product (local/L)", "0"),
            ("Project life in years", "20"), ("Discount rate in %", "0"),
            ("Depreciation period in years", "10"),
            ("Construction/Startup period (production starts in year)", "3"),
        ]
        for index, (key, default) in enumerate(econ_fields, start=1):
            row, col = divmod(index, 2)
            ttk.Label(econ_card, text=key).grid(row=row, column=col * 2, sticky="w", padx=8, pady=6)
            entry = ttk.Entry(econ_card, width=22)
            entry.insert(0, default)
            entry.grid(row=row, column=col * 2 + 1, sticky="ew", padx=8, pady=6)
            self.input_entries[key] = entry
            if key == "Local currency code":
                entry.configure(textvariable=self.currency_var)
        for col in range(4):
            econ_card.columnconfigure(col, weight=1 if col % 2 else 0)

        process_card = ttk.LabelFrame(economic_tab, text="Production and Process", padding=12)
        process_card.pack(fill=tk.X)
        process_fields = [
            ("Quantity of product manufactured in L/h", "0"),
            ("Particulate processing steps (P)", "0"),
            ("Non-particulate processing steps (Nnp)", "0"),
            ("Process Type: F-Fluid, S-Solid, M-Mixed", "F"),
            ("Total equipment cost ($)", "0"),
        ]
        for index, (key, default) in enumerate(process_fields):
            row, col = divmod(index, 2)
            ttk.Label(process_card, text=key).grid(row=row, column=col * 2, sticky="w", padx=8, pady=6)
            entry = ttk.Entry(process_card, width=22)
            entry.insert(0, default)
            entry.grid(row=row, column=col * 2 + 1, sticky="ew", padx=8, pady=6)
            self.input_entries[key] = entry
        process_card.columnconfigure(1, weight=1)
        process_card.columnconfigure(3, weight=1)

        ttk.Label(materials_tab,
                  text="Enter up to 10 raw materials and 10 utilities. Unused rows can remain at zero.",
                  style="SectionHint.TLabel").pack(anchor="w", pady=(0, 8))
        material_canvas_frame = ttk.Frame(materials_tab)
        material_canvas_frame.pack(fill=tk.BOTH, expand=True)
        material_canvas = tk.Canvas(material_canvas_frame, highlightthickness=0, bg="#f3f6fa")
        material_vsb = ttk.Scrollbar(material_canvas_frame, orient="vertical", command=material_canvas.yview)
        material_canvas.configure(yscrollcommand=material_vsb.set)
        material_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        material_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        material_content = ttk.Frame(material_canvas)
        material_window = material_canvas.create_window((0, 0), window=material_content, anchor="nw")
        material_content.bind("<Configure>",
                              lambda _e: material_canvas.configure(scrollregion=material_canvas.bbox("all")))
        material_canvas.bind("<Configure>",
                             lambda e: material_canvas.itemconfigure(material_window, width=e.width))

        consumable_card = ttk.LabelFrame(material_content, text="Process Consumables", padding=10)
        consumable_card.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(
            consumable_card,
            text=("Continuous: enter a consumption or make-up rate per operating hour; the interval is ignored. "
                  "Periodic replacement: enter the quantity replaced per event and the operating hours between events. "
                  "Annualized cost = quantity × local unit price × (operating hours/day × operating days/year) ÷ replacement interval, "
                  "then converted to USD. For example, 100 kg replaced every 4,000 operating h at 500 local/kg, "
                  "with 24 h/day and 330 days/year, costs 99,000 local/year. Give a positive interval when both quantity and price are entered."),
            style="SectionHint.TLabel", wraplength=1150, justify=tk.LEFT,
        ).grid(row=0, column=0, columnspan=6, sticky="w", padx=4, pady=(0, 8))
        consumable_headers = (
            "Consumable", "Costing mode", "Quantity (per h / event)", "Unit",
            "Price (local/unit)", "Replacement interval (operating h; periodic only)",
        )
        for col, heading in enumerate(consumable_headers):
            ttk.Label(consumable_card, text=heading).grid(row=1, column=col, sticky="w", padx=4)
        for i in range(1, 11):
            defaults = ("Catalyst" if i == 1 else f"Consumable {i}", "Continuous", "0",
                        "kg" if i == 1 else "", "0", "0")
            keys = (
                f"Process consumable {i} name", f"Process consumable {i} mode",
                f"Process consumable {i} quantity", f"Process consumable {i} unit",
                f"Process consumable {i} cost (local/unit)",
                f"Process consumable {i} replacement interval (operating h)",
            )
            for col, (key, default) in enumerate(zip(keys, defaults)):
                if col == 1:
                    entry = ttk.Combobox(consumable_card, values=("Continuous", "Periodic replacement"),
                                         width=23, state="readonly")
                    entry.set(default)
                else:
                    entry = ttk.Entry(consumable_card, width=(25, 24, 18, 10, 17, 31)[col])
                    entry.insert(0, default)
                entry.grid(row=i + 1, column=col, sticky="ew", padx=4, pady=2)
                self.input_entries[key] = entry

        raw_card = ttk.LabelFrame(material_content, text="Raw Materials (select rate and price units; use density for mass/volume conversion)", padding=10)
        raw_card.pack(fill=tk.X, pady=(0, 10))
        raw_headers = ("Material", "Flow rate", "Rate unit", "Price (local currency)", "Price basis", "Density (kg/L; if needed)")
        for col, heading in enumerate(raw_headers):
            ttk.Label(raw_card, text=heading).grid(row=0, column=col, sticky="w", padx=4)
        for i in range(1, 11):
            name_key = f"Raw material {i} name"
            row_keys = (
                name_key, f"Raw material {i} quantity (unit/h)", f"Raw material {i} rate unit",
                f"Raw material {i} cost (local/unit)", f"Raw material {i} price basis",
                f"Raw material {i} density (kg/L)",
            )
            defaults = (f"Raw material {i}", "0", "kg/h", "0", "kg", "0")
            for col, (key, default) in enumerate(zip(row_keys, defaults)):
                if col == 2:
                    entry = ttk.Combobox(raw_card, values=("kg/h", "t/h", "L/h", "m³/h"),
                                         width=10, state="readonly")
                    entry.set(default)
                elif col == 4:
                    entry = ttk.Combobox(raw_card, values=("kg", "t", "L", "m³"),
                                         width=10, state="readonly")
                    entry.set(default)
                else:
                    entry = ttk.Entry(raw_card, width=(24, 15, 10, 15, 10, 18)[col])
                    entry.insert(0, default)
                entry.grid(row=i, column=col, sticky="ew", padx=4, pady=2)
                self.input_entries[key] = entry

        util_card = ttk.LabelFrame(material_content, text="Utilities (enter quantity and local-currency price per unit)", padding=10)
        util_card.pack(fill=tk.X, pady=(0, 10))
        for col, heading in enumerate(("Utility", "Quantity (unit/h)", "Unit", "Price")):
            ttk.Label(util_card, text=heading).grid(row=0, column=col, sticky="w", padx=5)
        self.utility_price_heading = tk.StringVar(value="Price (NGN/unit)")
        ttk.Label(util_card, textvariable=self.utility_price_heading).grid(row=0, column=3, sticky="w", padx=5)
        utility_defaults = [
            ("Electricity", "kWh"), ("Cooling water", "kg"),
            ("Process water", "kg"), ("Steam", "kg"),
            ("Compressed air", "Nm3"), ("Chilled water", "kg"),
            ("Boiler fuel", "kg"), ("Instrument air", "Nm3"),
            ("Nitrogen", "Nm3"), ("Other utility", ""),
        ]
        for i in range(1, 11):
            utility_name, utility_unit = utility_defaults[i - 1]
            self._add_material_row(util_card, i,
                f"Utility {i} name", f"Utility {i} quantity (unit/h)",
                f"Utility {i} cost (local/unit)", utility_name, "0", "0",
                unit_key=f"Utility {i} unit", unit_default=utility_unit)
        self._bind_scroll_wheel(material_content, material_canvas)

        self.currency_var.trace_add("write", self._refresh_currency_labels)
        before_card = ttk.LabelFrame(self.tab_inputs, text="Before Analysis", padding=8)
        before_card.pack(fill=tk.X, padx=12, pady=(0, 4))
        ttk.Label(before_card,
                  text="When you start a calculation, TOPCEE will ask whether to save all current input data first.",
                  style="SectionHint.TLabel").pack(anchor="w")
        actions = ttk.Frame(self.tab_inputs, padding=(12, 2, 12, 10))
        actions.pack(fill=tk.X)
        ttk.Button(actions, text="Save Inputs", command=self._save_inputs).pack(side=tk.LEFT, padx=4)
        ttk.Button(actions, text="Load Inputs", command=self._load_inputs).pack(side=tk.LEFT, padx=4)
        ttk.Button(actions, text="Calculate Profitability", command=self._run_calculation,
                   style="Accent.TButton").pack(side=tk.RIGHT, padx=4)
        ttk.Button(actions, text="Go to Results", command=self._go_results).pack(side=tk.RIGHT, padx=4)

    def _add_material_row(self, parent, row, name_key, qty_key, price_key,
                          name_default, qty_default, price_default, unit_key=None,
                          unit_default=""):
        name = ttk.Entry(parent, width=25)
        name.insert(0, name_default)
        name.grid(row=row, column=0, sticky="ew", padx=5, pady=2)
        quantity = ttk.Entry(parent, width=16)
        quantity.insert(0, qty_default)
        quantity.grid(row=row, column=1, sticky="ew", padx=5, pady=2)
        self.input_entries[name_key] = name
        self.input_entries[qty_key] = quantity
        if unit_key:
            unit = ttk.Entry(parent, width=12)
            unit.insert(0, unit_default)
            unit.grid(row=row, column=2, sticky="ew", padx=5, pady=2)
            self.input_entries[unit_key] = unit
            price_col = 3
        else:
            price_col = 2
        price = ttk.Entry(parent, width=17)
        price.insert(0, price_default)
        price.grid(row=row, column=price_col, sticky="ew", padx=5, pady=2)
        self.input_entries[price_key] = price

    def _bind_scroll_wheel(self, content, canvas):
        """Scroll only the canvas whose content the pointer is currently over."""
        def scroll(event):
            if getattr(event, "num", None) == 4:
                canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                canvas.yview_scroll(1, "units")
            else:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"

        def bind_tree(widget):
            widget.bind("<MouseWheel>", scroll, add="+")
            widget.bind("<Button-4>", scroll, add="+")
            widget.bind("<Button-5>", scroll, add="+")
            for child in widget.winfo_children():
                bind_tree(child)

        bind_tree(content)

    def _refresh_currency_labels(self, *_args):
        code = self.currency_var.get().strip().upper() or "LOCAL"
        self.utility_price_heading.set(f"Price ({code}/unit)")

    # ---------------------------------------------------- Results Tab ------
    def _build_results_tab(self):
        frame = self.tab_results

        # Top: calculate button + result text
        top = tk.Frame(frame)
        top.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)
        tk.Button(top, text="Calculate Profitability", command=self._run_calculation,
                  bg="#673AB7", fg="white", font=("TkDefaultFont", 10, "bold")).pack(side=tk.LEFT, padx=4)
        tk.Button(top, text="Save Results (CSV)", command=self._save_results_csv).pack(side=tk.LEFT, padx=4)
        tk.Button(top, text="Go to Report",
                  command=lambda: self.notebook.select(self.tab_report)).pack(side=tk.LEFT, padx=4)
        tk.Button(top, text="Sensitivity Analysis",
                  command=lambda: self.notebook.select(self.tab_sensitivity)).pack(side=tk.LEFT, padx=4)

        # Keep the numeric results and chart on separate pages so neither crowds out the other.
        result_tabs = ttk.Notebook(frame)
        result_tabs.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
        summary_page = ttk.Frame(result_tabs)
        chart_page = ttk.Frame(result_tabs)
        result_tabs.add(summary_page, text="Analysis Results")
        result_tabs.add(chart_page, text="Cash Flow Diagram")

        # Scrollable results display
        canvas = tk.Canvas(summary_page, highlightthickness=0, bg="#f3f6fa")
        vsb = ttk.Scrollbar(summary_page, orient="vertical", command=canvas.yview)
        self.res_frame = tk.Frame(canvas)
        self.res_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        result_window = canvas.create_window((0, 0), window=self.res_frame, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(result_window, width=e.width))
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_canvas = canvas

        # Chart has its own page, with enough space to read axes and values.
        self.chart_frame = tk.LabelFrame(chart_page, text="Cash Flow Diagram")
        self.chart_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # Initial message
        tk.Label(self.res_frame, text="Run calculation to see CAPEX, OPEX, and profitability results.",
                 fg="#666", pady=20).pack()
        self._bind_scroll_wheel(self.res_frame, canvas)

    # ---------------------------------------------------- Report Tab -------
    def _build_report_tab(self):
        frame = self.tab_report

        # Folder selection (remembered)
        fold = tk.LabelFrame(frame, text="Output Folder")
        fold.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)
        inner = tk.Frame(fold)
        inner.pack(fill=tk.X, padx=8, pady=6)
        tk.Label(inner, text="Save reports to:").pack(side=tk.LEFT)
        self.report_folder_var = tk.StringVar(value=self.last_folder)
        tk.Entry(inner, textvariable=self.report_folder_var, width=70).pack(side=tk.LEFT, padx=6, fill=tk.X, expand=True)
        tk.Button(inner, text="Browse", command=self._browse_report_folder).pack(side=tk.LEFT, padx=4)
        tk.Button(inner, text="Remember this folder", command=self._remember_folder).pack(side=tk.LEFT, padx=4)

        # Report options
        opt = tk.LabelFrame(frame, text="Reports to Generate")
        opt.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)
        self.gen_equip_var = tk.BooleanVar(value=True)
        self.gen_full_report_var = tk.BooleanVar(value=True)
        chk = tk.Frame(opt)
        chk.pack(padx=8, pady=6)
        tk.Checkbutton(chk, text="Equipment Cost Schedule (CSV)",
                       variable=self.gen_equip_var).pack(side=tk.LEFT, padx=10)
        tk.Checkbutton(chk,
                       text="Full TEA Report (CAPEX, OPEX, Profitability, and Cash Flow)",
                       variable=self.gen_full_report_var).pack(side=tk.LEFT, padx=10)

        # Generate buttons
        gen = tk.Frame(frame)
        gen.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)
        tk.Button(gen, text="Generate All Selected Reports",
                  bg="#4CAF50", fg="white", font=("TkDefaultFont", 11, "bold"),
                  command=self._generate_reports).pack(side=tk.LEFT, padx=6)
        tk.Button(gen, text="Open Output Folder",
                  command=self._open_output_folder).pack(side=tk.LEFT, padx=6)

        # Status/log
        log = tk.LabelFrame(frame, text="Report Log")
        log.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=6)
        self.report_log = tk.Text(log, height=10, wrap=tk.WORD)
        self.report_log.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        tk.Label(log, text="Full analysis reports are saved as HTML, PDF, CSV, and Excel (if openpyxl is installed).",
                 fg="#666").pack(anchor=tk.W, padx=6)

    # ----------------------------------------- Sensitivity Analysis Tab ----
    def _build_sensitivity_tab(self):
        frame = self.tab_sensitivity
        intro = ttk.LabelFrame(frame, text="Sensitivity Study", padding=12)
        intro.pack(fill=tk.X, padx=10, pady=(10, 6))
        ttk.Label(
            intro,
            text=("Select one numeric input from Economic Parameters, or an entered or custom-named "
                  "Operating Cost Parameters row (process consumables, raw materials, or utilities), "
                  "and choose one, two, or three profitability response metrics. "
                  "TOPCEE will calculate 21 cases from -100% to +100% of the baseline, "
                  "in 10% increments, without changing the Process Inputs. Integer inputs are "
                  "rounded to whole values; invalid cases or undefined metrics are marked N/A. "
                  "A zero baseline remains zero throughout a percentage sweep."),
            style="SectionHint.TLabel", wraplength=1100, justify=tk.LEFT,
        ).pack(anchor="w")

        controls = ttk.LabelFrame(frame, text="Study Setup", padding=12)
        controls.pack(fill=tk.X, padx=10, pady=6)
        ttk.Label(controls, text="Input parameter").grid(row=0, column=0, sticky="w", padx=6, pady=5)
        self.sensitivity_parameter_var = tk.StringVar()
        self.sensitivity_parameter_combo = ttk.Combobox(
            controls, textvariable=self.sensitivity_parameter_var,
            state="readonly", width=48,
        )
        self.sensitivity_parameter_combo.grid(row=0, column=1, sticky="ew", padx=6, pady=5)
        self.sensitivity_parameter_combo.bind(
            "<<ComboboxSelected>>", lambda _event: self._update_sensitivity_baseline())
        self.sensitivity_baseline_var = tk.StringVar(value="Baseline: -")
        ttk.Label(controls, textvariable=self.sensitivity_baseline_var,
                  style="SectionHint.TLabel").grid(row=0, column=2, sticky="w", padx=6, pady=5)
        ttk.Button(controls, text="Refresh from Process Inputs",
                   command=self._refresh_sensitivity_parameters).grid(
                       row=0, column=3, padx=6, pady=5)

        ttk.Label(controls, text="Number of response variables").grid(
            row=1, column=0, sticky="w", padx=6, pady=5)
        self.sensitivity_metric_count_var = tk.StringVar(value="3")
        metric_count_combo = ttk.Combobox(
            controls, textvariable=self.sensitivity_metric_count_var,
            values=("1", "2", "3"), state="readonly", width=8,
        )
        metric_count_combo.grid(row=1, column=1, sticky="w", padx=6, pady=5)
        metric_count_combo.bind(
            "<<ComboboxSelected>>", lambda _event: self._update_sensitivity_metric_controls())

        metric_labels = (
            "Net Present Worth (Non-Discounted) — NNPW ($)",
            "Net Present Worth (Discounted) — DNPW ($)",
            "Internal Rate of Return — IRR (%)",
            "Simple Annual Return on Investment — ROI (%)",
            "Undiscounted Payback — NPBP (years)",
            "Discounted Payback — DPBP (years)",
            "Undiscounted Benefit-Cost Ratio",
            "Discounted Benefit-Cost Ratio",
        )
        self.sensitivity_metric_keys = {
            metric_labels[0]: "NNPW", metric_labels[1]: "DNPW",
            metric_labels[2]: "IRR", metric_labels[3]: "ROI",
            metric_labels[4]: "NPBP", metric_labels[5]: "DPBP",
            metric_labels[6]: "B_C_ratio_NCF", metric_labels[7]: "B_C_ratio_DCF",
        }
        defaults = (metric_labels[0], metric_labels[2], metric_labels[3])
        self.sensitivity_metric_vars = []
        self.sensitivity_metric_row_widgets = []
        for index, default in enumerate(defaults):
            metric_label_widget = ttk.Label(controls, text=f"Response metric {index + 1}")
            metric_label_widget.grid(row=index + 2, column=0, sticky="w", padx=6, pady=5)
            metric_var = tk.StringVar(value=default)
            metric_combo = ttk.Combobox(
                controls, textvariable=metric_var, values=metric_labels,
                state="readonly", width=66,
            )
            metric_combo.grid(row=index + 2, column=1, columnspan=3, sticky="ew", padx=6, pady=5)
            self.sensitivity_metric_vars.append(metric_var)
            self.sensitivity_metric_row_widgets.append((metric_label_widget, metric_combo))

        controls.columnconfigure(1, weight=1)
        self._update_sensitivity_metric_controls()

        output = ttk.LabelFrame(frame, text="Sensitivity Report Output", padding=10)
        output.pack(fill=tk.X, padx=10, pady=6)
        self.sensitivity_folder_var = tk.StringVar(value=self.last_folder)
        ttk.Label(output, text="Save reports to").pack(side=tk.LEFT, padx=(4, 6))
        ttk.Entry(output, textvariable=self.sensitivity_folder_var, width=76).pack(
            side=tk.LEFT, padx=4, fill=tk.X, expand=True)
        ttk.Button(output, text="Browse", command=self._browse_sensitivity_folder).pack(
            side=tk.LEFT, padx=4)
        ttk.Button(output, text="Run and Export CSV, HTML, Excel, and Figure",
                   command=self._run_sensitivity_analysis,
                   style="Accent.TButton").pack(side=tk.LEFT, padx=6)

        result_tabs = ttk.Notebook(frame)
        result_tabs.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        table_page = ttk.Frame(result_tabs)
        figure_page = ttk.Frame(result_tabs)
        result_tabs.add(table_page, text="Results Table")
        result_tabs.add(figure_page, text="Sensitivity Figure")

        table_frame = ttk.LabelFrame(table_page, text="Sensitivity Results", padding=8)
        table_frame.pack(fill=tk.BOTH, expand=True)
        columns = (
            "change", "parameter_value", "metric1_value", "metric1_change",
            "metric2_value", "metric2_change", "metric3_value", "metric3_change",
        )
        self.sensitivity_tree = ttk.Treeview(
            table_frame, columns=columns, show="headings", height=15)
        for column, heading, width in (
            ("change", "Change from baseline (%)", 165),
            ("parameter_value", "Input value", 165),
            ("metric1_value", "Response 1 value", 175),
            ("metric1_change", "Response 1 change (%)", 175),
            ("metric2_value", "Response 2 value", 175),
            ("metric2_change", "Response 2 change (%)", 175),
            ("metric3_value", "Response 3 value", 175),
            ("metric3_change", "Response 3 change (%)", 175),
        ):
            self.sensitivity_tree.heading(column, text=heading)
            self.sensitivity_tree.column(column, width=width, anchor="e" if column != "change" else "center")
        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.sensitivity_tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.sensitivity_tree.xview)
        self.sensitivity_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.sensitivity_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        ttk.Label(
            figure_page,
            text=("X-axis: selected input change from baseline (%). Y-axis: each response metric's "
                  "change from its base-case value (% of the absolute base value)."),
            style="SectionHint.TLabel", wraplength=1100,
        ).pack(anchor="w", padx=10, pady=(8, 0))
        # Standard landscape 5:3 figure keeps the -100%..+100% curves legible
        # while fitting the app pane and printable report layouts.
        self.sensitivity_figure = Figure(figsize=(8, 4.8), dpi=100)
        self.sensitivity_ax = self.sensitivity_figure.add_subplot(111)
        self.sensitivity_canvas = FigureCanvasTkAgg(self.sensitivity_figure, master=figure_page)
        self.sensitivity_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.sensitivity_ax.text(
            0.5, 0.5, "Run the sensitivity study to display response curves.",
            ha="center", va="center", transform=self.sensitivity_ax.transAxes,
        )
        self.sensitivity_ax.set_axis_off()
        self.sensitivity_canvas.draw()

        self.sensitivity_status_var = tk.StringVar(
            value="Operating-cost options reflect current row names and units; unused default rows are hidden.")
        ttk.Label(frame, textvariable=self.sensitivity_status_var,
                  style="SectionHint.TLabel", wraplength=1100).pack(
                      anchor="w", padx=16, pady=(0, 8))
        self._refresh_sensitivity_parameters()

    def _refresh_sensitivity_parameters(self):
        inputs = self._get_inputs()
        current_label = self.sensitivity_parameter_var.get()
        current_key = getattr(self, "sensitivity_parameter_key_by_label", {}).get(current_label)
        currency = inputs.get("Local currency code", "local currency").strip() or "local currency"
        numeric_keys = {
            "Exchange rate (local currency per $)", "Operating labor cost (local/hour)",
            "Subsidy (%)", "Tax (%)", "Working hour in h/day",
            "Number of shifts per day", "Operating days per year",
            "Selling price of your product (local/L)", "Project life in years",
            "Discount rate in %", "Depreciation period in years",
            "Construction/Startup period (production starts in year)",
            "Quantity of product manufactured in L/h", "Particulate processing steps (P)",
            "Non-particulate processing steps (Nnp)", "Total equipment cost ($)",
        }
        def input_number(key):
            try:
                return float(str(inputs.get(key, "")).replace(",", "").strip() or 0)
            except (TypeError, ValueError):
                return 0.0

        utility_defaults = (
            "Electricity", "Cooling water", "Process water", "Steam", "Compressed air",
            "Chilled water", "Boiler fuel", "Instrument air", "Nitrogen", "Other utility",
        )
        for index in range(1, 11):
            consumable_name = inputs.get(f"Process consumable {index} name", "").strip()
            consumable_default = "Catalyst" if index == 1 else f"Consumable {index}"
            consumable_fields = {
                f"Process consumable {index} quantity",
                f"Process consumable {index} cost (local/unit)",
                f"Process consumable {index} replacement interval (operating h)",
            }
            if (consumable_name != consumable_default or
                    any(input_number(key) != 0 for key in consumable_fields)):
                numeric_keys.update(consumable_fields)

            raw_name = inputs.get(f"Raw material {index} name", "").strip()
            raw_default = f"Raw material {index}"
            raw_fields = {
                f"Raw material {index} quantity (unit/h)",
                f"Raw material {index} cost (local/unit)",
                f"Raw material {index} density (kg/L)",
            }
            if raw_name != raw_default or any(input_number(key) != 0 for key in raw_fields):
                numeric_keys.update(raw_fields)

            utility_name = inputs.get(f"Utility {index} name", "").strip()
            utility_fields = {
                f"Utility {index} quantity (unit/h)",
                f"Utility {index} cost (local/unit)",
            }
            if (utility_name != utility_defaults[index - 1] or
                    any(input_number(key) != 0 for key in utility_fields)):
                numeric_keys.update(utility_fields)

        self.sensitivity_parameter_key_by_label = {}
        for key in numeric_keys:
            value = inputs.get(key, "")
            try:
                parsed_value = float(str(value).replace(",", "").strip() or 0)
            except (TypeError, ValueError):
                continue
            if not math.isfinite(parsed_value):
                continue
            if key.startswith("Process consumable "):
                index = key.split()[2]
                item_name = inputs.get(f"Process consumable {index} name", "").strip()
                item_name = item_name or f"Consumable {index}"
                unit = inputs.get(f"Process consumable {index} unit", "").strip() or "unit"
                mode = inputs.get(f"Process consumable {index} mode", "Continuous").strip().lower()
                if key.endswith(" quantity"):
                    detail = (f"Replacement quantity ({unit}/event)" if "periodic" in mode
                              else f"Usage rate ({unit}/h)")
                elif key.endswith(" replacement interval (operating h)"):
                    detail = "Replacement interval (operating h)"
                else:
                    detail = f"Unit price ({currency}/{unit})"
                label = (f"Operating Cost Parameters | Process Consumable {index}: "
                         f"{item_name} | {detail}")
            elif key.startswith("Raw material "):
                index = key.split()[2]
                item_name = inputs.get(f"Raw material {index} name", "").strip()
                item_name = item_name or f"Raw material {index}"
                rate_unit = inputs.get(f"Raw material {index} rate unit", "kg/h").strip() or "kg/h"
                price_basis = inputs.get(f"Raw material {index} price basis", "kg").strip() or "kg"
                if key.endswith(" quantity (unit/h)"):
                    detail = f"Flow rate ({rate_unit})"
                elif key.endswith(" density (kg/L)"):
                    detail = "Density (kg/L)"
                else:
                    detail = f"Unit price ({currency}/{price_basis})"
                label = (f"Operating Cost Parameters | Raw Material {index}: "
                         f"{item_name} | {detail}")
            elif key.startswith("Utility "):
                index = key.split()[1]
                item_name = inputs.get(f"Utility {index} name", "").strip()
                item_name = item_name or f"Utility {index}"
                unit = inputs.get(f"Utility {index} unit", "").strip() or "unit"
                detail = (f"Consumption rate ({unit}/h)" if "quantity" in key
                          else f"Unit price ({currency}/{unit})")
                label = (f"Operating Cost Parameters | Utility {index}: "
                         f"{item_name} | {detail}")
            else:
                label = f"Economic Parameters | {key}"
            self.sensitivity_parameter_key_by_label[label] = key
        candidates = sorted(self.sensitivity_parameter_key_by_label)
        self.sensitivity_parameter_combo["values"] = candidates
        selected_label = next(
            (label for label, key in self.sensitivity_parameter_key_by_label.items()
             if key == current_key),
            current_label if current_label in candidates else (candidates[0] if candidates else ""),
        )
        self.sensitivity_parameter_var.set(selected_label)
        self._update_sensitivity_baseline()
        if not candidates:
            self.sensitivity_status_var.set(
                "No numeric Process Inputs are available to sweep.")
        else:
            self.sensitivity_status_var.set(
                f"{len(candidates)} numeric inputs available from Economic Parameters and entered/custom-named operating-cost rows. "
                "Unused default rows are hidden; zero-valued active inputs remain selectable. "
                "Percentage sweep: -100%, -90%, …, 0%, …, +90%, +100%.")

    def _update_sensitivity_metric_controls(self):
        try:
            metric_count = int(self.sensitivity_metric_count_var.get())
        except (TypeError, ValueError):
            metric_count = 3
        metric_count = min(max(metric_count, 1), 3)
        self.sensitivity_metric_count_var.set(str(metric_count))
        for index, (label_widget, metric_combo) in enumerate(self.sensitivity_metric_row_widgets):
            if index < metric_count:
                label_widget.grid()
                metric_combo.grid()
            else:
                label_widget.grid_remove()
                metric_combo.grid_remove()

    def _update_sensitivity_baseline(self):
        label = self.sensitivity_parameter_var.get()
        key = self.sensitivity_parameter_key_by_label.get(label, "")
        value = self._get_inputs().get(key, "") if key else ""
        self.sensitivity_baseline_var.set(f"Baseline: {value or '0 (blank)'}")

    def _on_notebook_tab_changed(self, _event=None):
        if self.notebook.select() == str(self.tab_sensitivity):
            self._refresh_sensitivity_parameters()

    def _browse_sensitivity_folder(self):
        folder = filedialog.askdirectory(
            title="Select Sensitivity Report Folder",
            initialdir=self.sensitivity_folder_var.get() or self.last_folder,
        )
        if folder:
            self.sensitivity_folder_var.set(folder)

    @staticmethod
    def _sensitivity_display_value(value):
        if value is None:
            return "N/A"
        try:
            numeric = float(value)
        except (TypeError, ValueError):
            return str(value)
        return f"{numeric:,.6g}" if math.isfinite(numeric) else "N/A"

    @staticmethod
    def _write_sensitivity_xlsx(path, rows, chart_png_bytes):
        """Write a formatted XLSX workbook with the sensitivity figure embedded."""
        main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        pkg_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        content_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        drawingml_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        ET.register_namespace("", main_ns)
        ET.register_namespace("r", rel_ns)
        ET.register_namespace("xdr", xdr_ns)
        ET.register_namespace("a", drawingml_ns)

        def qn(namespace, tag):
            return f"{{{namespace}}}{tag}"

        def column_name(number):
            name = ""
            while number:
                number, remainder = divmod(number - 1, 26)
                name = chr(65 + remainder) + name
            return name

        worksheet = ET.Element(qn(main_ns, "worksheet"))
        ET.SubElement(worksheet, qn(main_ns, "sheetViews"))
        views = worksheet.find(qn(main_ns, "sheetViews"))
        view = ET.SubElement(views, qn(main_ns, "sheetView"), {"workbookViewId": "0"})
        ET.SubElement(view, qn(main_ns, "pane"), {
            "ySplit": "12", "topLeftCell": "A13", "activePane": "bottomLeft", "state": "frozen"
        })
        cols = ET.SubElement(worksheet, qn(main_ns, "cols"))
        column_widths = (24, 42, 30, 34, 30, 34, 30, 34)
        column_count = max(len(row) for row in rows)
        for index, width in enumerate(column_widths[:column_count], start=1):
            ET.SubElement(cols, qn(main_ns, "col"), {
                "min": str(index), "max": str(index), "width": str(width), "customWidth": "1"
            })
        sheet_data = ET.SubElement(worksheet, qn(main_ns, "sheetData"))
        for row_index, values in enumerate(rows, start=1):
            row_element = ET.SubElement(sheet_data, qn(main_ns, "row"), {"r": str(row_index)})
            style_id = 2 if row_index == 1 else (1 if row_index == 12 else 0)
            for col_index, value in enumerate(values, start=1):
                if value is None:
                    value = "N/A"
                cell = ET.SubElement(row_element, qn(main_ns, "c"), {
                    "r": f"{column_name(col_index)}{row_index}", "s": str(style_id)
                })
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    ET.SubElement(cell, qn(main_ns, "v")).text = str(value)
                else:
                    cell.set("t", "inlineStr")
                    inline = ET.SubElement(cell, qn(main_ns, "is"))
                    text_node = ET.SubElement(inline, qn(main_ns, "t"))
                    text_node.text = str(value)
        ET.SubElement(worksheet, qn(main_ns, "autoFilter"), {
            "ref": f"A12:{column_name(max(len(row) for row in rows))}{len(rows)}"
        })
        ET.SubElement(worksheet, qn(main_ns, "drawing"), {qn(rel_ns, "id"): "rId1"})

        sheet_rels = ET.Element(qn(pkg_rel_ns, "Relationships"))
        ET.SubElement(sheet_rels, qn(pkg_rel_ns, "Relationship"), {
            "Id": "rId1", "Type": f"{rel_ns}/drawing", "Target": "../drawings/drawing1.xml"
        })

        drawing = ET.Element(qn(xdr_ns, "wsDr"))
        anchor = ET.SubElement(drawing, qn(xdr_ns, "oneCellAnchor"))
        anchor_from = ET.SubElement(anchor, qn(xdr_ns, "from"))
        for tag, value in zip(("col", "row", "colOff", "rowOff"), (0, 35, 0, 0)):
            ET.SubElement(anchor_from, qn(xdr_ns, tag)).text = str(value)
        ET.SubElement(anchor, qn(xdr_ns, "ext"), {"cx": "9144000", "cy": "5486400"})
        picture = ET.SubElement(anchor, qn(xdr_ns, "pic"))
        nonvisual = ET.SubElement(picture, qn(xdr_ns, "nvPicPr"))
        ET.SubElement(nonvisual, qn(xdr_ns, "cNvPr"), {
            "id": "1", "name": "TOPCEE Sensitivity Figure"
        })
        nonvisual_picture = ET.SubElement(nonvisual, qn(xdr_ns, "cNvPicPr"))
        ET.SubElement(nonvisual_picture, qn(drawingml_ns, "picLocks"), {"noChangeAspect": "1"})
        blip_fill = ET.SubElement(picture, qn(xdr_ns, "blipFill"))
        ET.SubElement(blip_fill, qn(drawingml_ns, "blip"), {qn(rel_ns, "embed"): "rId1"})
        stretch = ET.SubElement(blip_fill, qn(drawingml_ns, "stretch"))
        ET.SubElement(stretch, qn(drawingml_ns, "fillRect"))
        shape_props = ET.SubElement(picture, qn(xdr_ns, "spPr"))
        transform = ET.SubElement(shape_props, qn(drawingml_ns, "xfrm"))
        ET.SubElement(transform, qn(drawingml_ns, "off"), {"x": "0", "y": "0"})
        ET.SubElement(transform, qn(drawingml_ns, "ext"), {"cx": "9144000", "cy": "5486400"})
        geometry = ET.SubElement(shape_props, qn(drawingml_ns, "prstGeom"), {"prst": "rect"})
        ET.SubElement(geometry, qn(drawingml_ns, "avLst"))
        ET.SubElement(anchor, qn(xdr_ns, "clientData"))

        drawing_rels = ET.Element(qn(pkg_rel_ns, "Relationships"))
        ET.SubElement(drawing_rels, qn(pkg_rel_ns, "Relationship"), {
            "Id": "rId1", "Type": f"{rel_ns}/image", "Target": "../media/image1.png"
        })

        workbook = ET.Element(qn(main_ns, "workbook"))
        sheets = ET.SubElement(workbook, qn(main_ns, "sheets"))
        ET.SubElement(sheets, qn(main_ns, "sheet"), {
            "name": "Sensitivity", "sheetId": "1", qn(rel_ns, "id"): "rId1"
        })

        workbook_rels = ET.Element(qn(pkg_rel_ns, "Relationships"))
        ET.SubElement(workbook_rels, qn(pkg_rel_ns, "Relationship"), {
            "Id": "rId1", "Type": f"{rel_ns}/worksheet", "Target": "worksheets/sheet1.xml"
        })
        ET.SubElement(workbook_rels, qn(pkg_rel_ns, "Relationship"), {
            "Id": "rId2", "Type": f"{rel_ns}/styles", "Target": "styles.xml"
        })
        root_rels = ET.Element(qn(pkg_rel_ns, "Relationships"))
        ET.SubElement(root_rels, qn(pkg_rel_ns, "Relationship"), {
            "Id": "rId1", "Type": f"{rel_ns}/officeDocument", "Target": "xl/workbook.xml"
        })
        content_types = ET.Element(qn(content_ns, "Types"))
        ET.SubElement(content_types, qn(content_ns, "Default"), {
            "Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"
        })
        ET.SubElement(content_types, qn(content_ns, "Default"), {
            "Extension": "xml", "ContentType": "application/xml"
        })
        ET.SubElement(content_types, qn(content_ns, "Default"), {
            "Extension": "png", "ContentType": "image/png"
        })
        ET.SubElement(content_types, qn(content_ns, "Override"), {
            "PartName": "/xl/workbook.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        })
        ET.SubElement(content_types, qn(content_ns, "Override"), {
            "PartName": "/xl/worksheets/sheet1.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
        })
        ET.SubElement(content_types, qn(content_ns, "Override"), {
            "PartName": "/xl/styles.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
        })
        ET.SubElement(content_types, qn(content_ns, "Override"), {
            "PartName": "/xl/drawings/drawing1.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.drawing+xml"
        })

        styles = ET.Element(qn(main_ns, "styleSheet"))
        fonts = ET.SubElement(styles, qn(main_ns, "fonts"), {"count": "3"})
        normal_font = ET.SubElement(fonts, qn(main_ns, "font"))
        ET.SubElement(normal_font, qn(main_ns, "sz"), {"val": "11"})
        ET.SubElement(normal_font, qn(main_ns, "name"), {"val": "Calibri"})
        header_font = ET.SubElement(fonts, qn(main_ns, "font"))
        ET.SubElement(header_font, qn(main_ns, "b"))
        ET.SubElement(header_font, qn(main_ns, "color"), {"rgb": "FF1A3C5A"})
        ET.SubElement(header_font, qn(main_ns, "sz"), {"val": "11"})
        ET.SubElement(header_font, qn(main_ns, "name"), {"val": "Calibri"})
        title_font = ET.SubElement(fonts, qn(main_ns, "font"))
        ET.SubElement(title_font, qn(main_ns, "b"))
        ET.SubElement(title_font, qn(main_ns, "color"), {"rgb": "FF164E78"})
        ET.SubElement(title_font, qn(main_ns, "sz"), {"val": "16"})
        ET.SubElement(title_font, qn(main_ns, "name"), {"val": "Calibri"})
        fills = ET.SubElement(styles, qn(main_ns, "fills"), {"count": "3"})
        ET.SubElement(fills, qn(main_ns, "fill")).append(ET.Element(qn(main_ns, "patternFill"), {"patternType": "none"}))
        ET.SubElement(fills, qn(main_ns, "fill")).append(ET.Element(qn(main_ns, "patternFill"), {"patternType": "gray125"}))
        header_fill = ET.SubElement(fills, qn(main_ns, "fill"))
        ET.SubElement(header_fill, qn(main_ns, "patternFill"), {"patternType": "solid"}).append(
            ET.Element(qn(main_ns, "fgColor"), {"rgb": "FFEEF2F7"}))
        borders = ET.SubElement(styles, qn(main_ns, "borders"), {"count": "1"})
        border = ET.SubElement(borders, qn(main_ns, "border"))
        for edge in ("left", "right", "top", "bottom", "diagonal"):
            ET.SubElement(border, qn(main_ns, edge))
        cell_style_xfs = ET.SubElement(styles, qn(main_ns, "cellStyleXfs"), {"count": "1"})
        ET.SubElement(cell_style_xfs, qn(main_ns, "xf"), {
            "numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0"
        })
        cell_xfs = ET.SubElement(styles, qn(main_ns, "cellXfs"), {"count": "3"})
        ET.SubElement(cell_xfs, qn(main_ns, "xf"), {
            "numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0", "xfId": "0"
        })
        header_xf = ET.SubElement(cell_xfs, qn(main_ns, "xf"), {
            "numFmtId": "0", "fontId": "1", "fillId": "2", "borderId": "0", "xfId": "0",
            "applyFont": "1", "applyFill": "1", "applyAlignment": "1"
        })
        ET.SubElement(header_xf, qn(main_ns, "alignment"), {"wrapText": "1", "horizontal": "center"})
        ET.SubElement(cell_xfs, qn(main_ns, "xf"), {
            "numFmtId": "0", "fontId": "2", "fillId": "0", "borderId": "0", "xfId": "0", "applyFont": "1"
        })
        cell_styles = ET.SubElement(styles, qn(main_ns, "cellStyles"), {"count": "1"})
        ET.SubElement(cell_styles, qn(main_ns, "cellStyle"), {
            "name": "Normal", "xfId": "0", "builtinId": "0"
        })

        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", ET.tostring(content_types, encoding="utf-8", xml_declaration=True))
            archive.writestr("_rels/.rels", ET.tostring(root_rels, encoding="utf-8", xml_declaration=True))
            archive.writestr("xl/workbook.xml", ET.tostring(workbook, encoding="utf-8", xml_declaration=True))
            archive.writestr("xl/_rels/workbook.xml.rels", ET.tostring(workbook_rels, encoding="utf-8", xml_declaration=True))
            archive.writestr("xl/worksheets/sheet1.xml", ET.tostring(worksheet, encoding="utf-8", xml_declaration=True))
            archive.writestr("xl/worksheets/_rels/sheet1.xml.rels", ET.tostring(sheet_rels, encoding="utf-8", xml_declaration=True))
            archive.writestr("xl/drawings/drawing1.xml", ET.tostring(drawing, encoding="utf-8", xml_declaration=True))
            archive.writestr("xl/drawings/_rels/drawing1.xml.rels", ET.tostring(drawing_rels, encoding="utf-8", xml_declaration=True))
            archive.writestr("xl/media/image1.png", chart_png_bytes)
            archive.writestr("xl/styles.xml", ET.tostring(styles, encoding="utf-8", xml_declaration=True))

    def _run_sensitivity_analysis(self):
        self._refresh_sensitivity_parameters()
        parameter_label = self.sensitivity_parameter_var.get()
        parameter = self.sensitivity_parameter_key_by_label.get(parameter_label, "")
        base_inputs = self._get_inputs()
        if not parameter or parameter not in base_inputs:
            messagebox.showerror("Sensitivity Analysis", "Select an input parameter first.")
            return
        try:
            baseline = float(str(base_inputs[parameter]).replace(",", "").strip() or 0)
        except (TypeError, ValueError):
            messagebox.showerror("Sensitivity Analysis", "The selected input must be numeric.")
            return
        if not math.isfinite(baseline):
            messagebox.showerror(
                "Sensitivity Analysis",
                "The selected input must have a finite numeric baseline.")
            return

        metric_count = int(self.sensitivity_metric_count_var.get())
        selected_labels = [variable.get() for variable in self.sensitivity_metric_vars[:metric_count]]
        if len(set(selected_labels)) != metric_count:
            messagebox.showerror(
                "Sensitivity Analysis",
                f"Select {metric_count} different response metric(s).")
            return
        metric_keys = [self.sensitivity_metric_keys[label] for label in selected_labels]

        folder = self.sensitivity_folder_var.get().strip() or self.last_folder
        if not os.path.isdir(folder):
            messagebox.showerror("Sensitivity Analysis", "The selected output folder does not exist.")
            return

        # Recalculate the base case from the current Process Inputs. The user's
        # entries remain untouched; each perturbation is applied to a copy.
        base_result = calculate_profitability(dict(base_inputs), show_errors=False)
        if base_result is None:
            messagebox.showerror(
                "Sensitivity Analysis",
                "The current base-case inputs could not be calculated. Correct the inputs and try again.")
            return

        base_metric_values = [base_result.get(key) for key in metric_keys]

        def response_change_percent(value, base_value):
            if value is None or base_value is None:
                return None
            try:
                value = float(value)
                base_value = float(base_value)
            except (TypeError, ValueError):
                return None
            if not math.isfinite(value) or not math.isfinite(base_value) or base_value == 0:
                return None
            # Absolute baseline preserves intuitive increase/decrease direction
            # when the base profitability metric is negative.
            return 100.0 * (value - base_value) / abs(base_value)

        integer_inputs = {
            "Project life in years", "Number of shifts per day",
            "Depreciation period in years",
            "Construction/Startup period (production starts in year)",
            "Particulate processing steps (P)", "Non-particulate processing steps (Nnp)",
        }
        sensitivity_rows = []
        for change_pct in range(-100, 101, 10):
            scenario_value = baseline * (1.0 + change_pct / 100.0)
            if parameter in integer_inputs:
                scenario_value = int(round(scenario_value))
            scenario_inputs = dict(base_inputs)
            scenario_inputs[parameter] = f"{scenario_value:.12g}"
            scenario_result = calculate_profitability(scenario_inputs, show_errors=False)
            values = [scenario_result.get(key) if scenario_result else None for key in metric_keys]
            sensitivity_rows.append({
                "change_pct": change_pct,
                "parameter_value": scenario_value,
                "metrics": values,
                "metric_change_pcts": [
                    response_change_percent(value, base_value)
                    for value, base_value in zip(values, base_metric_values)
                ],
            })

        visible_columns = ["change", "parameter_value"]
        for index, label in enumerate(selected_labels, start=1):
            self.sensitivity_tree.heading(f"metric{index}_value", text=f"{label} result")
            self.sensitivity_tree.heading(
                f"metric{index}_change", text=f"{label} change from base (%)")
            visible_columns.extend((f"metric{index}_value", f"metric{index}_change"))
        self.sensitivity_tree["displaycolumns"] = tuple(visible_columns)
        for item in self.sensitivity_tree.get_children():
            self.sensitivity_tree.delete(item)
        for row in sensitivity_rows:
            table_values = [
                f"{row['change_pct']}%",
                self._sensitivity_display_value(row["parameter_value"]),
            ]
            for index in range(3):
                if index < len(row["metrics"]):
                    table_values.extend((
                        self._sensitivity_display_value(row["metrics"][index]),
                        self._sensitivity_display_value(row["metric_change_pcts"][index]),
                    ))
                else:
                    table_values.extend(("", ""))
            self.sensitivity_tree.insert("", tk.END, values=(
                *table_values,
            ))

        self.sensitivity_ax.clear()
        plotted_series = 0
        short_metric_names = {
            "NNPW": "NNPW ($)", "DNPW": "DNPW ($)", "IRR": "IRR (%)",
            "ROI": "ROI (%)", "NPBP": "NPBP (years)", "DPBP": "DPBP (years)",
            "B_C_ratio_NCF": "Undiscounted B/C", "B_C_ratio_DCF": "Discounted B/C",
        }
        for row_index, label in enumerate(selected_labels):
            response_values = [row["metric_change_pcts"][row_index] for row in sensitivity_rows]
            if not any(value is not None for value in response_values):
                continue
            plot_values = [np.nan if value is None else value for value in response_values]
            self.sensitivity_ax.plot(
                [row["change_pct"] for row in sensitivity_rows], plot_values,
                marker="o", linewidth=2, markersize=4,
                label=short_metric_names.get(metric_keys[row_index], label),
            )
            plotted_series += 1
        self.sensitivity_ax.axhline(0, color="#536779", linewidth=1, linestyle="--")
        self.sensitivity_ax.axvline(0, color="#a0acb8", linewidth=1, linestyle=":")
        self.sensitivity_ax.set_xlabel("Selected input change from baseline (%)")
        self.sensitivity_ax.set_ylabel("Response change from baseline (%)")
        self.sensitivity_ax.set_title(f"Profitability sensitivity to {parameter_label}")
        self.sensitivity_ax.grid(True, alpha=0.28)
        if plotted_series:
            self.sensitivity_ax.legend(loc="best", fontsize=8)
        else:
            self.sensitivity_ax.text(
                0.5, 0.5,
                "Percentage response change is undefined for the selected base metric(s).\n"
                "A non-zero base result is required to plot percentage change.",
                ha="center", va="center", transform=self.sensitivity_ax.transAxes,
            )
        self.sensitivity_figure.tight_layout()
        self.sensitivity_canvas.draw_idle()

        chart_buffer = io.BytesIO()
        self.sensitivity_figure.savefig(chart_buffer, format="png", dpi=200, bbox_inches="tight")
        chart_png_bytes = chart_buffer.getvalue()
        chart_data_uri = "data:image/png;base64," + base64.b64encode(chart_png_bytes).decode("ascii")

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = f"TOPCEE_Sensitivity_{timestamp}"
        metric_units = {
            "NNPW": "USD", "DNPW": "USD", "IRR": "%", "ROI": "%",
            "NPBP": "years", "DPBP": "years",
            "B_C_ratio_NCF": "ratio", "B_C_ratio_DCF": "ratio",
        }
        headers = ["Change from baseline (%)", f"{parameter_label} input value"] + [
            heading
            for label, key in zip(selected_labels, metric_keys)
            for heading in (f"{label} result [{metric_units[key]}]",
                            f"{label} response change from baseline (%)")
        ]
        exported_rows = [[
            row["change_pct"], row["parameter_value"],
            *(value if value is not None else "N/A"
              for pair in zip(row["metrics"], row["metric_change_pcts"]) for value in pair)
        ] for row in sensitivity_rows]
        saved, errors = [], []

        figure_path = os.path.join(folder, f"{base_name}_Figure.png")
        csv_path = os.path.join(folder, f"{base_name}.csv")
        try:
            with open(csv_path, "w", newline="", encoding="utf-8") as stream:
                writer = csv.writer(stream)
                writer.writerow(["TOPCEE Sensitivity Analysis"])
                writer.writerow(["Project title", base_inputs.get("Project title", "")])
                writer.writerow(["Project summary", base_inputs.get("Project summary", "")])
                writer.writerow(["Selected parameter", parameter_label])
                writer.writerow(["Input field key", parameter])
                writer.writerow(["Baseline value", baseline])
                writer.writerow(["Method", "Multiplicative change from -100% to +100% in 10% increments; integer inputs rounded; zero baseline remains zero"])
                writer.writerow(["Response metrics", "; ".join(selected_labels)])
                writer.writerow(["Response change formula", "100 x (scenario result - base result) / abs(base result); N/A when base result is zero or undefined"])
                writer.writerow(["Figure Y-axis", "Response change from baseline (%)"])
                writer.writerow(["Figure image", os.path.basename(figure_path)])
                writer.writerow([])
                writer.writerow(headers)
                writer.writerows(exported_rows)
            saved.append(csv_path)
        except Exception as exc:
            errors.append(f"CSV: {exc}")

        try:
            with open(figure_path, "wb") as stream:
                stream.write(chart_png_bytes)
            saved.append(figure_path)
        except Exception as exc:
            errors.append(f"Figure PNG: {exc}")

        html_path = os.path.join(folder, f"{base_name}.html")
        try:
            table_rows = "".join(
                "<tr>" + "".join(f"<td>{html.escape(self._sensitivity_display_value(value))}</td>"
                                    for value in row) + "</tr>"
                for row in exported_rows
            )
            html_doc = f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><title>TOPCEE Sensitivity Analysis</title>
<style>body{{font:15px Arial,sans-serif;color:#24364b;margin:32px}}h1{{color:#164e78}}
table{{border-collapse:collapse;width:100%;margin-top:20px}}th,td{{border:1px solid #cbd5e1;padding:8px;text-align:right}}
th{{background:#eaf0f6}}td:nth-child(1){{text-align:center}}tr:nth-child(even){{background:#f7f9fc}}</style>
</head><body><h1>TOPCEE Sensitivity Analysis</h1>
<p><b>Project:</b> {html.escape(base_inputs.get('Project title', ''))}</p>
<p><b>Summary:</b> {html.escape(base_inputs.get('Project summary', ''))}</p>
<p><b>Selected input:</b> {html.escape(parameter_label)}<br><b>Baseline value:</b> {baseline:,.8g}</p>
<p>Each case changes the selected input multiplicatively from -100% to +100% in 10% increments. Integer inputs are rounded to whole values. For a zero baseline, all varied values remain zero. Invalid cases or undefined metrics are shown as N/A.</p>
<p>The figure's y-axis is response change from baseline, calculated as 100 × (scenario result - base result) ÷ |base result|. This denominator keeps the plotted direction intuitive when the base result is negative; zero or undefined base results produce N/A percentage changes.</p>
<img alt="Sensitivity response figure" style="display:block;max-width:100%;height:auto;margin:20px auto" src="{chart_data_uri}">
<table><thead><tr>{''.join(f'<th>{html.escape(header)}</th>' for header in headers)}</tr></thead>
<tbody>{table_rows}</tbody></table></body></html>"""
            with open(html_path, "w", encoding="utf-8") as stream:
                stream.write(html_doc)
            saved.append(html_path)
        except Exception as exc:
            errors.append(f"HTML: {exc}")

        xlsx_path = os.path.join(folder, f"{base_name}.xlsx")
        try:
            xlsx_rows = [
                ["TOPCEE Sensitivity Analysis"],
                ["Project title", base_inputs.get("Project title", "")],
                ["Project summary", base_inputs.get("Project summary", "")],
                ["Selected parameter", parameter_label],
                ["Input field key", parameter],
                ["Baseline value", baseline],
                ["Method", "-100% to +100% in 10% increments; integer inputs rounded; zero baseline remains zero"],
                ["Response metrics", "; ".join(selected_labels)],
                ["Response change formula", "100 x (scenario result - base result) / abs(base result); N/A when base result is zero or undefined"],
                ["Figure Y-axis", "Response change from baseline (%)"],
                [],
                headers,
                *exported_rows,
            ]
            self._write_sensitivity_xlsx(xlsx_path, xlsx_rows, chart_png_bytes)
            saved.append(xlsx_path)
        except Exception as exc:
            errors.append(f"Excel: {exc}")

        self.last_folder = folder
        save_config({"last_folder": self.last_folder, "db_path": self.db_path})
        zero_baseline_note = (
            " Because its baseline is zero, the input remains zero at every percentage level."
            if baseline == 0 else ""
        )
        self.sensitivity_status_var.set(
            f"Calculated 21 sensitivity cases for {parameter}. Generated {len(saved)} report file(s)."
            f"{zero_baseline_note}")
        if saved:
            message = "Sensitivity reports saved to:\n" + "\n".join(saved)
            if errors:
                message += "\n\nSome exports failed:\n" + "\n".join(errors)
                messagebox.showwarning("Sensitivity Analysis", message)
            else:
                messagebox.showinfo("Sensitivity Analysis", message)
        else:
            messagebox.showerror("Sensitivity Analysis", "No reports were generated.\n" + "\n".join(errors))

    # ---------------------------------------------- Equipment Tab Methods --




    # ---------------------------------------------- User Guide Tab ----------
    def _build_user_guide_tab(self):
        frame = self.tab_guide

        # A single calm page color, consistent white guide cards, and a readable scroll area.
        canvas = tk.Canvas(frame, bg="#f3f6fa", highlightthickness=0)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        guide_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(guide_window, width=e.width))
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Title
        title = ttk.Label(scroll_frame, text="TOPCEE User Guide", style="GuideTitle.TLabel")
        title.pack(anchor="w", fill=tk.X, padx=18, pady=(20, 4))

        subtitle = ttk.Label(
            scroll_frame,
            text="A practical guide to equipment costing, process inputs, TEA results, and report generation.",
            style="GuideSubtitle.TLabel", wraplength=1100, justify=tk.LEFT,
        )
        subtitle.pack(anchor="w", padx=18, pady=(0, 12))

        sections = []

        # Overview
        sections.append(self._guide_section(scroll_frame, "1. Overview",
            "TOPCEE v2.0 is a process economic analysis application for chemical and process plants. "
            "It helps you estimate equipment costs, perform CAPEX/OPEX analysis, evaluate profitability, "
            "and generate professional cost reports for your plant design project.\n\n"
            "Follow the tabs from left to right. The workflow has five stages:\n"
            "  Step 1: Equipment Costing  — Select plant equipment from a cost database and estimate costs automatically.\n"
            "  Step 2: Process Inputs     — Enter all economic parameters for your process.\n"
            "  Step 3: Results            — Calculate CAPEX, OPEX, cash flow, and profitability indicators.\n"
            "  Step 4: Sensitivity Analysis — Explore how selected inputs change one to three profitability results.\n"
            "  Step 5: Generate Report    — Save the analysis and cash-flow results in report formats."))

        sections.append(self._guide_section(scroll_frame, "2. Before You Start — Collect Data from Your Process Model",
            "TOPCEE does not read an Aspen, DWSIM, or other simulator project file directly. Use your process model, stream tables, equipment summaries, and project basis to collect the values below, then enter them in TOPCEE. You can begin with preliminary estimates; record the source and assumptions so you can update them later.\n\n"
            "Process and production basis:\n"
            "  • Project title and a short summary of the process and product.\n"
            "  • Product name, design production rate (TOPCEE asks for L/h), expected product selling price, and the basis for that price. Convert a mass-based product rate to L/h using the product density before entry.\n"
            "  • Expected operating hours per day, operating days per year, shifts per day, and project life. Confirm that hours/day × days/year is consistent with the annual operating schedule in your model.\n"
            "  • Process type (fluid, solid, or mixed), and consistent counts of particulate (P) and non-particulate (Nnp) processing steps for the labor correlation.\n\n"
            "Equipment and cost basis:\n"
            "  • A PFD or equipment list with equipment names, PFD/tag IDs, design capacities, units, and number of parallel or installed units. Keep the Process ID consistent with the tag on your PFD.\n"
            "  • Equipment reference cost, reference capacity, reference source/year, material and pressure assumptions if you are not using the database. For database costing, confirm the selected item and its units before adding it to the project list.\n"
            "  • Any independent fixed-capital or working-capital estimate you have, for comparison. TOPCEE estimates these from equipment purchase cost using the selected Lang-factor method; treat this as an early-stage estimate.\n\n"
            "Annual operating costs and financial assumptions:\n"
            "  • Each raw material's flow rate, flow unit, price, price basis (per kg, tonne, litre, or cubic metre), and density when a mass-to-volume conversion is needed.\n"
            "  • Each utility's use rate, unit, and unit price; and each catalyst or other process consumable's continuous use rate or replacement charge, price, and operating-hour replacement interval.\n"
            "  • Local currency code and exchange rate in local currency per USD. Enter product prices, wages, materials, utilities, and consumables in that local currency. TOPCEE reports the economic results in USD.\n"
            "  • Direct operator hourly wage, income tax and subsidy assumptions, discount rate, depreciation life, construction/startup year (the year production begins), and any basis for those values.\n\n"
            "Before running the TEA, check that units agree, every price has a clear unit basis, equipment IDs match the PFD, and the annual operating schedule matches the process model. Sensitivity analysis uses the entered values as its base case, so decide which uncertain inputs and profitability measures you want to examine."))

        # Step 1
        sections.append(self._guide_section(scroll_frame, "3. Step 1 — Equipment Costing",
            "Purpose: Build a list of equipment for your plant and calculate each item's cost automatically.\n\n"
            "How to use:\n"
            "  a. Click 'Load Sheet' to open your Capital-Equipment-Cost Excel database.\n"
            "     (Default path: C:\\\\Users\\\\Dr_Oyegoke\\\\Downloads\\\\Capital-Equipment-Cost-Database.xlsx)\n"
            "  b. If your workbook has multiple sheets, select the correct one from the 'Sheet' dropdown.\n"
            "  c. Scroll through the equipment list on the left. Use the Search box to filter.\n"
            "  d. Click any equipment row to select it — all column values appear in the right panel.\n"
            "     Ref Capacity and Ref Cost are read from the selected row; for range-based data, the matching Min Scale and Min Cost pair is used.\n"
            "     Scaling Exponent and scaling unit are also filled from the database where available.\n"
            "     Ref Year is extracted from the equipment's Source citation. Ref CEPCI is derived from that year using TOPCEE's reported 2007–2023 CEPCI data or its OLS regression; the equipment database's CEPCI column is not used for escalation.\n"
            "  e. In 'Your Design Specification', enter:\n"
            "     • Design Capacity — your plant's required capacity for this equipment\n"
            "     • Scaling Exponent — auto-filled from the database when listed (otherwise 0.6), and editable\n"
            "     • Quantity — number of units\n"
            "     • Ref Year — auto-filled from the source citation; if no year is found, enter it manually to generate Ref CEPCI\n"
            "     • Target Year — defaults to the current calendar year; Target CEPCI is auto-generated from reported data or regression\n"
            "     • Target CEPCI — auto-estimated and editable so you can replace a regression value with a published index\n"
            "     • Material Factor / Pressure Factor — adjustment factors (default 1.0)\n"
            "  f. Click 'Calculate Cost' — the app computes:\n"
            "        CI = Target CEPCI / Ref CEPCI\n"
            "        Unit Cost = Ref Cost × (Design Cap / Ref Cap)^Exponent × CI × Material × Pressure\n"
            "        Total Cost = Unit Cost × Quantity\n"
            "     For years absent from the reported 2007–2023 data, Target CEPCI uses OLS linear regression: CEPCI = 604.7176 + 13.0485 × (year − 2015), R² = 0.576.\n"
            "     This is a simple trend estimate, not an official CEPCI release or a strong long-range forecast. For 2026 it estimates about 748.25, below the reported 2023 value of 797.9; replace it with a published value when available.\n"
            "  g. Click 'Add to Project List' (green button) to add the item.\n"
            "  h. Repeat for all equipment in your plant.\n"
            "  i. The running total appears at the bottom.\n"
            "  j. Click 'Use Total as PCE in Process Inputs' to send the total equipment cost to Step 2.\n"
            "  k. You can also 'Export Equipment List (CSV)' for your records."))

        # Step 2
        sections.append(self._guide_section(scroll_frame, "4. Step 2 — Process Inputs",
            "Purpose: Enter all economic parameters needed for CAPEX, OPEX, and profitability analysis.\n\n"
            "The Process Inputs page is organized into three scrollable subtabs:\n"
            "  • Input Details — enter a project title and summary before other inputs, review guidance, and load previously saved JSON files, including older saved inputs. Default title and summary text are supplied if these are left blank.\n"
            "  • Economic Parameters — project life, discount rate, depreciation period, construction/startup year, local currency and exchange rate, operating hours/day, shifts/day, operating days/year, hourly operator wage, product price, process type, production rate, particulate/non-particulate labor-step counts, and equipment cost.\n"
            "  • Operating Cost Parameters — up to 10 process consumables, 10 raw materials, and 10 utilities. Consumables support continuous use or periodic replacement; local-currency prices are converted to USD. Suggested utility names are pre-filled and editable.\n\n"
            "Fields include:\n"
            "  • Local currency code and exchange rate (local currency per USD)\n"
            "  • Operating labor cost per hour, selling price, and material prices entered in your local currency; results are converted to USD\n"
            "  • Operating hours/day (default 24), shifts/day (default 3), and operating days/year (default 330)\n"
            "  • Hourly operating labor wage (default 3,850 local currency/hour, equivalent to 77,000 ÷ 20)\n"
            "  • Selling price of product (local currency/L)\n"
            "  • Project life (years), Discount rate (%)\n"
            "  • Depreciation period (default 10 years) and construction/startup period, entered as the year production begins (default year 3)\n"
            "  • Up to 10 raw materials, each with a selectable mass or volume flow rate (kg/h, t/h, L/h, or m³/h) and price basis (per kg, t, L, or m³). Enter density in kg/L only when converting between mass flow and volume-based price, or volume flow and mass-based price.\n"
            "  • Up to 10 utilities, each with its own quantity, unit, and local-currency price\n"
            "  • Process Consumables has up to 10 rows for catalysts, sorbents, resins, membranes, and similar items. Choose the costing mode that matches how the item is consumed or replenished:\n"
            "      - Continuous — use this for ongoing feed, make-up, or dosing during operation. Enter quantity per operating hour (for example, 2 kg/h), its unit, and local price per unit. The replacement interval is ignored and may remain 0. Annual cost (local/year) = quantity per hour × price per unit × operating hours/day × operating days/year.\n"
            "      - Periodic replacement — use this for a fixed charge replaced after a campaign or a stated amount of plant operating time, such as a catalyst bed or adsorbent charge. Enter the amount replaced each event (for example, 100 kg/event), the price per kg, and the interval in operating hours between replacement events. Do not enter a per-hour rate in this mode.\n"
            "      - Annualized cost (local/year) = quantity per event × price per unit × (operating hours/day × operating days/year) ÷ replacement interval. Example: with 24 operating h/day and 330 operating days/year, a 100 kg charge replaced every 4,000 operating hours at 500 local currency/kg costs 100 × 500 × (24 × 330) ÷ 4,000 = 99,000 local currency/year. TOPCEE then divides by the exchange rate to report USD/year.\n"
            "      - Enter a positive replacement interval when both event quantity and unit price are positive; otherwise TOPCEE asks for a valid interval. The interval is plant running time, not calendar days. This is an annualized operating-cost estimate: TOPCEE does not schedule individual replacement cash flows or add a separate initial charge at startup.\n"
            "  • Any raw-material, utility, or consumable row with both its quantity and price set to zero is ignored in calculations and reports, even if a suggested name remains in the row.\n"
            "  • Enter each catalyst, adsorbent, resin, membrane, or other consumable only once in Process Consumables; do not also list it as a raw material or utility. Suggested utility names and units can be edited or cleared.\n"
            "  • Production rate (L/h), and particulate (P) and non-particulate (Nnp) process-step counts\n"
            "\nHow to count steps for the operating-labor correlation:\n"
            "  • P counts process steps that handle particulate solids, such as solids feeding, conveying/distribution, size reduction or classification, and particulate removal/separation. This count is squared in the correlation, so enter only the particulate-solids steps.\n"
            "  • Nnp counts main process steps that handle non-particulate streams, such as compression, heating/cooling, mixing, reaction, and fluid separation. Do not put these steps in P. Pumps, storage vessels, pipes, and instruments are generally excluded from Nnp in the cited application examples.\n"
            "  • Count principal operations from the PFD at a consistent level; do not count pipes, valves, instruments, or every equipment tag as a separate operation. Apply one consistent convention to parallel or repeated operations.\n"
            "  • Example: a gas/liquid flowsheet with a compressor, heat exchanger, reactor, and distillation column could use P = 0 and Nnp = 4. A solids feed, milling, and screening train could add P = 3. These are empirical labor-correlation counts, not equipment-cost totals.\n"
            "  • The labor correlation estimates operators per shift: NOL_shift = ceil(√(6.29 + 31.7 × P² + 0.23 × Nnp)). Shift hours = operating hours/day ÷ shifts/day. The model assumes 1,960 paid hours/operator-year (49 weeks × 5 days × 8 hours). Total operators = ceil(NOL_shift × shifts/day × operating days/year × shift hours ÷ 1,960). Annual labor cost in USD = total operators × 1,960 × hourly wage in local currency ÷ exchange rate.\n"
            "  • Rounding up covers complete shift positions and whole employees. At 24 operating hours/day and 3 shifts/day, shifts are 8 hours; at 2 shifts/day, shifts are 12 hours. Hourly wage is multiplied by assumed paid employee hours/year, not plant operating hours. Enter the direct-operator hourly rate; operating supervision remains a separate estimated cost.\n"
            "  • Older saved inputs used monthly wages. On loading, TOPCEE converts monthly wage to hourly using 12 months ÷ 1,960 paid hours/year. The old processing count loads into Nnp for review; the old non-processing support count is ignored by this correlation. Recheck P and Nnp after loading an older file.\n"
            "  • Process Type: F = Fluid, S = Solid, M = Mixed\n"
            "  • Total equipment cost ($) — auto-filled from Step 1, or enter manually\n\n"
            "When you click Calculate Profitability, TOPCEE asks whether to save the current inputs before calculating. Sensitivity Analysis uses the current entries directly and does not change them. Save inputs as JSON for later reuse; Input Details also loads saved or older JSON input files.\n\n"
            "Tips:\n"
            "  • Use the save prompt or save control to store your entries as a JSON file.\n"
            "  • Load a previously saved file from Input Details.\n"
            "  • The Total equipment cost field is read-only if set from Step 1."))

        # Step 3
        sections.append(self._guide_section(scroll_frame, "5. Step 3 — Results",
            "Purpose: Run the full economic calculation and view results.\n\n"
            "How to use:\n"
            "  a. Click 'Calculate Profitability' (purple button). The calculation also runs from the button at the bottom of Process Inputs.\n"
            "  b. The Results page has separate Analysis Results and Cash Flow Diagram subtabs; the diagram is arranged so result figures remain visible.\n"
            "  c. Results are displayed in three categories:\n"
            "     CAPEX: PCE, DPC, FCI, WC, TCI\n"
            "     OPEX/COM: Raw materials, process consumables, utilities, labor, DMC, FMC, GE, COM, Cost per Liter\n"
            "     Profitability: NNPW, DNPW, Payback periods, ROI, IRR, B/C ratios\n"
            "  d. The Cash Flow Diagram subtab shows NCF, CNCF, DCF, and CDCF.\n"
            "  e. Click 'Save Results (CSV)' to export computed metrics and the annual cash-flow series.\n"
            "  f. After reviewing the base case, continue to Sensitivity Analysis if you want to explore uncertain inputs. The sensitivity tool recalculates from current Process Inputs, so a previous calculation is not technically required.\n"
            "  g. Use the Sensitivity Analysis tab for scenario studies. To skip sensitivity, click 'Go to Report' as a shortcut to the Generate Report tab."))

        # Step 4
        sections.append(self._guide_section(scroll_frame, "6. Step 4 — Sensitivity Analysis",
            "Purpose: See how changing a selected input affects one or more profitability measures while keeping the other inputs at their current values. For a useful comparison, first calculate and review the base case. Sensitivity Analysis can also calculate its own base case from the current inputs.\n\n"
            "How to use:\n"
            "  a. Select one numeric Process Input to vary. Economic Parameters appear by their field name. Operating Cost Parameters appear by the current name and unit in each active material, utility, or consumable row. If you edit a row name, use 'Refresh from Process Inputs' if the updated name is not listed.\n"
            "  b. Select one, two, or three different profitability response variables.\n"
            "  c. Use the built-in sweep from −100% to +100% in 10 percentage-point steps. A −100% change sets a positive base input to zero. A zero-valued base input cannot produce a meaningful percentage sweep; enter a non-zero base value first.\n"
            "  d. Read the table and figure together. The x-axis is the input change; the y-axis is percentage change in each selected response from its base-case value: 100 × (scenario − base) ÷ |base|. N/A means the base response is zero or undefined.\n"
            "  e. Sensitivity results are separate from the base-case TEA reports. Generate the sensitivity report from this tab; it produces HTML, Excel, and CSV files, and the CSV identifies its accompanying PNG figure. These files report scenario results and do not replace the main TEA report."))

        # Step 5
        sections.append(self._guide_section(scroll_frame, "7. Step 5 — Generate Report",
            "Purpose: Produce professional cost reports and save them to your chosen folder.\n\n"
            "How to use:\n"
            "  a. Select an Output Folder (use 'Browse' or type a path).\n"
            "     Click 'Remember this folder' to save it for future sessions.\n"
            "  b. Choose which reports to generate:\n"
            "     • Equipment Cost Schedule (CSV) — itemized equipment costs\n"
            "     • Full TEA Report — CAPEX, OPEX/COM, profitability metrics, cash-flow data, equipment, inputs, and notes\n"
            "  c. Click 'Generate All Selected Reports' (green button).\n"
            "  d. Reports are saved as:\n"
            "     • HTML and PDF — project title and summary, CAPEX, OPEX/COM, profitability metrics, annual cash-flow data, equipment, inputs, and notes; PDF also includes a cash-flow chart\n"
            "     • CSV — project title and summary, spreadsheet-compatible breakdowns, computed metrics, equipment, notes, and annual cash-flow rows\n"
            "     • Excel — project title and summary, formatted breakdowns plus annual cash-flow table and chart; Excel export requires openpyxl\n"
            "  e. The Report Log shows all saved file paths.\n"
            "  f. Click 'Open Output Folder' to view your saved files."))

        # Formulas
        sections.append(self._guide_section(scroll_frame, "8. Key Formulas and Cash Flow Analysis",
            "Equipment Cost Scaling:\n"
            "  C₂ = C₁ × (Q₂ / Q₁)^n × CI × MF × PF\n"
            "  Where: C₁ = reference cost, Q₁ = reference capacity, Q₂ = design capacity,\n"
            "         n = scaling exponent, CI = target CEPCI / reference CEPCI, MF = material factor, PF = pressure factor.\n"
            "  Target CEPCI is taken from reported annual values for 2007–2023 and otherwise estimated with OLS linear regression; model R² = 0.576.\n\n"
            "CAPEX (Lang Factor Method):\n"
            "  DPC = PCE × Lang factor[8]   (Direct Plant Cost)\n"
            "  FCI = PCE × Lang factor[14]  (Fixed Capital Investment)\n"
            "  TCI = FCI + Working Capital  (Total Capital Investment)\n\n"
            "OPEX / Cost of Manufacturing:\n"
            "  COM = Raw Material + Utilities + Labor + Maintenance + Overheads + Depreciation + Taxes + Insurance + General Expenses\n\n"
            "Annual cash flow (t = 0, …, project life − 1; production starts in the selected construction/startup year S):\n"
            "  For startup year S ≥ 3, capital outlays are I₀ = 0.6 × TDC and I₁ = 0.4 × TDC; for S = 1 or 2, TDC is funded in year 0. Working capital is added at I₍S−1₎ and recovered in the final year.\n"
            "  Depreciation is straight-line at TDC / selected depreciation life per operating year, beginning in year S and capped at the TDC basis (zero salvage assumption). The default life is 10 years.\n"
            "  Rₜ = annual revenue, including subsidy, for t ≥ S and zero before year S. COMcash = COM − the model's annual depreciation estimate. uₜ is 1 in an operating year and 0 before production.\n"
            "  Taxable income = Rₜ − COMcash × uₜ − dₜ; after-tax operating cash flow Yₜ = taxable income × (1 − tax rate) + dₜ. Tax depreciation dₜ stops when its basis is exhausted.\n"
            "  NCFₜ = Yₜ − Iₜ; CNCFₜ = Σⱼ₌₀ᵗ NCFⱼ.\n"
            "  DCFₜ = NCFₜ / (1 + discount rate)ᵗ; CDCFₜ = Σⱼ₌₀ᵗ DCFⱼ.\n"
            "  NNPW is the final CNCF and DNPW is the final CDCF. These annual series are included in generated reports.\n\n"
            "After-tax project IRR solves Σ NCFₜ/(1+IRR)ᵗ = 0; negative IRRs are allowed, and IRR is undefined for nonconventional cash flows with multiple sign changes or no bracketed root. Project payback is interpolated at the first cumulative cash-flow crossing, measured from year 0; discounted payback uses DCF. Simple annual after-tax ROI = annual net profit / TCI × 100; it does not account for cash-flow timing. Benefit-cost ratios use present-worth gross revenues plus tax credits divided by present-worth capital outlays, cash operating costs, and payable taxes, at 0% and the selected discount rate, respectively.\n\n"
            "Assumptions and limits: production begins in the selected startup year (default year 3); the 60/40 TDC funding schedule is a model assumption for startup year 3 or later, not a universal expenditure schedule. The model assumes constant annual operating performance, immediate tax effects on losses, zero separate salvage proceeds, and no financing cash flows or tax-loss carryforward. Working capital and land are recovered at closure as one combined estimate. Use jurisdiction-appropriate tax depreciation and project-specific capital timing for investment decisions.\n\n"
            "Equipment data acknowledgement: Capital Equipment Cost Database, Global CO₂ Initiative / AssessCCUS: https://assessccus.globalco2initiative.org/wp-content/uploads/Capital-Equipment-Cost-Database.xlsx\n\n"
            "Further process design and economics references: Peters, Timmerhaus & West, Plant Design and Economics for Chemical Engineers; Turton et al., Analysis, Synthesis, and Design of Chemical Processes; Towler & Sinnott, Chemical Engineering Design; and Harmsen et al., Product and Process Design: Driving Innovation (2018)."))

        sections.append(self._guide_section(scroll_frame, "9. Plain-Language Glossary",
            "CAPEX — capital expenditure: the estimated investment to build or install the plant.\n"
            "PCE — purchased equipment cost: the cost of the major equipment before plant-wide Lang-factor additions.\n"
            "DPC — direct plant cost: PCE plus estimated installation, instrumentation, piping, electrical, buildings, yard, and service-facility costs.\n"
            "IPC — indirect plant cost: estimated engineering/supervision and construction expenses.\n"
            "FCI — fixed capital investment: the estimated fixed plant investment, including direct and indirect costs and capital additions.\n"
            "WC — working capital and land allowance used by this estimate. TCI — total capital investment: FCI plus WC.\n"
            "OPEX — operating expenditure. COM — cost of manufacturing: the model's annual manufacturing, fixed, and general operating costs.\n"
            "NCF / DCF — net cash flow / discounted cash flow for an individual year. CNCF / CDCF — their cumulative totals through that year.\n"
            "NNPW / DNPW — final cumulative cash flow without / with discounting. NPV means net present value; these measures summarize project cash flows on the two bases.\n"
            "IRR — the discount rate that makes project NPV zero. ROI — annual after-tax net profit divided by TCI; this simple ratio does not account for timing. B/C — present-worth benefits divided by present-worth costs.\n"
            "Sensitivity analysis — a one-input-at-a-time scenario sweep that shows how selected profitability measures change when one input changes, holding other inputs constant."))

        # Tips
        sections.append(self._guide_section(scroll_frame, "10. Tips and Troubleshooting",
            "  • Start with the Equipment Costing tab to build your plant's equipment list.\n"
            "  • Use realistic design capacities based on your process simulation (e.g., Aspen HYSYS, DWSIM).\n"
            "  • The scaling exponent (typically 0.6–0.7) depends on equipment type.\n"
            "  • Save your inputs frequently using 'Save Inputs'.\n"
            "  • Always run 'Calculate Profitability' before generating reports.\n"
            "  • TOPCEE needs numpy and matplotlib to run its desktop interface and plots. Excel database reading and Excel report export require openpyxl; PDF export also requires matplotlib. If an export fails, read the Report Log for the error.\n"
            "  • The output folder is remembered across sessions once you click 'Remember this folder'.\n"
            "  • If an Excel database or workbook export does not work, install openpyxl. A developer setup also needs numpy and matplotlib. Typical command: python -m pip install numpy openpyxl matplotlib. Packaged app users should use the installer that includes these dependencies."))

        # Pack all sections
        for section in sections:
            section.pack(fill=tk.X, padx=18, pady=6)

        ttk.Label(scroll_frame, text="", padding=10).pack()
        self._bind_scroll_wheel(scroll_frame, canvas)

    def _build_about_tab(self):
        """Build a dedicated About page for the application and its developer."""
        frame = self.tab_about
        canvas = tk.Canvas(frame, bg="#f3f6fa", highlightthickness=0)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        content = ttk.Frame(canvas)
        content.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        content_window = canvas.create_window((0, 0), window=content, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(content_window, width=e.width))
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Label(content, text="About TOPCEE", style="GuideTitle.TLabel").pack(
            anchor="w", fill=tk.X, padx=18, pady=(20, 4))
        ttk.Label(
            content,
            text="Purpose, capabilities, and the people behind the Process Cost & Economic Evaluator.",
            style="GuideSubtitle.TLabel", wraplength=1100, justify=tk.LEFT,
        ).pack(anchor="w", padx=18, pady=(0, 12))

        sections = [
            ("What TOPCEE is",
             "TOPCEE means Toyese Oyegoke Process Cost & Economic Evaluator. It is a desktop tool for preliminary techno-economic evaluation of chemical and process plant designs. The application brings equipment cost scaling, plant capital and operating cost estimation, cash-flow analysis, and report generation into one guided workflow."),
            ("What the application helps you do",
             "• Build an equipment list with process IDs, design specifications, database reference capacity and cost, and scaled equipment costs.\n"
             "• Enter project economics in a selected local currency and provide the exchange rate to USD.\n"
             "• Enter up to 10 raw materials with selectable mass/volume flow and price units; density enables conversion when their bases differ. Unused zero-quantity/zero-price rows are excluded.\n"
             "• Estimate CAPEX using the selected process type's Lang-factor relationships and estimate OPEX/COM from entered operating costs and model correlations.\n"
             "• Review profitability indicators and yearly nominal and discounted cash flows, then export results to HTML, CSV, and Excel."),
            ("Calculation scope",
             "TOPCEE reports equipment cost estimates, direct and indirect plant costs, fixed and total capital investment, manufacturing cost components, and project cash-flow indicators. Cash-flow results include NCF, CNCF, DCF, and CDCF over the entered project life. The assumptions, correlations, cost database, currency conversion, and project-specific inputs affect the estimates; review them for each study and retain the generated input and result files with the project record."),
            ("Developer",
             "TOPCEE is developed by Engr. Dr. T. Oyegoke, Pencil Team - CAD Engineering of Processes & Reactive Interfaces (PT-CEPRIs) Group, Chemical Engineering Department, Ahmadu Bello University Zaria, NIGERIA\n\n"
             "ResearchGate: https://www.researchgate.net/profile/Toyese-Oyegoke/research\n"
             "Google Scholar: https://scholar.google.com/citations?hl=en&user=l5bQoIQAAAAJ&view_op=list_works&sortby=pubdate\n"
             "ORCID: https://orcid.org/0000-0002-2026-6864\n\n"
             "Contact: OyegokeToyese@gmail.com"),
            ("Using and interpreting results",
             "TOPCEE supports engineering-economic estimation and screening during process design. Results are estimates based on the supplied data and implemented methods. For project decisions, check input units and currency, verify database references and assumptions, and refine estimates with current vendor quotations and site-specific information."),
        ]
        for title, body in sections:
            self._guide_section(content, title, body).pack(fill=tk.X, padx=18, pady=6)
        ttk.Label(content, text="", padding=10).pack()
        self._bind_scroll_wheel(content, canvas)

    def _guide_section(self, parent, title, text):
        """Create a uniformly styled guide card."""
        lf = ttk.LabelFrame(parent, text=title, style="GuideCard.TLabelframe", padding=14)
        lbl = ttk.Label(lf, text=text, style="GuideBody.TLabel", justify=tk.LEFT,
                        wraplength=1100)
        lbl.pack(anchor="w", fill=tk.X)
        return lf

    # ---------------------------------------------- Inputs Tab Methods -----
    def _get_inputs(self) -> Dict[str, str]:
        return {k: e.get() for k, e in self.input_entries.items()}

    def _save_inputs(self):
        inputs = self._get_inputs()
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("TOPCEE Inputs", "*.json"), ("All files", "*.*")],
            title="Save Inputs",
            initialdir=self.last_folder,
        )
        if not path:
            return False
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(inputs, f, indent=2)
            self.last_folder = os.path.dirname(path)
            save_config({"last_folder": self.last_folder, "db_path": self.db_path})
            messagebox.showinfo("Saved", f"Inputs saved to:\n{path}")
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save inputs:\n{e}")
            return False

    def _load_inputs(self):
        path = filedialog.askopenfilename(
            filetypes=[("TOPCEE Inputs", "*.json"), ("All files", "*.*")],
            title="Load Inputs",
            initialdir=self.last_folder,
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                inputs = json.load(f)
            aliases = {
                "Exchange rate (N/$)": "Exchange rate (local currency per $)",
                "Wage (N/month)": "Operating labor cost (local/hour)",
                "Wage (local/month)": "Operating labor cost (local/hour)",
                "Working days in d/year": "Operating days per year",
                "Selling price of your product in $/L": "Selling price of your product (local/L)",
                # Older files' general processing-step count is best mapped to
                # Nnp; old "non-processing" support counts are not part of P/Nnp.
                "Number of processing steps/units": "Non-particulate processing steps (Nnp)",
                "Catalyst quantity (kg/h)": "Process consumable 1 quantity",
                "Catalyst consumption (kg/h)": "Process consumable 1 quantity",
                "Catalyst cost (local/kg)": "Process consumable 1 cost (local/unit)",
                "Catalyst cost (N/kg)": "Process consumable 1 cost (local/unit)",
                "Cooling water quantity (kg/h)": "Utility 2 quantity (unit/h)",
                "Cooling water cost (N/kg)": "Utility 2 cost (local/unit)",
                "Process water quantity (kg/h)": "Utility 3 quantity (unit/h)",
                "Process water cost (N/kg)": "Utility 3 cost (local/unit)",
                "Power demand quantity (kWh)": "Utility 4 quantity (unit/h)",
                "Power supply cost (N/kWh)": "Utility 4 cost (local/unit)",
                "Other utilities quantity (kg/h)": "Utility 5 quantity (unit/h)",
                "Other utilities cost (N/kg)": "Utility 5 cost (local/unit)",
            }
            for i in range(1, 11):
                old_qty = ("Raw material quantity (kg/h)" if i == 1
                           else f"Raw material {i} quantity (kg/h)")
                old_price = ("Raw material cost (N/kg)" if i == 1
                             else f"Raw material {i} cost (local/kg)")
                aliases[old_qty] = f"Raw material {i} quantity (unit/h)"
                aliases[old_price] = f"Raw material {i} cost (local/unit)"
            for k, v in inputs.items():
                target = k if k in self.input_entries else aliases.get(k)
                if target in self.input_entries:
                    if k in ("Wage (N/month)", "Wage (local/month)"):
                        try:
                            v = float(v) * 12 / OPERATOR_WORK_HOURS_PER_YEAR
                        except (TypeError, ValueError):
                            pass
                    if k == "Selling price of your product in $/L":
                        rate = inputs.get("Exchange rate (N/$)", "1")
                        try:
                            v = float(v) * float(rate)
                        except (TypeError, ValueError):
                            pass
                    self.input_entries[target].delete(0, tk.END)
                    self.input_entries[target].insert(0, str(v))
            legacy_utility_names = {
                "Cooling water quantity (kg/h)": ("Utility 2 name", "Cooling water"),
                "Process water quantity (kg/h)": ("Utility 3 name", "Process water"),
                "Power demand quantity (kWh)": ("Utility 4 name", "Power"),
                "Other utilities quantity (kg/h)": ("Utility 5 name", "Other utilities"),
            }
            for legacy_key, (name_key, name) in legacy_utility_names.items():
                if legacy_key in inputs and name_key in self.input_entries:
                    self.input_entries[name_key].delete(0, tk.END)
                    self.input_entries[name_key].insert(0, name)

            # Migrate saved catalyst utility rows into the consumables table.
            utility_defaults = [
                "Electricity", "Cooling water", "Process water", "Steam", "Compressed air",
                "Chilled water", "Boiler fuel", "Instrument air", "Nitrogen", "Other utility",
            ]
            for i in range(1, 11):
                # Legacy files used kg/h and local/kg; retain their original basis.
                for field, value in (("rate unit", "kg/h"), ("price basis", "kg"),
                                     ("density (kg/L)", "0")):
                    key = f"Raw material {i} {field}"
                    if key in self.input_entries and key not in inputs:
                        self.input_entries[key].delete(0, tk.END)
                        self.input_entries[key].insert(0, value)

            for i in range(1, 11):
                name_key = f"Utility {i} name"
                qty_key = f"Utility {i} quantity (unit/h)"
                unit_key = f"Utility {i} unit"
                cost_key = f"Utility {i} cost (local/unit)"
                if "catalyst" not in self.input_entries[name_key].get().lower():
                    continue
                qty = self.input_entries[qty_key].get()
                unit = self.input_entries[unit_key].get() or "kg"
                cost = self.input_entries[cost_key].get()
                target = None
                catalyst_already_loaded = False
                for j in range(1, 11):
                    c_name = self.input_entries[f"Process consumable {j} name"].get().lower()
                    c_qty = self.input_entries[f"Process consumable {j} quantity"].get()
                    c_price = self.input_entries[f"Process consumable {j} cost (local/unit)"].get()
                    try:
                        c_is_used = float(c_qty or 0) != 0 or float(c_price or 0) != 0
                    except ValueError:
                        c_is_used = bool(c_qty.strip() or c_price.strip())
                    if "catalyst" in c_name and c_is_used:
                        catalyst_already_loaded = True
                        break
                if not catalyst_already_loaded:
                    for j in range(1, 11):
                        target_qty = self.input_entries[f"Process consumable {j} quantity"].get()
                        target_cost = self.input_entries[f"Process consumable {j} cost (local/unit)"].get()
                        try:
                            row_is_empty = not float(target_qty or 0) and not float(target_cost or 0)
                        except ValueError:
                            row_is_empty = False
                        if row_is_empty:
                            target = j
                            break
                if target is not None:
                    for key, value in (
                        (f"Process consumable {target} name", "Catalyst"),
                        (f"Process consumable {target} mode", "Continuous"),
                        (f"Process consumable {target} quantity", qty),
                        (f"Process consumable {target} unit", unit),
                        (f"Process consumable {target} cost (local/unit)", cost),
                    ):
                        entry = self.input_entries[key]
                        if key.endswith(" mode"):
                            entry.set(value)
                        else:
                            entry.delete(0, tk.END)
                            entry.insert(0, str(value))
                for key, value in ((qty_key, "0"), (cost_key, "0"),
                                   (name_key, utility_defaults[i - 1])):
                    entry = self.input_entries[key]
                    entry.delete(0, tk.END)
                    entry.insert(0, value)
            self.last_folder = os.path.dirname(path)
            save_config({"last_folder": self.last_folder, "db_path": self.db_path})
            messagebox.showinfo("Loaded", f"Inputs loaded from:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load inputs:\n{e}")

    # ---------------------------------------------- Calculation ------------
    def _run_calculation(self):
        save_choice = messagebox.askyesnocancel(
            "Save Inputs Before Analysis?",
            "Would you like to save all current input data before running the analysis?",
        )
        if save_choice is None:
            return
        if save_choice and not self._save_inputs():
            return

        inputs = self._get_inputs()

        # Validate critical fields are present and numeric
        critical_fields = [
            'Local currency code', 'Exchange rate (local currency per $)',
            'Project life in years', 'Discount rate in %',
            'Working hour in h/day', 'Number of shifts per day',
            'Operating days per year', 'Operating labor cost (local/hour)',
            'Total equipment cost ($)',
        ]
        for field in critical_fields:
            val = inputs.get(field, '').strip()
            if val == '':
                messagebox.showerror("Input Error",
                                     f"Please enter a value for:\n{field}")
                return

        try:
            startup_value = float(inputs.get(
                "Construction/Startup period (production starts in year)", "3"))
            if not startup_value.is_integer():
                raise ValueError("Construction/Startup period must be a whole year number.")
            startup_year = int(startup_value)
            project_life = int(float(inputs.get("Project life in years", "0")))
            depreciation_value = float(inputs.get("Depreciation period in years", "10"))
            if not depreciation_value.is_integer():
                raise ValueError("Depreciation period must be a whole number of years.")
            depreciation_life = int(depreciation_value)
            if startup_year < 1:
                raise ValueError("Construction/Startup period must be year 1 or later.")
            if project_life <= startup_year:
                raise ValueError("Project life must extend beyond the startup year.")
            if depreciation_life <= 0:
                raise ValueError("Depreciation period must be greater than zero.")
        except (ValueError, TypeError) as exc:
            messagebox.showerror("Input Error", str(exc))
            return

        try:
            result = calculate_profitability(inputs)
        except Exception as e:
            messagebox.showerror("Calculation Error", str(e))
            return

        if result is None:
            # calculate_profitability already showed an error message
            return

        self.calc_result = result
        self._display_results(result)
        self.notebook.select(self.tab_results)
        self.status_var.set("Calculation complete.")

    def _display_results(self, result: Dict[str, Any]):
        # Clear
        for w in self.res_frame.winfo_children():
            w.destroy()

        # Result cards
        cards = [
            ("CAPEX / Total Capital Investment", [
                ("Purchased Equipment Cost (PCE)", result["PCE"], "$"),
                ("Direct Plant Cost (DPC)", result["DPC"], "$"),
                ("Fixed Capital Investment (FCI)", result["FCI"], "$"),
                ("Working Capital (WC)", result["WC"], "$"),
                ("Total Capital Investment (TCI)", result["TCI"], "$"),
            ]),
            ("OPEX / Cost of Manufacturing", [
                ("Raw Material Cost", result["RM"], "$/yr"),
                ("Process Consumables", result["process_consumable_cost_total"], "$/yr"),
                ("Other Utilities", result["UC"] - result["PC"], "$/yr"),
                ("Power Cost", result["PC"], "$/yr"),
                ("Operators per shift", result["operators_per_shift"], "operators"),
                ("Total operating operators", result["total_operators"], "operators"),
                ("Operating Labor", result["OL"], "$/yr"),
                ("Maintenance Supplies", result["MS"], "$/yr"),
                ("Direct Manufacturing Cost (DMC)", result["DMC"], "$/yr"),
                ("Fixed Manufacturing Cost (FMC)", result["FMC"], "$/yr"),
                ("General Expenses (GE)", result["GE"], "$/yr"),
                ("Cost of Manufacturing (COM)", result["COM"], "$/yr"),
                ("Cost per Liter (CoPv)", result["CoPv"], "$/L"),
            ]),
            ("Profitability", [
                ("Net Present Worth (Non-Discounted)", result["NNPW"], "$"),
                ("Net Present Worth (Discounted)", result["DNPW"], "$"),
                ("Project Payback (undiscounted, from year 0)", result["NPBP"], "years"),
                ("Project Payback (discounted, from year 0)", result["DPBP"], "years"),
                ("Simple annual after-tax ROI (net profit / TCI)", result["ROI"], "%"),
                ("Internal Rate of Return (IRR)", result["IRR"], "%"),
                ("Benefit-Cost Ratio (undiscounted)", result["B_C_ratio_NCF"], ""),
                ("Benefit-Cost Ratio (discounted)", result["B_C_ratio_DCF"], ""),
            ]),
        ]

        for title, rows in cards:
            lf = tk.LabelFrame(self.res_frame, text=title, font=("TkDefaultFont", 10, "bold"))
            lf.pack(fill=tk.X, padx=6, pady=5)
            inner = tk.Frame(lf)
            inner.pack(fill=tk.X, padx=6, pady=4)
            for i, (label, val, unit) in enumerate(rows):
                r = i // 2
                c = (i % 2) * 2
                tk.Label(inner, text=label, anchor='w', width=42).grid(
                    row=r, column=c, sticky='w', padx=6, pady=2)
                display_val = ("Not reached" if label.startswith("Payback") else "Not defined") if val is None else f"{val:,.4g}"
                tk.Label(inner, text=f"{display_val} {unit}".strip(),
                         anchor='w', width=22, font=("TkDefaultFont", 10, "bold")).grid(
                    row=r, column=c + 1, sticky='w', padx=6, pady=2)

        self._bind_scroll_wheel(self.res_frame, self.results_canvas)

        # Chart
        for w in self.chart_frame.winfo_children():
            w.destroy()
        self._draw_chart(result)

    def _draw_chart(self, result: Dict[str, Any]):
        fig = plt.Figure(figsize=(8, 4.5), dpi=100)
        ax = fig.add_subplot(111)
        y = result["years"]
        ax.plot(y, result["NCF"], label='NCF', marker='o')
        ax.plot(y, result["CNCF"], label='CNCF', marker='s')
        ax.plot(y, result["DCF"], label='DCF', marker='^')
        ax.plot(y, result["CDCF"], label='CDCF', marker='d')
        ax.set_xlabel('Year')
        ax.set_ylabel('Cash Flow ($)')
        ax.set_title('Cash Flow Diagram for Plant')
        ax.legend()
        ax.grid(True, alpha=0.3)
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        toolbar = NavigationToolbar2Tk(canvas, self.chart_frame)
        toolbar.update()
        toolbar.pack(side=tk.BOTTOM, fill=tk.X)

    def _save_results_csv(self):
        if not self.calc_result:
            messagebox.showwarning("Info", "No results to save.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save Results",
            initialdir=self.last_folder,
        )
        if not path:
            return
        try:
            r = self.calc_result
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Project title", r.get("inputs", {}).get("Project title", "Chemical Process Plant TEA Study")])
                w.writerow(["Project summary", r.get("inputs", {}).get("Project summary", "Preliminary techno-economic analysis of a modeled chemical process plant.")])
                w.writerow(["TOPCEE Results", "Amounts reported in USD unless stated otherwise"])
                w.writerow([])
                w.writerow(["OPERATING LABOR STAFFING"])
                w.writerow(["Metric", "Value"])
                w.writerow(["Operators required per shift", r["operators_per_shift"]])
                w.writerow(["Total operators", r["total_operators"]])
                w.writerow(["Shifts per day", r["inputs"].get("Number of shifts per day", "3")])
                w.writerow(["Operating days per year", r["inputs"].get("Operating days per year", "330")])
                w.writerow(["Operator wage (local currency/hour)", r["operator_hourly_wage_local"]])
                w.writerow(["Annual operating labor cost (USD/year)", r["OL"]])
                w.writerow([])
                w.writerow(["CAPEX BREAKDOWN"])
                w.writerow(["Level", "Row Type", "Item", "Amount ($)"])
                capex_items = [
                    (0, "Group", "Direct plant cost components", ""),
                    (1, "Detail", "Purchased equipment cost (PCE)", r["PCE"]),
                    (1, "Detail", "Equipment installation", r["CAPEX_installations"]),
                    (1, "Detail", "Instrumentation and control", r["CAPEX_instrumentation"]),
                    (1, "Detail", "Piping", r["CAPEX_piping"]),
                    (1, "Detail", "Electrical", r["CAPEX_electrical"]),
                    (1, "Detail", "Buildings and structures", r["CAPEX_buildings"]),
                    (1, "Detail", "Yard improvements", r["CAPEX_yard"]),
                    (1, "Detail", "Service facilities", r["CAPEX_services"]),
                    (0, "Subtotal", "Direct Plant Cost (DPC)", r["DPC"]),
                    (0, "Group", "Indirect plant cost components", ""),
                    (1, "Detail", "Engineering and supervision", r["CAPEX_engineering"]),
                    (1, "Detail", "Construction expenses", r["CAPEX_construction"]),
                    (0, "Subtotal", "Indirect Plant Cost (IPC)", r["IPC"]),
                    (0, "Subtotal", "Total Direct & Indirect Plant Cost", r["TPC"]),
                    (1, "Detail", "Contractor's fee and legal expenses", r["contractor_fee"]),
                    (1, "Detail", "Contingency", r["contingency"]),
                    (0, "Subtotal", "Fixed Capital Investment (FCI)", r["FCI"]),
                    (1, "Detail", "Working capital and land", r["WC"]),
                    (0, "Total", "Total Capital Investment (TCI)", r["TCI"]),
                ]
                for level, kind, label, value in capex_items:
                    w.writerow([level, kind, f'{"  " * level}{label}', value])
                w.writerow([])
                w.writerow(["OPEX / COST OF MANUFACTURING BREAKDOWN"])
                w.writerow(["Level", "Row Type", "Item", "Amount ($)"])
                opex_items = [
                    (0, "Group", "Direct manufacturing cost components", ""),
                    (1, "Subtotal", "Raw material cost", r["RM"]),
                ]
                opex_items.extend(
                    (2, "Detail", f"Raw material: {name}", amount)
                    for name, amount in r["raw_material_cost_items"]
                )
                opex_items.append((1, "Subtotal", "Process consumables cost",
                                   r["process_consumable_cost_total"]))
                opex_items.extend(
                    (2, "Detail", f"Process consumable: {name}", amount)
                    for name, amount in r["process_consumable_cost_items"]
                )
                opex_items.append(
                    (1, "Subtotal", "Utilities cost", r["UC"]),
                )
                opex_items.extend(
                    (2, "Detail", f"Utility: {name}", amount)
                    for name, amount in r["utility_cost_items"]
                )
                opex_items.extend([
                    (1, "Detail", "Operating labor", r["OL"]),
                    (1, "Detail", "Operating supervision", r["OS"]),
                    (1, "Detail", "Maintenance supplies", r["MS"]),
                    (1, "Detail", "Operating supplies", r["OSUP"]),
                    (1, "Detail", "Maintenance labor and supervision", r["MLS"]),
                    (1, "Detail", "Laboratory charges", r["LAB"]),
                    (1, "Detail", "Patents and royalties", r["PR"]),
                    (0, "Subtotal", "Direct Manufacturing Cost (DMC)", r["DMC"]),
                    (0, "Group", "Fixed manufacturing cost components", ""),
                    (1, "Detail", "Maintenance and operating overhead", r["maintenance_overhead"]),
                    (1, "Detail", "Depreciation", r["depreciation"]),
                    (1, "Detail", "Local taxes", r["local_taxes"]),
                    (1, "Detail", "Insurance", r["insurance"]),
                    (0, "Subtotal", "Fixed Manufacturing Cost (FMC)", r["FMC"]),
                    (0, "Group", "General expense components", ""),
                    (1, "Detail", "Administrative costs", r["admin_costs"]),
                    (1, "Detail", "Distribution and selling costs", r["distribution_costs"]),
                    (1, "Detail", "Research and development", r["research_costs"]),
                    (0, "Subtotal", "General Expenses (GE)", r["GE"]),
                    (0, "Total", "Cost of Manufacturing (COM)", r["COM"]),
                ])
                for level, kind, label, value in opex_items:
                    w.writerow([level, kind, f'{"  " * level}{label}', value])
                w.writerow([])
                w.writerow(["PROFITABILITY METRICS"])
                w.writerow(["Metric", "Value"])
                for k in ["NNPW", "DNPW", "NPBP", "DPBP", "ROI", "IRR",
                          "B_C_ratio_NCF", "B_C_ratio_DCF"]:
                    value = r.get(k, "")
                    if value is None:
                        value = "Not reached" if k in ("NPBP", "DPBP") else "Not defined"
                    w.writerow([k, value])
                w.writerow([])
                w.writerow(["Year", "NCF", "CNCF", "DCF", "CDCF"])
                for i, yr in enumerate(r["years"]):
                    w.writerow([yr, r["NCF"][i], r["CNCF"][i], r["DCF"][i], r["CDCF"][i]])
            self.last_folder = os.path.dirname(path)
            save_config({"last_folder": self.last_folder, "db_path": self.db_path})
            messagebox.showinfo("Saved", f"Results saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _go_results(self):
        self.notebook.select(self.tab_results)

    # ---------------------------------------------- Report Generation ------
    def _browse_report_folder(self):
        path = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=self.report_folder_var.get() or self.last_folder,
        )
        if path:
            self.report_folder_var.set(path)

    def _remember_folder(self):
        folder = self.report_folder_var.get().strip()
        if not folder:
            messagebox.showwarning("Info", "Please select a folder first.")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("Error", "Folder does not exist.")
            return
        self.last_folder = folder
        save_config({"last_folder": self.last_folder, "db_path": self.db_path})
        messagebox.showinfo("Remembered", f"Output folder remembered:\n{folder}")

    def _open_output_folder(self):
        folder = self.report_folder_var.get().strip() or self.last_folder
        if os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showwarning("Info", "Folder does not exist.")

    def _generate_reports(self):
        if not self.calc_result:
            messagebox.showwarning("No Results",
                                   "Please run 'Calculate Profitability' first before generating reports.")
            return

        folder = self.report_folder_var.get().strip() or self.last_folder
        if not os.path.isdir(folder):
            messagebox.showerror("Error", "Output folder does not exist.")
            return

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = f"TOPCEE_Report_{timestamp}"
        saved = []
        log_lines = []

        r = self.calc_result
        inputs = r.get("inputs", {})
        report_inputs = dict(inputs)

        def nonzero_input(key):
            try:
                return float(str(report_inputs.get(key, "0")).replace(",", "").strip() or 0) != 0
            except (TypeError, ValueError):
                return False

        for i in range(1, 11):
            raw_name = f"Raw material {i} name"
            raw_qty = f"Raw material {i} quantity (unit/h)"
            raw_price = f"Raw material {i} cost (local/unit)"
            if not nonzero_input(raw_qty) and not nonzero_input(raw_price):
                for key in (raw_name, raw_qty, f"Raw material {i} rate unit", raw_price,
                            f"Raw material {i} price basis", f"Raw material {i} density (kg/L)"):
                    report_inputs.pop(key, None)

            utility_keys = (
                f"Utility {i} name", f"Utility {i} quantity (unit/h)",
                f"Utility {i} unit", f"Utility {i} cost (local/unit)",
            )
            if not nonzero_input(utility_keys[1]) and not nonzero_input(utility_keys[3]):
                for key in utility_keys:
                    report_inputs.pop(key, None)

        for i in range(1, 11):
            consumable_qty = f"Process consumable {i} quantity"
            consumable_price = f"Process consumable {i} cost (local/unit)"
            if not nonzero_input(consumable_qty) and not nonzero_input(consumable_price):
                for key in (
                    f"Process consumable {i} name", f"Process consumable {i} mode",
                    consumable_qty, f"Process consumable {i} unit", consumable_price,
                    f"Process consumable {i} replacement interval (operating h)",
                ):
                    report_inputs.pop(key, None)

        # Build report builder
        builder = CostReportBuilder(
            project_name=(inputs.get("Project title", "").strip() or "Chemical Process Plant TEA Study"),
            prepared_by="TOPCEE v2.0",
            currency="$",
            project_summary=(inputs.get("Project summary", "").strip() or
                             "Preliminary techno-economic analysis of a modeled chemical process plant."),
        )
        builder.set_inputs(report_inputs)
        builder.add_note(
            f"User-entered prices and wages are in {r['currency_code']}; "
            f"the exchange rate is {r['exchange_rate']:,.6g} {r['currency_code']} per USD. "
            "Financial results and equipment costs are reported in USD."
        )
        builder.add_note(
            f"Operating labor estimate: {r['operators_per_shift']} operators per shift; "
            f"{r['total_operators']} total operators for "
            f"{inputs.get('Number of shifts per day', '3')} shifts/day and "
            f"{inputs.get('Operating days per year', '330')} operating days/year. "
            f"Annual operating labor cost is ${r['OL']:,.2f}, using "
            f"{r['operator_hourly_wage_local']:,.2f} {r['currency_code']}/hour and "
            f"{OPERATOR_WORK_HOURS_PER_YEAR:,} paid hours per operator-year."
        )
        builder.set_capex({
            "pce": r["PCE"], "dpc": r["DPC"], "ipc": r["IPC"],
            "installations": r["CAPEX_installations"],
            "instrumentation": r["CAPEX_instrumentation"],
            "piping": r["CAPEX_piping"], "electrical": r["CAPEX_electrical"],
            "buildings": r["CAPEX_buildings"], "yard_improvement": r["CAPEX_yard"],
            "service_facilities": r["CAPEX_services"],
            "engineering_supervision": r["CAPEX_engineering"],
            "construction_expenses": r["CAPEX_construction"],
            "dipc": r["TPC"], "contractor_fee": r["contractor_fee"],
            "contingency": r["contingency"], "fci": r["FCI"],
            "tci": r["TCI"], "wc": r["WC"],
        })
        builder.set_opex({
            "rm": r["RM"], "utilities": r["UC"] - r["PC"],
            "utility_total": r["UC"], "pc": r["PC"], "ol": r["OL"],
            "raw_material_cost_items": r["raw_material_cost_items"],
            "utility_cost_items": r["utility_cost_items"],
            "consumables": r["process_consumable_cost_total"],
            "process_consumable_cost_items": r["process_consumable_cost_items"],
            "os": r["OS"], "ms": r["MS"], "osup": r["OSUP"], "mls": r["MLS"],
            "lab": r["LAB"], "pr": r["PR"],
            "maintenance_overhead": r["maintenance_overhead"],
            "depreciation": r["depreciation"], "local_taxes": r["local_taxes"],
            "insurance": r["insurance"], "admin_costs": r["admin_costs"],
            "distribution_costs": r["distribution_costs"],
            "research_costs": r["research_costs"],
            "dmc": r["DMC"], "fmc": r["FMC"], "ge": r["GE"], "com": r["COM"],
        })
        builder.set_profitability({
            "nnpw": r["NNPW"], "dnpw": r["DNPW"], "npbp": r["NPBP"],
            "dpbp": r["DPBP"], "roi": r["ROI"], "irr": r["IRR"],
            "bcncf": r["B_C_ratio_NCF"], "bcdcf": r["B_C_ratio_DCF"],
            "com": r["COM"], "tci": r["TCI"],
        })
        builder.set_cash_flow(
            r["years"], [float(v) for v in r["NCF"]], [float(v) for v in r["CNCF"]],
            [float(v) for v in r["DCF"]], [float(v) for v in r["CDCF"]],
        )
        builder.set_equipment_list(self.equipment_items)
        if self.equipment_total > 0:
            builder.add_note(f"Total equipment cost from equipment database: $ {self.equipment_total:,.2f}")
        else:
            builder.add_note("Equipment cost entered manually (equipment database not used).")

        # 1. Equipment Cost Schedule (CSV)
        if self.gen_equip_var.get():
            eq_path = os.path.join(folder, f"{base_name}_EquipmentCost.csv")
            try:
                with open(eq_path, "w", newline="", encoding="utf-8") as f:
                    w = csv.writer(f)
                    w.writerow(["TOPCEE Equipment Cost Schedule"])
                    w.writerow(["Project title", inputs.get("Project title", "").strip() or "Chemical Process Plant TEA Study"])
                    w.writerow(["Project summary", inputs.get("Project summary", "").strip() or "Preliminary techno-economic analysis of a modeled chemical process plant."])
                    w.writerow(["Date", datetime.datetime.now().strftime("%d %b %Y %H:%M")])
                    w.writerow(["Total Equipment Cost", f"{self.equipment_total:,.2f}"])
                    w.writerow([])
                    w.writerow(EquipmentItem.header())
                    for it in self.equipment_items:
                        w.writerow(it.to_row())
                    w.writerow(["TOTAL"] + [""] * 17 +
                               [f"{self.equipment_total:,.2f}", ""])
                saved.append(eq_path)
                log_lines.append(f"[OK] Equipment Cost Schedule -> {eq_path}")
            except Exception as e:
                log_lines.append(f"[ERR] Equipment Cost Schedule: {e}")

        # 2. Full TEA Report (HTML + PDF + CSV + Excel)
        if self.gen_full_report_var.get():
            html_path = os.path.join(folder, f"{base_name}_CostReport.html")
            csv_path = os.path.join(folder, f"{base_name}_CostReport.csv")
            try:
                builder.save_html(html_path)
                saved.append(html_path)
                log_lines.append(f"[OK] Full TEA Report (HTML) -> {html_path}")
            except Exception as e:
                log_lines.append(f"[ERR] HTML Report: {e}")

            try:
                builder.save_csv(csv_path)
                saved.append(csv_path)
                log_lines.append(f"[OK] Full TEA Report (CSV) -> {csv_path}")
            except Exception as e:
                log_lines.append(f"[ERR] CSV Report: {e}")

            try:
                xlsx_path = os.path.join(folder, f"{base_name}_CostReport.xlsx")
                builder.save_excel(xlsx_path)
                saved.append(xlsx_path)
                log_lines.append(f"[OK] Full TEA Report (Excel) -> {xlsx_path}")
            except Exception as e:
                log_lines.append(f"[ERR] Excel Report: {e}")

            try:
                pdf_path = os.path.join(folder, f"{base_name}_CostReport.pdf")
                builder.save_pdf(pdf_path)
                saved.append(pdf_path)
                log_lines.append(f"[OK] Full TEA Report (PDF) -> {pdf_path}")
            except Exception as e:
                log_lines.append(f"[ERR] PDF Report: {e}")

        # Update log
        self.report_log.delete("1.0", tk.END)
        self.report_log.insert(tk.END, "\n".join(log_lines))
        self.last_folder = folder
        save_config({"last_folder": self.last_folder, "db_path": self.db_path})

        if saved:
            messagebox.showinfo("Reports Generated",
                                f"{len(saved)} file(s) saved to:\n{folder}\n\n"
                                f"{chr(10).join(saved[:5])}")
        else:
            messagebox.showwarning("Warning", "No reports were generated. Check the log for errors.")

    # ------------------------------------------------------------------ --


def main():
    root = tk.Tk()
    app = TOPCEEApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
