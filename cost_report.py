# -*- coding: utf-8 -*-
"""
Cost Report Generator for TOPCEE

Produces a professional process-economic cost report covering:
- CAPEX breakdown (equipment, direct plant cost, indirect plant cost,
  fixed capital investment, working capital, total capital investment)
- OPEX / Cost of Manufacturing breakdown (raw materials, utilities,
  labor, maintenance, overheads, depreciation, taxes, insurance, GE)
- Profitability results (NPV, IRR, ROI, payback, B/C ratio, cash flows)
- Equipment cost list (from the equipment costing module)

Exports:
- HTML report (browser-printable to PDF)
- CSV report
- Excel workbook (if openpyxl available)

Usage:
    from cost_report import CostReportBuilder
    builder = CostReportBuilder(project_name="My Plant")
    builder.set_capex(...)
    builder.set_opex(...)
    builder.set_profitability(...)
    builder.set_equipment_list(items)
    builder.save_html("report.html")
    builder.save_csv("report.csv")
"""

from __future__ import annotations

import os
import csv
import datetime
import html as html_lib
import textwrap
from typing import List, Dict, Any, Optional

try:
    import openpyxl
    from openpyxl.chart import LineChart, Reference
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False

try:
    from equipment_cost_db import EquipmentItem
    HAS_EQUIPMENT_DB = True
except Exception:
    HAS_EQUIPMENT_DB = False


def _fmt(value, digits=2, prefix="", suffix=""):
    if value is None:
        return "Not defined"
    try:
        v = float(value)
    except (TypeError, ValueError):
        return str(value)
    return f"{prefix}{v:,.{digits}f}{suffix}"


def _money(value, digits=2):
    return _fmt(value, digits=digits, prefix="$ ")


CAPEX_REPORT_LINES = [
    ("group", 0, "Direct plant cost components", None),
    ("detail", 1, "Purchased equipment cost (PCE)", "pce"),
    ("detail", 1, "Equipment installation", "installations"),
    ("detail", 1, "Instrumentation and control", "instrumentation"),
    ("detail", 1, "Piping", "piping"),
    ("detail", 1, "Electrical", "electrical"),
    ("detail", 1, "Buildings and structures", "buildings"),
    ("detail", 1, "Yard improvements", "yard_improvement"),
    ("detail", 1, "Service facilities", "service_facilities"),
    ("subtotal", 0, "Direct Plant Cost (DPC)", "dpc"),
    ("group", 0, "Indirect plant cost components", None),
    ("detail", 1, "Engineering and supervision", "engineering_supervision"),
    ("detail", 1, "Construction expenses", "construction_expenses"),
    ("subtotal", 0, "Indirect Plant Cost (IPC)", "ipc"),
    ("subtotal", 0, "Total Direct & Indirect Plant Cost", "dipc"),
    ("group", 0, "Capital additions", None),
    ("detail", 1, "Contractor's fee and legal expenses", "contractor_fee"),
    ("detail", 1, "Contingency", "contingency"),
    ("subtotal", 0, "Fixed Capital Investment (FCI)", "fci"),
    ("detail", 1, "Working capital and land", "wc"),
    ("total", 0, "Total Capital Investment (TCI)", "tci"),
]

OPEX_REPORT_LINES = [
    ("group", 0, "Direct manufacturing cost components", None),
    ("subtotal", 1, "Raw material cost", "rm"),
    ("dynamic", 2, "", "raw_material_cost_items"),
    ("subtotal", 1, "Process consumables cost", "consumables"),
    ("dynamic", 2, "", "process_consumable_cost_items"),
    ("subtotal", 1, "Utilities cost", "utility_total"),
    ("dynamic", 2, "", "utility_cost_items"),
    ("detail", 1, "Operating labor", "ol"),
    ("detail", 1, "Operating supervision", "os"),
    ("detail", 1, "Maintenance supplies", "ms"),
    ("detail", 1, "Operating supplies", "osup"),
    ("detail", 1, "Maintenance labor and supervision", "mls"),
    ("detail", 1, "Laboratory charges", "lab"),
    ("detail", 1, "Patents and royalties", "pr"),
    ("subtotal", 0, "Direct Manufacturing Cost (DMC)", "dmc"),
    ("group", 0, "Fixed manufacturing cost components", None),
    ("detail", 1, "Maintenance and operating overhead", "maintenance_overhead"),
    ("detail", 1, "Depreciation", "depreciation"),
    ("detail", 1, "Local taxes", "local_taxes"),
    ("detail", 1, "Insurance", "insurance"),
    ("subtotal", 0, "Fixed Manufacturing Cost (FMC)", "fmc"),
    ("group", 0, "General expense components", None),
    ("detail", 1, "Administrative costs", "admin_costs"),
    ("detail", 1, "Distribution and selling costs", "distribution_costs"),
    ("detail", 1, "Research and development", "research_costs"),
    ("subtotal", 0, "General Expenses (GE)", "ge"),
    ("total", 0, "Cost of Manufacturing (COM)", "com"),
]


# ---------------------------------------------------------------------------
# Report builder
# ---------------------------------------------------------------------------

class CostReportBuilder:
    """Assembles a complete process-economic cost report."""

    def __init__(self, project_name: str = "Process Plant",
                 prepared_by: str = "TOPCEE",
                 currency: str = "$",
                 project_summary: str = ""):
        self.project_name = project_name
        self.project_summary = project_summary
        self.prepared_by = prepared_by
        self.currency = currency
        self.date = datetime.datetime.now().strftime("%d %b %Y %H:%M")

        self.inputs: Dict[str, Any] = {}
        self.capex: Dict[str, float] = {}
        self.opex: Dict[str, Any] = {}
        self.profitability: Dict[str, float] = {}
        self.cash_flow_years: List[int] = []
        self.cash_flow_ncf: List[float] = []
        self.cash_flow_cncf: List[float] = []
        self.cash_flow_dcf: List[float] = []
        self.cash_flow_cdcf: List[float] = []
        self.equipment_items: List[Any] = []
        self.notes: List[str] = []

    # ---- Setters -----------------------------------------------------------

    def set_inputs(self, inputs: Dict[str, Any]):
        self.inputs = dict(inputs)

    def set_capex(self, capex: Dict[str, float]):
        self.capex = dict(capex)

    def set_opex(self, opex: Dict[str, Any]):
        self.opex = dict(opex)

    def set_profitability(self, prof: Dict[str, float]):
        self.profitability = dict(prof)

    def set_cash_flow(self, years: List[int], ncf: List[float],
                      cncf: List[float], dcf: List[float], cdcf: List[float]):
        self.cash_flow_years = list(years)
        self.cash_flow_ncf = list(ncf)
        self.cash_flow_cncf = list(cncf)
        self.cash_flow_dcf = list(dcf)
        self.cash_flow_cdcf = list(cdcf)

    def set_equipment_list(self, items: List[Any]):
        self.equipment_items = list(items)

    def add_note(self, text: str):
        self.notes.append(text)

    # ---- Helpers -----------------------------------------------------------

    def _opex_report_lines(self):
        """Return OPEX rows with named raw-material, consumable, and utility details expanded."""
        lines = []
        for kind, level, label, key in OPEX_REPORT_LINES:
            if kind != "dynamic":
                lines.append((kind, level, label, self.opex.get(key, 0.0) if key else ""))
                continue
            for item_name, amount in self.opex.get(key, []) or []:
                prefix = {
                    "raw_material_cost_items": "Raw material",
                    "process_consumable_cost_items": "Process consumable",
                    "utility_cost_items": "Utility",
                }.get(key, "Item")
                lines.append(("detail", level, f"{prefix}: {item_name}", amount))
        return lines

    def _equipment_rows(self) -> List[List[str]]:
        rows = []
        if not self.equipment_items:
            return rows
        for it in self.equipment_items:
            if hasattr(it, "to_row"):
                rows.append([str(v) for v in it.to_row()])
            else:
                rows.append([
                    str(getattr(it, "name", "")),
                    str(getattr(it, "category", "")),
                    str(getattr(it, "equipment_type", "")),
                    str(getattr(it, "capacity", "")),
                    str(getattr(it, "total_cost", lambda: 0)()),
                ])
        return rows

    def _total_equipment(self) -> float:
        total = 0.0
        for it in self.equipment_items:
            try:
                total += float(it.total_cost()) if hasattr(it, "total_cost") else float(getattr(it, "total_cost", 0))
            except Exception:
                pass
        return total

    # ---- HTML report -------------------------------------------------------

    def build_html(self) -> str:
        money = lambda v: _money(v)
        esc = lambda s: html_lib.escape(str(s), quote=True)

        total_eq = self._total_equipment()
        eq_rows = self._equipment_rows()

        capex_rows = []
        for kind, level, label, key in CAPEX_REPORT_LINES:
            value = self.capex.get(key, 0.0) if key else ""
            capex_rows.append(
                f'<tr class="{kind}"><td class="level-{level}">{esc(label)}</td>'
                f'<td>{money(value) if key else ""}</td></tr>'
            )

        opex_rows = []
        for kind, level, label, value in self._opex_report_lines():
            opex_rows.append(
                f'<tr class="{kind}"><td class="level-{level}">{esc(label)}</td>'
                f'<td>{money(value) if value != "" else ""}</td></tr>'
            )

        # Profitability rows
        prof_rows = []
        for key, label, digits in [
            ("nnpw", "Non-Discounted Net Present Worth (NNPW)", 2),
            ("dnpw", "Discounted Net Present Worth (DNPW)", 2),
            ("npbp", "Project Payback (undiscounted, from year 0; years)", 2),
            ("dpbp", "Project Payback (discounted, from year 0; years)", 2),
            ("roi", "Simple Annual After-Tax ROI (net profit / TCI, %)", 2),
            ("irr", "Internal Rate of Return (IRR, %)", 2),
            ("bcncf", "Benefit-Cost Ratio (undiscounted)", 3),
            ("bcdcf", "Benefit-Cost Ratio (discounted at selected rate)", 3),
            ("com", "Cost of Manufacturing (COM)", 2),
            ("tci", "Total Capital Investment (TCI)", 2),
        ]:
            val = self.profitability.get(key, 0.0)
            suffix = "%" if key in ("roi", "irr") else ""
            if val is None and key in ("npbp", "dpbp"):
                formatted = "Not reached"
            elif val is None:
                formatted = "Not defined"
            else:
                formatted = money(val) if key in ("nnpw", "dnpw", "com", "tci") else _fmt(
                    val, digits=digits, suffix=suffix
                )
            prof_rows.append(f"<tr><td>{esc(label)}</td><td>{formatted}</td></tr>")

        # Equipment rows HTML
        eq_html = ""
        if eq_rows:
            header = EquipmentItem.header() if HAS_EQUIPMENT_DB else tuple(
                f"Column {i + 1}" for i in range(len(eq_rows[0]))
            )
            eq_html = "<h3>Equipment Cost List</h3>"
            eq_html += '<table><thead><tr>' + "".join(f"<th>{esc(h)}</th>" for h in header) + "</tr></thead><tbody>"
            for row in eq_rows:
                eq_html += "<tr>" + "".join(f"<td>{esc(v)}</td>" for v in row) + "</tr>"
            eq_html += f"<tr class='total'><td colspan='{len(header) - 2}'><b>TOTAL EQUIPMENT COST</b></td><td><b>{money(total_eq)}</b></td><td></td></tr>"
            eq_html += "</tbody></table>"
        else:
            eq_html = "<p><i>No equipment items included in this report.</i></p>"

        # Cash flow table HTML
        cf_html = ""
        if self.cash_flow_years:
            cf_html = "<h3>Cash Flow Summary</h3>"
            cf_html += '<table><thead><tr><th>Year</th><th>NCF ($)</th><th>CNCF ($)</th><th>DCF ($)</th><th>CDCF ($)</th></tr></thead><tbody>'
            for i, yr in enumerate(self.cash_flow_years):
                ncf = self.cash_flow_ncf[i] if i < len(self.cash_flow_ncf) else 0
                cncf = self.cash_flow_cncf[i] if i < len(self.cash_flow_cncf) else 0
                dcf = self.cash_flow_dcf[i] if i < len(self.cash_flow_dcf) else 0
                cdcf = self.cash_flow_cdcf[i] if i < len(self.cash_flow_cdcf) else 0
                cf_html += (
                    f"<tr><td>{yr}</td><td>{money(ncf)}</td><td>{money(cncf)}</td>"
                    f"<td>{money(dcf)}</td><td>{money(cdcf)}</td></tr>"
                )
            cf_html += "</tbody></table>"
        else:
            cf_html = "<p><i>Cash flow data not available.</i></p>"

        # Inputs table
        inp_rows = ""
        for k, v in self.inputs.items():
            inp_rows += f"<tr><td>{esc(str(k))}</td><td>{esc(str(v))}</td></tr>"

        # Notes
        notes_html = ""
        if self.notes:
            notes_html = "<h3>Notes</h3><ul>" + "".join(f"<li>{esc(n)}</li>" for n in self.notes) + "</ul>"

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Cost Report - {esc(self.project_name)}</title>
<style>
  body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 30px; color: #222; }}
  h1 {{ color: #1a3c5a; border-bottom: 2px solid #1a3c5a; padding-bottom: 8px; }}
  h2 {{ color: #2c5f8a; margin-top: 24px; }}
  h3 {{ color: #333; margin-top: 18px; }}
  table {{ border-collapse: collapse; width: 100%; margin: 10px 0 20px 0; }}
  th, td {{ border: 1px solid #b9c6d4; padding: 6px 9px; text-align: left; font-size: 13px; }}
  th {{ background: #eef2f7; font-weight: 600; color: #1a3c5a; }}
  tr:nth-child(even) {{ background: #f7f9fb; }}
  tr.group td {{ background: #eaf0f6; color: #1a3c5a; font-weight: 700; }}
  tr.subtotal td {{ background: #edf5fa; border-top: 2px solid #8fa9c0; font-weight: 700; }}
  tr.total td {{ background: #d4e5f2; border-top: 2px solid #1a3c5a; font-weight: 700; }}
  td.level-1 {{ padding-left: 24px; }}
  td.level-2 {{ padding-left: 42px; }}
  .total-row td {{ background: #dbe7f3; font-weight: bold; }}
  .meta {{ display: flex; gap: 30px; margin-bottom: 20px; font-size: 13px; color: #555; }}
  .meta div {{ flex: 1; }}
  .kpi {{ display: flex; gap: 16px; margin: 18px 0; }}
  .kpi-box {{ border: 1px solid #b9c6d4; border-radius: 6px; padding: 12px 16px; flex: 1; background: #f7f9fb; }}
  .kpi-box .v {{ font-size: 20px; font-weight: bold; color: #1a3c5a; }}
  .kpi-box .l {{ font-size: 12px; color: #666; }}
  .note {{ background: #fff8e1; border-left: 4px solid #f0c419; padding: 10px; margin: 14px 0; }}
  .footer {{ margin-top: 40px; border-top: 1px solid #b9c6d4; padding-top: 10px; font-size: 11px; color: #777; }}
  @media print {{ body {{ margin: 0; }} h1 {{ border-bottom: none; }} }}
</style>
</head>
<body>

<div class="meta">
  <div><b>Project:</b> {esc(self.project_name)}</div>
  <div><b>Prepared by:</b> {esc(self.prepared_by)}</div>
  <div><b>Date:</b> {esc(self.date)}</div>
  <div><b>Currency:</b> {esc(self.currency)}</div>
</div>

<h1>Process Plant Cost &amp; Economic Analysis Report</h1>
<h2>{esc(self.project_name)}</h2>
<p>{esc(self.project_summary)}</p>

<div class="kpi">
  <div class="kpi-box"><div class="v">{money(total_eq)}</div><div class="l">Total Equipment Cost</div></div>
  <div class="kpi-box"><div class="v">{money(self.capex.get('tci', 0))}</div><div class="l">Total Capital Investment (TCI)</div></div>
  <div class="kpi-box"><div class="v">{money(self.opex.get('com', 0))}</div><div class="l">Cost of Manufacturing (COM)</div></div>
  <div class="kpi-box"><div class="v">{money(self.profitability.get('dnpw', 0))}</div><div class="l">Discounted NPV</div></div>
  <div class="kpi-box"><div class="v">{_fmt(self.profitability.get('irr', 0), suffix='%')}</div><div class="l">Internal Rate of Return</div></div>
  <div class="kpi-box"><div class="v">{_fmt(self.profitability.get('roi', 0), suffix='%')}</div><div class="l">Simple Annual After-Tax ROI (net profit / TCI)</div></div>
</div>

<h2>1. CAPEX / Total Capital Investment Breakdown</h2>
<table><thead><tr><th>Item</th><th>Amount ({esc(self.currency)})</th></tr></thead><tbody>
{chr(10).join(capex_rows)}
</tbody></table>

<h2>2. OPEX / Cost of Manufacturing Breakdown</h2>
<table><thead><tr><th>Item</th><th>Amount ({esc(self.currency)})</th></tr></thead><tbody>
{chr(10).join(opex_rows)}
</tbody></table>

<h2>3. Profitability &amp; Investment Metrics</h2>
<table><thead><tr><th>Metric</th><th>Value</th></tr></thead><tbody>
{chr(10).join(prof_rows)}
</tbody></table>

{cf_html}

<h2>4. Input Data Summary</h2>
<table><thead><tr><th>Parameter</th><th>Value</th></tr></thead><tbody>
{inp_rows}
</tbody></table>

{eq_html}

{notes_html}

<div class="footer">
  Generated by TOPCEE (Toyese Oyegoke Process Cost &amp; Economic Evaluator).
  This report is for engineering-economic estimation purposes only.
</div>

</body>
</html>"""
        return html

    # ---- Save methods ------------------------------------------------------

    def save_html(self, path: str):
        html = self.build_html()
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        return path

    def save_csv(self, path: str):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["TOPCEE Cost Report", self.project_name])
            w.writerow(["Project summary", self.project_summary])
            w.writerow(["Prepared by", self.prepared_by])
            w.writerow(["Date", self.date])
            w.writerow(["Currency", self.currency])
            w.writerow([])
            w.writerow(["CAPEX BREAKDOWN"])
            w.writerow(["Level", "Row Type", "Item", "Amount ($)"])
            for kind, level, label, key in CAPEX_REPORT_LINES:
                value = self.capex.get(key, "") if key else ""
                w.writerow([level, kind.title(), f'{"  " * level}{label}', value])
            w.writerow([])
            w.writerow(["OPEX / COM BREAKDOWN"])
            w.writerow(["Level", "Row Type", "Item", "Amount ($)"])
            for kind, level, label, value in self._opex_report_lines():
                w.writerow([level, kind.title(), f'{"  " * level}{label}', value])
            w.writerow([])
            w.writerow(["PROFITABILITY METRICS"])
            for key, label in [
                ("nnpw", "NNPW"), ("dnpw", "DNPW"), ("npbp", "Project Payback (undiscounted, from year 0; years)"),
                ("dpbp", "Project Payback (discounted, from year 0; years)"), ("roi", "Simple annual after-tax ROI (net profit / TCI, %)"), ("irr", "IRR (%)"),
                ("bcncf", "Benefit-Cost Ratio (undiscounted)"),
                ("bcdcf", "Benefit-Cost Ratio (discounted at selected rate)"),
                ("com", "Cost of Manufacturing (COM)"), ("tci", "Total Capital Investment (TCI)"),
            ]:
                value = self.profitability.get(key, 0.0)
                if value is None:
                    value = "Not reached" if key in ("npbp", "dpbp") else "Not defined"
                w.writerow([label, value])
            w.writerow([])
            w.writerow(["CASH FLOW ANALYSIS"])
            w.writerow(["Year", "NCF ($)", "CNCF ($)", "DCF ($)", "CDCF ($)"])
            for i, year in enumerate(self.cash_flow_years):
                w.writerow([year,
                            self.cash_flow_ncf[i] if i < len(self.cash_flow_ncf) else 0.0,
                            self.cash_flow_cncf[i] if i < len(self.cash_flow_cncf) else 0.0,
                            self.cash_flow_dcf[i] if i < len(self.cash_flow_dcf) else 0.0,
                            self.cash_flow_cdcf[i] if i < len(self.cash_flow_cdcf) else 0.0])
            w.writerow([])
            w.writerow(["EQUIPMENT COST LIST"])
            eq_headers = EquipmentItem.header() if HAS_EQUIPMENT_DB else [
                "Name", "Process ID", "Category", "Type", "Capacity", "Total Cost ($)"
            ]
            w.writerow(eq_headers)
            for it in self.equipment_items:
                w.writerow(it.to_row() if hasattr(it, "to_row") else [
                    getattr(it, "name", ""), getattr(it, "process_id", ""),
                    getattr(it, "category", ""), getattr(it, "equipment_type", ""),
                    getattr(it, "capacity", ""),
                    it.total_cost() if hasattr(it, "total_cost") else 0.0,
                ])
            w.writerow(["TOTAL"] + [""] * (len(eq_headers) - 3) + [self._total_equipment(), ""])
            if self.notes:
                w.writerow([])
                w.writerow(["REPORT NOTES"])
                for note in self.notes:
                    w.writerow([note])
        return path

    def save_excel(self, path: str):
        if not HAS_OPENPYXL:
            raise RuntimeError("openpyxl required for Excel export. pip install openpyxl")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Cost Report"
        ws.sheet_properties.outlinePr.summaryBelow = True

        bold = Font(bold=True, color="1A3C5A")
        hdr_fill = PatternFill("solid", fgColor="EEF2F7")
        thin = Side(style="thin", color="B9C6D4")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        r = 1
        ws.cell(row=r, column=1, value="TOPCEE Cost Report").font = Font(bold=True, size=14)
        r += 1
        for k, v in [("Project", self.project_name), ("Project summary", self.project_summary),
                     ("Prepared by", self.prepared_by),
                     ("Date", self.date), ("Currency", self.currency)]:
            ws.cell(row=r, column=1, value=k).font = bold
            ws.cell(row=r, column=2, value=v)
            r += 1
        r += 1

        def section(title):
            nonlocal r
            c = ws.cell(row=r, column=1, value=title); c.font = Font(bold=True, size=12, color="2C5F8A")
            r += 1

        def table(headers, rows, total_row=None, row_kinds=None,
                  row_levels=None, money_columns=None):
            nonlocal r
            for j, h in enumerate(headers, 1):
                c = ws.cell(row=r, column=j, value=h); c.font = bold; c.fill = hdr_fill; c.border = border
            r += 1
            for i, row in enumerate(rows):
                kind = row_kinds[i] if row_kinds and i < len(row_kinds) else "detail"
                level = row_levels[i] if row_levels and i < len(row_levels) else 0
                ws.row_dimensions[r].outlineLevel = min(level, 7)
                for j, v in enumerate(row, 1):
                    c = ws.cell(row=r, column=j, value=v)
                    c.border = border
                    if kind == "group":
                        c.fill = PatternFill("solid", fgColor="EAF0F6")
                        c.font = bold
                    elif kind == "subtotal":
                        c.fill = PatternFill("solid", fgColor="EDF5FA")
                        c.font = bold
                    elif kind == "total":
                        c.fill = PatternFill("solid", fgColor="164E78")
                        c.font = Font(bold=True, color="FFFFFF")
                    if j == 1 and level:
                        c.alignment = Alignment(indent=level)
                    if money_columns and j in money_columns and isinstance(v, (int, float)):
                        c.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);-'
                r += 1
            if total_row:
                for j, v in enumerate(total_row, 1):
                    c = ws.cell(row=r, column=j, value=v); c.font = bold; c.border = border
                r += 1
            r += 1

        # CAPEX
        section("1. CAPEX / TCI Breakdown")
        capex_rows = [[label, self.capex.get(key, "") if key else ""]
                      for _kind, _level, label, key in CAPEX_REPORT_LINES]
        capex_data_start = r + 1
        table(["CAPEX item", "Amount ($)"], capex_rows,
              row_kinds=[item[0] for item in CAPEX_REPORT_LINES],
              row_levels=[item[1] for item in CAPEX_REPORT_LINES], money_columns={2})
        capex_excel_row = lambda index: capex_data_start + index
        capex_formulas = {
            9: f"=SUM(B{capex_excel_row(1)}:B{capex_excel_row(8)})",
            13: f"=SUM(B{capex_excel_row(11)}:B{capex_excel_row(12)})",
            14: f"=SUM(B{capex_excel_row(9)},B{capex_excel_row(13)})",
            18: f"=SUM(B{capex_excel_row(14)},B{capex_excel_row(16)}:B{capex_excel_row(17)})",
            20: f"=SUM(B{capex_excel_row(18)}:B{capex_excel_row(19)})",
        }
        for index, formula in capex_formulas.items():
            cell = ws.cell(row=capex_excel_row(index), column=2, value=formula)
            cell.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);-'

        # OPEX
        section("2. OPEX / COM Breakdown")
        opex_lines = self._opex_report_lines()
        opex_rows = [[label, value] for _kind, _level, label, value in opex_lines]
        opex_data_start = r + 1
        table(["OPEX item", "Amount ($)"], opex_rows,
              row_kinds=[item[0] for item in opex_lines],
              row_levels=[item[1] for item in opex_lines], money_columns={2})
        opex_excel_row = lambda index: opex_data_start + index
        opex_label_rows = {label: opex_excel_row(i)
                           for i, (_kind, _level, label, _value) in enumerate(opex_lines)}
        raw_detail_indices = [i for i, (_kind, _level, label, _value) in enumerate(opex_lines)
                              if label.startswith("Raw material:")]
        consumable_detail_indices = [i for i, (_kind, _level, label, _value) in enumerate(opex_lines)
                                     if label.startswith("Process consumable:")]
        utility_detail_indices = [i for i, (_kind, _level, label, _value) in enumerate(opex_lines)
                                  if label.startswith("Utility:")]

        def subtotal_formula(detail_indices):
            if not detail_indices:
                return "=0"
            first, last = opex_excel_row(detail_indices[0]), opex_excel_row(detail_indices[-1])
            return f"=SUM(B{first}:B{last})"

        raw_total_row = opex_label_rows["Raw material cost"]
        consumables_total_row = opex_label_rows["Process consumables cost"]
        utility_total_row = opex_label_rows["Utilities cost"]
        dmc_row = opex_label_rows["Direct Manufacturing Cost (DMC)"]
        fmc_row = opex_label_rows["Fixed Manufacturing Cost (FMC)"]
        ge_row = opex_label_rows["General Expenses (GE)"]
        com_row = opex_label_rows["Cost of Manufacturing (COM)"]
        opex_formulas = {
            raw_total_row: subtotal_formula(raw_detail_indices),
            consumables_total_row: subtotal_formula(consumable_detail_indices),
            utility_total_row: subtotal_formula(utility_detail_indices),
            dmc_row: (f"=SUM(B{raw_total_row},B{consumables_total_row},B{utility_total_row},"
                      f"B{opex_label_rows['Operating labor']}:B{opex_label_rows['Patents and royalties']})"),
            fmc_row: (f"=SUM(B{opex_label_rows['Maintenance and operating overhead']}"
                      f":B{opex_label_rows['Insurance']})"),
            ge_row: (f"=SUM(B{opex_label_rows['Administrative costs']}"
                     f":B{opex_label_rows['Research and development']})"),
            com_row: f"=SUM(B{dmc_row},B{fmc_row},B{ge_row})",
        }
        for row, formula in opex_formulas.items():
            cell = ws.cell(row=row, column=2, value=formula)
            cell.number_format = '"$"#,##0.00;[Red]("$"#,##0.00);-'

        # Profitability
        section("3. Profitability Metrics")
        prof_rows = []
        for key, label in [
            ("nnpw", "NNPW"), ("dnpw", "DNPW"), ("npbp", "Project Payback (undiscounted, from year 0; years)"),
            ("dpbp", "Project Payback (discounted, from year 0; years)"), ("roi", "Simple annual after-tax ROI (net profit / TCI, %)"), ("irr", "IRR (%)"),
            ("bcncf", "Benefit-Cost Ratio (undiscounted)"),
            ("bcdcf", "Benefit-Cost Ratio (discounted at selected rate)"),
        ]:
            value = self.profitability.get(key, 0.0)
            if value is None:
                value = "Not reached" if key in ("npbp", "dpbp") else "Not defined"
            prof_rows.append([label, value])
        profitability_header_row = r
        table(["Metric", "Value"], prof_rows)
        for metric_index, (key, _label) in enumerate([
            ("nnpw", "NNPW"), ("dnpw", "DNPW"), ("npbp", "Project Payback (undiscounted, from year 0; years)"),
            ("dpbp", "Project Payback (discounted, from year 0; years)"), ("roi", "Simple annual after-tax ROI (net profit / TCI, %)"), ("irr", "IRR (%)"),
            ("bcncf", "Benefit-Cost Ratio (undiscounted)"),
            ("bcdcf", "Benefit-Cost Ratio (discounted at selected rate)"),
        ], start=1):
            if key in ("nnpw", "dnpw"):
                ws.cell(row=profitability_header_row + metric_index, column=2).number_format = (
                    '"$"#,##0.00;[Red]("$"#,##0.00);-'
                )

        # Equipment
        section("4. Equipment Cost List")
        eq_headers = EquipmentItem.header() if HAS_EQUIPMENT_DB else [
            "Name", "Process ID", "Category", "Type", "Capacity", "Total Cost ($)"
        ]
        eq_rows = []
        for it in self.equipment_items:
            eq_rows.append(it.to_row() if hasattr(it, "to_row") else [
                getattr(it, "name", ""), getattr(it, "process_id", ""),
                getattr(it, "category", ""), getattr(it, "equipment_type", ""),
                getattr(it, "capacity", ""),
                it.total_cost() if hasattr(it, "total_cost") else 0.0,
            ])
        table(eq_headers, eq_rows, total_row=["TOTAL"] + [""] * (len(eq_headers) - 3) + [self._total_equipment(), ""])

        # Excel workbook includes the same cash-flow data and diagram shown in the app.
        section("5. Cash Flow Diagram and Data")
        cash_flow_rows = []
        for i, year in enumerate(self.cash_flow_years):
            cash_flow_rows.append([
                year,
                self.cash_flow_ncf[i] if i < len(self.cash_flow_ncf) else 0.0,
                self.cash_flow_cncf[i] if i < len(self.cash_flow_cncf) else 0.0,
                self.cash_flow_dcf[i] if i < len(self.cash_flow_dcf) else 0.0,
                self.cash_flow_cdcf[i] if i < len(self.cash_flow_cdcf) else 0.0,
            ])
        cash_flow_header_row = r
        table(["Year", "NCF ($)", "CNCF ($)", "DCF ($)", "CDCF ($)"],
              cash_flow_rows, money_columns={2, 3, 4, 5})
        if cash_flow_rows:
            chart = LineChart()
            chart.title = "Cash Flow Diagram for Plant"
            chart.style = 13
            chart.y_axis.title = "Cash Flow (USD)"
            chart.x_axis.title = "Year"
            chart.height = 9
            chart.width = 18
            data = Reference(ws, min_col=2, max_col=5, min_row=cash_flow_header_row,
                             max_row=cash_flow_header_row + len(cash_flow_rows))
            categories = Reference(ws, min_col=1, min_row=cash_flow_header_row + 1,
                                   max_row=cash_flow_header_row + len(cash_flow_rows))
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(categories)
            chart.legend.position = "b"
            ws.add_chart(chart, f"G{cash_flow_header_row}")

        if self.notes:
            section("6. Report Notes")
            table(["Note"], [[note] for note in self.notes])

        # Auto column widths
        for col in ws.columns:
            max_len = 0
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 40)

        wb.save(path)
        return path

    def save_pdf(self, path: str):
        """Write a paginated PDF version of the report using matplotlib (a core TOPCEE dependency)."""
        try:
            from matplotlib.backends.backend_pdf import PdfPages
            from matplotlib.figure import Figure
        except Exception as exc:
            raise RuntimeError("Matplotlib is required for PDF export.") from exc

        with PdfPages(path) as pdf:
            info = pdf.infodict()
            info["Title"] = f"TOPCEE Report: {self.project_name}"
            info["Author"] = self.prepared_by
            info["Subject"] = self.project_summary

            def render_page(title, rows, headers=("Item", "Value"), first_page=False):
                chunk_size = 26 if first_page else 34
                chunks = [rows[i:i + chunk_size] for i in range(0, len(rows), chunk_size)] or [[]]
                for part, chunk in enumerate(chunks):
                    fig = Figure(figsize=(8.27, 11.69), constrained_layout=True)
                    ax = fig.add_subplot(111)
                    ax.axis("off")
                    if first_page:
                        heading = f"TOPCEE Process Cost & Economic Analysis\n{self.project_name}"
                        ax.text(0.02, 0.98, heading, transform=ax.transAxes,
                                va="top", fontsize=16, weight="bold", color="#1a3c5a")
                        summary = textwrap.fill(self.project_summary or "No project summary provided.", 95)
                        ax.text(0.02, 0.90, summary, transform=ax.transAxes,
                                va="top", fontsize=10, color="#333333")
                        ax.text(0.02, 0.84, f"Prepared by: {self.prepared_by}    Date: {self.date}    Currency: {self.currency}",
                                transform=ax.transAxes, va="top", fontsize=9)
                        top = 0.80
                    else:
                        top = 0.97
                    page_title = title if len(chunks) == 1 else f"{title} ({part + 1}/{len(chunks)})"
                    ax.text(0.02, top, page_title, transform=ax.transAxes,
                            va="top", fontsize=13, weight="bold", color="#2c5f8a")
                    table_top = top - 0.04
                    table = ax.table(cellText=chunk, colLabels=headers, cellLoc="left",
                                     colLoc="left", bbox=[0.02, 0.04, 0.96, table_top - 0.06])
                    table.auto_set_font_size(False)
                    table.set_fontsize(7.5)
                    for (row, _col), cell in table.get_celld().items():
                        if row == 0:
                            cell.set_facecolor("#eef2f7")
                            cell.set_text_props(weight="bold", color="#1a3c5a")
                    pdf.savefig(fig)
                    fig.clear()

            render_page("Project Overview", [
                ["Description", self.project_summary or "No project summary provided."],
                ["Prepared by", self.prepared_by], ["Date", self.date], ["Currency", self.currency],
            ], first_page=True)
            render_page("1. CAPEX / Total Capital Investment Breakdown", [
                [label, _money(self.capex.get(key, 0.0)) if key else ""]
                for _kind, _level, label, key in CAPEX_REPORT_LINES
            ])
            render_page("2. OPEX / Cost of Manufacturing Breakdown", [
                [label, _money(value) if value != "" else ""]
                for _kind, _level, label, value in self._opex_report_lines()
            ])
            prof_rows = []
            for key, label in [("nnpw", "NNPW"), ("dnpw", "DNPW"), ("npbp", "Project Payback (undiscounted, from year 0; years)"),
                               ("dpbp", "Project Payback (discounted, from year 0; years)"), ("roi", "Simple annual after-tax ROI (net profit / TCI, %)"), ("irr", "IRR (%)"),
                               ("bcncf", "Benefit-Cost Ratio (undiscounted)"),
                               ("bcdcf", "Benefit-Cost Ratio (discounted at selected rate)"),
                               ("com", "Cost of Manufacturing"), ("tci", "Total Capital Investment")]:
                value = self.profitability.get(key)
                if value is None:
                    shown = "Not reached" if key in ("npbp", "dpbp") else "Not defined"
                elif key in ("nnpw", "dnpw", "com", "tci"):
                    shown = _money(value)
                else:
                    shown = str(value)
                prof_rows.append([label, shown])
            render_page("3. Profitability & Investment Metrics", prof_rows)
            render_page("4. Annual Cash Flow", [
                [str(year), *[_money(values[i] if i < len(values) else 0.0) for values in
                              (self.cash_flow_ncf, self.cash_flow_cncf,
                               self.cash_flow_dcf, self.cash_flow_cdcf)]]
                for i, year in enumerate(self.cash_flow_years)
            ], headers=("Year", "NCF", "CNCF", "DCF", "CDCF"))
            if self.cash_flow_years:
                fig = Figure(figsize=(11.69, 8.27), constrained_layout=True)
                ax = fig.add_subplot(111)
                for label, values, marker in (
                    ("NCF", self.cash_flow_ncf, "o"), ("CNCF", self.cash_flow_cncf, "s"),
                    ("DCF", self.cash_flow_dcf, "^"), ("CDCF", self.cash_flow_cdcf, "d"),
                ):
                    ax.plot(self.cash_flow_years, values, label=label, marker=marker, markersize=3)
                ax.set_title(f"Cash Flow Diagram — {self.project_name}")
                ax.set_xlabel("Year")
                ax.set_ylabel("Cash Flow (USD)")
                ax.grid(True, alpha=0.3)
                ax.legend(loc="best")
                pdf.savefig(fig)
                fig.clear()
            eq_pdf_headers = ("Name", "Process ID", "Reference Source", "Ref Year", "Ref CEPCI",
                              "Target Year", "Target CEPCI", "Index Ratio", "Design Capacity", "Total Cost ($)", "Notes")
            eq_pdf_rows = [[row[index] if index < len(row) else ""
                            for index in (0, 1, 4, 5, 10, 11, 12, 13, 8, 18, 19)]
                           for row in self._equipment_rows()]
            render_page("5. Equipment Cost List", eq_pdf_rows or [["No equipment items"] + [""] * (len(eq_pdf_headers) - 1)],
                        headers=eq_pdf_headers)
            render_page("6. Input Data Summary", [[str(k), str(v)] for k, v in self.inputs.items()])
            if self.notes:
                render_page("Report Notes", [[str(i + 1), note] for i, note in enumerate(self.notes)],
                            headers=("#", "Note"))
        return path
