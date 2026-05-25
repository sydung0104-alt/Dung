#!/usr/bin/env python3
"""
Mechanical Quotation Analysis Tool v2.0
========================================
Reads Excel quotation files and produces supplier comparison, currency-
converted summaries, part-level comparisons, and formatted Excel reports.

Usage examples:
    python quotation_tool.py quotes.xlsx
    python quotation_tool.py quotes.xlsx --base-currency EUR --out report.xlsx
    python quotation_tool.py quotes.xlsx --filter-supplier "Alpha Valves" "ProPump Co."
    python quotation_tool.py quotes.xlsx --rates-file my_rates.json --top 5
    python quotation_tool.py quotes.xlsx --list-sheets
    python quotation_tool.py quotes.xlsx --col-price "Net Price" --col-qty "PCS"
"""

import sys
import json
import argparse
import urllib.request
from datetime import datetime
from pathlib import Path

import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────

VERSION = "2.0"

FALLBACK_RATES_TO_USD: dict[str, float] = {
    "USD": 1.000, "EUR": 1.080, "GBP": 1.270, "JPY": 0.00670,
    "CNY": 0.140,  "SGD": 0.740, "MYR": 0.210, "THB": 0.02800,
    "INR": 0.012,  "AUD": 0.650, "CAD": 0.740, "CHF": 1.120,
    "KRW": 0.00073,"HKD": 0.128, "TWD": 0.031, "VND": 0.0000400,
    "IDR": 0.000063,"PHP": 0.0175,"BRL": 0.200, "MXN": 0.0580,
    "AED": 0.272,  "SAR": 0.267, "TRY": 0.031, "ZAR": 0.0540,
    "SEK": 0.095,  "NOK": 0.094, "DKK": 0.145, "PLN": 0.2500,
    "CZK": 0.044,  "HUF": 0.0027,"RUB": 0.011, "UAH": 0.024,
}

_ALIASES: dict[str, list[str]] = {
    "supplier":    ["supplier", "vendor", "manufacturer", "brand", "seller", "source"],
    "description": ["description", "item name", "part name", "material", "product", "item desc"],
    "part_number": ["part number", "part no", "part#", "p/n", "pn", "model no", "item no", "ref"],
    "quantity":    ["quantity", "qty", "amount", "count", "pieces", "pcs", "nos", "no."],
    "unit_price":  ["unit price", "unit_price", "quoted price", "list price", "rate", "price", "cost"],
    "currency":    ["currency", "curr.", "ccy", "curr"],
    "lead_time":   ["lead time", "lead_time", "delivery", "weeks", "eta"],
    "discount":    ["discount", "disc", "disc.", "rebate", "%off"],
    "remarks":     ["remarks", "notes", "note", "comment"],
}

PALETTE = {
    "header_bg":  "#2E4057",
    "header_fg":  "#FFFFFF",
    "best_bg":    "#C6EFCE",
    "best_fg":    "#276221",
    "worst_bg":   "#FFC7CE",
    "worst_fg":   "#9C0006",
    "title_fg":   "#2E4057",
    "alt_row":    "#F2F2F2",
    "total_bg":   "#D9E1F2",
    "section_bg": "#4472C4",
}


# ─────────────────────────────────────────────────────────────────────────────
# Currency Converter
# ─────────────────────────────────────────────────────────────────────────────

class CurrencyConverter:
    """Live-rate currency converter with fallback to built-in rates."""

    def __init__(self, base: str = "USD", rates_file: Path | None = None, verbose: bool = False):
        self.base = base.upper()
        self.rates: dict[str, float] = {}   # each rate is value in USD
        self.source = "built-in fallback"
        self._verbose = verbose

        if rates_file:
            self._load_file(rates_file)
        else:
            self._fetch_live()

        if not self.rates:
            self.rates = FALLBACK_RATES_TO_USD.copy()

    def _load_file(self, path: Path):
        try:
            data = json.loads(path.read_text())
            data = data.get("rates", data)
            self.rates = {k.upper(): float(v) for k, v in data.items()}
            self.source = f"file: {path.name}"
        except Exception as exc:
            print(f"[WARN] Rates file error ({exc}); using built-in fallback.")
            self.rates = FALLBACK_RATES_TO_USD.copy()

    def _fetch_live(self):
        try:
            url = "https://open.er-api.com/v6/latest/USD"
            with urllib.request.urlopen(url, timeout=4) as resp:
                data = json.loads(resp.read())
            if data.get("result") == "success":
                self.rates = {k.upper(): float(v) for k, v in data["rates"].items()}
                self.source = f"live rates – open.er-api.com ({datetime.now():%Y-%m-%d})"
                if self._verbose:
                    print(f"[INFO] Fetched live exchange rates.")
        except Exception:
            if self._verbose:
                print("[INFO] Could not fetch live rates; using built-in fallback.")

    def convert(self, amount: float, from_ccy: str, to_ccy: str | None = None) -> float:
        to_ccy = (to_ccy or self.base).upper()
        from_ccy = from_ccy.upper()
        if from_ccy == to_ccy:
            return amount
        usd_val = amount / self.rates.get(from_ccy, 1.0)
        return usd_val * self.rates.get(to_ccy, 1.0)

    def rate(self, from_ccy: str) -> float:
        return self.convert(1.0, from_ccy)

    def symbol(self) -> str:
        symbols = {"USD": "$", "EUR": "€", "GBP": "£", "JPY": "¥", "CNY": "¥"}
        return symbols.get(self.base, self.base + " ")


# ─────────────────────────────────────────────────────────────────────────────
# Column Detection
# ─────────────────────────────────────────────────────────────────────────────

def _detect_col(columns: list[str], role: str) -> str | None:
    lower = [c.lower().strip() for c in columns]
    for alias in _ALIASES[role]:
        for i, col in enumerate(lower):
            if alias in col:
                return columns[i]
    return None


def detect_mapping(df: pd.DataFrame, overrides: dict[str, str]) -> dict[str, str]:
    cols = list(df.columns)
    mapping: dict[str, str] = {}
    for role in _ALIASES:
        found = overrides.get(role) or _detect_col(cols, role)
        if found and found in cols:
            mapping[role] = found
    return mapping


# ─────────────────────────────────────────────────────────────────────────────
# Data Loading & Cleaning
# ─────────────────────────────────────────────────────────────────────────────

def load_sheet(path: Path, sheet: str | int | None) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=sheet if sheet is not None else 0,
                           header=0, dtype=str)
    except Exception as exc:
        sys.exit(f"[ERROR] Cannot read '{path}': {exc}")
    df.dropna(how="all", inplace=True)
    df.dropna(axis=1, how="all", inplace=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def to_float(series: pd.Series) -> pd.Series:
    cleaned = series.astype(str).str.replace(r"[^\d.\-]", "", regex=True)
    return pd.to_numeric(cleaned, errors="coerce")


# ─────────────────────────────────────────────────────────────────────────────
# Analysis
# ─────────────────────────────────────────────────────────────────────────────

def build_line_items(df: pd.DataFrame, mapping: dict, conv: CurrencyConverter) -> pd.DataFrame:
    """Return a clean working DataFrame with numeric columns and converted prices."""
    w = df.copy()

    w["_qty"] = to_float(w[mapping["quantity"]]) if "quantity" in mapping else 1.0
    w["_qty"] = w["_qty"].fillna(1.0)

    w["_unit_price"] = to_float(w[mapping["unit_price"]]) if "unit_price" in mapping else float("nan")

    # Per-row currency
    if "currency" in mapping:
        w["_ccy"] = w[mapping["currency"]].astype(str).str.strip().str.upper()
    else:
        w["_ccy"] = conv.base

    w["_unit_price_base"] = w.apply(
        lambda r: conv.convert(r["_unit_price"], r["_ccy"])
        if pd.notna(r["_unit_price"]) else float("nan"),
        axis=1,
    )

    # Discount
    if "discount" in mapping:
        disc = to_float(w[mapping["discount"]]) / 100.0
        disc.fillna(0.0, inplace=True)
        w["_disc_pct"] = disc
        w["_net_price_base"] = w["_unit_price_base"] * (1.0 - w["_disc_pct"])
    else:
        w["_disc_pct"] = 0.0
        w["_net_price_base"] = w["_unit_price_base"]

    w["_line_total"] = w["_qty"] * w["_net_price_base"]

    # Supplier
    w["_supplier"] = (
        w[mapping["supplier"]].astype(str).str.strip()
        if "supplier" in mapping else "(no supplier)"
    )
    w["_supplier"] = w["_supplier"].replace({"nan": "(unknown)", "": "(unknown)"})

    # Part identifier (prefer part_number, fall back to description)
    if "part_number" in mapping:
        w["_part_id"] = w[mapping["part_number"]].astype(str).str.strip()
    elif "description" in mapping:
        w["_part_id"] = w[mapping["description"]].astype(str).str.strip()
    else:
        w["_part_id"] = w.index.astype(str)

    return w


def supplier_summary(li: pd.DataFrame) -> pd.DataFrame:
    useful = li.dropna(subset=["_unit_price_base"])
    grouped = (
        useful.groupby("_supplier", sort=False)
        .agg(
            items=("_line_total", "count"),
            total_qty=("_qty", "sum"),
            min_unit=("_unit_price_base", "min"),
            max_unit=("_unit_price_base", "max"),
            avg_unit=("_unit_price_base", "mean"),
            total_value=("_line_total", "sum"),
        )
        .reset_index()
        .rename(columns={"_supplier": "Supplier"})
        .sort_values("total_value", ascending=False)
    )
    grand = grouped["total_value"].sum()
    grouped["share_%"] = (grouped["total_value"] / grand * 100).round(1)
    return grouped


def part_comparison(li: pd.DataFrame) -> pd.DataFrame | None:
    """Build a part × supplier pivot for parts quoted by ≥ 2 suppliers."""
    useful = li.dropna(subset=["_unit_price_base", "_part_id"])
    useful = useful[useful["_part_id"].notna() & (useful["_part_id"] != "nan")]

    pivot = (
        useful.groupby(["_part_id", "_supplier"])["_net_price_base"]
        .mean()
        .unstack(level="_supplier")
    )
    # Keep only parts with ≥ 2 supplier quotes
    multi = pivot[pivot.notna().sum(axis=1) >= 2].copy()
    if multi.empty:
        return None

    multi["best_price"] = multi.min(axis=1)
    multi["best_supplier"] = multi.drop(columns="best_price").idxmin(axis=1)
    multi["savings_vs_worst"] = multi.drop(columns=["best_price", "best_supplier"]).max(axis=1) - multi["best_price"]
    multi["savings_%"] = (multi["savings_vs_worst"] / multi.drop(columns=["best_price", "best_supplier", "savings_vs_worst"]).max(axis=1) * 100).round(1)
    multi.reset_index(inplace=True)
    multi.rename(columns={"_part_id": "Part / Description"}, inplace=True)
    return multi


def overall_stats(li: pd.DataFrame, conv: CurrencyConverter) -> dict:
    useful = li.dropna(subset=["_line_total"])
    return {
        "total_value":      useful["_line_total"].sum(),
        "total_qty":        useful["_qty"].sum(),
        "line_items":       len(useful),
        "unique_suppliers": useful["_supplier"].nunique(),
        "currencies_found": sorted(useful["_ccy"].unique().tolist()),
        "base_currency":    conv.base,
        "rate_source":      conv.source,
        "generated":        datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


# ─────────────────────────────────────────────────────────────────────────────
# Console Report
# ─────────────────────────────────────────────────────────────────────────────

def _w(val, decimals=2, prefix="") -> str:
    if val is None or (isinstance(val, float) and val != val):
        return "N/A"
    return f"{prefix}{val:,.{decimals}f}"


def print_report(stats: dict, sup_df: pd.DataFrame, part_df: pd.DataFrame | None,
                 mapping: dict, path: Path, top: int | None):
    cur = stats["base_currency"]
    W = 74
    bar = "─" * W
    dbl = "═" * W

    def section(title):
        print(f"\n{bar}")
        print(f"  {title}")
        print(bar)

    print(f"\n{dbl}")
    print(f"  QUOTATION ANALYSIS  v{VERSION}  │  {path.name}")
    print(f"  Generated : {stats['generated']}")
    print(f"  Rates     : {stats['rate_source']}")
    print(dbl)

    # Detected columns
    print("\n  Detected columns:")
    for role, col in mapping.items():
        print(f"    {role:<14} → {col}")
    if stats["currencies_found"]:
        print(f"    {'currencies':<14} → {', '.join(stats['currencies_found'])}  → converted to {cur}")

    # Overall
    section("OVERALL TOTALS")
    print(f"  {'Unique suppliers':<22}: {stats['unique_suppliers']}")
    print(f"  {'Line items':<22}: {stats['line_items']}")
    print(f"  {'Total quantity':<22}: {stats['total_qty']:,.0f}")
    print(f"  {'Total value':<22}: {cur} {stats['total_value']:>16,.2f}")

    # By supplier
    section("SUPPLIER SUMMARY" + (f"  (top {top})" if top else ""))
    shown = sup_df.head(top) if top else sup_df

    cw = max((len(str(s)) for s in shown["Supplier"]), default=12) + 2
    cw = max(cw, 20)
    hdr = (f"  {'Supplier':<{cw}}{'Items':>6}  {'Qty':>10}  "
           f"{'Min Unit':>12}  {'Max Unit':>12}  {'Total':>14}  {'Share':>7}")
    print(hdr)
    print("  " + "─" * (len(hdr) - 2))

    for _, r in shown.iterrows():
        print(
            f"  {str(r['Supplier']):<{cw}}"
            f"{int(r['items']):>6}  "
            f"{r['total_qty']:>10,.0f}  "
            f"{cur}{r['min_unit']:>11,.2f}  "
            f"{cur}{r['max_unit']:>11,.2f}  "
            f"{cur}{r['total_value']:>13,.2f}  "
            f"{r['share_%']:>6.1f}%"
        )

    best_val   = sup_df.iloc[0]
    cheap_avg  = sup_df.loc[sup_df["avg_unit"].idxmin()]
    print(f"\n  ★  Highest-value supplier : {best_val['Supplier']}  ({cur} {best_val['total_value']:,.2f})")
    print(f"  ★  Lowest avg unit price  : {cheap_avg['Supplier']}  ({cur} {cheap_avg['avg_unit']:,.2f})")

    # Part comparison
    if part_df is not None:
        section(f"PART-LEVEL SUPPLIER COMPARISON  ({len(part_df)} parts, ≥2 suppliers)")
        suppliers = [c for c in part_df.columns
                     if c not in ("Part / Description", "best_price", "best_supplier",
                                  "savings_vs_worst", "savings_%")]
        pw = max((len(str(p)) for p in part_df["Part / Description"]), default=20) + 2
        pw = max(pw, 22)
        sup_head = "  ".join(f"{s[:14]:>14}" for s in suppliers)
        print(f"\n  {'Part / Description':<{pw}}  {sup_head}  {'Best Supplier':<18}  {'Saving'}")
        print("  " + "─" * (len(f"  {'Part / Description':<{pw}}  {sup_head}  {'Best Supplier':<18}  {'Saving'}") - 2))
        for _, r in part_df.iterrows():
            prices = "  ".join(
                f"{_w(r[s], 2, cur):>14}" if pd.notna(r.get(s)) else f"{'—':>14}"
                for s in suppliers
            )
            print(f"  {str(r['Part / Description'])[:pw-1]:<{pw}}  {prices}  "
                  f"{str(r['best_supplier'])[:17]:<18}  {r['savings_%']:.1f}%")

    print(f"\n{dbl}\n")


# ─────────────────────────────────────────────────────────────────────────────
# Excel Export
# ─────────────────────────────────────────────────────────────────────────────

class ExcelExporter:
    def __init__(self, path: Path, stats: dict, cur_sym: str):
        self.path = path
        self.stats = stats
        self.cur = stats["base_currency"]
        self.sym = cur_sym
        self.writer = pd.ExcelWriter(path, engine="xlsxwriter")
        self.wb = self.writer.book
        self._init_formats()

    def _init_formats(self):
        wb = self.wb
        self.fmt = {
            "title":    wb.add_format({"bold": True, "font_size": 14, "font_color": PALETTE["title_fg"]}),
            "subtitle": wb.add_format({"italic": True, "font_color": "#666666"}),
            "header":   wb.add_format({"bold": True, "bg_color": PALETTE["header_bg"],
                                        "font_color": PALETTE["header_fg"], "border": 1,
                                        "text_wrap": True, "valign": "vcenter"}),
            "cell":     wb.add_format({"border": 1, "valign": "vcenter"}),
            "num":      wb.add_format({"border": 1, "num_format": "#,##0.00"}),
            "int_":     wb.add_format({"border": 1, "num_format": "#,##0"}),
            "pct":      wb.add_format({"border": 1, "num_format": "0.0%"}),
            "total":    wb.add_format({"bold": True, "bg_color": PALETTE["total_bg"],
                                        "num_format": "#,##0.00", "border": 1}),
            "total_lbl":wb.add_format({"bold": True, "bg_color": PALETTE["total_bg"], "border": 1}),
            "best":     wb.add_format({"bold": True, "bg_color": PALETTE["best_bg"],
                                        "font_color": PALETTE["best_fg"], "border": 1,
                                        "num_format": "#,##0.00"}),
            "worst":    wb.add_format({"bg_color": PALETTE["worst_bg"],
                                        "font_color": PALETTE["worst_fg"], "border": 1,
                                        "num_format": "#,##0.00"}),
            "alt":      wb.add_format({"bg_color": PALETTE["alt_row"], "border": 1}),
            "alt_num":  wb.add_format({"bg_color": PALETTE["alt_row"], "border": 1,
                                        "num_format": "#,##0.00"}),
        }

    def _write_headers(self, ws, row: int, headers: list[str]):
        for ci, h in enumerate(headers):
            ws.write(row, ci, h, self.fmt["header"])

    # ── Sheet 1: Overview ────────────────────────────────────────────────────
    def write_overview(self, stats: dict, sup_df: pd.DataFrame, conv: CurrencyConverter):
        sup_df.to_excel(self.writer, sheet_name="Overview", index=False, startrow=10)
        ws = self.writer.sheets["Overview"]

        ws.write(0, 0, "Quotation Analysis Report", self.fmt["title"])
        ws.write(1, 0, f"Generated: {stats['generated']}  |  Rates: {stats['rate_source']}",
                 self.fmt["subtitle"])
        ws.write(3, 0, "KEY METRICS", self.fmt["header"])
        metrics = [
            ("Base currency",    stats["base_currency"]),
            ("Unique suppliers", stats["unique_suppliers"]),
            ("Total line items", stats["line_items"]),
            ("Total quantity",   f"{stats['total_qty']:,.0f}"),
            ("Total value",      f"{self.cur} {stats['total_value']:,.2f}"),
            ("Currencies found", ", ".join(stats["currencies_found"]) or self.cur),
        ]
        for i, (k, v) in enumerate(metrics, start=4):
            ws.write(i, 0, k, self.fmt["cell"])
            ws.write(i, 1, str(v), self.fmt["cell"])

        headers = ["Supplier", "Items", "Total Qty", "Min Unit Price",
                   "Max Unit Price", "Avg Unit Price", "Total Value", "Share %"]
        self._write_headers(ws, 10, headers)
        for ri, (_, r) in enumerate(sup_df.iterrows(), start=11):
            alt = ri % 2 == 0
            ws.write(ri, 0, r["Supplier"],          self.fmt["alt"] if alt else self.fmt["cell"])
            ws.write(ri, 1, int(r["items"]),         self.fmt["int_"])
            ws.write(ri, 2, r["total_qty"],          self.fmt["int_"])
            ws.write(ri, 3, r["min_unit"],           self.fmt["num"])
            ws.write(ri, 4, r["max_unit"],           self.fmt["num"])
            ws.write(ri, 5, r["avg_unit"],           self.fmt["num"])
            ws.write(ri, 6, r["total_value"],        self.fmt["num"])
            ws.write(ri, 7, r["share_%"] / 100,     self.fmt["pct"])

        tr = 11 + len(sup_df)
        ws.write(tr, 0, "TOTAL", self.fmt["total_lbl"])
        ws.write(tr, 1, int(sup_df["items"].sum()),       self.fmt["total"])
        ws.write(tr, 2, sup_df["total_qty"].sum(),        self.fmt["total"])
        ws.write(tr, 6, sup_df["total_value"].sum(),      self.fmt["total"])

        ws.set_column(0, 0, 26)
        ws.set_column(1, 7, 15)

        # Sparkline-style bar chart
        chart = self.wb.add_chart({"type": "bar"})
        n = len(sup_df)
        chart.add_series({
            "name":       "Total Value",
            "categories": ["Overview", 11, 0, 10 + n, 0],
            "values":     ["Overview", 11, 6, 10 + n, 6],
            "fill":       {"color": PALETTE["header_bg"]},
        })
        chart.set_title({"name": f"Total Value by Supplier ({self.cur})"})
        chart.set_legend({"none": True})
        chart.set_size({"width": 480, "height": 300})
        ws.insert_chart(tr + 2, 0, chart)

    # ── Sheet 2: Supplier Comparison ─────────────────────────────────────────
    def write_supplier_comparison(self, sup_df: pd.DataFrame):
        ws = self.wb.add_worksheet("Supplier Comparison")
        ws.write(0, 0, "Supplier Head-to-Head Comparison", self.fmt["title"])

        metrics = ["items", "total_qty", "min_unit", "max_unit", "avg_unit", "total_value", "share_%"]
        labels  = ["Line Items", "Total Qty", f"Min Unit ({self.cur})",
                   f"Max Unit ({self.cur})", f"Avg Unit ({self.cur})",
                   f"Total Value ({self.cur})", "Share (%)"]
        suppliers = sup_df["Supplier"].tolist()

        # Column headers = suppliers
        ws.write(2, 0, "Metric", self.fmt["header"])
        for ci, s in enumerate(suppliers, 1):
            ws.write(2, ci, s, self.fmt["header"])
        ws.write(2, len(suppliers) + 1, "Best", self.fmt["header"])

        for ri, (metric, label) in enumerate(zip(metrics, labels), start=3):
            ws.write(ri, 0, label, self.fmt["cell"])
            vals = sup_df[metric].tolist()
            best_idx = vals.index(min(vals)) if metric in ("min_unit", "avg_unit") else vals.index(max(vals))
            for ci, val in enumerate(vals, 1):
                fmt = self.fmt["best"] if (ci - 1) == best_idx else self.fmt["num"]
                if metric in ("items", "total_qty"):
                    fmt = self.fmt["int_"] if (ci - 1) != best_idx else self.fmt["best"]
                ws.write(ri, ci, val, fmt)
            ws.write(ri, len(suppliers) + 1, suppliers[best_idx], self.fmt["best"])

        ws.set_column(0, 0, 22)
        ws.set_column(1, len(suppliers) + 1, 18)

    # ── Sheet 3: Part Comparison ─────────────────────────────────────────────
    def write_part_comparison(self, part_df: pd.DataFrame | None):
        ws = self.wb.add_worksheet("Part Comparison")
        ws.write(0, 0, "Part-Level Supplier Comparison", self.fmt["title"])
        ws.write(1, 0, f"Green = best price  |  Red = highest price  |  All prices in {self.cur}",
                 self.fmt["subtitle"])

        if part_df is None:
            ws.write(3, 0, "No parts found with quotes from ≥ 2 suppliers.", self.fmt["cell"])
            return

        supplier_cols = [c for c in part_df.columns
                         if c not in ("Part / Description", "best_price",
                                      "best_supplier", "savings_vs_worst", "savings_%")]
        headers = ["Part / Description"] + supplier_cols + ["Best Supplier", "Max Saving", "Saving %"]
        self._write_headers(ws, 3, headers)

        for ri, (_, row) in enumerate(part_df.iterrows(), start=4):
            ws.write(ri, 0, str(row["Part / Description"]), self.fmt["cell"])
            prices = [row.get(s) for s in supplier_cols]
            valid  = [p for p in prices if pd.notna(p)]
            best   = min(valid) if valid else None
            worst  = max(valid) if valid else None
            for ci, (s, p) in enumerate(zip(supplier_cols, prices), start=1):
                if pd.isna(p):
                    ws.write(ri, ci, "—", self.fmt["cell"])
                elif p == best:
                    ws.write(ri, ci, p, self.fmt["best"])
                elif p == worst:
                    ws.write(ri, ci, p, self.fmt["worst"])
                else:
                    ws.write(ri, ci, p, self.fmt["num"])
            n = len(supplier_cols)
            ws.write(ri, n + 1, str(row["best_supplier"]),    self.fmt["cell"])
            ws.write(ri, n + 2, row["savings_vs_worst"],      self.fmt["num"])
            ws.write(ri, n + 3, row["savings_%"] / 100,       self.fmt["pct"])

        ws.set_column(0, 0, 30)
        ws.set_column(1, len(supplier_cols) + 3, 16)

    # ── Sheet 4: Line Items ──────────────────────────────────────────────────
    def write_line_items(self, li: pd.DataFrame, mapping: dict):
        keep = [c for c in li.columns if not c.startswith("_")]
        extra = pd.DataFrame({
            "Supplier (detected)":         li["_supplier"],
            f"Unit Price ({self.cur})":    li["_unit_price_base"].round(4),
            f"Net Price ({self.cur})":     li["_net_price_base"].round(4),
            f"Line Total ({self.cur})":    li["_line_total"].round(2),
        })
        out = pd.concat([li[keep].reset_index(drop=True), extra.reset_index(drop=True)], axis=1)
        out.to_excel(self.writer, sheet_name="Line Items", index=False, startrow=2)
        ws = self.writer.sheets["Line Items"]
        ws.write(0, 0, "All Line Items (parsed & currency-converted)", self.fmt["title"])
        for ci, col in enumerate(out.columns):
            ws.write(2, ci, col, self.fmt["header"])
        ws.set_column(0, len(out.columns), 16)

    # ── Sheet 5: Currency Rates ──────────────────────────────────────────────
    def write_currency_rates(self, conv: CurrencyConverter, found_ccys: list[str]):
        ws = self.wb.add_worksheet("Currency Rates")
        ws.write(0, 0, f"Exchange Rates Used  →  Base: {self.cur}", self.fmt["title"])
        ws.write(1, 0, f"Source: {conv.source}", self.fmt["subtitle"])
        headers = ["Currency", f"1 unit → {self.cur}", f"1 {self.cur} → currency"]
        self._write_headers(ws, 3, headers)
        all_ccys = sorted(set(found_ccys) | set(conv.rates.keys()))
        for ri, ccy in enumerate(all_ccys, start=4):
            rate_to_base = conv.rate(ccy)
            fmt = self.fmt["best"] if ccy in found_ccys else self.fmt["cell"]
            ws.write(ri, 0, ccy, fmt)
            ws.write(ri, 1, round(rate_to_base, 6),         self.fmt["num"])
            ws.write(ri, 2, round(1 / rate_to_base, 6) if rate_to_base else 0, self.fmt["num"])
        ws.set_column(0, 2, 20)
        ws.write(4 + len(all_ccys) + 1, 0,
                 "Highlighted rows = currencies found in this quotation file.", self.fmt["subtitle"])

    def save(self):
        self.writer.close()
        print(f"[✓] Report exported → {self.path}")


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="quotation_tool",
        description=f"Mechanical Quotation Analysis Tool v{VERSION}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    # Positional
    p.add_argument("file", help="Input Excel file (.xlsx / .xls)")

    # Sheet selection
    p.add_argument("--sheet",   default=None, help="Sheet name or 0-based index (default: first)")
    p.add_argument("--list-sheets", action="store_true", help="List all sheets and exit")

    # Output
    p.add_argument("--out",     default=None, help="Output .xlsx path (default: <input>_analysis.xlsx)")
    p.add_argument("--no-export", action="store_true", help="Print report only; skip Excel export")

    # Currency
    p.add_argument("--base-currency", default="USD", metavar="CCY",
                   help="Target currency for all conversions (default: USD)")
    p.add_argument("--rates-file",    default=None, type=Path, metavar="JSON",
                   help="JSON file with custom exchange rates {CCY: rate_vs_USD}")

    # Filtering / display
    p.add_argument("--filter-supplier", nargs="+", metavar="NAME",
                   help="Include only these suppliers (space-separated, case-insensitive)")
    p.add_argument("--top",  type=int, default=None, metavar="N",
                   help="Show only top N suppliers in the console report")
    p.add_argument("--min-value", type=float, default=None, metavar="AMT",
                   help="Exclude line items with total value below AMT (in base currency)")

    # Column overrides
    g = p.add_argument_group("column overrides")
    g.add_argument("--col-supplier",    dest="ov_supplier",    metavar="COL")
    g.add_argument("--col-qty",         dest="ov_quantity",    metavar="COL")
    g.add_argument("--col-price",       dest="ov_unit_price",  metavar="COL")
    g.add_argument("--col-currency",    dest="ov_currency",    metavar="COL")
    g.add_argument("--col-description", dest="ov_description", metavar="COL")
    g.add_argument("--col-part-number", dest="ov_part_number", metavar="COL")
    g.add_argument("--col-lead-time",   dest="ov_lead_time",   metavar="COL")

    p.add_argument("--verbose", action="store_true")
    return p


def main():
    parser = build_parser()
    args = parser.parse_args()

    path = Path(args.file)
    if not path.exists():
        sys.exit(f"[ERROR] File not found: {path}")

    # List sheets
    if args.list_sheets:
        xl = pd.ExcelFile(path)
        print(f"Sheets in '{path.name}':")
        for i, s in enumerate(xl.sheet_names):
            print(f"  [{i}]  {s}")
        return

    # Parse sheet arg (allow int or name)
    sheet_arg = args.sheet
    if sheet_arg is not None:
        try:
            sheet_arg = int(sheet_arg)
        except ValueError:
            pass

    print(f"[INFO] Loading: {path.name}")
    df = load_sheet(path, sheet_arg)
    if df.empty:
        sys.exit("[ERROR] Sheet is empty.")

    # Currency converter
    conv = CurrencyConverter(
        base=args.base_currency.upper(),
        rates_file=args.rates_file,
        verbose=args.verbose,
    )

    # Column mapping
    overrides = {k: v for k, v in {
        "supplier":    args.ov_supplier,
        "quantity":    args.ov_quantity,
        "unit_price":  args.ov_unit_price,
        "currency":    args.ov_currency,
        "description": args.ov_description,
        "part_number": args.ov_part_number,
        "lead_time":   args.ov_lead_time,
    }.items() if v}
    mapping = detect_mapping(df, overrides)

    if "unit_price" not in mapping:
        print("[WARN] No price column detected. Use --col-price to specify.")
    if "quantity" not in mapping:
        print("[WARN] No quantity column detected. Defaulting to 1 per row.")

    # Build line items
    li = build_line_items(df, mapping, conv)

    # Supplier filter
    if args.filter_supplier:
        filt = [s.lower() for s in args.filter_supplier]
        li = li[li["_supplier"].str.lower().isin(filt)]
        if li.empty:
            sys.exit("[ERROR] No rows match the supplier filter.")

    # Min-value filter
    if args.min_value is not None:
        li = li[li["_line_total"] >= args.min_value]

    # Compute results
    sup_df  = supplier_summary(li)
    part_df = part_comparison(li)
    stats   = overall_stats(li, conv)

    # Console report
    print_report(stats, sup_df, part_df, mapping, path, args.top)

    # Excel export
    if not args.no_export:
        out_path = Path(args.out) if args.out else path.with_stem(path.stem + "_analysis")
        out_path = out_path.with_suffix(".xlsx")
        exp = ExcelExporter(out_path, stats, conv.symbol())
        exp.write_overview(stats, sup_df, conv)
        exp.write_supplier_comparison(sup_df)
        exp.write_part_comparison(part_df)
        exp.write_line_items(li, mapping)
        exp.write_currency_rates(conv, stats["currencies_found"])
        exp.save()


if __name__ == "__main__":
    main()
