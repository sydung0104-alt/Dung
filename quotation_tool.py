"""
Mechanical Sales Engineer – Quotation Summarizer
Reads Excel quotation files and produces price/quantity/supplier summaries.

Usage:
    python quotation_tool.py quote.xlsx
    python quotation_tool.py quote.xlsx --sheet "Sheet2" --out summary.xlsx
    python quotation_tool.py quote.xlsx --currency USD --top 5
"""

import sys
import argparse
import re
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Column-detection aliases (case-insensitive substring match)
# ---------------------------------------------------------------------------
_ALIASES = {
    "supplier":    ["supplier", "vendor", "manufacturer", "brand", "source", "seller"],
    "description": ["description", "item name", "part name", "material", "product", "item", "part"],
    "part_number": ["part number", "part no", "part#", "p/n", "pn", "model no", "item no"],
    "quantity":    ["quantity", "qty", "amount", "count", "pieces", "pcs", "nos", "no."],
    "unit_price":  ["unit price", "unit_price", "quoted price", "list price", "rate", "price", "cost"],
    "currency":    ["currency", "curr", "ccy"],
    "lead_time":   ["lead time", "lead_time", "delivery", "weeks"],
}


def _detect_column(columns: list[str], role: str) -> str | None:
    """Return the first column whose name contains any alias for *role*."""
    lower_cols = [c.lower().strip() for c in columns]
    for alias in _ALIASES[role]:
        for i, col in enumerate(lower_cols):
            if alias in col:
                return columns[i]
    return None


def _to_numeric(series: pd.Series) -> pd.Series:
    """Strip currency symbols / commas and coerce to float."""
    cleaned = series.astype(str).str.replace(r"[^\d.\-]", "", regex=True)
    return pd.to_numeric(cleaned, errors="coerce")


# ---------------------------------------------------------------------------
# Core logic
# ---------------------------------------------------------------------------

def load_sheet(path: Path, sheet: str | None) -> pd.DataFrame:
    """Load an Excel sheet into a DataFrame, stripping blank rows/cols."""
    try:
        df = pd.read_excel(path, sheet_name=sheet or 0, header=0, dtype=str)
    except Exception as exc:
        sys.exit(f"[ERROR] Cannot read '{path}': {exc}")

    df.dropna(how="all", inplace=True)
    df.dropna(axis=1, how="all", inplace=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def detect_columns(df: pd.DataFrame, overrides: dict) -> dict:
    """Auto-detect role → column mapping, apply CLI overrides."""
    cols = list(df.columns)
    mapping = {}
    for role in _ALIASES:
        found = overrides.get(role) or _detect_column(cols, role)
        if found and found in cols:
            mapping[role] = found
    return mapping


def summarize(df: pd.DataFrame, mapping: dict, currency_symbol: str) -> dict:
    """
    Returns a dict with:
      - 'line_items': cleaned DataFrame
      - 'by_supplier': per-supplier aggregation
      - 'overall': scalar totals
    """
    df = df.copy()

    # ── Quantity ────────────────────────────────────────────────────────────
    if "quantity" in mapping:
        df["_qty"] = _to_numeric(df[mapping["quantity"]])
    else:
        df["_qty"] = 1.0  # assume 1 if not present

    # ── Unit price ──────────────────────────────────────────────────────────
    if "unit_price" in mapping:
        df["_unit_price"] = _to_numeric(df[mapping["unit_price"]])
    else:
        df["_unit_price"] = float("nan")

    df["_line_total"] = df["_qty"] * df["_unit_price"]

    # ── Supplier ────────────────────────────────────────────────────────────
    supplier_col = mapping.get("supplier")
    if supplier_col:
        df["_supplier"] = df[supplier_col].fillna("(unknown)").str.strip()
    else:
        df["_supplier"] = "(no supplier column)"

    # Drop rows with no price AND no quantity data (likely header repeats)
    useful = df.dropna(subset=["_unit_price", "_qty"], how="all").copy()

    # ── Per-supplier summary ────────────────────────────────────────────────
    by_supplier = (
        useful.groupby("_supplier", sort=False)
        .agg(
            line_items=("_line_total", "count"),
            total_qty=("_qty", "sum"),
            min_unit_price=("_unit_price", "min"),
            max_unit_price=("_unit_price", "max"),
            avg_unit_price=("_unit_price", "mean"),
            total_value=("_line_total", "sum"),
        )
        .reset_index()
        .rename(columns={"_supplier": "supplier"})
        .sort_values("total_value", ascending=False)
    )

    overall = {
        "total_line_items": len(useful),
        "total_qty": useful["_qty"].sum(),
        "total_value": useful["_line_total"].sum(),
        "unique_suppliers": useful["_supplier"].nunique(),
        "currency": currency_symbol,
    }

    return {"line_items": useful, "by_supplier": by_supplier, "overall": overall}


# ---------------------------------------------------------------------------
# Display
# ---------------------------------------------------------------------------

def fmt_num(val, decimals=2):
    if pd.isna(val):
        return "N/A"
    return f"{val:,.{decimals}f}"


def print_summary(result: dict, mapping: dict, path: Path):
    ov = result["overall"]
    cur = ov["currency"]
    bys = result["by_supplier"]

    sep = "─" * 72

    print(f"\n{'═' * 72}")
    print(f"  QUOTATION SUMMARY  │  {path.name}")
    print(f"{'═' * 72}")

    # Detected columns
    print("\nDetected columns:")
    for role, col in mapping.items():
        print(f"  {role:<14} → {col}")

    # Overall totals
    print(f"\n{sep}")
    print("OVERALL TOTALS")
    print(sep)
    print(f"  Suppliers          : {ov['unique_suppliers']}")
    print(f"  Line items         : {ov['total_line_items']}")
    print(f"  Total quantity     : {fmt_num(ov['total_qty'], 0)}")
    print(f"  Total value        : {cur} {fmt_num(ov['total_value'])}")

    # Per-supplier breakdown
    print(f"\n{sep}")
    print("BY SUPPLIER")
    print(sep)

    col_w = max(len(str(s)) for s in bys["supplier"]) + 2
    col_w = max(col_w, 20)

    header = (
        f"  {'Supplier':<{col_w}}"
        f"{'Items':>7}  {'Total Qty':>11}  "
        f"{'Min Price':>12}  {'Max Price':>12}  {'Total Value':>14}"
    )
    print(header)
    print("  " + "─" * (len(header) - 2))

    for _, row in bys.iterrows():
        print(
            f"  {str(row['supplier']):<{col_w}}"
            f"{int(row['line_items']):>7}  "
            f"{fmt_num(row['total_qty'], 0):>11}  "
            f"{cur}{fmt_num(row['min_unit_price']):>11}  "
            f"{cur}{fmt_num(row['max_unit_price']):>11}  "
            f"{cur}{fmt_num(row['total_value']):>13}"
        )

    # Best-value supplier
    best = bys.iloc[0]
    print(f"\n  ★  Highest-value supplier: {best['supplier']}  "
          f"({cur}{fmt_num(best['total_value'])})")

    cheapest_avg = bys.loc[bys["avg_unit_price"].idxmin()]
    print(f"  ★  Lowest avg unit price : {cheapest_avg['supplier']}  "
          f"({cur}{fmt_num(cheapest_avg['avg_unit_price'])} avg)")

    if "lead_time" in mapping:
        print(f"\n  ℹ  Lead-time column detected: '{mapping['lead_time']}' "
              "(see exported file for values)")

    print(f"{'═' * 72}\n")


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def export_excel(result: dict, out_path: Path, mapping: dict):
    ov = result["overall"]
    cur = ov["currency"]

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr_fmt = wb.add_format({"bold": True, "bg_color": "#2E4057", "font_color": "white", "border": 1})
        num_fmt = wb.add_format({"num_format": f'#,##0.00', "border": 1})
        int_fmt = wb.add_format({"num_format": "#,##0", "border": 1})
        cell_fmt = wb.add_format({"border": 1})
        title_fmt = wb.add_format({"bold": True, "font_size": 13, "font_color": "#2E4057"})
        total_fmt = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "num_format": "#,##0.00", "border": 1})

        # ── Sheet 1: Supplier Summary ────────────────────────────────────────
        bys = result["by_supplier"].copy()
        bys.to_excel(writer, sheet_name="Supplier Summary", index=False, startrow=3)
        ws = writer.sheets["Supplier Summary"]
        ws.write(0, 0, "Quotation Supplier Summary", title_fmt)
        ws.write(1, 0, f"Total value: {cur} {ov['total_value']:,.2f}   "
                       f"Suppliers: {ov['unique_suppliers']}   "
                       f"Line items: {ov['total_line_items']}")

        for ci, col_name in enumerate(bys.columns):
            ws.write(3, ci, col_name, hdr_fmt)
        for ri, (_, row) in enumerate(bys.iterrows(), start=4):
            for ci, val in enumerate(row):
                fmt = num_fmt if isinstance(val, float) else (int_fmt if isinstance(val, int) else cell_fmt)
                ws.write(ri, ci, val, fmt)

        # totals row
        tr = 4 + len(bys)
        ws.write(tr, 0, "TOTAL", total_fmt)
        ws.write(tr, 1, ov["total_line_items"], total_fmt)
        ws.write(tr, 2, ov["total_qty"], total_fmt)
        ws.write(tr, 5, ov["total_value"], total_fmt)

        ws.set_column(0, 0, 24)
        ws.set_column(1, 5, 15)

        # ── Sheet 2: Line Items ──────────────────────────────────────────────
        display_cols = {v: k for k, v in mapping.items()}
        li = result["line_items"].copy()
        keep = [c for c in li.columns if not c.startswith("_")]
        li = li[keep]
        li.to_excel(writer, sheet_name="Line Items", index=False, startrow=1)
        ws2 = writer.sheets["Line Items"]
        ws2.write(0, 0, "All line items (parsed)", title_fmt)
        for ci, col_name in enumerate(li.columns):
            label = f"{col_name} [{display_cols[col_name]}]" if col_name in display_cols else col_name
            ws2.write(1, ci, label, hdr_fmt)
        ws2.set_column(0, len(li.columns), 18)

    print(f"[✓] Summary exported to: {out_path}")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Summarize mechanical quotation Excel files.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("file", help="Input Excel file (.xlsx / .xls)")
    p.add_argument("--sheet", default=None, help="Sheet name or index (default: first sheet)")
    p.add_argument("--out", default=None, help="Export summary to this .xlsx file")
    p.add_argument("--currency", default="", help="Currency symbol prefix (e.g. USD, €)")
    p.add_argument("--top", type=int, default=None, help="Show only top N suppliers by value")
    # Manual column overrides
    p.add_argument("--col-supplier",    dest="supplier",    default=None, help="Column name for supplier")
    p.add_argument("--col-qty",         dest="quantity",    default=None, help="Column name for quantity")
    p.add_argument("--col-price",       dest="unit_price",  default=None, help="Column name for unit price")
    p.add_argument("--col-description", dest="description", default=None, help="Column name for description")
    p.add_argument("--list-sheets",     action="store_true", help="List available sheets and exit")
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
        print("Sheets in", path.name)
        for i, s in enumerate(xl.sheet_names):
            print(f"  [{i}] {s}")
        return

    df = load_sheet(path, args.sheet)

    if df.empty:
        sys.exit("[ERROR] The sheet appears to be empty.")

    overrides = {
        k: v for k, v in {
            "supplier":    args.supplier,
            "quantity":    args.quantity,
            "unit_price":  args.unit_price,
            "description": args.description,
        }.items() if v
    }

    mapping = detect_columns(df, overrides)

    if "unit_price" not in mapping and "quantity" not in mapping:
        print("[WARN] Could not detect price or quantity columns automatically.")
        print("       Columns found:", list(df.columns))
        print("       Use --col-price and --col-qty to specify them manually.")

    cur = args.currency if args.currency else ""
    result = summarize(df, mapping, cur)

    if args.top:
        result["by_supplier"] = result["by_supplier"].head(args.top)

    print_summary(result, mapping, path)

    out_path = Path(args.out) if args.out else path.with_stem(path.stem + "_summary")
    out_path = out_path.with_suffix(".xlsx")
    export_excel(result, out_path, mapping)


if __name__ == "__main__":
    main()
