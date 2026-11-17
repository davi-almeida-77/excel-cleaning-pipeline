import pandas as pd
from pathlib import Path
import numpy as np
import re
from datetime import datetime
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)

BASE = Path(__file__).resolve().parent

RAW = BASE / "data" / "raw"

df = pd.read_excel(RAW / "sales_january_2026.xlsx", header=5)

df_before = df.copy()

canonical_map = {
    "Order": "order",
    "Date": "date",
    "Customer": "customer",
    "Product" : "product",
    "Qty": "quantity",
    "Unit Price": "unit_price",
    "Total": "total",    
}

FORMATS = [
    "%B %d, %Y",
    "%d-%b-%Y",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%Y-%m-%d",
]


df = df.rename(columns=canonical_map)

def parse_number(raw):


    if pd.isna(raw):
        return float("nan"), ("empty")

    if isinstance(raw, (int, float)):
        return float(raw), None

    s = str(raw).strip()

    if s == "":
        return float("nan"), "empty"


    s = re.sub(r"[^0-9.,]", "", s)

    error = None
    
    last_comma = s.rfind(",")
    last_dot = s.rfind(".")

    if last_comma != -1 and last_dot != -1:
        if last_comma > last_dot:      
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")

    elif last_comma != -1:
        if s.count(",") > 1:
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
        
    else:
        if s.count(".") > 1:
            s = s.replace(".", "")
        elif last_dot != -1 and len(s.split(".")[-1]) == 3:
            error = "ambiguous"

    try:
        value = float(s)
        return float(s), error
    except ValueError:
        return float("nan"), "not a number"
    
def parse_date(raw):

    if pd.isna(raw):
        return pd.NaT, "empty"

    if isinstance(raw, (int, float)):
        return pd.to_datetime(raw, unit="D", origin="1899-12-30"), None

    s = str(raw).strip()

    for fmt in FORMATS:
        try:
            return pd.to_datetime(s, format=fmt), None
        except ValueError:
            continue

    if s.isdigit():
        return pd.to_datetime(float(s), unit="D", origin="1899-12-30"), None

    return pd.NaT, "unrecognized format"

def fill_totals(df):
    df["total_source"] = "original"

    missing_total = (df["total_error"] == "empty") & df["unit_price"].notna() & df["quantity"].notna()
    df.loc[missing_total, "total"] = df["quantity"] * df["unit_price"]
    df.loc[missing_total, "total_source"] = "calculated"

    return df

def clean_customer(df):
    df["customer"] = df["customer"].str.strip().str.title()
    df["customer"] = df["customer"].replace(["Null", "Customer", "N/a"], pd.NA)

    return df

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
  
    df = df.copy()
    df.columns = (
        df.columns.str.strip()
        .str.lower()
        .str.replace(r"\s+", "_", regex=True)
    )
    return df


df = df[df["order"].str.startswith("ORD-", na=False)]

def apply_parser(df, column, parser):

    df[f"{column}_raw"] = df[column]
    df[[column, f"{column}_error"]] = df[column].apply(parser).tolist()
    return df


PARSERS = {
    "date": parse_date,
    "quantity": parse_number,
    "unit_price": parse_number,
    "total": parse_number,
}


for col, parser in PARSERS.items():
    df = apply_parser(df, col, parser)

df = fill_totals(df)
df = clean_customer(df)

def parse_report(df, columns, n=20):
    cols = ["order"] + [c for col in columns for c in (f"{col}_raw", col)]
    print(df[cols].head(n))

    summary = pd.DataFrame({
        "ok": [df[f"{c}_error"].isna().sum() for c in columns],
        "errors": [df[f"{c}_error"].notna().sum() for c in columns],
    }, index=columns)
    print("\n  Summary " )
    print(summary)


CLEAN_COLUMNS = ["order", "date", "customer", "product", "quantity", "unit_price", "total"]
PROCESSED = BASE / "data" / "processed"

CHANGED_FILL = PatternFill("solid", start_color="FFF2CC")
CALCULATED_FILL = PatternFill("solid", start_color="DDEBF7")
ERROR_FILL = PatternFill("solid", start_color="F8CBAD")
DROPPED_FILL = PatternFill("solid", start_color="D9D9D9")
HEADER_FONT = Font(bold=True)


def export_clean(df, path):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(path, engine="openpyxl", datetime_format="YYYY-MM-DD") as writer:
        df[CLEAN_COLUMNS].to_excel(writer, sheet_name="Cleaned", index=False)
        autofit(writer.sheets["Cleaned"])

    return path


def same_value(original, cleaned):
    if pd.isna(original) and pd.isna(cleaned):
        return True
    if pd.isna(original) or pd.isna(cleaned):
        return False
    if isinstance(cleaned, pd.Timestamp):
        return isinstance(original, (pd.Timestamp, datetime)) and pd.Timestamp(original) == cleaned
    if isinstance(cleaned, (int, float)):
        try:
            return float(original) == float(cleaned)
        except (TypeError, ValueError):
            return False
    return str(original) == str(cleaned)


def find_changes(df, df_before):
    before = df_before.rename(columns=canonical_map)
    changes = []

    for idx, row in df.iterrows():
        for col in CLEAN_COLUMNS:
            original = before.at[idx, col]
            cleaned = row[col]
            error = row.get(f"{col}_error")
            if col == "total" and row["total_source"] == "calculated":
                status = "calculated"
            elif pd.notna(error) and error != "empty":
                status = "error"
            elif not same_value(original, cleaned):
                status = "changed"
            else:
                continue
            changes.append({
                "row": idx,
                "order": row["order"],
                "column": col,
                "original": original,
                "cleaned": cleaned,
                "status": status,
                "error": error if pd.notna(error) else None,
            })

    return pd.DataFrame(changes, columns=["row", "order", "column", "original", "cleaned", "status", "error"])


def export_report(df, df_before, path):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    changes = find_changes(df, df_before)
    dropped = df_before.index.difference(df.index)

    summary = pd.DataFrame({
        "changed": [(changes["column"].eq(c) & changes["status"].eq("changed")).sum() for c in CLEAN_COLUMNS],
        "calculated": [(changes["column"].eq(c) & changes["status"].eq("calculated")).sum() for c in CLEAN_COLUMNS],
        "errors": [(changes["column"].eq(c) & changes["status"].eq("error")).sum() for c in CLEAN_COLUMNS],
    }, index=pd.Index(CLEAN_COLUMNS, name="column"))
    totals = pd.DataFrame({
        "metric": ["original rows", "kept rows", "dropped rows", "cells changed"],
        "value": [len(df_before), len(df), len(dropped), len(changes)],
    })

    fills = {"changed": CHANGED_FILL, "calculated": CALCULATED_FILL, "error": ERROR_FILL}

    with pd.ExcelWriter(path, engine="openpyxl", datetime_format="YYYY-MM-DD") as writer:
        df_before.to_excel(writer, sheet_name="Original", index=False)
        df[CLEAN_COLUMNS].to_excel(writer, sheet_name="Cleaned", index=False)
        totals.to_excel(writer, sheet_name="Report", index=False)
        summary.to_excel(writer, sheet_name="Report", startrow=len(totals) + 2)
        changes.astype(object).to_excel(writer, sheet_name="Changes", index=False)

        original_ws = writer.sheets["Original"]
        cleaned_ws = writer.sheets["Cleaned"]

        original_row = {idx: pos + 2 for pos, idx in enumerate(df_before.index)}
        cleaned_row = {idx: pos + 2 for pos, idx in enumerate(df.index)}
        original_col = {canonical_map.get(c, c): i + 1 for i, c in enumerate(df_before.columns)}
        cleaned_col = {c: i + 1 for i, c in enumerate(CLEAN_COLUMNS)}

        for change in changes.itertuples():
            fill = fills[change.status]
            cleaned_ws.cell(cleaned_row[change.row], cleaned_col[change.column]).fill = fill
            original_ws.cell(original_row[change.row], original_col[change.column]).fill = fill

        for idx in dropped:
            for cell in original_ws[original_row[idx]]:
                cell.fill = DROPPED_FILL

        changes_ws = writer.sheets["Changes"]
        for r, status in enumerate(changes["status"], start=2):
            for cell in changes_ws[r]:
                cell.fill = fills[status]

        legend_ws = writer.sheets["Report"]
        legend_row = len(totals) + len(summary) + 6
        legend_ws.cell(legend_row, 1, "Legend").font = HEADER_FONT
        for i, (label, fill) in enumerate([
            ("changed by cleaning", CHANGED_FILL),
            ("calculated (was empty)", CALCULATED_FILL),
            ("could not be parsed", ERROR_FILL),
            ("row dropped", DROPPED_FILL),
        ], start=1):
            legend_ws.cell(legend_row + i, 1).fill = fill
            legend_ws.cell(legend_row + i, 2, label)

        for ws in writer.sheets.values():
            ws.freeze_panes = "A2"
            autofit(ws)

    return path


def autofit(ws):
    for column in ws.columns:
        width = max((len(str(c.value)) for c in column if c.value is not None), default=0)
        ws.column_dimensions[get_column_letter(column[0].column)].width = min(width + 2, 50)


print(df[["order", "date", "customer", "product", "quantity", "unit_price", "total", "total_source"]])

export_clean(df, PROCESSED / "sales_january_2026_clean.xlsx")
export_report(df, df_before, PROCESSED / "sales_january_2026_report.xlsx")
