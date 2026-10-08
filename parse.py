import pandas as pd
from pathlib import Path
import numpy as np

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


    s = s.replace("R$", "").replace("$", "").replace(" ", "")
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

def apply_parser(df, column, parser):

    df[f"{column}_raw"] = df[column]
    df[[column, f"{column}_error"]] = df[column].apply(parser).tolist()
    return df

def fill_totals(df):
    df["total_source"] = "original"

    missing_total = df["total"].isna() & df["unit_price"].notna()
    df.loc[missing_total, "total"] = df["quantity"] * df["unit_price"]
    df.loc[missing_total, "total_source"] = "calculated"

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

def parse_report(df, columns, n=20):
    cols = ["order"] + [c for col in columns for c in (f"{col}_raw", col)]
    print(df[cols].head(n))

    summary = pd.DataFrame({
        "ok": [df[f"{c}_error"].isna().sum() for c in columns],
        "errors": [df[f"{c}_error"].notna().sum() for c in columns],
    }, index=columns)
    print("\n=== Summary ===")
    print(summary)


parse_report(df, PARSERS)
print(df[["order", "quantity", "unit_price", "total", "total_source"]].head(20))


pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)

print(parse_report(df, PARSERS))