import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Orders Dashboard", layout="wide")
TABLE_WIDTH_MODE = "stretch"

# ---------------------
# Helpers
# ---------------------
def normalize_columns(df):
    df = df.copy()
    df.columns = (
        df.columns.str.strip()
        .str.replace(r"[_\s]+", " ", regex=True)
        .str.title()
    )
    return df

def robust_parse_date_col(series):
    """Safe robust date parsing (handles 01-09-2025 style, Excel serials, and weird dashes)."""
    s = series.copy()
    # Convert to string, strip spaces, normalize separators
    s_str = s.astype(str).str.strip()
    s_str = s_str.str.replace("[–—]", "-", regex=True)  # replace en-dash/em-dash with normal dash
    s_str = s_str.str.replace("/", "-", regex=True)     # unify slashes to dash

    # 1) Try explicit formats
    parsed = pd.to_datetime(s_str, format="%d-%m-%Y", errors="coerce")
    for fmt in ("%Y-%m-%d", "%d-%b-%Y", "%d %b %Y", "%m-%d-%Y", "%d-%m-%y", "%d-%m-%Y %H:%M:%S"):
        parsed = parsed.fillna(pd.to_datetime(s_str, format=fmt, errors="coerce"))

    # 2) Excel-style serial numbers
    s_num = pd.to_numeric(series, errors="coerce")
    mask = s_num.notna()
    if mask.any():
        try:
            nums = s_num[mask].astype("float64").values
            converted = pd.to_datetime(nums, unit="d", origin="1899-12-30", errors="coerce")
            converted_series = pd.Series(converted, index=s_num[mask].index)
            parsed = parsed.fillna(converted_series)
        except Exception:
            converted_series = s_num[mask].apply(
                lambda x: pd.to_datetime(float(x), unit="d", origin="1899-12-30", errors="coerce")
            )
            parsed = parsed.fillna(converted_series)

    # 3) Final fallback
    parsed = parsed.fillna(pd.to_datetime(s_str, errors="coerce"))
    return parsed.dt.date

def aggregate_secondary(df_secondary, join_keys):
    """Aggregate secondary to unique join_keys rows."""
    df = df_secondary.copy()
    nums = df.select_dtypes(include=["number"]).columns.tolist()
    objs = [c for c in df.select_dtypes(include=["object", "category"]).columns.tolist() if c not in join_keys]

    agg_dict = {c: "sum" for c in nums}
    for c in objs:
        agg_dict[c] = "first"

    grouped = df.groupby(join_keys, dropna=False).agg(agg_dict).reset_index()
    return grouped

def unify_columns_and_drop(df_local, base):
    sum_col, sec_col = f"{base}_Sum", f"{base}_Sec"
    if sum_col in df_local.columns and sec_col in df_local.columns:
        df_local[base] = df_local[sum_col].combine_first(df_local[sec_col])
        df_local.drop(columns=[sum_col, sec_col], inplace=True, errors="ignore")
    elif sum_col in df_local.columns:
        df_local[base] = df_local[sum_col]
        df_local.drop(columns=[sum_col], inplace=True, errors="ignore")
    elif sec_col in df_local.columns:
        df_local[base] = df_local[sec_col]
        df_local.drop(columns=[sec_col], inplace=True, errors="ignore")
    return df_local

def parse_time_series(series):
    parsed = pd.to_datetime(series, format="%H:%M", errors="coerce")
    if parsed.isna().all():
        parsed = pd.to_datetime(series, errors="coerce")
    return parsed

# ---------------------
# Load data
# ---------------------
@st.cache_data
def load_data():
    df_summary = pd.read_excel("Summary.xlsx", engine="openpyxl")
    df_secondary = pd.read_excel("Secondary.xlsx", engine="openpyxl")
    return df_summary, df_secondary

df_summary_raw, df_secondary_raw = load_data()
df_summary = normalize_columns(df_summary_raw)
df_secondary = normalize_columns(df_secondary_raw)

if "Date" in df_summary.columns and "Order Date" not in df_summary.columns:
    df_summary = df_summary.rename(columns={"Date": "Order Date"})
if "Date" in df_secondary.columns and "Order Date" not in df_secondary.columns:
    df_secondary = df_secondary.rename(columns={"Date": "Order Date"})

if "Order Date" in df_summary.columns:
    df_summary["Order Date"] = robust_parse_date_col(df_summary["Order Date"])
if "Order Date" in df_secondary.columns:
    df_secondary["Order Date"] = robust_parse_date_col(df_secondary["Order Date"])

if "User" in df_summary.columns:
    df_summary["User"] = df_summary["User"].astype(str).str.strip().str.upper()
if "User" in df_secondary.columns:
    df_secondary["User"] = df_secondary["User"].astype(str).str.strip().str.upper()

if "Order Date" in df_secondary.columns:
    join_keys = ["User", "Order Date"]
    df_secondary_agg = aggregate_secondary(df_secondary, join_keys)
else:
    join_keys = ["User"]
    df_secondary_agg = aggregate_secondary(df_secondary, join_keys)

df = pd.merge(df_summary, df_secondary_agg, on=join_keys, how="left", suffixes=("_Sum", "_Sec"))

for col_base in [
    "Region", "Territory", "Reporting Manager", "Distributor",
    "L4Position User", "L3Position User", "L2Position User", "Primary Category"
]:
    df = unify_columns_and_drop(df, col_base)

extra_cols = [
    "Tc","Pc","Ovc","First Call","Last Call","Total Retail Time(Hh:Mm)",
    "Ghee","Dw Primary Packs","Dw Consu","Dw Bulk","36 No","Smp","Gjm","Cream","Uht Milk","Flavored Milk"
]
for c in extra_cols:
    if c not in df.columns:
        df[c] = 0 if ("Call" not in c and "Time" not in c) else pd.NaT

if {"First Call", "Last Call"}.issubset(df.columns):
    first_parsed = parse_time_series(df["First Call"])
    last_parsed = parse_time_series(df["Last Call"])
    diff_minutes = (last_parsed - first_parsed).dt.total_seconds() / 60
    diff_minutes = pd.to_numeric(diff_minutes, errors="coerce").fillna(0).clip(lower=0).astype(int)
    df["Total Retail Time(Hh:Mm)"] = diff_minutes.apply(lambda x: f"{x//60:02d}:{x%60:02d}")

# ---------------------
# Quick Date Check
# ---------------------
with st.expander("🔎 Check for 01-Sep-2025 and 22-Sep-2025 in Summary"):
    for d in ["2025-09-01", "2025-09-22"]:
        d_parsed = pd.to_datetime(d).date()
        rows = df_summary[df_summary["Order Date"] == d_parsed]
        st.write(f"{d} → {len(rows)} rows")
        st.write(rows.head(10))

# ---------------------
# Filters
# ---------------------
st.markdown("### Filters")

required_filters = [
    "Order Date","Region","Territory","L4Position User","L3Position User",
    "L2Position User","Reporting Manager","Primary Category","User"
]

filter_selections = {}
available_cols = {c.lower().replace(" ", "").replace("_",""): c for c in df.columns}

if "orderdate" in available_cols:
    date_col = available_cols["orderdate"]
    min_date, max_date = df[date_col].dropna().min(), df[date_col].dropna().max()
    if pd.notna(min_date) and pd.notna(max_date):
        start, end = st.date_input("Order Date Range", value=(min_date, max_date),
                                   min_value=min_date, max_value=max_date)
        filter_selections[date_col] = (start, end)

for f in required_filters:
    key = f.lower().replace(" ", "").replace("_","")
    if key == "orderdate": continue
    if key in available_cols:
        col = available_cols[key]
        vals = df[col].dropna().unique().tolist()
        if vals:
            sel = st.multiselect(f, sorted(vals), key=f"f_{f}")
            filter_selections[col] = sel

df_filtered = df.copy()
for col, sel in filter_selections.items():
    if isinstance(sel, tuple):
        start, end = sel
        df_filtered = df_filtered[(df_filtered[col] >= start) & (df_filtered[col] <= end)]
    elif sel:
        df_filtered = df_filtered[df_filtered[col].isin(sel)]

# ---------------------
# Final Columns
# ---------------------
final_columns = [
    "Order Date","Region","Territory","L4Position User","L3Position User","L2Position User",
    "Reporting Manager","Primary Category","Distributor","Beat","Outlet Name","Address","Market","Product","User",
    "Tc","Pc","Ovc","First Call","Last Call","Total Retail Time(Hh:Mm)",
    "Ghee","Dw Primary Packs","Dw Consu","Dw Bulk","36 No","Smp","Gjm","Cream","Uht Milk","Flavored Milk"
]
final_columns = [c for c in final_columns if c in df_filtered.columns]
final_df = df_filtered[final_columns].reset_index(drop=True)

# ---------------------
# KPIs & Export
# ---------------------
st.markdown("### KPIs")
k1,k2,k3,k4 = st.columns(4)
k1.metric("Rows", len(final_df))
k2.metric("Unique Users", final_df["User"].nunique() if "User" in final_df.columns else 0)
k3.metric("Outlets", final_df["Outlet Name"].nunique() if "Outlet Name" in final_df.columns else 0)
k4.metric("Territories", final_df["Territory"].nunique() if "Territory" in final_df.columns else 0)

# ---------------------
# Styled Table
# ---------------------
st.markdown("### Results Table (Top 200 Rows)")

st.markdown(
    """
    <style>
    .stDataFrame table, .stDataFrame th, .stDataFrame td {
        border: 2px solid black !important;
    }
    .stDataFrame th, .stDataFrame td {
        font-weight: bold !important;
    }
    .stDataFrame td {
        text-align: center !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.dataframe(final_df.head(200), width="stretch")

def to_csv_bytes(df_obj): return df_obj.to_csv(index=False).encode("utf-8")
def to_excel_bytes(df_obj):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_obj.to_excel(writer, index=False)
    return out.getvalue()

st.download_button("Download CSV", to_csv_bytes(final_df), "filtered_export.csv", "text/csv")
st.download_button("Download Excel", to_excel_bytes(final_df),
                   "filtered_export.xlsx",
                   "application/vnd.openxmlformats-officedocument-spreadsheetml.sheet")
