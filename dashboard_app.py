import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Orders Dashboard", layout="wide")

TABLE_WIDTH_MODE = "stretch"

# --- Normalize column names ---
def normalize_columns(df):
    df = df.copy()
    df.columns = (
        df.columns.str.strip()
        .str.replace(r"[_\s]+", " ", regex=True)
        .str.title()
    )
    return df

# --- Load Excel files ---
@st.cache_data
def load_data():
    df_summary = pd.read_excel("Summary.xlsx", engine="openpyxl")
    df_secondary = pd.read_excel("Secondary.xlsx", engine="openpyxl")
    return normalize_columns(df_summary), normalize_columns(df_secondary)

df_summary, df_secondary = load_data()

# -------------------------
# Robust Order Date parsing + debug + left-merge (keep Summary rows)
# -------------------------
def robust_parse_date_col(series):
    """Try multiple formats, Excel serials, then fallback to generic parser.
       Returns a pd.Series of python date objects (or NaT)."""
    s = series.copy()
    s_str = s.astype(str).str.strip()

    # 1) Try explicit formats
    parsed = pd.to_datetime(s_str, format="%Y-%m-%d", errors="coerce")
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d", "%d-%b-%Y", "%d %b %Y", "%m/%d/%Y", "%d.%m.%Y"):
        parsed = parsed.fillna(pd.to_datetime(s_str, format=fmt, errors="coerce"))

    # 2) If values look numeric (Excel serial), attempt origin conversion
    s_num = pd.to_numeric(series, errors="coerce")
    if s_num.notna().any():
        parsed = parsed.fillna(pd.to_datetime(s_num, unit="d", origin="1899-12-30", errors="coerce"))

    # 3) Final fallback to generic parser (dateutil)
    parsed = parsed.fillna(pd.to_datetime(s_str, errors="coerce"))

    # Return date (not datetime) for consistent filtering
    return parsed.dt.date

# Show raw samples BEFORE parsing
with st.expander("🔎 Raw Order Date sample (Summary.xlsx)"):
    if "Order Date" in df_summary.columns:
        st.write(df_summary["Order Date"].head(200))
    else:
        st.write("No Order Date column in Summary.xlsx")

with st.expander("🔎 Raw Order Date sample (Secondary.xlsx)"):
    if "Order Date" in df_secondary.columns:
        st.write(df_secondary["Order Date"].head(200))
    else:
        st.write("No Order Date column in Secondary.xlsx")

# Apply robust parsing to both dataframes (in-place), keep raw for debug
for dname, dframe in [("Summary", df_summary), ("Secondary", df_secondary)]:
    if "Order Date" in dframe.columns:
        parsed = robust_parse_date_col(dframe["Order Date"])
        dframe["Order Date_Parsed"] = parsed
        dframe["Order Date_Raw"] = dframe["Order Date"].astype(str)
        dframe["Order Date"] = parsed

# Debug coverage BEFORE merge
with st.expander("🔎 Date coverage BEFORE merge"):
    if "Order Date" in df_summary.columns:
        st.write("Summary parsed min/max:", df_summary["Order Date"].min(), df_summary["Order Date"].max())
        st.write("Unique dates in Summary:", sorted(df_summary["Order Date"].dropna().unique()))
        st.write("Count parsed vs missing (Summary):",
                 {"parsed": int(df_summary["Order Date"].notna().sum()), "missing": int(df_summary["Order Date"].isna().sum())})
    if "Order Date" in df_secondary.columns:
        st.write("Secondary parsed min/max:", df_secondary["Order Date"].min(), df_secondary["Order Date"].max())
        st.write("Count parsed vs missing (Secondary):",
                 {"parsed": int(df_secondary["Order Date"].notna().sum()), "missing": int(df_secondary["Order Date"].isna().sum())})

# --- Merge on join keys, keep all Summary rows (left join) ---
join_keys = ["User", "Order Date"]
try:
    df = pd.merge(df_summary, df_secondary, on=join_keys, how="left", suffixes=("_Sum", "_Sec"))
except Exception as e:
    st.error(f"Merge failed on {join_keys}: {e}")
    st.stop()

# After merge debug
with st.expander("🔎 Date coverage AFTER merge"):
    if "Order Date" in df.columns:
        st.write("Merged parsed min/max:", df["Order Date"].min(), df["Order Date"].max())
        st.write("Unique dates in Merged:", sorted(df["Order Date"].dropna().unique()))
        st.write("Count parsed vs missing (Merged):",
                 {"parsed": int(df["Order Date"].notna().sum()), "missing": int(df["Order Date"].isna().sum())})

# --- Unify Sum/Sec columns into clean ones and drop originals ---
def unify_columns(df_local, base):
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

for col in ["Region", "Territory", "Reporting Manager", "Distributor",
            "L4Position User", "L3Position User", "L2Position User"]:
    df = unify_columns(df, col)

# --- Debug: Show merged column names (cleaned) ---
with st.expander("🔎 Debug: Show all merged column names"):
    st.write(list(df.columns))

# --- Ensure extra columns exist ---
extra_cols = [
    "Tc","Pc","Ovc","First Call","Last Call","Total Retail Time(Hh:Mm)",
    "Ghee","Dw Primary Packs","Dw Consu","Dw Bulk","36 No","Smp","Gjm","Cream","Uht Milk","Flavored Milk"
]
for c in extra_cols:
    if c not in df.columns:
        df[c] = 0 if "Call" not in c and "Time" not in c else pd.NaT

# --- Compute retail time safely (HH:MM) ---
if {"First Call", "Last Call"}.issubset(df.columns):
    first = pd.to_datetime(df["First Call"], format="%H:%M", errors="coerce")
    last = pd.to_datetime(df["Last Call"], format="%H:%M", errors="coerce")
    diff_minutes = (last - first).dt.total_seconds() / 60
    diff_minutes = pd.to_numeric(diff_minutes, errors="coerce").fillna(0).clip(lower=0).astype(int)
    df["Total Retail Time(Hh:Mm)"] = diff_minutes.apply(lambda x: f"{x // 60:02d}:{x % 60:02d}")

# --- Filters (robust matching + date range) ---
st.markdown("### Filters")

required_filters = [
    "Order Date","Region","Territory","L4Position User","L3Position User",
    "L2Position User","Reporting Manager","Primary Category","User"
]

filter_selections = {}
available_cols = {c.lower().replace(" ", "").replace("_",""): c for c in df.columns}

# Special: date range
if "orderdate" in available_cols:
    date_col = available_cols["orderdate"]
    # get min/max from merged df (skip NaT)
    min_date = df[date_col].dropna().min()
    max_date = df[date_col].dropna().max()
    if pd.isna(min_date) or pd.isna(max_date):
        # fallback to summary if merged lacks proper parsed dates
        if "Order Date" in df_summary.columns:
            min_date = df_summary["Order Date"].dropna().min()
            max_date = df_summary["Order Date"].dropna().max()
    # Protect against None
    if pd.notna(min_date) and pd.notna(max_date):
        start, end = st.date_input("Order Date Range", value=(min_date, max_date), min_value=min_date, max_value=max_date)
        filter_selections[date_col] = (start, end)
    else:
        st.info("Order Date range not available (parsing failed). See debug panels above.")

# Other filters
for f in required_filters:
    if f.lower().replace(" ", "").replace("_","") == "orderdate":
        continue
    key = f.lower().replace(" ", "").replace("_","")
    if key in available_cols:
        col = available_cols[key]
        vals = df[col].dropna().unique().tolist()
        if vals:
            sel = st.multiselect(f, sorted(vals), key=f"f_{f}")
            filter_selections[col] = sel
    else:
        st.info(f"⚠️ Column '{f}' not found in merged data.")

# --- Debug: Row counts at key steps ---
with st.expander("🔎 Row counts at each step"):
    st.write("Rows in Summary.xlsx:", len(df_summary))
    st.write("Rows in Secondary.xlsx:", len(df_secondary))
    st.write("Rows after merge:", len(df))

# --- Apply filters ---
df_filtered = df.copy()
for col, sel in filter_selections.items():
    if col == available_cols.get("orderdate") and isinstance(sel, tuple):
        start, end = sel
        if start and end:
            df_filtered = df_filtered[(df_filtered[col] >= start) & (df_filtered[col] <= end)]
    elif sel:
        df_filtered = df_filtered[df_filtered[col].isin(sel)]

# Debug rows after filters
with st.expander("🔎 Row counts after filters"):
    st.write("Rows after filters:", len(df_filtered))

# --- Final locked column order ---
final_columns = [
    "Order Date","Region","Territory","L4Position User","L3Position User","L2Position User",
    "Reporting Manager","Primary Category","Distributor","Beat","Outlet Name","Address","Market","Product","User",
    "Tc","Pc","Ovc","First Call","Last Call","Total Retail Time(Hh:Mm)",
    "Ghee","Dw Primary Packs","Dw Consu","Dw Bulk","36 No","Smp","Gjm","Cream","Uht Milk","Flavored Milk"
]
final_columns = [c for c in final_columns if c in df_filtered.columns]
final_df = df_filtered[final_columns].reset_index(drop=True)

# --- KPIs & Table & Export ---
st.markdown("### KPIs")
k1,k2,k3,k4 = st.columns(4)
k1.metric("Rows", len(final_df))
k2.metric("Unique Users", final_df["User"].nunique() if "User" in final_df.columns else 0)
k3.metric("Outlets", final_df["Outlet Name"].nunique() if "Outlet Name" in final_df.columns else 0)
k4.metric("Territories", final_df["Territory"].nunique() if "Territory" in final_df.columns else 0)

st.markdown("### Results Table (Top 200 Rows)")
st.dataframe(final_df.head(200), width=TABLE_WIDTH_MODE)

def to_csv_bytes(df_obj):
    return df_obj.to_csv(index=False).encode("utf-8")
def to_excel_bytes(df_obj):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_obj.to_excel(writer, index=False)
    return out.getvalue()

st.download_button("Download CSV", to_csv_bytes(final_df), "filtered_export.csv", "text/csv")
st.download_button("Download Excel", to_excel_bytes(final_df), "filtered_export.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
