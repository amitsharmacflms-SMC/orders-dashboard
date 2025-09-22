import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Orders Dashboard", layout="wide")

# --- Global table width mode ---
TABLE_WIDTH_MODE = "stretch"   # change to "content" if you don’t want full width

# --- Normalize column names ---
def normalize_columns(df):
    df = df.copy()
    df.columns = (
        df.columns.str.strip()
        .str.replace(r"[_\s]+", " ", regex=True)  # unify underscores/spaces
        .str.title()  # Title case
    )
    return df

# --- Load Excel files ---
@st.cache_data
def load_data():
    df_summary = pd.read_excel("Summary.xlsx", engine="openpyxl")
    df_secondary = pd.read_excel("Secondary.xlsx", engine="openpyxl")
    return normalize_columns(df_summary), normalize_columns(df_secondary)

df_summary, df_secondary = load_data()

# --- Merge on exact join keys ---
join_keys = ["User", "Order Date"]

for df in [df_summary, df_secondary]:
    if "Order Date" in df.columns:
        df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce").dt.date  # ✅ strip to date only

try:
    df = pd.merge(
        df_summary, df_secondary,
        on=join_keys, how="left",   # ✅ keep all rows from Summary.xlsx
        suffixes=("_Sum", "_Sec")
    )
except Exception as e:
    st.error(f"Merge failed on {join_keys}: {e}")
    st.stop()

# --- Unify Sum/Sec columns into clean ones ---
def unify_columns(df, base):
    sum_col, sec_col = f"{base}_Sum", f"{base}_Sec"
    if sum_col in df.columns and sec_col in df.columns:
        # Prefer Summary (_Sum), fallback to Secondary (_Sec)
        df[base] = df[sum_col].combine_first(df[sec_col])
        df.drop(columns=[sum_col, sec_col], inplace=True)
    elif sum_col in df.columns:
        df[base] = df[sum_col]
        df.drop(columns=[sum_col], inplace=True)
    elif sec_col in df.columns:
        df[base] = df[sec_col]
        df.drop(columns=[sec_col], inplace=True)
    return df

# Columns we want unified for filters
for col in ["Region", "Territory", "Reporting Manager", "Distributor",
            "L4Position User", "L3Position User", "L2Position User"]:
    df = unify_columns(df, col)

# --- Debug: Show merged column names ---
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

# --- Compute retail time safely (HH:MM format, default 00:00 if missing) ---
if {"First Call", "Last Call"}.issubset(df.columns):
    first = pd.to_datetime(df["First Call"], format="%H:%M", errors="coerce")
    last = pd.to_datetime(df["Last Call"], format="%H:%M", errors="coerce")

    diff_minutes = (last - first).dt.total_seconds() / 60
    diff_minutes = pd.to_numeric(diff_minutes, errors="coerce").fillna(0).clip(lower=0)
    diff_minutes = diff_minutes.astype(int)

    df["Total Retail Time(Hh:Mm)"] = diff_minutes.apply(
        lambda x: f"{x // 60:02d}:{x % 60:02d}"
    )

# --- Filters (robust matching + date range) ---
st.markdown("### Filters")

required_filters = [
    "Order Date","Region","Territory","L4Position User","L3Position User",
    "L2Position User","Reporting Manager","Primary Category","User"
]

filter_selections = {}
available_cols = {c.lower().replace(" ", "").replace("_",""): c for c in df.columns}

matched_filters = {}
missing_filters = []

for f in required_filters:
    f_key = f.lower().replace(" ", "").replace("_","")
    if f == "Order Date":
        if "orderdate" in available_cols:
            col = available_cols["orderdate"]
            min_date, max_date = df[col].min(), df[col].max()
            start, end = st.date_input(
                "Order Date Range",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date
            )
            filter_selections[col] = (start, end)
            matched_filters[f] = col
        else:
            missing_filters.append(f)
    else:
        if f_key in available_cols:
            col = available_cols[f_key]
            matched_filters[f] = col
            vals = df[col].dropna().unique().tolist()
            if vals:
                sel = st.multiselect(f, sorted(vals), key=f"f_{f}")
                filter_selections[col] = sel
        else:
            missing_filters.append(f)

# --- Debug panel for filter matching ---
with st.expander("🔎 Filter Matching Debug"):
    st.write("✅ Matched filters:", matched_filters)
    st.write("⚠️ Missing filters (not found in merged data):", missing_filters)

# --- Apply filters ---
df_filtered = df.copy()
for col, sel in filter_selections.items():
    if col == "Order Date" and isinstance(sel, tuple):  # ✅ date range
        start, end = sel
        if start and end:
            df_filtered = df_filtered[
                (df_filtered[col] >= start) & (df_filtered[col] <= end)
            ]
    elif sel:
        df_filtered = df_filtered[df_filtered[col].isin(sel)]

# --- Final locked column order ---
final_columns = [
    "Order Date","Region","Territory","L4Position User","L3Position User","L2Position User",
    "Reporting Manager","Primary Category","Distributor","Beat","Outlet Name","Address","Market","Product","User",
    "Tc","Pc","Ovc","First Call","Last Call","Total Retail Time(Hh:Mm)",
    "Ghee","Dw Primary Packs","Dw Consu","Dw Bulk","36 No","Smp","Gjm","Cream","Uht Milk","Flavored Milk"
]
final_columns = [c for c in final_columns if c in df_filtered.columns]
final_df = df_filtered[final_columns].reset_index(drop=True)

# --- KPIs ---
st.markdown("### KPIs")
k1,k2,k3,k4 = st.columns(4)
k1.metric("Rows", len(final_df))
k2.metric("Unique Users", final_df["User"].nunique() if "User" in final_df.columns else 0)
k3.metric("Outlets", final_df["Outlet Name"].nunique() if "Outlet Name" in final_df.columns else 0)
k4.metric("Territories", final_df["Territory"].nunique() if "Territory" in final_df.columns else 0)

# --- Table ---
st.markdown("### Results Table (Top 200 Rows)")
st.dataframe(final_df.head(200), width="stretch")

# --- Export ---
def to_csv_bytes(df_obj):
    return df_obj.to_csv(index=False).encode("utf-8")

def to_excel_bytes(df_obj):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_obj.to_excel(writer, index=False)
    return out.getvalue()

st.download_button("Download CSV", to_csv_bytes(final_df), "filtered_export.csv", "text/csv")
st.download_button("Download Excel", to_excel_bytes(final_df), "filtered_export.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
