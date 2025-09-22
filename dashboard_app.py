import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Stylish Orders Dashboard", layout="wide")

# --- Page style ---
st.markdown("""
    <style>
    .stApp { font-family: "Inter", sans-serif; background: #f7fafc; }
    .header { background: linear-gradient(90deg,#4f46e5,#06b6d4); padding: 18px; border-radius:12px; color:white; }
    .card { background: white; padding: 16px; border-radius: 10px; box-shadow: 0 6px 18px rgba(15,23,42,0.06); }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header"><h2>Stylish Orders Dashboard</h2><p>Interactive filters and exports</p></div>', unsafe_allow_html=True)
st.write("")

# --- Auto-load Excel files from repo ---
@st.cache_data
def load_data():
    df_summary = pd.read_excel("Summary.xlsx", engine="openpyxl")
    df_secondary = pd.read_excel("Secondary.xlsx", engine="openpyxl")
    df_summary.columns = [str(c).strip() for c in df_summary.columns]
    df_secondary.columns = [str(c).strip() for c in df_secondary.columns]
    return df_summary, df_secondary

df_summary, df_secondary = load_data()

# --- Merge on exact join keys ---
join_keys = ["User", "Order Date"]

# Ensure "Order Date" is datetime in both
for df in [df_summary, df_secondary]:
    if "Order Date" in df.columns:
        df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")

try:
    df = pd.merge(
        df_summary,
        df_secondary,
        on=join_keys,
        how="outer",
        suffixes=("_sum", "_sec")
    )
except Exception as e:
    st.error(f"Merge failed on {join_keys}: {e}")
    df = pd.concat([df_summary.reset_index(drop=True), df_secondary.reset_index(drop=True)], axis=1)

# --- Add extra columns if missing ---
extra_cols = [
    "TC","PC","OVC","First Call","Last Call","Total Retail Time(HH:MM)",
    "Ghee","Dw Primary Packs","Dw Consu","DW Bulk","36 No","SMP","GJM","Cream","UHT Milk","Flavored Milk"
]
for c in extra_cols:
    if c not in df.columns:
        df[c] = 0 if "Call" not in c and "Time" not in c else pd.NaT

# Compute retail time if possible
if "First Call" in df.columns and "Last Call" in df.columns:
    try:
        diff = pd.to_datetime(df["Last Call"], errors="coerce") - pd.to_datetime(df["First Call"], errors="coerce")
        df["Total Retail Time(HH:MM)"] = diff.dt.total_seconds().div(60).fillna(0).astype(int).apply(lambda x: f"{x//60:02d}:{x%60:02d}")
    except:
        pass

# --- Filters ---
st.markdown("### Filters")
filter_cols = ["Order Date","Region","Territory","L4Position User","L3Position User","L2Position User","Reporting Manager","PrimaryCategory","User"]

filter_selections = {}
for c in filter_cols:
    if c in df.columns:
        sel = st.multiselect(c, sorted(df[c].dropna().unique()), key=f"f_{c}")
        filter_selections[c] = sel

df_filtered = df.copy()
for col, sel in filter_selections.items():
    if sel:
        df_filtered = df_filtered[df_filtered[col].isin(sel)]

# --- Column visibility ---
st.markdown("### Column Visibility")
main_cols = [
    "Order Date","Region","Territory","L4Position User","L3Position User","L2Position User",
    "Reporting Manager","PrimaryCategory","Distributor","Beat","Outlet Name","Address","Market","Product","User"
]
main_cols = [c for c in main_cols if c in df.columns]

vis_cols = []
colL, colR = st.columns(2)
with colL:
    st.markdown("**Main Columns**")
    for c in main_cols:
        if st.checkbox(c, True, key=f"vis_{c}"):
            vis_cols.append(c)
with colR:
    st.markdown("**Extra Columns**")
    for c in extra_cols:
        if st.checkbox(c, c in ["TC","PC","OVC"], key=f"vis_extra_{c}"):
            vis_cols.append(c)

if not vis_cols:
    st.warning("No columns selected. Please pick at least one.")
    st.stop()

final_df = df_filtered[vis_cols].reset_index(drop=True)

# --- KPIs ---
st.markdown("### KPIs")
k1,k2,k3,k4 = st.columns(4)
k1.metric("Rows", len(final_df))
k2.metric("Unique Users", final_df["User"].nunique() if "User" in final_df.columns else 0)
k3.metric("Outlets", final_df["Outlet Name"].nunique() if "Outlet Name" in final_df.columns else 0)
k4.metric("Territories", final_df["Territory"].nunique() if "Territory" in final_df.columns else 0)

# --- Table ---
st.markdown("### Results Table (Top 200 Rows)")
st.dataframe(final_df.head(200), use_container_width=True)

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
