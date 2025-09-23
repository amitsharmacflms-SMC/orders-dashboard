import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date, timedelta

st.set_page_config(page_title="SMC-Madhusudan Daily Working Dashboard", layout="wide")

# ---------------------
# Header + Tabs
# ---------------------
st.markdown(
    """
    <style>
    .main-header {
        font-size: 32px;
        font-weight: bold;
        color: white;
        background: linear-gradient(90deg, #4a90e2, #9013fe);
        padding: 15px;
        border-radius: 10px;
        text-align: center;
    }
    .filter-row {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        margin-bottom: 20px;
    }
    .stDataFrame thead tr th {
        font-weight: bold !important;
        background-color: #4a90e2 !important;
        color: white !important;
        text-align: center !important;
    }
    .stDataFrame tbody td {
        font-weight: bold !important;
        text-align: center !important;
        border: 1.5px solid #444 !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown('<div class="main-header">SMC-Madhusudan Daily Working Dashboard</div>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📊 Daily Summary", "🏪 Outlet Wise Report"])

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
    s = series.copy()
    parsed = pd.to_datetime(s, errors="coerce")
    return parsed.dt.date

# ---------------------
# Load data
# ---------------------
@st.cache_data
def load_data():
    df_summary = pd.read_excel("Summary.xlsx", engine="openpyxl")
    df_secondary = pd.read_excel("Secondary.xlsx", engine="openpyxl")
    return normalize_columns(df_summary), normalize_columns(df_secondary)

df_summary, df_secondary = load_data()

# Rename Date -> Order Date
if "Date" in df_summary.columns and "Order Date" not in df_summary.columns:
    df_summary = df_summary.rename(columns={"Date": "Order Date"})
if "Date" in df_secondary.columns and "Order Date" not in df_secondary.columns:
    df_secondary = df_secondary.rename(columns={"Date": "Order Date"})

# Parse dates
if "Order Date" in df_summary.columns:
    df_summary["Order Date"] = robust_parse_date_col(df_summary["Order Date"])
if "Order Date" in df_secondary.columns:
    df_secondary["Order Date"] = robust_parse_date_col(df_secondary["Order Date"])

# Merge
join_keys = ["User", "Order Date"] if "Order Date" in df_secondary.columns else ["User"]
df = pd.merge(df_summary, df_secondary, on=join_keys, how="left", suffixes=("_Sum", "_Sec"))

# Drop unwanted cols from table
remove_cols = ["Outlet Name", "Address", "Market", "Product"]
df = df.drop(columns=[c for c in remove_cols if c in df.columns], errors="ignore")

# ---------------------
# DAILY SUMMARY TAB
# ---------------------
with tab1:
    st.subheader("📊 Daily Summary Report")

    # Filter order
    required_filters = [
        "Order Date", "Region", "User",
        "L4Position User", "L3Position User", "L2Position User",
        "Reporting Manager", "Primary Category"
    ]

    filter_selections = {}

    # ---- Date Filters ----
    min_date, max_date = df["Order Date"].min(), df["Order Date"].max()

    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        date_range = st.date_input("Order Date Range", value=(min_date, max_date), min_value=min_date, max_value=max_date)
    with col2:
        single_date = st.date_input("Single Date", value=max_date, min_value=min_date, max_value=max_date)
    with col3:
        date_group = st.selectbox("Date Group", ["All", "Last 7 Days", "Last 15 Days"])

    # Apply date filters
    df_filtered = df.copy()
    if isinstance(date_range, tuple):
        df_filtered = df_filtered[(df_filtered["Order Date"] >= date_range[0]) & (df_filtered["Order Date"] <= date_range[1])]
    if single_date:
        df_filtered = df_filtered[df_filtered["Order Date"] == single_date]
    if date_group == "Last 7 Days":
        df_filtered = df_filtered[df_filtered["Order Date"] >= date.today() - timedelta(days=7)]
    elif date_group == "Last 15 Days":
        df_filtered = df_filtered[df_filtered["Order Date"] >= date.today() - timedelta(days=15)]

    # ---- Other Filters (Smart Filtering) ----
    filter_cols = [f for f in required_filters if f != "Order Date"]

    for f in filter_cols:
        if f in df_filtered.columns:
            vals = sorted(df_filtered[f].dropna().unique().tolist())
            vals = ["All"] + vals
            sel = st.multiselect(f, vals, default="All")
            if "All" not in sel:
                df_filtered = df_filtered[df_filtered[f].isin(sel)]

    # ---- Column Selection ----
    cols_available = df_filtered.columns.tolist()
    cols_available = ["All"] + cols_available
    selected_cols = st.multiselect("Columns Wants in Table", cols_available, default="All")

    if "All" in selected_cols or not selected_cols:
        final_df = df_filtered
    else:
        final_df = df_filtered[selected_cols]

    st.markdown("### Results Table (Top 200 Rows)")
    st.dataframe(final_df.head(200), width="stretch")

    # Export
    def to_csv_bytes(df_obj): return df_obj.to_csv(index=False).encode("utf-8")
    def to_excel_bytes(df_obj):
        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df_obj.to_excel(writer, index=False)
        return out.getvalue()

    st.download_button("Download CSV", to_csv_bytes(final_df), "filtered_export.csv", "text/csv")
    st.download_button("Download Excel", to_excel_bytes(final_df),
                       "filtered_export.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------
# OUTLET WISE REPORT TAB
# ---------------------
with tab2:
    st.subheader("🏪 Outlet Wise Report")
    st.info("This section is a placeholder. You can plug in outlet-wise logic here.")
