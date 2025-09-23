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
# ---------------------
# DAILY SUMMARY TAB
# ---------------------
with tab1:
    st.subheader("📊 Daily Summary Report")

    # ---- Date Filter (combined) ----
    min_date, max_date = df["Order Date"].min(), df["Order Date"].max()

    col1, col2 = st.columns([2, 1])
    with col1:
        date_mode = st.radio("Date Selection Mode", ["None", "Single Date", "Date Range"], horizontal=True)
    with col2:
        date_group = st.selectbox("Date Group", ["All", "Last 7 Days", "Last 15 Days"])

    df_filtered = df.copy()

    if date_mode == "Single Date":
        single_date = st.date_input("Pick a Date", value=max_date, min_value=min_date, max_value=max_date)
        df_filtered = df_filtered[df_filtered["Order Date"] == single_date]

    elif date_mode == "Date Range":
        date_range = st.date_input("Pick a Date Range", value=(min_date, max_date),
                                   min_value=min_date, max_value=max_date)
        if isinstance(date_range, tuple) and len(date_range) == 2:
            df_filtered = df_filtered[(df_filtered["Order Date"] >= date_range[0]) & (df_filtered["Order Date"] <= date_range[1])]

    # ---- Date Group still applies after above ----
    if date_group == "Last 7 Days":
        df_filtered = df_filtered[df_filtered["Order Date"] >= date.today() - timedelta(days=7)]
    elif date_group == "Last 15 Days":
        df_filtered = df_filtered[df_filtered["Order Date"] >= date.today() - timedelta(days=15)]

    # ---- Other Filters (Smart Filtering) ----
    required_filters = [
        "Region","User","L4Position User","L3Position User","L2Position User",
        "Reporting Manager","Primary Category"
    ]

    for f in required_filters:
        if f in df_filtered.columns:
            vals = sorted(df_filtered[f].dropna().unique().tolist())
            vals = ["All"] + vals
            sel = st.multiselect(f, vals, default="All", key=f"f_{f}")
            if "All" not in sel:
                df_filtered = df_filtered[df_filtered[f].isin(sel)]

    # ---- Column Selection ----
    # curated list (Summary + Secondary)
curated_cols = [
    "Order Date","L4Position User","L3Position User","L2Position User","Region",
    "Reporting Manager","User","Selected Jw User","Type","Reason",
    "Tc","Pc","Ovc","First Call","Last Call","Total Retail Time(Hh:Mm)",
    "Ghee","Dw Primary Packs","Dw Consu","Dw Bulk","36 No","Smp","Gjm",
    "Cream","Uht Milk","Flavored Milk",
    "Distributor","Territory","Beat"
]

# keep only those present in df_filtered
allowed_cols = [c for c in curated_cols if c in df_filtered.columns]

cols_available = ["All"] + allowed_cols
selected_cols = st.multiselect("Columns Wants in Table", cols_available, default="All")

if "All" in selected_cols or not selected_cols:
    final_df = df_filtered[allowed_cols]
else:
    final_df = df_filtered[selected_cols]


    # ---- Results ----
    st.markdown("### Results Table (Top 200 Rows)")
    st.dataframe(final_df.head(200), width="stretch")

    # ---- Export ----
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


# ---------------------
# OUTLET WISE REPORT TAB
# ---------------------
with tab2:
    st.subheader("🏪 Outlet Wise Report")
    st.info("This section is a placeholder. You can plug in outlet-wise logic here.")
