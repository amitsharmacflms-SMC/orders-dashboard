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
    """
    Safer robust date parsing:
    - try several explicit formats
    - for numeric-like values, convert Excel serials (unit='d', origin='1899-12-30')
      but only for the numeric subset to avoid dtype-casting errors
    - fallback to generic parser
    Returns pd.Series of python date objects (or NaT)
    """
    s = series.copy()
    s_str = s.astype(str).str.strip()

    # 1) Try explicit common formats
    parsed = pd.to_datetime(s_str, format="%Y-%m-%d", errors="coerce")
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d", "%d-%b-%Y", "%d %b %Y", "%m/%d/%Y", "%d.%m.%Y"):
        parsed = parsed.fillna(pd.to_datetime(s_str, format=fmt, errors="coerce"))

    # 2) Try Excel-style serial numbers but only on numeric entries
    s_num = pd.to_numeric(series, errors="coerce")
    mask = s_num.notna()
    if mask.any():
        try:
            nums = s_num[mask].astype("float64").values
            converted = pd.to_datetime(nums, unit="d", origin="1899-12-30", errors="coerce")
            converted_series = pd.Series(converted, index=s_num[mask].index)
            parsed = parsed.fillna(converted_series)
        except Exception:
            converted_series = s_num[mask].apply(lambda x: pd.to_datetime(float(x), unit="d", origin="1899-12-30", errors="coerce"))
            parsed = parsed.fillna(converted_series)

    # 3) Final fallback to generic parser (dateutil)
    parsed = parsed.fillna(pd.to_datetime(s_str, errors="coerce"))

    return parsed.dt.date

def aggregate_secondary(df_secondary, join_keys):
    """Aggregate secondary to unique join_keys rows.
       Numeric columns -> sum, non-numeric -> first non-null."""
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
    """Parse times that are primarily HH:MM, fallback to generic parser."""
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

# Normalize headers
df_summary = normalize_columns(df_summary_raw)
df_secondary = normalize_columns(df_secondary_raw)

# If files use "Date" column name, rename to "Order Date"
if "Date" in df_summary.columns and "Order Date" not in df_summary.columns:
    df_summary = df_summary.rename(columns={"Date": "Order Date"})
if "Date" in df_secondary.columns and "Order Date" not in df_secondary.columns:
    df_secondary = df_secondary.rename(columns={"Date": "Order Date"})

# Show raw samples for debugging
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

# Robust parse Order Date in both dataframes if column exists
if "Order Date" in df_summary.columns:
    df_summary["Order Date_Parsed"] = robust_parse_date_col(df_summary["Order Date"])
    df_summary["Order Date_Raw"] = df_summary["Order Date"].astype(str)
    df_summary["Order Date"] = df_summary["Order Date_Parsed"]

if "Order Date" in df_secondary.columns:
    df_secondary["Order Date_Parsed"] = robust_parse_date_col(df_secondary["Order Date"])
    df_secondary["Order Date_Raw"] = df_secondary["Order Date"].astype(str)
    df_secondary["Order Date"] = df_secondary["Order Date_Parsed"]

# Normalize User fields (strip + uppercase) to avoid subtle mismatches
if "User" in df_summary.columns:
    df_summary["User"] = df_summary["User"].astype(str).str.strip().str.upper()
if "User" in df_secondary.columns:
    df_secondary["User"] = df_secondary["User"].astype(str).str.strip().str.upper()

# Debug: date coverage before merge
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

# ---------------------
# Diagnostic (pre-aggregation) to help if Secondary has multiple rows
# ---------------------
with st.expander("🔎 Pre-aggregation Secondary sample"):
    st.write("Secondary (raw) shape:", df_secondary.shape)
    st.write("Secondary (raw) columns (first 80):", list(df_secondary.columns)[:80])
    st.write("Secondary (raw) sample 20 rows:")
    st.write(df_secondary.head(20))

# ---------------------
# Decide merge keys and aggregate secondary as needed
# ---------------------
if "Order Date" in df_secondary.columns:
    join_keys = ["User", "Order Date"]
    st.info("Secondary.xlsx has an 'Order Date' column. Aggregating Secondary by (User, Order Date) and performing left-merge on these keys.")
    df_secondary_agg = aggregate_secondary(df_secondary, join_keys)
else:
    join_keys = ["User"]
    st.warning("Secondary.xlsx does NOT have 'Order Date'. Aggregating Secondary by 'User' only and falling back to merge on 'User'. This may attach aggregated Secondary data across multiple Summary dates for the same user.")
    df_secondary_agg = aggregate_secondary(df_secondary, join_keys)
    if "User" in df_secondary.columns:
        dup_counts = df_secondary["User"].value_counts().rename_axis("User").reset_index(name="Secondary_rows_count")
        with st.expander("🔎 Secondary rows per User (sample)"):
            st.write(dup_counts.head(200))

# Diagnostic: show aggregated secondary sample
with st.expander("🔎 Secondary agg sample & counts"):
    st.write("Secondary_agg shape:", df_secondary_agg.shape)
    st.write("Columns (first 80):", list(df_secondary_agg.columns)[:80])
    st.write(df_secondary_agg.head(20))

# ====== DIAGNOSTIC BLOCK: BEFORE MERGE ======
st.markdown("### 🔍 Secondary Merge Diagnostics (pre-merge)")

s = df_summary.copy()
t = df_secondary_agg.copy()

# normalized fields for diagnostics
s["__user__"] = s["User"].astype(str).str.strip().str.upper() if "User" in s.columns else None
t["__user__"] = t["User"].astype(str).str.strip().str.upper() if "User" in t.columns else None

s["__date__"] = pd.to_datetime(s["Order Date"], errors="coerce").dt.date if "Order Date" in s.columns else pd.NaT
t["__date__"] = pd.to_datetime(t["Order Date"], errors="coerce").dt.date if "Order Date" in t.columns else pd.NaT

st.write("Unique users in Summary (sample 20):", sorted(s["__user__"].dropna().unique())[:20])
st.write("Unique users in Secondary agg (sample 20):", sorted(t["__user__"].dropna().unique())[:20])
st.write("Counts -> summary:", len(s), " secondary_agg:", len(t))

merge_keys_debug = join_keys.copy()
st.write("Attempting debug merge on keys:", merge_keys_debug)
if merge_keys_debug:
    tmp = s.merge(t, on=merge_keys_debug, how="left", indicator=True)
    st.write("Merge indicator counts:", tmp["_merge"].value_counts().to_dict())
    st.write("Rows where nothing from Secondary matched (left_only) sample 20:")
    st.write(tmp[tmp["_merge"]=="left_only"][["User","Order Date"]].head(20))

    left_only_users = tmp[tmp["_merge"]=="left_only"]["__user__"].dropna().unique()[:5]
    for u in left_only_users:
        st.write(f"Secondary rows for USER (normalized) {u} (sample):")
        st.write(t[t["__user__"]==u].head(10))

    # per-user date mismatches (sample)
    if "Order Date" in df_secondary.columns:
        s_grp = s.groupby("__user__")["__date__"].agg(lambda x: set(x.dropna()))
        t_grp = t.groupby("__user__")["__date__"].agg(lambda x: set(x.dropna()))
        compare_users = list(s_grp.index)[:10]
        diffs = {}
        for u in compare_users:
            sd = s_grp.get(u, set())
            td = t_grp.get(u, set())
            only_in_summary = sorted(list(sd - td))[:10]
            only_in_secondary = sorted(list(td - sd))[:10]
            if only_in_summary or only_in_secondary:
                diffs[u] = {"only_in_summary": only_in_summary, "only_in_secondary": only_in_secondary}
        st.write("Per-user date mismatches (sample):")
        st.write(diffs)
else:
    st.warning("No merge keys available for debug (check 'User' presence).")

# ====== END DIAGNOSTIC BLOCK ======

# Perform left merge (keep all summary rows)
try:
    df = pd.merge(df_summary, df_secondary_agg, on=join_keys, how="left", suffixes=("_Sum", "_Sec"))
except Exception as e:
    st.error(f"Merge failed on {join_keys}: {e}")
    st.stop()

# Debug after merge
with st.expander("🔎 Merge result sample & counts"):
    st.write("Rows in Summary.xlsx:", len(df_summary))
    st.write("Rows in Secondary.xlsx (raw):", len(df_secondary))
    st.write("Rows in Secondary.xlsx (agg):", len(df_secondary_agg))
    st.write("Rows after merge:", len(df))
    st.write("Merged columns (first 120):", list(df.columns)[:120])
    st.write("Sample merged rows (first 20):")
    st.write(df.head(20))

# ---------------------
# Unify common columns (drop the original _Sum/_Sec)
# ---------------------
for col_base in [
    "Region", "Territory", "Reporting Manager", "Distributor",
    "L4Position User", "L3Position User", "L2Position User", "Primary Category"
]:
    df = unify_columns_and_drop(df, col_base)

# Ensure extra columns exist
extra_cols = [
    "Tc","Pc","Ovc","First Call","Last Call","Total Retail Time(Hh:Mm)",
    "Ghee","Dw Primary Packs","Dw Consu","Dw Bulk","36 No","Smp","Gjm","Cream","Uht Milk","Flavored Milk"
]
for c in extra_cols:
    if c not in df.columns:
        df[c] = 0 if ("Call" not in c and "Time" not in c) else pd.NaT

# ---------------------
# Compute retail time safely (First Call / Last Call)
# ---------------------
if {"First Call", "Last Call"}.issubset(df.columns):
    first_parsed = parse_time_series(df["First Call"])
    last_parsed = parse_time_series(df["Last Call"])

    diff_minutes = (last_parsed - first_parsed).dt.total_seconds() / 60
    diff_minutes = pd.to_numeric(diff_minutes, errors="coerce").fillna(0).clip(lower=0).astype(int)
    df["Total Retail Time(Hh:Mm)"] = diff_minutes.apply(lambda x: f"{x // 60:02d}:{x % 60:02d}")
else:
    if "Total Retail Time(Hh:Mm)" not in df.columns:
        df["Total Retail Time(Hh:Mm)"] = "00:00"

# ---------------------
# Filters + Date Range + Robust matching
# ---------------------
st.markdown("### Filters")

required_filters = [
    "Order Date","Region","Territory","L4Position User","L3Position User",
    "L2Position User","Reporting Manager","Primary Category","User"
]

filter_selections = {}
available_cols = {c.lower().replace(" ", "").replace("_",""): c for c in df.columns}

matched_filters = {}
missing_filters = []

# Date range special handling
if "orderdate" in available_cols:
    date_col = available_cols["orderdate"]
    min_date = df[date_col].dropna().min()
    max_date = df[date_col].dropna().max()
    if (pd.isna(min_date) or pd.isna(max_date)) and "Order Date" in df_summary.columns:
        min_date = df_summary["Order Date"].dropna().min()
        max_date = df_summary["Order Date"].dropna().max()
    if pd.notna(min_date) and pd.notna(max_date):
        start, end = st.date_input("Order Date Range", value=(min_date, max_date), min_value=min_date, max_value=max_date)
        filter_selections[date_col] = (start, end)
        matched_filters["Order Date"] = date_col
    else:
        st.info("Order Date range not available (parsing failed). See debug panels above.")
else:
    missing_filters.append("Order Date")

# Other filters
for f in required_filters:
    if f.lower().replace(" ", "").replace("_","") == "orderdate":
        continue
    key = f.lower().replace(" ", "").replace("_","")
    if key in available_cols:
        col = available_cols[key]
        matched_filters[f] = col
        vals = df[col].dropna().unique().tolist()
        if vals:
            sel = st.multiselect(f, sorted(vals), key=f"f_{f}")
            filter_selections[col] = sel
    else:
        missing_filters.append(f)

with st.expander("🔎 Filter Matching Debug"):
    st.write("✅ Matched filters:", matched_filters)
    st.write("⚠️ Missing filters (not found in merged data):", missing_filters)

# ---------------------
# Debug: Row counts at each step
# ---------------------
with st.expander("🔎 Row counts at each step"):
    st.write("Rows in Summary.xlsx:", len(df_summary))
    st.write("Rows in Secondary.xlsx (raw):", len(df_secondary))
    st.write("Rows in Secondary.xlsx (agg):", len(df_secondary_agg))
    st.write("Rows after merge:", len(df))

# ---------------------
# Apply filters
# ---------------------
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

# ---------------------
# Final locked column order
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
# KPIs, table, export
# ---------------------
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
st.download_button("Download Excel", to_excel_bytes(final_df), "filtered_export.xlsx", "application/vnd.openxmlformats-officedocument-spreadsheetml.sheet")
