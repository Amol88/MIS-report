import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import datetime
import numpy as np

# Page setup
st.set_page_config(page_title="Dynamic MIS Dashboard", layout="wide")

st.title("📊 Dynamic MIS Dashboard")
st.markdown("Apni Excel ya CSV file upload karein aur real-time business performance track karein.")

# Sample data generator (testing ke liye)
@st.cache_data
def load_sample_data():
    np.random.seed(42)
    dates = pd.date_range(start="2026-01-01", end="2026-11-25", freq="D")
    categories = ["Electronics", "FMCG", "Hardware", "Services", "Consultancy"]
    cities = ["Mumbai", "Pune", "Delhi", "Ahmedabad", "Bengaluru"]
    
    data = []
    for _ in range(400):
        rev = np.random.randint(15000, 150000)
        exp = int(rev * np.random.uniform(0.55, 0.85))
        data.append({
            "Date": np.random.choice(dates),
            "Category": np.random.choice(categories),
            "City": np.random.choice(cities),
            "Revenue": rev,
            "Expense": exp
        })
    return pd.DataFrame(data)

# Auto-detect header row helper
def detect_header_row(file_obj, is_csv=False):
    file_obj.seek(0)
    try:
        if is_csv:
            preview = pd.read_csv(file_obj, header=None, nrows=25)
        else:
            preview = pd.read_excel(file_obj, header=None, nrows=25)
    except Exception:
        return 1

    keywords = [
        "date", "dt", "particular", "party", "customer", "voucher", "invoice",
        "bill", "amount", "total", "gross", "net", "sale", "revenue", "income",
        "expense", "cost", "purchase", "debit", "credit", "qty", "quantity",
        "rate", "gst", "tax", "city", "state", "region", "branch", "category",
        "product", "item"
    ]

    best_row = 0
    max_score = 0

    for idx, row in preview.iterrows():
        row_values = [str(val).strip().lower() for val in row.dropna()]
        score = sum(1 for val in row_values if any(k in val for k in keywords))
        if score > max_score:
            max_score = score
            best_row = idx

    return best_row + 1 if max_score >= 2 else 1

# Duplicate column names ko unique banane ka helper
def make_unique_columns(columns):
    seen = {}
    unique_cols = []
    for col in columns:
        col_str = str(col).strip()
        if col_str in seen:
            seen[col_str] += 1
            unique_cols.append(f"{col_str}_{seen[col_str]}")
        else:
            seen[col_str] = 0
            unique_cols.append(col_str)
    return unique_cols

# Indian Rupee formatting helper
def format_inr(num):
    return f"₹{num:,.0f}"

# Sidebar - File Upload
st.sidebar.header("1. Data Upload")
uploaded_file = st.sidebar.file_uploader("Excel (.xlsx) ya CSV file upload karein", type=["xlsx", "csv"])

if uploaded_file is not None:
    is_csv = uploaded_file.name.lower().endswith(".csv")
    detected_row = detect_header_row(uploaded_file, is_csv=is_csv)
    
    header_row = st.sidebar.number_input(
        "Header Row Number (Auto-detected)",
        min_value=1,
        max_value=25,
        value=int(detected_row),
        step=1,
        key=f"header_{uploaded_file.name}_{uploaded_file.size}"
    )
    
    try:
        uploaded_file.seek(0)
        if is_csv:
            raw_df = pd.read_csv(uploaded_file, header=header_row - 1)
        else:
            raw_df = pd.read_excel(uploaded_file, header=header_row - 1)
        
        raw_df = raw_df.dropna(how="all", axis=1).dropna(how="all", axis=0)
        raw_df.columns = make_unique_columns(raw_df.columns)
        st.sidebar.success(f"File load ho gayi! (Header Row: {header_row})")
    except Exception as e:
        st.error(f"File read karne mein error aaya: {e}")
        st.stop()
else:
    st.sidebar.info("Abhi Sample Data dikh raha hai. Apna data dekhne ke liye file upload karein.")
    raw_df = load_sample_data()

# Check if dataframe has enough columns
cols = raw_df.columns.tolist()
if len(cols) < 3:
    st.error("File mein kam se kam 3 columns (Date, Revenue, Expense) hone chahiye. Header Row Number check karein.")
    st.stop()

# Sidebar - Dynamic Column Mapping
st.sidebar.header("2. Column Mapping")

def get_default_index(keywords, fallback=0):
    for i, col in enumerate(cols):
        if any(k in str(col).lower() for k in keywords):
            return i
    return min(fallback, len(cols) - 1)

date_col = st.sidebar.selectbox("Date Column", cols, index=get_default_index(["date", "dt", "month", "day"], 0))
cat_col = st.sidebar.selectbox("Category / Party / Product Column", cols, index=get_default_index(["cat", "party", "particular", "customer", "prod", "item", "dept"], 1))
city_col = st.sidebar.selectbox("City / Branch / Region Column", cols, index=get_default_index(["city", "branch", "region", "state", "loc", "voucher"], 2))
rev_col = st.sidebar.selectbox("Revenue / Sales Column", cols, index=get_default_index(["rev", "sale", "income", "gross", "total", "amount", "credit"], 3))
exp_col = st.sidebar.selectbox("Expense / Cost Column", cols, index=get_default_index(["exp", "cost", "purch", "spend", "debit", "tax", "gst"], 4))

# Duplicate column check (crash rokne ke liye)
if len({date_col, rev_col, exp_col}) < 3:
    st.warning("⚠️ Sidebar mein **Date**, **Revenue**, aur **Expense** ke liye alag-alag columns select karein.")
    st.stop()

# Data cleaning & preparation
df = raw_df.copy()
df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
df[rev_col] = pd.to_numeric(df[rev_col], errors="coerce").fillna(0)
df[exp_col] = pd.to_numeric(df[exp_col], errors="coerce").fillna(0)
df = df.dropna(subset=[date_col])

if df.empty or not pd.api.types.is_datetime64_any_dtype(df[date_col]):
    st.error(f"❌ Selected column **'{date_col}'** mein valid dates nahi mili. Sidebar mein **Header Row Number** ya **Date Column** check karein.")
    st.stop()

# Blank text values ko safe string mein convert karna
df[cat_col] = df[cat_col].fillna("N/A").astype(str).str.strip()
df[city_col] = df[city_col].fillna("N/A").astype(str).str.strip()

df["Profit"] = df[rev_col] - df[exp_col]
df["Month"] = df[date_col].dt.to_period("M").astype(str)

# Sidebar - Interactive Filters
st.sidebar.header("3. Filters")
min_date = df[date_col].min().date()
max_date = df[date_col].max().date()

date_range = st.sidebar.date_input("Date Range", [min_date, max_date], min_value=min_date, max_value=max_date)

cat_options = sorted([str(x) for x in df[cat_col].unique().tolist()])
selected_cats = st.sidebar.multiselect(f"Select {cat_col}", options=cat_options, default=cat_options)

city_options = sorted([str(x) for x in df[city_col].unique().tolist()])
selected_cities = st.sidebar.multiselect(f"Select {city_col}", options=city_options, default=city_options)

# Apply filters
if len(date_range) == 2:
    start_dt, end_dt = date_range
    filtered_df = df[
        (df[date_col].dt.date >= start_dt) &
        (df[date_col].dt.date <= end_dt) &
        (df[cat_col].isin(selected_cats)) &
        (df[city_col].isin(selected_cities))
    ]
else:
    filtered_df = df[
        (df[cat_col].isin(selected_cats)) &
        (df[city_col].isin(selected_cities))
    ]

if filtered_df.empty:
    st.warning("Selected filters ke liye koi data available nahi hai. Filters change karke dekhein.")
    st.stop()

# KPI Section
total_rev = filtered_df[rev_col].sum()
total_exp = filtered_df[exp_col].sum()
total_profit = filtered_df["Profit"].sum()
profit_margin = (total_profit / total_rev * 100) if total_rev != 0 else 0

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("Total Revenue", format_inr(total_rev))
kpi2.metric("Total Expense", format_inr(total_exp))
kpi3.metric("Net Profit", format_inr(total_profit))
kpi4.metric("Profit Margin", f"{profit_margin:.2f}%")

st.markdown("---")

# Charts Row 1
col1, col2 = st.columns(2)

with col1:
    st.subheader("Monthly Revenue, Expense & Profit Trend")
    monthly_df = filtered_df.groupby("Month", as_index=False)[[rev_col, exp_col, "Profit"]].sum()
    fig_trend = go.Figure()
    fig_trend.add_trace(go.Bar(x=monthly_df["Month"], y=monthly_df[rev_col], name="Revenue", marker_color="#2E86C1"))
    fig_trend.add_trace(go.Bar(x=monthly_df["Month"], y=monthly_df[exp_col], name="Expense", marker_color="#E74C3C"))
    fig_trend.add_trace(go.Scatter(x=monthly_df["Month"], y=monthly_df["Profit"], name="Net Profit", mode="lines+markers", line=dict(color="#28B463", width=3)))
    fig_trend.update_layout(barmode="group", xaxis_title="Month", yaxis_title="Amount (₹)", legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig_trend, use_container_width=True)

with col2:
    st.subheader(f"{cat_col}-wise Revenue vs Profit")
    cat_df = filtered_df.groupby(cat_col, as_index=False)[[rev_col, "Profit"]].sum().sort_values(by=rev_col, ascending=False).head(15)
    fig_cat = px.bar(cat_df, x=cat_col, y=[rev_col, "Profit"], barmode="group", color_discrete_sequence=["#3498DB", "#2ECC71"])
    fig_cat.update_layout(yaxis_title="Amount (₹)", legend_title="Metric")
    st.plotly_chart(fig_cat, use_container_width=True)

# Charts Row 2
col3, col4 = st.columns(2)

with col3:
    st.subheader(f"{city_col}-wise Revenue Share")
    city_df = filtered_df.groupby(city_col, as_index=False)[rev_col].sum().sort_values(by=rev_col, ascending=False).head(10)
    fig_pie = px.pie(city_df, names=city_col, values=rev_col, hole=0.4)
    st.plotly_chart(fig_pie, use_container_width=True)

with col4:
    st.subheader(f"Summary Table ({cat_col}-wise)")
    summary_table = filtered_df.groupby(cat_col, as_index=False)[[rev_col, exp_col, "Profit"]].sum()
    summary_table["Margin (%)"] = (summary_table["Profit"] / summary_table[rev_col].replace(0, np.nan) * 100).fillna(0).round(2)
    summary_table = summary_table.sort_values(by=rev_col, ascending=False)
    st.dataframe(summary_table, use_container_width=True, hide_index=True)

st.markdown("---")

# Excel Export Feature
st.subheader("📥 Download MIS Report (Excel)")

def generate_excel(filtered_data, monthly_summary, category_summary):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        monthly_summary.to_excel(writer, index=False, sheet_name="Monthly_Summary")
        category_summary.to_excel(writer, index=False, sheet_name="Category_Summary")
        filtered_data.to_excel(writer, index=False, sheet_name="Filtered_Raw_Data")
    return output.getvalue()

excel_bytes = generate_excel(filtered_df, monthly_df, summary_table)

st.download_button(
    label="Download Filtered MIS Report (.xlsx)",
    data=excel_bytes,
    file_name=f"MIS_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
