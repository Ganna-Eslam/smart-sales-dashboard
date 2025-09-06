
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, time

# ---- CONFIG ----
EXCEL_PATH = "01K42FVQY4PC6DEM0C7ZF0X8GX.xlsx"
 

st.set_page_config(layout="wide", page_title="SmartMart Sales Dashboard (Streamlit)")

# ---- HELPERS ----
@st.cache_data(show_spinner=False)
def load_data(excel_path=EXCEL_PATH):
    xls = pd.ExcelFile(excel_path)
    # Expecting sheets: Orders, Products, Sales_Reps, Regions, Customers
    orders = xls.parse("Orders")
    products = xls.parse("Products") if "Products" in xls.sheet_names else pd.DataFrame()
    reps = xls.parse("Sales_Reps") if "Sales_Reps" in xls.sheet_names else pd.DataFrame()
    regions = xls.parse("Regions") if "Regions" in xls.sheet_names else pd.DataFrame()
    customers = xls.parse("Customers") if "Customers" in xls.sheet_names else pd.DataFrame()
    return orders, products, reps, regions, customers

def clean_and_enrich_orders(df_orders, products):
    df = df_orders.copy()
    # Ensure date/time columns
    if "Order_Date" in df.columns:
        df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce")
        df["Year"] = df["Order_Date"].dt.year
        df["Month"] = df["Order_Date"].dt.month
        df["Month_Name"] = df["Order_Date"].dt.strftime("%b")
        df["Day"] = df["Order_Date"].dt.day
        df["Day_of_Week"] = df["Order_Date"].dt.day_name()
    # If Hour is not present, try to extract from Time or Order_Date
    if "Hour" not in df.columns and "Time" in df.columns:
        # Try parsing Time column
        try:
            df["Hour"] = pd.to_datetime(df["Time"], errors="coerce").dt.hour
        except Exception:
            df["Hour"] = pd.to_numeric(df["Time"], errors="coerce").fillna(0).astype(int)
    if "Hour" not in df.columns and "Order_Date" in df.columns:
        df["Hour"] = df["Order_Date"].dt.hour.fillna(0).astype(int)
    # Numeric conversions
    for c in ["Quantity_Sold", "Unit_Price", "Total_Sales", "Profit"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    # Recompute Total_Sales if missing or inconsistent
    if "Total_Sales" not in df.columns or df["Total_Sales"].isnull().any():
        if "Quantity_Sold" in df.columns and "Unit_Price" in df.columns:
            df["Total_Sales_calc"] = df["Quantity_Sold"] * df["Unit_Price"]
            # if Total_Sales exists but has big differences, prefer recalculated
            if "Total_Sales" in df.columns:
                diff = (df["Total_Sales"] - df["Total_Sales_calc"]).abs()
                if (diff > 1e-6).sum() > 0:
                    df["Total_Sales"] = df["Total_Sales_calc"]
            else:
                df["Total_Sales"] = df["Total_Sales_calc"]
    # Ensure Return_Flag is 0/1
    if "Return_Flag" in df.columns:
        df["Return_Flag"] = df["Return_Flag"].apply(lambda x: 1 if str(x).strip() in ["1", "True", "true", "Y", "y"] else 0)
    else:
        df["Return_Flag"] = 0
    # Time of day classification
    def tod(h):
        try:
            h = int(h)
        except:
            return "Unknown"
        if 6 <= h <= 11:
            return "Morning"
        if 12 <= h <= 16:
            return "Afternoon"
        if 17 <= h <= 20:
            return "Evening"
        if (21 <= h <= 23) or (0 <= h <= 5):
            return "Night"
        return "Unknown"
    df["Time_of_Day"] = df["Hour"].apply(tod)
    # Join product categories if available
    if not products.empty and "Product_ID" in df.columns:
        product_cols = [c for c in products.columns if c.lower() in ["product_id","productid","id"]]+[c for c in products.columns if c.lower().startswith("product")]
        # safe merge on Product_ID column name
        if "Product_ID" in products.columns:
            df = df.merge(products, on="Product_ID", how="left", suffixes=('', '_prod'))
        elif len(product_cols)>0:
            prod_key = product_cols[0]
            df = df.merge(products, left_on="Product_ID", right_on=prod_key, how="left", suffixes=('', '_prod'))
    # compute Average Order Size as Total_Sales / Quantity_Sold (handle div by zero)
    df["Avg_Order_Size"] = df.apply(lambda r: r["Total_Sales"] / r["Quantity_Sold"] if r["Quantity_Sold"]>0 else r["Total_Sales"], axis=1)
    # Profit margin
    df["Profit_Margin"] = df.apply(lambda r: r["Profit"] / r["Total_Sales"] if r["Total_Sales"]>0 else 0, axis=1)
    return df

def filter_df(df, sel_category, sel_regions, sel_reps, date_range, sel_tod, products):
    d = df.copy()
    # Category filter: try Product Category column names
    if sel_category and len(sel_category)>0 and "Product_Category" in d.columns:
        d = d[d["Product_Category"].isin(sel_category)]
    elif sel_category and len(sel_category)>0:
        # try common column names
        possible = [c for c in d.columns if "category" in c.lower()]
        if possible:
            d = d[d[possible[0]].isin(sel_category)]
    # Region filter
    if sel_regions and len(sel_regions)>0 and "Region_ID" in d.columns:
        d = d[d["Region_ID"].isin(sel_regions)]
    # Sales Rep filter
    if sel_reps and len(sel_reps)>0 and "Sales_Rep_ID" in d.columns:
        d = d[d["Sales_Rep_ID"].isin(sel_reps)]
    # Time of day filter
    if sel_tod and len(sel_tod)>0:
        d = d[d["Time_of_Day"].isin(sel_tod)]
    # Date range filter
    if date_range is not None and "Order_Date" in d.columns:
        start_date, end_date = date_range
        d = d[(d["Order_Date"]>=pd.to_datetime(start_date)) & (d["Order_Date"]<=pd.to_datetime(end_date))]
    return d

# ---- LOAD ----
try:
    orders, products, reps, regions, customers = load_data()
except Exception as e:
    st.error(f"Failed to load data from {EXCEL_PATH}: {e}")
    st.stop()

orders = clean_and_enrich_orders(orders, products)

# Sidebar - Filters
st.sidebar.header("Filters")
# Product Category options
cat_col = None
possible_cat_cols = [c for c in orders.columns if "category" in c.lower()]
if "Product_Category" in orders.columns:
    cat_col = "Product_Category"
elif len(possible_cat_cols)>0:
    cat_col = possible_cat_cols[0]

if cat_col:
    categories = sorted(orders[cat_col].dropna().unique().tolist())
    sel_category = st.sidebar.multiselect("Product Category", options=categories, default=[])
else:
    sel_category = []

# Region options
region_vals = orders["Region_ID"].unique().tolist() if "Region_ID" in orders.columns else []
sel_regions = st.sidebar.multiselect("Region (ID)", options=sorted(region_vals), default=[])

# Sales rep options
rep_vals = orders["Sales_Rep_ID"].unique().tolist() if "Sales_Rep_ID" in orders.columns else []
sel_reps = st.sidebar.multiselect("Sales Rep (ID)", options=sorted(rep_vals), default=[])

# Time of day
tod_options = ["Morning","Afternoon","Evening","Night"]
sel_tod = st.sidebar.multiselect("Time of Day", options=tod_options, default=[])

# Date range
min_date = orders["Order_Date"].min() if "Order_Date" in orders.columns else pd.to_datetime("2000-01-01")
max_date = orders["Order_Date"].max() if "Order_Date" in orders.columns else pd.to_datetime("today")
date_range = st.sidebar.date_input("Date range", value=(min_date, max_date))

# Apply filters
df = filter_df(orders, sel_category, sel_regions, sel_reps, date_range, sel_tod, products)

# ---- KPIs ----
st.title("SmartMart Sales Dashboard (Streamlit)")
st.markdown("Interactive dashboard generated from uploaded Excel. Use the sidebar to filter data.")

total_sales = df["Total_Sales"].sum()
total_qty = df["Quantity_Sold"].sum() if "Quantity_Sold" in df.columns else 0
total_profit = df["Profit"].sum() if "Profit" in df.columns else 0
avg_order_size = (total_sales / total_qty) if total_qty>0 else 0
return_rate = (df["Return_Flag"].sum() / len(df)) if len(df)>0 else 0
profit_margin = (total_profit / total_sales) if total_sales>0 else 0

# Sales Growth calculations (MoM and YoY)
def calc_growth(df_in, period="M"):
    tmp = df_in.set_index("Order_Date").resample(period)["Total_Sales"].sum().sort_index()
    mom = tmp.pct_change().replace([np.inf, -np.inf], np.nan)
    yoy = tmp.groupby([tmp.index.month, tmp.index.year]).sum() if False else None
    return tmp, mom

monthly_sales, mom = calc_growth(df[df["Total_Sales"]>0], period="M")
latest_mom = mom.dropna().iloc[-1] if not mom.dropna().empty else 0

# KPI cards
col1, col2, col3, col4, col5, col6 = st.columns(6)
col1.metric("Total Sales", f"{total_sales:,.2f}")
col2.metric("Total Quantity", f"{total_qty:,d}")
col3.metric("Total Profit", f"{total_profit:,.2f}")
col4.metric("Avg Order Size", f"{avg_order_size:,.2f}")
col5.metric("Return Rate", f"{return_rate:.2%}")
col6.metric("Profit Margin", f"{profit_margin:.2%}", delta=f"{latest_mom:.2%} MoM")

# ---- Charts & Analysis ----
st.markdown("### Sales Trend")
fig_trend = px.line(monthly_sales.reset_index(), x="Order_Date", y="Total_Sales", title="Monthly Sales Trend", labels={"Order_Date":"Month","Total_Sales":"Total Sales"})
st.plotly_chart(fig_trend, use_container_width=True)

# Hourly heatmap / distribution
st.markdown("### Sales by Hour of Day (Heatmap)")
if "Hour" in df.columns:
    pivot = df.pivot_table(index="Hour", columns="Day_of_Week", values="Total_Sales", aggfunc="sum", fill_value=0)
    # Reorder days to week order
    days_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    pivot = pivot.reindex(columns=[d for d in days_order if d in pivot.columns], fill_value=0)
    fig_heat = px.imshow(pivot.values, x=pivot.columns, y=pivot.index, aspect="auto", labels=dict(x="Day", y="Hour", color="Total Sales"), title="Sales Heatmap (Hour vs Day)")
    st.plotly_chart(fig_heat, use_container_width=True)
else:
    st.info("No Hour column available to compute hourly heatmap.")

# Sales by Product Category (donut)
st.markdown("### Sales by Product Category")
if cat_col:
    cat_agg = df.groupby(cat_col)["Total_Sales"].sum().reset_index().sort_values("Total_Sales", ascending=False)
    fig_cat = px.pie(cat_agg, names=cat_col, values="Total_Sales", hole=0.4, title="Sales by Product Category")
    st.plotly_chart(fig_cat, use_container_width=True)
else:
    st.info("No category column available in the data to show Category breakdown.")

# Sales vs Returns by Region
st.markdown("### Sales vs Returns by Region (Top Regions)")
if "Region_ID" in df.columns:
    region_agg = df.groupby("Region_ID").agg(Total_Sales=("Total_Sales","sum"), Returns=("Return_Flag","sum")).reset_index()
    region_agg = region_agg.sort_values("Total_Sales", ascending=False).head(10)
    fig_sr = go.Figure(data=[go.Bar(name="Total Sales", x=region_agg["Region_ID"].astype(str), y=region_agg["Total_Sales"]),
                             go.Bar(name="Returns", x=region_agg["Region_ID"].astype(str), y=region_agg["Returns"])])
    fig_sr.update_layout(barmode='group', title="Sales vs Returns by Region")
    st.plotly_chart(fig_sr, use_container_width=True)
else:
    st.info("No Region_ID column found.")

# Profit margin over time
st.markdown("### Profit Margin Over Time (Monthly)")
if "Order_Date" in df.columns:
    pm = df.set_index("Order_Date").resample("M").agg(Total_Profit=("Profit","sum"), Total_Sales=("Total_Sales","sum"))
    pm["Profit_Margin"] = pm.apply(lambda r: r["Total_Profit"]/r["Total_Sales"] if r["Total_Sales"]>0 else 0, axis=1)
    fig_pm = px.line(pm.reset_index(), x="Order_Date", y="Profit_Margin", title="Monthly Profit Margin", labels={"Order_Date":"Month","Profit_Margin":"Profit Margin"})
    st.plotly_chart(fig_pm, use_container_width=True)
else:
    st.info("No Order_Date to compute profit margin over time.")

# Sales Rep Performance
st.markdown("### Sales Rep Performance (by Total Sales)")
if "Sales_Rep_ID" in df.columns:
    rep_agg = df.groupby("Sales_Rep_ID").agg(Total_Sales=("Total_Sales","sum"), Quantity=("Quantity_Sold","sum")).reset_index().sort_values("Total_Sales", ascending=False)
    st.dataframe(rep_agg.head(20))
    fig_reps = px.bar(rep_agg.head(10), x="Sales_Rep_ID", y="Total_Sales", title="Top 10 Sales Reps by Sales")
    st.plotly_chart(fig_reps, use_container_width=True)
else:
    st.info("No Sales_Rep_ID column found.")

# Top Products, Regions, Customers
st.markdown("### Top Entities")
colA, colB, colC = st.columns(3)
with colA:
    st.write("Top Products (by Sales)")
    if "Product_ID" in df.columns:
        ptop = df.groupby("Product_ID")["Total_Sales"].sum().reset_index().sort_values("Total_Sales", ascending=False).head(5)
        st.table(ptop)
    else:
        st.write("No Product_ID")

with colB:
    st.write("Top Regions (by Sales)")
    if "Region_ID" in df.columns:
        rtop = df.groupby("Region_ID")["Total_Sales"].sum().reset_index().sort_values("Total_Sales", ascending=False).head(5)
        st.table(rtop)
    else:
        st.write("No Region_ID")

with colC:
    st.write("Top Customers (by Sales)")
    if "Customer_ID" in df.columns:
        ctop = df.groupby("Customer_ID")["Total_Sales"].sum().reset_index().sort_values("Total_Sales", ascending=False).head(5)
        st.table(ctop)
    else:
        st.write("No Customer_ID present in Orders to compute top customers.")

# Insights & Recommendations (basic automated examples)
st.markdown("## Insights & Recommendations")
insights = []
# Insight 1: peak day
if "Day_of_Week" in df.columns:
    day_agg = df.groupby("Day_of_Week")["Total_Sales"].sum().reindex(["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"], fill_value=0)
    peak_day = day_agg.idxmax()
    insights.append(f"Peak sales day: **{peak_day}** with total sales of {day_agg.max():,.2f}.")
# Insight 2: peak time of day
tod_agg = df.groupby("Time_of_Day")["Total_Sales"].sum()
if not tod_agg.empty:
    peak_tod = tod_agg.idxmax()
    insights.append(f"Peak time of day: **{peak_tod}** with total sales of {tod_agg.max():,.2f}.")
# Insight 3: best product
if "Product_ID" in df.columns:
    best_prod = df.groupby("Product_ID")["Total_Sales"].sum().idxmax()
    insights.append(f"Top product by sales: **{best_prod}**.")

for i, ins in enumerate(insights[:3]):
    st.write(f"{i+1}. {ins}")

st.markdown("**Recommendations:**")
st.write("1. Focus marketing and promotions during peak time periods and days to amplify conversion when demand is highest.")
st.write("2. Prioritize inventory and cross-selling for top products/regions shown above to increase revenue and reduce stockouts.")

# Footer - show filtered data option
if st.checkbox("Show filtered data table"):
    st.write(df.head(200))
