"""
Enhanced Superstore Sales Dashboard
A comprehensive analytics dashboard for visualizing and analyzing superstore sales data
with dynamic themes, interactive filters, and advanced visualizations.
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go
import plotly.figure_factory as ff
import os
import warnings


# Suppress warnings for cleaner output
warnings.filterwarnings("ignore")

# ============================================================================
# PAGE CONFIGURATION
# ============================================================================
st.set_page_config(
    page_title="Superstore Sales Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# CUSTOM CSS STYLING
# ============================================================================
st.markdown("""
    <style>
    /* Main background and text styling */
    .main {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
    }
    
    /* Card styling for metrics */
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin: 10px 0;
    }
    
    /* Title styling */
    .dashboard-title {
        color: white;
        text-align: center;
        font-size: 3rem;
        font-weight: bold;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        margin-bottom: 1rem;
    }
    
    /* Subtitle styling */
    .dashboard-subtitle {
        color: #f0f0f0;
        text-align: center;
        font-size: 1.2rem;
        margin-bottom: 2rem;
    }
    
    /* Sidebar styling */
    .css-1d391kg {
        background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Remove padding from top */
    .block-container {
        padding-top: 1rem;
    }
    
    /* Custom expander styling */
    .streamlit-expanderHeader {
        background-color: #667eea;
        color: white;
        border-radius: 5px;
    }
    </style>
""", unsafe_allow_html=True)

# ============================================================================
# HEADER SECTION
# ============================================================================
st.markdown('<h1 class="dashboard-title">📊 Superstore Sales Analytics</h1>', unsafe_allow_html=True)
st.markdown('<p class="dashboard-subtitle">Unlock insights from your sales data with interactive visualizations and real-time analytics</p>', unsafe_allow_html=True)

# ============================================================================
# SIDEBAR - THEME SELECTION
# ============================================================================
st.sidebar.title("⚙️ Dashboard Settings")
theme = st.sidebar.selectbox(
    "Select Visualization Theme:",
    ["plotly", "plotly_white", "plotly_dark", "ggplot2", "seaborn", "simple_white"]
)

# ============================================================================
# FILE UPLOAD SECTION
# ============================================================================
st.sidebar.markdown("---")
st.sidebar.subheader("📁 Data Upload")

f1 = st.file_uploader(
    "Upload your Superstore sales data:",
    type=["csv", "xlsx", "xls", "txt"],
    help="Supported formats: CSV, Excel (XLSX/XLS), TXT"
)

# Check if file is uploaded, otherwise stop execution
if f1 is not None:
    file_name = f1.name
    st.sidebar.success(f"✅ Loaded: {file_name}")
    
    # Read the uploaded file based on its extension
    try:
        if file_name.endswith('.csv') or file_name.endswith('.txt'):
            df = pd.read_csv(f1, encoding="ISO-8859-1")
        else:
            df = pd.read_excel(f1)
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()
else:
    st.warning("⚠️ Please upload a Superstore sales data file to proceed.")
    st.info("💡 **Tip**: Your file should contain columns like Order Date, Region, State, City, Category, Sales, etc.")
    st.stop()

# ============================================================================
# DATA PREPROCESSING
# ============================================================================
# Convert Order Date to datetime format
df["Order Date"] = pd.to_datetime(df["Order Date"], errors='coerce')

# Remove any rows with invalid dates
df = df.dropna(subset=["Order Date"])

# Extract min and max dates for date range filter
startDate = df["Order Date"].min()
endDate = df["Order Date"].max()

# ============================================================================
# DOWNLOAD FULL DATASET OPTION
# ============================================================================
st.sidebar.markdown("---")
st.sidebar.subheader("📥 Download Dataset")

# Convert full dataset to CSV
full_dataset_csv = df.to_csv(index=False).encode('utf-8')

st.sidebar.download_button(
    label="⬇️ Download Full Dataset",
    data=full_dataset_csv,
    file_name='superstore_full_dataset.csv',
    mime='text/csv',
    help='Download the complete dataset as CSV file for external use'
)

# ============================================================================
# DATE RANGE FILTER
# ============================================================================
st.sidebar.markdown("---")
st.sidebar.subheader("📅 Date Range Filter")

col1, col2 = st.columns(2)

with col1:
    date1 = pd.to_datetime(st.date_input("Start Date:", startDate))

with col2:
    date2 = pd.to_datetime(st.date_input("End Date:", endDate))

# Filter dataframe based on selected date range
df = df[(df["Order Date"] >= date1) & (df["Order Date"] <= date2)].copy()

# ============================================================================
# SIDEBAR FILTERS - REGION, STATE, CITY
# ============================================================================
st.sidebar.markdown("---")
st.sidebar.subheader("🌍 Location Filters")

# Region filter
region = st.sidebar.multiselect(
    "Pick Region(s):",
    options=df["Region"].unique(),
    help="Filter by one or more regions"
)

# Create filtered dataframe based on region selection
if not region:
    df2 = df.copy()
else:
    df2 = df[df["Region"].isin(region)]

# State filter (based on selected regions)
state = st.sidebar.multiselect(
    "Pick State(s):",
    options=df2["State"].unique(),
    help="Filter by one or more states"
)

if not state:
    df3 = df2.copy()
else:
    df3 = df2[df2["State"].isin(state)]

# City filter (based on selected states)
city = st.sidebar.multiselect(
    "Pick City(ies):",
    options=df3["City"].unique(),
    help="Filter by one or more cities"
)

if not city:
    df4 = df3.copy()
else:
    df4 = df3[df3["City"].isin(city)]

# ============================================================================
# APPLY ALL FILTERS TO CREATE FINAL FILTERED DATAFRAME
# ============================================================================
# Complex filtering logic to handle all combinations of filters
if not region and not state and not city:
    filtered_df = df
elif not state and not city:
    filtered_df = df[df["Region"].isin(region)]
elif not region and not city:
    filtered_df = df[df["State"].isin(state)]
elif state and city:
    filtered_df = df3[df["State"].isin(state) & df3["City"].isin(city)]
elif region and city:
    filtered_df = df3[df["Region"].isin(region) & df3["City"].isin(city)]
elif region and state:
    filtered_df = df3[df["Region"].isin(region) & df3["State"].isin(state)]
elif city:
    filtered_df = df3[df3["City"].isin(city)]
else:
    filtered_df = df3[df3["Region"].isin(region) & df3["State"].isin(state) & df3["City"].isin(city)]

# ============================================================================
# KEY PERFORMANCE INDICATORS (KPIs)
# ============================================================================
st.markdown("### 📈 Key Performance Indicators")

# Calculate KPI metrics
total_sales = filtered_df["Sales"].sum()
total_profit = filtered_df["Profit"].sum()
total_quantity = filtered_df["Quantity"].sum()
profit_margin = (total_profit / total_sales * 100) if total_sales > 0 else 0
avg_order_value = filtered_df.groupby("Order ID")["Sales"].sum().mean() if "Order ID" in filtered_df.columns else total_sales / len(filtered_df)

# Display KPIs in columns
kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)

with kpi1:
    st.metric(
        label="💰 Total Sales",
        value=f"${total_sales:,.2f}",
        delta=f"{len(filtered_df)} orders"
    )

with kpi2:
    st.metric(
        label="📊 Total Profit",
        value=f"${total_profit:,.2f}",
        delta=f"{profit_margin:.1f}% margin"
    )

with kpi3:
    st.metric(
        label="📦 Total Quantity",
        value=f"{int(total_quantity):,}",
        delta="units sold"
    )

with kpi4:
    st.metric(
        label="💵 Avg Order Value",
        value=f"${avg_order_value:,.2f}",
        delta="per order"
    )

with kpi5:
    # Calculate number of unique customers if Customer ID exists
    unique_customers = filtered_df["Customer ID"].nunique() if "Customer ID" in filtered_df.columns else len(filtered_df)
    st.metric(
        label="👥 Customers",
        value=f"{unique_customers:,}",
        delta="unique"
    )

st.markdown("---")

# ============================================================================
# CATEGORY AND CITY ANALYSIS
# ============================================================================
col1, col2 = st.columns(2)

# Category-wise sales aggregation
category_df = filtered_df.groupby(by=["Category"], as_index=False)["Sales"].sum()

with col1:
    st.subheader("🏷️ Category-wise Sales Performance")
    
    # Create bar chart with gradient colors
    fig = px.bar(
        category_df,
        x="Category",
        y="Sales",
        text=['${:,.0f}'.format(x) for x in category_df["Sales"]],
        template=theme,
        color="Sales",
        color_continuous_scale="Blues"
    )
    
    fig.update_traces(textposition='outside')
    fig.update_layout(
        showlegend=False,
        height=450,
        xaxis_title="Category",
        yaxis_title="Sales ($)"
    )
    
    st.plotly_chart(fig, use_container_width=True)

with col2:
    st.subheader("🏙️ Top 10 Cities by Sales")
    
    # Aggregate top 10 cities by sales
    top_cities = (
        filtered_df.groupby(by=["City"], as_index=False)["Sales"]
        .sum()
        .sort_values(by="Sales", ascending=False)
        .head(10)
    )
    
    # Create horizontal bar chart for better readability
    fig = px.bar(
        top_cities,
        x="Sales",
        y="City",
        text=['${:,.0f}'.format(x) for x in top_cities["Sales"]],
        template=theme,
        color="Sales",
        color_continuous_scale="Viridis",
        orientation='h'
    )
    
    fig.update_traces(textposition='outside')
    fig.update_layout(
        showlegend=False,
        height=450,
        xaxis_title="Sales ($)",
        yaxis_title="City",
        yaxis={'categoryorder':'total ascending'}
    )
    
    st.plotly_chart(fig, use_container_width=True)

# ============================================================================
# DATA TABLES WITH DOWNLOAD OPTIONS
# ============================================================================
cl1, cl2 = st.columns(2)

with cl1:
    with st.expander("📋 Category Sales Data Table"):
        st.write(
            category_df.style.background_gradient(cmap='YlGnBu')
            .format({'Sales': '${:,.2f}'})
        )
        csv = category_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            "⬇️ Download Category Data",
            data=csv,
            file_name='category_sales_data.csv',
            mime='text/csv',
            help='Download category sales data as CSV file'
        )

with cl2:
    with st.expander("📋 Top 10 Cities Data Table"):
        st.write(
            top_cities.style.background_gradient(cmap='YlGnBu')
            .format({'Sales': '${:,.2f}'})
        )
        csv = top_cities.to_csv(index=False).encode('utf-8')
        st.download_button(
            "⬇️ Download Top Cities Data",
            data=csv,
            file_name='top_cities_sales_data.csv',
            mime='text/csv',
            help='Download top cities sales data as CSV file'
        )

st.markdown("---")

# ============================================================================
# TIME SERIES ANALYSIS
# ============================================================================
st.subheader("📅 Sales Trends Over Time")

# Create month-year column for time series analysis
filtered_df["month_year"] = filtered_df["Order Date"].dt.to_period("M")

# Aggregate sales by month
linechart = pd.DataFrame(
    filtered_df.groupby(filtered_df["month_year"].dt.strftime("%Y : %b"))["Sales"].sum()
).reset_index()

# Create area chart for time series
fig2 = go.Figure()

fig2.add_trace(go.Scatter(
    x=linechart["month_year"],
    y=linechart["Sales"],
    mode='lines+markers',
    name='Sales',
    line=dict(color='#667eea', width=3),
    marker=dict(size=8, color='#764ba2'),
    fill='tonexty',
    fillcolor='rgba(102, 126, 234, 0.2)'
))

fig2.update_layout(
    template=theme,
    height=500,
    xaxis_title="Month-Year",
    yaxis_title="Sales ($)",
    hovermode='x unified'
)

st.plotly_chart(fig2, use_container_width=True)

# Time series data table with download option
with st.expander("📊 View Time Series Data"):
    st.write(linechart.T.style.background_gradient(cmap="Blues"))
    csv = linechart.to_csv(index=False).encode("utf-8")
    st.download_button(
        '⬇️ Download Time Series Data',
        data=csv,
        file_name="TimeSeries.csv",
        mime='text/csv'
    )

st.markdown("---")

# ============================================================================
# HIERARCHICAL TREEMAP VISUALIZATION
# ============================================================================
st.subheader("🗺️ Hierarchical Sales View - Region → Category → Sub-Category")

fig3 = px.treemap(
    filtered_df,
    path=["Region", "Category", "Sub-Category"],
    values="Sales",
    hover_data=["Sales"],
    color="Sales",
    color_continuous_scale="RdYlGn"
)

fig3.update_layout(
    height=650
)

st.plotly_chart(fig3, use_container_width=True)

st.markdown("---")

# ============================================================================
# SEGMENT AND CATEGORY PIE CHARTS
# ============================================================================
chart1, chart2 = st.columns(2)

with chart1:
    st.subheader("🎯 Sales Distribution by Segment")
    
    segment_df = filtered_df.groupby("Segment")["Sales"].sum().reset_index()
    
    fig = px.pie(
        segment_df,
        values="Sales",
        names="Segment",
        template=theme,
        hole=0.4,  # Creates a donut chart
        color_discrete_sequence=px.colors.sequential.RdBu
    )
    
    fig.update_traces(
        textposition='inside',
        textinfo='percent+label',
        textfont_size=12
    )
    
    st.plotly_chart(fig, use_container_width=True)

with chart2:
    st.subheader("📦 Sales Distribution by Category")
    
    category_pie_df = filtered_df.groupby("Category")["Sales"].sum().reset_index()
    
    fig = px.pie(
        category_pie_df,
        values="Sales",
        names="Category",
        template=theme,
        hole=0.4,
        color_discrete_sequence=px.colors.sequential.Viridis
    )
    
    fig.update_traces(
        textposition='inside',
        textinfo='percent+label',
        textfont_size=12
    )
    
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# ============================================================================
# REGIONAL PERFORMANCE COMPARISON
# ============================================================================
st.subheader("🌎 Regional Performance Comparison")

# Aggregate data by region
region_performance = filtered_df.groupby("Region").agg({
    "Sales": "sum",
    "Profit": "sum",
    "Quantity": "sum"
}).reset_index()

# Create grouped bar chart
fig_region = go.Figure()

fig_region.add_trace(go.Bar(
    name='Sales',
    x=region_performance['Region'],
    y=region_performance['Sales'],
    marker_color='#667eea'
))

fig_region.add_trace(go.Bar(
    name='Profit',
    x=region_performance['Region'],
    y=region_performance['Profit'],
    marker_color='#764ba2'
))

fig_region.update_layout(
    barmode='group',
    template=theme,
    height=400,
    xaxis_title="Region",
    yaxis_title="Amount ($)"
)

st.plotly_chart(fig_region, use_container_width=True)

st.markdown("---")

# ============================================================================
# MONTHLY SUB-CATEGORY ANALYSIS
# ============================================================================
st.subheader("📊 Month-wise Sub-Category Sales Analysis")

with st.expander("📋 Sample Data Preview"):
    df_sample = df.head(5)[["Region", "State", "City", "Category", "Sales", "Profit", "Quantity"]]
    fig_table = ff.create_table(df_sample, colorscale="Cividis")
    st.plotly_chart(fig_table, use_container_width=True)

# Create pivot table for month-wise sub-category sales
st.markdown("#### 📅 Monthly Sub-Category Sales Matrix")
filtered_df["month"] = filtered_df["Order Date"].dt.month_name()

sub_category_year = pd.pivot_table(
    data=filtered_df,
    values="Sales",
    index=["Sub-Category"],
    columns="month"
)

st.write(sub_category_year.style.background_gradient(cmap="Blues").format("${:,.0f}"))

# Download option for pivot table
csv_pivot = sub_category_year.to_csv().encode('utf-8')
st.download_button(
    "⬇️ Download Monthly Sub-Category Data",
    data=csv_pivot,
    file_name="monthly_subcategory_sales.csv",
    mime='text/csv'
)

# ============================================================================
# PROFIT ANALYSIS SECTION
# ============================================================================
st.markdown("---")
st.subheader("💹 Profit Analysis")

profit_col1, profit_col2 = st.columns(2)

with profit_col1:
    # Top 10 profitable products
    st.markdown("#### 🏆 Top 10 Profitable Products")
    
    if "Product Name" in filtered_df.columns:
        top_products = (
            filtered_df.groupby("Product Name")["Profit"]
            .sum()
            .sort_values(ascending=False)
            .head(10)
            .reset_index()
        )
        
        fig_profit = px.bar(
            top_products,
            x="Profit",
            y="Product Name",
            orientation='h',
            template=theme,
            color="Profit",
            color_continuous_scale="Greens"
        )
        
        fig_profit.update_layout(
            height=400,
            yaxis={'categoryorder': 'total ascending'}
        )
        
        st.plotly_chart(fig_profit, use_container_width=True)

with profit_col2:
    # Profit margin by category
    st.markdown("#### 📊 Profit Margin by Category")
    
    category_profit = filtered_df.groupby("Category").agg({
        "Sales": "sum",
        "Profit": "sum"
    }).reset_index()
    
    category_profit["Profit Margin %"] = (
        category_profit["Profit"] / category_profit["Sales"] * 100
    )
    
    fig_margin = px.bar(
        category_profit,
        x="Category",
        y="Profit Margin %",
        template=theme,
        color="Profit Margin %",
        color_continuous_scale="RdYlGn",
        text=category_profit["Profit Margin %"].apply(lambda x: f'{x:.1f}%')
    )
    
    fig_margin.update_traces(textposition='outside')
    fig_margin.update_layout(height=400)
    
    st.plotly_chart(fig_margin, use_container_width=True)

# ============================================================================
# FOOTER
# ============================================================================
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: white;'>
        <p>Amartay Kumar Dhar</p>
        <p>Email: antukumardhar100@gmail.com</p>
    </div>
    """,
    unsafe_allow_html=True
)