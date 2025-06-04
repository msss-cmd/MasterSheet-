# app.py (part 1 - setup and input)

import streamlit as st
import pandas as pd
import plotly.express as px
import openai

st.set_page_config(page_title="Company Dashboard", layout="wide")

st.title("📊 Company Master Excel Dashboard")

# Upload Excel file
uploaded_file = st.file_uploader("📁 Upload your Master Excel File (.xlsx)", type=["xlsx"])

# OpenAI API key input
openai_api_key = st.text_input("🔐 Enter your OpenAI API Key", type="password")

# Validate input
if not uploaded_file:
    st.warning("Please upload an Excel file to continue.")
elif not openai_api_key:
    st.warning("Please enter your OpenAI API key.")
else:
    openai.api_key = openai_api_key

    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        selected_sheet = st.selectbox("📄 Select a sheet to explore", sheet_names)
        df = xls.parse(selected_sheet)
        df.columns = df.columns.str.strip()  # Clean column names
        st.success(f"✅ Loaded sheet: {selected_sheet}")
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()
# app.py (part 2 - QT Register 2025)

if selected_sheet == "QT Register 2025":
    st.header("📄 Dashboard: QT Register 2025")

    # Parse date
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors='coerce')

    # KPIs
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Quotations", len(df))
    col2.metric("Unique Clients", df["Company"].nunique() if "Company" in df else "N/A")
    col3.metric("Sales People", df["Sales Person"].nunique() if "Sales Person" in df else "N/A")

    st.markdown("---")

    # 📈 Monthly Quotation Trend
    if "Date" in df.columns:
        trend = df.groupby(df["Date"].dt.to_period("M")).size().reset_index(name='Count')
        trend["Date"] = trend["Date"].astype(str)
        st.plotly_chart(px.bar(trend, x="Date", y="Count", title="📈 Monthly Quotations"))

    # 👤 By Sales Person
    if "Sales Person" in df.columns:
        sales = df["Sales Person"].value_counts().reset_index()
        sales.columns = ["Sales Person", "Count"]
        st.plotly_chart(px.bar(sales, x="Sales Person", y="Count", title="👤 Quotations by Sales Person"))

    # 📦 By Product
    if "Product" in df.columns:
        st.plotly_chart(px.pie(df, names="Product", title="📦 Product Distribution"))

    # 📋 Full Table
    st.markdown("### Full Quotation Table")
    st.dataframe(df, use_container_width=True)
# app.py (part 3 - 2025 INV)

if selected_sheet == "2025 INV":
    st.header("🧾 Dashboard: 2025 Invoices")

    # Clean column names and parse dates
    df.columns = df.columns.str.strip()
    if "Invoice Date" in df.columns:
        df["Invoice Date"] = pd.to_datetime(df["Invoice Date"], errors='coerce')

    # Check critical columns
    required_columns = ["Client", "Amount", "Invoice Date"]
    for col in required_columns:
        if col not in df.columns:
            st.warning(f"Missing column: {col} — dashboard may be incomplete.")

    # KPIs
    total_invoices = len(df)
    total_revenue = df["Amount"].sum() if "Amount" in df.columns else 0
    unique_clients = df["Client"].nunique() if "Client" in df.columns else 0

    col1, col2, col3 = st.columns(3)
    col1.metric("📦 Total Invoices", total_invoices)
    col2.metric("💰 Total Revenue", f"₹ {total_revenue:,.2f}")
    col3.metric("👤 Unique Clients", unique_clients)

    st.markdown("---")

    # 📈 Monthly Revenue Trend
    if "Invoice Date" in df.columns and "Amount" in df.columns:
        rev_trend = df.groupby(df["Invoice Date"].dt.to_period("M"))["Amount"].sum().reset_index()
        rev_trend["Invoice Date"] = rev_trend["Invoice Date"].astype(str)
        st.plotly_chart(px.line(rev_trend, x="Invoice Date", y="Amount", title="📈 Monthly Revenue Trend"))

    # 👤 Revenue by Client
    if "Client" in df.columns and "Amount" in df.columns:
        client_rev = df.groupby("Client")["Amount"].sum().sort_values(ascending=False).head(10).reset_index()
        st.plotly_chart(px.bar(client_rev, x="Client", y="Amount", title="🏢 Top 10 Clients by Revenue"))

    # ✅ Paid vs Unpaid (if applicable)
    if "Status" in df.columns:
        status_count = df["Status"].value_counts().reset_index()
        status_count.columns = ["Status", "Count"]
        st.plotly_chart(px.pie(status_count, names="Status", values="Count", title="✅ Invoice Status Distribution"))

    # 📋 Full Invoice Table
    st.markdown("### Full Invoice Table")
    st.dataframe(df, use_container_width=True)
# app.py (part 4 - Payment Pending)

if selected_sheet == "Payment Pending":
    st.header("💸 Dashboard: Payment Pending")

    # Clean columns and parse dates
    df.columns = df.columns.str.strip()

    if "Invoice Date" in df.columns:
        df["Invoice Date"] = pd.to_datetime(df["Invoice Date"], errors='coerce')
    if "Due Date" in df.columns:
        df["Due Date"] = pd.to_datetime(df["Due Date"], errors='coerce')

    # Default to 0 if Amount column is missing
    df["Amount"] = pd.to_numeric(df.get("Amount", 0), errors='coerce').fillna(0)

    # KPI Metrics
    total_pending = df["Amount"].sum()
    total_clients = df["Client"].nunique() if "Client" in df.columns else 0
    overdue_count = df[df["Due Date"] < pd.Timestamp.today()].shape[0] if "Due Date" in df.columns else "N/A"

    col1, col2, col3 = st.columns(3)
    col1.metric("💰 Total Pending (BHD)", f"BHD {total_pending:,.2f}")
    col2.metric("👥 Clients with Pending", total_clients)
    col3.metric("⏰ Overdue Invoices", overdue_count)

    st.markdown("---")

    # 🔻 Overdue Aging Analysis
    if "Due Date" in df.columns:
        today = pd.Timestamp.today()
        df["Days Overdue"] = (today - df["Due Date"]).dt.days
        aging_bins = [0, 30, 60, 90, 180, 365, float('inf')]
        aging_labels = ["0-30", "31-60", "61-90", "91-180", "181-365", "365+"]
        df["Aging Bucket"] = pd.cut(df["Days Overdue"], bins=aging_bins, labels=aging_labels, right=False)

        aging_summary = df.groupby("Aging Bucket")["Amount"].sum().reset_index()
        st.plotly_chart(px.bar(aging_summary, x="Aging Bucket", y="Amount", title="📊 Payment Aging (BHD)"))

    # 🧍 Top Clients with Pending Dues
    if "Client" in df.columns:
        top_clients = df.groupby("Client")["Amount"].sum().sort_values(ascending=False).head(10).reset_index()
        st.plotly_chart(px.bar(top_clients, x="Client", y="Amount", title="🏢 Top 10 Clients by Outstanding Payment"))

    # 📋 Full Payment Pending Table
    st.markdown("### Full Payment Pending Table")
    st.dataframe(df, use_container_width=True)
