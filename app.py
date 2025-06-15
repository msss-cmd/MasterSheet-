import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import io

st.set_page_config(layout="wide", page_title="SSS Master Sheet Dashboard")

st.title("SSS Master Sheet Reporting & Dashboard")
st.markdown("Upload your **single SSS Master Sheet Excel workbook** to get reports, visualizations, and a dashboard for its sheets.")

# --- Enhanced Helper Functions ---
def process_qt_register_2025(df: pd.DataFrame):
    st.header("1. Quotation Register 2025 Analysis")

    df['Date'] = df['Date'].astype(str).str.replace('//', '/', regex=False)
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df.dropna(subset=['Date'], inplace=True)

    for col in ['Company  Name', 'Product', 'Sales Person']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    st.subheader("Raw Data Preview")
    st.dataframe(df.head())

    st.subheader("Key Metrics & Reports")

    quotations_by_sales_person = df['Sales Person'].value_counts().reset_index()
    quotations_by_sales_person.columns = ['Sales Person', 'Number of Quotations']
    st.write("#### Number of Quotations by Sales Person:")
    st.dataframe(quotations_by_sales_person)

    quotations_by_product = df['Product'].value_counts().reset_index()
    quotations_by_product.columns = ['Product', 'Number of Quotations']
    st.write("#### Number of Quotations by Product:")
    st.dataframe(quotations_by_product.head(10))

    df['Month'] = df['Date'].dt.to_period('M')
    monthly_quotations = df['Month'].value_counts().sort_index().reset_index()
    monthly_quotations.columns = ['Month', 'Number of Quotations']
    monthly_quotations['Month'] = monthly_quotations['Month'].astype(str)
    st.write("#### Monthly Quotation Trends:")
    st.dataframe(monthly_quotations)

    # --- Enhanced Insights ---
    st.subheader("🧠 Advanced Sales Insights")

    if 'Status' in df.columns:
        win_rate = (df['Status'].str.lower() == 'won').mean()
        st.metric("Win Rate (%)", f"{win_rate * 100:.2f}")

    if 'Amount' in df.columns:
        try:
            df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
            avg_amount = df['Amount'].mean()
            st.metric("Average Quotation Value", f"$ {avg_amount:,.2f}")
        except:
            st.warning("Unable to parse 'Amount' column for average value computation.")

    if 'Submitted Date' in df.columns and 'Closed Date' in df.columns:
        df['Submitted Date'] = pd.to_datetime(df['Submitted Date'], errors='coerce')
        df['Closed Date'] = pd.to_datetime(df['Closed Date'], errors='coerce')
        df['Lead Time (days)'] = (df['Closed Date'] - df['Submitted Date']).dt.days
        avg_lead = df['Lead Time (days)'].mean()
        st.metric("Avg Lead Time (days)", f"{avg_lead:.1f}")

    if 'Company  Name' in df.columns and 'Amount' in df.columns:
        top_clients = df.groupby('Company  Name')['Amount'].sum().nlargest(5).reset_index()
        st.write("#### Top Clients by Quotation Amount:")
        st.dataframe(top_clients)

    # --- Plots ---
    st.subheader("Visualizations")

    fig1, ax1 = plt.subplots(figsize=(10, 6))
    sns.barplot(x='Number of Quotations', y='Sales Person', data=quotations_by_sales_person, palette='viridis', ax=ax1)
    ax1.set_title('Number of Quotations by Sales Person (2025)')
    st.pyplot(fig1)

    fig2, ax2 = plt.subplots(figsize=(12, 7))
    sns.barplot(x='Number of Quotations', y='Product', data=quotations_by_product.head(10), palette='magma', ax=ax2)
    ax2.set_title('Top 10 Products by Number of Quotations (2025)')
    st.pyplot(fig2)

    fig3, ax3 = plt.subplots(figsize=(10, 6))
    sns.lineplot(x='Month', y='Number of Quotations', data=monthly_quotations, marker='o', color='purple', ax=ax3)
    ax3.set_title('Monthly Quotation Trends (2025)')
    plt.xticks(rotation=45, ha='right')
    st.pyplot(fig3)

    if 'Sales Person' in df.columns and 'Product' in df.columns:
        st.write("#### Heatmap: Quotations by Product and Sales Person")
        heatmap_data = df.groupby(['Sales Person', 'Product']).size().unstack(fill_value=0)
        if not heatmap_data.empty:
            fig_heatmap, ax_heatmap = plt.subplots(figsize=(14, len(heatmap_data.index) * 0.7 + len(heatmap_data.columns) * 0.2))
            sns.heatmap(heatmap_data, annot=True, fmt="d", cmap="YlGnBu", linewidths=.5, ax=ax_heatmap)
            ax_heatmap.set_title('Number of Quotations per Product and Sales Person')
            plt.xticks(rotation=90)
            st.pyplot(fig_heatmap)

    if 'Date' in df.columns:
        daily_quotations = df.groupby(df['Date'].dt.date).size().reset_index(name='Daily Quotations')
        daily_quotations['Date'] = pd.to_datetime(daily_quotations['Date'])
        daily_quotations = daily_quotations.sort_values('Date')
        daily_quotations['Cumulative Quotations'] = daily_quotations['Daily Quotations'].cumsum()

        if not daily_quotations.empty:
            fig_cumulative, ax_cumulative = plt.subplots(figsize=(12, 6))
            sns.lineplot(x='Date', y='Cumulative Quotations', data=daily_quotations, marker='o', color='green', ax=ax_cumulative)
            ax_cumulative.set_title('Cumulative Number of Quotations (2025)')
            plt.xticks(rotation=45, ha='right')
            st.pyplot(fig_cumulative)
