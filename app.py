import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io

# Set page configuration
st.set_page_config(layout="wide", page_title="SSS Master Sheet Dashboard")

st.title("SSS Master Sheet Reporting & Dashboard")
st.markdown("Upload your Excel files to generate reports, visualizations, and a dashboard.")

# --- Helper Functions for Processing Each File Type ---

def process_qt_register_2025(df: pd.DataFrame):
    """
    Processes the 'QT Register 2025' data, performs analysis, and generates visualizations.
    """
    st.header("1. Quotation Register 2025 Analysis")

    # Data Cleaning for QT Register 2025
    df['Date'] = df['Date'].astype(str).str.replace('//', '/', regex=False)
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df.dropna(subset=['Date'], inplace=True)

    # Strip leading/trailing spaces from categorical columns
    df['Company  Name'] = df['Company  Name'].astype(str).str.strip()
    df['Product'] = df['Product'].astype(str).str.strip()
    df['Sales Person'] = df['Sales Person'].astype(str).str.strip()

    st.subheader("Raw Data Preview")
    st.dataframe(df.head())

    # Data Analysis
    st.subheader("Key Metrics & Reports")

    # 1. Quotations by Sales Person
    quotations_by_sales_person = df['Sales Person'].value_counts().reset_index()
    quotations_by_sales_person.columns = ['Sales Person', 'Number of Quotations']
    st.write("#### Number of Quotations by Sales Person:")
    st.dataframe(quotations_by_sales_person)

    # 2. Quotations by Product
    quotations_by_product = df['Product'].value_counts().reset_index()
    quotations_by_product.columns = ['Product', 'Number of Quotations']
    st.write("#### Number of Quotations by Product:")
    st.dataframe(quotations_by_product.head(10)) # Display top 10 for brevity

    # 3. Monthly Quotation Trends
    df['Month'] = df['Date'].dt.to_period('M')
    monthly_quotations = df['Month'].value_counts().sort_index().reset_index()
    monthly_quotations.columns = ['Month', 'Number of Quotations']
    monthly_quotations['Month'] = monthly_quotations['Month'].astype(str)
    st.write("#### Monthly Quotation Trends:")
    st.dataframe(monthly_quotations)

    # Visualizations
    st.subheader("Visualizations")

    # Plot 1: Quotations by Sales Person
    fig1, ax1 = plt.subplots(figsize=(10, 6))
    sns.barplot(x='Number of Quotations', y='Sales Person', data=quotations_by_sales_person, palette='viridis', ax=ax1)
    ax1.set_title('Number of Quotations by Sales Person (2025)')
    ax1.set_xlabel('Number of Quotations')
    ax1.set_ylabel('Sales Person')
    st.pyplot(fig1)

    # Plot 2: Top Products by Quotations
    fig2, ax2 = plt.subplots(figsize=(12, 7))
    sns.barplot(x='Number of Quotations', y='Product', data=quotations_by_product.head(10), palette='magma', ax=ax2)
    ax2.set_title('Top 10 Products by Number of Quotations (2025)')
    ax2.set_xlabel('Number of Quotations')
    ax2.set_ylabel('Product')
    st.pyplot(fig2)

    # Plot 3: Monthly Quotation Trends
    fig3, ax3 = plt.subplots(figsize=(10, 6))
    sns.lineplot(x='Month', y='Number of Quotations', data=monthly_quotations, marker='o', color='purple', ax=ax3)
    ax3.set_title('Monthly Quotation Trends (2025)')
    ax3.set_xlabel('Month')
    ax3.set_ylabel('Number of Quotations')
    plt.xticks(rotation=45, ha='right')
    st.pyplot(fig3)

def process_2025_inv(df: pd.DataFrame):
    """
    Processes the '2025 INV' data. (Placeholder for detailed analysis)
    """
    st.header("2. 2025 Invoices Analysis")

    # Basic Data Cleaning
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df.dropna(subset=['Date'], inplace=True)
    df.columns = df.columns.str.strip() # Clean column names

    st.subheader("Raw Data Preview")
    st.dataframe(df.head())

    st.subheader("Key Metrics & Reports (Example)")
    st.write("Total Invoices:", df['INV No.'].nunique())
    st.write("Unique Sales Persons:", df['Sales Person'].nunique())

    st.markdown("""
    **Further analysis could include:**
    - Total invoice amount over time (if 'Amount' column exists and is numeric)
    - Invoices by Reseller/End User
    - Top Suppliers and Products in invoices
    - Sales performance by Sales Person (based on invoice value)
    """)
    if 'Sales Person' in df.columns:
        st.write("#### Invoices by Sales Person:")
        st.bar_chart(df['Sales Person'].value_counts())

def process_meeting_agenda(df: pd.DataFrame):
    """
    Processes the 'Meeting Agenda' data. (Placeholder for detailed analysis)
    """
    st.header("3. Meeting Agenda Analysis")

    # Basic Data Cleaning for Meeting Agenda
    # Assuming 'Order Value Approx.' might be numeric, cleaning it
    df['Order Value Approx.'] = pd.to_numeric(
        df['Order Value Approx.'].astype(str).str.replace(',', ''), errors='coerce'
    )
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df.dropna(subset=['Date', 'Order Value Approx.'], inplace=True)
    df.columns = df.columns.str.strip()

    st.subheader("Raw Data Preview")
    st.dataframe(df.head())

    st.subheader("Key Metrics & Reports (Example)")
    st.write("Total Meeting Points:", df['No:'].nunique())
    st.write(f"Total Approximate Order Value: $ {df['Order Value Approx.'].sum():,.2f}")

    st.markdown("""
    **Further analysis could include:**
    - Average Order Value per meeting point
    - Meeting points by Action By person
    - Distribution of Margin
    - Trends in meeting topics or order values over time
    """)
    if 'Action By' in df.columns:
        st.write("#### Points by Action By:")
        st.bar_chart(df['Action By'].value_counts())
    if 'Margin' in df.columns and df['Margin'].dtype != 'object':
         st.write("#### Margin Distribution:")
         fig, ax = plt.subplots()
         sns.histplot(df['Margin'].dropna(), kde=True, ax=ax)
         st.pyplot(fig)


def process_payment_pending(df: pd.DataFrame):
    """
    Processes the 'Payment Pending' data. (Placeholder for detailed analysis)
    """
    st.header("4. Payment Pending Analysis")

    # Basic Data Cleaning for Payment Pending
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
    df.dropna(subset=['Amount', 'PARTY NAME'], inplace=True)
    df.columns = df.columns.str.strip()

    st.subheader("Raw Data Preview")
    st.dataframe(df.head())

    st.subheader("Key Metrics & Reports (Example)")
    total_pending_amount = df['Amount'].sum()
    st.write(f"Total Pending Amount: $ {total_pending_amount:,.2f}")
    st.write("Number of Parties with Pending Payments:", df['PARTY NAME'].nunique())

    st.markdown("""
    **Further analysis could include:**
    - Top parties with highest pending amounts
    - Distribution of pending amounts
    - Contact person analysis
    - Aging of payments (if a 'Due Date' or similar column is available)
    """)
    if 'PARTY NAME' in df.columns:
        st.write("#### Top 10 Parties by Pending Amount:")
        top_parties = df.groupby('PARTY NAME')['Amount'].sum().nlargest(10).reset_index()
        st.dataframe(top_parties)
        fig, ax = plt.subplots(figsize=(10, 6))
        sns.barplot(x='Amount', y='PARTY NAME', data=top_parties, palette='coolwarm', ax=ax)
        ax.set_title('Top 10 Parties by Pending Amount')
        st.pyplot(fig)


# --- Main Application Logic ---
uploaded_files = st.file_uploader(
    "Upload your SSS Master Sheet Excel files",
    type=["xlsx"], # Changed to accept Excel files
    accept_multiple_files=True
)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.sidebar.subheader(f"Processing: {uploaded_file.name}")
        try:
            # Read Excel with a specified header row based on initial inspection
            # Note: pd.read_excel might handle headers differently than pd.read_csv
            # The skiprows argument is still used for consistency based on previous CSV parsing.
            if "Payment Pending" in uploaded_file.name:
                df = pd.read_excel(uploaded_file, skiprows=1)
            elif "Quotation Register 2023" in uploaded_file.name:
                df = pd.read_excel(uploaded_file, skiprows=879)
            elif "Meeting Agenda" in uploaded_file.name:
                df = pd.read_excel(uploaded_file, skiprows=2)
            else: # Default for QT Register 2025, 2025 INV, or other general Excel sheets
                df = pd.read_excel(uploaded_file)

            # Use st.expander for a collapsible section for each file's analysis
            with st.expander(f"Analysis for {uploaded_file.name}"):
                if "QT Register 2025" in uploaded_file.name:
                    process_qt_register_2025(df.copy()) # Use .copy() to avoid SettingWithCopyWarning
                elif "2025 INV" in uploaded_file.name:
                    process_2025_inv(df.copy())
                elif "Meeting Agenda" in uploaded_file.name:
                    process_meeting_agenda(df.copy())
                elif "Payment Pending" in uploaded_file.name:
                    process_payment_pending(df.copy())
                else:
                    st.write(f"No specific processing logic defined for '{uploaded_file.name}'.")
                    st.dataframe(df.head()) # Show raw data if no specific logic

        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {e}")
            st.write("Please ensure the file is a valid Excel sheet and matches the expected structure.")

else:
    st.info("Please upload your Excel files to get started.")
