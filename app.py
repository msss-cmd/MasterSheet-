import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io

# Set page configuration
st.set_page_config(layout="wide", page_title="SSS Master Sheet Dashboard")

# Inject dark theme CSS
st.markdown("""
<style>
body, .stApp {
    background-color: #121212;
    color: #E0E0E0;
    font-family: 'Segoe UI', sans-serif;
}
h1, h2, h3, h4 {
    color: #ffffff;
    font-weight: 600;
}
section[data-testid="stSidebar"] {
    background-color: #1f1f1f;
    color: #f0f0f0;
}
.stMarkdown p {
    color: #d1d1d1;
}
button {
    background-color: #1abc9c;
    color: white;
    border-radius: 6px;
    padding: 8px 16px;
    font-weight: bold;
    transition: 0.3s;
}
button:hover {
    background-color: #16a085;
}
thead tr th {
    background-color: #2c2c2c !important;
    color: #f1f1f1 !important;
}
tbody tr td {
    background-color: #1e1e1e !important;
    color: #f1f1f1 !important;
}
.metric-card {
    background-color: #1f1f1f;
    padding: 20px;
    margin: 10px 0;
    border-radius: 10px;
    box-shadow: 0 0 10px rgba(0,0,0,0.4);
    transition: transform 0.2s;
}
.metric-card:hover {
    transform: scale(1.02);
}
details > summary {
    background-color: #2d2d2d;
    color: #ffffff;
    border: 1px solid #444;
    padding: 10px;
    border-radius: 5px;
    cursor: pointer;
}
details[open] > summary {
    background-color: #3a3a3a;
}
details > * {
    background-color: #181818 !important;
}
div[data-testid="stFileUploader"] > label {
    color: #cccccc;
    font-weight: 500;
}
.css-1cpxqw2, .stCaption, .stSubheader {
    color: #aaaaaa !important;
}
</style>
""", unsafe_allow_html=True)

# App header
st.markdown("""
<div style='text-align: center; padding: 30px 10px;'>
    <h1 style='color: #00ffe1; font-size: 42px;'>📊 SSS Master Sheet Dashboard</h1>
    <p style='font-size: 18px; color: #aaaaaa;'>Upload your <strong>SSS Master Sheet Excel workbook</strong> to analyze quotations, invoices, meetings, and payments in a beautiful dark interface.</p>
</div>
""", unsafe_allow_html=True)
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io

# Set page configuration
st.set_page_config(layout="wide", page_title="SSS Master Sheet Dashboard")

st.title("SSS Master Sheet Reporting & Dashboard")
st.markdown("Upload your **single SSS Master Sheet Excel workbook** to get reports, visualizations, and a dashboard for its sheets.")

# --- Helper Functions for Processing Each File Type (by Sheet Name) ---

def process_qt_register_2025(df: pd.DataFrame):
    """
    Processes the 'QT Register 2025' data, performs analysis, and generates visualizations.
    Includes advanced visualizations.
    """
    st.header("1. Quotation Register 2025 Analysis")

    # Data Cleaning for QT Register 2025
    df['Date'] = df['Date'].astype(str).str.replace('//', '/', regex=False)
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df.dropna(subset=['Date'], inplace=True)

    # Strip leading/trailing spaces from categorical columns
    for col in ['Company  Name', 'Product', 'Sales Person']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

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

    # --- Advanced Visualizations ---
    st.subheader("Advanced Visualizations")

    # Plot 1: Quotations by Sales Person (Basic)
    fig1, ax1 = plt.subplots(figsize=(10, 6))
    sns.barplot(x='Number of Quotations', y='Sales Person', data=quotations_by_sales_person, palette='viridis', ax=ax1)
    ax1.set_title('Number of Quotations by Sales Person (2025)')
    ax1.set_xlabel('Number of Quotations')
    ax1.set_ylabel('Sales Person')
    st.pyplot(fig1)

    # Plot 2: Top Products by Quotations (Basic)
    fig2, ax2 = plt.subplots(figsize=(12, 7))
    sns.barplot(x='Number of Quotations', y='Product', data=quotations_by_product.head(10), palette='magma', ax=ax2)
    ax2.set_title('Top 10 Products by Number of Quotations (2025)')
    ax2.set_xlabel('Number of Quotations')
    ax2.set_ylabel('Product')
    st.pyplot(fig2)

    # Plot 3: Monthly Quotation Trends (Basic)
    fig3, ax3 = plt.subplots(figsize=(10, 6))
    sns.lineplot(x='Month', y='Number of Quotations', data=monthly_quotations, marker='o', color='purple', ax=ax3)
    ax3.set_title('Monthly Quotation Trends (2025)')
    ax3.set_xlabel('Month')
    ax3.set_ylabel('Number of Quotations')
    plt.xticks(rotation=45, ha='right')
    st.pyplot(fig3)

    # Advanced Visualization 1: Heatmap of Quotations by Product & Sales Person
    if 'Sales Person' in df.columns and 'Product' in df.columns:
        st.write("#### Heatmap: Quotations by Product and Sales Person")
        # Create a pivot table to count quotations per Sales Person and Product
        heatmap_data = df.groupby(['Sales Person', 'Product']).size().unstack(fill_value=0)
        
        if not heatmap_data.empty:
            fig_heatmap, ax_heatmap = plt.subplots(figsize=(14, len(heatmap_data.index) * 0.7 + len(heatmap_data.columns) * 0.2)) # Dynamic sizing
            sns.heatmap(heatmap_data, annot=True, fmt="d", cmap="YlGnBu", linewidths=.5, ax=ax_heatmap)
            ax_heatmap.set_title('Number of Quotations per Product and Sales Person')
            ax_heatmap.set_xlabel('Product')
            ax_heatmap.set_ylabel('Sales Person')
            plt.xticks(rotation=90)
            plt.yticks(rotation=0)
            plt.tight_layout() # Adjust layout to prevent labels overlapping
            st.pyplot(fig_heatmap)
        else:
            st.info("No data available to generate Heatmap for Product and Sales Person.")


    # Advanced Visualization 2: Cumulative Sum of Quotations Over Time
    if 'Date' in df.columns:
        st.write("#### Cumulative Sum of Quotations Over Time")
        daily_quotations = df.groupby(df['Date'].dt.date).size().reset_index(name='Daily Quotations')
        daily_quotations['Date'] = pd.to_datetime(daily_quotations['Date'])
        daily_quotations = daily_quotations.sort_values('Date')
        daily_quotations['Cumulative Quotations'] = daily_quotations['Daily Quotations'].cumsum()

        if not daily_quotations.empty:
            fig_cumulative, ax_cumulative = plt.subplots(figsize=(12, 6))
            sns.lineplot(x='Date', y='Cumulative Quotations', data=daily_quotations, marker='o', color='green', ax=ax_cumulative)
            ax_cumulative.set_title('Cumulative Number of Quotations (2025)')
            ax_cumulative.set_xlabel('Date')
            ax_cumulative.set_ylabel('Cumulative Quotations')
            plt.xticks(rotation=45, ha='right')
            st.pyplot(fig_cumulative)
        else:
            st.info("No date data available to generate Cumulative Sum of Quotations.")


def process_2025_inv(df: pd.DataFrame):
    """
    Processes the '2025 INV' data.
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
    Processes the 'Meeting Agenda' data.
    """
    st.header("3. Meeting Agenda Analysis")

    # Basic Data Cleaning for Meeting Agenda
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
    Processes the 'Payment Pending' data.
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

def process_quotation_register_2023(df: pd.DataFrame):
    """
    Processes the 'Quotation Register 2023' data.
    Note: This sheet has generic column names, so specific analysis is limited without more info.
    """
    st.header("5. Quotation Register 2023 Analysis")
    st.warning("This sheet has generic column names (e.g., Unnamed: 0), making specific analysis challenging without more context.")
    st.subheader("Raw Data Preview")
    st.dataframe(df.head())
    st.write("Columns detected:", df.columns.tolist())
    st.markdown("""
    To perform meaningful analysis on this sheet, please provide more information about what each column represents.
    """)

# --- Main Application Logic ---

# Dictionary to map sheet names to their respective skiprows values
sheet_skiprows_map = {
    "Payment Pending": 1,
    "Quotation Register 2023": 879,
    "Meeting Agenda": 2,
    # For '2025 INV' and 'QT Register 2025', assuming header is at row 0 (default)
}

uploaded_file = st.file_uploader(
    "Upload your SSS Master Sheet Excel workbook (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=False
)

if uploaded_file:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names

        st.sidebar.subheader("Select Sheets to Analyze")
        selected_sheets = st.sidebar.multiselect(
            "Choose which sheets to analyze:",
            options=sheet_names,
            default=[s for s in sheet_names if any(keyword in s for keyword in ["QT Register 2025", "2025 INV", "Meeting Agenda", "Payment Pending", "Quotation Register 2023"])]
        )

        if not selected_sheets:
            st.info("Please select at least one sheet from the sidebar to begin analysis.")
        else:
            for sheet_name in selected_sheets:
                st.sidebar.write(f"Loading sheet: {sheet_name}")

                skip_rows = sheet_skiprows_map.get(sheet_name, 0)
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=skip_rows)

                with st.expander(f"Analysis for Sheet: '{sheet_name}'"):
                    if "QT Register 2025" in sheet_name:
                        process_qt_register_2025(df.copy())
                    elif "2025 INV" in sheet_name:
                        process_2025_inv(df.copy())
                    elif "Meeting Agenda" in sheet_name:
                        process_meeting_agenda(df.copy())
                    elif "Payment Pending" in sheet_name:
                        process_payment_pending(df.copy())
                    elif "Quotation Register 2023" in sheet_name:
                        process_quotation_register_2023(df.copy())
                    else:
                        st.write(f"No specific processing logic defined for sheet: '{sheet_name}'. Displaying raw data.")
                        st.dataframe(df.head())

    except Exception as e:
        st.error(f"Error processing the Excel file: {e}")
        st.write("Please ensure the uploaded file is a valid .xlsx workbook and that the sheet names and data formats are consistent.")

else:
    st.info("Please upload your Excel workbook (.xlsx) to get started.")
