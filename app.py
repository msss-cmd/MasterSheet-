import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io

# Set page configuration
st.set_page_config(layout="wide", page_title="SSS Master Sheet Dashboard")

st.title("SSS Master Sheet Reporting & Dashboard")
st.markdown("Upload your **single SSS Master Sheet Excel workbook** to get reports, visualizations, dashboard, and perform CRUD operations on its sheets.")

# --- Helper Functions for Processing Each File Type (now by Sheet Name) ---

def process_qt_register_2025(df: pd.DataFrame):
    """
    Processes the 'QT Register 2025' data, performs analysis, generates visualizations,
    and includes CRUD operations.
    """
    st.header("1. Quotation Register 2025 Analysis & CRUD")

    # Data Cleaning for QT Register 2025 (applied once on load)
    df['Date'] = df['Date'].astype(str).str.replace('//', '/', regex=False)
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    # Filling NaT with a placeholder or dropping, for robust analysis.
    # For CRUD, we might want to keep them or prompt user to fill. Dropping for analysis now.
    df.dropna(subset=['Date'], inplace=True)

    # Strip leading/trailing spaces from categorical columns
    for col in ['Company  Name', 'Product', 'Sales Person']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # --- CRUD Operations ---
    st.subheader("Perform CRUD Operations")
    st.info("Changes made here are applied to the data in your current session. Use the 'Download Modified Excel' button to save them.")

    # Store DataFrame in session state for persistence across interactions
    if 'qt_register_2025_df' not in st.session_state:
        st.session_state['qt_register_2025_df'] = df.copy() # Use a copy to avoid modifying original df directly

    # Display and allow editing
    st.write("#### Edit Data Directly:")
    edited_df = st.data_editor(st.session_state['qt_register_2025_df'], num_rows="dynamic")
    if not edited_df.equals(st.session_state['qt_register_2025_df']):
        st.session_state['qt_register_2025_df'] = edited_df
        st.success("Data updated in session.")

    # Add New Row
    st.write("#### Add New Row:")
    with st.form("add_qt_row_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            new_sl = st.number_input("Sl", min_value=1, value=len(st.session_state['qt_register_2025_df']) + 1)
            new_date = st.date_input("Date")
        with col2:
            new_ref_nos = st.text_input("Ref Nos.")
            new_company_name = st.text_input("Company Name")
        with col3:
            new_product = st.text_input("Product")
            new_sales_person = st.text_input("Sales Person")

        submitted = st.form_submit_button("Add Row")
        if submitted:
            new_row_data = {
                'Sl': new_sl,
                'Date': new_date,
                'Ref Nos.': new_ref_nos,
                'Company  Name': new_company_name,
                'Product': new_product,
                'Sales Person': new_sales_person
            }
            new_row_df = pd.DataFrame([new_row_data])
            st.session_state['qt_register_2025_df'] = pd.concat([st.session_state['qt_register_2025_df'], new_row_df], ignore_index=True)
            st.success("New row added.")
            st.rerun() # Rerun to refresh the data_editor and plots

    # Delete Row
    st.write("#### Delete Row by Index:")
    with st.form("delete_qt_row_form"):
        row_to_delete_index = st.number_input(
            "Enter the index of the row to delete (first row is 0)",
            min_value=0,
            max_value=len(st.session_state['qt_register_2025_df']) - 1,
            value=0
        )
        delete_submitted = st.form_submit_button("Delete Row")
        if delete_submitted:
            if row_to_delete_index in st.session_state['qt_register_2025_df'].index:
                st.session_state['qt_register_2025_df'].drop(row_to_delete_index, inplace=True)
                st.session_state['qt_register_2025_df'].reset_index(drop=True, inplace=True) # Reset index after dropping
                st.success(f"Row at index {row_to_delete_index} deleted.")
                st.rerun() # Rerun to refresh the data_editor and plots
            else:
                st.warning("Invalid index. Please enter a valid row index.")

    # Use the current state of the DataFrame for analysis and visualization
    current_df = st.session_state['qt_register_2025_df'].copy()

    # --- Data Analysis (using current_df) ---
    st.subheader("Key Metrics & Reports")

    # 1. Quotations by Sales Person
    quotations_by_sales_person = current_df['Sales Person'].value_counts().reset_index()
    quotations_by_sales_person.columns = ['Sales Person', 'Number of Quotations']
    st.write("#### Number of Quotations by Sales Person:")
    st.dataframe(quotations_by_sales_person)

    # 2. Quotations by Product
    quotations_by_product = current_df['Product'].value_counts().reset_index()
    quotations_by_product.columns = ['Product', 'Number of Quotations']
    st.write("#### Number of Quotations by Product:")
    st.dataframe(quotations_by_product.head(10)) # Display top 10 for brevity

    # 3. Monthly Quotation Trends
    current_df['Month'] = current_df['Date'].dt.to_period('M')
    monthly_quotations = current_df['Month'].value_counts().sort_index().reset_index()
    monthly_quotations.columns = ['Month', 'Number of Quotations']
    monthly_quotations['Month'] = monthly_quotations['Month'].astype(str)
    st.write("#### Monthly Quotation Trends:")
    st.dataframe(monthly_quotations)

    # Visualizations (using current_df)
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
    Processes the '2025 INV' data and allows editing. (Placeholder for full CRUD)
    """
    st.header("2. 2025 Invoices Analysis & Editing")

    # Basic Data Cleaning
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df.dropna(subset=['Date'], inplace=True)
    df.columns = df.columns.str.strip() # Clean column names

    if 'inv_2025_df' not in st.session_state:
        st.session_state['inv_2025_df'] = df.copy()

    st.subheader("Raw Data Preview & Editing")
    edited_df = st.data_editor(st.session_state['inv_2025_df'], num_rows="dynamic")
    if not edited_df.equals(st.session_state['inv_2025_df']):
        st.session_state['inv_2025_df'] = edited_df
        st.success("Data updated in session.")

    st.subheader("Key Metrics & Reports (Example)")
    current_df = st.session_state['inv_2025_df'].copy()
    st.write("Total Invoices:", current_df['INV No.'].nunique())
    st.write("Unique Sales Persons:", current_df['Sales Person'].nunique())

    st.markdown("""
    **Further analysis could include:**
    - Total invoice amount over time (if 'Amount' column exists and is numeric)
    - Invoices by Reseller/End User
    - Top Suppliers and Products in invoices
    - Sales performance by Sales Person (based on invoice value)
    """)
    if 'Sales Person' in current_df.columns:
        st.write("#### Invoices by Sales Person:")
        st.bar_chart(current_df['Sales Person'].value_counts())

def process_meeting_agenda(df: pd.DataFrame):
    """
    Processes the 'Meeting Agenda' data and allows editing. (Placeholder for full CRUD)
    """
    st.header("3. Meeting Agenda Analysis & Editing")

    # Basic Data Cleaning for Meeting Agenda
    df['Order Value Approx.'] = pd.to_numeric(
        df['Order Value Approx.'].astype(str).str.replace(',', ''), errors='coerce'
    )
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df.dropna(subset=['Date', 'Order Value Approx.'], inplace=True)
    df.columns = df.columns.str.strip()

    if 'meeting_agenda_df' not in st.session_state:
        st.session_state['meeting_agenda_df'] = df.copy()

    st.subheader("Raw Data Preview & Editing")
    edited_df = st.data_editor(st.session_state['meeting_agenda_df'], num_rows="dynamic")
    if not edited_df.equals(st.session_state['meeting_agenda_df']):
        st.session_state['meeting_agenda_df'] = edited_df
        st.success("Data updated in session.")

    st.subheader("Key Metrics & Reports (Example)")
    current_df = st.session_state['meeting_agenda_df'].copy()
    st.write("Total Meeting Points:", current_df['No:'].nunique())
    st.write(f"Total Approximate Order Value: $ {current_df['Order Value Approx.'].sum():,.2f}")

    st.markdown("""
    **Further analysis could include:**
    - Average Order Value per meeting point
    - Meeting points by Action By person
    - Distribution of Margin
    - Trends in meeting topics or order values over time
    """)
    if 'Action By' in current_df.columns:
        st.write("#### Points by Action By:")
        st.bar_chart(current_df['Action By'].value_counts())
    if 'Margin' in current_df.columns and current_df['Margin'].dtype != 'object':
         st.write("#### Margin Distribution:")
         fig, ax = plt.subplots()
         sns.histplot(current_df['Margin'].dropna(), kde=True, ax=ax)
         st.pyplot(fig)


def process_payment_pending(df: pd.DataFrame):
    """
    Processes the 'Payment Pending' data and allows editing. (Placeholder for full CRUD)
    """
    st.header("4. Payment Pending Analysis & Editing")

    # Basic Data Cleaning for Payment Pending
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
    df.dropna(subset=['Amount', 'PARTY NAME'], inplace=True)
    df.columns = df.columns.str.strip()

    if 'payment_pending_df' not in st.session_state:
        st.session_state['payment_pending_df'] = df.copy()

    st.subheader("Raw Data Preview & Editing")
    edited_df = st.data_editor(st.session_state['payment_pending_df'], num_rows="dynamic")
    if not edited_df.equals(st.session_state['payment_pending_df']):
        st.session_state['payment_pending_df'] = edited_df
        st.success("Data updated in session.")

    st.subheader("Key Metrics & Reports (Example)")
    current_df = st.session_state['payment_pending_df'].copy()
    total_pending_amount = current_df['Amount'].sum()
    st.write(f"Total Pending Amount: $ {total_pending_amount:,.2f}")
    st.write("Number of Parties with Pending Payments:", current_df['PARTY NAME'].nunique())

    st.markdown("""
    **Further analysis could include:**
    - Top parties with highest pending amounts
    - Distribution of pending amounts
    - Contact person analysis
    - Aging of payments (if a 'Due Date' or similar column is available)
    """)
    if 'PARTY NAME' in current_df.columns:
        st.write("#### Top 10 Parties by Pending Amount:")
        top_parties = current_df.groupby('PARTY NAME')['Amount'].sum().nlargest(10).reset_index()
        st.dataframe(top_parties)
        fig, ax = plt.subplots(figsize=(10, 6))
        sns.barplot(x='Amount', y='PARTY NAME', data=top_parties, palette='coolwarm', ax=ax)
        ax.set_title('Top 10 Parties by Pending Amount')
        st.pyplot(fig)

def process_quotation_register_2023(df: pd.DataFrame):
    """
    Processes the 'Quotation Register 2023' data and allows editing.
    Note: This sheet has generic column names, so specific analysis is limited without more info.
    """
    st.header("5. Quotation Register 2023 Analysis & Editing")
    st.warning("This sheet has generic column names (e.g., Unnamed: 0), making specific analysis challenging without more context.")

    if 'quotation_register_2023_df' not in st.session_state:
        st.session_state['quotation_register_2023_df'] = df.copy()

    st.subheader("Raw Data Preview & Editing")
    edited_df = st.data_editor(st.session_state['quotation_register_2023_df'], num_rows="dynamic")
    if not edited_df.equals(st.session_state['quotation_register_2023_df']):
        st.session_state['quotation_register_2023_df'] = edited_df
        st.success("Data updated in session.")

    st.write("Columns detected:", st.session_state['quotation_register_2023_df'].columns.tolist())
    st.markdown("""
    To perform meaningful analysis on this sheet, please provide more information about what each column represents.
    """)

# --- Function to convert DataFrame to Excel Bytes ---
def to_excel_bytes(dfs_dict: dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dfs_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    processed_data = output.getvalue()
    return processed_data

# --- Main Application Logic ---

# Dictionary to map sheet names to their respective skiprows values
# Based on the original CSV headers provided in the problem description
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

# Initialize session state for dataframes if not already present
if 'dataframes' not in st.session_state:
    st.session_state['dataframes'] = {}
    st.session_state['original_uploaded_file_name'] = None # To store original filename

if uploaded_file:
    # If a new file is uploaded, clear previous session state dataframes
    if uploaded_file.name != st.session_state['original_uploaded_file_name']:
        st.session_state['dataframes'] = {}
        st.session_state['original_uploaded_file_name'] = uploaded_file.name
        # Clear specific sheet DFs from session state too if new file is uploaded
        for key in list(st.session_state.keys()):
            if '_df' in key: # Assuming specific sheet DFs end with _df
                del st.session_state[key]


    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names

        st.sidebar.subheader("Select Sheets to Analyze & Edit")
        selected_sheets = st.sidebar.multiselect(
            "Choose which sheets to analyze:",
            options=sheet_names,
            default=[s for s in sheet_names if any(keyword in s for keyword in ["QT Register 2025", "2025 INV", "Meeting Agenda", "Payment Pending", "Quotation Register 2023"])]
        )

        if not selected_sheets:
            st.info("Please select at least one sheet from the sidebar to begin analysis.")
        else:
            current_processed_dfs = {} # To hold DataFrames that have been processed in this run

            for sheet_name in selected_sheets:
                st.sidebar.write(f"Loading sheet: {sheet_name}")

                # Use session state to manage the DataFrame for the current sheet
                # This allows changes to persist across interactions
                session_state_key = sheet_name.replace(" ", "_").lower() + "_df"

                if session_state_key not in st.session_state:
                    skip_rows = sheet_skiprows_map.get(sheet_name, 0)
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=skip_rows)
                    st.session_state[session_state_key] = df.copy()
                else:
                    df = st.session_state[session_state_key]

                with st.expander(f"Analysis & CRUD for Sheet: '{sheet_name}'"):
                    if "QT Register 2025" in sheet_name:
                        process_qt_register_2025(df) # Pass the session state df reference
                    elif "2025 INV" in sheet_name:
                        process_2025_inv(df)
                    elif "Meeting Agenda" in sheet_name:
                        process_meeting_agenda(df)
                    elif "Payment Pending" in sheet_name:
                        process_payment_pending(df)
                    elif "Quotation Register 2023" in sheet_name:
                        process_quotation_register_2023(df)
                    else:
                        st.write(f"No specific processing logic defined for sheet: '{sheet_name}'. Displaying raw data and enabling general editing.")
                        edited_df = st.data_editor(df, num_rows="dynamic")
                        if not edited_df.equals(df):
                            st.session_state[session_state_key] = edited_df
                            st.success("Data updated in session.")

                current_processed_dfs[sheet_name] = st.session_state[session_state_key]

            # --- Download Button for Modified Excel ---
            st.sidebar.markdown("---")
            st.sidebar.subheader("Download Modified Data")
            if current_processed_dfs:
                st.sidebar.download_button(
                    label="Download Modified Excel",
                    data=to_excel_bytes(current_processed_dfs),
                    file_name=f"Modified_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.sidebar.success("Download button active!")
            else:
                st.sidebar.info("No sheets processed yet for download.")

    except Exception as e:
        st.error(f"Error processing the Excel file: {e}")
        st.write("Please ensure the uploaded file is a valid .xlsx workbook and that the sheet names and data formats are consistent.")

else:
    st.info("Please upload your Excel workbook (.xlsx) to get started.")
