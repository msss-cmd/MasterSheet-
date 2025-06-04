import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
# No longer importing os or dotenv for API key

# --- Page Configuration ---
st.set_page_config(
    page_title="Salahuddin Softtech Solutions Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 SSS Internal Business Dashboard")
st.markdown("Welcome to the multi-page dashboard for Salahuddin Softtech Solutions.")

# --- Sidebar Inputs ---
st.sidebar.title("Configuration & Data Upload")

# OpenAI API Key Input
st.sidebar.subheader("API Key")
openai_api_key_input = st.sidebar.text_input(
    "Enter your OpenAI API Key",
    type="password", # Mask the input for security
    key="openai_api_key_input" # Unique key for this widget
)

# Store API key in session state once provided
if openai_api_key_input:
    st.session_state.openai_api_key = openai_api_key_input
else:
    st.session_state.openai_api_key = None # Clear if input is empty

# Excel File Uploader
st.sidebar.subheader("Upload Data")
uploaded_file = st.sidebar.file_uploader("Upload your SSS Master Sheet.xlsx", type=["xlsx"])

# Initialize session state for data storage and API key
if "data_loaded" not in st.session_state:
    st.session_state.data_loaded = False
if "sheets" not in st.session_state:
    st.session_state.sheets = {}
if "openai_api_key" not in st.session_state:
    st.session_state.openai_api_key = None

# Process uploaded file
if uploaded_file is not None:
    try:
        # Load all sheets into a dictionary of DataFrames
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        st.session_state.sheets = all_sheets
        st.session_state.data_loaded = True # Flag to indicate data is loaded

        st.sidebar.success("Excel file loaded successfully!")
        st.sidebar.write(f"Sheets detected: {', '.join(all_sheets.keys())}")

    except Exception as e:
        st.sidebar.error(f"Error loading Excel file: {e}")
        st.session_state.data_loaded = False # Reset data on error
        st.session_state.sheets = {}
else:
    # Reset data if no file is uploaded or file is cleared
    st.session_state.data_loaded = False
    st.session_state.sheets = {}


# --- Helper Function for Data Preprocessing ---
def preprocess_data(sheets):
    processed_sheets = {}

    # Preprocess Quotation Registers
    qt_dfs = []
    for sheet_name, df in sheets.items():
        if "QT Register" in sheet_name:
            if 'Date' in df.columns:
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            if 'Value' in df.columns:
                df['Value'] = pd.to_numeric(df['Value'], errors='coerce').fillna(0)
            if 'Salesperson' in df.columns:
                df['Salesperson'] = df['Salesperson'].astype(str).str.strip()
            qt_dfs.append(df)
    if qt_dfs:
        processed_sheets['quotations'] = pd.concat(qt_dfs, ignore_index=True)

    # Preprocess Invoices (assuming all sheets containing 'INV' are invoices)
    inv_dfs = []
    for sheet_name, df in sheets.items():
        # Added a broader check for 'INV' in sheet name for future-proofing,
        # but 2025 INV was specified as primary
        if "INV" in sheet_name:
            if 'Date' in df.columns:
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            if 'Amount' in df.columns:
                df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce').fillna(0)
            inv_dfs.append(df)
    if inv_dfs:
        processed_sheets['invoices'] = pd.concat(inv_dfs, ignore_index=True)


    # Preprocess Payment Pending
    if 'Payment Pending' in sheets:
        df = sheets['Payment Pending'].copy()
        if 'Pending Amount' in df.columns:
            df['Pending Amount'] = pd.to_numeric(df['Pending Amount'], errors='coerce').fillna(0)
        if 'Due Date' in df.columns:
            df['Due Date'] = pd.to_datetime(df['Due Date'], errors='coerce')
        processed_sheets['payment_pending'] = df

    # Preprocess Subscriptions
    if 'Subscriptions' in sheets:
        df = sheets['Subscriptions'].copy()
        if 'Start Date' in df.columns:
            df['Start Date'] = pd.to_datetime(df['Start Date'], errors='coerce')
        if 'End Date' in df.columns:
            df['End Date'] = pd.to_datetime(df['End Date'], errors='coerce')
        if 'Cost' in df.columns:
            df['Cost'] = pd.to_numeric(df['Cost'], errors='coerce').fillna(0)
        processed_sheets['subscriptions'] = df

    # Preprocess Meeting Agenda
    if 'Meeting Agenda' in sheets:
        df = sheets['Meeting Agenda'].copy()
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        processed_sheets['meeting_agenda'] = df
    
    # Preprocess To Remember
    if 'To Remember' in sheets:
        df = sheets['To Remember'].copy()
        if 'Due Date' in df.columns:
            df['Due Date'] = pd.to_datetime(df['Due Date'], errors='coerce')
        processed_sheets['to_remember'] = df

    return processed_sheets


# --- Main Content Area: Conditional Display ---
# Sidebar Navigation (only show if data is loaded)
if st.session_state.data_loaded and st.session_state.openai_api_key:
    st.sidebar.markdown("---")
    st.sidebar.title("Navigation")
    page = st.sidebar.radio("Go to", [
        "Dashboard Home",
        "Quotations",
        "Invoices",
        "Payments",
        "Subscriptions",
        "Tasks & Meetings",
        "Ask AI (GPT Insights)"
    ])

    processed_data = preprocess_data(st.session_state.sheets)

    # Dashboard Home
    if page == "Dashboard Home":
        st.header("Dashboard Home (KPI Overview)")

        if 'quotations' not in processed_data and 'invoices' not in processed_data and 'payment_pending' not in processed_data and 'subscriptions' not in processed_data:
            st.info("No relevant data found in the uploaded Excel file to display the Dashboard Home KPIs. Please ensure your Excel file contains 'QT Register', 'INV', 'Payment Pending', and 'Subscriptions' sheets with valid data.")
        else:
            # --- Filters ---
            st.subheader("Filters")
            available_years = []
            if 'quotations' in processed_data and not processed_data['quotations'].empty and 'Date' in processed_data['quotations'].columns:
                available_years.extend(processed_data['quotations']['Date'].dt.year.dropna().unique().astype(int).tolist())
            if 'invoices' in processed_data and not processed_data['invoices'].empty and 'Date' in processed_data['invoices'].columns:
                available_years.extend(processed_data['invoices']['Date'].dt.year.dropna().unique().astype(int).tolist())
            
            # Ensure unique and sorted years, add 'All' option
            available_years = sorted(list(set(available_years)))
            selected_year = st.selectbox("Select Year", ['All'] + available_years)

            available_salespersons = []
            if 'quotations' in processed_data and not processed_data['quotations'].empty and 'Salesperson' in processed_data['quotations'].columns:
                available_salespersons.extend(processed_data['quotations']['Salesperson'].dropna().unique().tolist())
            
            # Ensure unique and sorted salespersons, add 'All' option
            available_salespersons = sorted(list(set(available_salespersons)))
            selected_salesperson = st.selectbox("Select Salesperson", ['All'] + available_salespersons)

            # --- KPI Cards ---
            st.subheader("Key Performance Indicators")
            col1, col2, col3, col4 = st.columns(4)

            total_quotations_value = 0
            total_invoices_value = 0
            total_pending_payments = 0
            active_subscriptions_count = 0

            # Filter data based on selections
            filtered_quotations = pd.DataFrame()
            if 'quotations' in processed_data:
                filtered_quotations = processed_data['quotations'].copy()
                if selected_year != 'All' and 'Date' in filtered_quotations.columns:
                    filtered_quotations = filtered_quotations[filtered_quotations['Date'].dt.year == selected_year]
                if selected_salesperson != 'All' and 'Salesperson' in filtered_quotations.columns:
                    filtered_quotations = filtered_quotations[filtered_quotations['Salesperson'] == selected_salesperson]
                if 'Value' in filtered_quotations.columns:
                    total_quotations_value = filtered_quotations['Value'].sum()

            filtered_invoices = pd.DataFrame()
            if 'invoices' in processed_data:
                filtered_invoices = processed_data['invoices'].copy()
                if selected_year != 'All' and 'Date' in filtered_invoices.columns:
                    filtered_invoices = filtered_invoices[filtered_invoices['Date'].dt.year == selected_year]
                if 'Amount' in filtered_invoices.columns:
                    total_invoices_value = filtered_invoices['Amount'].sum()

            # Payment Pending (usually not filtered by year/salesperson for dashboard overview)
            if 'payment_pending' in processed_data and 'Pending Amount' in processed_data['payment_pending'].columns:
                total_pending_payments = processed_data['payment_pending']['Pending Amount'].sum()
            
            # Active Subscriptions (usually not filtered by year/salesperson for dashboard overview)
            if 'subscriptions' in processed_data and 'Status' in processed_data['subscriptions'].columns:
                active_subscriptions_count = processed_data['subscriptions'][
                    processed_data['subscriptions']['Status'].astype(str).str.strip().str.lower() == 'active'
                ].shape[0]

            with col1:
                st.metric("Total Quotations Value", f"BHD {total_quotations_value:,.2f}")
            with col2:
                st.metric("Total Invoices Value", f"BHD {total_invoices_value:,.2f}")
            with col3:
                st.metric("Total Pending Payments", f"BHD {total_pending_payments:,.2f}")
            with col4:
                st.metric("Active Subscriptions", active_subscriptions_count)

            # --- Monthly Trend Charts ---
            st.subheader("Monthly Trend Charts")

            # Quotation Trend
            if not filtered_quotations.empty and 'Date' in filtered_quotations.columns and 'Value' in filtered_quotations.columns:
                quotation_monthly_trend = filtered_quotations.set_index('Date').resample('M')['Value'].sum().reset_index()
                quotation_monthly_trend['Month'] = quotation_monthly_trend['Date'].dt.strftime('%Y-%m')
                fig_qt = px.line(
                    quotation_monthly_trend,
                    x='Month',
                    y='Value',
                    title='Monthly Quotation Trend',
                    labels={'Value': 'Quotation Value (BHD)', 'Month': 'Month'},
                    markers=True
                )
                fig_qt.update_xaxes(tickangle=45)
                fig_qt.update_layout(hovermode="x unified")
                st.plotly_chart(fig_qt, use_container_width=True)
            else:
                st.info("No quotation data available for the selected filters to display monthly trend.")

            # Invoice Trend
            if not filtered_invoices.empty and 'Date' in filtered_invoices.columns and 'Amount' in filtered_invoices.columns:
                invoice_monthly_trend = filtered_invoices.set_index('Date').resample('M')['Amount'].sum().reset_index()
                invoice_monthly_trend['Month'] = invoice_monthly_trend['Date'].dt.strftime('%Y-%m')
                fig_inv = px.line(
                    invoice_monthly_trend,
                    x='Month',
                    y='Amount',
                    title='Monthly Invoice Trend',
                    labels={'Amount': 'Invoice Amount (BHD)', 'Month': 'Month'},
                    markers=True,
                    color_discrete_sequence=['red']
                )
                fig_inv.update_xaxes(tickangle=45)
                fig_inv.update_layout(hovermode="x unified")
                st.plotly_chart(fig_inv, use_container_width=True)
            else:
                st.info("No invoice data available for the selected filters to display monthly trend.")

            # Quotation vs Invoice Trend (Combined)
            if not filtered_quotations.empty and not filtered_invoices.empty and 'Date' in filtered_quotations.columns and 'Value' in filtered_quotations.columns and 'Date' in filtered_invoices.columns and 'Amount' in filtered_invoices.columns:
                # Merge on month and year
                quotation_monthly_trend = filtered_quotations.set_index('Date').resample('M')['Value'].sum().reset_index()
                quotation_monthly_trend['Month'] = quotation_monthly_trend['Date'].dt.strftime('%Y-%m')

                invoice_monthly_trend = filtered_invoices.set_index('Date').resample('M')['Amount'].sum().reset_index()
                invoice_monthly_trend['Month'] = invoice_monthly_trend['Date'].dt.strftime('%Y-%m')

                combined_trend = pd.merge(
                    quotation_monthly_trend.rename(columns={'Value': 'Quotations'}),
                    invoice_monthly_trend.rename(columns={'Amount': 'Invoices'}),
                    on='Month',
                    how='outer'
                ).fillna(0) # Fill NaNs for months where one type of data is missing
                
                fig_combined = go.Figure()
                fig_combined.add_trace(go.Scatter(x=combined_trend['Month'], y=combined_trend['Quotations'], mode='lines+markers', name='Quotations'))
                fig_combined.add_trace(go.Scatter(x=combined_trend['Month'], y=combined_trend['Invoices'], mode='lines+markers', name='Invoices'))

                fig_combined.update_layout(
                    title='Monthly Quotation vs. Invoice Trend',
                    xaxis_title='Month',
                    yaxis_title='Value (BHD)',
                    hovermode="x unified"
                )
                fig_combined.update_xaxes(tickangle=45)
                st.plotly_chart(fig_combined, use_container_width=True)
            elif not filtered_quotations.empty or not filtered_invoices.empty:
                st.info("Both quotation and invoice data are needed to display a combined trend.")
            else:
                st.info("Not enough data to display combined quotation vs. invoice trend.")


    elif page == "Quotations":
        st.header("Quotations Analysis")
        st.info("Content for Quotations will go here in the next steps.")
        if 'quotations' in processed_data:
            st.subheader("Raw Quotation Data (Sample):")
            st.dataframe(processed_data['quotations'].head())

    elif page == "Invoices":
        st.header("Invoices Analysis")
        st.info("Content for Invoices will go here in the next steps.")
        if 'invoices' in processed_data:
            st.subheader("Raw Invoice Data (Sample):")
            st.dataframe(processed_data['invoices'].head())

    elif page == "Payments":
        st.header("Payments Overview")
        st.info("Content for Payments will go here in the next steps.")
        if 'payment_pending' in processed_data:
            st.subheader("Raw Payment Pending Data (Sample):")
            st.dataframe(processed_data['payment_pending'].head())

    elif page == "Subscriptions":
        st.header("Subscriptions Management")
        st.info("Content for Subscriptions will go here in the next steps.")
        if 'subscriptions' in processed_data:
            st.subheader("Raw Subscriptions Data (Sample):")
            st.dataframe(processed_data['subscriptions'].head())

    elif page == "Tasks & Meetings":
        st.header("Tasks & Meetings Tracker")
        st.info("Content for Tasks & Meetings will go here in the next steps.")
        if 'meeting_agenda' in processed_data:
            st.subheader("Raw Meeting Agenda Data (Sample):")
            st.dataframe(processed_data['meeting_agenda'].head())
        if 'to_remember' in processed_data:
            st.subheader("Raw To Remember Data (Sample):")
            st.dataframe(processed_data['to_remember'].head())

    elif page == "Ask AI (GPT Insights)":
        st.header("Ask AI (GPT Insights)")
        st.info("Content for AI Insights will go here in the next steps.")

else:
    # Display message if not all inputs are provided
    st.info("Please enter your OpenAI API Key and upload the 'SSS Master Sheet.xlsx' file using the sidebar to unlock the dashboard features.")


# --- Initial Run Instructions ---
st.sidebar.markdown("---")
st.sidebar.markdown("**How to Run:**")
st.sidebar.markdown("1. Make sure you have `SSS Master Sheet.xlsx` ready with the specified sheets.")
st.sidebar.markdown("2. Save this script as `app.py`.")
st.sidebar.markdown("3. Run `streamlit run app.py` in your terminal.")