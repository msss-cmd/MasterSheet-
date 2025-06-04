import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import openai
import os
from io import BytesIO

# Set page config
st.set_page_config(
    page_title="SSS Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load data function with caching
@st.cache_data
def load_data(uploaded_file):
    if uploaded_file is not None:
        try:
            excel_data = pd.ExcelFile(uploaded_file)
            sheets = {}
            
            # Load all required sheets
            required_sheets = [
                'QT Register 2023', 'QT Register 2024', 'QT Register 2025',
                '2025 INV', 'Payment Pending', 'Subscriptions',  # Note: Typo in original requirements ('Subscriptions' vs 'Subscriptions')
                'Meeting Agenda', 'To Remember'
            ]
            
            for sheet in required_sheets:
                try:
                    sheets[sheet] = pd.read_excel(excel_data, sheet_name=sheet)
                    # Basic cleaning - remove empty rows
                    sheets[sheet] = sheets[sheet].dropna(how='all')
                except Exception as e:
                    st.warning(f"Could not load sheet '{sheet}': {str(e)}")
                    sheets[sheet] = pd.DataFrame()
            
            # Combine quotation registers
            quotation_sheets = [s for s in sheets.keys() if 'QT Register' in s]
            combined_quotations = pd.concat([sheets[sheet] for sheet in quotation_sheets], ignore_index=True)
            sheets['Combined Quotations'] = combined_quotations
            
            return sheets
        except Exception as e:
            st.error(f"Error loading Excel file: {str(e)}")
            return None
    return None

# Initialize session state for OpenAI API key and chat history
if 'openai_api_key' not in st.session_state:
    st.session_state.openai_api_key = None
if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []

# Sidebar for navigation and settings
with st.sidebar:
    st.image("https://via.placeholder.com/150x50?text=SSS+Logo", width=150)
    st.title("SSS Dashboard")
    
    # File uploader
    uploaded_file = st.file_uploader("Upload SSS Master Sheet", type=["xlsx", "xls"])
    
    # Navigation
    page = st.radio(
        "Go to",
        ["Dashboard", "Quotations", "Invoices", "Payments", "Subscriptions", "Tasks & Meetings", "AI Assistant"],
        index=0
    )
    
    # OpenAI API key input
    with st.expander("OpenAI Settings"):
        api_key = st.text_input("Enter OpenAI API Key", type="password", value=st.session_state.openai_api_key or "")
        if api_key:
            st.session_state.openai_api_key = api_key
            openai.api_key = api_key
            st.success("API key saved!")
    
    st.markdown("---")
    st.markdown("**Salahuddin Softtech Solutions**")
    st.markdown("v1.0 | © 2025")

# Load data
data = load_data(uploaded_file) if uploaded_file else None

# Helper functions for data processing
def process_quotations_data(quotations_df):
    if quotations_df.empty:
        return quotations_df
    
    # Convert date columns if they exist
    if 'Date' in quotations_df.columns:
        quotations_df['Date'] = pd.to_datetime(quotations_df['Date'], errors='coerce')
        quotations_df['Year'] = quotations_df['Date'].dt.year
        quotations_df['Month'] = quotations_df['Date'].dt.month_name()
    
    return quotations_df

def process_invoices_data(invoices_df):
    if invoices_df.empty:
        return invoices_df
    
    # Convert date columns if they exist
    if 'Date' in invoices_df.columns:
        invoices_df['Date'] = pd.to_datetime(invoices_df['Date'], errors='coerce')
        invoices_df['Year'] = invoices_df['Date'].dt.year
        invoices_df['Month'] = invoices_df['Date'].dt.month_name()
    
    # Convert amount columns to numeric if they exist
    if 'Amount' in invoices_df.columns:
        invoices_df['Amount'] = pd.to_numeric(invoices_df['Amount'], errors='coerce')
    
    return invoices_df

def process_payments_data(payments_df):
    if payments_df.empty:
        return payments_df
    
    # Convert date columns if they exist
    date_columns = ['Due Date', 'Payment Date']
    for col in date_columns:
        if col in payments_df.columns:
            payments_df[col] = pd.to_datetime(payments_df[col], errors='coerce')
    
    # Convert amount columns to numeric if they exist
    if 'Pending Amount' in payments_df.columns:
        payments_df['Pending Amount'] = pd.to_numeric(payments_df['Pending Amount'], errors='coerce')
    
    # Calculate days overdue if possible
    if 'Due Date' in payments_df.columns:
        today = pd.to_datetime('today')
        payments_df['Days Overdue'] = (today - payments_df['Due Date']).dt.days
        payments_df['Aging Bucket'] = pd.cut(
            payments_df['Days Overdue'],
            bins=[-1, 0, 30, 60, 90, np.inf],
            labels=['Not Due', '0-30', '31-60', '61-90', '90+']
        )
    
    return payments_df

def process_subscriptions_data(subscriptions_df):
    if subscriptions_df.empty:
        return subscriptions_df
    
    # Convert date columns if they exist
    date_columns = ['Start Date', 'End Date']
    for col in date_columns:
        if col in subscriptions_df.columns:
            subscriptions_df[col] = pd.to_datetime(subscriptions_df[col], errors='coerce')
    
    # Convert amount columns to numeric if they exist
    if 'Cost' in subscriptions_df.columns:
        subscriptions_df['Cost'] = pd.to_numeric(subscriptions_df['Cost'], errors='coerce')
    
    # Calculate status
    if 'End Date' in subscriptions_df.columns:
        today = pd.to_datetime('today')
        subscriptions_df['Status'] = subscriptions_df['End Date'].apply(
            lambda x: 'Active' if x >= today else 'Expired'
        )
        subscriptions_df['Days Until Expiry'] = (subscriptions_df['End Date'] - today).dt.days
    
    return subscriptions_df

def process_meetings_data(meetings_df):
    if meetings_df.empty:
        return meetings_df
    
    # Convert date columns if they exist
    if 'Date' in meetings_df.columns:
        meetings_df['Date'] = pd.to_datetime(meetings_df['Date'], errors='coerce')
    
    return meetings_df

def process_tasks_data(tasks_df):
    if tasks_df.empty:
        return tasks_df
    
    # Convert date columns if they exist
    if 'Due Date' in tasks_df.columns:
        tasks_df['Due Date'] = pd.to_datetime(tasks_df['Due Date'], errors='coerce')
    
    # Calculate status if possible
    if 'Status' not in tasks_df.columns:
        tasks_df['Status'] = 'Pending'
    
    return tasks_df

# Dashboard Home Page
if page == "Dashboard":
    st.title("SSS Dashboard Overview")
    
    if data is None:
        st.warning("Please upload the SSS Master Sheet Excel file to view data.")
    else:
        # KPIs
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_quotations = len(data.get('Combined Quotations', pd.DataFrame()))
            st.metric("Total Quotations", f"{total_quotations:,}")
        
        with col2:
            total_invoices = len(data.get('2025 INV', pd.DataFrame()))
            st.metric("Total Invoices (2025)", f"{total_invoices:,}")
        
        with col3:
            payments_df = process_payments_data(data.get('Payment Pending', pd.DataFrame()))
            total_pending = payments_df['Pending Amount'].sum() if not payments_df.empty and 'Pending Amount' in payments_df.columns else 0
            st.metric("Total Payments Pending", f"${total_pending:,.2f}")
        
        with col4:
            subs_df = process_subscriptions_data(data.get('Subscriptions', pd.DataFrame()))
            total_subs = len(subs_df) if not subs_df.empty else 0
            st.metric("Total Subscriptions", f"{total_subs:,}")
        
        st.markdown("---")
        
        # Monthly Quotation Trends
        st.subheader("Monthly Quotation Trends")
        quotations_df = process_quotations_data(data.get('Combined Quotations', pd.DataFrame()))
        
        if not quotations_df.empty and 'Date' in quotations_df.columns:
            # Group by month and year
            monthly_quotations = quotations_df.groupby(['Year', 'Month']).size().reset_index(name='Count')
            monthly_quotations['Month_Year'] = monthly_quotations['Month'] + ' ' + monthly_quotations['Year'].astype(str)
            
            # Create plot
            fig = px.line(
                monthly_quotations,
                x='Month_Year',
                y='Count',
                title='Monthly Quotation Count',
                markers=True
            )
            fig.update_layout(xaxis_title='Month', yaxis_title='Quotation Count')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("No date information available in quotations data.")
        
        # Monthly Invoice Trends
        st.subheader("Monthly Invoice Trends (2025)")
        invoices_df = process_invoices_data(data.get('2025 INV', pd.DataFrame()))
        
        if not invoices_df.empty and 'Date' in invoices_df.columns and 'Amount' in invoices_df.columns:
            # Group by month and sum amounts
            monthly_invoices = invoices_df.groupby(['Month']).agg({'Amount': 'sum'}).reset_index()
            
            # Create plot
            fig = px.bar(
                monthly_invoices,
                x='Month',
                y='Amount',
                title='Monthly Invoice Amounts (2025)',
                text='Amount'
            )
            fig.update_traces(texttemplate='%{text:$,.0f}', textposition='outside')
            fig.update_layout(xaxis_title='Month', yaxis_title='Total Amount ($)')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("No date or amount information available in 2025 invoices data.")

# Quotations Page
elif page == "Quotations":
    st.title("Quotations Analysis")
    
    if data is None:
        st.warning("Please upload the SSS Master Sheet Excel file to view data.")
    else:
        quotations_df = process_quotations_data(data.get('Combined Quotations', pd.DataFrame()))
        
        if quotations_df.empty:
            st.warning("No quotation data available.")
        else:
            # Filters
            with st.expander("Filters"):
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    years = sorted(quotations_df['Year'].unique()) if 'Year' in quotations_df.columns else []
                    selected_years = st.multiselect("Select Years", years, default=years)
                    if selected_years:
                        quotations_df = quotations_df[quotations_df['Year'].isin(selected_years)]
                
                with col2:
                    clients = sorted(quotations_df['Client'].unique()) if 'Client' in quotations_df.columns else []
                    selected_clients = st.multiselect("Select Clients", clients, default=clients)
                    if selected_clients:
                        quotations_df = quotations_df[quotations_df['Client'].isin(selected_clients)]
                
                with col3:
                    salespeople = sorted(quotations_df['Salesperson'].unique()) if 'Salesperson' in quotations_df.columns else []
                    selected_salespeople = st.multiselect("Select Salespeople", salespeople, default=salespeople)
                    if selected_salespeople:
                        quotations_df = quotations_df[quotations_df['Salesperson'].isin(selected_salespeople)]
            
            # Top Clients by Quotation Count
            st.subheader("Top Clients by Quotation Count")
            if 'Client' in quotations_df.columns:
                top_clients = quotations_df['Client'].value_counts().reset_index()
                top_clients.columns = ['Client', 'Quotation Count']
                
                fig = px.bar(
                    top_clients.head(10),
                    x='Client',
                    y='Quotation Count',
                    title='Top 10 Clients by Quotation Count'
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Client information not available in quotations data.")
            
            # Conversion Rate Analysis
            st.subheader("Quotation to Invoice Conversion")
            
            # Check if we have invoice data to compare
            invoices_df = process_invoices_data(data.get('2025 INV', pd.DataFrame()))
            if not invoices_df.empty and 'Quotation ID' in invoices_df.columns and 'Quotation ID' in quotations_df.columns:
                # Get unique quotation IDs that became invoices
                converted_quotations = invoices_df['Quotation ID'].unique()
                
                # Calculate conversion rate
                total_quotations = len(quotations_df)
                converted_count = len(quotations_df[quotations_df['Quotation ID'].isin(converted_quotations)])
                conversion_rate = (converted_count / total_quotations) * 100 if total_quotations > 0 else 0
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Total Quotations", total_quotations)
                with col2:
                    st.metric("Conversion Rate", f"{conversion_rate:.1f}%")
                
                # Conversion by salesperson if available
                if 'Salesperson' in quotations_df.columns:
                    conversion_by_salesperson = quotations_df.groupby('Salesperson').apply(
                        lambda x: pd.Series({
                            'Total Quotations': len(x),
                            'Converted Quotations': len(x[x['Quotation ID'].isin(converted_quotations)]),
                            'Conversion Rate': (len(x[x['Quotation ID'].isin(converted_quotations)]) / len(x)) * 100 if len(x) > 0 else 0
                        })
                    ).reset_index()
                    
                    fig = px.bar(
                        conversion_by_salesperson,
                        x='Salesperson',
                        y='Conversion Rate',
                        title='Conversion Rate by Salesperson',
                        text='Conversion Rate'
                    )
                    fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Unable to calculate conversion rate without invoice data or quotation IDs.")

# Invoices Page
elif page == "Invoices":
    st.title("Invoices Analysis")
    
    if data is None:
        st.warning("Please upload the SSS Master Sheet Excel file to view data.")
    else:
        invoices_df = process_invoices_data(data.get('2025 INV', pd.DataFrame()))
        
        if invoices_df.empty:
            st.warning("No invoice data available.")
        else:
            # Filters
            with st.expander("Filters"):
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    clients = sorted(invoices_df['Client'].unique()) if 'Client' in invoices_df.columns else []
                    selected_clients = st.multiselect("Select Clients", clients, default=clients)
                    if selected_clients:
                        invoices_df = invoices_df[invoices_df['Client'].isin(selected_clients)]
                
                with col2:
                    salespeople = sorted(invoices_df['Salesperson'].unique()) if 'Salesperson' in invoices_df.columns else []
                    selected_salespeople = st.multiselect("Select Salespeople", salespeople, default=salespeople)
                    if selected_salespeople:
                        invoices_df = invoices_df[invoices_df['Salesperson'].isin(selected_salespeople)]
                
                with col3:
                    po_numbers = sorted(invoices_df['PO Number'].unique()) if 'PO Number' in invoices_df.columns else []
                    selected_po = st.multiselect("Select PO Numbers", po_numbers, default=po_numbers)
                    if selected_po:
                        invoices_df = invoices_df[invoices_df['PO Number'].isin(selected_po)]
            
            # Total Invoice Value by Month
            st.subheader("Total Invoice Value by Month")
            if 'Month' in invoices_df.columns and 'Amount' in invoices_df.columns:
                monthly_totals = invoices_df.groupby('Month')['Amount'].sum().reset_index()
                
                fig = px.bar(
                    monthly_totals,
                    x='Month',
                    y='Amount',
                    title='Monthly Invoice Totals',
                    text='Amount'
                )
                fig.update_traces(texttemplate='%{text:$,.0f}', textposition='outside')
                fig.update_layout(yaxis_title='Total Amount ($)')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Month or Amount information not available in invoice data.")
            
            # Average Invoice Size
            st.subheader("Average Invoice Size")
            if 'Amount' in invoices_df.columns:
                avg_invoice_size = invoices_df['Amount'].mean()
                median_invoice_size = invoices_df['Amount'].median()
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Average Invoice Size", f"${avg_invoice_size:,.2f}")
                with col2:
                    st.metric("Median Invoice Size", f"${median_invoice_size:,.2f}")
                
                # Distribution of invoice sizes
                fig = px.histogram(
                    invoices_df,
                    x='Amount',
                    nbins=20,
                    title='Distribution of Invoice Amounts'
                )
                fig.update_layout(xaxis_title='Amount ($)', yaxis_title='Count')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Amount information not available in invoice data.")

# Payments Page
elif page == "Payments":
    st.title("Payments Analysis")
    
    if data is None:
        st.warning("Please upload the SSS Master Sheet Excel file to view data.")
    else:
        payments_df = process_payments_data(data.get('Payment Pending', pd.DataFrame()))
        
        if payments_df.empty:
            st.warning("No payment data available.")
        else:
            # Total and client-wise pending amounts
            st.subheader("Pending Payments Summary")
            
            if 'Pending Amount' in payments_df.columns:
                total_pending = payments_df['Pending Amount'].sum()
                st.metric("Total Pending Amount", f"${total_pending:,.2f}")
                
                # Client-wise summary
                if 'Client' in payments_df.columns:
                    client_summary = payments_df.groupby('Client')['Pending Amount'].sum().reset_index().sort_values('Pending Amount', ascending=False)
                    
                    fig = px.bar(
                        client_summary,
                        x='Client',
                        y='Pending Amount',
                        title='Pending Amount by Client',
                        text='Pending Amount'
                    )
                    fig.update_traces(texttemplate='%{text:$,.0f}', textposition='outside')
                    fig.update_layout(yaxis_title='Pending Amount ($)')
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("Client information not available in payment data.")
            else:
                st.warning("Pending Amount information not available in payment data.")
            
            # Overdue payments and aging buckets
            st.subheader("Overdue Payments Analysis")
            
            if 'Days Overdue' in payments_df.columns and 'Aging Bucket' in payments_df.columns:
                # Aging bucket summary
                aging_summary = payments_df['Aging Bucket'].value_counts().reset_index()
                aging_summary.columns = ['Aging Bucket', 'Count']
                
                fig = px.pie(
                    aging_summary,
                    names='Aging Bucket',
                    values='Count',
                    title='Payment Aging Distribution'
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # Overdue payments table
                overdue_payments = payments_df[payments_df['Days Overdue'] > 0]
                if not overdue_payments.empty:
                    st.write("Overdue Payments Details:")
                    st.dataframe(
                        overdue_payments.sort_values('Days Overdue', ascending=False),
                        column_config={
                            "Client": "Client",
                            "Pending Amount": st.column_config.NumberColumn("Amount", format="$%.2f"),
                            "Due Date": "Due Date",
                            "Days Overdue": "Days Overdue",
                            "Aging Bucket": "Aging Bucket",
                            "Remarks": "Remarks"
                        },
                        hide_index=True,
                        use_container_width=True
                    )
                else:
                    st.success("No overdue payments found!")
            else:
                st.warning("Unable to analyze overdue payments without date information.")

# Subscriptions Page
elif page == "Subscriptions":
    st.title("Subscriptions Management")
    
    if data is None:
        st.warning("Please upload the SSS Master Sheet Excel file to view data.")
    else:
        subs_df = process_subscriptions_data(data.get('Subscriptions', pd.DataFrame()))
        
        if subs_df.empty:
            st.warning("No subscription data available.")
        else:
            # Active vs Expired
            st.subheader("Subscription Status")
            
            if 'Status' in subs_df.columns:
                status_counts = subs_df['Status'].value_counts().reset_index()
                status_counts.columns = ['Status', 'Count']
                
                col1, col2 = st.columns(2)
                with col1:
                    fig = px.pie(
                        status_counts,
                        names='Status',
                        values='Count',
                        title='Active vs Expired Subscriptions'
                    )
                    st.plotly_chart(fig, use_container_width=True)
                
                with col2:
                    active_subs = subs_df[subs_df['Status'] == 'Active']
                    if not active_subs.empty and 'Cost' in active_subs.columns:
                        total_monthly_cost = active_subs['Cost'].sum()
                        st.metric("Total Monthly Cost (Active)", f"${total_monthly_cost:,.2f}")
                    else:
                        st.warning("Cost information not available for active subscriptions.")
            else:
                st.warning("Status information not available in subscription data.")
            
            # Upcoming Expirations
            st.subheader("Upcoming Subscription Expirations")
            
            if 'End Date' in subs_df.columns and 'Days Until Expiry' in subs_df.columns:
                upcoming_subs = subs_df[
                    (subs_df['Status'] == 'Active') & 
                    (subs_df['Days Until Expiry'] <= 60)
                ].sort_values('Days Until Expiry')
                
                if not upcoming_subs.empty:
                    st.write("Subscriptions expiring in the next 60 days:")
                    st.dataframe(
                        upcoming_subs[['Client', 'Service', 'End Date', 'Days Until Expiry', 'Cost']],
                        column_config={
                            "Client": "Client",
                            "Service": "Service",
                            "End Date": "End Date",
                            "Days Until Expiry": "Days Until Expiry",
                            "Cost": st.column_config.NumberColumn("Cost", format="$%.2f")
                        },
                        hide_index=True,
                        use_container_width=True
                    )
                else:
                    st.success("No subscriptions expiring in the next 60 days!")
            else:
                st.warning("Unable to identify upcoming expirations without end date information.")
            
            # Subscription Explorer
            st.subheader("Subscription Explorer")
            
            # Filters
            with st.expander("Filters"):
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    status_options = sorted(subs_df['Status'].unique()) if 'Status' in subs_df.columns else []
                    selected_status = st.multiselect("Select Status", status_options, default=status_options)
                    if selected_status:
                        subs_df = subs_df[subs_df['Status'].isin(selected_status)]
                
                with col2:
                    if 'Cost' in subs_df.columns:
                        min_cost, max_cost = st.slider(
                            "Cost Range",
                            min_value=float(subs_df['Cost'].min()),
                            max_value=float(subs_df['Cost'].max()),
                            value=(float(subs_df['Cost'].min()), float(subs_df['Cost'].max()))
                        )
                        subs_df = subs_df[(subs_df['Cost'] >= min_cost) & (subs_df['Cost'] <= max_cost)]
                
                with col3:
                    clients = sorted(subs_df['Client'].unique()) if 'Client' in subs_df.columns else []
                    selected_clients = st.multiselect("Select Clients", clients, default=clients)
                    if selected_clients:
                        subs_df = subs_df[subs_df['Client'].isin(selected_clients)]
            
            # Display filtered subscriptions
            st.dataframe(
                subs_df,
                column_config={
                    "Client": "Client",
                    "Service": "Service",
                    "Start Date": "Start Date",
                    "End Date": "End Date",
                    "Status": "Status",
                    "Cost": st.column_config.NumberColumn("Cost", format="$%.2f"),
                    "Days Until Expiry": "Days Until Expiry"
                },
                hide_index=True,
                use_container_width=True
            )

# Tasks & Meetings Page
elif page == "Tasks & Meetings":
    st.title("Tasks & Meetings")
    
    if data is None:
        st.warning("Please upload the SSS Master Sheet Excel file to view data.")
    else:
        # Tasks
        st.subheader("Pending Tasks")
        tasks_df = process_tasks_data(data.get('To Remember', pd.DataFrame()))
        
        if tasks_df.empty:
            st.warning("No task data available.")
        else:
            # Filter for pending tasks
            pending_tasks = tasks_df[tasks_df['Status'] == 'Pending']
            
            if not pending_tasks.empty:
                # Group by responsible person if available
                if 'Responsible' in pending_tasks.columns:
                    task_counts = pending_tasks['Responsible'].value_counts().reset_index()
                    task_counts.columns = ['Responsible', 'Task Count']
                    
                    fig = px.bar(
                        task_counts,
                        x='Responsible',
                        y='Task Count',
                        title='Pending Tasks by Responsible Person'
                    )
                    st.plotly_chart(fig, use_container_width=True)
                
                # Display task details
                st.write("Pending Task Details:")
                st.dataframe(
                    pending_tasks,
                    column_config={
                        "Task": "Task",
                        "Responsible": "Responsible",
                        "Due Date": "Due Date",
                        "Status": "Status",
                        "Remarks": "Remarks"
                    },
                    hide_index=True,
                    use_container_width=True
                )
            else:
                st.success("No pending tasks found!")
        
        # Meetings
        st.subheader("Upcoming Meetings")
        meetings_df = process_meetings_data(data.get('Meeting Agenda', pd.DataFrame()))
        
        if meetings_df.empty:
            st.warning("No meeting data available.")
        else:
            # Filter for upcoming meetings
            if 'Date' in meetings_df.columns:
                today = pd.to_datetime('today')
                upcoming_meetings = meetings_df[meetings_df['Date'] >= today].sort_values('Date')
                
                if not upcoming_meetings.empty:
                    # Display in a timeline
                    upcoming_meetings['Date_str'] = upcoming_meetings['Date'].dt.strftime('%Y-%m-%d')
                    
                    fig = px.timeline(
                        upcoming_meetings,
                        x_start="Date_str",
                        x_end="Date_str",
                        y="Title",
                        title="Upcoming Meetings Timeline"
                    )
                    fig.update_yaxes(autorange="reversed")
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Display meeting details
                    st.write("Meeting Details:")
                    st.dataframe(
                        upcoming_meetings,
                        column_config={
                            "Date": "Date",
                            "Title": "Title",
                            "Participants": "Participants",
                            "Agenda": "Agenda"
                        },
                        hide_index=True,
                        use_container_width=True
                    )
                else:
                    st.success("No upcoming meetings scheduled!")
            else:
                st.warning("Unable to identify upcoming meetings without date information.")

# AI Assistant Page
elif page == "AI Assistant":
    st.title("AI Business Insights Assistant")
    
    if data is None:
        st.warning("Please upload the SSS Master Sheet Excel file to enable AI insights.")
    elif not st.session_state.openai_api_key:
        st.warning("Please enter your OpenAI API key in the sidebar to enable AI insights.")
    else:
        # Display chat history
        for message in st.session_state.chat_history:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
        
        # User input
        if prompt := st.chat_input("Ask a question about SSS business data..."):
            # Add user message to chat history
            st.session_state.chat_history.append({"role": "user", "content": prompt})
            
            # Display user message
            with st.chat_message("user"):
                st.markdown(prompt)
            
            # Prepare context from data
            context = "SSS Business Data Summary:\n"
            
            # Add quotations summary
            quotations_df = process_quotations_data(data.get('Combined Quotations', pd.DataFrame()))
            if not quotations_df.empty:
                context += f"- Quotations: {len(quotations_df)} total quotations\n"
                if 'Year' in quotations_df.columns:
                    context += f"  - Years: {', '.join(map(str, sorted(quotations_df['Year'].unique())))}\n"
                if 'Client' in quotations_df.columns:
                    top_clients = quotations_df['Client'].value_counts().head(3).index.tolist()
                    context += f"  - Top clients: {', '.join(top_clients)}\n"
            
            # Add invoices summary
            invoices_df = process_invoices_data(data.get('2025 INV', pd.DataFrame()))
            if not invoices_df.empty:
                context += f"- Invoices: {len(invoices_df)} total invoices\n"
                if 'Amount' in invoices_df.columns:
                    total_invoice_amount = invoices_df['Amount'].sum()
                    avg_invoice = invoices_df['Amount'].mean()
                    context += f"  - Total amount: ${total_invoice_amount:,.2f}\n"
                    context += f"  - Average invoice: ${avg_invoice:,.2f}\n"
            
            # Add payments summary
            payments_df = process_payments_data(data.get('Payment Pending', pd.DataFrame()))
            if not payments_df.empty and 'Pending Amount' in payments_df.columns:
                total_pending = payments_df['Pending Amount'].sum()
                context += f"- Pending Payments: {len(payments_df)} pending payments totaling ${total_pending:,.2f}\n"
                if 'Days Overdue' in payments_df.columns:
                    overdue = payments_df[payments_df['Days Overdue'] > 0]
                    context += f"  - Overdue payments: {len(overdue)} totaling ${overdue['Pending Amount'].sum():,.2f}\n"
            
            # Add subscriptions summary
            subs_df = process_subscriptions_data(data.get('Subscriptions', pd.DataFrame()))
            if not subs_df.empty:
                context += f"- Subscriptions: {len(subs_df)} total subscriptions\n"
                if 'Status' in subs_df.columns:
                    active = subs_df[subs_df['Status'] == 'Active']
                    context += f"  - Active: {len(active)}\n"
                    if 'Cost' in active.columns:
                        context += f"  - Monthly cost: ${active['Cost'].sum():,.2f}\n"
                if 'End Date' in subs_df.columns:
                    upcoming = subs_df[(subs_df['Status'] == 'Active') & (subs_df['Days Until Expiry'] <= 30)]
                    context += f"  - Expiring in 30 days: {len(upcoming)}\n"
            
            # Add tasks/meetings summary
            tasks_df = process_tasks_data(data.get('To Remember', pd.DataFrame()))
            if not tasks_df.empty:
                pending_tasks = tasks_df[tasks_df['Status'] == 'Pending']
                context += f"- Tasks: {len(pending_tasks)} pending tasks\n"
            
            meetings_df = process_meetings_data(data.get('Meeting Agenda', pd.DataFrame()))
            if not meetings_df.empty and 'Date' in meetings_df.columns:
                upcoming_meetings = meetings_df[meetings_df['Date'] >= pd.to_datetime('today')]
                context += f"- Meetings: {len(upcoming_meetings)} upcoming meetings\n"
            
            # Prepare messages for OpenAI
            messages = [
                {"role": "system", "content": "You are a helpful business intelligence assistant for Salahuddin Softtech Solutions (SSS). Analyze the provided business data and answer questions concisely with relevant numbers when possible."},
                {"role": "user", "content": f"{context}\n\nQuestion: {prompt}"}
            ]
            
            # Call OpenAI API
            try:
                with st.spinner("Generating response..."):
                    response = openai.ChatCompletion.create(
                        model="gpt-4",
                        messages=messages,
                        temperature=0.3
                    )
                    ai_response = response.choices[0].message['content']
                
                # Add AI response to chat history
                st.session_state.chat_history.append({"role": "assistant", "content": ai_response})
                
                # Display AI response
                with st.chat_message("assistant"):
                    st.markdown(ai_response)
            
            except Exception as e:
                st.error(f"Error calling OpenAI API: {str(e)}")
