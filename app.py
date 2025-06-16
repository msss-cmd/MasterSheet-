import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io

# --- FIX: Move st.set_page_config to the very top ---
st.set_page_config(layout="wide", page_title="Salahuddin Softech Solutions Dashboard")

# --- Custom CSS for Streamlit App (from provided HTML/CSS concepts) ---
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?display=swap&family=Noto+Sans:wght@400;500;700;900&family=Space+Grotesk:wght@400;500;700');

    :root {
        --primary-bg: #141f1f;
        --secondary-bg: #294242;
        --border-color: #3b5e5e;
        --text-color: #ffffff;
        --secondary-text-color: #9bc0c0;
        --accent-button-bg: #d2f3f3;
        --accent-button-text: #141f1f;
        --positive-change: #0bda50;
        --negative-change: #fa5c38;
    }

    /* Overall page background and font */
    html, body, [data-testid="stAppViewContainer"] {
        background-color: var(--primary-bg);
        font-family: 'Space Grotesk', 'Noto Sans', sans-serif;
        color: var(--text-color);
    }

    /* Main container padding adjustments */
    .stApp {
        max-width: 960px; /* Equivalent to your max-w-[960px] */
        padding-left: 40px; /* px-40 is too much for standard screens, adjusted */
        padding-right: 40px; /* px-40 is too much for standard screens, adjusted */
        margin: auto; /* Center the content */
    }

    /* Header styling */
    h1, h2, h3, h4, h5, h6 {
        color: var(--text-color);
        font-weight: 700;
        line-height: tight;
        letter-spacing: -0.015em;
    }
    h1 { font-size: 28px; } /* Adjusted from px-40 */
    h2 { font-size: 22px; } /* Adjusted */

    /* Paragraph text */
    p, .stMarkdown, .stText {
        color: var(--text-color);
        font-weight: 400;
        line-height: normal;
    }

    /* Customizing file uploader */
    [data-testid="stFileUploaderDropzone"] {
        background-color: var(--primary-bg); /* Match app background */
        border: 2px dashed var(--border-color);
        border-radius: 0.75rem; /* rounded-xl */
        padding: 3.5rem 1.5rem; /* py-14 px-6 */
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 1.5rem; /* gap-6 */
    }
    [data-testid="stFileUploaderDropzone"] p {
        color: var(--text-color);
        font-weight: 700; /* font-bold */
        font-size: 1.125rem; /* text-lg */
        line-height: 1.75rem; /* leading-tight */
    }
    [data-testid="stFileUploaderDropzone"] button {
        background-color: var(--secondary-bg) !important;
        color: var(--text-color) !important;
        border: none !important;
        border-radius: 9999px !important; /* rounded-full */
        height: 2.5rem !important; /* h-10 */
        padding-left: 1rem !important; /* px-4 */
        padding-right: 1rem !important;
        font-size: 0.875rem !important; /* text-sm */
        font-weight: 700 !important; /* font-bold */
        line-height: 1.25rem !important; /* leading-normal */
        letter-spacing: 0.015em !important; /* tracking-[0.015em] */
        min-width: 84px; /* min-w-[84px] */
    }
     [data-testid="stFileUploaderUploadButton"] div {
        background-color: var(--secondary-bg) !important;
        color: var(--text-color) !important;
        border-radius: 9999px !important; /* rounded-full */
        font-size: 0.875rem !important; /* text-sm */
        font-weight: 700 !important; /* font-bold */
        height: 2.5rem !important; /* h-10 */
        line-height: 1.25rem !important; /* leading-normal */
        letter-spacing: 0.015em !important; /* tracking-[0.015em] */
     }


    /* Sidebar styling - Removed since sidebar is minimal */
    /* [data-testid="stSidebar"] {
        background-color: var(--primary-bg);
    } */

    /* Metric cards (st.container used for this) */
    .metric-card {
        background-color: var(--secondary-bg);
        border-radius: 0.75rem; /* rounded-xl */
        padding: 1.5rem; /* p-6 */
        display: flex;
        flex-direction: column;
        gap: 0.5rem; /* gap-2 */
        min-width: 158px; /* min-w-[158px] */
    }
    .metric-card p:first-child { /* Metric Title */
        color: var(--text-color);
        font-size: 1rem; /* text-base */
        font-weight: 500; /* font-medium */
        line-height: normal;
    }
    .metric-card .metric-value { /* Metric Value */
        color: var(--text-color);
        font-size: 1.5rem; /* text-2xl */
        font-weight: 700; /* font-bold */
        line-height: 1.75rem; /* leading-tight */
        letter-spacing: light; /* tracking-light */
    }

    /* Table styling */
    .stDataFrame {
        border-radius: 0.75rem; /* rounded-xl */
        border: 1px solid var(--border-color);
        overflow: hidden;
    }
    .stDataFrame table {
        background-color: var(--primary-bg);
    }
    .stDataFrame th {
        background-color: #1d2f2f; /* bg-[#1d2f2f] */
        color: var(--text-color);
        font-size: 0.875rem; /* text-sm */
        font-weight: 500; /* font-medium */
        line-height: normal;
        padding: 0.75rem 1rem; /* px-4 py-3 */
        text-align: left;
        border-bottom: none;
    }
    .stDataFrame td {
        background-color: var(--primary-bg);
        color: var(--secondary-text-color);
        font-size: 0.875rem; /* text-sm */
        font-weight: 400; /* font-normal */
        line-height: normal;
        padding: 0.5rem 1rem; /* px-4 py-2 */
        border-top: 1px solid var(--border-color); /* border-t-[#3b5e5e] */
    }
    .stDataFrame tbody tr:first-child td {
        border-top: none; /* No top border for the very first row */
    }

    /* Multiselect / Selectbox styling */
    [data-testid="stMultiSelect"] > div > div {
        background-color: var(--secondary-bg) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 0.75rem !important; /* rounded-xl */
    }
    [data-testid="stMultiSelect"] label {
        color: var(--text-color);
        font-size: 1rem; /* text-base */
        font-weight: 500; /* font-medium */
        line-height: normal;
    }
    [data-testid="stMultiSelect"] [data-testid="stMultiSelectOptions"] {
        background-color: var(--secondary-bg) !important;
        border: 1px solid var(--border-color) !important;
    }
     [data-testid="stMultiSelect"] [data-testid="stMultiSelectOptions"] div {
        color: var(--secondary-text-color) !important;
    }
    [data-testid="stMultiSelect"] [data-testid="stMultiSelectOptions"] div:hover {
        background-color: var(--border-color) !important;
    }
     [data-testid="stMultiSelect"] [data-testid="stMultiSelectChip"] {
        background-color: var(--border-color) !important;
        color: var(--text-color) !important;
        border-radius: 9999px !important;
    }


    /* Apply button style */
    .stButton > button {
        background-color: var(--accent-button-bg) !important;
        color: var(--accent-button-text) !important;
        border: none !important;
        border-radius: 9999px !important; /* rounded-full */
        height: 2.5rem !important; /* h-10 */
        padding-left: 1rem !important; /* px-4 */
        padding-right: 1rem !important;
        font-size: 0.875rem !important; /* text-sm */
        font-weight: 700 !important; /* font-bold */
        line-height: 1.25rem !important; /* leading-normal */
        letter-spacing: 0.015em !important; /* tracking-[0.015em] */
        min-width: 84px; /* min-w-[84px] */
        margin-top: 1rem; /* Adjust spacing */
    }

    /* Expander styling */
    [data-testid="stExpander"] {
        background-color: var(--primary-bg); /* Match overall background */
        border: 1px solid var(--border-color); /* Add border */
        border-radius: 0.75rem; /* rounded-xl */
        margin-bottom: 1rem; /* Add some spacing below expanders */
    }
    [data-testid="stExpanderSummary"] {
        background-color: var(--secondary-bg); /* Header background */
        border-radius: 0.75rem 0.75rem 0 0; /* Rounded top corners */
        padding: 0.75rem 1rem; /* Add padding */
    }
    [data-testid="stExpanderSummary"] p {
        color: var(--text-color) !important;
        font-weight: 600;
    }
    [data-testid="stExpanderContent"] {
        padding: 1rem; /* Adjust content padding */
    }

    /* Plot styling for Matplotlib/Seaborn to fit dark theme */
    .stPlotlyChart, .stImage {
        background-color: var(--secondary-bg); /* Chart background */
        border-radius: 0.75rem;
        padding: 1rem; /* Add padding around the plot */
        margin-bottom: 1rem; /* Spacing between plots */
    }

    /* Info/Warning/Error boxes */
    [data-testid="stAlert"] {
        background-color: var(--secondary-bg);
        color: var(--text-color);
        border: 1px solid var(--border-color);
    }
    [data-testid="stAlert"] [data-testid="stMarkdownContainer"] p {
        color: var(--text-color);
    }

    /* Horizontal line separator */
    hr {
        border-top: 1px solid var(--border-color);
    }

    /* Specific element styling adjustments */
    .stMarkdownContainer h2 {
        padding-left: 1rem; /* px-4 */
        padding-right: 1rem; /* px-4 */
        padding-top: 1.25rem; /* pt-5 */
        padding-bottom: 0.75rem; /* pb-3 */
        margin-top: 0;
    }
    .stMarkdownContainer p {
        padding-left: 1rem; /* px-4 */
        padding-right: 1rem; /* px-4 */
    }

    /* Adjust main content padding to match design */
    [data-testid="stVerticalBlock"] > div:first-child {
        padding-top: 1.25rem; /* py-5 */
        padding-bottom: 1.25rem; /* py-5 */
    }
    [data-testid="stHorizontalBlock"] {
        gap: 1rem; /* gap-4 */
    }
    [data-testid="stVerticalBlock"] {
        gap: 1rem; /* gap-4 */
    }

    /* Adjust padding for individual content sections based on HTML structure */
    .main-content-section {
        padding: 1rem; /* p-4 */
    }

    .key-metrics-container {
        display: flex;
        flex-wrap: wrap;
        gap: 1rem; /* gap-4 */
        padding-left: 1rem; /* px-4 */
        padding-right: 1rem; /* px-4 */
        padding-bottom: 1rem; /* p-4 for the section */
        padding-top: 1rem;
    }
    .table-container-div {
        padding-left: 1rem; /* px-4 */
        padding-right: 1rem; /* px-4 */
        padding-bottom: 0.75rem; /* py-3 */
        padding-top: 0.75rem;
    }
    .chart-panel-container {
        display: flex;
        flex-wrap: wrap;
        gap: 1rem; /* gap-4 */
        padding-left: 1rem; /* px-4 */
        padding-right: 1rem; /* px-4 */
        padding-top: 1.5rem; /* py-6 */
        padding-bottom: 1.5rem;
    }

    </style>
    """,
    unsafe_allow_html=True
)

# --- Helper Functions for Processing Each File Type (moved to top) ---

def process_qt_register_2025(df: pd.DataFrame):
    """
    Processes the 'QT Register 2025' data, performs analysis, and generates visualizations.
    Includes advanced visualizations.
    """
    st.markdown("<h2 class='st-emotion-cache-10grg6x e10grg6x4'>Quotation Register 2025</h2>", unsafe_allow_html=True)

    # Data Cleaning for QT Register 2025
    df['Date'] = df['Date'].astype(str).str.replace('//', '/', regex=False)
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df.dropna(subset=['Date'], inplace=True)

    # Strip leading/trailing spaces from categorical columns
    for col in ['Company  Name', 'Product', 'Sales Person']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # --- Key Metrics ---
    st.markdown("<div class='key-metrics-container'>", unsafe_allow_html=True)
    num_quotations = df['Quotation ID'].nunique() if 'Quotation ID' in df.columns else len(df)
    top_salesperson = df['Sales Person'].value_counts().index[0] if 'Sales Person' in df.columns and not df['Sales Person'].empty else "N/A"

    col1, col2, col3 = st.columns([1,1,1]) # Use 3 columns for metrics if needed

    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <p>Number of Quotations</p>
            <p class="metric-value">{num_quotations}</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <p>Top Salesperson</p>
            <p class="metric-value">{top_salesperson}</p>
        </div>
        """, unsafe_allow_html=True)
    # Adding a placeholder third metric or expanding if data allows
    if 'Value' in df.columns:
        total_value = df['Value'].sum() if pd.api.types.is_numeric_dtype(df['Value']) else 0
        with col3:
            st.markdown(f"""
            <div class="metric-card">
                <p>Total Quotation Value</p>
                <p class="metric-value">${total_value:,.0f}</p>
            </div>
            """, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Raw Data Preview (styled as a table container)
    st.subheader("Raw Data Preview")
    st.markdown("<div class='table-container-div'>", unsafe_allow_html=True)
    st.dataframe(df.head())
    st.markdown("</div>", unsafe_allow_html=True)


    # Data Analysis & Visualizations
    st.subheader("Key Metrics & Reports")

    # 1. Quotations by Sales Person
    if 'Sales Person' in df.columns:
        quotations_by_sales_person = df['Sales Person'].value_counts().reset_index()
        quotations_by_sales_person.columns = ['Sales Person', 'Number of Quotations']
        st.write("#### Number of Quotations by Sales Person:")
        st.markdown("<div class='table-container-div'>", unsafe_allow_html=True)
        st.dataframe(quotations_by_sales_person)
        st.markdown("</div>", unsafe_allow_html=True)

    # 2. Quotations by Product
    if 'Product' in df.columns:
        quotations_by_product = df['Product'].value_counts().reset_index()
        quotations_by_product.columns = ['Product', 'Number of Quotations']
        st.write("#### Number of Quotations by Product:")
        st.markdown("<div class='table-container-div'>", unsafe_allow_html=True)
        st.dataframe(quotations_by_product.head(10)) # Display top 10 for brevity
        st.markdown("</div>", unsafe_allow_html=True)

    # 3. Monthly Quotation Trends
    if 'Date' in df.columns:
        df['Month'] = df['Date'].dt.to_period('M')
        monthly_quotations = df['Month'].value_counts().sort_index().reset_index()
        monthly_quotations.columns = ['Month', 'Number of Quotations']
        monthly_quotations['Month'] = monthly_quotations['Month'].astype(str)
        st.write("#### Monthly Quotation Trends:")
        st.markdown("<div class='table-container-div'>", unsafe_allow_html=True)
        st.dataframe(monthly_quotations)
        st.markdown("</div>", unsafe_allow_html=True)

    # --- Advanced Visualizations ---
    st.subheader("Advanced Visualizations")

    # Set Matplotlib style for dark theme
    plt.style.use('dark_background')
    plt.rcParams.update({
        'axes.facecolor': '#294242', # secondary-bg
        'figure.facecolor': '#294242', # secondary-bg
        'text.color': 'white',
        'axes.labelcolor': 'white',
        'xtick.color': 'white',
        'ytick.color': 'white',
        'grid.color': '#3b5e5e', # border-color
        'grid.linestyle': '--',
        'grid.linewidth': 0.5,
        'axes.edgecolor': '#3b5e5e',
        'boxplot.boxprops.color': 'white',
        'boxplot.whiskerprops.color': 'white',
        'boxplot.capprops.color': 'white',
        'boxplot.medianprops.color': 'white',
        'patch.edgecolor': 'white',
    })

    st.markdown("<div class='chart-panel-container'>", unsafe_allow_html=True)
    # Plot 1: Quotations by Sales Person
    if 'Sales Person' in df.columns:
        fig1, ax1 = plt.subplots(figsize=(10, 6))
        sns.barplot(x='Number of Quotations', y='Sales Person', data=quotations_by_sales_person, palette='viridis', ax=ax1)
        ax1.set_title('Number of Quotations by Sales Person (2025)')
        ax1.set_xlabel('Number of Quotations')
        ax1.set_ylabel('Sales Person')
        st.pyplot(fig1)

    # Plot 2: Top Products by Quotations
    if 'Product' in df.columns:
        fig2, ax2 = plt.subplots(figsize=(12, 7))
        sns.barplot(x='Number of Quotations', y='Product', data=quotations_by_product.head(10), palette='magma', ax=ax2)
        ax2.set_title('Top 10 Products by Number of Quotations (2025)')
        ax2.set_xlabel('Number of Quotations')
        ax2.set_ylabel('Product')
        st.pyplot(fig2)

    # Plot 3: Monthly Quotation Trends
    if 'Date' in df.columns:
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
    st.markdown("</div>", unsafe_allow_html=True) # Close chart-panel-container


def process_2025_inv(df: pd.DataFrame):
    """
    Processes the '2025 INV' data.
    """
    st.markdown("<h2 class='st-emotion-cache-10grg6x e10grg6x4'>Invoices 2025</h2>", unsafe_allow_html=True)

    # Basic Data Cleaning
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df.dropna(subset=['Date'], inplace=True)
    df.columns = df.columns.str.strip() # Clean column names

    # --- Key Metrics ---
    st.markdown("<div class='key-metrics-container'>", unsafe_allow_html=True)
    invoice_count = df['INV No.'].nunique() if 'INV No.' in df.columns else len(df)
    salesperson_count = df['Sales Person'].nunique() if 'Sales Person' in df.columns else 0

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <p>Invoice Count</p>
            <p class="metric-value">{invoice_count}</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <p>Salespeople Count</p>
            <p class="metric-value">{salesperson_count}</p>
        </div>
        """, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.subheader("Raw Data Preview")
    st.markdown("<div class='table-container-div'>", unsafe_allow_html=True)
    st.dataframe(df.head())
    st.markdown("</div>", unsafe_allow_html=True)

    st.subheader("Key Metrics & Reports (Example)")
    st.write("Total Invoices:", invoice_count)
    st.write("Unique Sales Persons:", salesperson_count)

    st.markdown("""
    **Further analysis could include:**
    - Total invoice amount over time (if 'Amount' column exists and is numeric)
    - Invoices by Reseller/End User
    - Top Suppliers and Products in invoices
    - Sales performance by Sales Person (based on invoice value)
    """)
    if 'Sales Person' in df.columns:
        st.write("#### Invoices by Sales Person:")
        plt.style.use('dark_background')
        fig, ax = plt.subplots(figsize=(10, 6))
        sns.countplot(y='Sales Person', data=df, palette='GnBu', ax=ax)
        ax.set_title('Invoices by Sales Person')
        ax.set_xlabel('Number of Invoices')
        ax.set_ylabel('Sales Person')
        st.pyplot(fig)


def process_meeting_agenda(df: pd.DataFrame):
    """
    Processes the 'Meeting Agenda' data.
    """
    st.markdown("<h2 class='st-emotion-cache-10grg6x e10grg6x4'>Meeting Agenda</h2>", unsafe_allow_html=True)

    # Basic Data Cleaning for Meeting Agenda
    df['Order Value Approx.'] = pd.to_numeric(
        df['Order Value Approx.'].astype(str).str.replace(',', ''), errors='coerce'
    )
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df.dropna(subset=['Date', 'Order Value Approx.'], inplace=True)
    df.columns = df.columns.str.strip()

    # --- Key Metrics ---
    st.markdown("<div class='key-metrics-container'>", unsafe_allow_html=True)
    total_meeting_points = df['No:'].nunique() if 'No:' in df.columns else len(df)
    total_approx_order_value = df['Order Value Approx.'].sum() if 'Order Value Approx.' in df.columns and pd.api.types.is_numeric_dtype(df['Order Value Approx.']) else 0

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <p>Total Meeting Points</p>
            <p class="metric-value">{total_meeting_points}</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <p>Total Approximate Order Value</p>
            <p class="metric-value">${total_approx_order_value:,.0f}</p>
        </div>
        """, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


    st.subheader("Raw Data Preview")
    st.markdown("<div class='table-container-div'>", unsafe_allow_html=True)
    st.dataframe(df.head())
    st.markdown("</div>", unsafe_allow_html=True)

    st.subheader("Key Metrics & Reports (Example)")
    st.write(f"Total Meeting Points: {total_meeting_points}")
    st.write(f"Total Approximate Order Value: $ {total_approx_order_value:,.2f}")

    st.markdown("""
    **Further analysis could include:**
    - Average Order Value per meeting point
    - Meeting points by Action By person
    - Distribution of Margin
    - Trends in meeting topics or order values over time
    """)
    if 'Action By' in df.columns:
        st.write("#### Points by Action By:")
        plt.style.use('dark_background')
        fig, ax = plt.subplots(figsize=(10, 6))
        sns.countplot(y='Action By', data=df, palette='Spectral', ax=ax)
        ax.set_title('Meeting Points by Action By Person')
        ax.set_xlabel('Number of Points')
        ax.set_ylabel('Action By')
        st.pyplot(fig)

    if 'Margin' in df.columns and pd.api.types.is_numeric_dtype(df['Margin']):
         st.write("#### Margin Distribution:")
         plt.style.use('dark_background')
         fig, ax = plt.subplots(figsize=(10, 6))
         sns.histplot(df['Margin'].dropna(), kde=True, ax=ax, color='teal')
         ax.set_title('Margin Distribution')
         ax.set_xlabel('Margin')
         ax.set_ylabel('Frequency')
         st.pyplot(fig)


def process_payment_pending(df: pd.DataFrame):
    """
    Processes the 'Payment Pending' data.
    """
    st.markdown("<h2 class='st-emotion-cache-10grg6x e10grg6x4'>Payment Pending</h2>", unsafe_allow_html=True)

    # Basic Data Cleaning for Payment Pending
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
    df.dropna(subset=['Amount', 'PARTY NAME'], inplace=True)
    df.columns = df.columns.str.strip()

    # --- Key Metrics ---
    st.markdown("<div class='key-metrics-container'>", unsafe_allow_html=True)
    total_pending_amount = df['Amount'].sum()
    num_parties_pending = df['PARTY NAME'].nunique()

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <p>Total Pending Amount</p>
            <p class="metric-value">${total_pending_amount:,.0f}</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <p>Parties with Pending Payments</p>
            <p class="metric-value">{num_parties_pending}</p>
        </div>
        """, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


    st.subheader("Raw Data Preview")
    st.markdown("<div class='table-container-div'>", unsafe_allow_html=True)
    st.dataframe(df.head())
    st.markdown("</div>", unsafe_allow_html=True)

    st.subheader("Key Metrics & Reports (Example)")
    st.write(f"Total Pending Amount: $ {total_pending_amount:,.2f}")
    st.write("Number of Parties with Pending Payments:", num_parties_pending)

    st.markdown("""
    **Further analysis could include:**
    - Top parties with highest pending amounts
    - Distribution of pending amounts
    - Contact person analysis
    - Aging of payments (if a 'Due Date' or similar column is available)
    """)
    if 'PARTY NAME' in df.columns:
        st.write("#### Top 10 Parties by Pending Amount:")
        st.markdown("<div class='table-container-div'>", unsafe_allow_html=True)
        top_parties = df.groupby('PARTY NAME')['Amount'].sum().nlargest(10).reset_index()
        st.dataframe(top_parties)
        st.markdown("</div>", unsafe_allow_html=True)

        plt.style.use('dark_background')
        fig, ax = plt.subplots(figsize=(10, 6))
        sns.barplot(x='Amount', y='PARTY NAME', data=top_parties, palette='coolwarm', ax=ax)
        ax.set_title('Top 10 Parties by Pending Amount')
        ax.set_xlabel('Amount Pending ($)')
        ax.set_ylabel('Party Name')
        st.pyplot(fig)


def process_quotation_register_2023(df: pd.DataFrame):
    """
    Processes the 'Quotation Register 2023' data.
    Note: This sheet has generic column names, so specific analysis is limited without more info.
    """
    st.markdown("<h2 class='st-emotion-cache-10grg6x e10grg6x4'>Quotation Register 2023</h2>", unsafe_allow_html=True)
    st.warning("This sheet has generic column names (e.g., Unnamed: 0), making specific analysis challenging without more context.")
    st.subheader("Raw Data Preview")
    st.markdown("<div class='table-container-div'>", unsafe_allow_html=True)
    st.dataframe(df.head())
    st.markdown("</div>", unsafe_allow_html=True)
    st.write("Columns detected:", df.columns.tolist())
    st.markdown("""
    To perform meaningful analysis on this sheet, please provide more information about what each column represents.
    """)


# --- Streamlit UI Components (rebuilt using design) ---

st.markdown(
    "<h2 class='st-emotion-cache-10grg6x e10grg6x4'>Salahuddin Softech Solutions Dashboard</h2>",
    unsafe_allow_html=True
)
st.markdown(
    "<p class='st-emotion-cache-10grg6x e10grg6x4' style='text-align: center; margin-bottom: 1.5rem;'>Upload your Excel workbook to unlock powerful data analysis and visualization.</p>",
    unsafe_allow_html=True
)

# File Upload Section
st.markdown(
    """
    <div class="flex flex-col p-4 main-content-section">
      <div class="flex flex-col items-center gap-6 rounded-xl border-2 border-dashed border-[#3b5e5e] px-6 py-14">
        <div class="flex max-w-[480px] flex-col items-center gap-2">
          <p class="text-white text-lg font-bold leading-tight tracking-[-0.015em] max-w-[480px] text-center">Drag and drop your Excel file here</p>
          <p class="text-white text-sm font-normal leading-normal max-w-[480px] text-center">Or click to browse your files</p>
        </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)
# Inject the Streamlit file uploader *after* the custom HTML for the dropzone
uploaded_file = st.file_uploader(
    "Upload your SSS Master Sheet Excel workbook (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=False,
    key="excel_uploader",
    label_visibility="collapsed" # Hide default label, as we have custom text
)

# Dictionary to map sheet names to their respective skiprows values
sheet_skiprows_map = {
    "Payment Pending": 1,
    "Quotation Register 2023": 879,
    "Meeting Agenda": 2,
    # For '2025 INV' and 'QT Register 2025', assuming header is at row 0 (default)
}

# Initialize session state for processed dataframes
if 'processed_dataframes' not in st.session_state:
    st.session_state['processed_dataframes'] = {}

if uploaded_file:
    # Check if a new file is uploaded or if the file has changed
    if st.session_state.get('last_uploaded_file_id') != uploaded_file.file_id:
        st.session_state['processed_dataframes'] = {} # Clear dataframes if new file
        st.session_state['last_uploaded_file_id'] = uploaded_file.file_id # Store current file ID

        try:
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names

            # Process all sheets and store them in session state for later use
            for sheet_name in sheet_names:
                skip_rows = sheet_skiprows_map.get(sheet_name, 0)
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=skip_rows)
                st.session_state['processed_dataframes'][sheet_name] = df

            st.success("Excel file loaded and sheets processed!")

        except Exception as e:
            st.error(f"Error processing the Excel file: {e}")
            st.write("Please ensure the uploaded file is a valid .xlsx workbook and that the sheet names and data formats are consistent.")
            st.session_state['processed_dataframes'] = {} # Clear on error

    # --- Data Analysis & Reports Section ---
    if st.session_state['processed_dataframes']:
        st.markdown(
            "<div class='main-content-section flex h-full min-h-[700px] flex-col justify-between bg-[--primary-bg] p-4'>",
            unsafe_allow_html=True
        )
        st.markdown("<h1 class='st-emotion-cache-10grg6x e10grg6x4'>Data Sheets</h1>", unsafe_allow_html=True)

        # Sheet Selection
        available_sheets = list(st.session_state['processed_dataframes'].keys())
        # Provide a default selection for common sheets if they exist
        default_selection = [s for s in available_sheets if any(keyword in s for keyword in ["QT Register 2025", "2025 INV", "Meeting Agenda", "Payment Pending", "Quotation Register 2023"])]
        if not default_selection and available_sheets: # If no common sheets, select the first one
            default_selection = [available_sheets[0]]

        selected_sheets_for_analysis = st.multiselect(
            "Sheet Selection", # Label for the multiselect
            options=available_sheets,
            default=default_selection,
            label_visibility="collapsed" # Hide default label to use custom one
        )
        st.markdown(
            """
            <div class="flex items-center gap-3 px-3 py-2 rounded-full bg-[--secondary-bg]" style="margin-top: -30px; margin-bottom: 1rem;">
                <div class="text-white" data-icon="File" data-size="24px" data-weight="fill">
                    <svg xmlns="http://www.w3.org/2000/svg" width="24px" height="24px" fill="currentColor" viewBox="0 0 256 256">
                        <path d="M213.66,82.34l-56-56A8,8,0,0,0,152,24H56A16,16,0,0,0,40,40V216a16,16,0,0,0,16,16H200a16,16,0,0,0,16-16V88A8,8,0,0,0,213.66,82.34ZM152,88V44l44,44Z"></path>
                    </svg>
                </div>
                <p class="text-white text-sm font-medium leading-normal">Sheet Selection</p>
            </div>
            """, unsafe_allow_html=True
        )


        # Apply button (just for visual representation, Streamlit multiselect updates instantly)
        st.button(
            "Apply",
            key="apply_sheets",
            help="Click to apply selected sheets for analysis (selection applies automatically).",
            use_container_width=True # Make button full width as per design
        )

        if not selected_sheets_for_analysis:
            st.info("Please select at least one sheet for analysis from the multi-select above.")
        else:
            for sheet_name in selected_sheets_for_analysis:
                df = st.session_state['processed_dataframes'][sheet_name].copy() # Use a copy for display/analysis
                with st.expander(f"Analysis for Sheet: '{sheet_name}'", expanded=True):
                    if "QT Register 2025" in sheet_name:
                        # Reusing internal functions with new styling
                        process_qt_register_2025(df)
                    elif "2025 INV" in sheet_name:
                        process_2025_inv(df)
                    elif "Meeting Agenda" in sheet_name:
                        process_meeting_agenda(df)
                    elif "Payment Pending" in sheet_name:
                        process_payment_pending(df)
                    elif "Quotation Register 2023" in sheet_name:
                        process_quotation_register_2023(df)
                    else:
                        st.write(f"No specific processing logic defined for sheet: '{sheet_name}'. Displaying raw data.")
                        st.dataframe(df.head())
        st.markdown("</div>", unsafe_allow_html=True) # Close the main-content-container


    else:
        st.info("Please upload an Excel workbook from the file uploader above to view Data Analysis & Reports.")

else:
    st.info("Please upload your Excel workbook (.xlsx) from the file uploader above to get started.")
