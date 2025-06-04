import streamlit as st
import pandas as pd
import openai
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO

st.set_page_config(page_title="SSS Operations Dashboard", layout="wide")
st.title("📊 SSS Operations Dashboard")

# Upload OpenAI API key
openai_api_key = st.text_input("Enter your OpenAI API Key:", type="password")
if not openai_api_key:
    st.warning("Please enter your OpenAI API key to continue.")
    st.stop()
openai.api_key = openai_api_key

# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel File", type=["xlsx"])
if not uploaded_file:
    st.warning("Please upload the SSS Master Sheet Excel file.")
    st.stop()

# Read all sheets from the Excel file
xlsx = pd.ExcelFile(uploaded_file)
data = {sheet: xlsx.parse(sheet) for sheet in xlsx.sheet_names}

st.sidebar.header("Sheets Overview")
selected_sheet = st.sidebar.selectbox("Select Sheet to View", options=list(data.keys()))
st.subheader(f"📄 Viewing Sheet: {selected_sheet}")
st.dataframe(data[selected_sheet])

# --- Insights ---
if "Payment Pending" in data:
    payment_pending = data["Payment Pending"]
    if "Amount" in payment_pending.columns:
        total_pending = payment_pending["Amount"].sum()
        st.metric("💰 Total Pending Amount", f"BHD {total_pending:,.2f}")

if "Quotation" in data:
    all_qts = data["Quotation"].copy()
    if "Date" in all_qts.columns:
        try:
            all_qts["Date"] = pd.to_datetime(all_qts["Date"], errors='coerce')
            qt_monthly = (
                all_qts.dropna(subset=["Date"])
                .assign(Month=lambda df: df["Date"].dt.to_period("M"))
                .groupby("Month").size().reset_index(name="Quotations")
            )
            qt_monthly["Month"] = qt_monthly["Month"].astype(str)
            st.subheader("📈 Quotations Over Time")
            st.line_chart(qt_monthly.set_index("Month"))
        except Exception as e:
            st.warning(f"Could not generate quotation timeline: {e}")

# --- AI Assistant ---
st.subheader("🤖 Ask Questions About Your Data")
user_question = st.text_input("Ask a question about your operations data:")

if user_question:
    context_text = "\n\n".join([
        f"Sheet: {sheet}\n" + data[sheet].head(5).to_string(index=False)
        for sheet in data
    ])
    try:
        response = openai.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a helpful assistant specialized in reading and summarizing business operations data from Excel sheets."},
                {"role": "user", "content": f"Here is the data:\n{context_text}\n\nNow answer this question: {user_question}"},
            ],
            temperature=0.3
        )
        st.success(response.choices[0].message.content)
    except Exception as e:
        st.error(f"Failed to get response from OpenAI: {str(e)}")
