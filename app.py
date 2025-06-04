import streamlit as st
import pandas as pd
import plotly.express as px
import openai
from datetime import datetime

# --- CONFIG ---
st.set_page_config(page_title="SSS Internal Dashboard", layout="wide")

# --- FILE UPLOAD ---
st.sidebar.title("📁 Upload Files")
uploaded_file = st.sidebar.file_uploader("Upload the 'SSS Master Sheet.xlsx'", type=["xlsx"])
api_key = st.sidebar.text_input("🔑 Enter your OpenAI API Key", type="password")

if not uploaded_file or not api_key:
    st.warning("Please upload the Excel file and enter your OpenAI API key to continue.")
    st.stop()

openai.api_key = api_key

# --- LOAD SHEETS ---
def load_data(file):
    xl = pd.read_excel(file, sheet_name=None)
    return xl

data = load_data(uploaded_file)

# --- UTILS ---
def parse_date(date):
    try:
        return pd.to_datetime(date)
    except:
        return pd.NaT

def monthly_summary(df, date_column, value_column):
    df[date_column] = pd.to_datetime(df[date_column], errors='coerce')
    df = df.dropna(subset=[date_column])
    df['Month'] = df[date_column].dt.to_period('M')
    return df.groupby('Month')[value_column].sum().reset_index()

# --- PAGES ---
st.sidebar.title("📊 Dashboard Pages")
page = st.sidebar.selectbox("Choose a section", [
    "Home Dashboard",
    "Quotations",
    "Invoices",
    "Payments",
    "Subscriptions",
    "Tasks & Meetings",
    "Ask AI Assistant"
])

# --- HOME PAGE ---
if page == "Home Dashboard":
    st.title("📊 SSS Internal Business Dashboard")

    total_quotations = sum(len(data[sheet]) for sheet in data if sheet.startswith("QT Register"))
    total_invoices = len(data.get("2025 INV", pd.DataFrame()))
    total_pending = data.get("Payment Pending", pd.DataFrame())["Amount"].sum()
    total_subs = len(data.get("Subscriptions", pd.DataFrame()))

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Quotations", total_quotations)
    col2.metric("Total Invoices", total_invoices)
    col3.metric("Pending Payments", f"BHD {total_pending:,.2f}")
    col4.metric("Subscriptions", total_subs)

    all_qts = pd.concat([data[sheet] for sheet in data if sheet.startswith("QT Register")])
    all_qts["Date"] = pd.to_datetime(all_qts["Date"], errors="coerce")
    qt_monthly = all_qts.dropna(subset=["Date"]).groupby(all_qts["Date"].dt.to_period("M")).size().reset_index(name="Quotations")
    qt_monthly["Date"] = qt_monthly["Date"].astype(str)
    st.plotly_chart(px.bar(qt_monthly, x="Date", y="Quotations", title="Monthly Quotations"), use_container_width=True)

elif page == "Quotations":
    st.title("📄 Quotation Register")
    qts = pd.concat([data[s] for s in data if s.startswith("QT Register")])
    st.dataframe(qts, use_container_width=True)

elif page == "Invoices":
    st.title("🧾 Invoice Records")
    inv = data.get("2025 INV", pd.DataFrame())
    st.dataframe(inv, use_container_width=True)
    if "Date" in inv.columns:
        inv["Date"] = pd.to_datetime(inv["Date"], errors="coerce")
        inv_summary = inv.dropna(subset=["Date"]).groupby(inv["Date"].dt.to_period("M")).size().reset_index(name="Invoices")
        inv_summary["Date"] = inv_summary["Date"].astype(str)
        st.plotly_chart(px.line(inv_summary, x="Date", y="Invoices", title="Monthly Invoices"), use_container_width=True)

elif page == "Payments":
    st.title("💰 Payment Pending")
    pay = data.get("Payment Pending", pd.DataFrame())
    if not pay.empty:
        st.dataframe(pay, use_container_width=True)
        client_summary = pay.groupby("Client")["Amount"].sum().reset_index()
        fig = px.bar(client_summary, x="Client", y="Amount", title="Pending Amount by Client")
        st.plotly_chart(fig, use_container_width=True)

elif page == "Subscriptions":
    st.title("📦 Active Subscriptions")
    subs = data.get("Subscriptions", pd.DataFrame())
    if not subs.empty:
        st.dataframe(subs, use_container_width=True)
        if "End Date" in subs.columns:
            subs["End Date"] = pd.to_datetime(subs["End Date"], errors="coerce")
            upcoming = subs[subs["End Date"] > datetime.now()]
            st.subheader("Upcoming Expirations")
            st.dataframe(upcoming.sort_values("End Date"), use_container_width=True)

elif page == "Tasks & Meetings":
    st.title("📋 Tasks and Meetings")
    tasks = data.get("To Remember", pd.DataFrame())
    meetings = data.get("Meeting Agenda", pd.DataFrame())
    if not tasks.empty:
        st.subheader("📝 To-Do List")
        st.dataframe(tasks, use_container_width=True)
    if not meetings.empty:
        st.subheader("📅 Meeting Agenda")
        st.dataframe(meetings, use_container_width=True)

elif page == "Ask AI Assistant":
    st.title("🤖 Ask AI about Your Business")
    user_question = st.text_area("Type your question about your data")

    if st.button("Get Insight") and user_question:
        context_str = ""
        for sheet_name, df in data.items():
            context_str += f"\nSheet: {sheet_name}\n{df.head(50).to_csv(index=False)}"

        messages = [
            {"role": "system", "content": "You are a smart business analyst assistant. Answer user questions based only on the provided Excel sheet data."},
            {"role": "user", "content": f"Here is the data:\n{context_str}"},
            {"role": "user", "content": user_question}
        ]

        with st.spinner("Thinking..."):
            response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=messages
            )
            answer = response.choices[0].message.content
            st.success("AI Insight:")
            st.write(answer)
