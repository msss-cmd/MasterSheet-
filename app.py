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

sheet_df = data[selected_sheet].copy()
st.subheader(f"📄 Viewing Sheet: {selected_sheet}")
st.dataframe(sheet_df)

# --- Auto Visualization ---
st.subheader("📊 Auto Visualizations")
try:
    numeric_cols = sheet_df.select_dtypes(include='number').columns.tolist()
    date_cols = sheet_df.select_dtypes(include='datetime').columns.tolist()

    if numeric_cols:
        st.markdown("#### 💸 Summary Stats")
        summary = sheet_df[numeric_cols].describe().transpose()[["mean", "sum"]]
        summary.columns = ["Average", "Total (BHD)"]
        st.dataframe(summary)

        for col in numeric_cols:
            st.markdown(f"**{col} Distribution**")
            fig, ax = plt.subplots()
            sns.histplot(sheet_df[col].dropna(), kde=True, ax=ax)
            ax.set_title(f"Distribution of {col}")
            ax.set_xlabel(f"{col} (BHD)" if "amount" in col.lower() else col)
            st.pyplot(fig)

    if date_cols and numeric_cols:
        date_col = date_cols[0]
        st.markdown(f"#### ⏱️ Time Series - {numeric_cols[0]} over Time")
        sheet_df[date_col] = pd.to_datetime(sheet_df[date_col], errors='coerce')
        df_time = sheet_df[[date_col, numeric_cols[0]]].dropna()
        df_time = df_time.groupby(df_time[date_col].dt.to_period("M")).sum().reset_index()
        df_time[date_col.name] = df_time[date_col.name].astype(str)
        st.line_chart(data=df_time.set_index(date_col.name))
except Exception as e:
    st.warning(f"Could not auto-generate visualizations: {e}")

# --- AI Assistant per Sheet ---
st.subheader("🤖 Ask a Question About This Sheet")
user_question = st.text_input("Ask a question related to this sheet:")

if user_question:
    try:
        sample_data = sheet_df.head(10).to_string(index=False)
        messages = [
            {"role": "system", "content": "You are a business data analyst helping the user understand Excel sheets with financial and operational data. All currency is in Bahraini Dinar (BHD)."},
            {"role": "user", "content": f"Here is the data from the '{selected_sheet}' sheet:\n\n{sample_data}\n\nNow answer this question:\n{user_question}"}
        ]
        response = openai.chat.completions.create(
            model="gpt-4",
            messages=messages,
            temperature=0.3
        )
        st.success(response.choices[0].message.content)
    except Exception as e:
        st.error(f"Failed to get response from OpenAI: {str(e)}")
