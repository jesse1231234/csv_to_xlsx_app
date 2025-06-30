import streamlit as st
import pandas as pd
from io import BytesIO

# import your existing processing functions
from script1 import process_echo360
from script2 import process_gradebook

# --- UI ---
st.set_page_config(page_title="CSV → XLSX Processor", layout="wide")
st.title("📊 CSV → XLSX Processor")

st.markdown("""
Upload your two CSVs below. When both are provided, this app will:
1. Run your existing processing logic  
2. Bundle the results into one `.xlsx` with separate sheets  
3. Offer it for download
""")

col1, col2 = st.columns(2)
with col1:
    uploaded_echo = st.file_uploader("1️⃣ Echo360 CSV", type="csv")
with col2:
    uploaded_grade = st.file_uploader("2️⃣ Gradebook CSV", type="csv")

# --- Processing & Download ---
if uploaded_echo and uploaded_grade:
    try:
        # read into DataFrames
        df_echo  = pd.read_csv(uploaded_echo)
        df_grade = pd.read_csv(uploaded_grade)

        # apply your logic
        out1 = process_echo360(df_echo)
        out2 = process_gradebook(df_grade)

        # write to BytesIO
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            out1.to_excel(writer, sheet_name="Echo360 Output", index=False)
            out2.to_excel(writer, sheet_name="Gradebook Output", index=False)
        buf.seek(0)

        # download button
        st.download_button(
            label="💾 Download combined XLSX",
            data=buf,
            file_name="combined_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Processing failed: {e}")
