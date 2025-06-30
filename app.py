import streamlit as st
import pandas as pd
from io import BytesIO
from script1 import process_echo360
from script2 import process_gradebook


def main():
    # Page config
    st.set_page_config(page_title="CSV → XLSX Processor", layout="wide")
    st.title("📊 CSV → XLSX Processor")
    st.markdown(
        """
        Upload your two CSVs below. When both are provided, this app will:
        1. Run your existing processing logic  
        2. Bundle the results into one `.xlsx` with separate sheets  
        3. Offer it for download
        """
    )

    # File upload widgets
    col1, col2 = st.columns(2)
    with col1:
        uploaded_echo = st.file_uploader("1️⃣ Echo360 CSV", type="csv")
    with col2:
        uploaded_grade = st.file_uploader("2️⃣ Gradebook CSV", type="csv")

    # Process and generate Excel
    if uploaded_echo and uploaded_grade:
        try:
            # Read CSVs into DataFrames
            df_echo = pd.read_csv(uploaded_echo)
            df_grade = pd.read_csv(uploaded_grade)

            # Apply processing functions
            out1 = process_echo360(df_echo)
            out2 = process_gradebook(df_grade)

            # Prepare in-memory Excel file
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                # Write DataFrames to separate sheets
                out1.to_excel(writer, sheet_name="Echo360 Output", index=False)
                out2.to_excel(writer, sheet_name="Gradebook Output", index=False)

                # Access workbook & worksheets
                workbook = writer.book
                ws1 = writer.sheets["Echo360 Output"]
                ws2 = writer.sheets["Gradebook Output"]

                # --- Echo360 sheet formatting & charts ---
                # 1) Format Video Duration column as hh:mm:ss
                time_fmt = workbook.add_format({"num_format": "hh:mm:ss"})
                ws1.set_column("B:B", 15, time_fmt)

                # 2) Data bars on Video Duration (col B)
                max_row = len(out1) + 1  # account for header row
                ws1.conditional_format(f"B2:B{max_row}", {"type": "data_bar"})

                # 3) Line chart: View % Over Time
                chart1 = workbook.add_chart({"type": "line"})
                chart1.add_series({
                    "name":       "Average View %",
                    "categories": ["Echo360 Output", 1, 0, max_row, 0],  # Media Title
                    "values":     ["Echo360 Output", 1, 3, max_row, 3],  # Average View %
                })
                chart1.set_title({"name": "View % Over Time"})
                chart1.set_style(9)
                ws1.insert_chart("J2", chart1)

                # 4) Line chart: Unique Viewers Over Time
                chart2 = workbook.add_chart({"type": "line"})
                chart2.add_series({
                    "name":       "Number of Unique Viewers",
                    "categories": ["Echo360 Output", 1, 0, max_row, 0],
                    "values":     ["Echo360 Output", 1, 2, max_row, 2],  # Unique Viewers
                })
                chart2.set_title({"name": "Unique Viewers Over Time"})
                chart2.set_style(9)
                ws1.insert_chart("J20", chart2)

                # --- Gradebook sheet formatting ---
                header_fmt = workbook.add_format({"bold": True})
                ws2.set_row(0, None, header_fmt)

            # Seek to start of the stream
            buffer.seek(0)

            # Download button
            st.download_button(
                label="💾 Download combined XLSX",
                data=buffer,
                file_name="combined_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Processing failed: {e}")


if __name__ == "__main__":
    main()
