import streamlit as st
import pandas as pd
from io import BytesIO
from script1 import process_echo360
from script2 import process_gradebook


def main():
    # Configure page
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

    # Upload widgets
    col1, col2 = st.columns(2)
    with col1:
        uploaded_echo = st.file_uploader("1️⃣ Echo360 CSV", type="csv")
    with col2:
        uploaded_grade = st.file_uploader("2️⃣ Gradebook CSV", type="csv")

    # Generate and download Excel when both files are uploaded
    if uploaded_echo and uploaded_grade:
        try:
            # Read CSVs
            df_echo = pd.read_csv(uploaded_echo)
            df_grade = pd.read_csv(uploaded_grade)

            # Process with user functions
            out1 = process_echo360(df_echo)
            out2 = process_gradebook(df_grade)

            # Prepare in-memory Excel file
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                # Write data
                out1.to_excel(writer, sheet_name="Echo360 Output", index=False)
                out2.to_excel(writer, sheet_name="Gradebook Output", index=False)

                # Access workbook and worksheets
                workbook = writer.book
                ws1 = writer.sheets["Echo360 Output"]
                ws2 = writer.sheets["Gradebook Output"]

                # Calculate table ranges
                rows1 = len(out1) + 1  # include header
                cols1 = len(out1.columns)
                rows2 = len(out2) + 1  # include header
                cols2 = len(out2.columns)

                # Format as Excel Tables
                ws1.add_table(0, 0, rows1, cols1 - 1, {
                    "columns": [{"header": h} for h in out1.columns],
                    "style":   "Table Style Medium 9"
                })
                ws2.add_table(0, 0, rows2, cols2 - 1, {
                    "columns": [{"header": h} for h in out2.columns],
                    "style":   "Table Style Medium 2"
                })

                # Conditional formatting: Echo360
                # Data bars on Video Duration (col B)
                ws1.conditional_format(f"B2:B{rows1}", {"type": "data_bar"})
                # 3-color scale on Average View % (col D)
                ws1.conditional_format(f"D2:D{rows1}", {
                    "type":      "3_color_scale",
                    "min_type":  "percentile", "min_value": 10,
                    "mid_type":  "percentile", "mid_value": 50,
                    "max_type":  "percentile", "max_value": 90,
                })

                # Conditional formatting: Gradebook
                # Highlight Final Grade < 60% (assuming column G)
                ws2.conditional_format(f"G2:G{rows2}", {
                    "type":     "cell",
                    "criteria": "<",
                    "value":    60,
                    "format":   workbook.add_format({"bg_color": "#FFC7CE"})
                })

                # Formatting: time column in Echo360
                time_fmt = workbook.add_format({"num_format": "hh:mm:ss"})
                ws1.set_column("B:B", 15, time_fmt)

                # Add charts: Echo360
                # Chart 1: View % Over Time
                chart1 = workbook.add_chart({"type": "line"})
                chart1.add_series({
                    "name":       "Average View %",
                    "categories": ["Echo360 Output", 1, 0, rows1, 0],
                    "values":     ["Echo360 Output", 1, 3, rows1, 3],
                })
                chart1.set_title({"name": "View % Over Time"})
                chart1.set_style(9)
                ws1.insert_chart("J2", chart1)

                # Chart 2: Unique Viewers Over Time
                chart2 = workbook.add_chart({"type": "line"})
                chart2.add_series({
                    "name":       "Number of Unique Viewers",
                    "categories": ["Echo360 Output", 1, 0, rows1, 0],
                    "values":     ["Echo360 Output", 1, 2, rows1, 2],
                })
                chart2.set_title({"name": "Unique Viewers Over Time"})
                chart2.set_style(9)
                ws1.insert_chart("J20", chart2)

                # Gradebook: bold header
                header_fmt = workbook.add_format({"bold": True})
                ws2.set_row(0, None, header_fmt)

            # Seek to start and offer download
            buffer.seek(0)
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