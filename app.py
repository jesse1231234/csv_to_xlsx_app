import streamlit as st
import pandas as pd
from io import BytesIO
from script1 import process_echo360
from script2 import process_gradebook


def main():
    # Configure the Streamlit page
    st.set_page_config(page_title="CSV → XLSX Processor", layout="wide")
    st.title("📊 CSV → XLSX Processor")
    st.markdown(
        """
        Upload your two CSVs below. Once both are provided, this app will:
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

    # When both CSVs are uploaded, process and generate the Excel
    if uploaded_echo and uploaded_grade:
        try:
            # Read CSVs into DataFrames
            df_echo = pd.read_csv(uploaded_echo)
            df_grade = pd.read_csv(uploaded_grade)

            # Apply your existing processing functions
            out1 = process_echo360(df_echo)
            out2 = process_gradebook(df_grade)

            # Create an in-memory buffer for the Excel file
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                # Write data to separate sheets
                out1.to_excel(writer, sheet_name="Echo360 Output", index=False)
                out2.to_excel(writer, sheet_name="Gradebook Output", index=False)

                # Grab the workbook and worksheets
                workbook = writer.book
                ws1 = writer.sheets["Echo360 Output"]
                ws2 = writer.sheets["Gradebook Output"]

                # Calculate table dimensions
                nrows1, ncols1 = out1.shape
                nrows2, ncols2 = out2.shape

                # 1) Format as native Excel Tables
                ws1.add_table(0, 0, nrows1, ncols1 - 1, {
                    "columns": [{"header": h} for h in out1.columns],
                    "style":   "Table Style Medium 9"
                })
                ws2.add_table(0, 0, nrows2, ncols2 - 1, {
                    "columns": [{"header": h} for h in out2.columns],
                    "style":   "Table Style Medium 2"
                })

                # 2) Conditional formatting for Echo360 sheet
                # Data bars on Video Duration (col B)
                ws1.conditional_format(f"B2:B{nrows1+1}", {"type": "data_bar"})
                # 3-color scale on Average View % (col D)
                ws1.conditional_format(f"D2:D{nrows1+1}", {
                    "type":      "3_color_scale",
                    "min_type":  "percentile", "min_value": 10,
                    "mid_type":  "percentile", "mid_value": 50,
                    "max_type":  "percentile", "max_value": 90,
                })

                # 3) Conditional formatting for Gradebook sheet
                # Identify the column index for 'Final Grade'
                try:
                    idx = list(out2.columns).index('Final Grade')
                    col_letter = chr(ord('A') + idx)
                    ws2.conditional_format(
                        f"{col_letter}2:{col_letter}{nrows2+1}", {
                            "type":     "cell",
                            "criteria": "<",
                            "value":    60,
                            "format":   workbook.add_format({"bg_color": "#FFC7CE"})
                        }
                    )
                except ValueError:
                    pass  # 'Final Grade' column not found

                # 4) Time formatting on Video Duration (col B)
                time_fmt = workbook.add_format({"num_format": "hh:mm:ss"})
                ws1.set_column("B:B", 15, time_fmt)

                # 5) Add charts to Echo360 sheet
                chart1 = workbook.add_chart({"type": "line"})
                chart1.add_series({
                    "name":       "Average View %",
                    "categories": ["Echo360 Output", 1, 0, nrows1, 0],
                    "values":     ["Echo360 Output", 1, 3, nrows1, 3],
                })
                chart1.set_title({"name": "View % Over Time"})
                chart1.set_style(9)
                ws1.insert_chart("J2", chart1)

                chart2 = workbook.add_chart({"type": "line"})
                chart2.add_series({
                    "name":       "Number of Unique Viewers",
                    "categories": ["Echo360 Output", 1, 0, nrows1, 0],
                    "values":     ["Echo360 Output", 1, 2, nrows1, 2],
                })
                chart2.set_title({"name": "Unique Viewers Over Time"})
                chart2.set_style(9)
                ws1.insert_chart("J20", chart2)

                # 6) Bold the header row on Gradebook sheet
                header_fmt = workbook.add_format({"bold": True})
                ws2.set_row(0, None, header_fmt)

            # Move buffer cursor to the start
            buffer.seek(0)

            # Provide download button
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