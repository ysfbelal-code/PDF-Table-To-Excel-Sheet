import streamlit as st
import pandas as pd
import tempfile
import os
import sys
from pathlib import Path

# Ensure backend is importable (same directory)
sys.path.insert(0, str(Path(__file__).parent))
from backend import enterprise_grade_table_miner

st.set_page_config(page_title="PDF Table to Excel", layout="centered")

st.title("📄 PDF Table → Excel")
st.markdown("Extract tables from any PDF (scanned or digital, English + Arabic)")

uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

mode = st.radio(
    "Table handling",
    options=["Separate sheets", "Combined table"],
    index=0,
    help="Separate = each table as its own sheet · Combined = same headers merged, near‑duplicates (1 cell diff) keep first"
)

if uploaded_file is not None:
    # Show file info
    st.info(f"File: {uploaded_file.name} ({uploaded_file.size / 1024:.1f} KB)")

    if st.button("Convert to Excel", type="primary"):
        # Save uploaded PDF to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(uploaded_file.getbuffer())
            tmp_pdf_path = tmp_pdf.name

        # Output Excel in temporary location
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
            tmp_xlsx_path = tmp_xlsx.name

        # Determine combine mode
        combine_mode = "combined" if mode == "Combined table" else "separate"

        # Run conversion with spinner
        with st.spinner("Running OCR & table detection — this may take 30–60s..."):
            try:
                enterprise_grade_table_miner(tmp_pdf_path, tmp_xlsx_path, combine_mode=combine_mode)
                st.success("Conversion completed successfully!")

                # Read the generated Excel to display preview
                excel_data = pd.read_excel(tmp_xlsx_path, sheet_name=None)

                # Show preview of each sheet
                for sheet_name, df in excel_data.items():
                    st.subheader(f"Sheet: {sheet_name}")
                    st.dataframe(df.head(10), use_container_width=True)
                    if len(df) > 10:
                        st.caption(f"Showing first 10 of {len(df)} rows")

                # Provide download button for the Excel file
                with open(tmp_xlsx_path, "rb") as f:
                    excel_bytes = f.read()
                st.download_button(
                    label="📥 Download Excel file",
                    data=excel_bytes,
                    file_name=f"{Path(uploaded_file.name).stem}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Conversion failed: {e}")
            finally:
                # Clean up temporary files
                os.unlink(tmp_pdf_path)
                if os.path.exists(tmp_xlsx_path):
                    os.unlink(tmp_xlsx_path)
else:
    st.info("Please upload a PDF file to begin.")