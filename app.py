import streamlit as st
import pandas as pd
import tempfile
import os
import sys
import smtplib
from pathlib import Path
from email.message import EmailMessage
from datetime import datetime

# Ensure backend is importable
sys.path.insert(0, str(Path(__file__).parent))
from backend import enterprise_grade_table_miner

st.set_page_config(page_title="PDF Table to Excel", layout="centered")

st.title("📄 PDF Table → Excel")
st.markdown("Extract tables from any PDF (scanned or digital, English + Arabic)")

# ---------- Main conversion UI ----------
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

mode = st.radio(
    "Table handling",
    options=["Separate sheets", "Combined table"],
    index=0,
    help="Separate = each table as its own sheet · Combined = same headers merged, near‑duplicates (1 cell diff) keep first"
)

if uploaded_file is not None:
    st.info(f"File: {uploaded_file.name} ({uploaded_file.size / 1024:.1f} KB)")

    if st.button("Convert to Excel", type="primary"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(uploaded_file.getbuffer())
            tmp_pdf_path = tmp_pdf.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
            tmp_xlsx_path = tmp_xlsx.name

        combine_mode = "combined" if mode == "Combined table" else "separate"

        with st.spinner("Running OCR & table detection — this may take 30–60s..."):
            try:
                enterprise_grade_table_miner(tmp_pdf_path, tmp_xlsx_path, combine_mode=combine_mode)
                st.success("Conversion completed successfully!")

                excel_data = pd.read_excel(tmp_xlsx_path, sheet_name=None)

                for sheet_name, df in excel_data.items():
                    st.subheader(f"Sheet: {sheet_name}")
                    st.dataframe(df.head(10), use_container_width=True)
                    if len(df) > 10:
                        st.caption(f"Showing first 10 of {len(df)} rows")

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
                os.unlink(tmp_pdf_path)
                if os.path.exists(tmp_xlsx_path):
                    os.unlink(tmp_xlsx_path)
else:
    st.info("Please upload a PDF file to begin.")

# ---------- Report an Issue section ----------
st.divider()
with st.expander("📬 Report an Issue", expanded=False):
    st.markdown("If you encountered an error or unexpected behaviour, please submit a report below. "
                "We'll get back to you as soon as possible.")
    
    with st.form("report_form"):
        user_email = st.text_input("Your email (optional)", placeholder="so we can reply")
        problem_description = st.text_area("Describe the issue in detail", height=150,
                                           placeholder="What happened? What did you expect?")
        attach_pdf = st.file_uploader("Attach the PDF that caused the problem (optional)", type=["pdf"])
        
        submitted = st.form_submit_button("Send Report")
        
        if submitted:
            if not problem_description.strip():
                st.warning("Please describe the issue before submitting.")
            else:
                # Send email
                try:
                    # Use secrets for SMTP configuration
                    smtp_server = st.secrets["smtp"]["server"]
                    smtp_port = st.secrets["smtp"]["port"]
                    smtp_username = st.secrets["smtp"]["username"]
                    smtp_password = st.secrets["smtp"]["password"]
                    sender_email = st.secrets["smtp"]["sender"]
                    recipient_email = "ysfbelal@gmail.com"
                    
                    msg = EmailMessage()
                    msg["Subject"] = f"PDF Table Extractor Report - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                    msg["From"] = sender_email
                    msg["To"] = recipient_email
                    msg["Reply-To"] = user_email if user_email.strip() else sender_email
                    
                    body = f"Report submitted via Streamlit app:\n\n"
                    body += f"User email: {user_email if user_email.strip() else 'Not provided'}\n"
                    body += f"Description:\n{problem_description}\n\n"
                    body += f"---\nSystem info:\n"
                    body += f"Python version: {sys.version}\n"
                    body += f"Platform: {sys.platform}\n"
                    msg.set_content(body)
                    
                    # Attach PDF if provided
                    if attach_pdf is not None:
                        pdf_data = attach_pdf.getvalue()
                        msg.add_attachment(pdf_data,
                                           maintype='application',
                                           subtype='pdf',
                                           filename=attach_pdf.name)
                    
                    # Send via SMTP
                    with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                        server.login(smtp_username, smtp_password)
                        server.send_message(msg)
                    
                    st.success("✅ Report sent successfully! Thank you for your feedback.")
                
                except Exception as e:
                    st.error(f"Failed to send report: {e}")
                    st.info("Please try again later or contact us directly.")
