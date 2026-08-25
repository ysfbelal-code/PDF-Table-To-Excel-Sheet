import streamlit as st
import pandas as pd
import tempfile
import os
import sys
import smtplib
from pathlib import Path
from email.message import EmailMessage
from datetime import datetime
import zipfile
import io

# Ensure backend is importable
sys.path.insert(0, str(Path(__file__).parent))
from backend import enterprise_grade_table_miner

st.set_page_config(page_title="PDF Table to Excel", layout="centered")

st.title("📄 PDF Table → Excel")
st.markdown("Extract tables from any PDF (scanned or digital, English + Arabic) · **Now with batch processing!**")

# ---------- Main conversion UI ----------
MAX_TOTAL_SIZE = 250 * 1024 * 1024  # 250 MB

uploaded_files = st.file_uploader(
    "Choose one or more PDF files",
    type="pdf",
    accept_multiple_files=True,
    help="Upload up to 250 MB total. Each PDF will be converted to its own Excel file."
)

mode = st.radio(
    "Table handling",
    options=["Separate sheets", "Combined table"],
    index=0,
    help="Separate = each table as its own sheet · Combined = same headers merged, near‑duplicates (1 cell diff) keep first"
)

if uploaded_files:
    # Calculate total size
    total_size = sum(f.size for f in uploaded_files)
    total_size_mb = total_size / (1024 * 1024)
    
    if total_size > MAX_TOTAL_SIZE:
        st.error(
            f"❌ Total file size ({total_size_mb:.1f} MB) exceeds 250 MB limit. "
            f"Please remove some files and try again."
        )
    else:
        st.info(
            f"📊 {len(uploaded_files)} file(s) selected · "
            f"{total_size_mb:.1f} MB total"
        )
        
        if st.button("Convert to Excel", type="primary"):
            combine_mode = "combined" if mode == "Combined table" else "separate"
            
            excel_files = {}  # {filename: bytes}
            progress_bar = st.progress(0)
            status_container = st.empty()
            
            for idx, uploaded_file in enumerate(uploaded_files):
                status_container.info(
                    f"Processing: {uploaded_file.name} ({idx + 1}/{len(uploaded_files)})"
                )
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    tmp_pdf.write(uploaded_file.getbuffer())
                    tmp_pdf_path = tmp_pdf.name
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
                    tmp_xlsx_path = tmp_xlsx.name
                
                try:
                    with st.spinner(f"Running OCR & table detection on {uploaded_file.name}..."):
                        enterprise_grade_table_miner(tmp_pdf_path, tmp_xlsx_path, combine_mode=combine_mode)
                    
                    # Read the generated Excel file
                    with open(tmp_xlsx_path, "rb") as f:
                        excel_bytes = f.read()
                    
                    # Use original filename stem + .xlsx
                    output_filename = f"{Path(uploaded_file.name).stem}.xlsx"
                    excel_files[output_filename] = excel_bytes
                    
                    status_container.success(f"✅ {uploaded_file.name} converted successfully!")
                    
                except Exception as e:
                    st.error(f"❌ Conversion failed for {uploaded_file.name}: {e}")
                
                finally:
                    os.unlink(tmp_pdf_path)
                    if os.path.exists(tmp_xlsx_path):
                        os.unlink(tmp_xlsx_path)
                
                # Update progress
                progress_bar.progress((idx + 1) / len(uploaded_files))
            
            st.divider()
            
            # Display results and downloads
            if excel_files:
                st.success(f"✅ Conversion completed! {len(excel_files)} Excel file(s) ready.")
                
                # Show preview for each file
                for excel_filename, excel_bytes in excel_files.items():
                    with st.expander(f"📋 Preview: {excel_filename}", expanded=False):
                        try:
                            excel_data = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=None)
                            for sheet_name, df in excel_data.items():
                                st.subheader(f"Sheet: {sheet_name}")
                                st.dataframe(df.head(10), use_container_width=True)
                                if len(df) > 10:
                                    st.caption(f"Showing first 10 of {len(df)} rows")
                        except Exception as e:
                            st.warning(f"Could not preview {excel_filename}: {e}")
                
                st.divider()
                
                # Download options
                col1, col2 = st.columns(2)
                
                with col1:
                    # Individual downloads
                    st.subheader("📥 Individual Downloads")
                    for excel_filename, excel_bytes in excel_files.items():
                        st.download_button(
                            label=f"Download {excel_filename}",
                            data=excel_bytes,
                            file_name=excel_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"download_{excel_filename}"
                        )
                
                with col2:
                    # Batch download as ZIP
                    st.subheader("📦 Batch Download")
                    if len(excel_files) > 1:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for excel_filename, excel_bytes in excel_files.items():
                                zip_file.writestr(excel_filename, excel_bytes)
                        
                        zip_buffer.seek(0)
                        st.download_button(
                            label="📦 Download All as ZIP",
                            data=zip_buffer.getvalue(),
                            file_name="pdf_tables_export.zip",
                            mime="application/zip"
                        )
                    else:
                        st.caption("(Single file: use individual download above)")
            
            progress_bar.empty()
            status_container.empty()
else:
    st.info("Please upload one or more PDF files to begin.")

# ---------- Report an Issue section ----------
st.divider()
with st.expander("📬 Feedback", expanded=False):
    st.markdown("If you encountered an error, unexpected behaviour, or want to request any new features, please submit a report below. "
                "We'll get back to you as soon as possible.")
    
    with st.form("report_form"):
        user_email = st.text_input("Your email (optional)", placeholder="Example: john_doe@example.com")
        problem_description = st.text_area("Insert your feedback here", height=150,
                                           placeholder="What happened? What did you expect?")
        attach_pdf = st.file_uploader("Attach a PDF that caused the problem (optional)", type=["pdf"])
        
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
                    st.info("Please try again later or contact us directly at ysfbelal@gmail.com.")
