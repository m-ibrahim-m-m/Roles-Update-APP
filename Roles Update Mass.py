import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
import time
import re
import os
from difflib import get_close_matches

# Configure the page
st.set_page_config(
    page_title="Excel Role Consolidator",
    layout="wide",
)

# ---------- Custom CSS ----------
st.markdown("""
    <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .main-header { font-size: 2.2em; font-weight: bold; color: #4a90e2; text-align: center; margin-bottom: 20px; }
    .sub-header { font-size: 1.4em; color: #333; margin-top: 25px; margin-bottom: 10px; }
    .uploaded-file { background-color: #f0f8ff; padding: 8px; border-radius: 6px; margin: 5px 0; font-size: 0.9em; }
    .dataframe { font-size: 0.85em; }
    .stButton>button { border-radius: 10px; background-color: #4a90e2; color: white; padding: 6px 18px; font-size: 0.9em; }
    .stButton>button:hover { background-color: #357abd; }
    .file-card { border: 1px solid #ddd; border-radius: 8px; padding: 10px; margin: 5px 0; background-color: #f9f9f9; }
    .file-name { font-weight: bold; color: #2c3e50; }
    .file-size { font-size: 0.85em; color: #7f8c8d; }
    .footer { text-align: center; font-size: 0.85em; color: #999; margin-top: 30px; }
    .stDownloadButton>button { background: linear-gradient(90deg, #4a90e2, #357abd); color: white; font-size: 0.9em; padding: 8px 20px; border-radius: 8px; }
    .stDownloadButton>button:hover { background: linear-gradient(90deg, #357abd, #2c5aa0); }
    </style>
""", unsafe_allow_html=True)

# ---------- Helper Functions ----------
def format_file_size(size_bytes):
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 ** 2:
        return f"{size_bytes / 1024:.2f} KB"
    else:
        return f"{size_bytes / (1024**2):.2f} MB"

def normalize_name(name: str) -> str:
    """Normalize file/sheet names for matching (spaces, underscores, case)."""
    return name.strip().replace("_", " ").upper()

# ---------- Main App ----------
st.markdown("<div class='main-header'>📊 Excel Role Consolidator</div>", unsafe_allow_html=True)

# File uploader
uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    with st.expander("📂 Uploaded Files", expanded=False):
        for file in uploaded_files:
            file_size = format_file_size(file.size)
            st.markdown(f"<div class='file-card'><span class='file-name'>{file.name}</span> - <span class='file-size'>{file_size}</span></div>", unsafe_allow_html=True)

    # ---------- Step 1: Generate Template ----------
    st.markdown("<div class='sub-header'>Step 1️⃣: Generate Excel Template</div>", unsafe_allow_html=True)
    if st.button("Generate Template"):
        wb = Workbook()
        wb.remove(wb.active)

        for file in uploaded_files:
            xls = pd.ExcelFile(file)
            for sheet in xls.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet)
                ws = wb.create_sheet(sheet[:31])  # Excel limit

                # Write headers
                for i, col in enumerate(df.columns, start=1):
                    ws.cell(row=1, column=i, value=col)

                # Write data
                for r in range(len(df)):
                    for c in range(len(df.columns)):
                        ws.cell(row=r + 2, column=c + 1, value=df.iloc[r, c])

                # Add metadata
                ws.cell(row=1, column=len(df.columns) + 1, value="Source_File")
                ws.cell(row=1, column=len(df.columns) + 2, value="Source_Sheet")
                ws.cell(row=1, column=len(df.columns) + 3, value="Actions")

                for r in range(len(df)):
                    ws.cell(row=r + 2, column=len(df.columns) + 1, value=os.path.splitext(file.name)[0])
                    ws.cell(row=r + 2, column=len(df.columns) + 2, value=sheet)
                    ws.cell(row=r + 2, column=len(df.columns) + 3, value="")

                # Dropdown for Actions
                dv = DataValidation(type="list", formula1='"Add,Remove"', allow_blank=True)
                ws.add_data_validation(dv)
                dv.add(f"{chr(65 + len(df.columns) + 2)}2:{chr(65 + len(df.columns) + 2)}{len(df) + 1}")

        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        st.download_button("⬇️ Download Template", buffer, "template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ---------- Step 2: Upload Filled Template ----------
    st.markdown("<div class='sub-header'>Step 2️⃣: Upload Filled Template</div>", unsafe_allow_html=True)
    template_file = st.file_uploader("Upload filled template", type=["xlsx"], key="template")

    if template_file:
        df_template = pd.read_excel(template_file)

        consolidated_data = []

        all_uploaded_names = [normalize_name(os.path.splitext(f.name)[0]) for f in uploaded_files]

        for _, row in df_template.iterrows():
            src_file = normalize_name(str(row.get("Source_File", "")))
            src_sheet = str(row.get("Source_Sheet", ""))
            action = row.get("Actions", "")

            # Match file robustly
            match = get_close_matches(src_file, all_uploaded_names, n=1, cutoff=0.9)
            if not match:
                st.warning(f"⚠️ Source file '{row.get('Source_File')}' not found in uploaded files.")
                continue

            matched_file_name = match[0]
            matched_file = next(f for f in uploaded_files if normalize_name(os.path.splitext(f.name)[0]) == matched_file_name)

            try:
                df_src = pd.read_excel(matched_file, sheet_name=src_sheet)
                df_src["Source_File"] = os.path.splitext(matched_file.name)[0]
                df_src["Source_Sheet"] = src_sheet
                df_src["Actions"] = action
                consolidated_data.append(df_src)
            except Exception as e:
                st.error(f"Error reading sheet {src_sheet} from {matched_file.name}: {e}")

        if consolidated_data:
            final_df = pd.concat(consolidated_data, ignore_index=True)

            st.markdown("<div class='sub-header'>📊 Consolidated Results</div>", unsafe_allow_html=True)
            st.dataframe(final_df, use_container_width=True)

            output = BytesIO()
            final_df.to_excel(output, index=False)
            output.seek(0)
            st.download_button("⬇️ Download Consolidated Excel", output, "consolidated_roles.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Footer
st.markdown("<div class='footer'>Excel Role Consolidator v2.1 | Built with ❤️ using Streamlit</div>", unsafe_allow_html=True)
