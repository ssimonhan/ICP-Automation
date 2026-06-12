import streamlit as st
import tempfile
import os
from pathlib import Path
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook

from process_icp_data import process_icp_file
from generate_icp_concentrations import build_concentration_workbook

st.set_page_config(page_title="ICP Automation", layout="wide")

st.title("🧪 ICP Automation Tool (Editable Workflow)")

uploaded_file = st.file_uploader("Upload ICP export", type=["csv", "xlsx"])
initials = st.text_input("Sample initials", value="SH")

# Store intermediate file path
if "mid_path" not in st.session_state:
    st.session_state.mid_path = None

# =========================
# STEP 1
# =========================

if uploaded_file and st.button("Step 1: Generate ICP Sheet"):

    with st.spinner("Processing raw file..."):

        with tempfile.NamedTemporaryFile(delete=False, suffix=uploaded_file.name) as tmp_in:
            tmp_in.write(uploaded_file.getvalue())
            input_path = Path(tmp_in.name)

        tmp_mid = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        mid_path = Path(tmp_mid.name)
        tmp_mid.close()

        process_icp_file(
            input_path=input_path,
            output_path=mid_path,
            sheet_name=None,
            sample_initials=initials
        )

        os.remove(input_path)

    # Save path
    st.session_state.mid_path = str(mid_path)

    # ✅ Store file bytes persistently
    with open(mid_path, "rb") as f:
        st.session_state.cleaned_file = f.read()

    st.success("✅ ICP sheet generated. Now edit dilution and sample below.")
    
    st.session_state.pop("final_file", None)

# ✅ ALWAYS show download button if file exists
if "cleaned_file" in st.session_state:
    st.download_button(
        "⬇️ Download Cleaned ICP File",
        data=st.session_state.cleaned_file,
        file_name=f"{initials}_ICP_cleaned.xlsx"
    )

# =========================
# STEP 2 — EDIT TABLE
# =========================

if st.session_state.mid_path:

    wb = load_workbook(st.session_state.mid_path)
    
    # Find ICP sheet
    icp_sheet_name = None
    for name in wb.sheetnames:
        if name.upper().endswith(" ICP"):
            icp_sheet_name = name
            break

    if icp_sheet_name is None:
        st.error("No ICP sheet found")
        st.stop()

    ws = wb[icp_sheet_name]

    # Convert to DataFrame
    data = []

    for row in range(3, ws.max_row + 1):
        sample_name = ws.cell(row, 3).value

        if sample_name:
            data.append({
                "Row": row,
                "Dilution": ws.cell(row, 1).value,
                "Sample": ws.cell(row, 2).value or "",
                "Sample Name": sample_name
            })

    df = pd.DataFrame(data)

    st.subheader("✏️ Edit Dilution and Sample")

    edited_df = st.data_editor(
        df,
        width='stretch',
        disabled=["Row", "Sample Name"],  # prevent breaking keys        
        column_config={
            "Dilution": st.column_config.NumberColumn("Dilution", step=1)
        }
    )

    # =========================
    # STEP 3 — RUN FINAL
    # =========================

    if st.button("Step 2: Generate Final Concentrations"):

        if edited_df["Dilution"].isnull().any():
            st.error("❌ Please fill in all dilution values.")
            st.stop()

        with st.spinner("Running full analysis..."):

            # Write edits back to workbook
            wb = load_workbook(st.session_state.mid_path)
            ws = wb[icp_sheet_name]

            for _, row in edited_df.iterrows():
                excel_row = int(row["Row"])
                ws.cell(excel_row, 1).value = row["Dilution"]
                ws.cell(excel_row, 2).value = row["Sample"]

            wb.save(st.session_state.mid_path)

            # Output file
            tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            output_path = Path(tmp_out.name)
            tmp_out.close()

            build_concentration_workbook(
                source_path=Path(st.session_state.mid_path),
                output_path=output_path,
                source_sheet_name=icp_sheet_name
            )

        with open(output_path, "rb") as f:
            st.session_state.final_file = f.read()

        st.success("✅ Final concentrations generated!")

# ✅ Persistent download button
if "final_file" in st.session_state:
    st.download_button(
        "⬇️ Download Final Excel",
        data=st.session_state.final_file,
        file_name="icp_results.xlsx"
    )