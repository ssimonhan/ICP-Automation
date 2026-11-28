import streamlit as st
import hashlib
import os
import re
import tempfile
from collections import OrderedDict
from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from streamlit_sortables import sort_items

from calibration_settings import CalibrationRange, HighlightThresholds
from process_icp_data import normalize_sample_initials, process_icp_file
from generate_icp_concentrations import build_concentration_workbook
from recovery_calculations import (
    DEFAULT_RECOVERY_TYPE,
    DEPLETED_STREAM,
    ICPResults,
    MEASUREMENT_BASES,
    PRODUCT_STREAM,
    RECOVERY_TYPES,
    build_recovery_workbook,
    calculate_recoveries,
    get_recovery_type_config,
    parse_feed_concentrations,
    parse_number,
    order_selected_elements,
    read_selected_concentrations,
)
from excel_backend import excel_backend


MONTH_NUMBERS = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}
MONTH_PATTERN = "|".join(MONTH_NUMBERS)
SORTABLE_COMPACT_STYLE = """
.sortable-component.vertical .sortable-container {
    min-width: 150px;
    padding: 4px;
}
.sortable-container-header {
    padding: 2px 4px;
    font-weight: 600;
}
.sortable-container-body {
    display: flex;
    flex-wrap: wrap;
    align-content: flex-start;
    gap: 4px;
    min-height: 42px;
    padding: 4px;
}
.sortable-component.vertical .sortable-item {
    display: inline-flex;
    width: fit-content;
    max-width: 100%;
    height: auto;
    align-items: center;
    margin: 0;
    padding: 3px 8px;
    border-radius: 14px;
    font-size: 0.85rem;
    line-height: 1.2;
}
"""


def date_from_filename(filename: str) -> str | None:
    """Return a YYYYMMDD date found in an uploaded filename, if present."""

    stem = Path(filename).stem
    numeric_patterns = (
        r"(?<!\d)(?P<year>20\d{2})(?P<month>0[1-9]|1[0-2])(?P<day>0[1-9]|[12]\d|3[01])(?!\d)",
        r"(?<!\d)(?P<year>20\d{2})[_. -](?P<month>0[1-9]|1[0-2])[_. -](?P<day>0[1-9]|[12]\d|3[01])(?!\d)",
    )
    for pattern in numeric_patterns:
        match = re.search(pattern, stem)
        if match:
            try:
                return date(**{field: int(value) for field, value in match.groupdict().items()}).strftime("%Y%m%d")
            except ValueError:
                continue

    written_date = re.search(
        rf"(?<!\d)(?P<day>0?[1-9]|[12]\d|3[01])[_. -]*(?P<month>{MONTH_PATTERN})[_. -]*(?P<year>20\d{{2}})(?!\d)",
        stem,
        flags=re.IGNORECASE,
    )
    if written_date:
        groups = written_date.groupdict()
        try:
            return date(
                int(groups["year"]),
                MONTH_NUMBERS[groups["month"].lower()],
                int(groups["day"]),
            ).strftime("%Y%m%d")
        except ValueError:
            pass
    return None


def final_download_filename(sample_initials: str, uploaded_filename: str) -> str:
    """Return the final workbook download name with an upload or run date."""

    output_date = date_from_filename(uploaded_filename) or date.today().strftime("%Y%m%d")
    return f"{sample_initials}_ICP_results_{output_date}.xlsx"


def cleaned_download_filename(sample_initials: str, uploaded_filename: str) -> str:
    """Return the cleaned workbook download name with an upload or run date."""

    output_date = date_from_filename(uploaded_filename) or date.today().strftime("%Y%m%d")
    return f"{sample_initials}_ICP_cleaned_{output_date}.xlsx"


def recovery_download_filename(uploaded_filename: str) -> str:
    """Return a dated recovery-workbook name based on the uploaded ICP results."""

    initials_match = re.match(r"([A-Za-z]+)_ICP_results", Path(uploaded_filename).name)
    initials = initials_match.group(1).upper() if initials_match else "ICP"
    output_date = date_from_filename(uploaded_filename) or date.today().strftime("%Y%m%d")
    return f"{initials}_recovery_{output_date}.xlsx"


st.set_page_config(page_title="ICP Automation", layout="wide")

analysis_tab, recovery_tab = st.tabs(["ICP Analysis", "Metal Recovery"])

# Keep the existing ICP workflow together in its own tab.
analysis_tab.__enter__()

st.title("🧪 ICP Automation Tool")

if excel_backend() == "cloud":
    st.info(
        "ℹ️ Running in cloud mode. "
        "Excel automatic formatting is approximated for compatibility."
    )

uploaded_file = st.file_uploader("Upload ICP export", type=["csv", "xlsx"])
initials = st.text_input("Sample initials", value="SH")

with st.expander("Advanced settings"):
    col1, col2 = st.columns(2)
    with col1:
        st.caption("Customize ppb ranges used for ICP highlighting and final concentration selection.")
        orange_min = st.number_input(
            "Orange range minimum ppb",
            min_value=0.0,
            value=1.0,
            step=0.1,
            help="Values from this number up to the green minimum are lower-confidence and highlighted orange.",
        )
        green_min = st.number_input(
            "Green range minimum ppb",
            min_value=0.0,
            value=10.0,
            step=1.0,
            help="Values from this number through the green maximum are treated as accurate.",
        )
        green_max = st.number_input(
            "Green range maximum ppb",
            min_value=0.0,
            value=400.0,
            step=10.0,
            help="Values greater than this are outside the accurate range.",
        )
    with col2:
        st.caption("Customize internal-standard and blank-sample warning thresholds.")
        internal_light_orange = st.number_input(
            "Internal standard tolerance before orange (%)",
            min_value=0.0,
            value=20.0,
            step=1.0,
        )
        internal_orange = st.number_input(
            "Internal standard orange threshold (%)",
            min_value=0.0,
            value=40.0,
            step=1.0,
            help="Values above this threshold are highlighted dark orange.",
        )
        blank_light_orange = st.number_input(
            "Blank sample tolerance before orange (ppb)",
            min_value=0.0,
            value=0.2,
            step=0.1,
        )
        blank_orange = st.number_input(
            "Blank sample orange threshold (ppb)",
            min_value=0.0,
            value=1.0,
            step=0.1,
            help="Values above this threshold are highlighted dark orange.",
        )

try:
    sample_initials = normalize_sample_initials(initials)
    calibration_range = CalibrationRange(
        orange_min=orange_min,
        green_min=green_min,
        green_max=green_max,
    )
    highlight_thresholds = HighlightThresholds(
        internal_light_orange=internal_light_orange,
        internal_orange=internal_orange,
        blank_light_orange=blank_light_orange,
        blank_orange=blank_orange,
    )
except ValueError as exc:
    st.error(f"Settings error: {exc}")
    st.stop()

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

        result_code = process_icp_file(
            input_path=input_path,
            output_path=mid_path,
            sheet_name=None,
            sample_initials=sample_initials,
            calibration_range=calibration_range,
            highlight_thresholds=highlight_thresholds,
            source_name=uploaded_file.name,
        )

        os.remove(input_path)

        if result_code != 0:
            if mid_path.exists():
                os.remove(mid_path)
            st.session_state.mid_path = None
            st.session_state.pop("cleaned_file", None)
            st.session_state.pop("final_file", None)
            st.error("ICP worksheet layout validation failed. Please check the messages above and the raw export format.")
            st.stop()

    # Save path
    st.session_state.mid_path = str(mid_path)
    st.session_state.source_filename = uploaded_file.name

    # Store file bytes persistently
    with open(mid_path, "rb") as f:
        st.session_state.cleaned_file = f.read()

    st.success("✅ ICP sheet generated. Now edit dilution and sample below.")
    
    st.session_state.pop("final_file", None)

# ALWAYS show download button if file exists
if "cleaned_file" in st.session_state:
    st.download_button(
        "⬇️ Download Cleaned ICP File",
        data=st.session_state.cleaned_file,
        file_name=cleaned_download_filename(
            sample_initials,
            st.session_state.get("source_filename", ""),
        )
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
                source_sheet_name=icp_sheet_name,
                calibration_range=calibration_range,
            )

        with open(output_path, "rb") as f:
            st.session_state.final_file = f.read()

        st.success("✅ Final concentrations generated!")

# Persistent download button
if "final_file" in st.session_state:
    st.download_button(
        "⬇️ Download Final Excel",
        data=st.session_state.final_file,
        file_name=final_download_filename(
            sample_initials,
            st.session_state.get("source_filename", ""),
        )
    )

analysis_tab.__exit__(None, None, None)

with recovery_tab:
    st.title("Metal Recovery Calculator")
    st.caption(
        "Upload a final ICP results workbook, configure the feed and stream types, "
        "then calculate metal recovery by element."
    )
    st.info(
        "If you edit selected concentrations in Excel, save the ICP results workbook "
        "before uploading it here so the linked summary values are updated."
    )
    use_current_result = st.checkbox(
        "Use the final ICP results generated in ICP Analysis",
        value=False,
        disabled="final_file" not in st.session_state,
        help="Available after you generate final concentrations in the ICP Analysis tab.",
        key="use_current_icp_result",
    )
    recovery_upload = None if use_current_result else st.file_uploader(
        "Upload ICP results workbook",
        type=["xlsx"],
        key="recovery_upload",
        help="Use a file named like [initials]_ICP_results_YYYYMMDD.xlsx.",
    )
    recovery_file_bytes = st.session_state.get("final_file") if use_current_result else (
        recovery_upload.getvalue() if recovery_upload else None
    )
    recovery_source_filename = (
        final_download_filename(sample_initials, st.session_state.get("source_filename", ""))
        if use_current_result else recovery_upload.name if recovery_upload else ""
    )

    if recovery_file_bytes:
        try:
            icp_results = read_selected_concentrations(BytesIO(recovery_file_bytes))
        except Exception as exc:
            st.error(f"Could not read selected ICP concentrations: {exc}")
        else:
            samples = list(icp_results.sample_concentrations)
            st.success(f"Loaded {len(samples)} samples and {len(icp_results.elements)} elements.")

            with st.container(border=True):
                st.subheader("Recovery settings")
                recovery_type_column, measurement_basis_column = st.columns(2)
                with recovery_type_column:
                    recovery_type = st.selectbox(
                        "Recovery Type",
                        RECOVERY_TYPES,
                        index=RECOVERY_TYPES.index(DEFAULT_RECOVERY_TYPE),
                        key="recovery_type",
                        help=(
                            "Choose the physical phases of the feed and the product stream. "
                            "The first phase is the feed and the second is the product "
                            "(for example, Solid → Liquid for leaching)."
                        ),
                    )
                with measurement_basis_column:
                    measurement_basis = st.radio(
                        "Measurement Basis",
                        MEASUREMENT_BASES,
                        horizontal=True,
                        key="recovery_measurement_basis",
                        help=(
                            "Product Stream calculates recovered metal directly from the product. "
                            "Depleted Stream uses the remaining metal in the depleted stream and "
                            "calculates 100% minus that fraction."
                        ),
                    )
                st.caption(
                    "Recovery (%) = recovered metal / feed metal × 100"
                    if measurement_basis == PRODUCT_STREAM
                    else "Recovery (%) = 100 - (remaining metal / feed metal × 100)"
                )
            recovery_config = get_recovery_type_config(recovery_type)

            with st.container(border=True):
                st.subheader("Feed concentrations")
                feed_mode = st.radio(
                    "Feed concentration source",
                    ["Select a feed sample from ICP results", "Paste feed concentrations"],
                    horizontal=True,
                    key="feed_mode",
                )
                feed_sample = None
                if feed_mode == "Select a feed sample from ICP results":
                    feed_sample = st.selectbox(
                        recovery_config.ui_feed_sample_label,
                        samples,
                        key="feed_sample",
                    )
                    feed_sample_name = feed_sample
                    feed_concentrations = dict(icp_results.sample_concentrations[feed_sample])
                else:
                    feed_sample_name = st.text_input(
                        "Feed sample name",
                        key="recovery_feed_sample_name",
                        help="This name is recorded in the recovery workbook.",
                    ).strip()

                if feed_mode == "Paste feed concentrations":
                    feed_text = st.text_area(
                        recovery_config.ui_feed_concentrations_label,
                        placeholder="Li\tFe\tNd\n19.7\t352411.6\t120.4",
                        help="Paste two rows from Excel: element names first, concentrations second. Element=ppm is also supported.",
                        key="feed_concentrations",
                    )
                    feed_concentrations = parse_feed_concentrations(feed_text, icp_results.elements)

            default_samples = [sample for sample in samples if sample != feed_sample]
            sample_order_key = "recovery_sample_order_" + hashlib.sha256(
                repr((samples, feed_sample)).encode("utf-8")
            ).hexdigest()[:12]
            selected_samples_header = (
                f"{recovery_config.stream_samples_label(measurement_basis)} "
                "for recovery calculation"
            )
            previous_sort = st.session_state.get(sample_order_key)
            if previous_sort:
                available_samples = [
                    sample for sample in previous_sort[0]["items"] if sample in samples
                ]
                selected_samples = [
                    sample for sample in previous_sort[1]["items"] if sample in samples
                ]
                assigned_samples = set(available_samples + selected_samples)
                for sample in samples:
                    if sample not in assigned_samples:
                        if sample in default_samples:
                            selected_samples.append(sample)
                        else:
                            available_samples.append(sample)
            else:
                selected_samples = list(default_samples)
                available_samples = [
                    sample for sample in samples if sample not in selected_samples
                ]
            selected_header_hash = hashlib.sha256(
                selected_samples_header.encode("utf-8")
            ).hexdigest()[:8]

            with st.container(border=True):
                st.subheader("Samples for recovery")
                st.caption(
                    f"{feed_sample_name} is excluded from recovery samples by default."
                    if feed_sample
                    else "All ICP samples are selected by default when feed concentrations are pasted."
                )
                sample_column, parameter_column = st.columns(2)
                with sample_column:
                    st.caption(
                        f"Drag samples into or out of the selected list. "
                        f"Reorder {recovery_config.stream_samples_label(measurement_basis).lower()} "
                        "by dragging them within that list."
                    )
                    sortable_samples = sort_items(
                        [
                            {"header": "Available samples", "items": available_samples},
                            {
                                "header": selected_samples_header,
                                "items": selected_samples,
                            },
                        ],
                        multi_containers=True,
                        direction="vertical",
                        custom_style=SORTABLE_COMPACT_STYLE,
                        key=f"recovery_samples_{sample_order_key}_{selected_header_hash}",
                    )
                    st.session_state[sample_order_key] = sortable_samples
                    selected_samples = sortable_samples[1]["items"]
                with parameter_column:
                    parameter_label = st.text_input(
                        "Experiment parameter name",
                        value="pH",
                        key="recovery_parameter_label",
                        help="For example: pH, temperature (°C), acid concentration, or leaching time (h).",
                    ).strip() or "Experiment parameter"
            input_rows = [
                {
                    "Sample": sample,
                    recovery_config.ui_feed_quantity_label: None,
                    recovery_config.stream_quantity_label(measurement_basis, ui=True): None,
                    "Experiment parameter": None,
                }
                for sample in selected_samples
            ]
            feed_quantity_label = recovery_config.ui_feed_quantity_label
            stream_quantity_label = recovery_config.stream_quantity_label(measurement_basis, ui=True)
            stream_basis = (
                recovery_config.product_basis
                if measurement_basis == PRODUCT_STREAM
                else recovery_config.depleted_basis
            )
            st.subheader("Stream quantities")
            st.caption(
                f"Enter {feed_quantity_label.lower()}, {stream_quantity_label.lower()}, "
                f"and {parameter_label} for each selected stream."
            )
            input_df = st.data_editor(
                pd.DataFrame(input_rows),
                width="stretch",
                hide_index=True,
                disabled=["Sample"],
                column_config={
                    feed_quantity_label: st.column_config.NumberColumn(
                        min_value=0.0,
                        step=0.0001 if recovery_config.feed_basis == "mass" else 0.01,
                        format="%.4f" if recovery_config.feed_basis == "mass" else "%.2f",
                    ),
                    stream_quantity_label: st.column_config.NumberColumn(
                        min_value=0.0,
                        step=0.0001 if stream_basis == "mass" else 0.01,
                        format="%.4f" if stream_basis == "mass" else "%.2f",
                    ),
                    "Experiment parameter": st.column_config.NumberColumn(
                        label=parameter_label,
                        step=0.01,
                        format="%.2f",
                    ),
                },
                key=(
                    f"recovery_input_table_{RECOVERY_TYPES.index(recovery_type)}_"
                    f"{MEASUREMENT_BASES.index(measurement_basis)}"
                ),
            )

            st.subheader("Elements for recovery summary and chart")
            selected_element_options = st.pills(
                "Click elements to include",
                icp_results.elements,
                selection_mode="multi",
                default=None,
                key="recovery_elements",
                help="Selected elements are used in the recovery summary and chart.",
            )
            selected_elements = order_selected_elements(
                icp_results.elements,
                selected_element_options or [],
            )

            if selected_samples and selected_elements:
                sample_inputs = {
                    row["Sample"]: {
                        "feed_quantity": parse_number(row[feed_quantity_label]),
                        "stream_quantity": parse_number(row[stream_quantity_label]),
                        "parameter": parse_number(row["Experiment parameter"]),
                    }
                    for _, row in input_df.iterrows()
                }
                selected_results = ICPResults(
                    icp_results.elements,
                    OrderedDict(
                        (sample, icp_results.sample_concentrations[sample])
                        for sample in selected_samples
                    ),
                )
                recoveries = calculate_recoveries(
                    selected_results,
                    feed_concentrations,
                    sample_inputs,
                    recovery_type,
                    measurement_basis,
                )
                preview_rows = []
                for sample in selected_samples:
                    preview_rows.append(
                        {
                            "Sample": sample,
                            parameter_label: sample_inputs[sample]["parameter"],
                            **{element: recoveries[sample][element] for element in selected_elements},
                        }
                    )
                st.subheader(f"Recovery calculated from {measurement_basis} (%)")
                st.dataframe(pd.DataFrame(preview_rows), width="stretch", hide_index=True)

                if st.button("Generate Recovery Workbook", key="generate_recovery"):
                    missing_inputs = [
                        sample for sample, values in sample_inputs.items()
                        if values["feed_quantity"] is None or values["stream_quantity"] is None
                        or values["feed_quantity"] <= 0 or values["stream_quantity"] < 0
                    ]
                    missing_feed = [
                        element for element in selected_elements
                        if feed_concentrations.get(element) in (None, 0)
                    ]
                    if missing_inputs:
                        st.error(
                            f"Enter a positive {feed_quantity_label.lower()} and a non-negative "
                            f"{stream_quantity_label.lower()} for: " + ", ".join(missing_inputs)
                        )
                    elif missing_feed:
                        st.error("Provide non-zero feed concentrations for: " + ", ".join(missing_feed))
                    else:
                        with st.spinner("Building recovery workbook..."):
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as output_file:
                                output_path = Path(output_file.name)
                            try:
                                build_recovery_workbook(
                                    output_path,
                                    selected_results,
                                    feed_concentrations,
                                    sample_inputs,
                                    parameter_label,
                                    selected_elements,
                                    recovery_type,
                                    measurement_basis,
                                    feed_sample_name,
                                )
                                st.session_state.recovery_file = output_path.read_bytes()
                                st.session_state.recovery_source_filename = recovery_source_filename
                            finally:
                                if output_path.exists():
                                    os.remove(output_path)

                        st.success("Recovery workbook generated.")
            else:
                st.info("Select at least one stream sample and one element to continue.")

    if "recovery_file" in st.session_state:
        st.download_button(
            "Download Recovery Workbook",
            data=st.session_state.recovery_file,
            file_name=recovery_download_filename(st.session_state.get("recovery_source_filename", "")),
            key="download_recovery",
        )
