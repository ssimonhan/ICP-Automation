"""Build leaching-recovery workbooks from ICP final-results workbooks."""

from __future__ import annotations

import re
from collections import OrderedDict
from pathlib import Path
from typing import BinaryIO, NamedTuple

from openpyxl import Workbook, load_workbook
from openpyxl.chart import ScatterChart, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.reference import Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from calibration_settings import DEFAULT_CALIBRATION_RANGE, CalibrationRange
from generate_icp_concentrations import read_groups, resolve_source_sheet, selected_index


ICP_SUMMARY_SHEET = "ICP summary"
RECOVERY_SHEET = "Recovery calculation"
ELEMENT_START_COL = 4

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
TITLE_FONT = Font(bold=True, size=14)
HEADER_FONT = Font(bold=True, color="FFFFFF")


class ICPResults(NamedTuple):
    elements: list[str]
    sample_concentrations: OrderedDict[str, dict[str, float | None]]


def normalize(value: object) -> str:
    return "" if value is None else str(value).strip()


def parse_number(value: object) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if text.startswith("="):
        return None
    match = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", text.replace(",", ""))
    return float(match.group(0)) if match else None


def parse_feed_concentrations(text: str, elements: list[str]) -> dict[str, float]:
    """Parse pasted two-row Excel data or Element=ppm feed values."""

    canonical = {element.casefold(): element for element in elements}
    values: dict[str, float] = {}
    lines = [line.strip() for line in (text or "").splitlines() if line.strip()]
    if len(lines) >= 2:
        element_cells = [cell.strip() for cell in re.split(r"\t|,", lines[0])]
        value_cells = [cell.strip() for cell in re.split(r"\t|,", lines[1])]
        for element, raw_value in zip(element_cells, value_cells):
            matched_element = canonical.get(element.casefold())
            number = parse_number(raw_value)
            if matched_element and number is not None:
                values[matched_element] = number
        if values:
            return values

    for line in re.split(r"[\n;,]+", text or ""):
        if not line.strip():
            continue
        if "=" in line:
            element, raw_value = line.split("=", 1)
        elif ":" in line:
            element, raw_value = line.split(":", 1)
        elif "\t" in line:
            element, raw_value = line.split("\t", 1)
        else:
            continue
        matched_element = canonical.get(element.strip().casefold())
        number = parse_number(raw_value)
        if matched_element and number is not None:
            values[matched_element] = number
    return values


def _read_summary_values(workbook) -> ICPResults | None:
    """Read cached numeric values from an existing ICP summary sheet."""

    if ICP_SUMMARY_SHEET not in workbook.sheetnames:
        return None
    ws = workbook[ICP_SUMMARY_SHEET]
    elements = [normalize(ws.cell(1, col).value) for col in range(4, ws.max_column + 1)]
    elements = [element for element in elements if element]
    if not elements:
        return None

    values: OrderedDict[str, dict[str, float | None]] = OrderedDict()
    for row_idx in range(2, ws.max_row + 1):
        sample = normalize(ws.cell(row_idx, 2).value)
        if not sample:
            continue
        row_values = {
            element: parse_number(ws.cell(row_idx, 4 + offset).value)
            for offset, element in enumerate(elements)
        }
        if any(value is not None for value in row_values.values()):
            values[sample] = row_values
    return ICPResults(elements, values) if values else None


def read_selected_concentrations(
    workbook_source: str | Path | BinaryIO,
    calibration_range: CalibrationRange = DEFAULT_CALIBRATION_RANGE,
) -> ICPResults:
    """Read selected sample concentrations, recalculating from the ICP sheet if needed."""

    cached_workbook = load_workbook(workbook_source, data_only=True)
    summary = _read_summary_values(cached_workbook)
    if summary:
        return summary

    if hasattr(workbook_source, "seek"):
        workbook_source.seek(0)
    workbook = load_workbook(workbook_source, data_only=False)
    source_sheet_name = resolve_source_sheet(workbook, None)
    elements, groups = read_groups(workbook[source_sheet_name])
    results: OrderedDict[str, dict[str, float | None]] = OrderedDict()
    for sample, rows in groups.items():
        dilutions = [row["dilution"] for row in rows]
        selected: dict[str, float | None] = {}
        for element_idx, element in enumerate(elements):
            ppb_values = [row["ppb_values"][element_idx] for row in rows]
            choice = selected_index(ppb_values, dilutions, calibration_range)
            if choice is None:
                selected[element] = None
                continue
            ppb = ppb_values[choice]
            dilution = dilutions[choice]
            selected[element] = None if ppb is None or dilution is None else ppb * dilution / 1000
        results[sample] = selected
    if not results:
        raise ValueError("No selected sample concentrations were found in the uploaded workbook.")
    return ICPResults(elements, results)


def calculate_recoveries(
    results: ICPResults,
    feed_concentrations: dict[str, float],
    sample_inputs: dict[str, dict[str, float | None]],
) -> OrderedDict[str, dict[str, float | None]]:
    """Calculate V × c_leachate / (m_feed × c_feed) × 100 for each sample and element."""

    recoveries: OrderedDict[str, dict[str, float | None]] = OrderedDict()
    for sample, concentrations in results.sample_concentrations.items():
        inputs = sample_inputs.get(sample, {})
        volume_ml = parse_number(inputs.get("volume_ml"))
        feed_mass_g = parse_number(inputs.get("feed_mass_g"))
        values: dict[str, float | None] = {}
        for element in results.elements:
            leachate_concentration = concentrations.get(element)
            feed_concentration = feed_concentrations.get(element)
            if (
                volume_ml is None or feed_mass_g is None or feed_mass_g <= 0
                or leachate_concentration is None or feed_concentration is None or feed_concentration == 0
            ):
                values[element] = None
            else:
                values[element] = volume_ml * leachate_concentration / (feed_mass_g * feed_concentration) * 100
        recoveries[sample] = values
    return recoveries


def _write_header(ws, row: int, labels: list[str]) -> None:
    for col, label in enumerate(labels, start=1):
        cell = ws.cell(row, col, label)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT


def build_recovery_workbook(
    output_path: Path,
    results: ICPResults,
    feed_concentrations: dict[str, float],
    sample_inputs: dict[str, dict[str, float | None]],
    parameter_label: str,
    selected_elements: list[str],
) -> None:
    """Create a template-style, standalone leaching-recovery workbook."""

    workbook = Workbook()
    recovery_ws = workbook.active
    recovery_ws.title = RECOVERY_SHEET
    recovery_ws["A1"] = "Leaching recovery"
    recovery_ws["A1"].font = TITLE_FONT
    recovery_ws["A2"] = "Recovery (%) = final volume (mL) × leachate concentration (ppm) / [feed mass (g) × feed concentration (ppm)] × 100"

    # 1. Preserve the feed concentrations in the order supplied by the user.
    recovery_ws["A4"] = "Feed concentrations (ppm)"
    recovery_ws["A4"].font = Font(bold=True)
    feed_elements = list(feed_concentrations)
    for col, element in enumerate(feed_elements, start=ELEMENT_START_COL):
        recovery_ws.cell(5, col, element).font = Font(bold=True)
        recovery_ws.cell(6, col, feed_concentrations[element]).fill = INPUT_FILL

    # 2. Match feed concentrations to the ICP leachate elements with XLOOKUP.
    recovery_ws["A8"] = "Feed concentrations matched to leachate elements (ppm)"
    recovery_ws["A8"].font = Font(bold=True)
    feed_last_col = get_column_letter(ELEMENT_START_COL + len(feed_elements) - 1)
    for col, element in enumerate(results.elements, start=ELEMENT_START_COL):
        col_letter = get_column_letter(col)
        recovery_ws.cell(9, col, element).font = Font(bold=True)
        recovery_ws.cell(
            10,
            col,
            f'=_xlfn.XLOOKUP({col_letter}$9,${get_column_letter(ELEMENT_START_COL)}$5:${feed_last_col}$5,${get_column_letter(ELEMENT_START_COL)}$6:${feed_last_col}$6,"")',
        )

    # 3. Copy selected ICP leachate concentrations in the template layout.
    leachate_title_row = 12
    leachate_header_row = leachate_title_row + 1
    leachate_start_row = leachate_header_row + 1
    recovery_ws.cell(leachate_title_row, 1, "Leachate concentrations (ppm)").font = Font(bold=True)
    _write_header(
        recovery_ws,
        leachate_header_row,
        ["Final volume (mL)", "Sample", parameter_label, *results.elements],
    )
    for row_idx, (sample, concentrations) in enumerate(results.sample_concentrations.items(), start=leachate_start_row):
        inputs = sample_inputs.get(sample, {})
        recovery_ws.cell(row_idx, 1, parse_number(inputs.get("volume_ml"))).fill = INPUT_FILL
        recovery_ws.cell(row_idx, 2, sample)
        recovery_ws.cell(row_idx, 3, parse_number(inputs.get("parameter"))).fill = INPUT_FILL
        for offset, element in enumerate(results.elements, start=ELEMENT_START_COL):
            cell = recovery_ws.cell(row_idx, offset, concentrations[element])
            cell.number_format = "0.0000"

    # 4. Calculate recovery with formulas, leaving feed mass as a distinct input.
    recovery_title_row = leachate_start_row + len(results.sample_concentrations) + 2
    recovery_header_row = recovery_title_row + 1
    recovery_start_row = recovery_header_row + 1
    recovery_ws.cell(recovery_title_row, 1, "Leaching recovery (%)").font = Font(bold=True)
    _write_header(
        recovery_ws,
        recovery_header_row,
        ["Feed mass (g)", "Sample", parameter_label, *results.elements],
    )
    for offset, (sample, _) in enumerate(results.sample_concentrations.items()):
        leachate_row = leachate_start_row + offset
        row_idx = recovery_start_row + offset
        inputs = sample_inputs.get(sample, {})
        recovery_ws.cell(row_idx, 1, parse_number(inputs.get("feed_mass_g"))).fill = INPUT_FILL
        recovery_ws.cell(row_idx, 2, sample)
        recovery_ws.cell(row_idx, 3, parse_number(inputs.get("parameter"))).fill = INPUT_FILL
        for col, _ in enumerate(results.elements, start=ELEMENT_START_COL):
            col_letter = get_column_letter(col)
            cell = recovery_ws.cell(
                row_idx,
                col,
                f'=IFERROR(${get_column_letter(1)}{leachate_row}*{col_letter}{leachate_row}/(${get_column_letter(1)}{row_idx}*{col_letter}$10)*100,"")',
            )
            cell.number_format = "0.00"

    # 5. Summarize selected elements on this same worksheet with XLOOKUP.
    summary_title_row = recovery_start_row + len(results.sample_concentrations) + 2
    summary_header_row = summary_title_row + 1
    summary_start_row = summary_header_row + 1
    recovery_ws.cell(summary_title_row, 1, "Recovery summary (%)").font = Font(bold=True)
    _write_header(
        recovery_ws,
        summary_header_row,
        ["Final volume (mL)", "Sample", parameter_label, *selected_elements],
    )
    recovery_last_col = get_column_letter(ELEMENT_START_COL + len(results.elements) - 1)
    for offset, (sample, _) in enumerate(results.sample_concentrations.items()):
        leachate_row = leachate_start_row + offset
        recovery_row = recovery_start_row + offset
        row_idx = summary_start_row + offset
        recovery_ws.cell(row_idx, 1, f"=A{leachate_row}")
        recovery_ws.cell(row_idx, 2, sample)
        recovery_ws.cell(row_idx, 3, f"=C{leachate_row}")
        for col, _ in enumerate(selected_elements, start=ELEMENT_START_COL):
            col_letter = get_column_letter(col)
            cell = recovery_ws.cell(
                row_idx,
                col,
                f'=_xlfn.XLOOKUP({col_letter}${summary_header_row},${get_column_letter(ELEMENT_START_COL)}${recovery_header_row}:${recovery_last_col}${recovery_header_row},${get_column_letter(ELEMENT_START_COL)}{recovery_row}:${recovery_last_col}{recovery_row},"")',
            )
            cell.number_format = "0.00"

    numeric_parameters = all(parse_number(sample_inputs.get(sample, {}).get("parameter")) is not None for sample in results.sample_concentrations)
    if selected_elements and numeric_parameters:
        chart = ScatterChart()
        chart.title = "Recovery vs. " + parameter_label
        chart.x_axis.title = parameter_label
        chart.y_axis.title = "Recovery (%)"
        chart.y_axis.majorGridlines = ChartLines(
            spPr=GraphicalProperties(ln=LineProperties(solidFill="D9D9D9"))
        )
        chart.x_axis.majorGridlines = ChartLines(
            spPr=GraphicalProperties(ln=LineProperties(solidFill="D9D9D9"))
        )
        categories = Reference(recovery_ws, min_col=3, min_row=summary_start_row, max_row=summary_start_row + len(results.sample_concentrations) - 1)
        for offset, _ in enumerate(selected_elements, start=ELEMENT_START_COL):
            values = Reference(recovery_ws, min_col=offset, min_row=summary_header_row, max_row=summary_start_row + len(results.sample_concentrations) - 1)
            series = Series(values, title_from_data=True)
            chart.series.append(series)
        chart.set_categories(categories)
        recovery_ws.add_chart(chart, "A" + str(summary_start_row + len(results.sample_concentrations) + 2))

    recovery_ws.freeze_panes = "D14"
    for col in range(1, recovery_ws.max_column + 1):
        recovery_ws.column_dimensions[get_column_letter(col)].width = 18 if col < ELEMENT_START_COL else 12

    workbook.save(output_path)
