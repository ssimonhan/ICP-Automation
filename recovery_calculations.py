"""Calculate and build metal-recovery workbooks from ICP final-results workbooks."""

from __future__ import annotations

import math
import re
from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO, NamedTuple

from openpyxl import Workbook, load_workbook
from openpyxl.chart import ScatterChart, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.marker import Marker
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

PRODUCT_STREAM = "Product Stream"
DEPLETED_STREAM = "Depleted Stream"
MEASUREMENT_BASES = (PRODUCT_STREAM, DEPLETED_STREAM)
DEFAULT_RECOVERY_TYPE = "Solid → Liquid"


@dataclass(frozen=True)
class RecoveryTypeConfig:
    label: str
    feed_basis: str
    product_basis: str
    depleted_basis: str
    ui_feed_quantity_label: str
    ui_product_quantity_label: str
    ui_depleted_quantity_label: str
    excel_feed_quantity_label: str
    excel_product_quantity_label: str
    excel_depleted_quantity_label: str
    ui_feed_concentrations_label: str
    excel_feed_concentrations_label: str
    ui_feed_sample_label: str
    ui_product_samples_label: str
    ui_depleted_samples_label: str
    excel_product_stream_label: str
    excel_depleted_stream_label: str

    def stream_quantity_label(self, measurement_basis: str, *, ui: bool) -> str:
        if measurement_basis == PRODUCT_STREAM:
            return self.ui_product_quantity_label if ui else self.excel_product_quantity_label
        if measurement_basis == DEPLETED_STREAM:
            return self.ui_depleted_quantity_label if ui else self.excel_depleted_quantity_label
        raise ValueError(f"Unknown measurement basis: {measurement_basis}")

    def stream_samples_label(self, measurement_basis: str) -> str:
        if measurement_basis == PRODUCT_STREAM:
            return self.ui_product_samples_label
        if measurement_basis == DEPLETED_STREAM:
            return self.ui_depleted_samples_label
        raise ValueError(f"Unknown measurement basis: {measurement_basis}")

    def stream_concentrations_label(self, measurement_basis: str) -> str:
        if measurement_basis == PRODUCT_STREAM:
            return self.excel_product_stream_label
        if measurement_basis == DEPLETED_STREAM:
            return self.excel_depleted_stream_label
        raise ValueError(f"Unknown measurement basis: {measurement_basis}")


RECOVERY_TYPE_CONFIGS = {
    "Solid → Solid": RecoveryTypeConfig(
        label="Solid → Solid",
        feed_basis="mass",
        product_basis="mass",
        depleted_basis="mass",
        ui_feed_quantity_label="Feed solid mass (g)",
        ui_product_quantity_label="Product solid mass (g)",
        ui_depleted_quantity_label="Remaining solid mass (g)",
        excel_feed_quantity_label="Feed mass (g)",
        excel_product_quantity_label="Product mass (g)",
        excel_depleted_quantity_label="Remaining solid mass (g)",
        ui_feed_concentrations_label="Feed solid concentrations (ppm)",
        excel_feed_concentrations_label="Feed solid concentrations (ppm)",
        ui_feed_sample_label="Feed solid sample",
        ui_product_samples_label="Product solid samples",
        ui_depleted_samples_label="Remaining solid samples",
        excel_product_stream_label="Product solid concentrations (ppm)",
        excel_depleted_stream_label="Remaining solid concentrations (ppm)",
    ),
    "Solid → Liquid": RecoveryTypeConfig(
        label="Solid → Liquid",
        feed_basis="mass",
        product_basis="volume",
        depleted_basis="mass",
        ui_feed_quantity_label="Feed mass (g)",
        ui_product_quantity_label="Final volume (mL)",
        ui_depleted_quantity_label="Undissolved solid mass (g)",
        excel_feed_quantity_label="Feed mass (g)",
        excel_product_quantity_label="Final volume (mL)",
        excel_depleted_quantity_label="Undissolved solid mass (g)",
        ui_feed_concentrations_label="Feed concentrations (ppm)",
        excel_feed_concentrations_label="Feed concentrations (ppm)",
        ui_feed_sample_label="Feed sample",
        ui_product_samples_label="Leachate samples",
        ui_depleted_samples_label="Undissolved solid samples",
        excel_product_stream_label="Leachate concentrations (ppm)",
        excel_depleted_stream_label="Undissolved solid concentrations (ppm)",
    ),
    "Liquid → Solid": RecoveryTypeConfig(
        label="Liquid → Solid",
        feed_basis="volume",
        product_basis="mass",
        depleted_basis="volume",
        ui_feed_quantity_label="Feed volume (mL)",
        ui_product_quantity_label="Precipitate mass (g)",
        ui_depleted_quantity_label="Barren volume (mL)",
        excel_feed_quantity_label="Feed volume (mL)",
        excel_product_quantity_label="Precipitate mass (g)",
        excel_depleted_quantity_label="Barren volume (mL)",
        ui_feed_concentrations_label="Feed liquid concentrations (ppm)",
        excel_feed_concentrations_label="Feed liquid concentrations (ppm)",
        ui_feed_sample_label="Feed liquid sample",
        ui_product_samples_label="Precipitate samples",
        ui_depleted_samples_label="Barren liquid samples",
        excel_product_stream_label="Precipitate concentrations (ppm)",
        excel_depleted_stream_label="Barren liquid concentrations (ppm)",
    ),
    "Liquid → Liquid": RecoveryTypeConfig(
        label="Liquid → Liquid",
        feed_basis="volume",
        product_basis="volume",
        depleted_basis="volume",
        ui_feed_quantity_label="Feed volume (mL)",
        ui_product_quantity_label="Organic phase volume (mL)",
        ui_depleted_quantity_label="Raffinate volume (mL)",
        excel_feed_quantity_label="Feed volume (mL)",
        excel_product_quantity_label="Organic phase volume (mL)",
        excel_depleted_quantity_label="Raffinate volume (mL)",
        ui_feed_concentrations_label="Feed liquid concentrations (ppm)",
        excel_feed_concentrations_label="Feed liquid concentrations (ppm)",
        ui_feed_sample_label="Feed liquid sample",
        ui_product_samples_label="Organic phase samples",
        ui_depleted_samples_label="Raffinate samples",
        excel_product_stream_label="Organic phase concentrations (ppm)",
        excel_depleted_stream_label="Raffinate concentrations (ppm)",
    ),
}
RECOVERY_TYPES = tuple(RECOVERY_TYPE_CONFIGS)


def get_recovery_type_config(recovery_type: str) -> RecoveryTypeConfig:
    try:
        return RECOVERY_TYPE_CONFIGS[recovery_type]
    except KeyError as exc:
        raise ValueError(f"Unknown recovery type: {recovery_type}") from exc


def order_selected_elements(elements: list[str], selected: list[str]) -> list[str]:
    """Keep selected elements in the source workbook's canonical order."""

    selected_set = set(selected)
    return [element for element in elements if element in selected_set]


class ICPResults(NamedTuple):
    elements: list[str]
    sample_concentrations: OrderedDict[str, dict[str, float | None]]


def normalize(value: object) -> str:
    return "" if value is None else str(value).strip()


def parse_number(value: object) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return number if math.isfinite(number) else None
    text = str(value).strip()
    if not text or text.startswith("="):
        return None
    match = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", text.replace(",", ""))
    return float(match.group(0)) if match else None


def _nice_axis_step(values: list[float], target_ticks: int = 5) -> float:
    """Choose a readable major tick interval for an axis."""

    if not values:
        return 1.0
    minimum = min(values)
    maximum = max(values)
    span = maximum - minimum
    if span <= 0:
        span = abs(maximum) if maximum else 1.0
    rough_step = span / target_ticks
    if rough_step <= 0:
        return 1.0
    magnitude = 10 ** math.floor(math.log10(rough_step))
    normalized = rough_step / magnitude
    if normalized <= 1:
        return magnitude
    if normalized <= 2:
        return 2 * magnitude
    if normalized <= 5:
        return 5 * magnitude
    return 10 * magnitude


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
    recovery_type: str = DEFAULT_RECOVERY_TYPE,
    measurement_basis: str = PRODUCT_STREAM,
) -> OrderedDict[str, dict[str, float | None]]:
    """Calculate metal recovery from a product stream or by depletion."""

    get_recovery_type_config(recovery_type)
    if measurement_basis not in MEASUREMENT_BASES:
        raise ValueError(f"Unknown measurement basis: {measurement_basis}")
    recoveries: OrderedDict[str, dict[str, float | None]] = OrderedDict()
    for sample, concentrations in results.sample_concentrations.items():
        inputs = sample_inputs.get(sample, {})
        feed_quantity = parse_number(inputs.get("feed_quantity", inputs.get("feed_mass_g")))
        stream_quantity = parse_number(inputs.get("stream_quantity", inputs.get("volume_ml")))
        values: dict[str, float | None] = {}
        for element in results.elements:
            stream_concentration = concentrations.get(element)
            feed_concentration = feed_concentrations.get(element)
            if (
                feed_quantity is None or feed_quantity <= 0
                or stream_quantity is None or stream_quantity < 0
                or stream_concentration is None or feed_concentration is None or feed_concentration == 0
            ):
                values[element] = None
            else:
                stream_remainder = (
                    stream_quantity * stream_concentration / (feed_quantity * feed_concentration) * 100
                )
                values[element] = (
                    stream_remainder
                    if measurement_basis == PRODUCT_STREAM
                    else 100 - stream_remainder
                )
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
    recovery_type: str = DEFAULT_RECOVERY_TYPE,
    measurement_basis: str = PRODUCT_STREAM,
    feed_sample_name: str = "",
) -> None:
    """Create a template-style, standalone recovery workbook."""

    config = get_recovery_type_config(recovery_type)
    if measurement_basis not in MEASUREMENT_BASES:
        raise ValueError(f"Unknown measurement basis: {measurement_basis}")
    stream_quantity_label = config.stream_quantity_label(measurement_basis, ui=False)
    stream_concentrations_label = config.stream_concentrations_label(measurement_basis)
    if measurement_basis == PRODUCT_STREAM:
        calculation_text = (
            "Recovery calculated from Product Stream (%) = product amount × product concentration "
            "/ (feed amount × feed concentration) × 100"
        )
        recovery_title = "Recovery calculated from Product Stream (%)"
    else:
        calculation_text = (
            "Recovery calculated from Depleted Stream (%) = 100 - "
            "(depleted amount × depleted concentration "
            "/ (feed amount × feed concentration) × 100)"
        )
        recovery_title = "Recovery calculated from Depleted Stream (%)"

    workbook = Workbook()
    recovery_ws = workbook.active
    recovery_ws.title = RECOVERY_SHEET
    recovery_ws["A1"] = "Metal recovery calculation"
    recovery_ws["A1"].font = TITLE_FONT
    recovery_ws["A2"] = calculation_text
    feed_sample_name = feed_sample_name.strip()

    # 1. Preserve the feed concentrations in the order supplied by the user.
    recovery_ws["A4"] = config.excel_feed_concentrations_label
    recovery_ws["A4"].font = Font(bold=True)
    recovery_ws.cell(5, 2, "Sample").font = Font(bold=True)
    recovery_ws.cell(6, 2, feed_sample_name)
    feed_elements = list(feed_concentrations)
    for col, element in enumerate(feed_elements, start=ELEMENT_START_COL):
        recovery_ws.cell(5, col, element).font = Font(bold=True)
        recovery_ws.cell(6, col, feed_concentrations[element]).fill = INPUT_FILL

    # 2. Match feed concentrations to ICP elements with XLOOKUP.
    recovery_ws["A8"] = "Feed concentrations matched to ICP elements (ppm)"
    recovery_ws["A8"].font = Font(bold=True)
    recovery_ws.cell(9, 2, "Sample").font = Font(bold=True)
    recovery_ws.cell(10, 2, feed_sample_name)
    feed_last_col = get_column_letter(ELEMENT_START_COL + len(feed_elements) - 1)
    for col, element in enumerate(results.elements, start=ELEMENT_START_COL):
        col_letter = get_column_letter(col)
        recovery_ws.cell(9, col, element).font = Font(bold=True)
        recovery_ws.cell(
            10,
            col,
            f'=_xlfn.XLOOKUP({col_letter}$9,${get_column_letter(ELEMENT_START_COL)}$5:${feed_last_col}$5,${get_column_letter(ELEMENT_START_COL)}$6:${feed_last_col}$6,"")',
        )

    # 3. Copy selected stream concentrations in the template layout.
    stream_title_row = 12
    stream_header_row = stream_title_row + 1
    stream_start_row = stream_header_row + 1
    recovery_ws.cell(stream_title_row, 1, stream_concentrations_label).font = Font(bold=True)
    _write_header(
        recovery_ws,
        stream_header_row,
        [stream_quantity_label, "Sample", parameter_label, *results.elements],
    )
    for row_idx, (sample, concentrations) in enumerate(results.sample_concentrations.items(), start=stream_start_row):
        inputs = sample_inputs.get(sample, {})
        recovery_ws.cell(
            row_idx,
            1,
            parse_number(inputs.get("stream_quantity", inputs.get("volume_ml"))),
        ).fill = INPUT_FILL
        recovery_ws.cell(row_idx, 2, sample)
        recovery_ws.cell(row_idx, 3, parse_number(inputs.get("parameter"))).fill = INPUT_FILL
        for offset, element in enumerate(results.elements, start=ELEMENT_START_COL):
            cell = recovery_ws.cell(row_idx, offset, concentrations[element])
            cell.number_format = "0.0000"

    # 4. Calculate recovery with formulas, keeping feed quantity as an input.
    recovery_title_row = stream_start_row + len(results.sample_concentrations) + 2
    recovery_header_row = recovery_title_row + 1
    recovery_start_row = recovery_header_row + 1
    recovery_ws.cell(recovery_title_row, 1, recovery_title).font = Font(bold=True)
    _write_header(
        recovery_ws,
        recovery_header_row,
        [config.excel_feed_quantity_label, "Sample", parameter_label, *results.elements],
    )
    for offset, (sample, _) in enumerate(results.sample_concentrations.items()):
        stream_row = stream_start_row + offset
        row_idx = recovery_start_row + offset
        inputs = sample_inputs.get(sample, {})
        feed_quantity = parse_number(inputs.get("feed_quantity", inputs.get("feed_mass_g")))
        recovery_ws.cell(row_idx, 1, feed_quantity).fill = INPUT_FILL
        recovery_ws.cell(row_idx, 2, sample)
        recovery_ws.cell(row_idx, 3, parse_number(inputs.get("parameter"))).fill = INPUT_FILL
        for col, _ in enumerate(results.elements, start=ELEMENT_START_COL):
            col_letter = get_column_letter(col)
            stream_ratio = (
                f"${get_column_letter(1)}{stream_row}*{col_letter}{stream_row}"
                f"/(${get_column_letter(1)}{row_idx}*{col_letter}$10)*100"
            )
            formula_value = (
                stream_ratio
                if measurement_basis == PRODUCT_STREAM
                else f"100-({stream_ratio})"
            )
            cell = recovery_ws.cell(
                row_idx,
                col,
                f'=IFERROR({formula_value},"")',
            )
            cell.number_format = "0.00"

    # 5. Summarize selected elements on this same worksheet with XLOOKUP.
    summary_title_row = recovery_start_row + len(results.sample_concentrations) + 2
    summary_header_row = summary_title_row + 1
    summary_start_row = summary_header_row + 1
    recovery_ws.cell(
        summary_title_row,
        1,
        f"Recovery summary calculated from {measurement_basis} (%)",
    ).font = Font(bold=True)
    _write_header(
        recovery_ws,
        summary_header_row,
        [stream_quantity_label, "Sample", parameter_label, *selected_elements],
    )
    recovery_last_col = get_column_letter(ELEMENT_START_COL + len(results.elements) - 1)
    for offset, (sample, _) in enumerate(results.sample_concentrations.items()):
        stream_row = stream_start_row + offset
        recovery_row = recovery_start_row + offset
        row_idx = summary_start_row + offset
        recovery_ws.cell(row_idx, 1, f"=A{stream_row}")
        recovery_ws.cell(row_idx, 2, sample)
        recovery_ws.cell(row_idx, 3, f"=C{stream_row}")
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
        chart.height = 8
        chart.width = 13
        chart.x_axis.axPos = "b"
        chart.y_axis.axPos = "l"
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.x_axis.majorTickMark = "out"
        chart.x_axis.minorTickMark = "out"
        chart.y_axis.majorTickMark = "out"
        chart.y_axis.minorTickMark = "out"
        chart.y_axis.numFmt = "0"
        chart.y_axis.numFmtIsLinked = False
        chart.x_axis.tickLblPos = "nextTo"
        chart.y_axis.tickLblPos = "nextTo"
        chart.x_axis.crosses = "autoZero"
        chart.y_axis.crosses = "autoZero"
        chart.x_axis.crossBetween = "midCat"
        chart.y_axis.crossBetween = "midCat"
        chart.y_axis.majorGridlines = ChartLines(
            spPr=GraphicalProperties(ln=LineProperties(solidFill="D9D9D9"))
        )
        chart.x_axis.majorGridlines = ChartLines(
            spPr=GraphicalProperties(ln=LineProperties(solidFill="D9D9D9"))
        )
        parameter_values = [
            float(parse_number(sample_inputs[sample].get("parameter")))
            for sample in results.sample_concentrations
            if parse_number(sample_inputs[sample].get("parameter")) is not None
        ]
        if parameter_values:
            x_min = min(parameter_values)
            x_max = max(parameter_values)
            x_step = _nice_axis_step(parameter_values)
            if x_min == x_max:
                chart.x_axis.scaling.min = x_min - x_step
                chart.x_axis.scaling.max = x_max + x_step
            else:
                chart.x_axis.scaling.min = math.floor(x_min / x_step) * x_step
                chart.x_axis.scaling.max = math.ceil(x_max / x_step) * x_step
            chart.x_axis.majorUnit = x_step

        recoveries = calculate_recoveries(
            results,
            feed_concentrations,
            sample_inputs,
            recovery_type,
            measurement_basis,
        )
        recovery_values = [
            float(value)
            for sample_values in recoveries.values()
            for element in selected_elements
            if (value := sample_values.get(element)) is not None
        ]
        y_min = min([0.0, *recovery_values])
        y_max = max([0.0, *recovery_values])
        if y_min == y_max:
            y_max = 100.0
        y_step = _nice_axis_step([y_min, y_max])
        y_min = math.floor(y_min / y_step) * y_step
        y_max = math.ceil(y_max / y_step) * y_step
        chart.y_axis.scaling.min = y_min
        chart.y_axis.scaling.max = y_max
        chart.y_axis.majorUnit = y_step

        categories = Reference(
            recovery_ws,
            min_col=3,
            min_row=summary_start_row,
            max_row=summary_start_row + len(results.sample_concentrations) - 1,
        )
        for offset, element in enumerate(selected_elements, start=ELEMENT_START_COL):
            values = Reference(
                recovery_ws,
                min_col=offset,
                min_row=summary_start_row,
                max_row=summary_start_row + len(results.sample_concentrations) - 1,
            )
            series = Series(values, categories, title=element)
            series.marker = Marker("circle")
            series.marker.size = 7
            chart.series.append(series)
        chart.set_categories(categories)
        recovery_ws.add_chart(chart, "A" + str(summary_start_row + len(results.sample_concentrations) + 2))
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"

    recovery_ws.freeze_panes = "D14"
    for col in range(1, recovery_ws.max_column + 1):
        recovery_ws.column_dimensions[get_column_letter(col)].width = 18 if col < ELEMENT_START_COL else 12

    workbook.save(output_path)
