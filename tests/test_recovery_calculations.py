import tempfile
import unittest
from collections import OrderedDict
from pathlib import Path

from openpyxl import load_workbook

from recovery_calculations import (
    DEFAULT_RECOVERY_TYPE,
    DEPLETED_STREAM,
    ICPResults,
    PRODUCT_STREAM,
    RECOVERY_TYPES,
    build_recovery_workbook,
    calculate_recoveries,
    get_recovery_type_config,
    order_selected_elements,
)


class RecoveryCalculationTests(unittest.TestCase):
    def setUp(self):
        self.results = ICPResults(
            ["Cu"],
            OrderedDict([("Stream 1", {"Cu": 5.0})]),
        )
        self.feed_concentrations = {"Cu": 10.0}
        self.sample_inputs = {
            "Stream 1": {
                "feed_quantity": 2.0,
                "stream_quantity": 1.0,
                "parameter": 2.0,
            }
        }

    def test_all_recovery_types_support_both_measurement_bases(self):
        for recovery_type in RECOVERY_TYPES:
            with self.subTest(recovery_type=recovery_type):
                config = get_recovery_type_config(recovery_type)
                self.assertIn(config.feed_basis, ("mass", "volume"))
                for basis in (PRODUCT_STREAM, DEPLETED_STREAM):
                    with self.subTest(measurement_basis=basis):
                        result = calculate_recoveries(
                            self.results,
                            self.feed_concentrations,
                            self.sample_inputs,
                            recovery_type,
                            basis,
                        )
                        expected = 25.0 if basis == PRODUCT_STREAM else 75.0
                        self.assertEqual(result["Stream 1"]["Cu"], expected)

    def test_legacy_leaching_input_keys_preserve_product_recovery(self):
        result = calculate_recoveries(
            self.results,
            self.feed_concentrations,
            {"Stream 1": {"feed_mass_g": 2.0, "volume_ml": 1.0}},
        )
        self.assertEqual(result["Stream 1"]["Cu"], 25.0)

    def test_selected_elements_follow_source_workbook_order(self):
        result = order_selected_elements(
            ["La", "Pr", "Nd", "Sm"],
            ["Sm", "Pr"],
        )
        self.assertEqual(result, ["Pr", "Sm"])

    def test_depleted_workbook_uses_depleted_formula_and_type_labels(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "recovery.xlsx"
            build_recovery_workbook(
                output_path,
                self.results,
                self.feed_concentrations,
                self.sample_inputs,
                "pH",
                ["Cu"],
                "Liquid → Solid",
                DEPLETED_STREAM,
                "Feed sample 1",
            )
            sheet = load_workbook(output_path, data_only=False)["Recovery calculation"]

        self.assertEqual(sheet["B5"].value, "Sample")
        self.assertEqual(sheet["B6"].value, "Feed sample 1")
        self.assertEqual(sheet["B9"].value, "Sample")
        self.assertEqual(sheet["B10"].value, "Feed sample 1")
        self.assertEqual(sheet["A13"].value, "Barren volume (mL)")
        self.assertEqual(sheet["A18"].value, "Feed volume (mL)")
        self.assertEqual(sheet["A17"].value, "Recovery calculated from Depleted Stream (%)")
        self.assertEqual(sheet["A22"].value, "Recovery summary calculated from Depleted Stream (%)")
        self.assertIn("100-(", sheet["D19"].value)
        self.assertIn("XLOOKUP", sheet["D24"].value)

    def test_default_workbook_keeps_leaching_labels_and_product_formula(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "leaching.xlsx"
            build_recovery_workbook(
                output_path,
                self.results,
                self.feed_concentrations,
                self.sample_inputs,
                "pH",
                ["Cu"],
                DEFAULT_RECOVERY_TYPE,
                PRODUCT_STREAM,
                "Feed sample 2",
            )
            sheet = load_workbook(output_path, data_only=False)["Recovery calculation"]

        self.assertEqual(sheet["B5"].value, "Sample")
        self.assertEqual(sheet["B6"].value, "Feed sample 2")
        self.assertEqual(sheet["B9"].value, "Sample")
        self.assertEqual(sheet["B10"].value, "Feed sample 2")
        self.assertEqual(sheet["A13"].value, "Final volume (mL)")
        self.assertEqual(sheet["A18"].value, "Feed mass (g)")
        self.assertEqual(sheet["A17"].value, "Recovery calculated from Product Stream (%)")
        self.assertEqual(sheet["A22"].value, "Recovery summary calculated from Product Stream (%)")
        self.assertIn("IFERROR", sheet["D19"].value)
        self.assertNotIn("100-(", sheet["D19"].value)


if __name__ == "__main__":
    unittest.main()
