from __future__ import annotations

import csv
import tempfile
import threading
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from qar_analyzer import AircraftType, CONFIGS, analyze_flight, analyze_selected_folder
from qar_output import WORKBOOK_HEADERS, create_log_path, write_log, write_workbook


def set_value(row: list[str], index: int, value: object) -> None:
    row[index] = str(value)


def synthetic_cfm(path: Path, aircraft_heading: float = 359, runway_heading: float = 1) -> None:
    config = CONFIGS[AircraftType.CFM]
    width = max(
        config.altitude_col,
        config.stick_cols[1],
        config.aircraft_heading_col,
        config.runway_heading_col,
        config.pitch_rate_cols[1],
        config.baro_altitude_col,
    ) + 1
    header = [f"column_{index}" for index in range(width)]
    rows: list[list[str]] = []
    for altitude in (2100, 300, -1, -1):
        row = [""] * width
        set_value(row, config.altitude_col, altitude)
        rows.append(row)

    # The first row is the takeoff row for CFM.
    for index in range(config.takeoff_group_1[0] + 1, config.takeoff_group_1[0] + 3):
        set_value(rows[0], index, 3)
    for index in range(config.takeoff_group_2[0], config.takeoff_group_2[0] + 2):
        set_value(rows[0], index, 3)
    set_value(rows[0], config.baro_altitude_col, 8000)
    for row in rows:
        for index in range(config.pitch_rate_cols[0], config.pitch_rate_cols[1] + 1):
            set_value(row, index, 3.5)

    set_value(rows[1], config.stick_cols[0], 15)
    set_value(rows[1], config.stick_cols[0] + 1, -15)
    set_value(rows[1], config.stick_cols[0] + 2, 16)
    set_value(rows[2], config.aircraft_heading_col, aircraft_heading)
    set_value(rows[2], config.runway_heading_col, runway_heading)

    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerow(["units"] + [""] * (width - 1))
        writer.writerows(rows)


class QarAnalyzerTests(unittest.TestCase):
    def test_strict_thresholds_and_circular_heading(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "B-TEST_20260101010101___S1_E1_YZDCFM_CA0001.csv"
            synthetic_cfm(path)
            result = analyze_flight("测试", path, AircraftType.CFM, threading.Event())
            self.assertTrue(result.metrics["stick"].success)
            self.assertEqual(result.metrics["stick"].value["beff"], 1)
            self.assertTrue(result.metrics["cross_angle"].success)
            self.assertFalse(result.metrics["cross_angle"].value)
            self.assertTrue(result.metrics["pitch_rate"].success)
            self.assertFalse(result.metrics["pitch_rate"].value)
            self.assertFalse(result.metrics["high_altitude"].applicable)

    def test_cross_angle_strictly_greater_than_six(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            six = folder / "B-A_20260101010101___S1_E1_YZDCFM_CA0001.csv"
            seven = folder / "B-B_20260101010102___S1_E1_YZDCFM_CA0002.csv"
            synthetic_cfm(six, 6, 0)
            synthetic_cfm(seven, 7, 0)
            self.assertFalse(analyze_flight("P", six, AircraftType.CFM, threading.Event()).metrics["cross_angle"].value)
            self.assertTrue(analyze_flight("P", seven, AircraftType.CFM, threading.Event()).metrics["cross_angle"].value)

    def test_cancel_does_not_analyze_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            person = Path(temp) / "人员"
            person.mkdir()
            synthetic_cfm(person / "B-A_20260101010101___S1_E1_YZDCFM_CA0001.csv")
            event = threading.Event()
            event.set()
            result = analyze_selected_folder(temp, event)
            self.assertTrue(result.cancelled)
            self.assertEqual(result.flights, [])

    def test_workbook_and_log_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            person = root / "input" / "甲"
            person.mkdir(parents=True)
            synthetic_cfm(person / "B-A_20260101010101___S1_E1_YZDCFM_CA0001.csv")
            result = analyze_selected_folder(root / "input")

            template = root / "template.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            for column in range(1, 12):
                sheet.cell(1, column, f"H{column}")
            workbook.save(template)
            workbook.close()

            output_dir = root / "result"
            output = write_workbook(template, output_dir, result.summaries)
            saved = load_workbook(output, data_only=False)
            sheet = saved.active
            self.assertEqual([sheet.cell(1, column).value for column in range(1, 12)], WORKBOOK_HEADERS)
            self.assertEqual(sheet.cell(2, 1).value, "甲")
            self.assertIsNone(sheet.cell(2, 4).value)
            self.assertIsNone(sheet.cell(2, 5).value)
            self.assertEqual(sheet.cell(2, 3).number_format, "0.00%")
            saved.close()

            log_path = write_log(output_dir, result)
            content = log_path.read_text(encoding="utf-8")
            self.assertIn("\t", content.splitlines()[0])
            self.assertTrue(log_path.name.startswith("qar_log_"))
            second = create_log_path(output_dir, result.started_at)
            self.assertNotEqual(log_path, second)


if __name__ == "__main__":
    unittest.main()
