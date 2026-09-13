from __future__ import annotations

import csv
import tempfile
import threading
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from docx import Document
from openpyxl import Workbook, load_workbook

from qar_analyzer import (
    AircraftType,
    CONFIGS,
    calculate_crosswind,
    analyze_flight,
    analyze_selected_folder,
    parse_flight_number,
    parse_flight_time,
)
from qar_output import (
    LOG_HEADERS,
    WORKBOOK_HEADERS,
    OutputError,
    create_log_path,
    write_docx_report,
    write_log,
    write_workbook,
)


def set_value(row: list[str], index: int, value: object) -> None:
    row[index] = str(value)


def synthetic_cfm(
    path: Path,
    aircraft_heading: float = 359,
    runway_heading: float = 1,
    pitch_rate: float = 3.5,
    baro_altitude: float = 8000,
    wind_rows: dict[int, tuple[float, float]] | None = None,
) -> None:
    config = CONFIGS[AircraftType.CFM]
    width = max(
        config.altitude_col,
        config.stick_cols[1],
        config.aircraft_heading_col,
        config.runway_heading_col,
        config.pitch_rate_cols[1],
        config.baro_altitude_col,
        config.wind_speed_col,
        config.wind_direction_col,
    ) + 1
    header = [f"column_{index}" for index in range(width)]
    rows: list[list[str]] = []
    for altitude in (2100, 300, -1, -1):
        row = [""] * width
        set_value(row, config.altitude_col, altitude)
        rows.append(row)

    for index in range(config.takeoff_group_1[0] + 1, config.takeoff_group_1[0] + 3):
        set_value(rows[0], index, 3)
    for index in range(config.takeoff_group_2[0], config.takeoff_group_2[0] + 2):
        set_value(rows[0], index, 3)
    set_value(rows[0], config.baro_altitude_col, baro_altitude)
    for row in rows:
        for index in range(config.pitch_rate_cols[0], config.pitch_rate_cols[1] + 1):
            set_value(row, index, pitch_rate)

    set_value(rows[1], config.stick_cols[0], 15)
    set_value(rows[1], config.stick_cols[0] + 1, -15)
    set_value(rows[1], config.stick_cols[0] + 2, 16)
    set_value(rows[2], config.aircraft_heading_col, aircraft_heading)
    set_value(rows[2], config.runway_heading_col, runway_heading)

    if wind_rows is None:
        wind_rows = {2: (10, 90)}
    for row_index, (speed, direction) in wind_rows.items():
        set_value(rows[row_index], config.wind_speed_col, speed)
        set_value(rows[row_index], config.wind_direction_col, direction)

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
            self.assertEqual(result.metrics["stick"].value["beff"], 1)
            self.assertFalse(result.metrics["cross_angle"].value)
            self.assertFalse(result.metrics["pitch_rate"].value)
            self.assertFalse(result.metrics["high_altitude"].applicable)
            self.assertFalse(result.crosswind_metric.applicable)

    def test_cross_angle_strictly_greater_than_six(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            six = folder / "B-A_20260101010101___S1_E1_YZDCFM_CA0001.csv"
            seven = folder / "B-B_20260101010102___S1_E1_YZDCFM_CA0002.csv"
            synthetic_cfm(six, 6, 0)
            synthetic_cfm(seven, 7, 0)
            self.assertFalse(
                analyze_flight("P", six, AircraftType.CFM, threading.Event())
                .metrics["cross_angle"].value
            )
            self.assertTrue(
                analyze_flight("P", seven, AircraftType.CFM, threading.Event())
                .metrics["cross_angle"].value
            )

    def test_crosswind_direction_examples(self) -> None:
        self.assertAlmostEqual(calculate_crosswind(10, 90, 0), 10)
        self.assertAlmostEqual(calculate_crosswind(10, 270, 0), -10)
        self.assertEqual(calculate_crosswind(10, 0, 0), 0)
        self.assertEqual(calculate_crosswind(10, 180, 0), 0)
        self.assertAlmostEqual(calculate_crosswind(10, -90, 0), -10)
        self.assertAlmostEqual(calculate_crosswind(10, 450, 0), 10)

    def test_nearest_wind_pair_tie_prefers_upward_row(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "B-TEST_20260101010101___S1_E1_YZDCFM_CA0001.csv"
            synthetic_cfm(
                path,
                aircraft_heading=10,
                runway_heading=0,
                wind_rows={1: (10, 90), 3: (20, 270)},
            )
            result = analyze_flight("测试", path, AircraftType.CFM, threading.Event())
            self.assertTrue(result.crosswind_metric.success)
            self.assertEqual(result.wind_source_row, 4)
            self.assertEqual(result.wind_row_offset, -1)
            self.assertAlmostEqual(result.crosswind or 0, 10)

    def test_pitch_event_records_peak_and_high_altitude(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "B-TEST_20260101010101___S1_E1_YZDCFM_CA0001.csv"
            synthetic_cfm(path, pitch_rate=3.6, baro_altitude=8001)
            result = analyze_flight("测试", path, AircraftType.CFM, threading.Event())
            self.assertTrue(result.metrics["pitch_rate"].value)
            self.assertAlmostEqual(result.max_pitch_rate or 0, 3.6)
            self.assertTrue(result.metrics["high_altitude"].value)

    def test_metadata_parsing_uses_first_legal_values(self) -> None:
        name = "B-X_20261301010101_20260102030405_YZDCFM_CA1234_CA9999.csv"
        self.assertEqual(parse_flight_time(name), datetime(2026, 1, 2, 3, 4, 5))
        self.assertEqual(parse_flight_number(name, AircraftType.CFM), "CA1234")

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

    def test_timestamped_outputs_docx_and_log(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            person = root / "input" / "甲"
            person.mkdir(parents=True)
            synthetic_cfm(
                person / "B-A_20260101010101___S1_E1_YZDCFM_CA0001.csv",
                aircraft_heading=10,
                runway_heading=0,
                pitch_rate=3.6,
                baro_altitude=8001,
            )
            result = analyze_selected_folder(root / "input")

            template = root / "template.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            for column in range(1, 12):
                sheet.cell(1, column, f"H{column}")
            workbook.save(template)
            workbook.close()

            output_dir = root / "result"
            fixed_save_time = datetime(2026, 1, 2, 3, 4, 5)
            with patch("qar_output._local_now", return_value=fixed_save_time):
                output = write_workbook(template, output_dir, result.summaries)
                second_output = write_workbook(template, output_dir, result.summaries)
            self.assertEqual(output.name, "qar_20260102_030405.xlsx")
            self.assertEqual(second_output.name, "qar_20260102_030405_1.xlsx")

            saved = load_workbook(output, data_only=False)
            sheet = saved.active
            self.assertEqual(
                [sheet.cell(1, column).value for column in range(1, 12)],
                WORKBOOK_HEADERS,
            )
            self.assertEqual(sheet.cell(2, 1).value, "甲")
            self.assertIsNone(sheet.cell(2, 4).value)
            self.assertIsNone(sheet.cell(2, 5).value)
            self.assertEqual(sheet.cell(2, 3).number_format, "0.00%")
            self.assertTrue(sheet.cell(1, 1).alignment.wrap_text)
            self.assertGreaterEqual(sheet.row_dimensions[1].height or 0, 96)
            saved.close()

            report = write_docx_report(output_dir, result)
            second_report = write_docx_report(output_dir, result)
            self.assertTrue(report.name.startswith("qar_report_"))
            self.assertTrue(second_report.stem.endswith("_1"))
            document = Document(report)
            paragraph_text = "\n".join(paragraph.text for paragraph in document.paragraphs)
            table_text = "\n".join(
                cell.text
                for table in document.tables
                for row in table.rows
                for cell in row.cells
            )
            self.assertIn("甲", paragraph_text)
            self.assertIn("接地交叉角事件", paragraph_text)
            self.assertIn("抬头速率事件", paragraph_text)
            self.assertIn("暂未分析", table_text)
            self.assertNotIn("侧风分量", paragraph_text.split("抬头速率事件", 1)[1])

            log_path = write_log(output_dir, result)
            content = log_path.read_text(encoding="utf-8")
            self.assertEqual(content.splitlines()[0].split("\t"), LOG_HEADERS)
            self.assertTrue(log_path.name.startswith("qar_log_"))
            second_log = create_log_path(output_dir, result.started_at)
            self.assertNotEqual(log_path, second_log)
            self.assertEqual(list(output_dir.glob(".*.tmp.xlsx")), [])
            self.assertEqual(list(output_dir.glob(".*.tmp.docx")), [])

    def test_cancelled_result_cannot_create_docx(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            result = analyze_selected_folder(temp, threading.Event())
            result.cancelled = True
            with self.assertRaises(OutputError):
                write_docx_report(Path(temp) / "result", result)
            self.assertEqual(list((Path(temp) / "result").glob("*.docx")), [])


if __name__ == "__main__":
    unittest.main()
