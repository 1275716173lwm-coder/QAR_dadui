from __future__ import annotations

import csv
import math
import re
import threading
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any, Callable, Iterable


class AircraftType(str, Enum):
    CFM = "CFM"
    LEAP = "LEAP"
    PW = "PW"


class AnalysisCancelled(Exception):
    """Raised internally when the user requests a safe cancellation."""


@dataclass(frozen=True)
class AircraftConfig:
    marker: str
    altitude_col: int
    stick_cols: tuple[int, int]
    aircraft_heading_col: int
    runway_heading_col: int
    takeoff_group_1: tuple[int, int]
    takeoff_group_2: tuple[int, int]
    takeoff_min_1: int
    takeoff_min_2: int
    pitch_rate_cols: tuple[int, int]
    baro_altitude_col: int


@dataclass
class MetricResult:
    success: bool
    value: Any = None
    error_code: str = ""
    error_message: str = ""
    applicable: bool = True


@dataclass
class FlightResult:
    person: str
    path: Path
    aircraft_type: AircraftType
    flight_time: datetime | None
    metrics: dict[str, MetricResult]
    warnings: list[str] = field(default_factory=list)


@dataclass
class MetricStats:
    candidate: int = 0
    valid: int = 0
    triggered: int = 0
    failed: int = 0


@dataclass
class PersonSummary:
    person: str
    stick_exceed_count: int
    stick_exceed_ratio: float | None
    cross_angle_count: int
    cross_angle_ratio: float | None
    pitch_rate_count: int
    pitch_rate_ratio: float | None
    high_altitude_count: int
    high_altitude_ratio: float | None
    stats: dict[str, MetricStats]


@dataclass
class LogRecord:
    record_type: str
    timestamp: datetime
    person: str = ""
    file_path: str = ""
    aircraft_type: str = ""
    metric: str = ""
    status: str = ""
    error_code: str = ""
    message: str = ""
    candidate: str = ""
    valid: str = ""
    triggered: str = ""
    failed: str = ""


@dataclass
class ProgressEvent:
    stage: str
    person: str = ""
    file_name: str = ""
    completed: int = 0
    total: int = 0
    message: str = ""


@dataclass
class AnalysisResult:
    selected_folder: Path
    started_at: datetime
    ended_at: datetime
    summaries: list[PersonSummary]
    flights: list[FlightResult]
    log_records: list[LogRecord]
    cancelled: bool = False


def _column_number(label: str) -> int:
    value = 0
    for char in label.upper():
        value = value * 26 + ord(char) - ord("A") + 1
    return value - 1


CONFIGS: dict[AircraftType, AircraftConfig] = {
    AircraftType.CFM: AircraftConfig(
        marker="YZDCFM",
        altitude_col=_column_number("D"),
        stick_cols=(_column_number("AM"), _column_number("BB")),
        aircraft_heading_col=_column_number("BG"),
        runway_heading_col=_column_number("BH"),
        takeoff_group_1=(_column_number("D"), _column_number("J")),
        takeoff_group_2=(_column_number("K"), _column_number("Q")),
        takeoff_min_1=2,
        takeoff_min_2=2,
        pitch_rate_cols=(_column_number("BL"), _column_number("BO")),
        baro_altitude_col=_column_number("BK"),
    ),
    AircraftType.LEAP: AircraftConfig(
        marker="YZDLEAP",
        altitude_col=_column_number("D"),
        stick_cols=(_column_number("AE"), _column_number("AT")),
        aircraft_heading_col=_column_number("AW"),
        runway_heading_col=_column_number("AX"),
        takeoff_group_1=(_column_number("D"), _column_number("H")),
        takeoff_group_2=(_column_number("I"), _column_number("M")),
        takeoff_min_1=1,
        takeoff_min_2=1,
        pitch_rate_cols=(_column_number("BB"), _column_number("BE")),
        baro_altitude_col=_column_number("BA"),
    ),
    AircraftType.PW: AircraftConfig(
        marker="YZDPW",
        altitude_col=_column_number("D"),
        stick_cols=(_column_number("AE"), _column_number("AT")),
        aircraft_heading_col=_column_number("AY"),
        runway_heading_col=_column_number("AZ"),
        takeoff_group_1=(_column_number("D"), _column_number("H")),
        takeoff_group_2=(_column_number("I"), _column_number("M")),
        takeoff_min_1=1,
        takeoff_min_2=2,
        pitch_rate_cols=(_column_number("BD"), _column_number("BG")),
        baro_altitude_col=_column_number("BC"),
    ),
}

METRIC_LABELS = {
    "stick": "300ft以下杆量",
    "cross_angle": "接地交叉角",
    "pitch_rate": "PF抬头速率",
    "high_altitude": "高高原抬头速率",
}

_TIME_PATTERN = re.compile(r"_(\d{14})_")


def parse_flight_time(name: str) -> datetime | None:
    match = _TIME_PATTERN.search(name)
    if not match:
        return None
    try:
        return datetime.strptime(match.group(1), "%Y%m%d%H%M%S")
    except ValueError:
        return None


def identify_aircraft(name: str) -> AircraftType | None:
    upper_name = name.upper()
    matches = [kind for kind, config in CONFIGS.items() if config.marker in upper_name]
    return matches[0] if len(matches) == 1 else None


def discover_people(selected_folder: Path) -> list[tuple[Path, list[tuple[Path, AircraftType]]]]:
    if not selected_folder.exists() or not selected_folder.is_dir():
        raise ValueError(f"所选数据目录不存在或不是文件夹：{selected_folder}")

    people: list[tuple[Path, list[tuple[Path, AircraftType]]]] = []
    person_dirs = sorted(
        (path for path in selected_folder.iterdir() if path.is_dir()),
        key=lambda path: (path.name.casefold(), path.name),
    )
    for person_dir in person_dirs:
        flights: list[tuple[Path, AircraftType]] = []
        for path in person_dir.iterdir():
            if not path.is_file() or path.suffix.lower() != ".csv":
                continue
            aircraft_type = identify_aircraft(path.name)
            if aircraft_type is not None:
                flights.append((path, aircraft_type))
        flights.sort(
            key=lambda item: (
                parse_flight_time(item[0].name) is None,
                parse_flight_time(item[0].name) or datetime.max,
                item[0].name.casefold(),
                item[0].name,
            )
        )
        people.append((person_dir, flights))
    return people


def _finite_number(value: str | None) -> float | None:
    if value is None:
        return None
    try:
        number = float(value.strip())
    except (AttributeError, TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def _cell(row: list[str], index: int) -> str | None:
    return row[index] if 0 <= index < len(row) else None


def _numbers(row: list[str], start: int, end: int) -> list[float | None]:
    return [_finite_number(_cell(row, index)) for index in range(start, end + 1)]


def _select_encoding(path: Path, cancel_event: threading.Event) -> str:
    for encoding in ("utf-8-sig", "utf-8", "gb18030", "cp936"):
        try:
            with path.open("r", encoding=encoding, newline="") as handle:
                while handle.read(1024 * 1024):
                    if cancel_event.is_set():
                        raise AnalysisCancelled
            return encoding
        except UnicodeDecodeError:
            continue
    raise UnicodeError("文件无法使用支持的编码读取")


def _failed_metrics(code: str, message: str) -> dict[str, MetricResult]:
    return {
        key: MetricResult(False, error_code=code, error_message=message)
        for key in METRIC_LABELS
    }


def analyze_flight(
    person: str,
    path: Path,
    aircraft_type: AircraftType,
    cancel_event: threading.Event,
) -> FlightResult:
    config = CONFIGS[aircraft_type]
    warnings: list[str] = []
    expected_width: int | None = None
    short_rows = 0
    long_rows = 0

    seen_above_2000 = False
    approach_rows: list[tuple[float | None, list[float | None], float | None, float | None]] = []
    touchdown_found = False

    takeoff_row: int | None = None
    previous_pitch_rates: list[float | None] | None = None
    pitch_window: list[float | None] = []
    pitch_rows_remaining = 0
    takeoff_baro: float | None = None

    try:
        encoding = _select_encoding(path, cancel_event)
        with path.open("r", encoding=encoding, newline="") as handle:
            reader = csv.reader(handle)
            try:
                header = next(reader)
            except StopIteration:
                raise ValueError("CSV为空文件")
            try:
                next(reader)  # units row
            except StopIteration:
                raise ValueError("CSV缺少单位行和数据行")

            expected_width = len(header)
            if expected_width == 0:
                raise ValueError("CSV表头为空")

            for data_index, row in enumerate(reader):
                if data_index % 250 == 0 and cancel_event.is_set():
                    raise AnalysisCancelled

                if len(row) < expected_width:
                    short_rows += 1
                elif len(row) > expected_width:
                    long_rows += 1

                radio_height = _finite_number(_cell(row, config.altitude_col))
                if not seen_above_2000 and radio_height is not None and radio_height > 2000:
                    seen_above_2000 = True

                if seen_above_2000 and not touchdown_found:
                    approach_rows.append(
                        (
                            radio_height,
                            _numbers(row, *config.stick_cols),
                            _finite_number(_cell(row, config.aircraft_heading_col)),
                            _finite_number(_cell(row, config.runway_heading_col)),
                        )
                    )
                    if radio_height is not None and radio_height < 0:
                        touchdown_found = True

                current_pitch_rates = _numbers(row, *config.pitch_rate_cols)
                if takeoff_row is None:
                    group_1 = _numbers(row, *config.takeoff_group_1)
                    group_2 = _numbers(row, *config.takeoff_group_2)
                    count_1 = sum(value is not None and value > 2 for value in group_1)
                    count_2 = sum(value is not None and value > 2 for value in group_2)
                    if count_1 >= config.takeoff_min_1 and count_2 >= config.takeoff_min_2:
                        takeoff_row = data_index
                        if previous_pitch_rates is not None:
                            pitch_window.extend(previous_pitch_rates)
                        pitch_window.extend(current_pitch_rates)
                        pitch_rows_remaining = 2
                        takeoff_baro = _finite_number(_cell(row, config.baro_altitude_col))
                    else:
                        previous_pitch_rates = current_pitch_rates
                elif pitch_rows_remaining > 0:
                    pitch_window.extend(current_pitch_rates)
                    pitch_rows_remaining -= 1

                if touchdown_found and takeoff_row is not None and pitch_rows_remaining == 0:
                    break
    except AnalysisCancelled:
        raise
    except Exception as exc:
        message = f"读取或解析CSV失败：{exc}"
        return FlightResult(
            person=person,
            path=path,
            aircraft_type=aircraft_type,
            flight_time=parse_flight_time(path.name),
            metrics=_failed_metrics("CSV_READ_ERROR", message),
            warnings=warnings,
        )

    if short_rows:
        warnings.append(f"发现{short_rows}行字段数少于表头，缺失字段按无效值处理")
    if long_rows:
        warnings.append(f"发现{long_rows}行字段数多于表头，多余字段已忽略")

    metrics: dict[str, MetricResult] = {}
    if not seen_above_2000:
        landing_error = ("NO_ABOVE_2000", "D列未找到大于2000的有效高度")
    elif not touchdown_found:
        landing_error = ("NO_TOUCHDOWN", "从首次大于2000后未找到D列首次负值")
    else:
        landing_error = None

    if landing_error:
        metrics["stick"] = MetricResult(False, error_code=landing_error[0], error_message=landing_error[1])
        metrics["cross_angle"] = MetricResult(False, error_code=landing_error[0], error_message=landing_error[1])
    else:
        nearest_index = min(
            (index for index, item in enumerate(approach_rows) if item[0] is not None),
            key=lambda index: abs(float(approach_rows[index][0]) - 300),
        )
        stick_values = [
            value
            for _, row_values, _, _ in approach_rows[nearest_index:]
            for value in row_values
            if value is not None
        ]
        if stick_values:
            exceed = sum(abs(value) > 15 for value in stick_values)
            metrics["stick"] = MetricResult(True, {"beff": exceed, "ball": len(stick_values)})
        else:
            metrics["stick"] = MetricResult(
                False,
                error_code="NO_STICK_VALUES",
                error_message="300ft至接地区域没有有效杆量采样点",
            )

        aircraft_heading = approach_rows[-1][2]
        runway_heading = approach_rows[-1][3]
        if runway_heading is None:
            runway_heading = next(
                (item[3] for item in reversed(approach_rows[:-1]) if item[3] is not None),
                None,
            )
        if aircraft_heading is None:
            metrics["cross_angle"] = MetricResult(
                False, error_code="NO_AIRCRAFT_HEADING", error_message="接地行飞机磁航向无有效值"
            )
        elif runway_heading is None:
            metrics["cross_angle"] = MetricResult(
                False, error_code="NO_RUNWAY_HEADING", error_message="接地行及其上方没有有效跑道磁航向"
            )
        else:
            angle_difference = abs(((aircraft_heading - runway_heading + 180) % 360) - 180)
            metrics["cross_angle"] = MetricResult(True, angle_difference > 6)

    if takeoff_row is None:
        metrics["pitch_rate"] = MetricResult(
            False, error_code="NO_TAKEOFF_ROW", error_message="未找到满足条件的首次离地行"
        )
    else:
        valid_pitch_rates = [value for value in pitch_window if value is not None]
        if not valid_pitch_rates:
            metrics["pitch_rate"] = MetricResult(
                False, error_code="NO_PITCH_RATE_VALUES", error_message="离地行前1至后2行没有有效抬头速率"
            )
        else:
            metrics["pitch_rate"] = MetricResult(True, max(valid_pitch_rates) > 3.5)

    pitch_metric = metrics["pitch_rate"]
    if pitch_metric.success and pitch_metric.value is True:
        if takeoff_baro is None:
            metrics["high_altitude"] = MetricResult(
                False, error_code="NO_BARO_ALTITUDE", error_message="离地行气压高度无有效值"
            )
        else:
            metrics["high_altitude"] = MetricResult(True, takeoff_baro > 8000)
    else:
        metrics["high_altitude"] = MetricResult(
            True,
            value=None,
            error_code="NOT_APPLICABLE",
            error_message="抬头速率未超限或抬头速率指标失败",
            applicable=False,
        )

    return FlightResult(
        person=person,
        path=path,
        aircraft_type=aircraft_type,
        flight_time=parse_flight_time(path.name),
        metrics=metrics,
        warnings=warnings,
    )


def _summarize_person(person: str, flights: Iterable[FlightResult]) -> PersonSummary:
    flight_list = list(flights)
    stats = {key: MetricStats() for key in METRIC_LABELS}
    stick_exceed = 0
    stick_total = 0

    for flight in flight_list:
        for key in ("stick", "cross_angle", "pitch_rate"):
            metric = flight.metrics[key]
            item = stats[key]
            item.candidate += 1
            if metric.success:
                item.valid += 1
                if key == "stick":
                    if metric.value["beff"] > 0:
                        item.triggered += 1
                elif metric.value is True:
                    item.triggered += 1
            else:
                item.failed += 1

        high_metric = flight.metrics["high_altitude"]
        if high_metric.applicable:
            stats["high_altitude"].candidate += 1
            if high_metric.success:
                stats["high_altitude"].valid += 1
                if high_metric.value is True:
                    stats["high_altitude"].triggered += 1
            else:
                stats["high_altitude"].failed += 1

        stick_metric = flight.metrics["stick"]
        if stick_metric.success:
            stick_exceed += int(stick_metric.value["beff"])
            stick_total += int(stick_metric.value["ball"])

    cross = stats["cross_angle"]
    pitch = stats["pitch_rate"]
    high = stats["high_altitude"]
    return PersonSummary(
        person=person,
        stick_exceed_count=stick_exceed,
        stick_exceed_ratio=(stick_exceed / stick_total) if stick_total else None,
        cross_angle_count=cross.triggered,
        cross_angle_ratio=(cross.triggered / cross.valid) if cross.valid else None,
        pitch_rate_count=pitch.triggered,
        pitch_rate_ratio=(pitch.triggered / pitch.valid) if pitch.valid else None,
        high_altitude_count=high.triggered,
        high_altitude_ratio=(high.triggered / high.valid) if high.valid else None,
        stats=stats,
    )


def _metric_log_record(flight: FlightResult, key: str, metric: MetricResult) -> LogRecord:
    if not metric.applicable:
        status = "NOT_APPLICABLE"
    else:
        status = "SUCCESS" if metric.success else "FAILED"
    value = ""
    if metric.value is not None:
        value = str(metric.value)
    message = metric.error_message
    if value:
        message = f"值={value}" + (f"；{message}" if message else "")
    return LogRecord(
        record_type="DETAIL",
        timestamp=datetime.now(),
        person=flight.person,
        file_path=str(flight.path),
        aircraft_type=flight.aircraft_type.value,
        metric=METRIC_LABELS[key],
        status=status,
        error_code=metric.error_code,
        message=message,
    )


def analyze_selected_folder(
    path: str | Path,
    cancel_event: threading.Event | None = None,
    progress_callback: Callable[[ProgressEvent], None] | None = None,
) -> AnalysisResult:
    selected_folder = Path(path).resolve()
    cancel_event = cancel_event or threading.Event()
    callback = progress_callback or (lambda event: None)
    started_at = datetime.now()
    log_records = [
        LogRecord(
            record_type="RUN",
            timestamp=started_at,
            status="STARTED",
            message=f"所选数据目录={selected_folder}",
        )
    ]
    people = discover_people(selected_folder)
    total_files = sum(len(items) for _, items in people)
    completed_files = 0
    all_flights: list[FlightResult] = []
    summaries: list[PersonSummary] = []
    cancelled = False

    callback(ProgressEvent("started", total=total_files, message=f"发现{len(people)}名人员、{total_files}个数据文件"))
    for person_dir, files in people:
        if cancel_event.is_set():
            cancelled = True
            break
        callback(
            ProgressEvent(
                "person", person=person_dir.name, completed=completed_files, total=total_files,
                message=f"开始分析人员：{person_dir.name}",
            )
        )
        person_flights: list[FlightResult] = []
        for file_path, aircraft_type in files:
            if cancel_event.is_set():
                cancelled = True
                break
            callback(
                ProgressEvent(
                    "file", person=person_dir.name, file_name=file_path.name,
                    completed=completed_files, total=total_files, message=f"正在分析：{file_path.name}",
                )
            )
            try:
                flight = analyze_flight(person_dir.name, file_path, aircraft_type, cancel_event)
            except AnalysisCancelled:
                cancelled = True
                break
            person_flights.append(flight)
            all_flights.append(flight)
            for key, metric in flight.metrics.items():
                log_records.append(_metric_log_record(flight, key, metric))
            for warning in flight.warnings:
                log_records.append(
                    LogRecord(
                        record_type="WARNING", timestamp=datetime.now(), person=flight.person,
                        file_path=str(flight.path), aircraft_type=flight.aircraft_type.value,
                        status="WARNING", error_code="ROW_WIDTH_MISMATCH", message=warning,
                    )
                )
            completed_files += 1
            callback(
                ProgressEvent(
                    "progress", person=person_dir.name, file_name=file_path.name,
                    completed=completed_files, total=total_files, message=f"已完成：{file_path.name}",
                )
            )
        summaries.append(_summarize_person(person_dir.name, person_flights))
        if cancelled:
            break

    for summary in summaries:
        for key, item in summary.stats.items():
            log_records.append(
                LogRecord(
                    record_type="SUMMARY", timestamp=datetime.now(), person=summary.person,
                    metric=METRIC_LABELS[key], status="SUMMARY", candidate=str(item.candidate),
                    valid=str(item.valid), triggered=str(item.triggered), failed=str(item.failed),
                )
            )

    ended_at = datetime.now()
    failed_count = sum(1 for record in log_records if record.status == "FAILED")
    if cancelled:
        status = "CANCELLED"
        message = "用户取消分析，未生成或替换结果表"
    elif failed_count:
        status = "COMPLETED_WITH_ERRORS"
        message = f"分析完成，共记录{failed_count}项失败"
    else:
        status = "COMPLETED"
        message = "分析完成，未发现错误"
    log_records.append(LogRecord(record_type="RUN", timestamp=ended_at, status=status, message=message))
    callback(
        ProgressEvent(
            "cancelled" if cancelled else "analyzed",
            completed=completed_files,
            total=total_files,
            message=message,
        )
    )
    return AnalysisResult(
        selected_folder=selected_folder,
        started_at=started_at,
        ended_at=ended_at,
        summaries=summaries,
        flights=all_flights,
        log_records=log_records,
        cancelled=cancelled,
    )
