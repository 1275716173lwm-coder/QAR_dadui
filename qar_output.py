from __future__ import annotations

import os
import uuid
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from qar_analyzer import AnalysisResult, LogRecord, PersonSummary


class OutputError(RuntimeError):
    pass


LOG_HEADERS = [
    "记录类型", "时间", "人员", "文件完整路径", "机型", "指标", "状态", "错误代码",
    "说明", "候选文件数", "有效文件数", "触发文件数", "失败文件数",
]

WORKBOOK_HEADERS = [
    "姓名",
    "低空 300ft AGL 以下杆量超过 15 个单位采样点数",
    "低空 300ft AGL 以下杆量超过 15 个单位采样点占比",
    "低空 50ft AGL 以下舵量超过 15 个单位次数",
    "低空 50ft AGL 以下舵量超过 15 个单位占比",
    "接地交叉角大于 6° 次数",
    "接地交叉角大于 6° 占比",
    "PF 抬头速率大于 3.5°/s 次数",
    "PF 抬头速率大于 3.5°/s 占比",
    "抬头速率大中高高原次数",
    "抬头速率大中高高原占比",
]


def _safe_text(value: object) -> str:
    return str(value).replace("\t", " ").replace("\r", " ").replace("\n", " ")


def create_log_path(output_dir: Path, started_at: datetime) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    stem = f"qar_log_{started_at:%Y%m%d_%H%M%S}"
    candidate = output_dir / f"{stem}.txt"
    counter = 1
    while candidate.exists():
        candidate = output_dir / f"{stem}_{counter}.txt"
        counter += 1
    return candidate


def write_log(output_dir: Path, result: AnalysisResult, extra_records: list[LogRecord] | None = None) -> Path:
    path = create_log_path(output_dir, result.started_at)
    records = [*result.log_records, *(extra_records or [])]
    with path.open("x", encoding="utf-8", newline="") as handle:
        handle.write("\t".join(LOG_HEADERS) + "\n")
        for record in records:
            row = [
                record.record_type,
                record.timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                record.person,
                record.file_path,
                record.aircraft_type,
                record.metric,
                record.status,
                record.error_code,
                record.message,
                record.candidate,
                record.valid,
                record.triggered,
                record.failed,
            ]
            handle.write("\t".join(_safe_text(value) for value in row) + "\n")
    return path


def _write_summary_row(sheet, row: int, summary: PersonSummary) -> None:
    values = [
        summary.person,
        summary.stick_exceed_count,
        summary.stick_exceed_ratio,
        None,
        None,
        summary.cross_angle_count,
        summary.cross_angle_ratio,
        summary.pitch_rate_count,
        summary.pitch_rate_ratio,
        summary.high_altitude_count,
        summary.high_altitude_ratio,
    ]
    for column, value in enumerate(values, start=1):
        sheet.cell(row=row, column=column, value=value)
    for column in (3, 7, 9, 11):
        sheet.cell(row=row, column=column).number_format = "0.00%"


def write_workbook(template_path: Path, output_dir: Path, summaries: list[PersonSummary]) -> Path:
    if not template_path.exists() or not template_path.is_file():
        raise OutputError(f"结果模板不存在：{template_path}")
    output_dir.mkdir(parents=True, exist_ok=True)
    temporary_path = output_dir / f".qar_{uuid.uuid4().hex}.tmp.xlsx"
    target_path = output_dir / "qar.xlsx"

    try:
        workbook = load_workbook(template_path)
        if not workbook.worksheets:
            raise OutputError("结果模板不包含工作表")
        sheet = workbook.worksheets[0]
        if sheet.max_column < 11:
            raise OutputError(f"结果模板至少需要11列，当前只有{sheet.max_column}列")
        if sheet.max_row > 1:
            sheet.delete_rows(2, sheet.max_row - 1)
        for column, header in enumerate(WORKBOOK_HEADERS, start=1):
            sheet.cell(row=1, column=column, value=header)
        for row, summary in enumerate(summaries, start=2):
            _write_summary_row(sheet, row, summary)
        workbook.save(temporary_path)
        workbook.close()
        try:
            os.replace(temporary_path, target_path)
        except PermissionError as exc:
            raise OutputError("无法替换result/qar.xlsx；请关闭Excel中已打开的该文件后重试") from exc
        except OSError as exc:
            raise OutputError(f"无法保存结果表：{exc}") from exc
        return target_path
    except OutputError:
        raise
    except Exception as exc:
        raise OutputError(f"生成结果表失败：{exc}") from exc
    finally:
        try:
            temporary_path.unlink(missing_ok=True)
        except OSError:
            pass
