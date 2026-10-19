"""해석 완료보고서 Excel 파일에서 표지·개요에 넣을 값을 읽는 모듈.

'해석 완료 보고서' 시트를 우선 읽고, 값이 없으면 '해석 결과 보고서' 시트의
대응 셀을 대신 읽는다. 셀 좌표는 보고서 양식에 맞춰 고정되어 있다.
"""

from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook


def _format_value(value):

    """셀 값을 'YYYY-MM-DD' 문자열이나 다듬은 문자열로 변환한다."""

    if value is None:
        return None

    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")

    text = str(value).strip()

    return text or None


def _first_value(workbook, candidates):

    """후보 (시트, 셀) 목록을 순서대로 읽어 처음 나온 값을 반환한다."""

    for sheet_name, coordinate in candidates:
        if sheet_name not in workbook.sheetnames:
            continue

        value = _format_value(
            workbook[sheet_name][coordinate].value
        )

        if value:
            return value

    return None


def read_report_data(file_path):

    """보고서 Excel을 열어 PPT에 채울 9개 항목을 dict로 반환한다."""

    path = Path(file_path)

    workbook = load_workbook(
        path,
        data_only=True,
        read_only=True,
        keep_vba=path.suffix.lower() == ".xlsm"
    )

    try:
        first_sheet = "해석 완료 보고서"
        summary_sheet = "해석 결과 보고서"

        return {
            "title": _first_value(
                workbook,
                [(first_sheet, "B5")]
            ),
            "analysis_number": _first_value(
                workbook,
                [
                    (first_sheet, "F2"),
                    (summary_sheet, "H2")
                ]
            ),
            "contact": _first_value(
                workbook,
                [
                    (first_sheet, "T2"),
                    (summary_sheet, "X2")
                ]
            ),
            "end_date": _first_value(
                workbook,
                [(first_sheet, "AD2")]
            ),
            "agenda": _first_value(
                workbook,
                [
                    (first_sheet, "J23"),
                    (summary_sheet, "N7")
                ]
            ),
            "background": _first_value(
                workbook,
                [
                    (first_sheet, "J30"),
                    (summary_sheet, "N9")
                ]
            ),
            "objective": _first_value(
                workbook,
                [
                    (first_sheet, "J38"),
                    (summary_sheet, "N14")
                ]
            ),
            "conclusion": _first_value(
                workbook,
                [
                    (first_sheet, "J46"),
                    (summary_sheet, "N19")
                ]
            )
        }
    finally:
        workbook.close()