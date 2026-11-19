"""중기부 조간 기사 목록을 원본 매크로 양식의 TXT로 변환."""
from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
from io import BytesIO
from urllib.parse import urlsplit

KST = timezone(timedelta(hours=9))
HEADERS = ("기사 제목", "매체명", "기사URL", "Naver수집URL")


def _cell(value):
    if isinstance(value, tuple):
        value = value[0]
    return "" if value is None else str(value).strip()


def _http_url(value):
    if isinstance(value, tuple):
        displayed, target = value
        return _http_url(displayed) or _http_url(target)
    value = _cell(value)
    if not value or any(c.isspace() for c in value):
        return ""
    try:
        url = urlsplit(value)
        return value if url.scheme in ("http", "https") and url.hostname else ""
    except ValueError:
        return ""


def _label(value):
    return "".join(_cell(value).split()).casefold()


def _find_layout(sheets):
    if "기사 목록" in sheets:
        rows = sheets["기사 목록"]
        candidates = [rows]
    else:
        candidates = [rows for name, rows in sheets.items() if name not in ("최종", "통계")]
    for rows in candidates:
        for i, row in enumerate(rows[:30]):
            labels = [_label(x) for x in row]
            names = [_label(x) for x in HEADERS]
            if all(x in labels for x in names[:2]):
                return rows[i + 1:], [labels.index(x) if x in labels else None for x in names]
    if "기사 목록" in sheets:
        rows = sheets["기사 목록"]
        if rows and row_is_section(rows[0]) and "장관" in _label(rows[0][0]):
            rows = rows[1:]
        return rows, [0, 1, 7, 8]
    raise ValueError("'기사 목록' 시트 또는 '기사 제목'과 '매체명' 열이 있는 시트가 필요합니다.")


def row_is_section(row):
    return bool(row and _cell(row[0]) and (len(row) < 2 or not _cell(row[1])))


def format_workbook(data: bytes, filename: str, report_date: date | None = None):
    """TXT와 포함된 기사 수를 반환합니다. 매크로를 실행하지 않습니다."""
    ext = filename.rsplit(".", 1)[-1].lower()
    if ext in {"xlsx", "xlsm"}:
        from openpyxl import load_workbook
        with BytesIO(data) as stream:
            workbook = load_workbook(stream, data_only=True)
            try:
                sheets = {}
                for sheet in workbook.worksheets:
                    if sheet.title in {"최종", "통계"}:
                        continue
                    sheets[sheet.title] = [
                        [(c.value, c.hyperlink.target) if c.hyperlink and c.hyperlink.target else c.value for c in row]
                        for row in sheet.iter_rows()
                    ]
            finally:
                workbook.close()
    elif ext == "xls":
        import pandas as pd
        sheets = {}
        with pd.ExcelFile(BytesIO(data), engine="xlrd") as book:
            for name in book.sheet_names:
                if name not in {"최종", "통계"}:
                    sheets[name] = book.parse(name, header=None, keep_default_na=False).values.tolist()
    else:
        raise ValueError(".xlsm, .xlsx, .xls 파일을 선택해 주세요.")
    rows, columns = _find_layout(sheets)
    report_date = report_date or datetime.now(KST).date()
    day = "월화수목금토일"[report_date.weekday()]
    lines = [f"< '{report_date:%y}.{report_date.month}.{report_date.day}({day}), 조간 주요 보도 >", "", ""]
    count = 0
    for row in rows:
        title, press, original, naver = (row[i] if i is not None and i < len(row) else None for i in columns)
        title, press = _cell(title), _cell(press)
        if not title:
            continue
        if not press:
            if len(lines) > 3:
                lines.append("")
            lines.extend([title, ""])
            continue
        lines.extend([f"ㅇ {press}, {title}", _http_url(naver) or _http_url(original), ""])
        count += 1
    if not count:
        raise ValueError("변환할 기사 행이 없습니다. 기사 제목과 매체명을 확인해 주세요.")
    return "\r\n".join(lines).rstrip("\r\n") + "\r\n", count
