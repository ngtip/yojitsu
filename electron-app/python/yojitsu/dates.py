"""日付・年月の共通処理"""

import calendar
from datetime import date, datetime, timedelta
from typing import Iterator, Tuple

WEEKDAY_JA = ("月", "火", "水", "木", "金", "土", "日")


def parse_yyyymm(value: str) -> Tuple[int, int]:
    text = str(value or "").strip()
    if len(text) != 6 or not text.isdigit():
        raise ValueError(f"対象年月は YYYYMM 形式で指定してください: {value}")
    year, month = int(text[:4]), int(text[4:])
    if not 1 <= month <= 12:
        raise ValueError(f"対象年月の月が不正です: {value}")
    return year, month


def to_yyyymm(year: int, month: int) -> str:
    return f"{year:04d}{month:02d}"


def next_month(year: int, month: int) -> Tuple[int, int]:
    return (year + 1, 1) if month == 12 else (year, month + 1)


def prev_month(year: int, month: int) -> Tuple[int, int]:
    return (year - 1, 12) if month == 1 else (year, month - 1)


def month_range(year: int, month: int) -> Tuple[date, date]:
    return date(year, month, 1), date(year, month, calendar.monthrange(year, month)[1])


def iter_dates(start: date, end: date) -> Iterator[date]:
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def iter_year_months(start: date, end: date) -> Iterator[Tuple[int, int]]:
    year, month = start.year, start.month
    while date(year, month, 1) <= end:
        yield year, month
        year, month = next_month(year, month)


def parse_date_arg(value: str) -> date:
    """YYYY-MM-DD / YYYY/MM/DD / YYYYMMDD を受け付ける"""
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.strptime(value, fmt).date()
        except (TypeError, ValueError):
            continue
    raise ValueError(f"日付形式が不正です: {value} (YYYY-MM-DD / YYYY/MM/DD / YYYYMMDD)")


def is_weekend(target: date) -> bool:
    return target.weekday() >= 5
