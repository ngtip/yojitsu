"""祝日・休暇カレンダー

DB の holidays（グローバル＋PJ固有）を日付単位に展開して保持する。
旧実装はスクリプトごとに対象区分がばらばらだったため、ここで用途別に定義を一本化する。
"""

import logging
from dataclasses import dataclass, field
from datetime import date
from typing import Dict, Iterable, Optional, Set

from .db import Database
from .dates import is_weekend, iter_dates

logger = logging.getLogger(__name__)

PUBLIC_HOLIDAY = '祝日'
# 営業日から除外する区分（「イベント」等は稼働日扱い）
NON_WORKING_CATEGORIES = frozenset({'祝日', '休暇', '公休', '会社公休'})


@dataclass
class HolidayCalendar:
    by_date: Dict[date, Set[str]] = field(default_factory=dict)

    @classmethod
    def from_rows(cls, rows: Iterable[dict]) -> 'HolidayCalendar':
        calendar = cls()
        for row in rows:
            category = str(row.get('category') or '').strip()
            try:
                start = date.fromisoformat(str(row['start_date']))
                end = date.fromisoformat(str(row['end_date'])) if row.get('end_date') else start
            except (KeyError, ValueError):
                logger.warning(f"休日レコードの日付形式が不正なためスキップ: holiday_id={row.get('holiday_id')}")
                continue
            for day in iter_dates(start, end):
                calendar.by_date.setdefault(day, set()).add(category)
        return calendar

    @classmethod
    def load(cls, db: Database, project_id: Optional[str], start: date, end: date) -> 'HolidayCalendar':
        rows = db.get_holidays(project_id, start.isoformat(), end.isoformat(), include_global=True)
        return cls.from_rows(rows)

    def is_public_holiday(self, day: date) -> bool:
        return PUBLIC_HOLIDAY in self.by_date.get(day, ())

    def is_day_off(self, day: date) -> bool:
        """祝日・休暇・公休など（土日は含まない）"""
        return bool(self.by_date.get(day, set()) & NON_WORKING_CATEGORIES)

    def is_business_day(self, day: date) -> bool:
        return not is_weekend(day) and not self.is_day_off(day)

    def business_days(self, start: date, end: date) -> int:
        return sum(1 for day in iter_dates(start, end) if self.is_business_day(day))
