"""PJメンバ"""

from dataclasses import dataclass
from datetime import date
from typing import List, Optional

from .db import Database
from .settings import SiteSettings


@dataclass(frozen=True)
class Member:
    member_id: str
    full_name: str
    display_name: str
    file_name: str          # 個別予定ファイル名（拡張子なし）。PJ所属の assignment_name
    group: str
    is_proprietary: bool    # True=プロパー, False=BP
    in_monthly: bool        # 月間カレンダーに載せるか
    organization: str       # 所属（BPは会社名）
    abbreviation: str       # 所属の略称
    order: int              # 定義順

    @classmethod
    def from_row(cls, row: dict, order: int, default_group: str) -> 'Member':
        return cls(
            member_id=str(row['member_id']),
            full_name=row.get('full_name') or '',
            display_name=row.get('display_name') or row.get('full_name') or '',
            file_name=row.get('assignment_name') or '',
            group=row.get('group_name') or default_group,
            is_proprietary=bool(row.get('is_proprietary')),
            in_monthly=bool(row.get('include_in_monthly')),
            organization=row.get('organization') or '',
            abbreviation=row.get('abbreviation') or '',
            order=order,
        )


def load_members(db: Database, project_id: str, site: SiteSettings,
                 period_start: date, period_end: Optional[date] = None) -> List[Member]:
    rows = db.get_project_members(
        project_id, period_start.isoformat(), period_end.isoformat() if period_end else None
    )
    return [Member.from_row(row, index, site.default_group) for index, row in enumerate(rows)]


def sort_by_group(members: List[Member], site: SiteSettings) -> List[Member]:
    """拠点（site-settings の定義順）→ プロパー優先 → 定義順"""
    return sorted(members, key=lambda m: (site.group_order(m.group), not m.is_proprietary, m.order))
