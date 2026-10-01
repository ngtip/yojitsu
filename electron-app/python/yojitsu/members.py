"""PJメンバ

参画は1回につき1レコード（project_members）。帳票は月単位なので、対象期間に1日でも参画していれば
そのメンバを載せ、複数回の参画は1人分にまとめて参画期間の一覧として持つ。
"""

from dataclasses import dataclass
from datetime import date
from typing import Dict, List, Optional, Tuple

from .db import Database
from .settings import SiteSettings

Period = Tuple[date, Optional[date]]   # (開始日, 終了日。None は継続中)


@dataclass(frozen=True)
class Member:
    member_id: str
    full_name: str
    display_name: str
    file_name: str          # 個別予定ファイル名（拡張子なし）。メンバで固定
    group: str
    is_proprietary: bool    # True=プロパー, False=BP
    in_monthly: bool        # 月間カレンダーに載せるか
    organization: str       # 所属（BPは会社名）
    abbreviation: str       # 所属の略称
    order: int              # 定義順
    periods: Tuple[Period, ...] = ()

    def participates(self, day: date) -> bool:
        """参画期間の情報が無いメンバ（単体で作ったもの等）は常に参画中とみなす"""
        if not self.periods:
            return True
        return any(start <= day and (end is None or day <= end) for start, end in self.periods)

    def periods_within(self, start: date, end: date) -> List[Period]:
        """start〜end に切り詰めた参画期間"""
        if not self.periods:
            return [(start, end)]
        return [(max(s, start), min(e or end, end)) for s, e in self.periods if s <= end and (e is None or e >= start)]

    def covers(self, start: date, end: date) -> bool:
        return self.periods_within(start, end) == [(start, end)]


def load_members(db: Database, project_id: str, site: SiteSettings,
                 period_start: date, period_end: Optional[date] = None) -> List[Member]:
    rows = db.get_project_members(
        project_id, period_start.isoformat(), period_end.isoformat() if period_end else None
    )
    by_member: Dict[str, List[dict]] = {}
    for row in rows:
        by_member.setdefault(str(row['member_id']), []).append(row)

    members = []
    for order, stints in enumerate(by_member.values()):
        latest = stints[-1]     # 参画区分などは期間内で最後の参画のものを使う
        members.append(Member(
            member_id=str(latest['member_id']),
            full_name=latest.get('full_name') or '',
            display_name=latest.get('display_name') or latest.get('full_name') or '',
            file_name=latest.get('schedule_file_name') or '',
            group=latest.get('group_name') or site.default_group,
            is_proprietary=bool(latest.get('is_proprietary')),
            in_monthly=bool(latest.get('include_in_monthly')),
            organization=latest.get('organization') or '',
            abbreviation=latest.get('abbreviation') or '',
            order=order,
            periods=tuple(
                (date.fromisoformat(s['start_date']), date.fromisoformat(s['end_date']) if s.get('end_date') else None)
                for s in stints
            ),
        ))
    return members


def sort_by_group(members: List[Member], site: SiteSettings) -> List[Member]:
    """拠点（site-settings の定義順。設定に無い拠点はその後ろに拠点名ごと）→ プロパー優先 → 定義順"""
    return sorted(members, key=lambda m: (site.group_order(m.group), m.group, not m.is_proprietary, m.order))
