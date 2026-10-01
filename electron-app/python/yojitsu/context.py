"""1回の実行に必要な情報（DB・PJ・パス・設定）をまとめる"""

import logging
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import List, Optional

from .db import Database
from .holidays import HolidayCalendar
from .members import Member, load_members
from .schedule import ScheduleStore
from .settings import SiteSettings, ToolPaths, load_site_settings, resolve_tool_paths

logger = logging.getLogger(__name__)


@dataclass
class Warning_:
    message: str
    member: str = ''

    def to_dict(self) -> dict:
        return {'member': self.member, 'message': self.message}


class RunContext:
    def __init__(self, db: Database, project_id: Optional[str] = None,
                 site: Optional[SiteSettings] = None, paths: Optional[ToolPaths] = None):
        self.db = db
        self.project = db.resolve_project(project_id)
        self.project_id: str = self.project['project_id']
        self.project_name: str = self.project.get('project_name') or 'PJ'
        self.project_code: str = self.project.get('project_code') or ''
        self.site = site or load_site_settings()
        self.paths = paths or resolve_tool_paths(db)
        self.schedules = ScheduleStore(self.paths.schedules)
        self.warnings: List[Warning_] = []
        self._unknown_groups = set()
        if self.site.problem:
            self.warnings.append(Warning_(self.site.problem))

    def members(self, start: date, end: Optional[date] = None) -> List[Member]:
        members = load_members(self.db, self.project_id, self.site, start, end)
        # site-settings の拠点名と DB の拠点名が一致しないと、並び順と人数集計から外れる
        configured = {group.name for group in self.site.groups}
        for group in sorted({m.group for m in members} - configured - self._unknown_groups):
            if configured:
                self.warn(f"拠点「{group}」が site-settings.json の groups にありません（並び順・人数集計の対象外）")
            self._unknown_groups.add(group)
        return members

    def holidays(self, start: date, end: date) -> HolidayCalendar:
        return HolidayCalendar.load(self.db, self.project_id, start, end)

    def warn(self, message: str, member: str = '') -> None:
        logger.warning(f"{member}: {message}" if member else message)
        self.warnings.append(Warning_(message, member))

    def output_dir(self, *parts: str) -> Path:
        path = self.paths.output.joinpath(*parts)
        path.mkdir(parents=True, exist_ok=True)
        return path
