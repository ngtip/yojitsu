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

    def members(self, start: date, end: Optional[date] = None) -> List[Member]:
        return load_members(self.db, self.project_id, self.site, start, end)

    def holidays(self, start: date, end: date) -> HolidayCalendar:
        return HolidayCalendar.load(self.db, self.project_id, start, end)

    def warn(self, message: str, member: str = '') -> None:
        logger.warning(f"{member}: {message}" if member else message)
        self.warnings.append(Warning_(message, member))

    def output_dir(self, *parts: str) -> Path:
        path = self.paths.output.joinpath(*parts)
        path.mkdir(parents=True, exist_ok=True)
        return path
