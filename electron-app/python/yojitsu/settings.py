"""設定の解決

- 環境固有設定（自社名・テンプレート名・拠点）: config/site-settings.json（Git管理外）
- 作業ディレクトリ: DB の tool_settings（相対パスは tool_root_dir 基準）
- ツール別設定: DB の tool_config
"""

import json
import logging
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from .db import Database

logger = logging.getLogger(__name__)

APP_ROOT = Path(__file__).resolve().parents[2]  # electron-app/
CONFIG_DIR = APP_ROOT / 'config'

DEFAULT_MONTHLY_TEMPLATE = '月間カレンダーテンプレ.xlsx'
DEFAULT_LIST_TEMPLATE = '一覧カレンダーテンプレ.xlsx'


@dataclass(frozen=True)
class Group:
    name: str                   # メンバ定義の拠点名
    label: str                  # 一覧カレンダーに出す表示名
    threshold: int              # この人数を超えたら黄色
    keywords: Tuple[str, ...]   # 行先にこれを含む人を数える（未指定なら label）


@dataclass
class SiteSettings:
    own_company_shortname: str = ''
    own_company_fullname: str = ''
    own_template_file: str = '作業実績表テンプレート_自社向け.xlsx'
    pj_template_file: str = '作業実績表テンプレート_PJ向け.xlsx'
    groups: List[Group] = field(default_factory=list)
    date_highlight_label: str = ''
    date_highlight_min_count: int = 0

    @property
    def default_group(self) -> str:
        return self.groups[0].name if self.groups else ''

    def group_order(self, group_name: Optional[str]) -> int:
        for index, group in enumerate(self.groups):
            if group.name == (group_name or self.default_group):
                return index
        return len(self.groups)


def load_site_settings(path: Optional[Path] = None) -> SiteSettings:
    path = path or Path(os.environ.get('YOJITSU_SITE_SETTINGS') or CONFIG_DIR / 'site-settings.json')
    raw: Dict[str, Any] = {}
    if path.exists():
        try:
            raw = json.loads(path.read_text(encoding='utf-8'))
        except (OSError, ValueError) as e:
            logger.warning(f"site-settings の読み込みに失敗したため既定値を使います: {e}")
    else:
        logger.warning(f"site-settings が無いため既定値を使います: {path}")

    defaults = SiteSettings()
    highlight = raw.get('date_highlight') or {}
    return SiteSettings(
        own_company_shortname=raw.get('own_company_shortname', defaults.own_company_shortname),
        own_company_fullname=raw.get('own_company_fullname', defaults.own_company_fullname),
        own_template_file=raw.get('own_template_file', defaults.own_template_file),
        pj_template_file=raw.get('pj_template_file', defaults.pj_template_file),
        groups=[
            Group(
                name=g['name'],
                label=g.get('label') or g['name'],
                threshold=int(g.get('threshold', 999)),
                keywords=tuple(g.get('keywords') or [g.get('label') or g['name']]),
            )
            for g in raw.get('groups') or []
        ],
        date_highlight_label=highlight.get('label', ''),
        date_highlight_min_count=int(highlight.get('min_count') or 0),
    )


@dataclass(frozen=True)
class ToolPaths:
    root: Path
    templates: Path
    output: Path
    schedules: Path

    def template(self, file_name: str) -> Path:
        return self.templates / file_name


def resolve_tool_paths(db: Database) -> ToolPaths:
    settings = db.get_tool_settings()
    root = Path(settings['tool_root_dir']).expanduser()

    def resolve(value: str) -> Path:
        path = Path(str(value or '').strip()).expanduser()
        return path if path.is_absolute() else root / path

    return ToolPaths(
        root=root,
        templates=resolve(settings['template_base_dir']),
        output=resolve(settings['output_base_dir']),
        schedules=resolve(settings['schedule_base_dir']),
    )


def template_file_name(db: Database, tool_name: str, default: str, project_id: Optional[str] = None) -> str:
    return db.get_config(tool_name, 'template_file', project_id=project_id) or default
