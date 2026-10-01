"""SQLite アクセス層

正本データ（PJ・メンバ・PJ所属・祝日休暇・ツール設定）はすべてこの DB にある。
既存の実環境 DB とスキーマを共有するため、CREATE は IF NOT EXISTS のみで既存列には触れない。
"""

import logging
import sqlite3
from pathlib import Path
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)

SCHEMA = """
CREATE TABLE IF NOT EXISTS tool_settings (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    tool_root_dir TEXT NOT NULL,
    template_base_dir TEXT NOT NULL,
    output_base_dir TEXT NOT NULL,
    schedule_base_dir TEXT NOT NULL,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS tool_config (
    config_id INTEGER PRIMARY KEY AUTOINCREMENT,
    tool_name TEXT NOT NULL,
    project_id TEXT,
    config_key TEXT NOT NULL,
    config_value TEXT,
    data_type TEXT DEFAULT 'string',
    description TEXT
);

CREATE TABLE IF NOT EXISTS admins (
    admin_id TEXT PRIMARY KEY,
    admin_name TEXT NOT NULL,
    organization TEXT,
    is_active INTEGER DEFAULT 1,
    notes TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS projects (
    project_id TEXT PRIMARY KEY,
    project_name TEXT NOT NULL UNIQUE,
    admin_id TEXT,
    description TEXT,
    is_active INTEGER NOT NULL DEFAULT 1,
    start_date TEXT,
    end_date TEXT,
    project_code TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS members (
    member_id TEXT PRIMARY KEY,
    member_no TEXT,
    full_name TEXT NOT NULL,
    display_name TEXT,
    abbreviation TEXT,
    group_name TEXT,
    organization TEXT,
    notes TEXT,
    file_storage_location TEXT NOT NULL DEFAULT 'internal',
    is_proprietary INTEGER NOT NULL DEFAULT 0,
    schedule_file_name TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

-- 参画1回につき1レコード（同じメンバが期間を空けて複数回参画できる）
CREATE TABLE IF NOT EXISTS project_members (
    project_member_id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT NOT NULL,
    member_id TEXT NOT NULL,
    start_date TEXT NOT NULL,
    end_date TEXT,
    assignment_name TEXT,
    is_proprietary INTEGER DEFAULT 0,
    include_in_monthly INTEGER DEFAULT 1,
    notes TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (project_id, member_id, start_date)
);

CREATE TABLE IF NOT EXISTS holidays (
    holiday_id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT,
    name TEXT NOT NULL,
    category TEXT NOT NULL,
    start_date TEXT NOT NULL,
    end_date TEXT,
    notes TEXT,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);
"""

MEMBER_UPDATABLE_FIELDS = (
    'full_name', 'display_name', 'group_name', 'organization', 'abbreviation',
    'member_no', 'file_storage_location', 'is_proprietary', 'schedule_file_name',
)
HOLIDAY_FIELDS = ('project_id', 'name', 'category', 'start_date', 'end_date', 'notes')


class Database:
    def __init__(self, db_file, create: bool = False):
        self.db_file = Path(db_file)
        if not create and not self.db_file.exists():
            raise FileNotFoundError(f"DB ファイルが見つかりません: {self.db_file}")
        self.db_file.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(self.db_file))
        self.conn.row_factory = sqlite3.Row
        if create:
            self.conn.executescript(SCHEMA)
        self._migrate()

    def close(self) -> None:
        self.conn.close()

    def __enter__(self) -> 'Database':
        return self

    def __exit__(self, *_exc) -> None:
        self.close()

    def _columns(self, table: str) -> set:
        return {row['name'] for row in self.conn.execute(f"PRAGMA table_info({table})")}

    def _migrate(self) -> None:
        if self._columns('projects') and 'project_code' not in self._columns('projects'):
            with self.conn:
                self.conn.execute("ALTER TABLE projects ADD COLUMN project_code TEXT")
            logger.info("スキーマ更新: projects.project_code を追加")
        # 個別予定のファイル名は参画ごとではなくメンバで固定する。旧来の project_members.assignment_name から移す
        if self._columns('members') and 'schedule_file_name' not in self._columns('members'):
            with self.conn:
                self.conn.execute("ALTER TABLE members ADD COLUMN schedule_file_name TEXT")
                self.conn.execute(
                    """
                    UPDATE members SET schedule_file_name = (
                        SELECT pm.assignment_name FROM project_members pm
                        WHERE pm.member_id = members.member_id AND pm.assignment_name IS NOT NULL
                        ORDER BY pm.start_date DESC LIMIT 1)
                    """
                )
            logger.info("スキーマ更新: members.schedule_file_name を追加し、PJ所属のファイル名を移行")

    def _all(self, query: str, params=()) -> List[Dict[str, Any]]:
        return [dict(row) for row in self.conn.execute(query, params)]

    def _one(self, query: str, params=()) -> Optional[Dict[str, Any]]:
        row = self.conn.execute(query, params).fetchone()
        return dict(row) if row else None

    # ---------- 設定 ----------

    def get_config(self, tool_name: str, config_key: str, project_id: Optional[str] = None) -> Optional[str]:
        """PJ固有 → グローバルの順に参照する"""
        if project_id:
            row = self._one(
                "SELECT config_value FROM tool_config WHERE tool_name = ? AND project_id = ? AND config_key = ?",
                (tool_name, project_id, config_key),
            )
            if row:
                return row['config_value']
        # 旧実装の NULL UPSERT で重複行が残っている DB があるため最新行を採用
        row = self._one(
            "SELECT config_value FROM tool_config WHERE tool_name = ? AND project_id IS NULL AND config_key = ?"
            " ORDER BY config_id DESC LIMIT 1",
            (tool_name, config_key),
        )
        return row['config_value'] if row else None

    def set_config(self, tool_name: str, config_key: str, config_value: str,
                   project_id: Optional[str] = None, data_type: str = 'string',
                   description: Optional[str] = None) -> None:
        with self.conn:
            # SQLite < 3.39 では NULL project_id の ON CONFLICT が効かないため DELETE + INSERT
            self.conn.execute(
                "DELETE FROM tool_config WHERE tool_name = ? AND project_id IS ? AND config_key = ?",
                (tool_name, project_id, config_key),
            )
            self.conn.execute(
                "INSERT INTO tool_config (tool_name, project_id, config_key, config_value, data_type, description)"
                " VALUES (?, ?, ?, ?, ?, ?)",
                (tool_name, project_id, config_key, config_value, data_type or 'string', description),
            )

    def get_tool_settings(self) -> Dict[str, Any]:
        row = self._one("SELECT * FROM tool_settings WHERE id = 1")
        if not row:
            raise RuntimeError("tool_settings が初期化されていません")
        return row

    def update_tool_settings(self, **values: Optional[str]) -> None:
        current = self.get_tool_settings()
        keys = ('tool_root_dir', 'template_base_dir', 'output_base_dir', 'schedule_base_dir')
        merged = [values.get(key) or current[key] for key in keys]
        with self.conn:
            self.conn.execute(
                "UPDATE tool_settings SET tool_root_dir = ?, template_base_dir = ?, output_base_dir = ?,"
                " schedule_base_dir = ?, updated_at = CURRENT_TIMESTAMP WHERE id = 1",
                merged,
            )

    def init_tool_settings(self, tool_root_dir: str, template_base_dir: str,
                           output_base_dir: str, schedule_base_dir: str) -> None:
        with self.conn:
            self.conn.execute(
                "INSERT OR REPLACE INTO tool_settings"
                " (id, tool_root_dir, template_base_dir, output_base_dir, schedule_base_dir)"
                " VALUES (1, ?, ?, ?, ?)",
                (tool_root_dir, template_base_dir, output_base_dir, schedule_base_dir),
            )

    # ---------- プロジェクト ----------

    _PROJECT_SELECT = """
        SELECT p.*, a.admin_name FROM projects p LEFT JOIN admins a ON a.admin_id = p.admin_id
    """

    def get_projects(self, active_only: bool = True) -> List[Dict[str, Any]]:
        where = "WHERE p.is_active = 1 " if active_only else ""
        return self._all(f"{self._PROJECT_SELECT} {where}ORDER BY p.project_name")

    def get_project(self, project_id: str) -> Optional[Dict[str, Any]]:
        return self._one(f"{self._PROJECT_SELECT} WHERE p.project_id = ?", (project_id,))

    def resolve_project(self, project_id: Optional[str] = None) -> Dict[str, Any]:
        """指定がなければ有効なPJの先頭を使う"""
        if project_id:
            project = self.get_project(project_id)
            if not project:
                raise ValueError(f"プロジェクトが見つかりません: {project_id}")
            return project
        projects = self.get_projects(active_only=True)
        if not projects:
            raise RuntimeError("有効なプロジェクトがありません")
        return projects[0]

    # ---------- メンバ ----------

    def get_members(self) -> List[Dict[str, Any]]:
        return self._all("SELECT * FROM members ORDER BY member_id")

    def get_member(self, member_id: str) -> Optional[Dict[str, Any]]:
        return self._one("SELECT * FROM members WHERE member_id = ?", (member_id,))

    def update_member(self, member_id: str, **changes: Any) -> Dict[str, Any]:
        current = self.get_member(member_id)
        if not current:
            raise ValueError(f"メンバが見つかりません: {member_id}")

        merged = {key: current.get(key) for key in MEMBER_UPDATABLE_FIELDS}
        merged.update({key: value for key, value in changes.items()
                       if key in MEMBER_UPDATABLE_FIELDS and value is not None})
        merged['is_proprietary'] = int(merged['is_proprietary'] or 0)
        merged['file_storage_location'] = merged['file_storage_location'] or 'internal'

        if not str(merged['full_name'] or '').strip():
            raise ValueError("full_name は必須です")
        if merged['file_storage_location'] not in ('internal', 'external'):
            raise ValueError("file_storage_location は internal / external のいずれかです")
        if merged['is_proprietary'] not in (0, 1):
            raise ValueError("is_proprietary は 0 / 1 のいずれかです")

        assignments = ", ".join(f"{key} = ?" for key in MEMBER_UPDATABLE_FIELDS)
        with self.conn:
            self.conn.execute(
                f"UPDATE members SET {assignments}, updated_at = CURRENT_TIMESTAMP WHERE member_id = ?",
                [merged[key] for key in MEMBER_UPDATABLE_FIELDS] + [member_id],
            )
        return {'member_id': member_id, **merged}

    def get_project_members(self, project_id: str, period_start: str,
                            period_end: Optional[str] = None) -> List[Dict[str, Any]]:
        """期間に1日でも重なる参画レコード（1人が複数回参画していれば複数行）。
        period_end 省略時は period_start 時点の在籍者"""
        period_end = period_end or period_start
        return self._all(
            """
            SELECT m.member_id, m.member_no, m.full_name, m.display_name, m.group_name,
                   m.organization, m.abbreviation, m.file_storage_location,
                   COALESCE(m.schedule_file_name, pm.assignment_name) AS schedule_file_name,
                   pm.project_member_id, pm.is_proprietary, pm.include_in_monthly,
                   pm.start_date, pm.end_date
            FROM project_members pm
            JOIN members m ON pm.member_id = m.member_id
            WHERE pm.project_id = ?
              AND pm.start_date <= ?
              AND (pm.end_date IS NULL OR pm.end_date >= ?)
            ORDER BY m.member_no, m.member_id, pm.start_date
            """,
            (project_id, period_end, period_start),
        )

    def get_external_member_aliases(self) -> set:
        """ファイル格納区分=external のメンバを識別する名前の集合（同期対象から除外する）"""
        aliases = set()
        for row in self._all(
            """
            SELECT m.member_id, m.display_name, m.full_name, m.schedule_file_name, pm.assignment_name
            FROM members m LEFT JOIN project_members pm ON pm.member_id = m.member_id
            WHERE m.file_storage_location = 'external'
            """
        ):
            aliases.update(str(value).strip() for value in row.values() if value and str(value).strip())
        return aliases

    # ---------- 祝日・休暇 ----------

    def get_holidays(self, project_id: Optional[str] = None, start_date: Optional[str] = None,
                     end_date: Optional[str] = None, include_global: bool = False) -> List[Dict[str, Any]]:
        """project_id=None はグローバル。include_global=True ならPJ固有とグローバルを合わせて返す"""
        query = "SELECT * FROM holidays WHERE "
        params: list = []
        if project_id is None:
            query += "project_id IS NULL"
        elif include_global:
            query += "(project_id IS NULL OR project_id = ?)"
            params.append(project_id)
        else:
            query += "project_id = ?"
            params.append(project_id)

        # end_date が NULL の行は単日扱い
        if start_date:
            query += " AND COALESCE(end_date, start_date) >= ?"
            params.append(start_date)
        if end_date:
            query += " AND start_date <= ?"
            params.append(end_date)
        return self._all(query + " ORDER BY start_date", params)

    def add_holiday(self, name: str, category: str, start_date: str, end_date: Optional[str] = None,
                    project_id: Optional[str] = None, notes: Optional[str] = None) -> int:
        values = _validate_holiday(name, category, start_date, end_date)
        existing = self._one(
            """
            SELECT holiday_id FROM holidays
            WHERE project_id IS ? AND name = ? AND category = ? AND start_date = ? AND end_date IS ?
            """,
            (project_id, values['name'], values['category'], values['start_date'], values['end_date']),
        )
        if existing:
            return int(existing['holiday_id'])
        with self.conn:
            cursor = self.conn.execute(
                "INSERT INTO holidays (project_id, name, category, start_date, end_date, notes)"
                " VALUES (?, ?, ?, ?, ?, ?)",
                (project_id, values['name'], values['category'], values['start_date'], values['end_date'], notes),
            )
        return int(cursor.lastrowid)

    def update_holiday(self, holiday_id: int, **changes: Any) -> None:
        current = self._one("SELECT * FROM holidays WHERE holiday_id = ?", (holiday_id,))
        if not current:
            raise ValueError(f"holiday_id が見つかりません: {holiday_id}")
        merged = {key: current.get(key) for key in HOLIDAY_FIELDS}
        merged.update({key: value for key, value in changes.items() if key in HOLIDAY_FIELDS and value is not None})
        # 終了日は空文字で「単日に戻す」指定になる
        if 'end_date' in changes:
            merged['end_date'] = changes['end_date'] or None
        merged.update(_validate_holiday(merged['name'], merged['category'], merged['start_date'], merged['end_date']))

        assignments = ", ".join(f"{key} = ?" for key in HOLIDAY_FIELDS)
        with self.conn:
            self.conn.execute(
                f"UPDATE holidays SET {assignments}, updated_at = CURRENT_TIMESTAMP WHERE holiday_id = ?",
                [merged[key] for key in HOLIDAY_FIELDS] + [holiday_id],
            )

    def delete_holiday(self, holiday_id: int) -> None:
        with self.conn:
            cursor = self.conn.execute("DELETE FROM holidays WHERE holiday_id = ?", (holiday_id,))
        if cursor.rowcount == 0:
            raise ValueError(f"holiday_id が見つかりません: {holiday_id}")


def _validate_holiday(name, category, start_date, end_date) -> Dict[str, Optional[str]]:
    values = {
        'name': str(name or '').strip(),
        'category': str(category or '').strip(),
        'start_date': str(start_date or '').strip(),
        'end_date': str(end_date or '').strip() or None,
    }
    for key in ('name', 'category', 'start_date'):
        if not values[key]:
            raise ValueError(f"{key} は必須です")
    if values['end_date'] and values['end_date'] < values['start_date']:
        raise ValueError("end_date は start_date 以降を指定してください")
    return values
