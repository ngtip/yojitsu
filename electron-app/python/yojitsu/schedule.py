"""個別予定ファイル（メンバごとの .xlsx）の読み込み

旧実装は機能ごとに別々の読み込み処理を持ち、列の解釈も食い違っていた。
ここで1回だけ読み、全機能が同じ DayEntry を使う。

シート構成（シート名 = YYYYMM。新しい月が左）:
  1行目: 区分見出し（行動予定 / 実績） / 2行目: 列見出し / 3行目以降: 1日1行
  現行: A 日付 | B 勤怠 | C 行先 | D PC持出 | E 入館証持出 | F wifi持出 |
        G〜K 実績（外部設計/内部設計/製造・単体テスト/会議/その他） | L PJ外作業 | M 備考
  旧レイアウトのシートは「PJ外作業」列が無く L が備考。そのため列の位置は2行目の見出しで決める。
"""

import logging
import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Optional, Tuple

from openpyxl import load_workbook

from .dates import to_yyyymm

logger = logging.getLogger(__name__)

HEADER_ROW = 2
FIRST_DATA_ROW = 3
HOUR_CATEGORIES = ('外部設計', '内部設計', '製造/単体テスト', '会議', 'その他')
VACATION = '休暇'
# 持出欄で「持ち出す」を表す値（実データは 有/無。－ や全角空白は「なし」）
YES_MARKS = frozenset({'有', '有り', 'あり', '〇', '○', '◯'})

# 見出し（空白・改行・記号を除いたもの） -> 項目
_HEADER_KEYS = {
    '日付': 'date', '勤怠': 'attendance', '行先': 'location',
    'pc持出': 'pc', '入館証持出': 'badge', 'wifi持出': 'wifi', 'wifi': 'wifi',
    '外部設計': 'hours0', '内部設計': 'hours1', '製造単体テスト': 'hours2', '会議': 'hours3', 'その他': 'hours4',
    'pj外作業': 'pj_outside', '備考': 'remarks',
}
# 見出しが読めないときの既定の列位置（現行レイアウト）
DEFAULT_LAYOUT = {
    'date': 1, 'attendance': 2, 'location': 3, 'pc': 4, 'badge': 5, 'wifi': 6,
    'hours0': 7, 'hours1': 8, 'hours2': 9, 'hours3': 10, 'hours4': 11, 'pj_outside': 12, 'remarks': 13,
}

_DATE_FORMULA = re.compile(r"=A(\d+)\+(\d+)", re.IGNORECASE)


@dataclass(frozen=True)
class DayEntry:
    day: date
    attendance: str = ''
    location: str = ''
    pc: str = ''
    badge: str = ''       # 入館証持出（旧コードでは「宿泊」として扱っていた E 列）
    wifi: str = ''
    hours: Tuple[float, ...] = (0.0,) * len(HOUR_CATEGORIES)
    pj_outside: float = 0.0
    remarks: str = ''

    @property
    def is_vacation(self) -> bool:
        return VACATION in self.attendance

    @property
    def carries_pc(self) -> bool:
        return self.pc in YES_MARKS

    @property
    def project_hours(self) -> float:
        return sum(self.hours)

    @property
    def has_actual(self) -> bool:
        return any(h > 0 for h in self.hours) or self.pj_outside > 0


MonthSchedule = Dict[date, DayEntry]


def _text(value) -> str:
    return str(value).strip() if value is not None else ''


def _number(value) -> float:
    if value is None or value == '':
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _normalize_header(value) -> str:
    return re.sub(r'[\s/／・]', '', _text(value)).lower()


def detect_layout(header_cells) -> Dict[str, int]:
    """2行目の見出しから 項目 -> 列番号 を作る。日付列が見つからなければ既定の並びを使う"""
    layout: Dict[str, int] = {}
    for col, value in enumerate(header_cells, start=1):
        key = _HEADER_KEYS.get(_normalize_header(value))
        if key and key not in layout:
            layout[key] = col
    return layout if 'date' in layout else dict(DEFAULT_LAYOUT)


def _resolve_date(value, formula, row: int, year: int, month: int,
                  previous: Optional[Tuple[int, date]]) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and 1 <= int(value) <= 31:
        return date(year, month, int(value))
    if isinstance(value, str) and value.strip().isdigit():
        return date(year, month, int(value.strip()))
    # openpyxl で作られ Excel で一度も開かれていないファイルは数式の計算結果が無い。
    # 日付列によくある =A3+1 形式だけ自前で解釈する
    if isinstance(formula, str) and previous:
        match = _DATE_FORMULA.fullmatch(formula.strip().replace('$', ''))
        prev_row, prev_date = previous
        if match and int(match.group(1)) == prev_row:
            return prev_date + timedelta(days=int(match.group(2)))
    return None


def read_month(path: Path, year: int, month: int) -> Optional[MonthSchedule]:
    """対象月シートを読む。シートが無ければ None"""
    sheet_name = to_yyyymm(year, month)
    values_wb = load_workbook(path, data_only=True, read_only=True)
    formulas_wb = load_workbook(path, data_only=False, read_only=True)
    try:
        if sheet_name not in values_wb.sheetnames:
            return None
        values_ws = values_wb[sheet_name]
        header = next(values_ws.iter_rows(min_row=HEADER_ROW, max_row=HEADER_ROW, values_only=True), ())
        layout = detect_layout(header)
        width = max(layout.values())
        date_col = layout['date']

        def cell(cells, key: str):
            col = layout.get(key)
            return cells[col - 1] if col else None

        value_rows = values_ws.iter_rows(min_row=FIRST_DATA_ROW, max_col=width, values_only=True)
        formula_rows = formulas_wb[sheet_name].iter_rows(
            min_row=FIRST_DATA_ROW, min_col=date_col, max_col=date_col, values_only=True)

        entries: MonthSchedule = {}
        previous: Optional[Tuple[int, date]] = None
        for row_no, (cells, formula_cells) in enumerate(zip(value_rows, formula_rows), start=FIRST_DATA_ROW):
            cells = tuple(cells) + (None,) * (width - len(cells))
            day = _resolve_date(cells[date_col - 1], formula_cells[0], row_no, year, month, previous)
            if day is None:
                continue
            previous = (row_no, day)
            if (day.year, day.month) != (year, month):
                continue
            entries[day] = DayEntry(
                day=day,
                attendance=_text(cell(cells, 'attendance')),
                location=_text(cell(cells, 'location')),
                pc=_text(cell(cells, 'pc')),
                badge=_text(cell(cells, 'badge')),
                wifi=_text(cell(cells, 'wifi')),
                hours=tuple(_number(cell(cells, f'hours{i}')) for i in range(len(HOUR_CATEGORIES))),
                pj_outside=_number(cell(cells, 'pj_outside')),
                remarks=_text(cell(cells, 'remarks')),
            )
        return entries
    finally:
        values_wb.close()
        formulas_wb.close()


class ScheduleStore:
    """個別予定ディレクトリ。ファイル名（拡張子なし）= メンバの assignment_name"""

    def __init__(self, directory: Path):
        self.directory = Path(directory)
        self._files = {p.stem: p for p in self.directory.glob('*.xlsx') if not p.name.startswith('~$')} \
            if self.directory.exists() else {}
        self._cache: Dict[Tuple[str, int, int], Optional[MonthSchedule]] = {}

    def has_file(self, file_name: str) -> bool:
        return file_name in self._files

    def month(self, file_name: str, year: int, month: int) -> Optional[MonthSchedule]:
        """ファイルもしくはシートが無ければ None。読み込み失敗は警告して None"""
        key = (file_name, year, month)
        if key not in self._cache:
            path = self._files.get(file_name)
            result = None
            if path:
                try:
                    result = read_month(path, year, month)
                except Exception as e:  # 壊れたファイル1件で全体を止めない
                    logger.warning(f"個別予定の読み込みに失敗: {path.name} ({e})")
            self._cache[key] = result
        return self._cache[key]
