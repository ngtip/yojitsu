"""実績時間集計（BP → プロパーの順に1シートへ出力）

メンバごとに、当月の実績合計・PJ外作業・入力済み営業日数・休日出勤を集計し、
残り営業日から月末の予測値を2通り出す。
  予測①: 実績 + 残営業日 × 8h
  予測②: 実績 + 残営業日 × (8h + 1日平均の残業)
"""

import logging
from dataclasses import dataclass
from datetime import date
from typing import Dict, List, Optional

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font

from ..context import RunContext
from ..dates import month_range
from ..excel import save_with_fallback, solid
from ..holidays import HolidayCalendar
from ..members import Member
from ..schedule import DayEntry

logger = logging.getLogger(__name__)

STANDARD_HOURS_PER_DAY = 8
# 月末予測がこの範囲を外れたら強調する
FORECAST_LOWER, FORECAST_UPPER = 140, 180
HOLIDAY_WORK_KEYWORD = '出勤'


@dataclass
class MemberHours:
    member: Member
    actual: float = 0.0             # G〜K の合計
    pj_outside: float = 0.0         # L
    business_input_days: int = 0    # 入力済みの営業日数（休暇は実績が無くても入力済み）
    holiday_work_planned: int = 0   # 土日祝で勤怠に「出勤」
    holiday_work_actual: int = 0    # そのうち実績あり
    last_actual_day: Optional[date] = None
    avg_overtime: float = 0.0       # 1日平均の残業（0.5h 単位）
    business_days: int = 0
    forecast_standard: float = 0.0
    forecast_with_overtime: float = 0.0


def summarize(member: Member, entries: Dict[date, DayEntry], holidays: HolidayCalendar,
              year: int, month: int) -> MemberHours:
    start, end = month_range(year, month)
    result = MemberHours(member, business_days=holidays.business_days(start, end))
    overtime_total = 0.0

    for day in sorted(entries):
        entry = entries[day]
        result.actual += entry.project_hours
        result.pj_outside += entry.pj_outside
        if entry.has_actual:
            result.last_actual_day = day
        overtime_total += max(0.0, entry.project_hours - STANDARD_HOURS_PER_DAY)

        if holidays.is_business_day(day):
            if entry.is_vacation or (entry.attendance and entry.has_actual):
                result.business_input_days += 1
        elif HOLIDAY_WORK_KEYWORD in entry.attendance:
            result.holiday_work_planned += 1
            if entry.has_actual:
                result.holiday_work_actual += 1

    result.actual = round(result.actual, 2)
    result.pj_outside = round(result.pj_outside, 2)
    if result.business_input_days:
        # 0.25 は 0.5 に寄せる（四捨五入ではなく 0.5 刻みの丸め）
        result.avg_overtime = round(overtime_total / result.business_input_days * 2) / 2
    remaining = max(result.business_days - result.business_input_days, 0)
    result.forecast_standard = round(result.actual + remaining * STANDARD_HOURS_PER_DAY, 2)
    result.forecast_with_overtime = round(
        result.actual + remaining * (STANDARD_HOURS_PER_DAY + result.avg_overtime), 2)
    return result


HEADERS = (
    ('メンバー名', 20), ('実績合計（h）', 16), ('月末予測値（8h/日）', 22), ('平均残業実績（h/日）', 18),
    ('月末予測値（実績平均）', 22), ('PJ外作業合計（h）', 18), ('PJ外作業込実績合計（h）', 24),
    ('営業日数（月）', 16), ('営業日入力済み日数', 16), ('休出予定日', 14), ('休出実績日', 14),
)
HEADER_ROW = 4
GRAY = solid('D3D3D3')


def write_sheet(ws, rows: List[MemberHours], year: int, month: int, business_days: int) -> None:
    ws['A1'], ws['B1'] = '対象年月', f"{year}年{month:02d}月"
    ws['A2'], ws['B2'] = '標準時間', business_days * STANDARD_HOURS_PER_DAY
    ws['C2'] = f"（営業日{business_days}日 × {STANDARD_HOURS_PER_DAY}h）"
    ws['B2'].number_format = '0"h"'
    for ref in ('A1', 'A2'):
        ws[ref].font, ws[ref].fill = Font(bold=True), GRAY
        ws[ref].alignment = Alignment(horizontal='center', vertical='center')

    for col, (title, width) in enumerate(HEADERS, start=1):
        cell = ws.cell(HEADER_ROW, col, title)
        cell.font, cell.fill = Font(bold=True), GRAY
        cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.column_dimensions[cell.column_letter].width = width

    for row, item in enumerate(rows, start=HEADER_ROW + 1):
        values = (
            item.member.display_name, item.actual, item.forecast_standard, item.avg_overtime,
            item.forecast_with_overtime, item.pj_outside, item.actual + item.pj_outside,
            item.business_days, item.business_input_days, item.holiday_work_planned, item.holiday_work_actual,
        )
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row, col, value)
            if 2 <= col <= 7:
                cell.number_format = '0.0'

    last_row = HEADER_ROW + len(rows)
    if rows:
        warn_fill, warn_font = solid('FFFF00'), Font(color='FF0000')
        for col in ('C', 'E'):
            for operator, bound in (('lessThan', FORECAST_LOWER), ('greaterThan', FORECAST_UPPER)):
                ws.conditional_formatting.add(
                    f"{col}{HEADER_ROW + 1}:{col}{last_row}",
                    CellIsRule(operator=operator, formula=[str(bound)], fill=warn_fill, font=warn_font),
                )
    ws.freeze_panes = f'A{HEADER_ROW + 1}'


def run(ctx: RunContext, year: int, month: int) -> dict:
    start, end = month_range(year, month)
    holidays = ctx.holidays(start, end)
    members = ctx.members(start, end)
    # BP を先、プロパーを後（それぞれ定義順）
    members = sorted(members, key=lambda m: (m.is_proprietary, m.order))

    rows = []
    for member in members:
        if not ctx.schedules.has_file(member.file_name):
            ctx.warn(f"個別予定ファイルが見つかりません（{member.file_name}.xlsx）", member.display_name)
        entries = ctx.schedules.month(member.file_name, year, month) or {}
        item = summarize(member, entries, holidays, year, month)
        logger.info(f"{member.display_name}: 実績{item.actual}h / PJ外{item.pj_outside}h / "
                    f"入力済み{item.business_input_days}日 / 休出{item.holiday_work_planned}日")
        rows.append(item)

    wb = Workbook()
    ws = wb.active
    ws.title = '集計結果'
    write_sheet(ws, rows, year, month, holidays.business_days(start, end))
    saved = save_with_fallback(wb, ctx.output_dir() / f"実績集計_{year}{month:02d}.xlsx")
    return {'outputs': [str(saved)]}
