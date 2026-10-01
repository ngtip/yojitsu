"""XML 直接編集で、画像などテンプレートの他の要素が壊れないこと"""

import zipfile

import pytest
from lxml import etree
from openpyxl import load_workbook

from yojitsu.features.submit_files import SheetData, XmlWriter
from yojitsu.xlsx_patch import XlsxPackage, XlsxPatchError

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


@pytest.fixture
def template(sample_env):
    return sample_env / 'templates' / '作業実績表テンプレート_自社向け.xlsx'


def _zip(path):
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _write(template, output, cells, text=None):
    XmlWriter().write(template, output, SheetData('10月実績', cells, text or {}))
    return _zip(output)


def test_image_and_other_parts_survive(template, tmp_path):
    before = _zip(template)
    after = _write(template, tmp_path / 'out.xlsx', {'B7': 2026, 'N7': '架空 太郎', 'F14': 7.5})

    assert any(name.startswith('xl/media/') for name in before), 'サンプルテンプレートに画像が入っていない'
    assert set(after) - {'xl/calcChain.xml'} == set(before) - {'xl/calcChain.xml'}   # パーツの増減なし
    # セル・シート名・書式以外のパーツ（画像・図形・関連付け・テーマ）は1バイトも変えない
    untouched = [n for n in before if n.startswith(('xl/drawings/', 'xl/media/', 'xl/worksheets/_rels/', 'xl/theme/'))]
    assert untouched
    for name in untouched:
        assert after[name] == before[name], name


def test_values_and_formats(template, tmp_path):
    output = tmp_path / 'out.xlsx'
    _write(template, output, {'B7': 2026, 'E7': 10, 'N7': '架空 太郎', 'F14': 7.5, 'B44': ''},
           text={'F10': '00123'})
    wb = load_workbook(output)
    assert wb.sheetnames == ['10月実績']                       # XX月実績 は残さない
    ws = wb['10月実績']
    assert (ws['B7'].value, ws['E7'].value, ws['N7'].value, ws['F14'].value) == (2026, 10, '架空 太郎', 7.5)
    assert ws['F10'].value == '00123' and ws['F10'].number_format == '@'
    assert ws['F13'].value == '=SUM(F14:F44)'                  # 合計の数式は残る
    assert ws['O14'].value.startswith('=IF(SUM(F14:N14)')
    assert ws['J4'].value == '作 業 実 績 表' and ws['J4'].font.b   # 見出しの書式も残る
    assert len(ws._images) == 1
    assert str(ws.print_area) == "'10月実績'!$B$1:$O$47"        # 印刷範囲も新しいシート名を指す


def test_workbook_recalculates_on_open(template, tmp_path):
    parts = _write(template, tmp_path / 'out.xlsx', {'F14': 8})
    workbook = etree.fromstring(parts['xl/workbook.xml'])
    assert workbook.find('m:calcPr', NS).get('fullCalcOnLoad') == '1'
    assert 'xl/calcChain.xml' not in parts


def test_rows_and_cells_are_inserted_in_order(template, tmp_path):
    parts = _write(template, tmp_path / 'out.xlsx', {'Z100': 'end', 'A3': 'top', 'C14': 1})
    sheet = etree.fromstring(parts['xl/worksheets/sheet1.xml'])
    rows = [int(r.get('r')) for r in sheet.find('m:sheetData', NS)]
    assert rows == sorted(rows)
    cols = [c.get('r') for c in sheet.find("m:sheetData/m:row[@r='14']", NS)]
    assert cols == sorted(cols, key=lambda ref: (len(ref.rstrip('0123456789')), ref))


def test_existing_sheet_is_reused(template, tmp_path):
    first = tmp_path / 'first.xlsx'
    _write(template, first, {'F14': 1})
    package = XlsxPackage(first)
    package.set_cells('10月実績', {'F14': 2})
    package.save(tmp_path / 'second.xlsx')
    wb = load_workbook(tmp_path / 'second.xlsx')
    assert wb.sheetnames == ['10月実績']
    assert wb['10月実績']['F14'].value == 2


def test_missing_sheet_is_rejected(template):
    with pytest.raises(XlsxPatchError):
        XlsxPackage(template).rename_sheet('存在しない', '10月実績')


def test_template_sheet_name_is_case_insensitive(template, tmp_path):
    """本番テンプレートは「XX月実績」。Excel と同じく大文字・小文字を区別せずに探す"""
    package = XlsxPackage(template)
    assert package.find_sheet('xx月実績') == 'XX月実績'
    package.rename_sheet('xx月実績', '10月実績')
    package.save(tmp_path / 'out.xlsx')
    assert load_workbook(tmp_path / 'out.xlsx').sheetnames == ['10月実績']
