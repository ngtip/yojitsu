"""xlsx を XML レベルで直接編集する

openpyxl はブックを読み込んで作り直すため、画像・図形などを含むテンプレートが壊れる。
ここでは zip 内の XML のうち必要な箇所だけを書き換え、他のパーツ（画像・図形・印刷設定など）は
バイト単位でそのまま残す。名前空間の接頭辞（mc:Ignorable が参照する x14ac 等）を保つため lxml を使う。

できること:
  - シート名の変更（印刷範囲などの参照も追従）
  - セルへの数値・文字列の書き込み / 値のクリア（書式は残す）
  - セルを文字列書式（@）にする
"""

import posixpath
import re
import zipfile
from copy import deepcopy
from pathlib import Path
from typing import Dict, Optional, Tuple

from lxml import etree

NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS_XML = 'http://www.w3.org/XML/1998/namespace'
REL_CALC_CHAIN = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain'
TEXT_FORMAT_ID = '49'   # 組み込みの「@」

M = f'{{{NS_MAIN}}}'
_CELL_REF = re.compile(r'([A-Z]+)(\d+)$')

# workbook.xml で calcPr より後ろに来る要素
_AFTER_CALC_PR = {'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes',
                  'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'}


class XlsxPatchError(RuntimeError):
    pass


def _rels_path(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')


def _resolve(base_part: str, target: str) -> str:
    if target.startswith('/'):
        return target.lstrip('/')
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def _col_number(letters: str) -> int:
    number = 0
    for ch in letters:
        number = number * 26 + ord(ch) - 64
    return number


def _split_ref(ref: str) -> Tuple[int, int]:
    match = _CELL_REF.match(ref.upper())
    if not match:
        raise XlsxPatchError(f"セル番地が不正です: {ref}")
    return int(match.group(2)), _col_number(match.group(1))


class XlsxPackage:
    def __init__(self, path: Path):
        with zipfile.ZipFile(path) as zf:
            self._order = zf.namelist()
            self._parts: Dict[str, bytes] = {name: zf.read(name) for name in self._order}
        self._trees: Dict[str, etree._ElementTree] = {}
        self._text_styles: Dict[str, str] = {}

    # ---------- 基本操作 ----------

    def _tree(self, part: str) -> etree._Element:
        if part not in self._trees:
            if part not in self._parts:
                raise XlsxPatchError(f"パーツがありません: {part}")
            parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
            self._trees[part] = etree.fromstring(self._parts[part], parser).getroottree()
        return self._trees[part].getroot()

    def _remove_part(self, part: str) -> None:
        self._parts.pop(part, None)
        self._trees.pop(part, None)
        if part in self._order:
            self._order.remove(part)

    def _rels(self, part: str) -> Optional[etree._Element]:
        path = _rels_path(part)
        return self._tree(path) if path in self._parts else None

    def _content_override(self, part: str) -> Optional[etree._Element]:
        types = self._tree('[Content_Types].xml')
        for override in types.findall(f'{{{NS_CT}}}Override'):
            if override.get('PartName') == '/' + part:
                return override
        return None

    # ---------- シート ----------

    def _workbook_part(self) -> str:
        for rel in self._tree('_rels/.rels'):
            if rel.get('Type', '').endswith('/officeDocument'):
                return _resolve('', rel.get('Target'))
        raise XlsxPatchError('workbook.xml が見つかりません')

    def _sheets(self):
        workbook = self._workbook_part()
        rels = {rel.get('Id'): rel for rel in self._rels(workbook)}
        for sheet in self._tree(workbook).find(f'{M}sheets'):
            rel = rels[sheet.get(f'{{{NS_REL}}}id')]
            yield sheet, _resolve(workbook, rel.get('Target'))

    @property
    def sheet_names(self):
        return [sheet.get('name') for sheet, _ in self._sheets()]

    def find_sheet(self, name: str) -> Optional[str]:
        """Excel と同じく大文字・小文字を区別せずにシートを探し、実際のシート名を返す"""
        for actual in self.sheet_names:
            if actual.casefold() == name.casefold():
                return actual
        return None

    def sheet_part(self, name: str) -> str:
        actual = self.find_sheet(name)
        for sheet, part in self._sheets():
            if sheet.get('name') == actual:
                return part
        raise XlsxPatchError(f"シートがありません: {name}")

    def rename_sheet(self, old_name: str, new_name: str) -> None:
        """シート名を変え、印刷範囲などの名前定義・数式・文書プロパティ内の参照も合わせて直す"""
        actual = self.find_sheet(old_name)
        if not actual:
            raise XlsxPatchError(f"シートがありません: {old_name}")
        existing = self.find_sheet(new_name)
        if existing and existing != actual:
            raise XlsxPatchError(f"同名のシートがあります: {new_name}")

        def replace_refs(text: str) -> str:
            # 'シート名'!A1 と シート名!A1 の両方の書き方がある。新しい名前は常に引用符で囲む
            return text.replace(f"'{actual}'!", f"'{new_name}'!").replace(f"{actual}!", f"'{new_name}'!")

        workbook = self._workbook_part()
        for sheet, part in list(self._sheets()):
            if sheet.get('name') == actual:
                sheet.set('name', new_name)
            for formula in self._tree(part).iter(f'{M}f'):
                if formula.text and actual in formula.text:
                    formula.text = replace_refs(formula.text)
        defined_names = self._tree(workbook).find(f'{M}definedNames')
        for defined in defined_names if defined_names is not None else ():
            if defined.text and actual in defined.text:
                defined.text = replace_refs(defined.text)
        if 'docProps/app.xml' in self._parts:
            for el in self._tree('docProps/app.xml').iter():
                if isinstance(el.tag, str) and etree.QName(el).localname == 'lpstr' and el.text == actual:
                    el.text = new_name

    # ---------- セル ----------

    def _text_style(self, style_index: Optional[str]) -> str:
        """既存の書式を元に、表示形式だけ「@」にした書式を作る"""
        key = style_index or '0'
        if key not in self._text_styles:
            styles_part = 'xl/styles.xml'
            cell_xfs = self._tree(styles_part).find(f'{M}cellXfs')
            xf = deepcopy(cell_xfs[int(key)])
            xf.set('numFmtId', TEXT_FORMAT_ID)
            xf.set('applyNumberFormat', '1')
            cell_xfs.append(xf)
            cell_xfs.set('count', str(len(cell_xfs)))
            self._text_styles[key] = str(len(cell_xfs) - 1)
        return self._text_styles[key]

    def _cell(self, sheet_data: etree._Element, ref: str) -> etree._Element:
        row_no, col_no = _split_ref(ref)
        row = None
        for candidate in sheet_data.findall(f'{M}row'):
            r = int(candidate.get('r'))
            if r == row_no:
                row = candidate
                break
            if r > row_no:
                row = etree.Element(f'{M}row', r=str(row_no))
                candidate.addprevious(row)
                break
        if row is None:
            row = etree.SubElement(sheet_data, f'{M}row', r=str(row_no))
        row.attrib.pop('spans', None)   # 列範囲のヒント。書き込み範囲が変わるので外す

        for cell in row.findall(f'{M}c'):
            if cell.get('r') == ref:
                return cell
            if _split_ref(cell.get('r'))[1] > col_no:
                new_cell = etree.Element(f'{M}c', r=ref)
                cell.addprevious(new_cell)
                return new_cell
        return etree.SubElement(row, f'{M}c', r=ref)

    def set_cells(self, sheet_name: str, values: Dict[str, object], text_refs=()) -> None:
        part = self.sheet_part(sheet_name)
        sheet_data = self._tree(part).find(f'{M}sheetData')
        text_refs = set(text_refs)
        for ref, value in values.items():
            cell = self._cell(sheet_data, ref.upper())
            formula = cell.find(f'{M}f')
            if formula is not None and formula.get('t') == 'shared' and formula.get('ref'):
                raise XlsxPatchError(f"共有数式の基準セルは上書きできません: {sheet_name}!{ref}")
            for child in list(cell):
                cell.remove(child)
            cell.attrib.pop('t', None)
            if ref in text_refs:
                cell.set('s', self._text_style(cell.get('s')))

            if value is None or value == '':
                continue
            if isinstance(value, bool):
                cell.set('t', 'b')
                etree.SubElement(cell, f'{M}v').text = '1' if value else '0'
            elif isinstance(value, (int, float)) and ref not in text_refs:
                etree.SubElement(cell, f'{M}v').text = repr(value) if isinstance(value, float) else str(value)
            else:
                cell.set('t', 'inlineStr')
                text = etree.SubElement(etree.SubElement(cell, f'{M}is'), f'{M}t')
                text.text = str(value)
                text.set(f'{{{NS_XML}}}space', 'preserve')

    # ---------- 保存 ----------

    def _drop_calc_chain(self) -> None:
        """数式の計算順序キャッシュ。シート追加で古くなるので消して Excel に作り直させる"""
        workbook = self._workbook_part()
        wb_rels = self._rels(workbook)
        for rel in list(wb_rels):
            if rel.get('Type') == REL_CALC_CHAIN:
                part = _resolve(workbook, rel.get('Target'))
                wb_rels.remove(rel)
                override = self._content_override(part)
                if override is not None:
                    override.getparent().remove(override)
                self._remove_part(part)
        root = self._tree(workbook)
        calc = root.find(f'{M}calcPr')
        if calc is None:
            # スキーマ上の要素順を守る（順序違いは Excel の「修復」対象になる）
            calc = etree.Element(f'{M}calcPr')
            following = next((child for child in root if etree.QName(child).localname in _AFTER_CALC_PR), None)
            if following is not None:
                following.addprevious(calc)
            else:
                root.append(calc)
        calc.set('fullCalcOnLoad', '1')

    def save(self, path: Path) -> None:
        self._drop_calc_chain()
        for part, tree in self._trees.items():
            self._parts[part] = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
        order = ['[Content_Types].xml'] + [p for p in self._order if p != '[Content_Types].xml']
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for part in order:
                zf.writestr(part, self._parts[part])
