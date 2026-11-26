"""Small, read-only XLSX reader using the Python standard library.

Formula cells use their saved value; this module never recalculates formulas.
Worksheets are streamed and numeric text is preserved to avoid money rounding.
"""
import datetime as dt
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
DATE_FORMATS = set(range(14, 23)) | {27, 30, 36, 45, 46, 47, 50, 57}


def column_index(reference):
    letters = re.match(r'[A-Za-z]+', reference or '')
    value = 0
    for letter in letters.group(0).upper() if letters else 'A':
        value = value * 26 + ord(letter) - 64
    return value - 1


def column_name(index):
    result = ''
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


class XLSXReader:
    def __init__(self, path):
        self.path = str(path)
        self.archive = zipfile.ZipFile(path)
        if 'xl/workbook.xml' not in self.archive.namelist():
            self.archive.close()
            raise ValueError('O arquivo não é uma planilha XLSX válida.')
        # Reject unexpectedly huge decompressed content before parsing.
        if sum(i.file_size for i in self.archive.infolist()) > 1_200_000_000:
            self.archive.close()
            raise ValueError('A planilha ultrapassa o limite de leitura local.')
        workbook = ET.fromstring(self.archive.read('xl/workbook.xml'))
        props = workbook.find(NS + 'workbookPr')
        self.date1904 = props is not None and props.get('date1904') in ('1', 'true')
        relationships = ET.fromstring(self.archive.read('xl/_rels/workbook.xml.rels'))
        targets = {}
        for item in relationships:
            target = item.get('Target', '')
            target = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
            targets[item.get('Id')] = target
        self.sheets = []
        for item in workbook.find(NS + 'sheets'):
            target = targets.get(item.get(REL + 'id'))
            if target in self.archive.namelist():
                self.sheets.append({'name': item.get('name', ''), 'path': target,
                                    'hidden': item.get('state', 'visible') != 'visible'})
        self.shared = []
        if 'xl/sharedStrings.xml' in self.archive.namelist():
            with self.archive.open('xl/sharedStrings.xml') as handle:
                parser = ET.iterparse(handle, events=('start', 'end'))
                _, shared_root = next(parser)
                for event, elem in parser:
                    if event == 'end' and elem.tag == NS + 'si':
                        self.shared.append(''.join(e.text or '' for e in elem.iter(NS + 't')))
                        shared_root.clear()
        self.date_styles = set()
        if 'xl/styles.xml' in self.archive.namelist():
            styles = ET.fromstring(self.archive.read('xl/styles.xml'))
            custom = {}
            for item in styles.findall(NS + 'numFmts/' + NS + 'numFmt'):
                fmt = item.get('formatCode', '').lower()
                # Remove literal strings and colour/locale annotations.
                fmt = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', '', fmt)
                custom[int(item.get('numFmtId', 0))] = bool(re.search(r'[ydh]|m[/\-]|[/\-]m', fmt))
            xfs = styles.find(NS + 'cellXfs')
            for index, item in enumerate(xfs if xfs is not None else []):
                fmt = int(item.get('numFmtId', 0))
                if fmt in DATE_FORMATS or custom.get(fmt, False):
                    self.date_styles.add(index)

    def close(self):
        self.archive.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()

    def _value(self, cell):
        kind = cell.get('t', '')
        if kind == 'inlineStr':
            return ''.join(e.text or '' for e in cell.iter(NS + 't'))
        item = cell.find(NS + 'v')
        value = item.text if item is not None and item.text is not None else ''
        if kind == 's':
            try:
                return self.shared[int(value)]
            except (ValueError, IndexError):
                return ''
        if kind == 'b':
            return '1' if value == '1' else '0'
        if kind == 'e':
            return ''  # Saved formula error is not usable business data.
        if kind in ('str', 'd'):
            return value
        if value and int(cell.get('s', 0)) in self.date_styles:
            try:
                serial = float(value)
                origin = dt.datetime(1904, 1, 1) if self.date1904 else dt.datetime(1899, 12, 30)
                if not self.date1904 and serial < 60:
                    origin = dt.datetime(1899, 12, 31)
                date = origin + dt.timedelta(days=serial)
                return date.date().isoformat()
            except (ValueError, OverflowError):
                pass
        return value

    def iter_rows(self, sheet_name, max_rows=None):
        sheet = next((s for s in self.sheets if s['name'] == sheet_name), None)
        if sheet is None:
            raise ValueError('A aba selecionada não existe na planilha.')
        emitted = 0
        with self.archive.open(sheet['path']) as handle:
            data_element = None
            for event, elem in ET.iterparse(handle, events=('start', 'end')):
                if event == 'start' and elem.tag == NS + 'sheetData':
                    data_element = elem
                if event != 'end' or elem.tag != NS + 'row':
                    continue
                row = []
                for cell in elem.findall(NS + 'c'):
                    index = column_index(cell.get('r'))
                    if index > 16383:
                        continue
                    if index >= len(row):
                        row.extend([''] * (index + 1 - len(row)))
                    row[index] = self._value(cell)
                number = int(elem.get('r', emitted + 1))
                elem.clear()
                if data_element is not None:
                    data_element.remove(elem)
                if any(str(v).strip() for v in row):
                    emitted += 1
                    yield number, row
                    if max_rows is not None and emitted >= max_rows:
                        break
