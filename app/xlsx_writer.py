"""Audit Fluxo XLSX snapshot export. Python standard library only.

Values, not executable formulas, macros or external connections. Integer cents
are authoritative. Located records are distinct from human-selected records.
"""
import datetime as dt
import os
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
BAD_XML = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff]')
STATUS = {
    'pendente': 'Pendente', 'em_conferencia': 'Em conferência',
    'pronto': 'Conferido', 'conferido': 'Conferido',
    'nao_encontrado': 'Sem correspondência',
    'conferir': 'Conferir', 'encontrado': 'Encontrado',
}
SUMMARY_HEADERS = [
    'Data da solicitação', 'CÓD_SIST', 'Nota fiscal', 'Cliente / fornecedor',
    'Filial solicitada', 'Valor da nota (R$)', 'Qtde. baixas do tipo aceito',
    'Primeira baixa', 'Última baixa', 'Líquido localizado (R$)',
    'Selecionado na conferência (R$)', 'Banco / conta encontrados',
    'Resultado da busca', 'Conferência', 'Observações / justificativa',
    'Alertas', 'Arquivo da solicitação', 'Aba da solicitação', 'Linha de origem',
]
DETAIL_HEADERS = [
    'ID da solicitação', 'CÓD_SIST solicitado', 'Nota fiscal',
    'Cliente / fornecedor', 'Filial solicitada', 'Filial da baixa',
    'CÓD_SIST da baixa', 'Tipo de documento', 'CÓD_TÍTULO', 'Data da baixa',
    'Valor da baixa (R$)', 'Juros (R$)', 'Descontos (R$)', 'Multa (R$)',
    'Valor líquido (R$)', 'CÓD_CONTA', 'Banco / conta original',
    'Selecionada na conferência', 'Tipo aceito', 'Permuta / adiantamento',
    'Documento de permuta', 'Título de permuta', 'Borderô', 'Tipo de borderô',
    'Arquivo da base', 'Aba da base', 'Linha de origem', 'ID da baixa', 'Alertas',
]


def _xml(element):
    return ET.tostring(element, encoding='utf-8', xml_declaration=True)


def _text(value):
    return '' if value is None else BAD_XML.sub('', str(value))[:32767]


def _pick(item, *keys):
    for key in keys:
        value = item.get(key)
        if value is not None and value != '':
            return value
    return None


def _truth(value):
    if isinstance(value, bool):
        return value
    return str(value or '').strip().lower() in ('1', 'true', 'sim', 's', 'yes')


def _cents(item, key='amount'):
    name = key + '_cents'
    if name in item:
        value = item[name]
        if value is None or value == '':
            return None
        number = Decimal(str(value))
        if not number.is_finite() or number != number.to_integral_value():
            raise ValueError('O valor em centavos deve ser um número inteiro.')
        return int(number)
    value = item.get(key)
    if value is None or value == '':
        return None
    raw = str(value).strip().replace('R$', '').replace(' ', '').replace('\u00a0', '')
    negative = raw.startswith('(') and raw.endswith(')')
    raw = raw.strip('()')
    if ',' in raw:
        raw = raw.replace('.', '').replace(',', '.')
    try:
        number = Decimal(raw)
        if not number.is_finite():
            raise ValueError('Valor monetário inválido na exportação.')
        result = int((number * 100).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
        return -result if negative else result
    except InvalidOperation as exc:
        raise ValueError('Valor monetário inválido na exportação.') from exc


def _date(value):
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    raw = str(value or '').strip()
    if not raw:
        return None
    try:
        return dt.datetime.fromisoformat(raw.replace('Z', '+00:00')).date()
    except ValueError:
        pass
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%d/%m/%y'):
        try:
            return dt.datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    return None


def _col(index):
    result = ''
    while index:
        index, digit = divmod(index - 1, 26)
        result = chr(65 + digit) + result
    return result


def _warnings(*groups):
    labels = []
    for group in groups:
        if isinstance(group, str):
            group = [group]
        for warning in group or []:
            if isinstance(warning, dict):
                label = warning.get('label') or warning.get('message') or warning.get('code')
            else:
                label = warning
            label = _text(label).strip()
            if label and label not in labels:
                labels.append(label)
    return ' '.join(labels)


def _origin(item):
    original = item.get('origin') or {}
    return (
        _pick(item, 'source_file', 'file_name') or original.get('file', ''),
        _pick(item, 'source_sheet', 'sheet') or original.get('sheet', ''),
        _pick(item, 'source_row', 'origin_row') or original.get('row', ''),
    )


def _identity(match):
    known = _pick(match, 'id', 'payment_id')
    if known is not None:
        return ('id', str(known))
    origin = _origin(match)
    if all(origin):
        return ('origin',) + tuple(str(v) for v in origin)
    return ('record',) + tuple(_text(match.get(key)) for key in (
        'branch', 'code', 'title', 'document_type', 'date', 'payment_date',
        'amount_cents', 'amount', 'account_code', 'account', 'exchange',
        'related_code', 'related_title',
    ))


def _eligible(match, config):
    if 'eligible' in match:
        return _truth(match['eligible'])
    expected = config.get('document_type')
    if expected is None or expected == '':
        return True
    return str(match.get('document_type') or '').strip() == str(expected).strip()


def _selected(match, request):
    if 'selected' in match:
        return _truth(match['selected'])
    identities = request.get('selected_match_ids', request.get('confirmed_match_ids', [])) or []
    if isinstance(identities, str):
        identities = []
    current = _pick(match, 'id', 'payment_id')
    return current is not None and str(current) in {str(value) for value in identities}


def _sum_money(matches):
    if not matches:
        return None
    values = [_cents(match) for match in matches]
    return None if any(value is None for value in values) else sum(values)


def _build_rows(items, context):
    config = context.get('config') or {}
    summary = []
    detail = []
    globally_eligible = {}
    globally_selected = {}
    without = 0
    multiple = 0
    for index, item in enumerate(items, 1):
        request = item.get('request') if isinstance(item.get('request'), dict) else item
        matches = item.get('matches') or request.get('matches') or []
        unique = {}
        for match in matches:
            unique.setdefault(_identity(match), match)
        matches = list(unique.values())
        eligible = [match for match in matches if _eligible(match, config)]
        selected = [match for match in eligible if _selected(match, request)]
        without += not bool(eligible)
        multiple += len(eligible) > 1
        for match in eligible:
            identity = _identity(match)
            if identity in globally_eligible and _cents(globally_eligible[identity]) != _cents(match):
                raise ValueError('Uma mesma baixa tem valores diferentes no retrato exportado.')
            globally_eligible[identity] = match
        for match in selected:
            globally_selected[_identity(match)] = match
        dates = sorted(value for value in (
            _date(_pick(match, 'payment_date', 'date')) for match in eligible
        ) if value is not None)
        accounts = list(dict.fromkeys(
            _text(match.get('account')).strip() or 'Conta não informada'
            for match in eligible
        ))
        alert = _warnings(item.get('warnings'), request.get('warnings'))
        match_status = request.get('match_status')
        if not eligible:
            label = 'Sem correspondência'
        elif match_status:
            label = STATUS.get(match_status, _text(match_status))
        else:
            label = 'Conferir' if alert or len(eligible) > 1 else 'Encontrado'
        notes = ' '.join(_text(request.get(key)).strip() for key in (
            'notes', 'justification',
        ) if _text(request.get(key)).strip())
        request_origin = _origin(request)
        summary.append([
            ('date', request.get('received_date')), ('text', request.get('code')),
            ('text', request.get('invoice')), ('text', request.get('supplier')),
            ('text', request.get('branch')), ('money', _cents(request)),
            ('count', len(eligible)), ('date', dates[0] if dates else None),
            ('date', dates[-1] if dates else None), ('money', _sum_money(eligible)),
            ('money', _sum_money(selected)), ('text', '\n'.join(accounts)),
            ('text', label), ('text', STATUS.get(request.get('status'), _text(request.get('status')))),
            ('text', notes), ('text', alert), ('text', request_origin[0]),
            ('text', request_origin[1]), ('count', request_origin[2]),
        ])
        request_id = request.get('id') or request.get('request_id') or str(index)
        for match in matches:
            origin = _origin(match)
            detail.append([
                ('text', request_id), ('text', request.get('code')), ('text', request.get('invoice')),
                ('text', request.get('supplier')), ('text', request.get('branch')),
                ('text', match.get('branch')), ('text', _pick(match, 'code', 'document_code')),
                ('text', match.get('document_type')), ('text', _pick(match, 'title', 'title_code')),
                ('date', _pick(match, 'payment_date', 'date')),
                ('money', _cents(match, 'gross_amount')), ('money', _cents(match, 'interest')),
                ('money', _cents(match, 'discount')), ('money', _cents(match, 'fine')),
                ('money', _cents(match)), ('text', match.get('account_code')),
                ('text', match.get('account')), ('text', 'Sim' if _selected(match, request) else 'Não'),
                ('text', 'Sim' if _eligible(match, config) else 'Não'),
                ('text', match.get('exchange')), ('text', match.get('related_code')),
                ('text', match.get('related_title')), ('text', match.get('bordero')),
                ('text', match.get('bordero_type')), ('text', origin[0]), ('text', origin[1]),
                ('count', origin[2]), ('text', _pick(match, 'id', 'payment_id')),
                ('text', _warnings(match.get('warnings'))),
            ])
    metrics = {
        'requests': len(items), 'located': len(globally_eligible), 'without': without,
        'multiple': multiple, 'located_cents': _sum_money(list(globally_eligible.values())),
        'selected_cents': _sum_money(list(globally_selected.values())),
        'associations': len(detail),
    }
    return summary, detail, metrics


class _Sheet:
    def __init__(self, name, headers, widths):
        self.name = name
        self.headers = headers
        self.widths = widths
        self.rows = {}
        self.heights = {}
        self.merges = []
        self.header_row = 12

    def put(self, row, col, kind, value, style=None):
        self.rows.setdefault(row, {})[col] = (kind, value, style)

    def line(self, row, text, style=8, height=30):
        self.put(row, 1, 'text', text, style)
        self.merges.append('A{}:{}{}'.format(row, _col(len(self.headers)), row))
        self.heights[row] = height

    def table(self, records):
        for col, header in enumerate(self.headers, 1):
            self.put(self.header_row, col, 'text', header, 3)
        self.heights[self.header_row] = 38
        for rownum, values in enumerate(records, self.header_row + 1):
            for col, (kind, value) in enumerate(values, 1):
                self.put(rownum, col, kind, value)
            maxlines = max((sum(max(1, (len(segment) + int(self.widths[col - 1]) - 1)
                                       // int(self.widths[col - 1]))
                               for segment in _text(value).split('\n'))
                            for col, (kind, value) in enumerate(values, 1)
                            if kind == 'text'), default=1)
            self.heights[rownum] = min(150, max(30, 14 * maxlines + 8))

    def as_xml(self):
        ncols = len(self.headers)
        maxrow = max(self.rows, default=self.header_row)
        root = ET.Element('worksheet', {'xmlns': MAIN, 'xmlns:r': REL})
        properties = ET.SubElement(root, 'sheetPr')
        ET.SubElement(properties, 'tabColor', {'rgb': 'FFD6074C'})
        ET.SubElement(properties, 'pageSetUpPr', {'fitToPage': '1'})
        ET.SubElement(root, 'dimension', {'ref': 'A1:{}{}'.format(_col(ncols), maxrow)})
        views = ET.SubElement(root, 'sheetViews')
        view = ET.SubElement(views, 'sheetView', {
            'workbookViewId': '0', 'showGridLines': '0', 'zoomScale': '85',
        })
        ET.SubElement(view, 'pane', {
            'xSplit': '2', 'ySplit': str(self.header_row), 'topLeftCell': 'C13',
            'activePane': 'bottomRight', 'state': 'frozen',
        })
        ET.SubElement(view, 'selection', {'pane': 'bottomRight', 'activeCell': 'C13', 'sqref': 'C13'})
        ET.SubElement(root, 'sheetFormatPr', {'defaultRowHeight': '24'})
        columns = ET.SubElement(root, 'cols')
        for col, width in enumerate(self.widths, 1):
            ET.SubElement(columns, 'col', {
                'min': str(col), 'max': str(col), 'width': str(width), 'customWidth': '1',
            })
        data = ET.SubElement(root, 'sheetData')
        for rownum in sorted(self.rows):
            attrs = {'r': str(rownum)}
            if rownum in self.heights:
                attrs.update(ht=str(self.heights[rownum]), customHeight='1')
            row = ET.SubElement(data, 'row', attrs)
            for col, (kind, value, style) in sorted(self.rows[rownum].items()):
                if style is None:
                    style = {'text': 0, 'count': 4, 'money': 5, 'date': 6}.get(kind, 0)
                attrs = {'r': '{}{}'.format(_col(col), rownum), 's': str(style)}
                if value is None or value == '':
                    ET.SubElement(row, 'c', attrs)
                    continue
                parsed = _date(value) if kind == 'date' else None
                if kind == 'money':
                    attrs['t'] = 'n'
                    cell = ET.SubElement(row, 'c', attrs)
                    ET.SubElement(cell, 'v').text = format(Decimal(int(value)) / 100, 'f')
                elif kind == 'count' and str(value).isdigit():
                    attrs['t'] = 'n'
                    cell = ET.SubElement(row, 'c', attrs)
                    ET.SubElement(cell, 'v').text = str(int(value))
                elif kind == 'date' and parsed:
                    attrs['t'] = 'n'
                    cell = ET.SubElement(row, 'c', attrs)
                    serial = (parsed - dt.date(1899, 12, 30)).days
                    if parsed < dt.date(1900, 3, 1):
                        serial -= 1
                    ET.SubElement(cell, 'v').text = str(serial)
                else:
                    attrs['t'] = 'inlineStr'
                    cell = ET.SubElement(row, 'c', attrs)
                    inline = ET.SubElement(cell, 'is')
                    ET.SubElement(inline, 't', {XML_SPACE: 'preserve'}).text = _text(value)
        ET.SubElement(root, 'autoFilter', {'ref': 'A{}:{}{}'.format(
            self.header_row, _col(ncols), max(maxrow, self.header_row))})
        if self.merges:
            merges = ET.SubElement(root, 'mergeCells', {'count': str(len(self.merges))})
            for reference in self.merges:
                ET.SubElement(merges, 'mergeCell', {'ref': reference})
        ET.SubElement(root, 'printOptions', {'horizontalCentered': '1'})
        ET.SubElement(root, 'pageMargins', {
            'left': '0.3', 'right': '0.3', 'top': '0.5', 'bottom': '0.5',
            'header': '0.2', 'footer': '0.2',
        })
        ET.SubElement(root, 'pageSetup', {
            'paperSize': '8', 'orientation': 'landscape', 'fitToWidth': '1', 'fitToHeight': '0',
        })
        footer = ET.SubElement(root, 'headerFooter')
        ET.SubElement(footer, 'oddFooter').text = 'Audit Fluxo     Página &P de &N'
        return _xml(root)


def _styles():
    root = ET.Element('styleSheet', {'xmlns': MAIN})
    formats = ET.SubElement(root, 'numFmts', {'count': '2'})
    ET.SubElement(formats, 'numFmt', {
        'numFmtId': '164', 'formatCode': '"R$" #,##0.00;[Red]("R$" #,##0.00);"R$" 0.00',
    })
    ET.SubElement(formats, 'numFmt', {'numFmtId': '165', 'formatCode': 'dd/mm/yyyy'})
    fonts = ET.SubElement(root, 'fonts', {'count': '5'})
    for size, color, bold, italic in (
        (10, '263238', False, False), (16, '172F46', True, False),
        (10, 'FFFFFF', True, False), (10, '172F46', True, False),
        (10, '59666E', False, True),
    ):
        font = ET.SubElement(fonts, 'font')
        ET.SubElement(font, 'sz', {'val': str(size)})
        ET.SubElement(font, 'color', {'rgb': 'FF' + color})
        ET.SubElement(font, 'name', {'val': 'Arial'})
        if bold:
            ET.SubElement(font, 'b')
        if italic:
            ET.SubElement(font, 'i')
    fills = ET.SubElement(root, 'fills', {'count': '5'})
    for pattern, color in (('none', None), ('gray125', None), ('solid', '172F46'),
                           ('solid', 'F2F5F7'), ('solid', 'D6074C')):
        fill = ET.SubElement(fills, 'fill')
        patternfill = ET.SubElement(fill, 'patternFill', {'patternType': pattern})
        if color:
            ET.SubElement(patternfill, 'fgColor', {'rgb': 'FF' + color})
            ET.SubElement(patternfill, 'bgColor', {'indexed': '64'})
    borders = ET.SubElement(root, 'borders', {'count': '3'})
    for which in ('none', 'rule', 'header'):
        border = ET.SubElement(borders, 'border')
        for edge in ('left', 'right', 'top', 'bottom', 'diagonal'):
            attrs = {'style': 'thin'} if (which == 'rule' and edge == 'bottom') or (
                which == 'header' and edge == 'right') else {}
            part = ET.SubElement(border, edge, attrs)
            if attrs:
                ET.SubElement(part, 'color', {'rgb': 'FFD6074C' if which == 'rule' else 'FFFFFFFF'})
    base = ET.SubElement(root, 'cellStyleXfs', {'count': '1'})
    ET.SubElement(base, 'xf', {'numFmtId': '0', 'fontId': '0', 'fillId': '0', 'borderId': '0'})
    # 0 body; 1 title; 2 context; 3 header; 4 integer; 5 BRL; 6 date;
    # 7 KPI label; 8 explanatory line; 9 KPI count; 10 KPI BRL; 11 brand rule.
    specs = (
        (0, 0, 0, 0, 'left', True), (0, 1, 0, 0, 'left', False),
        (0, 4, 0, 0, 'left', False), (0, 2, 2, 2, 'center', True),
        (3, 0, 0, 0, 'right', False), (164, 0, 0, 0, 'right', False),
        (165, 0, 0, 0, 'right', False), (0, 3, 3, 0, 'left', True),
        (0, 4, 0, 0, 'left', True), (3, 3, 3, 0, 'left', False),
        (164, 3, 3, 0, 'left', False), (0, 0, 0, 1, 'left', False),
    )
    cells = ET.SubElement(root, 'cellXfs', {'count': str(len(specs))})
    for numfmt, font, fill, border, horizontal, wrap in specs:
        xf = ET.SubElement(cells, 'xf', {
            'numFmtId': str(numfmt), 'fontId': str(font), 'fillId': str(fill),
            'borderId': str(border), 'xfId': '0', 'applyAlignment': '1',
            'applyNumberFormat': '1', 'applyFill': '1',
        })
        ET.SubElement(xf, 'alignment', {
            'horizontal': horizontal, 'vertical': 'center', 'wrapText': '1' if wrap else '0',
        })
    styles = ET.SubElement(root, 'cellStyles', {'count': '1'})
    ET.SubElement(styles, 'cellStyle', {'name': 'Normal', 'xfId': '0', 'builtinId': '0'})
    return _xml(root)


def _assemble(sheets, timestamp):
    workbook = ET.Element('workbook', {'xmlns': MAIN, 'xmlns:r': REL})
    ET.SubElement(workbook, 'workbookPr', {'date1904': '0'})
    views = ET.SubElement(workbook, 'bookViews')
    ET.SubElement(views, 'workbookView', {'activeTab': '0'})
    sheetlist = ET.SubElement(workbook, 'sheets')
    for index, sheet in enumerate(sheets, 1):
        ET.SubElement(sheetlist, 'sheet', {
            'name': sheet.name, 'sheetId': str(index), 'r:id': 'rId{}'.format(index),
        })
    names = ET.SubElement(workbook, 'definedNames')
    for index, sheet in enumerate(sheets):
        ET.SubElement(names, 'definedName', {
            'name': '_xlnm.Print_Titles', 'localSheetId': str(index),
        }).text = ("'{}'!".format(sheet.name) + '$' + str(sheet.header_row)
                   + ':$' + str(sheet.header_row))
    relationships = ET.Element('Relationships', {'xmlns': PKG})
    for index in range(1, len(sheets) + 1):
        ET.SubElement(relationships, 'Relationship', {
            'Id': 'rId{}'.format(index), 'Type': REL + '/worksheet',
            'Target': 'worksheets/sheet{}.xml'.format(index),
        })
    ET.SubElement(relationships, 'Relationship', {
        'Id': 'rId{}'.format(len(sheets) + 1), 'Type': REL + '/styles', 'Target': 'styles.xml',
    })
    package = ET.Element('Relationships', {'xmlns': PKG})
    for ident, typ, target in (
        ('rId1', REL + '/officeDocument', 'xl/workbook.xml'),
        ('rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
         'docProps/core.xml'),
        ('rId3', REL + '/extended-properties', 'docProps/app.xml'),
    ):
        ET.SubElement(package, 'Relationship', {'Id': ident, 'Type': typ, 'Target': target})
    content = ET.Element('Types', {'xmlns': 'http://schemas.openxmlformats.org/package/2006/content-types'})
    for ext, mimetype in (('rels', 'application/vnd.openxmlformats-package.relationships+xml'),
                           ('xml', 'application/xml')):
        ET.SubElement(content, 'Default', {'Extension': ext, 'ContentType': mimetype})
    for part, mimetype in (
        ('/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
        ('/xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'),
        ('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'),
        ('/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml'),
    ):
        ET.SubElement(content, 'Override', {'PartName': part, 'ContentType': mimetype})
    for index in range(1, len(sheets) + 1):
        ET.SubElement(content, 'Override', {
            'PartName': '/xl/worksheets/sheet{}.xml'.format(index),
            'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
        })
    core = ET.Element('cp:coreProperties', {
        'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
        'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
        'xmlns:dcterms': 'http://purl.org/dc/terms/',
        'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    })
    ET.SubElement(core, 'dc:creator').text = 'Audit Fluxo'
    ET.SubElement(core, 'dc:title').text = 'Atendimento à auditoria'
    ET.SubElement(core, 'dcterms:created', {'xsi:type': 'dcterms:W3CDTF'}).text = timestamp
    app = ET.Element('Properties', {
        'xmlns': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    })
    ET.SubElement(app, 'Application').text = 'Audit Fluxo'
    ET.SubElement(app, 'AppVersion').text = '1.0'
    result = {
        '[Content_Types].xml': _xml(content), '_rels/.rels': _xml(package),
        'xl/workbook.xml': _xml(workbook), 'xl/_rels/workbook.xml.rels': _xml(relationships),
        'xl/styles.xml': _styles(), 'docProps/core.xml': _xml(core), 'docProps/app.xml': _xml(app),
    }
    for index, sheet in enumerate(sheets, 1):
        result['xl/worksheets/sheet{}.xml'.format(index)] = sheet.as_xml()
    return result


def write_audit_workbook(path, items, context=None):
    """Write an XLSX snapshot and return the absolute output path.

    Supports flattened engine.request_detail dictionaries and {"request": ...,
    "matches": ...}. Cents are authoritative. Replaces the destination atomically.
    """
    context = dict(context or {})
    items = list(items)
    summary_rows, detail_rows, metrics = _build_rows(items, context)
    generated = context.get('generated_at') or dt.datetime.now().isoformat(timespec='seconds')
    subtitle = 'Lote: {}    Auditor: {}    Gerado: {}'.format(
        _text(context.get('batch_name') or 'Todos os documentos'),
        _text(context.get('auditor') or 'Não informado'), _text(generated),
    )
    summary = _Sheet('Resumo', SUMMARY_HEADERS,
                     [16, 14, 15, 38, 15, 22, 18, 16, 16, 24, 27, 40, 23, 22, 48, 65, 36, 26, 16])
    detail = _Sheet('Detalhamento', DETAIL_HEADERS,
                    [24, 17, 15, 38, 15, 15, 16, 15, 15, 16, 22, 20, 20, 20, 23, 15, 40,
                     20, 15, 20, 19, 19, 15, 15, 36, 26, 16, 24, 65])
    for sheet in (summary, detail):
        sheet.line(2, 'AUDIT FLUXO · AMOSTRA FICTÍCIA', 1, 28)
        sheet.line(3, subtitle, 2, 24)
        sheet.line(4, '', 11, 5)
    kpis = [
        (1, 3, 'Documentos solicitados', 'count', metrics['requests']),
        (4, 6, 'Baixas localizadas (sem duplicação)', 'count', metrics['located']),
        (7, 9, 'Sem correspondência', 'count', metrics['without']),
        (10, 12, 'Documentos com múltiplas baixas', 'count', metrics['multiple']),
        (13, 15, 'Líquido localizado (R$)', 'money', metrics['located_cents']),
        (16, 19, 'Selecionado na conferência (R$)', 'money', metrics['selected_cents']),
    ]
    for left, right, label, kind, value in kpis:
        summary.put(5, left, 'text', label, 7)
        summary.put(6, left, kind, value, 10 if kind == 'money' else 9)
        summary.merges.extend(['{}5:{}5'.format(_col(left), _col(right)),
                               '{}6:{}6'.format(_col(left), _col(right))])
    summary.heights.update({5: 32, 6: 28})
    summary.line(8, 'Valores de baixas são registros contábeis localizados. A conferência não comprova a existência do comprovante.', 8, 30)
    summary.line(9, 'Totais globais contam cada baixa uma vez. O mesmo registro pode aparecer em mais de uma solicitação no detalhamento.', 8, 30)
    expected = (context.get('config') or {}).get('document_type')
    summary.line(10, 'Critério: filial + CÓD_SIST; sem filial, apresentar candidatos por CÓD_SIST. Tipo de documento aceito: {}.'.format(
        expected if expected is not None and expected != '' else 'conforme importação'), 8, 30)
    detail.line(6, 'Linhas vinculadas a solicitações: {}. Registros do tipo não aceito permanecem visíveis e ficam fora dos totais.'.format(
        metrics['associations']), 8, 30)
    detail.line(8, 'Banco / conta conserva a descrição original. Contas transitórias, permutas e outras contas exigem conferência.', 8, 30)
    detail.line(9, 'Selecionada registra a escolha humana na data da exportação. Origem identifica arquivo, aba e linha da base.', 8, 30)
    source_names = context.get('source_names') or []
    if isinstance(source_names, str):
        source_names = [source_names]
    detail.line(10, 'Fontes: ' + ', '.join(_text(value) for value in source_names), 8, 30)
    summary.table(summary_rows)
    detail.table(detail_rows)
    timestamp = dt.datetime.now(dt.timezone.utc).isoformat(timespec='seconds').replace('+00:00', 'Z')
    parts = _assemble([summary, detail], timestamp)
    target = Path(path).expanduser().resolve()
    if target.suffix.lower() != '.xlsx':
        raise ValueError('A exportação deve usar a extensão .xlsx.')
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(prefix='.audit-fluxo-', suffix='.tmp',
                                         dir=target.parent, delete=False) as handle:
            temporary = Path(handle.name)
        with zipfile.ZipFile(temporary, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
        os.replace(str(temporary), str(target))
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()
    return str(target)
