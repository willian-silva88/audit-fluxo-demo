"""Local audit workflow. No internet, AI service or third-party dependency.

Amounts are stored as integer cents. Saved spreadsheet formulas are never
trusted as reconciliation logic: matches are calculated from the raw base.
The active payments snapshot is replaced, never accumulated. Reviewed links
do not prove payment or delivery of bank receipts.
"""
import csv
import codecs
import datetime as dt
import hashlib
import json
import re
import sqlite3
import secrets
import threading
import unicodedata
from contextlib import contextmanager
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

from xlsx_reader import XLSXReader, column_name

FIELD_ALIASES = {
    'supplier': ['supplier', 'cliente', 'fornecedor', 'nome fornecedor', 'razao social'],
    'invoice': ['invoice', 'n nf', 'n nota fiscal', 'numero nf', 'nota fiscal', 'nf'],
    'code': ['code', 'cod sist', 'codigo sistema', 'codigo documento', 'cod documento', 'documento'],
    'branch': ['branch', 'filial', 'codigo filial', 'cod filial'],
    'amount': ['amount', 'total nota', 'valor nota', 'valor liquido', 'valor baixa', 'valor'],
    'issue_date': ['issue date', 'emissao', 'data emissao'],
    'process': ['process', 'processo1', 'processo'],
    'company': ['company', 'empresa'],
    'received_date': ['received date', 'data solicitacao', 'data recebimento'],
    'auditor': ['auditor', 'auditoria'],
    'payment_date': ['payment date', 'data baixa', 'data pagamento'],
    'account': ['account', 'descricao conta', 'banco conta', 'conta bancaria', 'banco'],
    'account_code': ['account code', 'codigo conta', 'cod conta'],
    'document_type': ['document type', 'tipo documento', 'tipo doc'],
    'exchange': ['exchange', 'permuta adiantamento', 'permuta'],
    'related_code': ['related code', 'codigo documento permuta', 'documento permuta'],
    'related_title': ['related title', 'codigo titulo permuta'],
    'title': ['title', 'codigo titulo', 'cod titulo', 'titulo'],
    'bordero': ['bordero', 'codigo bordero'],
    'bordero_type': ['bordero type', 'tipo bordero'],
    'gross_amount': ['gross amount', 'valor baixa', 'valor principal', 'principal'],
    'interest': ['interest', 'valor juros', 'juros'],
    'discount': ['discount', 'valor descontos', 'descontos'],
    'fine': ['fine', 'valor multa', 'multa'],
}
REQUEST_FIELDS = ['supplier', 'invoice', 'code', 'branch', 'amount', 'issue_date', 'process', 'company', 'received_date', 'auditor']
PAYMENT_FIELDS = ['code', 'branch', 'amount', 'payment_date', 'account', 'account_code', 'document_type', 'exchange', 'related_code', 'related_title', 'title', 'bordero', 'bordero_type', 'gross_amount', 'interest', 'discount', 'fine']
DEFAULT_CONFIG = {
    'document_type': '2',
    'tolerance': 0,
    'required_documents': [],
    'non_bank_terms': ['transitoria', 'adiantamento', 'fornecedores', 'permuta'],
    'search_folders': [],
    'allowed_roots': [],
    'ai_enabled': False,
    'ai_provider': None,
    'profiles': {},
}
WARNING_LABELS = {
    'missing_code': 'Código do sistema ausente; identificação manual necessária.',
    'missing_branch': 'Filial não informada; os resultados são apenas candidatos.',
    'not_found': 'Nenhuma baixa encontrada com os critérios informados.',
    'multiple': 'Mais de uma baixa encontrada; conferir a composição.',
    'value_diff': 'A soma dos registros difere do valor da nota.',
    'permuta': 'Há registro de permuta ou adiantamento; conferir a composição.',
    'non_bank_account': 'A descrição da conta não identifica um banco de pagamento; confirmar a origem da baixa.',
    'missing_document_type': 'Tipo de documento não informado na baixa.',
    'document_type_diff': 'Tipo de documento fora do critério configurado; candidato não selecionável.',
    'review_invalidated': 'Uma atualização alterou os vínculos ou regras; conferir novamente.',
    'multiple_branches': 'O código foi encontrado em filiais diferentes.',
    'multiple_types': 'Há tipos de documento diferentes entre os registros localizados.',
    'multiple_accounts': 'Os registros localizados utilizam contas diferentes.',
    'duplicate_record': 'Há registros de baixa idênticos; conferir se são duplicações da origem.',
    'missing_payment_date': 'Data de baixa ausente no registro.',
    'net_difference': 'Valor líquido difere do principal com juros, descontos e multa.',
}


def normalize(value):
    text = unicodedata.normalize('NFKD', str(value or ''))
    text = ''.join(c for c in text if not unicodedata.combining(c)).lower()
    return re.sub(r'[^a-z0-9]+', ' ', text).strip()


def identifier(value):
    value = str('' if value is None else value).strip()
    if re.fullmatch(r'\d+\.0+', value):
        value = value.split('.')[0]
    # Numeric codes, unlike invoice display labels, are canonical integer keys.
    if re.fullmatch(r'\d+', value):
        return str(int(value))
    return value


def money(value):
    if value is None or str(value).strip() == '':
        return None
    text = str(value).strip().replace('\u00a0', '').replace(' ', '').replace('R$', '')
    negative = text.startswith('(') and text.endswith(')')
    text = text.strip('()')
    if ',' in text:
        text = text.replace('.', '').replace(',', '.')
    try:
        decimal = Decimal(text)
        if not decimal.is_finite():
            raise ValueError()
        cents = int((decimal * 100).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
        return -cents if negative else cents
    except (InvalidOperation, ValueError, OverflowError):
        raise ValueError('Foi encontrado um valor monetário inválido. Confira o mapeamento das colunas.')


def date_value(value):
    text = str(value or '').strip()
    if not text:
        return ''
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%d/%m/%y', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M:%S'):
        try:
            return dt.datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    if re.fullmatch(r'\d{5}(\.\d+)?', text):
        try:
            return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(text))).date().isoformat()
        except (ValueError, OverflowError):
            pass
    raise ValueError('Foi encontrada uma data inválida. Confira o mapeamento das colunas.')


def now():
    return dt.datetime.now().isoformat(timespec='seconds')


def file_hash(path):
    digest = hashlib.sha256()
    with open(path, 'rb') as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


class CSVReader:
    def __init__(self, path):
        self.path = str(path)
        self.sheets = [{'name': 'Dados', 'hidden': False}]
        with open(path, 'rb') as handle:
            sample = handle.read(65536)
        try:
            codecs.getincrementaldecoder('utf-8-sig')().decode(sample, final=False)
            self.encoding = 'utf-8-sig'
        except UnicodeDecodeError:
            self.encoding = 'cp1252'

    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass

    def iter_rows(self, sheet_name, max_rows=None):
        if sheet_name != 'Dados':
            raise ValueError('A aba selecionada não existe no arquivo.')
        with open(self.path, 'r', encoding=self.encoding, newline='') as handle:
            sample = handle.read(8192)
            handle.seek(0)
            try:
                dialect = csv.Sniffer().sniff('\n'.join(line for line in sample.splitlines() if line.strip()), delimiters=';,\t')
            except csv.Error:
                first_line = next((line for line in sample.splitlines() if line.strip()), '')
                delimiter = max((';', ',', '\t'), key=lambda separator: len(next(csv.reader([first_line], delimiter=separator))))
                dialect = type('LocalDialect', (csv.excel,), {'delimiter': delimiter})
            emitted = 0
            for row_number, row in enumerate(csv.reader(handle, dialect), 1):
                if any(str(v).strip() for v in row):
                    emitted += 1
                    yield row_number, row
                    if max_rows is not None and emitted >= max_rows:
                        break


def reader_for(path):
    extension = Path(path).suffix.lower()
    if extension == '.xlsx':
        return XLSXReader(path)
    if extension == '.csv':
        return CSVReader(path)
    raise ValueError('Envie um arquivo .xlsx ou .csv.')


def guess_mapping(headers, fields=None):
    headers_normalized = [normalize(v) for v in headers]
    found = {}
    for field in fields or FIELD_ALIASES:
        for alias in FIELD_ALIASES[field]:
            alias = normalize(alias)
            if alias in headers_normalized:
                found[field] = headers_normalized.index(alias)
                break
    return found


def analyze_sheet(reader, sheet_name):
    best = None
    rows = list(reader.iter_rows(sheet_name, max_rows=35))
    for number, row in rows:
        mapping = guess_mapping(row)
        p_score = sum(k in mapping for k in ('branch', 'code', 'payment_date', 'document_type', 'title', 'exchange', 'account'))
        r_score = sum(k in mapping for k in ('supplier', 'invoice', 'code', 'amount', 'issue_date', 'process'))
        kind = 'payments' if 'payment_date' in mapping and 'code' in mapping else ('requests' if r_score >= 3 else None)
        score = max(p_score, r_score)
        if best is None or score > best[0]:
            best = (score, number, row, mapping, kind)
    if best is None:
        return {'name': sheet_name, 'header_row': None, 'columns': [], 'mapping': {}, 'suggested_kind': None, 'recognized_fields': [], 'rows_preview': []}
    _, number, headers, mapping, kind = best
    fields = PAYMENT_FIELDS if kind == 'payments' else REQUEST_FIELDS
    return {'name': sheet_name, 'header_row': number,
            'columns': [{'index': i, 'label': str(value), 'letter': column_name(i)} for i, value in enumerate(headers)],
            'mapping': {key: value for key, value in mapping.items() if key in fields},
            'suggested_kind': kind, 'recognized_fields': [key for key in mapping if key in fields], 'rows_preview': [{'row': num, 'values': row} for num, row in rows if num > number][:5]}


class Engine:
    def __init__(self, data_dir):
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.db_path = self.data_dir / 'auditoria.sqlite3'
        self.lock = threading.RLock()
        with self.connection() as conn:
            conn.executescript('''
            CREATE TABLE IF NOT EXISTS config (key TEXT PRIMARY KEY, value TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS sources (id TEXT PRIMARY KEY, kind TEXT, file_name TEXT, file_hash TEXT, sheet TEXT, imported_at TEXT, row_count INTEGER, mapping TEXT, active INTEGER DEFAULT 1);
            CREATE TABLE IF NOT EXISTS batches (id TEXT PRIMARY KEY, source_id TEXT, name TEXT, auditor TEXT, received_date TEXT, created_at TEXT);
            CREATE TABLE IF NOT EXISTS requests (
                id TEXT PRIMARY KEY, batch_id TEXT, source_id TEXT, origin_row INTEGER,
                supplier TEXT, invoice TEXT, code TEXT, branch TEXT, amount INTEGER,
                issue_date TEXT, process TEXT, company TEXT, received_date TEXT, auditor TEXT,
                status TEXT DEFAULT 'pendente', notes TEXT DEFAULT '', justification TEXT DEFAULT '',
                selected_match_ids TEXT DEFAULT '[]', reviewed_at TEXT, raw TEXT,
                review_invalidated INTEGER DEFAULT 0);
            CREATE INDEX IF NOT EXISTS ix_requests_batch ON requests(batch_id);
            CREATE TABLE IF NOT EXISTS payments (
                id TEXT PRIMARY KEY, source_id TEXT, origin_row INTEGER, code TEXT, branch TEXT,
                amount INTEGER, payment_date TEXT, account TEXT, account_code TEXT, document_type TEXT,
                exchange TEXT, related_code TEXT, related_title TEXT, title TEXT, bordero TEXT,
                bordero_type TEXT, gross_amount INTEGER, interest INTEGER, discount INTEGER, fine INTEGER,
                raw TEXT);
            CREATE INDEX IF NOT EXISTS ix_payments_match ON payments(code, branch, document_type);
            CREATE TABLE IF NOT EXISTS attachments (
                id TEXT PRIMARY KEY, request_id TEXT, name TEXT, path TEXT, kind TEXT, created_at TEXT);
            CREATE INDEX IF NOT EXISTS ix_attachments_request ON attachments(request_id);
            CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY AUTOINCREMENT, request_id TEXT, event TEXT, created_at TEXT, details TEXT);
            ''')
            for key, value in DEFAULT_CONFIG.items():
                conn.execute('INSERT OR IGNORE INTO config(key,value) VALUES (?,?)', (key, json.dumps(value, ensure_ascii=False)))

    @contextmanager
    def connection(self):
        conn = sqlite3.connect(str(self.db_path), timeout=60)
        conn.row_factory = sqlite3.Row
        conn.execute('PRAGMA journal_mode=WAL')
        conn.execute('PRAGMA busy_timeout=60000')
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def reset_work(self):
        """Save a consistent local copy before clearing the working records."""
        with self.lock, self.connection() as conn:
            counts = {table: conn.execute("SELECT COUNT(*) FROM " + table).fetchone()[0]
                      for table in ("requests", "payments", "sources", "batches", "history", "attachments")}
            folder = self.data_dir / "copias" / (dt.datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + secrets.token_hex(4))
            folder.mkdir(parents=True, exist_ok=False)
            backup = folder / "auditoria.sqlite3"
            target = sqlite3.connect(str(backup))
            try:
                conn.backup(target)
                if target.execute("PRAGMA quick_check").fetchone()[0] != "ok":
                    raise ValueError("A copia dos dados nao passou na verificacao. O atendimento foi mantido.")
            finally:
                target.close()
            manifest = {"created_at": now(), "reason": "Comecar do zero", "records": counts,
                        "database": backup.name, "original_files_preserved": True}
            (folder / "LEIA-ME.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
            conn.execute("BEGIN IMMEDIATE")
            for table in ("attachments", "history", "requests", "payments", "batches", "sources", "config"):
                conn.execute("DELETE FROM " + table)
            conn.execute("DELETE FROM sqlite_sequence WHERE name='history'")
            for key, value in DEFAULT_CONFIG.items():
                conn.execute("INSERT INTO config(key,value) VALUES (?,?)",
                             (key, json.dumps(value, ensure_ascii=False)))
        return {"cleared": True, "backup": str(backup.resolve()), "previous": counts}

    def _config(self, conn):
        result = {row['key']: json.loads(row['value']) for row in conn.execute('SELECT key,value FROM config')}
        result['ai_enabled'] = False
        result['ai_provider'] = None
        return result

    def get_config(self):
        with self.connection() as conn:
            return self._config(conn)

    def update_config(self, payload):
        with self.lock, self.connection() as conn:
            previous = self._config(conn)
            signatures = self._review_signatures(conn)
            for key, value in payload.items():
                if key not in ('document_type', 'tolerance', 'non_bank_terms', 'profiles'):
                    continue
                if key == 'document_type':
                    value = identifier(value)
                elif key == 'tolerance':
                    cents = money(value)
                    if cents is None or cents < 0:
                        raise ValueError('A tolerância deve ser um valor não negativo.')
                    value = cents / 100
                elif key == 'non_bank_terms':
                    if not isinstance(value, list):
                        raise ValueError('Informe os termos de contas especiais em uma lista.')
                    value = [str(item).strip() for item in value if str(item).strip()]
                elif key == 'profiles' and not isinstance(value, dict):
                    raise ValueError('Os modelos de importação precisam ser um objeto.')
                conn.execute('INSERT OR REPLACE INTO config(key,value) VALUES (?,?)', (key, json.dumps(value, ensure_ascii=False)))
            config = self._config(conn)
            changed = any(config.get(key) != previous.get(key) for key in ('document_type', 'non_bank_terms', 'tolerance'))
            self._invalidate_reviews(conn, signatures, force=changed)
            self._history(conn, None, 'configuration_updated', {'changed_keys': [key for key in config if config[key] != previous.get(key)]})
            return config

    def preview_upload(self, path):
        with reader_for(path) as reader:
            sheets = [analyze_sheet(reader, item['name']) for item in reader.sheets]
            hidden = {item['name']: item.get('hidden', False) for item in reader.sheets}
            for sheet in sheets:
                sheet['hidden'] = hidden[sheet['name']]
        return {'filename': Path(path).name, 'hash': file_hash(path), 'sheets': sheets,
                'fields': {'requests': REQUEST_FIELDS, 'payments': PAYMENT_FIELDS},
                'required_fields': {'requests': ['code', 'amount'], 'payments': ['code', 'branch', 'amount', 'payment_date']},
                'warnings': ['Fórmulas usam os valores salvos; o painel não atualiza consultas externas.']}

    def _prepare_import(self, path, sheet, mapping, kind, metadata):
        preview = self.preview_upload(path)
        if not sheet:
            possibilities = [s for s in preview['sheets'] if s['suggested_kind'] == kind]
            if len(possibilities) == 1:
                sheet = possibilities[0]['name']
            elif len(preview['sheets']) == 1:
                sheet = preview['sheets'][0]['name']
            else:
                raise ValueError('Escolha a aba que contém os dados que deseja importar.')
        selected = next((s for s in preview['sheets'] if s['name'] == sheet), None)
        if not selected or not selected['header_row']:
            raise ValueError('Não foi possível identificar o cabeçalho da aba.')
        fields = REQUEST_FIELDS if kind == 'requests' else PAYMENT_FIELDS
        supplied = dict(mapping if mapping is not None else selected['mapping'])
        header_row = int(metadata.get('header_row') or supplied.pop('header_row', selected['header_row']))
        if header_row < 1:
            raise ValueError('A linha do cabeçalho deve ser positiva.')
        headers = None
        with reader_for(path) as reader:
            for number, row in reader.iter_rows(sheet):
                if number == header_row:
                    headers = row
                    break
                if number > header_row:
                    break
        if headers is None:
            raise ValueError('A linha indicada não contém um cabeçalho.')
        clean = {}
        for field, value in supplied.items():
            if field not in fields or value is None or value == '':
                continue
            if isinstance(value, dict):
                value = value.get('index')
            if isinstance(value, str) and not value.isdigit():
                if value in headers:
                    value = headers.index(value)
                elif re.fullmatch(r'[A-Za-z]{1,3}', value):
                    value = sum((ord(letter) - 64) * 26 ** power for power, letter in enumerate(value.upper()[::-1])) - 1
                else:
                    raise ValueError('Uma coluna selecionada não está disponível no cabeçalho.')
            clean[field] = int(value)
            if clean[field] < 0 or clean[field] >= len(headers):
                raise ValueError('Uma coluna selecionada está fora do cabeçalho.')
        required = ('code', 'amount') if kind == 'requests' else ('code', 'branch', 'amount', 'payment_date')
        missing = [key for key in required if key not in clean]
        if missing:
            raise ValueError('Mapeie os campos obrigatórios: ' + ', '.join(missing) + '.')
        digest = preview['hash']
        settings = {'mapping': clean, 'header_row': header_row}
        source_id = hashlib.sha256(json.dumps([kind, digest, sheet, settings], sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:32]
        return sheet, clean, header_row, headers, digest, source_id

    def _iter_mapped(self, path, sheet, header_row, headers, mapping):
        with reader_for(path) as reader:
            for number, row in reader.iter_rows(sheet):
                if number <= header_row:
                    continue
                values = {key: row[index] if index < len(row) else '' for key, index in mapping.items()}
                if not any(str(v).strip() for v in values.values()):
                    continue
                raw = {str(headers[i]) or column_name(i): str(value) for i, value in enumerate(row) if i < len(headers)}
                yield number, values, raw

    def import_requests(self, path, sheet=None, mapping=None, metadata=None, progress=None):
        metadata = dict(metadata or {})
        sheet, mapping, header_row, headers, digest, source_id = self._prepare_import(path, sheet, mapping, 'requests', metadata)
        with self.lock, self.connection() as conn:
            # Compare file content and reading settings, including sources whose
            # older IDs incorporated reception date or audit batch metadata.
            # Changing today's default date must never create another request set.
            existing = None
            for source in conn.execute("SELECT * FROM sources WHERE kind='requests' AND file_hash=? AND sheet=? ORDER BY imported_at,id", (digest, sheet)):
                saved_mapping = json.loads(source['mapping'] or '{}')
                saved_header = saved_mapping.pop('header_row', header_row)
                if saved_mapping == mapping and int(saved_header) == header_row:
                    existing = source
                    break
            if existing:
                batch = conn.execute('SELECT id FROM batches WHERE source_id=?', (existing['id'],)).fetchone()
                return {'imported': 0, 'existing': existing['row_count'], 'duplicate': True, 'source_id': existing['id'],
                        'batch_id': batch['id'] if batch else existing['id'], 'metadata_preserved': True,
                        'message': 'Esta solicitação já foi importada; dados do atendimento, histórico e decisões anteriores foram preservados.'}
            batch_id = source_id
            stamp = now()
            batch_name = str(metadata.get('batch_name') or Path(str(metadata.get('filename') or path)).stem)
            received = date_value(metadata.get('received_date') or dt.date.today().isoformat())
            auditor = str(metadata.get('auditor') or '')
            conn.execute('INSERT INTO batches VALUES (?,?,?,?,?,?)', (batch_id, source_id, batch_name, auditor, received, stamp))
            count = skipped = 0
            for number, values, raw in self._iter_mapped(path, sheet, header_row, headers, mapping):
                code = identifier(values.get('code'))
                # Ignore footer/heading rows; missing codes are allowed when an invoice identifies a request.
                if not code and not str(values.get('invoice', '')).strip():
                    skipped += 1
                    continue
                amount = money(values.get('amount'))
                if amount is None:
                    skipped += 1
                    continue
                request_id = hashlib.sha256((source_id + ':' + str(number)).encode()).hexdigest()[:32]
                conn.execute('''INSERT INTO requests(id,batch_id,source_id,origin_row,supplier,invoice,code,branch,amount,issue_date,process,company,received_date,auditor,raw)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                    (request_id, batch_id, source_id, number, str(values.get('supplier') or '').strip(),
                     str(values.get('invoice') or '').strip(), code, identifier(values.get('branch')), amount,
                     date_value(values.get('issue_date')), str(values.get('process') or '').strip(),
                     str(values.get('company') or '').strip(), date_value(values.get('received_date')) or received,
                     str(values.get('auditor') or '').strip() or auditor, json.dumps(raw, ensure_ascii=False)))
                self._history(conn, request_id, 'imported', {'source_id': source_id, 'row': number})
                count += 1
                if progress and count % 100 == 0:
                    progress({'processed': count, 'message': 'Importando solicitações'})
            if not count:
                raise ValueError('Nenhuma solicitação válida foi identificada. Confira a aba e as colunas.')
            conn.execute('INSERT INTO sources VALUES (?,?,?,?,?,?,?,?,?)',
                         (source_id, 'requests', str(metadata.get('filename') or Path(path).name), digest, sheet, stamp, count, json.dumps(dict(mapping, header_row=header_row)), 1))
            return {'imported': count, 'skipped': skipped, 'duplicate': False, 'source_id': source_id, 'batch_id': batch_id}

    def import_payments(self, path, sheet=None, mapping=None, metadata=None, progress=None):
        metadata = dict(metadata or {})
        sheet, mapping, header_row, headers, digest, source_id = self._prepare_import(path, sheet, mapping, 'payments', metadata)
        with self.lock, self.connection() as conn:
            existing = conn.execute('SELECT row_count,active FROM sources WHERE id=?', (source_id,)).fetchone()
            if existing and existing['active']:
                return {'imported': 0, 'existing': existing['row_count'], 'duplicate': True, 'source_id': source_id, 'message': 'Esta base já está carregada.'}
            old_signatures = self._review_signatures(conn)
            conn.execute('DELETE FROM payments')
            conn.execute("UPDATE sources SET active=0 WHERE kind='payments'")
            count = skipped = 0
            records = []
            duplicate_counts = {}
            sql = '''INSERT INTO payments VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'''
            for number, values, raw in self._iter_mapped(path, sheet, header_row, headers, mapping):
                code, branch = identifier(values.get('code')), identifier(values.get('branch'))
                if not code or not branch:
                    skipped += 1
                    continue
                amount = money(values.get('amount'))
                if amount is None:
                    skipped += 1
                    continue
                canonical = {
                    'code': code, 'branch': branch, 'amount': amount,
                    'payment_date': date_value(values.get('payment_date')),
                    'account': str(values.get('account') or '').strip(),
                    'account_code': identifier(values.get('account_code')),
                    'document_type': identifier(values.get('document_type')),
                    'exchange': str(values.get('exchange') or '').strip().upper(),
                    'related_code': identifier(values.get('related_code')),
                    'related_title': identifier(values.get('related_title')),
                    'title': identifier(values.get('title')),
                    'bordero': identifier(values.get('bordero')),
                    'bordero_type': identifier(values.get('bordero_type')),
                    'gross_amount': money(values.get('gross_amount')),
                    'interest': money(values.get('interest')),
                    'discount': money(values.get('discount')),
                    'fine': money(values.get('fine')),
                }
                content_id = hashlib.sha256(json.dumps(canonical, sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:28]
                ordinal = duplicate_counts.get(content_id, 0)
                duplicate_counts[content_id] = ordinal + 1
                payment_id = content_id + ':' + str(ordinal)
                records.append((payment_id, source_id, number, canonical['code'], canonical['branch'], canonical['amount'],
                    canonical['payment_date'], canonical['account'], canonical['account_code'], canonical['document_type'],
                    canonical['exchange'], canonical['related_code'], canonical['related_title'], canonical['title'],
                    canonical['bordero'], canonical['bordero_type'], canonical['gross_amount'], canonical['interest'],
                    canonical['discount'], canonical['fine'], json.dumps(raw, ensure_ascii=False)))
                count += 1
                if len(records) >= 2000:
                    conn.executemany(sql, records)
                    records.clear()
                    if progress:
                        progress({'processed': count, 'message': 'Importando base de baixas'})
            if records:
                conn.executemany(sql, records)
            if not count:
                raise ValueError('Nenhuma baixa válida foi identificada. Confira a aba e as colunas.')
            conn.execute('INSERT OR REPLACE INTO sources VALUES (?,?,?,?,?,?,?,?,?)',
                         (source_id, 'payments', str(metadata.get('filename') or Path(path).name), digest, sheet, now(), count, json.dumps(dict(mapping, header_row=header_row)), 1))
            self._invalidate_reviews(conn, old_signatures)
            self._history(conn, None, 'payments_imported', {'source_id': source_id, 'rows': count, 'skipped': skipped})
            return {'imported': count, 'skipped': skipped, 'duplicate': False, 'source_id': source_id, 'replaces_previous': True}

    def _history(self, conn, request_id, event, details):
        conn.execute('INSERT INTO history(request_id,event,created_at,details) VALUES (?,?,?,?)',
                     (request_id, event, now(), json.dumps(details, ensure_ascii=False)))

    def _matches(self, conn, request, config):
        if not request['code']:
            return []
        params = [request['code']]
        sql = 'SELECT * FROM payments WHERE code=?'
        if request['branch']:
            sql += ' AND branch=?'
            params.append(request['branch'])
        return [dict(row) for row in conn.execute(sql + ' ORDER BY payment_date,branch,document_type,title,id', params)]

    def _eligible(self, match, config):
        wanted = identifier(config['document_type'])
        return not wanted or identifier(match['document_type']) == wanted

    def _match_warnings(self, match, config):
        codes = []
        if normalize(match.get('exchange')) in ('s', 'sim', '1', 'true') or match.get('related_code'):
            codes.append('permuta')
        account = normalize(match.get('account'))
        if not account or any(normalize(term) in account for term in config.get('non_bank_terms', [])):
            codes.append('non_bank_account')
        if not match.get('document_type'):
            codes.append('missing_document_type')
        elif not self._eligible(match, config):
            codes.append('document_type_diff')
        if not match.get('payment_date'):
            codes.append('missing_payment_date')
        if match.get('gross_amount') is not None and all(match.get(key) is not None for key in ('interest', 'discount', 'fine')):
            if match['gross_amount'] + match['interest'] - match['discount'] + match['fine'] != match['amount']:
                codes.append('net_difference')
        return [{'code': code, 'label': WARNING_LABELS[code]} for code in codes]

    def _request_summary(self, conn, request, config, include_matches=False, sources=None):
        result = dict(request)
        candidates = self._matches(conn, request, config)
        matches = [match for match in candidates if self._eligible(match, config)]
        codes = []
        if not request['code']:
            codes.append('missing_code')
        if not request['branch']:
            codes.append('missing_branch')
        if not matches:
            codes.append('not_found')
        if len(matches) > 1:
            codes.append('multiple')
        if len({match['branch'] for match in candidates}) > 1:
            codes.append('multiple_branches')
        if len({match['document_type'] for match in candidates}) > 1:
            codes.append('multiple_types')
        if len({(match['account_code'], match['account']) for match in matches}) > 1:
            codes.append('multiple_accounts')
        if any(match['id'].rsplit(':', 1)[-1] != '0' for match in matches):
            codes.append('duplicate_record')
        total = sum(match['amount'] for match in matches)
        if matches and abs(total - request['amount']) > (money(config.get('tolerance', 0)) or 0):
            codes.append('value_diff')
        for match in candidates:
            for warning in self._match_warnings(match, config):
                if warning['code'] not in codes:
                    codes.append(warning['code'])
        if request['review_invalidated']:
            codes.append('review_invalidated')
        result['warnings'] = [{'code': code, 'label': WARNING_LABELS[code]} for code in codes]
        result['amount_cents'] = result['amount']
        result['amount'] = result['amount'] / 100
        result['matches_count'] = len(matches)
        result['candidates_count'] = len(candidates)
        result['matched_amount_cents'] = total
        result['matched_amount'] = total / 100
        result['eligible_amount_cents'] = total
        result['first_payment_date'] = min((match['payment_date'] for match in matches if match['payment_date']), default='')
        result['last_payment_date'] = max((match['payment_date'] for match in matches if match['payment_date']), default='')
        result['match_status'] = 'nao_encontrado' if not matches else ('conferir' if codes else 'encontrado')
        selections = json.loads(result['selected_match_ids'] or '[]')
        result['selected_match_ids'] = selections
        result['confirmed_match_ids'] = selections
        result['selected_amount_cents'] = sum(match['amount'] for match in matches if match['id'] in selections)
        result['selected_amount'] = result['selected_amount_cents'] / 100
        result['accounts'] = sorted({match['account'] for match in matches if match['account']})
        result['accounts_text'] = '; '.join(result['accounts'])
        result['bank_accounts'] = [{'code': code, 'description': description} for code, description in sorted({(match['account_code'], match['account']) for match in matches})]
        result['review_status'] = 'conferido' if result['status'] == 'pronto' else result['status']
        result['receipt_status'] = 'fase_seguinte'
        result['required_documents'], result['attachments'], result['missing_documents'] = [], [], []
        source = (sources or {}).get(request['source_id']) if sources is not None else conn.execute('SELECT file_name,sheet FROM sources WHERE id=?', (request['source_id'],)).fetchone()
        result['origin'] = {'file': source['file_name'], 'sheet': source['sheet'], 'row': request['origin_row']} if source else {'row': request['origin_row']}
        result['raw'] = json.loads(result['raw'] or '{}') if include_matches else None
        result['matched_payment_ids'] = [match['id'] for match in matches]
        if include_matches:
            result['matches'] = [self._payment_json(conn, match, config, selections, sources) for match in candidates]
        return result

    def _payment_json(self, conn, match, config, selections=None, sources=None):
        result = dict(match)
        for key in ('amount', 'gross_amount', 'interest', 'discount', 'fine'):
            result[key + '_cents'] = result[key]
            result[key] = result[key] / 100 if result[key] is not None else None
        result['date'] = result['payment_date']
        result['net_amount'] = result['amount']
        result['net_amount_cents'] = result['amount_cents']
        result['principal_amount'] = result['gross_amount']
        result['principal_amount_cents'] = result['gross_amount_cents']
        result['account_description'] = result['account']
        result['selected'] = result['id'] in (selections or [])
        result['eligible'] = self._eligible(match, config)
        result['warnings'] = self._match_warnings(match, config)
        result['raw'] = json.loads(result['raw'] or '{}')
        source = (sources or {}).get(match['source_id']) if sources is not None else conn.execute('SELECT file_name,sheet FROM sources WHERE id=?', (match['source_id'],)).fetchone()
        result['origin'] = {'file': source['file_name'], 'sheet': source['sheet'], 'row': match['origin_row']} if source else {'row': match['origin_row']}
        return result

    def state(self):
        with self.lock, self.connection() as conn:
            config = self._config(conn)
            sources = [dict(row) for row in conn.execute('SELECT id,kind,file_name,sheet,imported_at,row_count,active,mapping FROM sources ORDER BY imported_at DESC')]
            source_map = {source['id']: source for source in sources}
            for source in sources:
                source['mapping'] = json.loads(source['mapping'])
            requests = [self._request_summary(conn, row, config, sources=source_map) for row in conn.execute('SELECT * FROM requests ORDER BY received_date DESC,batch_id,origin_row')]
            batches = [dict(row) for row in conn.execute('SELECT * FROM batches ORDER BY created_at DESC')]
            matched_ids = {payment_id for request in requests for payment_id in request['matched_payment_ids']}
            matched_total = 0
            id_list = list(matched_ids)
            for start in range(0, len(id_list), 500):
                group = id_list[start:start + 500]
                matched_total += conn.execute('SELECT COALESCE(SUM(amount),0) FROM payments WHERE id IN (' + ','.join('?' for _ in group) + ')', group).fetchone()[0]
            totals = {
                'requests': len(requests), 'pending': sum(item['status'] == 'pendente' for item in requests),
                'in_review': sum(item['status'] == 'em_conferencia' for item in requests), 'ready': sum(item['status'] == 'pronto' for item in requests),
                'reviewed': sum(item['status'] == 'pronto' for item in requests), 'missing_documents': 0,
                'exceptions': sum(bool(item['warnings']) for item in requests), 'not_found': sum(item['matches_count'] == 0 for item in requests),
                'multiple': sum(item['matches_count'] > 1 for item in requests), 'matched_records': len(matched_ids),
                'requested_amount_cents': sum(item['amount_cents'] for item in requests), 'matched_records_amount_cents': matched_total,
                'payments': conn.execute('SELECT COUNT(*) FROM payments').fetchone()[0],
            }
            totals['requested_amount'] = totals['requested_amount_cents'] / 100
            totals['matched_records_amount'] = matched_total / 100
            return {'requests': requests, 'sources': sources, 'batches': batches, 'config': config, 'stats': totals, 'totals': totals,
                    'required_documents': [], 'ai': {'enabled': False, 'status': 'desativada'},
                    'note': 'Conferido significa vínculo revisado; localizar uma baixa não comprova pagamento nem entrega de comprovante.'}

    def reconcile(self):
        return self.state()

    def request_detail(self, request_id):
        with self.lock, self.connection() as conn:
            row = conn.execute('SELECT * FROM requests WHERE id=?', (str(request_id),)).fetchone()
            if row is None:
                raise ValueError('Solicitação não encontrada.')
            result = self._request_summary(conn, row, self._config(conn), True)
            result['history'] = [dict(h) for h in conn.execute('SELECT event,created_at,details FROM history WHERE request_id=? ORDER BY id DESC', (str(request_id),))]
            for event in result['history']:
                event['details'] = json.loads(event['details'] or '{}')
                event['action'] = event['event']
            return result

    def update_request(self, request_id, payload=None):
        if payload is None and isinstance(request_id, dict):
            payload = request_id
            request_id = payload.get('id') or payload.get('request_id')
        payload = dict(payload or {})
        with self.lock, self.connection() as conn:
            row = conn.execute('SELECT * FROM requests WHERE id=?', (str(request_id),)).fetchone()
            if row is None:
                raise ValueError('Solicitação não encontrada.')
            updated = dict(row)
            for key in ('branch', 'code'):
                if key in payload:
                    updated[key] = identifier(payload[key])
            for key in ('notes', 'justification'):
                if key in payload:
                    updated[key] = str(payload[key] or '')[:10000]
            if 'status' in payload:
                updated['status'] = 'pronto' if payload['status'] == 'conferido' else str(payload['status'])
            if updated['status'] not in ('pendente', 'em_conferencia', 'pronto'):
                raise ValueError('Status de conferência inválido.')
            config = self._config(conn)
            candidates = self._matches(conn, updated, config)
            eligible_ids = {match['id'] for match in candidates if self._eligible(match, config)}
            previous_selections = json.loads(row['selected_match_ids'] or '[]')
            selections = payload.get('selected_match_ids', payload.get('confirmed_match_ids', previous_selections))
            if not isinstance(selections, list):
                raise ValueError('Selecione os registros conferidos em uma lista.')
            selections = list(dict.fromkeys(str(item) for item in selections))
            if any(item not in eligible_ids for item in selections):
                raise ValueError('Uma baixa selecionada não corresponde à filial, código ou tipo configurado.')
            changed_identifiers = updated['branch'] != row['branch'] or updated['code'] != row['code']
            changed_selections = selections != previous_selections
            updated['selected_match_ids'] = json.dumps(selections)
            if updated['status'] == 'pronto':
                if not updated['branch']:
                    raise ValueError('Confirme a filial antes de marcar como conferido.')
                if not selections:
                    raise ValueError('Selecione pelo menos uma baixa antes de marcar como conferido.')
                summary = self._request_summary(conn, updated, config)
                selected_amount = sum(match['amount'] for match in candidates if match['id'] in selections)
                exceptional = bool([warning for warning in summary['warnings'] if warning['code'] != 'review_invalidated']) or abs(selected_amount - updated['amount']) > (money(config.get('tolerance', 0)) or 0)
                if exceptional and not updated['justification'].strip():
                    raise ValueError('Registre uma justificativa para conferir um item com exceções ou valores divergentes.')
                confirmed = payload.get('review_confirmed', payload.get('match_reviewed', False))
                if not confirmed and (row['status'] != 'pronto' or changed_identifiers or changed_selections or row['review_invalidated']):
                    raise ValueError('Confirme que conferiu os vínculos antes de marcar como conferido.')
                updated['reviewed_at'] = now() if confirmed else row['reviewed_at']
                updated['review_invalidated'] = 0
            elif changed_selections or changed_identifiers or row['status'] == 'pronto':
                updated['reviewed_at'] = None
            conn.execute('UPDATE requests SET branch=?,code=?,status=?,notes=?,justification=?,selected_match_ids=?,reviewed_at=?,review_invalidated=? WHERE id=?',
                (updated['branch'], updated['code'], updated['status'], updated['notes'], updated['justification'], updated['selected_match_ids'], updated['reviewed_at'], updated['review_invalidated'], request_id))
            self._history(conn, request_id, 'review_updated', {'status': updated['status'], 'selected_match_ids': selections,
                'previous_branch': row['branch'], 'branch': updated['branch'], 'previous_code': row['code'], 'code': updated['code'],
                'notes': updated['notes'], 'justification': updated['justification']})
        return self.request_detail(request_id)

    def _review_signatures(self, conn):
        config = self._config(conn)
        result = {}
        for row in conn.execute("SELECT * FROM requests WHERE reviewed_at IS NOT NULL OR selected_match_ids != '[]'"):
            result[row['id']] = sorted(m['id'] for m in self._matches(conn, row, config))
        return result

    def _invalidate_reviews(self, conn, old_signatures=None, force=False):
        config = self._config(conn)
        for row in list(conn.execute("SELECT * FROM requests WHERE reviewed_at IS NOT NULL OR selected_match_ids != '[]'")):
            candidates = self._matches(conn, row, config)
            available = {match['id'] for match in candidates}
            eligible = {match['id'] for match in candidates if self._eligible(match, config)}
            selected = json.loads(row['selected_match_ids'] or '[]')
            changed = force or any(item not in eligible for item in selected)
            if old_signatures is not None and old_signatures.get(row['id']) != sorted(available):
                changed = True
            if changed:
                retained = [item for item in selected if item in eligible]
                conn.execute("UPDATE requests SET status='em_conferencia',reviewed_at=NULL,review_invalidated=1,selected_match_ids=? WHERE id=?", (json.dumps(retained), row['id']))
                self._history(conn, row['id'], 'review_invalidated', {'reason': 'base_or_configuration_changed'})

def initialize(data_dir):
    return Engine(data_dir)
