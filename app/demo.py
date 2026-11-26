"""Serve the synthetic workspace on loopback, using only the standard library."""

import argparse
import json
import tempfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from engine import Engine
from xlsx_writer import write_audit_workbook

ROOT = Path(__file__).resolve().parents[1]
STATIC = {
    '/': ('index.html', 'text/html; charset=utf-8'),
    '/estilo.css': ('estilo.css', 'text/css; charset=utf-8'),
    '/painel.js': ('painel.js', 'text/javascript; charset=utf-8'),
}
BOOK_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'


class Demo:
    def __init__(self, directory):
        self.directory = Path(directory)
        self.engine = Engine(self.directory)

    def load_sample(self):
        with self.engine.lock:
            self.engine.import_payments(ROOT / 'exemplos/baixas-sinteticas.csv')
            self.engine.import_requests(
                ROOT / 'exemplos/solicitacao-sintetica.csv',
                metadata={'batch_name': 'Lote de demonstração', 'received_date': '2026-04-20'},
            )
            return self.engine.state()

    def review(self, payload):
        request_id = payload.get('id')
        if not isinstance(request_id, str):
            raise ValueError('Selecione um documento da amostra.')
        fields = {key: payload[key] for key in (
            'branch', 'justification', 'selected_match_ids',
        ) if key in payload}
        confirming = payload.get('review_confirmed') is True
        fields.update(status='conferido' if confirming else 'em_conferencia',
                      review_confirmed=confirming)
        if 'branch' in fields:
            detail = self.engine.request_detail(request_id)
            if (detail['branch'] or confirming or not isinstance(fields['branch'], str) or fields['branch'] not in
                    {match['branch'] for match in detail['matches'] if match['eligible']}):
                raise ValueError('Escolha uma filial elegível para resolver a identificação.')
        return self.engine.update_request(request_id, fields)

    def workbook(self):
        with self.engine.lock:
            state = self.engine.state()
            if not state['requests']:
                raise ValueError('Carregue os dados de exemplo antes de exportar.')
            items = [self.engine.request_detail(item['id']) for item in state['requests']]
            path = self.directory / 'audit-fluxo-amostra.xlsx'
            write_audit_workbook(path, items, {
                'batch_name': 'Lote de demonstração · dados fictícios',
                'config': state['config'],
                'source_names': sorted(source['file_name'] for source in state['sources']),
            })
            return path.read_bytes()


def handler_for(demo):
    class Handler(BaseHTTPRequestHandler):
        def respond(self, payload, content_type, status=200, filename=None):
            self.send_response(status)
            self.send_header('Content-Type', content_type)
            self.send_header('Content-Length', str(len(payload)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.send_header('Content-Security-Policy', "default-src 'self'; frame-ancestors 'none'; form-action 'none'")
            if filename:
                self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.end_headers()
            self.wfile.write(payload)

        def json_response(self, result, status=200):
            self.respond(json.dumps(result, ensure_ascii=False).encode('utf-8'),
                         'application/json; charset=utf-8', status)

        def allowed(self):
            port = self.server.server_port
            host = self.headers.get('Host', '')
            origin = self.headers.get('Origin')
            if host not in (f'127.0.0.1:{port}', f'localhost:{port}') or (
                    origin is not None and origin != 'http://' + host):
                self.json_response({'erro': 'A demo aceita somente acesso local.'}, 403)
                return False
            return True

        def do_GET(self):
            if not self.allowed():
                return
            try:
                if self.path in STATIC:
                    name, content_type = STATIC[self.path]
                    self.respond((ROOT / 'static' / name).read_bytes(), content_type)
                elif self.path == '/api/painel':
                    self.json_response(demo.engine.state())
                elif self.path.startswith('/api/documento/'):
                    request_id = self.path.removeprefix('/api/documento/')
                    self.json_response(demo.engine.request_detail(request_id))
                elif self.path == '/api/planilha':
                    self.respond(demo.workbook(), BOOK_TYPE, filename='audit-fluxo-amostra.xlsx')
                else:
                    self.json_response({'erro': 'Página não encontrada.'}, 404)
            except ValueError as error:
                self.json_response({'erro': str(error)}, 400)

        def do_POST(self):
            if not self.allowed():
                return
            try:
                size = int(self.headers.get('Content-Length', '0'))
                if not 0 < size <= 8192 or self.headers.get_content_type() != 'application/json':
                    raise ValueError('Envie uma solicitação JSON válida.')
                payload = json.loads(self.rfile.read(size))
                if not isinstance(payload, dict):
                    raise ValueError('Envie um objeto JSON.')
                if self.path == '/api/exemplo':
                    self.json_response(demo.load_sample())
                elif self.path == '/api/conferencia':
                    with demo.engine.lock:
                        self.json_response(demo.review(payload))
                else:
                    self.json_response({'erro': 'Operação não encontrada.'}, 404)
            except ValueError as error:
                self.json_response({'erro': str(error)}, 400)

    return Handler


def main():
    parser = argparse.ArgumentParser(description='Painel local Audit Fluxo')
    parser.add_argument('--port', type=int, default=8765)
    args = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix='audit-fluxo-') as directory:
        with ThreadingHTTPServer(('127.0.0.1', args.port), handler_for(Demo(directory))) as server:
            print(f'Abra http://127.0.0.1:{server.server_port} no navegador.', flush=True)
            try:
                server.serve_forever()
            except KeyboardInterrupt:
                pass


if __name__ == '__main__':
    main()
