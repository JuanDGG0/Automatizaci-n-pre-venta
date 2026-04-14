import json
import base64
import traceback
from pathlib import Path
from http.server import BaseHTTPRequestHandler, HTTPServer
from socketserver import ThreadingMixIn
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
from generators import generate
from generators.fda_perfiles import _load_generales

PROJECT_PATH = Path(__file__).resolve().parent
STATIC_PATH  = PROJECT_PATH / 'static'

MIME_TYPES = {
    '.html': 'text/html',
    '.css':  'text/css',
    '.js':   'application/javascript',
    '.png':  'image/png',
    '.ico':  'image/x-icon',
}

# 🔥 Servidor multihilo (evita bloqueos)
class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True  # mata threads al cerrar

class Handler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        print(f"[{self.address_string()}] {format % args}")

    def _send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self._send_json({}, 200)

    def do_GET(self):
        path = self.path.split('?')[0]
        if path == '/':
            path = '/home.html'

        # Catálogo de perfiles para el buscador del frontend
        if path == '/api/perfiles-catalog':
            try:
                _, perf_db = _load_generales()
                perfiles = [
                    {'torre': torre, 'rol': p['rol'], 'desc': p['desc']}
                    for torre, profs in perf_db.items()
                    for p in profs
                ]
                self._send_json({'ok': True, 'perfiles': perfiles})
            except Exception as e:
                self._send_json({'ok': False, 'error': str(e)}, 500)
            return

        file_path = STATIC_PATH / path.lstrip('/')

        if not file_path.exists() or not file_path.is_file():
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b'Not found')
            return

        suffix = file_path.suffix.lower()
        mime   = MIME_TYPES.get(suffix, 'application/octet-stream')
        body   = file_path.read_bytes()

        self.send_response(200)
        self.send_header('Content-Type', mime)
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        if self.path != '/generate':
            self._send_json({'ok': False, 'error': 'Ruta no encontrada'}, 404)
            return
        try:
            length = int(self.headers.get('Content-Length', '0'))
            raw    = self.rfile.read(length).decode('utf-8')
            config = json.loads(raw) if raw else {}
            print(f'[REQUEST] filial={config.get("filial")}')

            out_dir = '/tmp/periferia_out'
            Path(out_dir).mkdir(parents=True, exist_ok=True)

            result = generate(config, out_dir)

            files = []
            for tipo, file_path in result.items():
                p = Path(file_path)
                if not p.exists():
                    raise FileNotFoundError(f'Archivo no encontrado: {file_path}')
                with open(p, 'rb') as f:
                    b64 = base64.b64encode(f.read()).decode('utf-8')
                files.append({
                    'tipo': tipo,
                    'name': p.name,
                    'url':  'data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,' + b64
                })

            print(f'[OK] {len(files)} archivo(s)')
            self._send_json({'ok': True, 'files': files})

        except Exception as e:
            tb = traceback.format_exc()
            print(f'[ERROR] {e}\n{tb}')
            self._send_json({'ok': False, 'error': str(e)}, 500)


if __name__ == '__main__':
    server = ThreadedHTTPServer(('0.0.0.0', 8090), Handler)
    server.allow_reuse_address = True

    print('Servidor listo en:')
    print('  → http://localhost:8090')
    print('  → http://localhost:8090/generate')

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n[OK] Cerrando servidor...')
        server.shutdown()   # 🔑 clave para cerrar correctamente
    finally:
        server.server_close()
        print('[OK] Servidor detenido.')