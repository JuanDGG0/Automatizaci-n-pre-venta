import json
import base64
import traceback
from pathlib import Path
from http.server import BaseHTTPRequestHandler, HTTPServer
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
from generators import generate

PROJECT_PATH = Path(__file__).resolve().parent

class Handler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        print(f"[{self.address_string()}] {format % args}")

    def _send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self._send_json({}, 200)

    def do_POST(self):
        if self.path != "/generate":
            self._send_json({"ok": False, "error": "Ruta no encontrada"}, 404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            raw    = self.rfile.read(length).decode("utf-8")
            config = json.loads(raw) if raw else {}
            print(f"[REQUEST] filial={config.get('filial')}")

            out_dir = "/tmp/periferia_out"
            Path(out_dir).mkdir(parents=True, exist_ok=True)

            result = generate(config, out_dir)

            files = []
            for tipo, file_path in result.items():
                p = Path(file_path)
                if not p.exists():
                    raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
                with open(p, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode("utf-8")
                files.append({
                    "tipo": tipo,
                    "name": p.name,
                    "url":  "data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64," + b64
                })

            print(f"[OK] {len(files)} archivo(s)")
            self._send_json({"ok": True, "files": files})

        except Exception as e:
            tb = traceback.format_exc()
            print(f"[ERROR] {e}\n{tb}")
            self._send_json({"ok": False, "error": str(e)}, 500)

if __name__ == "__main__":
    server = HTTPServer(("0.0.0.0", 8090), Handler)
    print("Servidor listo en http://localhost:8090/generate")
    server.serve_forever()