import json
import subprocess
import base64
import traceback
from pathlib import Path
from http.server import BaseHTTPRequestHandler, HTTPServer

PROJECT_PATH = "/mnt/c/Users/heidyromero/periferia_preventa"


class Handler(BaseHTTPRequestHandler):

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
            raw = self.rfile.read(length).decode("utf-8")
            print(f"\n[REQUEST] Body recibido: {raw[:300]}")
            body = json.loads(raw) if raw else {}

            mode = (body.get("mode") or "fda").strip().lower()
            torres = body.get("torres", []) or []

            payload = {
                "mode": mode,
                "torres": torres,
            }

            print(f"[INFO] mode={mode}, torres={[t.get('name') for t in torres]}")

            out_dir = "/tmp/periferia_api"
            Path(out_dir).mkdir(parents=True, exist_ok=True)

            cmd = [
                "python3",
                f"{PROJECT_PATH}/generator.py",
                json.dumps(payload, ensure_ascii=False),
                out_dir,
            ]

            print(f"[CMD] {' '.join(cmd[:2])} ...")
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                cwd=PROJECT_PATH
            )

            print(f"[RETURNCODE] {result.returncode}")
            print(f"[STDOUT] {result.stdout[:500]}")
            print(f"[STDERR] {result.stderr[:500]}")

            if result.returncode != 0:
                error_msg = result.stderr or result.stdout or "Error ejecutando generator.py"
                print(f"[ERROR] returncode != 0: {error_msg}")
                self._send_json({"ok": False, "error": error_msg}, 500)
                return

            stdout = (result.stdout or "").strip()
            if not stdout:
                print("[ERROR] stdout vacío")
                self._send_json({"ok": False, "error": "generator.py no devolvió salida"}, 500)
                return

            try:
                generated = json.loads(stdout)
            except Exception as e:
                print(f"[ERROR] JSON parse failed: {e} | stdout: {stdout[:200]}")
                self._send_json({"ok": False, "error": "Salida inválida de generator.py", "raw": stdout}, 500)
                return

            if not isinstance(generated, dict) or not generated:
                print(f"[ERROR] generated vacío o no es dict: {generated}")
                self._send_json({"ok": False, "error": "generator.py no devolvió archivos"}, 500)
                return

            files = []

            for tipo, file_path in generated.items():
                p = Path(file_path)
                print(f"[FILE] tipo={tipo}, path={file_path}, exists={p.exists()}")
                if not p.exists():
                    self._send_json({"ok": False, "error": f"El archivo generado no existe: {file_path}"}, 500)
                    return

                with open(p, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode("utf-8")

                if tipo == "fda":
                    name = "Periferia_IT_Fuera_del_Alcance.pptx"
                elif tipo == "perfiles":
                    name = "Periferia_IT_Perfiles.pptx"
                elif tipo == "ambos":
                    name = "Periferia_IT_Propuesta_Completa.pptx"
                else:
                    name = p.name

                files.append({
                    "tipo": tipo,
                    "name": name,
                    "url": "data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64," + b64
                })

            print(f"[OK] Enviando {len(files)} archivo(s)")
            self._send_json({"ok": True, "files": files})

        except Exception as e:
            tb = traceback.format_exc()
            print(f"[EXCEPTION] {e}\n{tb}")
            self._send_json({"ok": False, "error": str(e), "traceback": tb}, 500)


if __name__ == "__main__":
    server = HTTPServer(("0.0.0.0", 8090), Handler)
    print("Servidor listo en http://localhost:8090/generate")
    server.serve_forever()