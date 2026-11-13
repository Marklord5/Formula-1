#!/bin/bash
# Doble clic: sirve esta carpeta en http://localhost:8765 y abre el reporte (lee data/ automáticamente).
# Para detenerlo, cierra esta ventana de Terminal.
cd "$(dirname "$0")"
PORT=8765
if ! curl -s "http://localhost:$PORT" >/dev/null; then
  # Sin caché: así el navegador siempre toma el CSS/JS más reciente después de un cambio
  python3 -c 'import http.server as h, sys
class H(h.SimpleHTTPRequestHandler):
    def end_headers(self): self.send_header("Cache-Control", "no-store"); super().end_headers()
h.test(H, port=int(sys.argv[1]))' "$PORT" >/dev/null 2>&1 &
  sleep 1
fi
open "http://localhost:$PORT/index.html"
wait
