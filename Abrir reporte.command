#!/bin/bash
# Doble clic: sirve esta carpeta en http://localhost:8765 y abre el reporte (lee data/ automáticamente).
# Para detenerlo, cierra esta ventana de Terminal.
cd "$(dirname "$0")"
PORT=8765
if ! curl -s "http://localhost:$PORT" >/dev/null; then
  python3 -m http.server "$PORT" >/dev/null 2>&1 &
  sleep 1
fi
open "http://localhost:$PORT/index.html"
wait
