#!/usr/bin/env bash
# Usa o Python do Homebrew com Tk (evita python3 das Command Line Tools → abort Tk no macOS 26).
set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"

pick_python() {
  local base
  for base in /opt/homebrew /usr/local; do
    for v in 3.13 3.12 3.11; do
      local py="${base}/opt/python@${v}/bin/python${v}"
      if [[ -x "$py" ]]; then
        echo "$py"
        return 0
      fi
    done
  done
  if command -v python3.13 >/dev/null 2>&1; then command -v python3.13; return 0; fi
  if command -v python3.12 >/dev/null 2>&1; then command -v python3.12; return 0; fi
  echo "python3"
}

PY="$(pick_python)"
if [[ "$PY" == "python3" ]]; then
  echo "Aviso: usando 'python3' do PATH. No macOS 26, se aparecer 'macOS 26 … required, have instead 16', instale:" >&2
  echo "  brew install python@3.13 python-tk@3.13" >&2
  echo "e rode de novo este script ou: /opt/homebrew/opt/python@3.13/bin/python3.13 index.py" >&2
fi

exec "$PY" index.py "$@"
