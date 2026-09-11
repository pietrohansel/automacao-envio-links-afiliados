#!/bin/bash
set -euo pipefail

cd "$(dirname "$0")"

if [ ! -x .venv/bin/python ]; then
  echo "Ambiente virtual .venv não encontrado. Crie-o primeiro com: python -m venv .venv"
  exit 1
fi

.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r requirements.txt -r requirements-linux.txt pyinstaller

rm -rf dist build automacao_afiliados.spec

.venv/bin/python -m PyInstaller \
  --onefile \
  --name automacao_afiliados \
  --add-data "imagens:imagens" \
  --hidden-import=PIL._tkinter_finder \
  main.py

echo "\nExecutável gerado em: dist/automacao_afiliados"

