#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DEST_DIR="/opt/automacao_afiliados"
ICON_SOURCE="$SCRIPT_DIR/imagens/icone.png"
ICON_FALLBACK="$SCRIPT_DIR/imagens/whatsapp.png"
ICON_TARGET="/usr/share/icons/automacao_afiliados.png"
APP_TARGET="/usr/share/applications/automacao_afiliados.desktop"

if [ ! -f "$SCRIPT_DIR/dist/automacao_afiliados" ]; then
  echo "Executável não encontrado em $SCRIPT_DIR/dist/automacao_afiliados"
  echo "Primeiro rode: ./build_executable.sh"
  exit 1
fi

if [ "$(id -u)" -ne 0 ]; then
  if command -v sudo >/dev/null 2>&1; then
    SUDO="sudo"
  else
    echo "Este script precisa de privilégios administrativos para instalar em /opt e /usr/share."
    exit 1
  fi
else
  SUDO=""
fi

run_as_root() {
  if [ -n "$SUDO" ]; then
    "$SUDO" "$@"
  else
    "$@"
  fi
}

run_as_root mkdir -p "$DEST_DIR"
run_as_root cp "$SCRIPT_DIR/dist/automacao_afiliados" "$DEST_DIR/automacao_afiliados"

OWNER_USER="${SUDO_USER:-$(id -un)}"
if [ -n "$OWNER_USER" ] && [ "$OWNER_USER" != "root" ]; then
  run_as_root chown -R "$OWNER_USER:$OWNER_USER" "$DEST_DIR"
  run_as_root chmod -R u+rwX,g+rX,o+rX "$DEST_DIR"
fi

if [ -f "$ICON_SOURCE" ]; then
  run_as_root mkdir -p /usr/share/icons
  run_as_root cp "$ICON_SOURCE" "$ICON_TARGET"
elif [ -f "$ICON_FALLBACK" ]; then
  run_as_root mkdir -p /usr/share/icons
  run_as_root cp "$ICON_FALLBACK" "$ICON_TARGET"
else
  echo "Erro: não foi encontrado um ícone válido em $SCRIPT_DIR/imagens/icone.png nem em $SCRIPT_DIR/imagens/whatsapp.png."
  exit 1
fi

run_as_root mkdir -p /usr/share/applications
run_as_root cp "$SCRIPT_DIR/automacao_afiliados.desktop" "$APP_TARGET"

if command -v update-desktop-database >/dev/null 2>&1; then
  run_as_root update-desktop-database /usr/share/applications
fi

cat <<'EOF'

Instalação concluída.

Agora você pode:
- abrir o app no menu de aplicativos;
- ou executar manualmente com:
  /opt/automacao_afiliados/automacao_afiliados
EOF
