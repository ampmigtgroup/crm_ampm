#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"

if [ -f "$HOME/.cargo/env" ]; then
  . "$HOME/.cargo/env"
fi

command -v patchelf >/dev/null 2>&1 || {
  echo "patchelf não encontrado. Rode: sudo apt install -y patchelf"
  exit 2
}

if [ ! -f "$HOME/.tauri/crm_ampm_updater.key" ]; then
  echo "Chave privada do updater não encontrada."
  exit 3
fi

export TAURI_SIGNING_PRIVATE_KEY="$HOME/.tauri/crm_ampm_updater.key"

if [ -z "${TAURI_SIGNING_PRIVATE_KEY_PASSWORD+x}" ]; then
  read -s -p "Senha da chave do updater (Enter se não houver): " TAURI_SIGNING_PRIVATE_KEY_PASSWORD
  echo
  export TAURI_SIGNING_PRIVATE_KEY_PASSWORD
fi

rm -rf src-tauri/target/release/bundle/appimage
rm -rf src-tauri/target/release/bundle/deb

npm run build

echo
echo "Arquivos V6 Linux:"
find src-tauri/target/release/bundle \
  -type f \( -name "*6.0.0*.AppImage" -o -name "*6.0.0*.deb" -o -name "*6.0.0*.sig" \) \
  -print 2>/dev/null || true
