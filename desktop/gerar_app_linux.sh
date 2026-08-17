#!/usr/bin/env bash
set -e

if [ -f "$HOME/.cargo/env" ]; then
  . "$HOME/.cargo/env"
fi

echo "Gerando CRM Operacional AmPm Desktop Executivo V3..."
npm run build

echo
echo "Arquivos gerados:"
find src-tauri/target/release/bundle -type f \( -name "*.AppImage" -o -name "*.deb" \) -print 2>/dev/null || true
