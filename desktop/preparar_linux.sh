#!/usr/bin/env bash
set -e

echo "============================================="
echo " CRM Operacional AmPm - Desktop Executivo V3"
echo "============================================="

if [ -f "$HOME/.cargo/env" ]; then
  . "$HOME/.cargo/env"
fi

command -v cargo >/dev/null 2>&1 || { echo "Rust/Cargo não encontrado."; exit 1; }
command -v npm >/dev/null 2>&1 || { echo "Node/NPM não encontrado."; exit 1; }

npm install

echo
echo "Preparação concluída."
echo "Teste: npm run dev"
echo "Build: ./gerar_app_linux.sh"
