#!/usr/bin/env bash
set -e

if [ -f "$HOME/.cargo/env" ]; then
  . "$HOME/.cargo/env"
fi

KEY="$HOME/.tauri/crm_ampm_updater.key"

if [ ! -f "$KEY" ]; then
  echo "Chave privada não encontrada em $KEY"
  echo "Execute primeiro: ./gerar_chave_atualizador.sh"
  exit 1
fi

if ! grep -q '"createUpdaterArtifacts": true' src-tauri/tauri.conf.json; then
  echo "O updater ainda não foi ativado no tauri.conf.json."
  echo "Execute configurar_updater.py com o arquivo da chave pública."
  exit 1
fi

export TAURI_SIGNING_PRIVATE_KEY="$KEY"

echo
echo "A senha da chave (se houver) será solicitada pelo processo de build."
echo "Não salve a senha dentro do projeto."
echo

npm run build

echo
echo "✅ Build finalizado."
echo "Assinaturas .sig e pacotes de update devem estar dentro de:"
echo "  src-tauri/target/release/bundle/"
find src-tauri/target/release/bundle -type f \( -name "*.sig" -o -name "*.AppImage" -o -name "*.deb" -o -name "*setup.exe" \) -print 2>/dev/null || true
