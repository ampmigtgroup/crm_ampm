#!/usr/bin/env bash
set -e

if [ -f "$HOME/.cargo/env" ]; then
  . "$HOME/.cargo/env"
fi

mkdir -p "$HOME/.tauri"

KEY="$HOME/.tauri/crm_ampm_updater.key"

echo "======================================================"
echo " CRM Operacional AmPm - Chave de Atualização"
echo "======================================================"
echo
echo "A chave PRIVADA será criada somente neste computador."
echo "NÃO envie essa chave por chat, e-mail ou GitHub."
echo
echo "Arquivo privado previsto:"
echo "  $KEY"
echo

if [ -e "$KEY" ]; then
  echo "Já existe uma chave privada nesse caminho."
  echo "Para proteger instalações futuras, o script NÃO vai sobrescrevê-la."
  exit 1
fi

npm run tauri signer generate -- -w "$KEY"

echo
echo "✅ Geração concluída."
echo
echo "Agora liste os arquivos criados com:"
echo "  ls -la \"$HOME/.tauri\""
echo
echo "A chave pública pode ser compartilhada/configurada."
echo "A chave privada $KEY deve permanecer secreta e ter backup seguro."
