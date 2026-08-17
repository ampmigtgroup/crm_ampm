# CRM Operacional AmPm — Desktop V5 Corporativo

## Objetivo desta etapa

A V5 prepara a distribuição centralizada e o atualizador assinado do aplicativo.

O updater do Tauri exige assinatura criptográfica. Por isso a chave privada NÃO está dentro deste pacote.

## Estado ao baixar este ZIP

- O programa continua compilando e rodando como a V4.
- O plugin do updater já está incluído no Rust.
- `createUpdaterArtifacts` permanece `false` até a chave pública ser configurada.
- Nenhuma chave privada está no projeto.
- O endpoint previsto é o GitHub Releases do projeto AmPm.
- Windows está preparado para instalação passiva de updates.
- Há workflow de GitHub Actions para releases assinados.

## 1. Teste normal

```bash
npm install
npm run dev
```

## 2. Gerar a chave de atualização

Faça isso UMA única vez:

```bash
chmod +x gerar_chave_atualizador.sh gerar_build_assinado.sh
./gerar_chave_atualizador.sh
```

A chave privada deve ficar fora do GitHub.

Depois:

```bash
ls -la ~/.tauri
```

Identifique o arquivo de chave pública gerado pelo Tauri.

## 3. Ativar o updater com a chave pública

Exemplo:

```bash
python3 configurar_updater.py ~/.tauri/ARQUIVO_DA_CHAVE_PUBLICA
```

Esse script altera apenas o `tauri.conf.json`.

## 4. Build assinado local

```bash
./gerar_build_assinado.sh
```

## 5. GitHub Actions

O workflow está em:

```text
.github/workflows/release-desktop.yml
```

Para uma release automática, o repositório precisa ter estes Secrets:

```text
TAURI_SIGNING_PRIVATE_KEY
TAURI_SIGNING_PRIVATE_KEY_PASSWORD
```

A chave pública NÃO é segredo e fica no `tauri.conf.json`.
A chave privada É segredo e nunca deve ser commitada.

## Importante

Faça backup seguro da chave privada. Se ela for perdida, os aplicativos já instalados com a chave pública correspondente não poderão aceitar novos updates assinados por outra chave.
