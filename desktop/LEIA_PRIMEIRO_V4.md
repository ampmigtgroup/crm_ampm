# CRM Operacional AmPm — Desktop V4 Resiliente

URL central:
`https://crmampm-operacional.streamlit.app/`

## O que muda nesta etapa

- Mantém splash, menu, bandeja, Sobre e Single Instance da V3.1.
- Adiciona detecção de perda de rede dentro do CRM.
- Exibe uma tela elegante de reconexão sem fechar o aplicativo.
- Faz verificação periódica de conectividade.
- Adiciona configuração específica para Windows.
- Windows gera instalador NSIS `-setup.exe`.
- Linux continua gerando AppImage e `.deb`.
- Inclui workflow do GitHub Actions para compilar Linux e Windows.
- Nenhuma chave do Supabase é colocada no desktop.
- O updater automático ainda NÃO foi ativado: ele será a próxima etapa, depois de definirmos assinatura e canal de releases.

## Testar no Linux

```bash
cd ~/Documentos/crm_ampm_desktop_v4
npm install
npm run dev
```

### Teste de resiliência

Com o CRM aberto:
1. desligue temporariamente o Wi-Fi/rede;
2. em poucos segundos a tela "Conexão interrompida" deve aparecer;
3. religue a rede;
4. clique em "Tentar reconectar".

## Build Linux

```bash
npm run build
```

## Windows

O arquivo `src-tauri/tauri.windows.conf.json` é mesclado automaticamente pelo Tauri quando o build roda no Windows.

A primeira distribuição Windows está configurada para NSIS, gerando um `-setup.exe`.

## GitHub Actions

O workflow está em:

`.github/workflows/build-desktop.yml`

Ele pode ser executado manualmente na aba Actions e compila os pacotes Linux e Windows como artifacts.

### Importante sobre assinatura

O instalador Windows de teste pode ser compilado sem certificado, mas para distribuição corporativa profissional é recomendável assinatura de código. O updater automático também deve usar pacotes assinados.
