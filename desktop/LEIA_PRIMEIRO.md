# CRM Operacional AmPm — Desktop Executivo V3

URL central do CRM:
`https://crmampm-operacional.streamlit.app/`

## Novidades

- Splash premium preservada.
- Menu nativo do aplicativo:
  - Abrir CRM
  - Recarregar CRM
  - Tela cheia
  - Minimizar
  - Sobre
  - Sair
- Ícone na bandeja do sistema com menu rápido.
- Tela "Sobre" própria.
- Single Instance: impede abrir várias cópias do programa.
- Ao tentar abrir outra cópia, a instância existente recebe foco.
- Janela remota continua sem capacidades Tauri adicionais.
- Nenhuma chave do Supabase é embutida.
- Versão 3.0.0.
- Build Linux: AppImage e .deb.

## Teste

```bash
cd ~/Documentos/crm_ampm_desktop_v3
chmod +x preparar_linux.sh gerar_app_linux.sh
./preparar_linux.sh
npm run dev
```

## Build

```bash
./gerar_app_linux.sh
```

## Próxima etapa prevista

O atualizador automático será configurado somente quando definirmos o canal de releases e a assinatura criptográfica dos pacotes. Isso evita ativar um mecanismo de atualização incompleto.
