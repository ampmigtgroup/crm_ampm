# CRM AmPm · IGT Group — Documentação de Operação

## Estado atual

O CRM oficial deste repositório continua baseado em **Streamlit + Supabase**, com a camada Enterprise/ Gestão 360 integrada na branch `feat/enterprise-crm-2026-10-03`.

### Módulos Enterprise
- Dashboard executivo.
- Agenda inteligente por proximidade.
- Carteira por consultor.
- SLA configurável (padrão: 180 dias).
- Auditoria de alterações no Supabase.
- Mala direta / prévia de e-mail.
- Modelos e abertura de WhatsApp.
- Gerador de orçamento com itens, desconto, anexos e exportação.
- Integração centralizada com Supabase.
- Identidade visual IGT Group.

### Dados e tabelas principais
- `crm_lojas`
- `crm_fila_callcenter`
- `crm_contatos`
- `crm_instrutores`
- `crm_agenda`
- `crm_recomendacao_deslocamento`
- `crm_orcamentos`
- `crm_orcamento_itens`
- `crm_orcamento_documentos`
- `crm_eventos`
- `crm_configuracoes`
- `crm_templates_email`
- `crm_campanhas_email`
- `crm_campanha_envios`
- `crm_resumo_envios`
- `crm_modelos_whatsapp`
- `crm_usuarios`
- `crm_permissoes_usuarios`

## Segurança

- Segredos não devem ser gravados no código nem neste documento.
- Credenciais administrativas devem permanecer em Secrets/.env.
- A função de sincronização do formulário não é executável diretamente por usuários públicos.
- A view unificada usa `security_invoker`.
- Tabelas comerciais, auditoria, agenda, configuração e campanhas estão protegidas por RLS.
- Índices de chaves estrangeiras foram reforçados.

## Custos configuráveis

- Km: R$ 2,20.
- Diária do instrutor: R$ 550,00.
- Hospedagem: R$ 280,00/noite.
- Limite de pernoite: 250 km.

## Operação

A aplicação usa o projeto Supabase `nptazzfvwhhmotfrvgdj` e deve receber URL/chaves por Secrets. Não publicar senhas, tokens ou chaves no GitHub.

## Validação

Executar:

```bash
python scripts/validate_crm.py app_crm.py
```

A documentação descreve a arquitetura efetivamente presente no repositório; funcionalidades de uma arquitetura FastAPI/React/MongoDB não devem ser tratadas como presentes neste código Streamlit até que sejam migradas de forma explícita.
