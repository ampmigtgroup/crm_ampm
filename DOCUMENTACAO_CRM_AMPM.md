# CRM AmPm · IGT Group — Arquitetura Nova

## Arquitetura ativa
- Backend: FastAPI.
- Banco: MongoDB.
- Frontend: React 19 + Vite.
- UI: Tailwind CSS v4 + componentes no padrão shadcn/ui.
- PWA: manifest instalável.
- Autenticação: JWT.
- Auditoria: eventos de criação, edição e exclusão registrados pela API.

## Interface visual
A referência visual aprovada foi preservada: fundo claro, sidebar grafite, identidade AmPm em amarelo/laranja, identidade IGT em azul, cards arredondados, sombras suaves, cabeçalho superior, filtros, KPIs, gráficos, tabelas e ações rápidas.

## Módulos
Dashboard Executivo; Pipeline AmPm; PROCV Gestão e Franquia AMPM; Calculadora & Otimizador de Custos; Call Center; Equipe de Instrutores; Orçamentos; Relatórios & Exportação; Agenda Inteligente; Lojas & Mapa; Mala Direta; WhatsApp; Carteira; Resumo Semanal; Auditoria; Usuários; SLA.

## API
- /api/auth/login
- /api/auth/me
- /api/dashboard
- /api/lojas
- /api/contatos
- /api/instrutores
- /api/agenda
- /api/pipeline
- /api/orcamentos
- /api/campanhas
- /api/auditoria
- /api/relatorios
- /api/usuarios
- /api/sla

## Banco
Coleções MongoDB: users, sessions, lojas, contatos, instrutores, agenda, oportunidades, orcamentos, campanhas, auditoria, relatorios, carteira, resumo_semanal e sla.

## Segurança
Nenhuma senha, token ou chave privada deve ser commitada. O administrador inicial é provisionado somente quando ADMIN_USERNAME e ADMIN_PASSWORD são fornecidos via ambiente/Secrets.

## Backup
A implementação anterior permanece preservada em: backup-streamlit-supabase-2026-10-04.

## Referência visual
A arquitetura visual enviada pelo usuário foi a referência do novo frontend, incluindo o Dashboard Executivo, sidebar AmPm/IGT, KPIs, gráficos, filtros e ações rápidas.