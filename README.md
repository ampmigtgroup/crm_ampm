# CRM Operacional AmPm — IGT Group

## Arquitetura ativa
- Frontend: React 19 + Vite + Tailwind CSS v4.
- Componentes e padrões visuais no estilo shadcn/ui.
- Backend: FastAPI.
- Banco: MongoDB.
- PWA instalável.
- API REST em /api.
- Autenticação JWT e auditoria.

## Interface
Identidade visual AmPm/IGT: sidebar grafite, amarelo/laranja AmPm, azul IGT, fundo claro, cards arredondados, sombras leves, cabeçalho superior e navegação modular.

## Módulos
Dashboard Executivo, Pipeline AmPm, PROCV Gestão e Franquia AMPM, Calculadora & Otimizador de Custos, Call Center, Equipe de Instrutores, Orçamentos, Relatórios & Exportação, Agenda Inteligente, Lojas & Mapa, Mala Direta, WhatsApp, Carteira, Resumo Semanal, Auditoria, Usuários e SLA.

## Segurança
Nenhuma credencial é gravada no código. Defina JWT_SECRET, MONGO_URI e demais variáveis por ambiente/Secrets.

## Backup
A implementação anterior permanece preservada na branch backup-streamlit-supabase-2026-10-04.
