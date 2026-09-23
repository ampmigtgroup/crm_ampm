# Fluxo Único de Dados — CRM Operacional AmPm

## Fonte de verdade

O Supabase/PostgreSQL passa a ser a fonte oficial de dados do CRM.

- **Rede de lojas:** `crm_lojas`
- **Call Center/Pipeline:** `crm_fila_callcenter`
- **Inaugurações:** `crm_inauguracoes`
- **Instrutores:** `crm_instrutores`
- **Recomendações de deslocamento:** `crm_recomendacao_deslocamento`
- **Contatos:** `crm_contatos`
- **Orçamentos:** `crm_orcamentos` + `crm_orcamento_itens`
- **Documentos de orçamento:** Supabase Storage + `crm_orcamento_documentos`
- **Eventos/auditoria:** `crm_eventos`
- **Respostas do formulário público:** `crm_form_respostas`
- **Visão consolidada:** `crm_visao_unificada`

## Fluxo operacional

1. O Excel entra somente pelo **Importador Inteligente**.
2. O importador normaliza os dados e grava no Supabase.
3. O formulário público grava a resposta em `crm_form_respostas`.
4. Um trigger do banco replica os dados operacionais relevantes para Rede de Lojas e Call Center e registra o evento.
5. O Call Center altera o registro central e registra o evento.
6. Pipeline, PROCV, Dashboard, Calculadora e Relatórios leem o mesmo banco central.
7. Orçamentos e itens são gravados no PostgreSQL.
8. Documentos de orçamento são gravados no Storage privado.
9. A visão `crm_visao_unificada` consolida o estado atual por PV para consultas gerenciais.
10. `crm_eventos` mantém o histórico operacional sem substituir as tabelas transacionais.

## Regra importante

`st.session_state` é apenas estado de interface. Ele não é fonte permanente.

O Excel não é mais banco de dados.

Arquivos de orçamento não são mais gravados no filesystem do Streamlit.

## Compatibilidade

As funções antigas de orçamento foram mantidas como wrappers para não quebrar os módulos existentes. O restante do CRM continua usando os DataFrames internos, mas a persistência oficial é o Supabase.

## Próximo passo de segurança

As quatro tabelas novas de operação (`crm_orcamentos`, `crm_orcamento_itens`, `crm_orcamento_documentos`, `crm_eventos`) precisam ter RLS habilitado antes de qualquer exposição por chave anon/authenticated. O CRM usa a chave server-side e, depois que RLS for habilitado sem políticas públicas, o acesso por clientes não privilegiados ficará bloqueado.
