import streamlit as st
import pandas as pd
import os

# Configuração da página
st.set_page_config(page_title="CRM Operacional AmPm", layout="wide")

# Estilização visual (Cores AmPm)
st.markdown("""
    <style>
    .stButton>button {
        background-color: #e0a96d;
        color: #000;
        font-weight: bold;
        border-radius: 5px;
    }
    </style>
""", unsafe_allow_html=True)

# --- CARREGAMENTO DA BASE DE DADOS (PRESERVA A BASE ORIGINAL) ---
@st.cache_data
def carregar_dados():
    # Caminhos possíveis para encontrar sua planilha oficial no projeto
    caminhos = ["base_crm.xlsx", "dados.xlsx", "fila.csv", "crm_data.xlsx"]
    for caminho in caminhos:
        if os.path.exists(caminho):
            if caminho.endswith('.csv'):
                return pd.read_csv(caminho)
            return pd.read_excel(caminho)
    
    # Se não encontrar arquivo local, tenta ler no session_state existente
    return None

df_base = carregar_dados()

if df_base is None:
    if 'df_fila' in st.session_state:
        df_fila = st.session_state['df_fila']
    else:
        # Fallback de segurança apenas se nenhum arquivo for localizado
        df_fila = pd.DataFrame([
            {'id_atendimento': 1, 'pv_abadi': 621193, 'loja': 'Conveniencia Rodrigues E Companhia Ltda', 'municipio': 'Atalaia', 'uf': 'AL', 'tipo_necessidade': 'Retreinamento', 'instrutor_sugerido': 'Isabela Paim Ricardo', 'semana_sugerida': '19/04/2027 a 23/04/2027', 'status_contato': 'A Contatar', 'observacao': None}
        ])
else:
    df_fila = df_base.copy()

# Garantir colunas padrão padronizadas
if 'status_contato' not in df_fila.columns:
    df_fila['status_contato'] = 'A Contatar'

# --- DIÁLOGO / MODAL DE ATENDIMENTO ---
@st.dialog("📝 Registrar Contato / Atendimento")
def abrir_modal_contato(loja_dados, lista_instrutores):
    st.markdown(f"### PV: {loja_dados.get('pv_abadi')} — {loja_dados.get('loja')}")
    st.caption(f"📍 Município: {loja_dados.get('municipio')} | UF: {loja_dados.get('uf')}")
    
    with st.form("form_registro_crm"):
        col1, col2 = st.columns(2)
        
        with col1:
            nome_contato = st.text_input("Nome do Decisor / Contato", value=str(loja_dados.get("nome_contato", "") if pd.notna(loja_dados.get("nome_contato")) else ""))
            telefone = st.text_input("Telefone / WhatsApp", value=str(loja_dados.get("telefone", "") if pd.notna(loja_dados.get("telefone")) else ""))
            
            opcoes_status = ["A Contatar", "Interessado - Aguardando confirmação", "Agendado", "Recusou", "Sem Resposta", "Loja Inativa"]
            status_atual = loja_dados.get("status_contato", "A Contatar")
            idx_status = opcoes_status.index(status_atual) if status_atual in opcoes_status else 0
            
            status = st.selectbox("Status da Abordagem", opcoes_status, index=idx_status)
            
        with col2:
            instrutor_def = loja_dados.get("instrutor_sugerido")
            idx_inst = lista_instrutores.index(instrutor_def) if instrutor_def in lista_instrutores else 0
            instrutor_alocado = st.selectbox("Instrutor Alocado", options=lista_instrutores, index=idx_inst)
            data_agendamento = st.date_input("Data Prevista para Treinamento")
            
        obs = st.text_area("Observações do Atendimento", value=str(loja_dados.get("observacao", "") if pd.notna(loja_dados.get("observacao")) else ""))
        
        salvar = st.form_submit_button("💾 Salvar Registro")
        if salvar:
            # Atualiza no DataFrame em memória
            idx = df_fila[df_fila['pv_abadi'] == loja_dados['pv_abadi']].index
            if not idx.empty:
                df_fila.loc[idx[0], 'status_contato'] = status
                df_fila.loc[idx[0], 'observacao'] = obs
                df_fila.loc[idx[0], 'instrutor_sugerido'] = instrutor_alocado
                st.session_state['df_fila'] = df_fila
            st.success("Registro atualizado com sucesso!")
            st.rerun()

# --- TÍTULO PRINCIPAL ---
st.title("⛽ CRM Operacional AmPm — Treinamentos & Inteligência de Roteamento")

# --- FILTROS NA BARRA LATERAL ---
st.sidebar.header("🔍 Filtros da Fila")
busca = st.sidebar.text_input("Buscar por PV Abadi ou Nome:")

status_filtro = st.sidebar.selectbox(
    "Status do Contato:",
    ["Todos"] + list(df_fila['status_contato'].dropna().unique())
)

lista_instrutores_filtro = sorted(df_fila['instrutor_sugerido'].dropna().astype(str).unique().tolist())
instrutor_filtro = st.sidebar.selectbox(
    "Instrutor Sugerido:",
    ["Todos"] + lista_instrutores_filtro
)

uf_filtro = st.sidebar.selectbox(
    "UF (Estado):",
    ["Todos"] + sorted(list(df_fila['uf'].dropna().unique()))
)

# Aplicação dos Filtros
df_filtrado = df_fila.copy()

if busca:
    df_filtrado = df_filtrado[
        df_filtrado['loja'].astype(str).str.contains(busca, case=False, na=False) | 
        df_filtrado['pv_abadi'].astype(str).str.contains(busca, na=False)
    ]

if status_filtro != "Todos":
    df_filtrado = df_filtrado[df_filtrado['status_contato'] == status_filtro]

if instrutor_filtro != "Todos":
    df_filtrado = df_filtrado[df_filtrado['instrutor_sugerido'] == instrutor_filtro]

if uf_filtro != "Todos":
    df_filtrado = df_filtrado[df_filtrado['uf'] == uf_filtro]

# --- CARDS DE MÉTRICAS (EXATAMENTE COMO NO SEU DASHBOARD) ---
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total Fila", len(df_filtrado))
c2.metric("A Contatar", len(df_filtrado[df_filtrado['status_contato'] == 'A Contatar']))
c3.metric("Contatados", len(df_filtrado[df_filtrado['status_contato'].isin(['Interessado - Aguardando confirmação', 'Sem Resposta'])]))
c4.metric("Confirmados", len(df_filtrado[df_filtrado['status_contato'] == 'Agendado']))
c5.metric("Concluídos", len(df_filtrado[df_filtrado['status_contato'] == 'Concluído']))

st.write("")

# --- ABAS DE NAVEGAÇÃO ---
aba_fila, aba_menor_custo, aba_historico, aba_cadastro, aba_exportar = st.tabs([
    "📋 Fila & Atualização", 
    "📍 Menor Custo (Instrutor)",
    "📊 Histórico & Custos",
    "➕ Novo Cadastro Manual", 
    "📥 Exportar Relatórios"
])

# --- ABA 1: FILA DE ATENDIMENTO ---
with aba_fila:
    st.subheader(f"Fila Prioritária de Contatos ({len(df_filtrado)} registros)")
    
    st.dataframe(df_filtrado, use_container_width=True)
    
    st.divider()
    st.subheader("📋 Gestão da Fila de Atendimento")
    
    if not df_filtrado.empty:
        # Formata o seletor exibindo NOME DA LOJA (PV)
        opcoes_lojas = {f"{row['loja']} (PV: {row['pv_abadi']})": row['pv_abadi'] for _, row in df_filtrado.iterrows()}
        
        col_sel, col_btn = st.columns([3, 1])
        
        with col_sel:
            loja_selecionada_texto = st.selectbox(
                "Selecione o Posto/Loja para atualizar registro:",
                options=list(opcoes_lojas.keys())
            )
            pv_selecionado = opcoes_lojas[loja_selecionada_texto]
            
        with col_btn:
            st.write("")
            st.write("")
            if st.button("📝 Registrar Contato"):
                dados_loja = df_filtrado[df_filtrado['pv_abadi'] == pv_selecionado].iloc[0].to_dict()
                abrir_modal_contato(dados_loja, lista_instrutores_filtro)
    else:
        st.warning("Nenhum posto encontrado para os filtros selecionados.")

# --- ABA 2: MENOR CUSTO (INSTRUTOR) ---
with aba_menor_custo:
    st.subheader("📍 Recomendação por Menor Distância e Custo")
    
    # Puxa exclusivamente a lista única de instrutores
    st.write("**Instrutores Disponíveis na Base:**")
    df_instrutores_resumo = df_fila[['instrutor_sugerido', 'uf']].dropna().drop_duplicates()
    st.dataframe(df_instrutores_resumo, use_container_width=True)

# --- DEMASI ABAS ---
with aba_historico:
    st.subheader("📊 Histórico de Atendimentos")
    st.dataframe(df_fila[df_fila['status_contato'] != 'A Contatar'], use_container_width=True)

with aba_cadastro:
    st.subheader("➕ Incluir Novo Posto")
    st.info("Utilize este espaço para cadastro manual fora da fila padrão.")

with aba_exportar:
    st.subheader("📥 Exportação de Dados")
    csv = df_filtrado.to_csv(index=False).encode('utf-8')
    st.download_button("📥 Baixar Base Filtrada (CSV)", data=csv, file_name="fila_crm_ampm.csv", mime="text/csv")
