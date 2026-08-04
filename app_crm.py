import streamlit as st
import pandas as pd
import glob
import os

# --- 1. CONFIGURAÇÃO E ESTILO DA PÁGINA ---
st.set_page_config(
    page_title="CRM Operacional AmPm — Treinamentos & Inteligência de Roteamento",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .stButton>button {
        background-color: #e0a96d;
        color: #000;
        font-weight: bold;
        border-radius: 5px;
        width: 100%;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. MOTOR DE CARREGAMENTO ROBUSTO DA BASE REAL ---
@st.cache_data(ttl=300)
def carregar_base_oficial():
    """
    Busca e carrega automaticamente qualquer arquivo Excel/CSV no diretório
    priorizando a base completa original.
    """
    # Procura arquivos no diretório do projeto
    arquivos_excel = glob.glob("*.xlsx") + glob.glob("*.xls")
    arquivos_csv = glob.glob("*.csv")
    
    # Ignora temporários do Excel
    arquivos_excel = [f for f in arquivos_excel if not os.path.basename(f).startswith("~$")]

    if arquivos_excel:
        return pd.read_excel(arquivos_excel[0])
    elif arquivos_csv:
        return pd.read_csv(arquivos_csv[0])
    
    return None

df_raw = carregar_base_oficial()

if df_raw is not None and not df_raw.empty:
    df_fila = df_raw.copy()
else:
    # Se não encontrar arquivo no disco, tenta ler da sessão Streamlit existente
    if 'df_fila' in st.session_state:
        df_fila = st.session_state['df_fila']
    else:
        st.error("⚠️ Nenhuma base de dados (.xlsx ou .csv) foi localizada na pasta do projeto.")
        st.stop()

# Garantia de integridade de colunas essenciais
colunas_necessarias = {
    'pv_abadi': 'pv_abadi',
    'loja': 'loja',
    'status_contato': 'status_contato',
    'instrutor_sugerido': 'instrutor_sugerido'
}

if 'status_contato' not in df_fila.columns:
    df_fila['status_contato'] = 'A Contatar'

if 'id_atendimento' not in df_fila.columns:
    df_fila['id_atendimento'] = df_fila.index + 1

# --- 3. MODAL DE ATENDIMENTO REVISADO ---
@st.dialog("📝 Registrar Contato / Atendimento")
def abrir_modal_contato(loja_dados, lista_instrutores_opcoes):
    st.markdown(f"### PV: {loja_dados.get('pv_abadi')} — {loja_dados.get('loja')}")
    st.caption(f"📍 Município: {loja_dados.get('municipio', 'N/A')} | UF: {loja_dados.get('uf', 'N/A')}")
    
    with st.form("form_registro_crm"):
        col1, col2 = st.columns(2)
        
        with col1:
            nome_contato = st.text_input(
                "Nome do Decisor / Contato", 
                value=str(loja_dados.get("nome_contato", "")) if pd.notna(loja_dados.get("nome_contato")) else ""
            )
            telefone = st.text_input(
                "Telefone / WhatsApp", 
                value=str(loja_dados.get("telefone", "")) if pd.notna(loja_dados.get("telefone")) else ""
            )
            
            opcoes_status = ["A Contatar", "Interessado - Aguardando confirmação", "Agendado", "Recusou", "Sem Resposta", "Loja Inativa"]
            status_atual = loja_dados.get("status_contato", "A Contatar")
            idx_status = opcoes_status.index(status_atual) if status_atual in opcoes_status else 0
            
            status = st.selectbox("Status da Abordagem", opcoes_status, index=idx_status)
            
        with col2:
            instrutor_def = loja_dados.get("instrutor_sugerido", "")
            idx_inst = lista_instrutores_opcoes.index(instrutor_def) if instrutor_def in lista_instrutores_opcoes else 0
            instrutor_alocado = st.selectbox("Instrutor Alocado", options=lista_instrutores_opcoes, index=idx_inst)
            data_agendamento = st.date_input("Data Prevista para Treinamento")
            
        obs = st.text_area(
            "Observações do Atendimento", 
            value=str(loja_dados.get("observacao", "")) if pd.notna(loja_dados.get("observacao")) else ""
        )
        
        salvar = st.form_submit_button("💾 Salvar Registro")
        if salvar:
            # Atualiza no estado global em memória
            idx_match = df_fila[df_fila['pv_abadi'] == loja_dados['pv_abadi']].index
            if not idx_match.empty:
                df_fila.loc[idx_match[0], 'status_contato'] = status
                df_fila.loc[idx_match[0], 'observacao'] = obs
                df_fila.loc[idx_match[0], 'instrutor_sugerido'] = instrutor_alocado
                st.session_state['df_fila'] = df_fila
            st.success("Registro atualizado com sucesso!")
            st.rerun()

# --- 4. CABEÇALHO PRINCIPAL ---
st.title("⛽ CRM Operacional AmPm — Treinamentos & Inteligência de Roteamento")

# --- 5. BARRA LATERAL (FILTROS) ---
st.sidebar.header("🔍 Filtros da Fila")

busca = st.sidebar.text_input("Buscar por PV Abadi ou Nome:")

opcoes_status_filtro = ["Todos"] + sorted([str(x) for x in df_fila['status_contato'].dropna().unique()])
status_filtro = st.sidebar.selectbox("Status do Contato:", opcoes_status_filtro)

# Filtro de Instrutor (Puxa estritamente instrutores)
lista_instrutores_unicos = sorted([str(x) for x in df_fila['instrutor_sugerido'].dropna().unique() if "Consultor" not in str(x)])
instrutor_filtro = st.sidebar.selectbox("Instrutor Sugerido:", ["Todos"] + lista_instrutores_unicos)

opcoes_uf = ["Todos"] + sorted([str(x) for x in df_fila['uf'].dropna().unique()]) if 'uf' in df_fila.columns else ["Todos"]
uf_filtro = st.sidebar.selectbox("UF (Estado):", opcoes_uf)

# Aplicação em cadeia dos filtros
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

if 'uf' in df_filtrado.columns and uf_filtro != "Todos":
    df_filtrado = df_filtrado[df_filtrado['uf'] == uf_filtro]

# --- 6. CARDS DE MÉTRICAS OPERACIONAIS ---
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total Fila", len(df_filtrado))
c2.metric("A Contatar", len(df_filtrado[df_filtrado['status_contato'] == 'A Contatar']))
c3.metric("Contatados", len(df_filtrado[df_filtrado['status_contato'].isin(['Interessado - Aguardando confirmação', 'Sem Resposta'])]))
c4.metric("Confirmados", len(df_filtrado[df_filtrado['status_contato'] == 'Agendado']))
c5.metric("Concluídos", len(df_filtrado[df_filtrado['status_contato'] == 'Concluído']))

st.write("")

# --- 7. ABAS DO SISTEMA ---
aba_fila, aba_menor_custo, aba_historico, aba_cadastro, aba_exportar = st.tabs([
    "📋 Fila & Atualização", 
    "📍 Menor Custo (Instrutor)",
    "📊 Histórico & Custos",
    "➕ Novo Cadastro Manual", 
    "📥 Exportar Relatórios"
])

# --- ABA 1: FILA & ATUALIZAÇÃO ---
with aba_fila:
    st.subheader(f"Fila Prioritária de Contatos ({len(df_filtrado)} registros)")
    
    # Exibição da tabela principal da fila
    st.dataframe(df_filtrado, use_container_width=True, height=350)
    
    st.divider()
    st.subheader("📋 Gestão da Fila de Atendimento")
    
    if not df_filtrado.empty:
        # Requisito 1: Seleção por Nome do Posto/Loja + PV formatado
        opcoes_lojas_map = {
            f"{row['loja']} (PV: {row['pv_abadi']})": row['pv_abadi'] 
            for _, row in df_filtrado.iterrows()
        }
        
        col_sel, col_btn = st.columns([3, 1])
        
        with col_sel:
            loja_selecionada_label = st.selectbox(
                "Selecione o Posto/Loja para atualizar registro:",
                options=list(opcoes_lojas_map.keys())
            )
            pv_selecionado = opcoes_lojas_map[loja_selecionada_label]
            
        with col_btn:
            st.write("")
            st.write("")
            # Requisito 2: Manutenção apenas do botão único da Gestão de Fila
            if st.button("📝 Registrar Contato"):
                dados_loja = df_filtrado[df_filtrado['pv_abadi'] == pv_selecionado].iloc[0].to_dict()
                abrir_modal_contato(dados_loja, lista_instrutores_unicos)
    else:
        st.warning("Nenhum posto encontrado para os filtros selecionados.")

# --- ABA 2: MENOR CUSTO (INSTRUTOR) ---
with aba_menor_custo:
    st.subheader("📍 Roteamento e Recomendação por Menor Distância (Instrutores)")
    
    # Requisito 3: Puxa e filtra estritamente a equipe de Instrutores (sem Consultores)
    if not df_filtrado.empty:
        df_instrutores_unicos = (
            df_filtrado[['instrutor_sugerido', 'uf']]
            .dropna()
            .drop_duplicates()
            .rename(columns={'instrutor_sugerido': 'Nome do Instrutor', 'uf': 'Estado (UF)'})
        )
        
        # Otimização: remove linhas onde o nome seja 'Consultor'
        df_instrutores_unicos = df_instrutores_unicos[
            ~df_instrutores_unicos['Nome do Instrutor'].astype(str).str.contains('Consultor', case=False, na=False)
        ]
        
        st.markdown("**Quadro de Instrutores Sugeridos na Base Atual:**")
        st.dataframe(df_instrutores_unicos, use_container_width=True)
    else:
        st.info("Nenhum dado de instrutor disponível para os filtros atuais.")

# --- ABA 3: HISTÓRICO & CUSTOS ---
with aba_historico:
    st.subheader("📊 Histórico de Atendimentos Realizados")
    df_historico = df_fila[df_fila['status_contato'] != 'A Contatar']
    if not df_historico.empty:
        st.dataframe(df_historico, use_container_width=True)
    else:
        st.info("Nenhum histórico registrado até o momento.")

# --- ABA 4: NOVO CADASTRO MANUAL ---
with aba_cadastro:
    st.subheader("➕ Inclusão Manual na Fila")
    st.info("Utilize este espaço para adicionar registros fora do fluxo padrão.")

# --- ABA 5: EXPORTAR RELATÓRIOS ---
with aba_exportar:
    st.subheader("📥 Exportação da Base Operacional")
    csv_data = df_filtrado.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Baixar Fila Filtrada (CSV)",
        data=csv_data,
        file_name="fila_atendimento_ampm.csv",
        mime="text/csv"
    )
