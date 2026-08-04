import streamlit as st
import pandas as pd

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

# --- DIÁLOGO / MODAL DE ATENDIMENTO ---
@st.dialog("📝 Registrar Contato / Atendimento")
def abrir_modal_contato(loja_dados, lista_instrutores):
    st.markdown(f"### PV: {loja_dados.get('pv_abadi')} — {loja_dados.get('loja')}")
    st.caption(f"📍 Município: {loja_dados.get('municipio')} | UF: {loja_dados.get('uf')}")
    
    with st.form("form_registro_crm"):
        col1, col2 = st.columns(2)
        
        with col1:
            nome_contato = st.text_input("Nome do Decisor / Contato", value=str(loja_dados.get("nome_contato", "")))
            telefone = st.text_input("Telefone / WhatsApp", value=str(loja_dados.get("telefone", "")))
            status = st.selectbox(
                "Status da Abordagem",
                ["A Contatar", "Interessado - Aguardando confirmação", "Agendado", "Recusou", "Sem Resposta", "Loja Inativa"]
            )
            
        with col2:
            instrutor_alocado = st.selectbox("Instrutor Alocado", options=lista_instrutores if lista_instrutores else ["Nenhum disponível"])
            data_agendamento = st.date_input("Data Prevista para Treinamento")
            
        obs = st.text_area("Observações do Atendimento", value=str(loja_dados.get("observacao", "")))
        
        salvar = st.form_submit_button("💾 Salvar Registro")
        if salvar:
            st.success("Registro atualizado com sucesso!")
            st.rerun()

# --- CARREGAMENTO SIMULADO / BASE ---
# Substitua pela sua leitura de dados (ex: pd.read_excel ou conexão SQL)
if 'df_fila' not in st.session_state:
    st.session_state['df_fila'] = pd.DataFrame([
        {
            'pv_abadi': 621193, 
            'loja': 'Conveniencia Rodrigues E Companhia Ltda', 
            'municipio': 'Atalaia', 
            'uf': 'AL',
            'id_atendimento': 1,
            'instrutor_sugerido': 'Isabela Paim Ricardo',
            'nome_contato': '',
            'telefone': '',
            'observacao': ''
        },
        {
            'pv_abadi': 621194, 
            'loja': 'Posto Central AmPm', 
            'municipio': 'Maceió', 
            'uf': 'AL',
            'id_atendimento': 2,
            'instrutor_sugerido': 'Carla Fernandes Dionizio',
            'nome_contato': '',
            'telefone': '',
            'observacao': ''
        }
    ])

df_fila = st.session_state['df_fila']
df_filtrado = df_fila.copy()

# Lista extraída diretamente dos instrutores sugeridos na fila
lista_instrutores = df_fila['instrutor_sugerido'].dropna().unique().tolist()

# --- CABEÇALHO PRINCIPAL ---
st.title("⛽ CRM Operacional AmPm — Treinamentos & Inteligência")

# --- BARRA LATERAL (FILTROS) ---
st.sidebar.header("🔍 Filtros da Fila")
busca = st.sidebar.text_input("Buscar por PV Abadi ou Nome:")
if busca:
    df_filtrado = df_filtrado[
        df_filtrado['loja'].str.contains(busca, case=False, na=False) | 
        df_filtrado['pv_abadi'].astype(str).str.contains(busca)
    ]

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
    st.subheader("📋 Gestão da Fila de Atendimento")
    
    if not df_filtrado.empty:
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
                abrir_modal_contato(dados_loja, lista_instrutores)
        
        st.divider()
        st.dataframe(df_filtrado, use_container_width=True)
    else:
        st.warning("Nenhum posto encontrado com os filtros aplicados.")

# --- ABA 2: MENOR CUSTO (INSTRUTOR) ---
with aba_menor_custo:
    st.subheader("📍 Recomendação por Menor Distância e Custo")
    try:
        lista_atendimentos = df_filtrado['id_atendimento'].tolist()
        if lista_atendimentos:
            id_selecionado = st.selectbox("Selecione o ID do atendimento para calcular custo:", lista_atendimentos)
            item = df_filtrado[df_filtrado['id_atendimento'] == id_selecionado].iloc[0]
            
            st.info(f"**Posto selecionado:** {item['loja']} | **Instrutor Recomendado:** {item['instrutor_sugerido']}")
        else:
            st.warning("Nenhum atendimento na lista para exibir.")
    except Exception as e:
        st.error(f"Erro ao carregar dados do instrutor: {e}")

# --- DEMAIS ABAS ---
with aba_historico:
    st.subheader("📊 Histórico de Atendimentos")
    st.write("Dados consolidados de interações.")

with aba_cadastro:
    st.subheader("➕ Incluir Novo Posto")
    st.write("Formulário de inclusão manual.")

with aba_exportar:
    st.subheader("📥 Exportação")
    st.write("Baixe a base atualizada em Excel ou CSV.")
