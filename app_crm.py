import streamlit as st
import pandas as pd

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="CRM Operacional AmPm — Treinamentos & Inteligência de Roteamento",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilização visual (Cores oficiais AmPm)
st.markdown("""
    <style>
    .stButton>button {
        background-color: #e0a96d;
        color: #000;
        font-weight: bold;
        border-radius: 5px;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem;
    }
    </style>
""", unsafe_allow_html=True)

# --- CARREGAMENTO DA BASE UNIFICADA DE DADOS ---
if 'df_fila' not in st.session_state:
    dados_iniciais = [
        {'id_atendimento': 1, 'pv_abadi': 621193, 'loja': 'Conveniencia Rodrigues E Companhia Ltda', 'municipio': 'Atalaia', 'uf': 'AL', 'tipo_necessidade': 'Retreinamento', 'instrutor_sugerido': 'Isabela Paim Ricardo', 'semana_sugerida': '19/04/2027 a 23/04/2027', 'status_contato': 'A Contatar', 'observacao': None},
        {'id_atendimento': 2, 'pv_abadi': 621258, 'loja': 'New Star Com De Combs E Lubrifi Ltda', 'municipio': 'Maceio', 'uf': 'AL', 'tipo_necessidade': 'Retreinamento', 'instrutor_sugerido': 'Isabela Paim Ricardo', 'semana_sugerida': '05/04/2027 a 09/04/2027', 'status_contato': 'A Contatar', 'observacao': None},
        {'id_atendimento': 3, 'pv_abadi': 621112, 'loja': 'Auto Posto Bariloche Eireli', 'municipio': 'Maceio', 'uf': 'AL', 'tipo_necessidade': 'Retreinamento', 'instrutor_sugerido': 'Isabela Paim Ricardo', 'semana_sugerida': '05/04/2027 a 09/04/2027', 'status_contato': 'A Contatar', 'observacao': None},
        {'id_atendimento': 4, 'pv_abadi': 709048, 'loja': 'Beluma Comercio De Combustiveis Ltda', 'municipio': 'Maceio', 'uf': 'AL', 'tipo_necessidade': 'Retreinamento', 'instrutor_sugerido': 'Isabela Paim Ricardo', 'semana_sugerida': '12/04/2027 a 16/04/2027', 'status_contato': 'A Contatar', 'observacao': None},
        {'id_atendimento': 5, 'pv_abadi': 665493, 'loja': 'Pratagy Frances Ltda', 'municipio': 'Marechal Deodoro', 'uf': 'AL', 'tipo_necessidade': 'Retreinamento', 'instrutor_sugerido': 'Carla Fernandes Dionizio', 'semana_sugerida': '22/03/2027 a 26/03/2027', 'status_contato': 'A Contatar', 'observacao': None},
        {'id_atendimento': 6, 'pv_abadi': 724183, 'loja': 'POSTO SANTA RITA', 'municipio': 'Rio Largo', 'uf': 'AL', 'tipo_necessidade': 'Treinamento Nova Loja (Implantação)', 'instrutor_sugerido': 'Isabela Paim Ricardo', 'semana_sugerida': '12/04/2027 a 16/04/2027', 'status_contato': 'A Contatar', 'observacao': None},
        {'id_atendimento': 7, 'pv_abadi': 702900, 'loja': 'TULEMON COMERCIO', 'municipio': 'Rio Largo', 'uf': 'AL', 'tipo_necessidade': 'Treinamento Nova Loja (Implantação)', 'instrutor_sugerido': 'Isabela Paim Ricardo', 'semana_sugerida': '19/04/2027 a 23/04/2027', 'status_contato': 'A Contatar', 'observacao': None},
        {'id_atendimento': 8, 'pv_abadi': 654251, 'loja': 'POSTO ALEXA LTDA', 'municipio': 'Sao Sebastiao', 'uf': 'AL', 'tipo_necessidade': 'Treinamento Nova Loja (Implantação)', 'instrutor_sugerido': 'Isabela Paim Ricardo', 'semana_sugerida': '15/03/2027 a 19/03/2027', 'status_contato': 'A Contatar', 'observacao': None},
        {'id_atendimento': 9, 'pv_abadi': 646306, 'loja': 'Auto Posto Sabalanga Ltda - Me', 'municipio': 'Vicosa', 'uf': 'AL', 'tipo_necessidade': 'Retreinamento', 'instrutor_sugerido': 'Carla Fernandes Dionizio', 'semana_sugerida': '22/03/2027 a 26/03/2027', 'status_contato': 'A Contatar', 'observacao': None},
        {'id_atendimento': 10, 'pv_abadi': 684108, 'loja': 'Mucuripe Varejo Ltda', 'municipio': 'Manaus', 'uf': 'AM', 'tipo_necessidade': 'Retreinamento', 'instrutor_sugerido': 'Carla Fernandes Dionizio', 'semana_sugerida': '05/04/2027 a 09/04/2027', 'status_contato': 'A Contatar', 'observacao': None}
    ]
    st.session_state['df_fila'] = pd.DataFrame(dados_iniciais * 38).iloc[:371].reset_index(drop=True)
    st.session_state['df_fila']['id_atendimento'] = st.session_state['df_fila'].index + 1

df_fila = st.session_state['df_fila']

# Extração de lista limpa de instrutores
lista_instrutores_unicos = sorted(
    [str(x) for x in df_fila['instrutor_sugerido'].dropna().unique() if "Consultor" not in str(x)]
)

# --- MODAL DE ATENDIMENTO NATIVO ---
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
            instrutor_def = loja_dados.get("instrutor_sugerido", "")
            idx_inst = lista_instrutores.index(instrutor_def) if instrutor_def in lista_instrutores else 0
            instrutor_alocado = st.selectbox("Instrutor Alocado", options=lista_instrutores, index=idx_inst)
            data_agendamento = st.date_input("Data Prevista para Treinamento")
            
        obs = st.text_area("Observações do Atendimento", value=str(loja_dados.get("observacao", "") if pd.notna(loja_dados.get("observacao")) else ""))
        
        salvar = st.form_submit_button("💾 Salvar Registro")
        if salvar:
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

# --- CARDS DE MÉTRICAS COMPLETO ---
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total Fila", len(df_fila))
c2.metric("A Contatar", len(df_fila[df_fila['status_contato'] == 'A Contatar']))
c3.metric("Contatados", len(df_fila[df_fila['status_contato'].isin(['Interessado - Aguardando confirmação', 'Sem Resposta'])]))
c4.metric("Confirmados", len(df_fila[df_fila['status_contato'] == 'Agendado']))
c5.metric("Concluídos", len(df_fila[df_fila['status_contato'] == 'Concluído']))

st.write("")

# --- PAINEL DE BUSCA RÁPIDA (INTEGRADO NO TOPO DA PÁGINA) ---
f1, f2, f3 = st.columns([2, 1, 1])
with f1:
    busca = st.text_input("🔍 Buscar por PV Abadi ou Nome do Posto:", "")
with f2:
    status_filtro = st.selectbox("Status:", ["Todos"] + sorted([str(x) for x in df_fila['status_contato'].dropna().unique()]))
with f3:
    uf_filtro = st.selectbox("UF:", ["Todos"] + sorted([str(x) for x in df_fila['uf'].dropna().unique()]))

# Filtragem dinâmica
df_filtrado = df_fila.copy()
if busca:
    df_filtrado = df_filtrado[
        df_filtrado['loja'].astype(str).str.contains(busca, case=False, na=False) | 
        df_filtrado['pv_abadi'].astype(str).str.contains(busca, na=False)
    ]
if status_filtro != "Todos":
    df_filtrado = df_filtrado[df_filtrado['status_contato'] == status_filtro]
if uf_filtro != "Todos":
    df_filtrado = df_filtrado[df_filtrado['uf'] == uf_filtro]

# --- ABAS DO CRM ---
aba_fila, aba_menor_custo, aba_historico, aba_cadastro, aba_exportar = st.tabs([
    "📋 Fila & Atualização Directa", 
    "📍 Menor Custo (Instrutores)",
    "📊 Histórico & Custos",
    "➕ Novo Cadastro Manual", 
    "📥 Exportar Relatórios"
])

# --- ABA 1: SELEÇÃO E AÇÃO INTERATIVA DIRETA ---
with aba_fila:
    st.subheader(f"Fila Operacional ({len(df_filtrado)} registros)")
    st.caption("👇 **Clique em qualquer linha da tabela para selecionar o posto e interagir diretamente:**")
    
    # Tabela com seleção de linha nativa ativada (on_select="rerun")
    event = st.dataframe(
        df_filtrado,
        use_container_width=True,
        hide_index=True,
        selection_mode="single-row",
        on_select="rerun"
    )
    
    # Captura a linha selecionada pelo clique direto do usuário
    selected_rows = event.selection.get("rows", [])
    
    st.divider()
    
    if selected_rows:
        index_selecionado = selected_rows[0]
        posto_selecionado = df_filtrado.iloc[index_selecionado].to_dict()
        
        # Painel Informativo Interativo Reativo
        st.subheader("📍 Detalhes do Posto Selecionado")
        col_a, col_b, col_c = st.columns([2, 2, 1])
        
        with col_a:
            st.markdown(f"**Loja:** {posto_selecionado['loja']}")
            st.markdown(f"**PV Abadi:** {posto_selecionado['pv_abadi']}")
            st.markdown(f"**Localização:** {posto_selecionado['municipio']} - {posto_selecionado['uf']}")
            
        with col_b:
            st.markdown(f"**Necessidade:** {posto_selecionado['tipo_necessidade']}")
            st.markdown(f"**Instrutor Sugerido:** {posto_selecionado['instrutor_sugerido']}")
            st.markdown(f"**Semana Sugerida:** {posto_selecionado['semana_sugerida']}")
            
        with col_c:
            st.write("")
            if st.button("📝 Registrar Contato", use_container_width=True):
                abrir_modal_contato(posto_selecionado, lista_instrutores_unicos)
    else:
        st.info("💡 Clique em qualquer linha da tabela acima para carregar o posto e abrir o registro.")

# --- ABA 2: QUADRO DE INSTRUTORES CONSOLIDADO ---
with aba_menor_custo:
    st.subheader("📍 Quadro Consolidado de Instrutores na Base")
    
    # Agrupa por instrutor para mostrar todas as UFs consolidadas sem repetir linhas
    df_instrutores_quadro = (
        df_fila[['instrutor_sugerido', 'uf']]
        .dropna()
        .drop_duplicates()
    )
    df_instrutores_quadro = df_instrutores_quadro[
        ~df_instrutores_quadro['instrutor_sugerido'].astype(str).str.contains('Consultor', case=False, na=False)
    ]
    
    df_consolidado = df_instrutores_quadro.groupby('instrutor_sugerido')['uf'].apply(lambda x: ', '.join(sorted(x.unique()))).reset_index()
    df_consolidado.columns = ['Nome do Instrutor', 'Estados de Atuação (UF)']
    
    st.dataframe(df_consolidado, use_container_width=True, hide_index=True)

# --- DEMAIS ABAS ---
with aba_historico:
    st.subheader("📊 Histórico de Atendimentos")
    st.dataframe(df_fila[df_fila['status_contato'] != 'A Contatar'], use_container_width=True, hide_index=True)

with aba_cadastro:
    st.subheader("➕ Novo Cadastro Manual")
    st.info("Formulário para inclusão de novos postos.")

with aba_exportar:
    st.subheader("📥 Exportar Relatórios")
    csv_data = df_filtrado.to_csv(index=False).encode('utf-8')
    st.download_button("📥 Baixar Fila (CSV)", data=csv_data, file_name="fila_crm_ampm.csv", mime="text/csv")
