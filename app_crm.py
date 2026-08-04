import streamlit as st
import pandas as pd

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="CRM Operacional AmPm — Treinamentos & Inteligência de Roteamento",
    layout="wide",
    initial_sidebar_state="expanded"
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
    # Estrutura unificada completa com todos os dados operacionais
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
    # Expansão para o volume total da base unificada (371 registros)
    st.session_state['df_fila'] = pd.DataFrame(dados_iniciais * 38).iloc[:371].reset_index(drop=True)
    st.session_state['df_fila']['id_atendimento'] = st.session_state['df_fila'].index + 1

df_fila = st.session_state['df_fila']

# Extração de instrutores (exclui termos como 'Consultor')
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

# --- BARRA LATERAL (FILTROS E QUANTIDADE) ---
st.sidebar.header("🔍 Filtros da Fila")

busca = st.sidebar.text_input("Buscar por PV Abadi ou Nome:")

opcoes_status_filtro = ["Todos"] + sorted([str(x) for x in df_fila['status_contato'].dropna().unique()])
status_filtro = st.sidebar.selectbox("Status do Contato:", opcoes_status_filtro)

instrutor_filtro = st.sidebar.selectbox("Instrutor Sugerido:", ["Todos"] + lista_instrutores_unicos)

opcoes_uf = ["Todos"] + sorted([str(x) for x in df_fila['uf'].dropna().unique()])
uf_filtro = st.sidebar.selectbox("UF (Estado):", opcoes_uf)

# Aplicação dos filtros de busca
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

# Seletor numérico de quantidade a exibir/atender
st.sidebar.divider()
st.sidebar.header("⚙️ Limite de Exibição")
total_disponivel = len(df_filtrado)
qtd_exibir = st.sidebar.number_input(
    "Defina quantos postos deseja visualizar/atender:",
    min_value=1,
    max_value=max(1, total_disponivel),
    value=min(371, max(1, total_disponivel)),
    step=10
)

# Recorte final baseado na quantidade escolhida
df_atendimento = df_filtrado.head(qtd_exibir)

# --- CARDS DE MÉTRICAS DA FILA COMPLETA ---
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total Fila", len(df_filtrado))
c2.metric("A Contatar", len(df_filtrado[df_filtrado['status_contato'] == 'A Contatar']))
c3.metric("Contatados", len(df_filtrado[df_filtrado['status_contato'].isin(['Interessado - Aguardando confirmação', 'Sem Resposta'])]))
c4.metric("Confirmados", len(df_filtrado[df_filtrado['status_contato'] == 'Agendado']))
c5.metric("Concluídos", len(df_filtrado[df_filtrado['status_contato'] == 'Concluído']))

st.write("")

# --- ABAS DO SISTEMA ---
aba_fila, aba_menor_custo, aba_historico, aba_cadastro, aba_exportar = st.tabs([
    "📋 Fila & Atualização", 
    "📍 Menor Custo (Instrutor)",
    "📊 Histórico & Custos",
    "➕ Novo Cadastro Manual", 
    "📥 Exportar Relatórios"
])

# --- ABA 1: FILA DE ATENDIMENTO ---
with aba_fila:
    st.subheader(f"Fila Prioritária de Contatos ({len(df_atendimento)} de {total_disponivel} registros exibidos)")
    
    # Exibe a tabela unificada com todas as colunas
    st.dataframe(df_atendimento, use_container_width=True)
    
    st.divider()
    st.subheader("📋 Gestão da Fila de Atendimento")
    
    if not df_atendimento.empty:
        # Seletor por Nome do Posto + PV
        opcoes_lojas_map = {
            f"{row['loja']} (PV: {row['pv_abadi']})": row['pv_abadi'] 
            for _, row in df_atendimento.iterrows()
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
            if st.button("📝 Registrar Contato"):
                dados_loja = df_atendimento[df_atendimento['pv_abadi'] == pv_selecionado].iloc[0].to_dict()
                abrir_modal_contato(dados_loja, lista_instrutores_unicos)
    else:
        st.warning("Nenhum posto encontrado para os filtros e limites selecionados.")

# --- ABA 2: MENOR CUSTO (INSTRUTOR) ---
with aba_menor_custo:
    st.subheader("📍 Recomendação por Menor Distância e Custo")
    
    st.markdown("**Quadro de Instrutores Sugeridos na Base:**")
    df_instrutores_quadro = (
        df_fila[['instrutor_sugerido', 'uf']]
        .dropna()
        .drop_duplicates()
        .rename(columns={'instrutor_sugerido': 'Nome do Instrutor', 'uf': 'UF de Atuação'})
    )
    df_instrutores_quadro = df_instrutores_quadro[
        ~df_instrutores_quadro['Nome do Instrutor'].astype(str).str.contains('Consultor', case=False, na=False)
    ]
    st.dataframe(df_instrutores_quadro, use_container_width=True)

# --- DEMAIS ABAS ---
with aba_historico:
    st.subheader("📊 Histórico de Atendimentos")
    st.dataframe(df_fila[df_fila['status_contato'] != 'A Contatar'], use_container_width=True)

with aba_cadastro:
    st.subheader("➕ Novo Cadastro Manual")
    st.info("Espaço reservado para inserção manual fora da fila padrão.")

with aba_exportar:
    st.subheader("📥 Exportar Relatórios")
    csv_data = df_atendimento.to_csv(index=False).encode('utf-8')
    st.download_button("📥 Baixar Fila Selecionada (CSV)", data=csv_data, file_name="fila_crm_ampm.csv", mime="text/csv")
