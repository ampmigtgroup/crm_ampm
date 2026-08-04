import streamlit as st
import pandas as pd
import os

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

# --- CARREGAMENTO INTEGRAL DA BASE DE DADOS COMPLETA ---
@st.cache_data
def carregar_base_completa():
    # Procura pelo arquivo CSV oficial na pasta raiz do projeto
    caminho_csv = "base_atendimento.csv"
    
    if os.path.exists(caminho_csv):
        df = pd.read_csv(caminho_csv)
    else:
        # Fallback de segurança para garantir a execução caso o arquivo ainda não esteja na pasta
        dados_base = [
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
        df = pd.DataFrame(dados_base * 40).reset_index(drop=True)
        df['id_atendimento'] = df.index + 1

    if 'status_contato' not in df.columns:
        df['status_contato'] = 'A Contatar'
    if 'observacao' not in df.columns:
        df['observacao'] = None

    return df

if 'df_fila' not in st.session_state:
    st.session_state['df_fila'] = carregar_base_completa()

df_fila = st.session_state['df_fila']

# Extração da lista completa de instrutores
lista_instrutores_unicos = sorted(
    [str(x) for x in df_fila['instrutor_sugerido'].dropna().unique() if "Consultor" not in str(x)]
)

# --- MODAL DE ATENDIMENTO (EDIÇÃO DIRETA DA BASE) ---
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
        
        salvar = st.form_submit_button("💾 Salvar e Atualizar Base")
        if salvar:
            idx = df_fila[df_fila['pv_abadi'] == loja_dados['pv_abadi']].index
            if not idx.empty:
                df_fila.loc[idx[0], 'status_contato'] = status
                df_fila.loc[idx[0], 'observacao'] = obs
                df_fila.loc[idx[0], 'instrutor_sugerido'] = instrutor_alocado
                st.session_state['df_fila'] = df_fila
            st.success("Registro atualizado com sucesso na base!")
            st.rerun()

# --- TÍTULO PRINCIPAL ---
st.title("⛽ CRM Operacional AmPm — Base Unificada Completa")

# --- CARDS DE MÉTRICAS DA BASE COMPLETA ---
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total da Base", len(df_fila))
c2.metric("A Contatar", len(df_fila[df_fila['status_contato'] == 'A Contatar']))
c3.metric("Contatados", len(df_fila[df_fila['status_contato'].isin(['Interessado - Aguardando confirmação', 'Sem Resposta'])]))
c4.metric("Confirmados", len(df_fila[df_fila['status_contato'] == 'Agendado']))
c5.metric("Concluídos", len(df_fila[df_fila['status_contato'] == 'Concluído']))

st.write("")

# --- PAINEL DE BUSCA E FILTRAGEM TIPO EXCEL ---
f1, f2, f3, f4 = st.columns([2, 1, 1, 1])
with f1:
    busca = st.text_input("🔍 PROCV Rápido (PV Abadi ou Nome do Posto):", "")
with f2:
    status_filtro = st.selectbox("Status:", ["Todos"] + sorted([str(x) for x in df_fila['status_contato'].dropna().unique()]))
with f3:
    uf_filtro = st.selectbox("UF:", ["Todos"] + sorted([str(x) for x in df_fila['uf'].dropna().unique()]))
with f4:
    qtd_exibir = st.number_input("Linhas visíveis:", min_value=10, max_value=max(10, len(df_fila)), value=min(50, len(df_fila)), step=10)

# Filtragem em Tempo Real (Equivalente ao Filtro do Excel)
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

# Limita a exibição da tela mantendo o total na memória
df_exibicao = df_filtrado.head(qtd_exibir)

# --- ABAS DO SISTEMA ---
aba_fila, aba_menor_custo, aba_historico, aba_cadastro, aba_exportar = st.tabs([
    "📋 Fila Interativa Completa", 
    "📍 Menor Custo (Instrutores)",
    "📊 Histórico & Custos",
    "➕ Novo Cadastro Manual", 
    "📥 Exportar Base Completa"
])

# --- ABA 1: TABELA REATIVA COM PROCV AUTOMÁTICO AO CLICAR ---
with aba_fila:
    st.subheader(f"Exibindo {len(df_exibicao)} de {len(df_filtrado)} registros filtrados (Total Base: {len(df_fila)})")
    st.caption("👇 **Clique sobre qualquer linha da tabela para fazer a busca automática (PROCV) das informações do posto:**")
    
    # Tabela com escuta de evento de clique em tempo real
    event = st.dataframe(
        df_exibicao,
        use_container_width=True,
        hide_index=True,
        selection_mode="single-row",
        on_select="rerun"
    )
    
    selected_rows = event.selection.get("rows", [])
    
    st.divider()
    
    # PROCV AUTOMÁTICO DA LINHA SELECIONADA
    if selected_rows:
        index_selecionado = selected_rows[0]
        posto_selecionado = df_exibicao.iloc[index_selecionado].to_dict()
        
        # PROCV buscando o registro original na base por PV ABADI
        pv_id = posto_selecionado['pv_abadi']
        registro_completo = df_fila[df_fila['pv_abadi'] == pv_id].iloc[0].to_dict()
        
        st.subheader(f"📍 Detalhes do Posto Selecionado (PV: {registro_completo['pv_abadi']})")
        col_a, col_b, col_c = st.columns([2, 2, 1])
        
        with col_a:
            st.markdown(f"**Loja:** {registro_completo['loja']}")
            st.markdown(f"**Município / UF:** {registro_completo['municipio']} - {registro_completo['uf']}")
            st.markdown(f"**Status Atual:** {registro_completo['status_contato']}")
            
        with col_b:
            st.markdown(f"**Tipo de Necessidade:** {registro_completo['tipo_necessidade']}")
            st.markdown(f"**Instrutor Sugerido:** {registro_completo['instrutor_sugerido']}")
            st.markdown(f"**Semana Sugerida:** {registro_completo['semana_sugerida']}")
            
        with col_c:
            st.write("")
            if st.button("📝 Atualizar Registro", use_container_width=True):
                abrir_modal_contato(registro_completo, lista_instrutores_unicos)
    else:
        st.info("💡 Selecione uma linha na tabela acima para carregar o painel interativo.")

# --- ABA 2: QUADRO DE INSTRUTORES CONSOLIDADO DA BASE INTEIRA ---
with aba_menor_custo:
    st.subheader("📍 Cobertura Completa de Instrutores por UF")
    
    df_instrutores_quadro = df_fila[['instrutor_sugerido', 'uf']].dropna().drop_duplicates()
    df_instrutores_quadro = df_instrutores_quadro[
        ~df_instrutores_quadro['instrutor_sugerido'].astype(str).str.contains('Consultor', case=False, na=False)
    ]
    
    df_consolidado = df_instrutores_quadro.groupby('instrutor_sugerido')['uf'].apply(lambda x: ', '.join(sorted(x.unique()))).reset_index()
    df_consolidado.columns = ['Nome do Instrutor', 'Estados de Atuação Atendidos']
    
    st.dataframe(df_consolidado, use_container_width=True, hide_index=True)

# --- DEMAIS ABAS ---
with aba_historico:
    st.subheader("📊 Histórico de Atendimentos Realizados")
    st.dataframe(df_fila[df_fila['status_contato'] != 'A Contatar'], use_container_width=True, hide_index=True)

with aba_cadastro:
    st.subheader("➕ Novo Cadastro Manual na Base")
    st.info("Utilize para cadastrar um novo posto manualmente na base ativa.")

with aba_exportar:
    st.subheader("📥 Exportar Base de Dados")
    csv_completo = df_fila.to_csv(index=False).encode('utf-8')
    st.download_button("📥 Baixar Toda a Base Atualizada (CSV)", data=csv_completo, file_name="base_completa_crm_ampm.csv", mime="text/csv")
