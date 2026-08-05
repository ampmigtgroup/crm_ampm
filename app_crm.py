import streamlit as st
import pandas as pd
import os

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="CRM Operacional AmPm — Base Unificada Completa",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilização (Cores oficiais AmPm)
st.markdown("""
    <style>
    .stButton>button {
        background-color: #e0a96d;
        color: #000;
        font-weight: bold;
        border-radius: 5px;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.6rem;
    }
    .procv-card {
        background-color: #1e222a;
        padding: 15px;
        border-radius: 8px;
        border-left: 4px solid #e0a96d;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# --- CARREGAMENTO E PROCV INTEGRADO DA PLANILHA OFICIAL ---
@st.cache_data
def carregar_e_cruzar_base():
    caminho_excel = "Base_Unificada_AmPm.xlsx"
    
    if os.path.exists(caminho_excel):
        try:
            xls = pd.ExcelFile(caminho_excel, engine='openpyxl')
            df_lojas = pd.read_excel(xls, sheet_name='Rede_de_Lojas')
            df_fila = pd.read_excel(xls, sheet_name='Fila_CallCenter')
            df_inaug = pd.read_excel(xls, sheet_name='Previsao_Inauguracao')
            
            # Padronização de chaves para o PROCV (PV Abadi)
            df_lojas['PV Abadi'] = pd.to_numeric(df_lojas['PV Abadi'], errors='coerce')
            df_fila['PV_Abadi'] = pd.to_numeric(df_fila['PV_Abadi'], errors='coerce')
            df_inaug['PV ABADI'] = pd.to_numeric(df_inaug['PV ABADI'], errors='coerce')
            
            # PROCV 1: Cruzando Fila com Rede de Lojas Completa
            df_base = pd.merge(
                df_lojas,
                df_fila[['PV_Abadi', 'Tipo_Necessidade', 'Data_Ultimo_Treinamento', 
                         'Dias_desde_Ultimo_Treinamento', 'Instrutor_Sugerido', 
                         'Semana_Sugerida', 'Status_Contato', 'Observacoes']],
                left_on='PV Abadi',
                right_on='PV_Abadi',
                how='left'
            )
            
            # PROCV 2: Cruzando com Previsão de Inauguração
            df_base = pd.merge(
                df_base,
                df_inaug[['PV ABADI', 'Previsão Inauguração', 'Pipeline']],
                left_on='PV Abadi',
                right_on='PV ABADI',
                how='left'
            )
            
            # Preenchimento de padrões para colunas PROCV não encontradas
            df_base['Status_Contato'] = df_base['Status_Contato'].fillna('A Contatar')
            df_base['Tipo_Necessidade'] = df_base['Tipo_Necessidade'].fillna('Sem Pendência Identificada')
            df_base['Instrutor_Sugerido'] = df_base['Instrutor_Sugerido'].fillna('Pendente de Alocação')
            
            return df_base
        except Exception as e:
            st.error(f"Erro ao ler planilha Excel: {e}")
            return pd.DataFrame()
    else:
        st.warning("⚠️ Planilha 'Base_Unificada_AmPm.xlsx' não localizada no repositório. Realize o 'git add -f Base_Unificada_AmPm.xlsx'.")
        return pd.DataFrame()

if 'df_unificado' not in st.session_state:
    st.session_state['df_unificado'] = carregar_e_cruzar_base()

df_base = st.session_state['df_unificado']

# --- TÍTULO PRINCIPAL ---
st.title("⛽ CRM Operacional AmPm — Inteligência PROCV em Tempo Real")

if not df_base.empty:
    # --- MÉTRICAS GERAIS DERAVIDAS DO PROCV ---
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Rede Total de Lojas", len(df_base))
    c2.metric("Fila Prevista CallCenter", len(df_base[df_base['Tipo_Necessidade'] != 'Sem Pendência Identificada']))
    c3.metric("A Contatar", len(df_base[df_base['Status_Contato'] == 'A Contatar']))
    c4.metric("Agendados / Confirmados", len(df_base[df_base['Status_Contato'] == 'Agendado']))
    c5.metric("Previsão Inaugurações", len(df_base[df_base['Previsão Inauguração'].notna()]))

    st.write("")

    # --- BARRA DE PROCV RÁPIDO E FILTROS MULTICOLUNAS ---
    st.subheader("🔍 PROCV Multicoluna em Tempo Real")
    col_busca, col_uf, col_tipo, col_status = st.columns([2, 1, 1, 1])
    
    with col_busca:
        termo_busca = st.text_input("Digite PV Abadi, Nome da Loja ou Cidade:", "")
    with col_uf:
        ufs_unicas = ["Todos"] + sorted([str(x) for x in df_base['UF'].dropna().unique()])
        filtro_uf = st.selectbox("UF:", ufs_unicas)
    with col_tipo:
        tipos_unicos = ["Todos"] + sorted([str(x) for x in df_base['Tipo_Necessidade'].dropna().unique()])
        filtro_tipo = st.selectbox("Necessidade:", tipos_unicos)
    with col_status:
        status_unicos = ["Todos"] + sorted([str(x) for x in df_base['Status_Contato'].dropna().unique()])
        filtro_status = st.selectbox("Status Contato:", status_unicos)

    # Execução das regras do PROCV nos filtros
    df_filtrado = df_base.copy()
    
    if termo_busca:
        df_filtrado = df_filtrado[
            df_filtrado['Razão Social'].astype(str).str.contains(termo_busca, case=False, na=False) |
            df_filtrado['PV Abadi'].astype(str).str.contains(termo_busca, na=False) |
            df_filtrado['Municipio'].astype(str).str.contains(termo_busca, case=False, na=False)
        ]
        
    if filtro_uf != "Todos":
        df_filtrado = df_filtrado[df_filtrado['UF'] == filtro_uf]
        
    if filtro_tipo != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Tipo_Necessidade'] == filtro_tipo]
        
    if filtro_status != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Status_Contato'] == filtro_status]

    # --- TABELA INTERATIVA DE SELEÇÃO PROCV ---
    st.markdown(f"**Resultado do PROCV:** {len(df_filtrado)} postos encontrados")
    
    colunas_exibicao = [
        'PV Abadi', 'Razão Social', 'Municipio', 'UF', 
        'Status Loja', 'Tipo_Necessidade', 'Instrutor_Sugerido', 
        'Semana_Sugerida', 'Status_Contato'
    ]
    
    event = st.dataframe(
        df_filtrado[colunas_exibicao],
        use_container_width=True,
        hide_index=True,
        selection_mode="single-row",
        on_select="rerun"
    )

    selected_rows = event.selection.get("rows", [])

    st.divider()

    # --- CARD DE DETALHES DO PROCV (TODAS AS COLUNAS CRUZADAS) ---
    if selected_rows:
        idx = selected_rows[0]
        posto = df_filtrado.iloc[idx].to_dict()

        st.subheader(f"📋 Painel PROCV Completo — PV: {posto['PV Abadi']} | {posto['Razão Social']}")
        
        col_p1, col_p2, col_p3 = st.columns(3)
        
        with col_p1:
            st.markdown("<div class='procv-card'>", unsafe_allow_html=True)
            st.markdown("### 🏪 Cadastro de Loja (Rede)")
            st.markdown(f"**Status Loja:** {posto.get('Status Loja', '-')}")
            st.markdown(f"**Endereço:** {posto.get('Endereço', '-')}")
            st.markdown(f"**Município/UF:** {posto.get('Municipio', '-')}/{posto.get('UF', '-')}")
            st.markdown(f"**CEP:** {posto.get('CEP', '-')}")
            st.markdown(f"**Modelo de Loja:** {posto.get('Modelo Loja Gerencial', '-')}")
            st.markdown("</div>", unsafe_allow_html=True)

        with col_p2:
            st.markdown("<div class='procv-card'>", unsafe_allow_html=True)
            st.markdown("### 👔 Franquia & Gestão")
            st.markdown(f"**Gerência de Franquia (GF):** {posto.get('GF', '-')}")
            st.markdown(f"**Consultor de Franquia (CF):** {posto.get('CF', '-')}")
            st.markdown(f"**Previsão Inauguração:** {posto.get('Previsão Inauguração', 'N/A')}")
            st.markdown(f"**Pipeline:** {posto.get('Pipeline', 'N/A')}")
            st.markdown(f"**Coordenadas GPS:** {posto.get('Latitude', '-')}, {posto.get('Longitude', '-')}")
            st.markdown("</div>", unsafe_allow_html=True)

        with col_p3:
            st.markdown("<div class='procv-card'>", unsafe_allow_html=True)
            st.markdown("### 📞 Fila do Call Center & Treinamentos")
            st.markdown(f"**Tipo de Necessidade:** {posto.get('Tipo_Necessidade', '-')}")
            st.markdown(f"**Instrutor Sugerido:** {posto.get('Instrutor_Sugerido', '-')}")
            st.markdown(f"**Semana Sugerida:** {posto.get('Semana_Sugerida', '-')}")
            st.markdown(f"**Status do Contato:** {posto.get('Status_Contato', '-')}")
            st.markdown(f"**Último Treinamento:** {posto.get('Data_Ultimo_Treinamento', 'Sem Registro')}")
            st.markdown("</div>", unsafe_allow_html=True)

    else:
        st.info("💡 Clique em qualquer linha da tabela acima para disparar o PROCV em todas as colunas.")
