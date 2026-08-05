import streamlit as st
import pandas as pd
import os

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="CRM Operacional AmPm — Sistema Integrado",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização visual refinada (Padrão AmPm)
st.markdown("""
    <style>
    .stButton>button {
        background-color: #e0a96d;
        color: #000;
        font-weight: bold;
        border-radius: 6px;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.5rem;
    }
    .procv-card {
        background-color: #1e222a;
        padding: 18px;
        border-radius: 10px;
        border-left: 5px solid #e0a96d;
        margin-bottom: 12px;
    }
    .main-header {
        font-weight: 800;
        color: #e0a96d;
    }
    </style>
""", unsafe_allow_html=True)

# --- CARREGAMENTO E CRUZAMENTO DE DADOS (PROCV INTEGRADO) ---
@st.cache_data
def carregar_bases_integradas():
    caminho = "Base_Unificada_AmPm.xlsx"
    if os.path.exists(caminho):
        try:
            xls = pd.ExcelFile(caminho, engine='openpyxl')
            df_lojas = pd.read_excel(xls, sheet_name='Rede_de_Lojas')
            df_fila = pd.read_excel(xls, sheet_name='Fila_CallCenter')
            df_inaug = pd.read_excel(xls, sheet_name='Previsao_Inauguracao')
            df_instrutores = pd.read_excel(xls, sheet_name='Instrutores')
            df_treinamentos = pd.read_excel(xls, sheet_name='Treinamentos_Feitos')
            
            # Normalização de chaves numéricas para o PROCV
            df_lojas['PV Abadi'] = pd.to_numeric(df_lojas['PV Abadi'], errors='coerce')
            df_fila['PV_Abadi'] = pd.to_numeric(df_fila['PV_Abadi'], errors='coerce')
            df_inaug['PV ABADI'] = pd.to_numeric(df_inaug['PV ABADI'], errors='coerce')
            
            # PROCV 1: Cruzando Rede de Lojas com Fila do Call Center
            df_base = pd.merge(
                df_lojas,
                df_fila[['PV_Abadi', 'Tipo_Necessidade', 'Data_Ultimo_Treinamento', 
                         'Dias_desde_Ultimo_Treinamento', 'Instrutor_Sugerido', 
                         'Semana_Sugerida', 'Status_Contato', 'Observacoes']],
                left_on='PV Abadi', right_on='PV_Abadi', how='left'
            )
            
            # PROCV 2: Cruzando com Previsão de Inaugurações
            df_base = pd.merge(
                df_base,
                df_inaug[['PV ABADI', 'Previsão Inauguração', 'Pipeline', 'Consultor_Possivel_Instrutor']],
                left_on='PV Abadi', right_on='PV ABADI', how='left'
            )
            
            # Tratamento de valores nulos nas colunas PROCV
            df_base['Status_Contato'] = df_base['Status_Contato'].fillna('A Contatar')
            df_base['Tipo_Necessidade'] = df_base['Tipo_Necessidade'].fillna('Rede Ativa (Sem Pendência)')
            df_base['Instrutor_Sugerido'] = df_base['Instrutor_Sugerido'].fillna('Pendente de Alocação')
            
            return df_base, df_instrutores, df_treinamentos
        except Exception as e:
            st.error(f"Erro ao processar as planilhas: {e}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    else:
        st.warning("⚠️ Arquivo 'Base_Unificada_AmPm.xlsx' não encontrado.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

df_base, df_instrutores, df_treinamentos = carregar_bases_integradas()

# --- MENU LATERAL DE NAVEGAÇÃO ---
st.sidebar.title("⛽ Menu CRM AmPm")
modulo = st.sidebar.radio(
    "Selecione o Módulo:",
    ["📊 Dashboard Executivo", "🔍 PROCV & Gestão de Lojas", "📞 Fila Call Center & Contatos", "👔 Gestão de Instrutores"]
)

st.sidebar.divider()
st.sidebar.markdown("**Status do Sistema:** Operacional 🟢")
st.sidebar.markdown(f"**Total de Lojas na Base:** {len(df_base)}")

# ==========================================
# MÓDULO 1: DASHBOARD EXECUTIVO
# ==========================================
if modulo == "📊 Dashboard Executivo":
    st.title("📊 Dashboard Executivo — AmPm")
    
    if not df_base.empty:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Lojas na Rede", len(df_base))
        c2.metric("Fila CallCenter Ativa", len(df_base[df_base['Tipo_Necessidade'] != 'Rede Ativa (Sem Pendência)']))
        c3.metric("Postos a Contatar", len(df_base[df_base['Status_Contato'] == 'A Contatar']))
        c4.metric("Inaugurações Previstas", len(df_base[df_base['Previsão Inauguração'].notna()]))
        
        st.divider()
        
        col_A, col_B = st.columns(2)
        
        with col_A:
            st.subheader("📍 Distribuição por UF (Rede Completa)")
            if 'UF' in df_base.columns:
                uf_counts = df_base['UF'].value_counts().head(10)
                st.bar_chart(uf_counts)
                
        with col_B:
            st.subheader("📋 Status dos Contatos no Call Center")
            if 'Status_Contato' in df_base.columns:
                status_counts = df_base['Status_Contato'].value_counts()
                st.bar_chart(status_counts)

# ==========================================
# MÓDULO 2: PROCV & GESTÃO DE LOJAS
# ==========================================
elif modulo == "🔍 PROCV & Gestão de Lojas":
    st.title("🔍 CRM PROCV — Cruzamento por Coluna de Postos")
    st.markdown("Pesquise por qualquer posto para disparar o **PROCV automático** em todas as abas da base unificada.")
    
    if not df_base.empty:
        col_busca, col_uf, col_status = st.columns([2, 1, 1])
        with col_busca:
            termo = st.text_input("Busca rápida (PV Abadi, Nome ou Cidade):", "")
        with col_uf:
            lista_ufs = ["Todas"] + sorted([str(x) for x in df_base['UF'].dropna().unique()])
            f_uf = st.selectbox("Filtrar UF:", lista_ufs)
        with col_status:
            lista_status = ["Todos"] + sorted([str(x) for x in df_base['Status_Contato'].dropna().unique()])
            f_status = st.selectbox("Filtrar Status:", lista_status)
            
        df_view = df_base.copy()
        if termo:
            df_view = df_view[
                df_view['Razão Social'].astype(str).str.contains(termo, case=False, na=False) |
                df_view['PV Abadi'].astype(str).str.contains(termo, na=False) |
                df_view['Municipio'].astype(str).str.contains(termo, case=False, na=False)
            ]
        if f_uf != "Todas":
            df_view = df_view[df_view['UF'] == f_uf]
        if f_status != "Todos":
            df_view = df_view[df_view['Status_Contato'] == f_status]
            
        st.markdown(f"**Registros encontrados:** {len(df_view)}")
        
        colunas_tabela = ['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Status Loja', 'Tipo_Necessidade', 'Instrutor_Sugerido', 'Status_Contato']
        
        evento = st.dataframe(
            df_view[colunas_tabela],
            use_container_width=True,
            hide_index=True,
            selection_mode="single-row",
            on_select="rerun"
        )
        
        linhas_selecionadas = evento.selection.get("rows", [])
        
        if linhas_selecionadas:
            idx = linhas_selecionadas[0]
            p = df_view.iloc[idx].to_dict()
            
            st.divider()
            st.subheader(f"📋 Painel PROCV Detalhado — PV: {p['PV Abadi']} | {p['Razão Social']}")
            
            k1, k2, k3 = st.columns(3)
            with k1:
                st.markdown("<div class='procv-card'>", unsafe_allow_html=True)
                st.markdown("### 🏪 Dados Cadastrais")
                st.markdown(f"**Razão Social:** {p.get('Razão Social', '-')}")
                st.markdown(f"**Status Loja:** {p.get('Status Loja', '-')}")
                st.markdown(f"**Endereço:** {p.get('Endereço', '-')}")
                st.markdown(f"**Município/UF:** {p.get('Municipio', '-')}/{p.get('UF', '-')}")
                st.markdown(f"**CEP:** {p.get('CEP', '-')}")
                st.markdown("</div>", unsafe_allow_html=True)
                
            with k2:
                st.markdown("<div class='procv-card'>", unsafe_allow_html=True)
                st.markdown("### 👔 Franquia & Modelo")
                st.markdown(f"**Modelo Gerencial:** {p.get('Modelo Loja Gerencial', '-')}")
                st.markdown(f"**Gerência (GF):** {p.get('GF', '-')}")
                st.markdown(f"**Consultor (CF):** {p.get('CF', '-')}")
                st.markdown(f"**Previsão Inauguração:** {p.get('Previsão Inauguração', 'N/A')}")
                st.markdown(f"**Pipeline:** {p.get('Pipeline', 'N/A')}")
                st.markdown("</div>", unsafe_allow_html=True)
                
            with k3:
                st.markdown("<div class='procv-card'>", unsafe_allow_html=True)
                st.markdown("### 📞 Atendimento & Treinamento")
                st.markdown(f"**Necessidade:** {p.get('Tipo_Necessidade', '-')}")
                st.markdown(f"**Instrutor Sugerido:** {p.get('Instrutor_Sugerido', '-')}")
                st.markdown(f"**Semana Programada:** {p.get('Semana_Sugerida', '-')}")
                st.markdown(f"**Status do Contato:** {p.get('Status_Contato', '-')}")
                st.markdown(f"**Último Treinamento:** {p.get('Data_Ultimo_Treinamento', 'Sem Registro')}")
                st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.info("💡 Clique em uma linha da tabela acima para carregar o PROCV completo do posto selecionado.")

# ==========================================
# MÓDULO 3: FILA CALL CENTER
# ==========================================
elif modulo == "📞 Fila Call Center & Contatos":
    st.title("📞 Gestão da Fila de Call Center")
    st.markdown("Acompanhamento das lojas agendadas, pendentes de contato e demandas prioritárias.")
    
    if not df_base.empty:
        df_fila_view = df_base[df_base['Tipo_Necessidade'] != 'Rede Ativa (Sem Pendência)'].copy()
        
        st.markdown(f"**Total de Postos na Fila:** {len(df_fila_view)}")
        
        cols_fila = ['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Tipo_Necessidade', 'Instrutor_Sugerido', 'Semana_Sugerida', 'Status_Contato']
        st.dataframe(df_fila_view[cols_fila], use_container_width=True, hide_index=True)

# ==========================================
# MÓDULO 4: GESTÃO DE INSTRUTORES
# ==========================================
elif modulo == "👔 Gestão de Instrutores":
    st.title("👔 Quadro de Instrutores AmPm")
    st.markdown("Relação de instrutores cadastrados, status e contatos operacionais.")
    
    if not df_instrutores.empty:
        colunas_inst = ['NOME_COMPLETO', 'STATUS', 'TELEFONE', 'EMAIL', 'Cidade', 'UF']
        colunas_disponiveis = [c for c in colunas_inst if c in df_instrutores.columns]
        st.dataframe(df_instrutores[colunas_disponiveis], use_container_width=True, hide_index=True)
    else:
        st.warning("Nenhum dado de instrutor encontrado.")
