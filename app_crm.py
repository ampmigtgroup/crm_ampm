import streamlit as st
import pandas as pd
import os

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="CRM Operacional AmPm — Sistema Integrado",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização visual (Padrão AmPm)
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
    .top-instructor {
        background-color: #262b36;
        padding: 12px;
        border-radius: 8px;
        margin-bottom: 8px;
        border-left: 4px solid #4CAF50;
    }
    </style>
""", unsafe_allow_html=True)

# --- CARREGAMENTO E PROCV MULTICOLUNA INTEGRADO ---
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
            df_rec = pd.read_excel(xls, sheet_name='Recomendacao_Deslocamento')
            
            # Normalização de chaves numéricas
            df_lojas['PV Abadi'] = pd.to_numeric(df_lojas['PV Abadi'], errors='coerce')
            df_fila['PV_Abadi'] = pd.to_numeric(df_fila['PV_Abadi'], errors='coerce')
            df_inaug['PV ABADI'] = pd.to_numeric(df_inaug['PV ABADI'], errors='coerce')
            df_rec['PV_ABADI'] = pd.to_numeric(df_rec['PV_ABADI'], errors='coerce')
            
            # PROCV 1: Cruzando Rede de Lojas com Call Center
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
            
            # Tratamento de valores nulos
            df_base['Status_Contato'] = df_base['Status_Contato'].fillna('A Contatar')
            df_base['Tipo_Necessidade'] = df_base['Tipo_Necessidade'].fillna('Rede Ativa (Sem Pendência)')
            df_base['Instrutor_Sugerido'] = df_base['Instrutor_Sugerido'].fillna('Pendente de Alocação')
            
            return df_base, df_instrutores, df_rec
        except Exception as e:
            st.error(f"Erro ao ler planilhas: {e}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    else:
        st.warning("⚠️ Arquivo 'Base_Unificada_AmPm.xlsx' não localizado.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

df_base, df_instrutores, df_rec = carregar_bases_integradas()

# --- MENU LATERAL DE NAVEGAÇÃO ---
st.sidebar.title("⛽ Menu CRM AmPm")
modulo = st.sidebar.radio(
    "Selecione o Módulo:",
    [
        "📊 Dashboard Executivo", 
        "🔍 PROCV & Gestão de Lojas", 
        "📍 Menor Custo & Geodeslocamento (Top 3)", 
        "📞 Fila Call Center & Contatos", 
        "👔 Gestão de Instrutores"
    ]
)

st.sidebar.divider()
st.sidebar.markdown("**Status do Sistema:** Operacional 🟢")
st.sidebar.markdown(f"**Lojas na Base:** {len(df_base)}")

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
            st.subheader("📍 Lojas por UF")
            st.bar_chart(df_base['UF'].value_counts().head(10))
        with col_B:
            st.subheader("📋 Status no Call Center")
            st.bar_chart(df_base['Status_Contato'].value_counts())

# ==========================================
# MÓDULO 2: PROCV & GESTÃO DE LOJAS
# ==========================================
elif modulo == "🔍 PROCV & Gestão de Lojas":
    st.title("🔍 CRM PROCV — Busca Multicoluna")
    if not df_base.empty:
        col_busca, col_uf = st.columns([3, 1])
        with col_busca:
            termo = st.text_input("Busca (PV Abadi, Nome ou Cidade):", "")
        with col_uf:
            f_uf = st.selectbox("UF:", ["Todas"] + sorted([str(x) for x in df_base['UF'].dropna().unique()]))
            
        df_view = df_base.copy()
        if termo:
            df_view = df_view[
                df_view['Razão Social'].astype(str).str.contains(termo, case=False, na=False) |
                df_view['PV Abadi'].astype(str).str.contains(termo, na=False) |
                df_view['Municipio'].astype(str).str.contains(termo, case=False, na=False)
            ]
        if f_uf != "Todas":
            df_view = df_view[df_view['UF'] == f_uf]
            
        colunas_tabela = ['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Status Loja', 'Tipo_Necessidade', 'Instrutor_Sugerido', 'Status_Contato']
        
        evento = st.dataframe(
            df_view[colunas_tabela],
            use_container_width=True,
            hide_index=True,
            selection_mode="single-row",
            on_select="rerun"
        )
        
        linhas = evento.selection.get("rows", [])
        if linhas:
            p = df_view.iloc[linhas[0]].to_dict()
            st.divider()
            st.subheader(f"📋 Painel PROCV Completo — PV: {p['PV Abadi']} | {p['Razão Social']}")
            
            k1, k2, k3 = st.columns(3)
            with k1:
                st.markdown("<div class='procv-card'>", unsafe_allow_html=True)
                st.markdown("### 🏪 Cadastro")
                st.markdown(f"**Status Loja:** {p.get('Status Loja', '-')}")
                st.markdown(f"**Endereço:** {p.get('Endereço', '-')}")
                st.markdown(f"**Município/UF:** {p.get('Municipio', '-')}/{p.get('UF', '-')}")
                st.markdown("</div>", unsafe_allow_html=True)
            with k2:
                st.markdown("<div class='procv-card'>", unsafe_allow_html=True)
                st.markdown("### 👔 Franquia")
                st.markdown(f"**Gerência (GF):** {p.get('GF', '-')}")
                st.markdown(f"**Consultor (CF):** {p.get('CF', '-')}")
                st.markdown(f"**Previsão Inauguração:** {p.get('Previsão Inauguração', 'N/A')}")
                st.markdown("</div>", unsafe_allow_html=True)
            with k3:
                st.markdown("<div class='procv-card'>", unsafe_allow_html=True)
                st.markdown("### 📞 Atendimento")
                st.markdown(f"**Necessidade:** {p.get('Tipo_Necessidade', '-')}")
                st.markdown(f"**Instrutor Sugerido:** {p.get('Instrutor_Sugerido', '-')}")
                st.markdown(f"**Status Contato:** {p.get('Status_Contato', '-')}")
                st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# MÓDULO 3: MENOR CUSTO & GEODESLOCAMENTO (TOP 3)
# ==========================================
elif modulo == "📍 Menor Custo & Geodeslocamento (Top 3)":
    st.title("📍 Otimizador de Deslocamento — Top 3 Instrutores Mais Próximos")
    st.markdown("Cruzamento geográfico de latitude/longitude para minimização do custo de passagem e hospedagem.")
    
    if not df_rec.empty:
        # Seleção do Posto para análise logística
        postos_unicos = df_rec[['PV_ABADI', 'Razao_Social', 'Municipio_Loja', 'UF_Loja']].drop_duplicates()
        postos_unicos['label'] = postos_unicos['PV_ABADI'].astype(str) + " - " + postos_unicos['Razao_Social'] + " (" + postos_unicos['Municipio_Loja'] + "/" + postos_unicos['UF_Loja'] + ")"
        
        posto_selecionado = st.selectbox("Selecione o Posto/Cliente para analisar as rotas:", postos_unicos['label'].tolist())
        
        pv_sel = int(posto_selecionado.split(" - ")[0])
        top_3 = df_rec[df_rec['PV_ABADI'] == pv_sel].sort_values(by='Ranking_Proximidade').head(3)
        
        if not top_3.empty:
            info_posto = top_3.iloc[0]
            st.markdown(f"### ⛽ Posto: **{info_posto['Razao_Social']}** (PV: {info_posto['PV_ABADI']})")
            st.markdown(f"📍 **Localização:** {info_posto['Municipio_Loja']}/{info_posto['UF_Loja']} | **Dias de Treinamento Necessários:** {info_posto['Dias_Treinamento_Necessarios']}")
            
            st.write("")
            st.subheader("🥇 Top 3 Opções de Instrutores por Menor Distância (Menor Custo Logístico)")
            
            col1, col2, col3 = st.columns(3)
            cols = [col1, col2, col3]
            
            for idx, (_, row) in enumerate(top_3.iterrows()):
                if idx < 3:
                    with cols[idx]:
                        st.markdown(f"<div class='top-instructor'>", unsafe_allow_html=True)
                        st.markdown(f"#### #{row['Ranking_Proximidade']}º Opção — {row['Instrutor_Sugerido']}")
                        st.markdown(f"📍 **Origem Instrutor:** {row['Cidade_Instrutor']} / {row['UF_Instrutor']}")
                        st.markdown(f"📏 **Distância Linear:** `{row['Distancia_km_linha_reta']} km`")
                        st.markdown("</div>", unsafe_allow_html=True)
                        
            st.divider()
            st.markdown("### 📊 Tabela de Comparação de Custo e Deslocamento")
            st.dataframe(
                top_3[['Ranking_Proximidade', 'Instrutor_Sugerido', 'Cidade_Instrutor', 'UF_Instrutor', 'Distancia_km_linha_reta']],
                use_container_width=True,
                hide_index=True
            )

# ==========================================
# MÓDULO 4: FILA CALL CENTER
# ==========================================
elif modulo == "📞 Fila Call Center & Contatos":
    st.title("📞 Gestão da Fila do Call Center")
    if not df_base.empty:
        df_fila_view = df_base[df_base['Tipo_Necessidade'] != 'Rede Ativa (Sem Pendência)'].copy()
        st.dataframe(
            df_fila_view[['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Tipo_Necessidade', 'Instrutor_Sugerido', 'Semana_Sugerida', 'Status_Contato']],
            use_container_width=True,
            hide_index=True
        )

# ==========================================
# MÓDULO 5: GESTÃO DE INSTRUTORES
# ==========================================
elif modulo == "👔 Gestão de Instrutores":
    st.title("👔 Relação de Instrutores")
    if not df_instrutores.empty:
        st.dataframe(df_instrutores[['NOME_COMPLETO', 'STATUS', 'TELEFONE', 'EMAIL', 'Cidade', 'UF']], use_container_width=True, hide_index=True)

