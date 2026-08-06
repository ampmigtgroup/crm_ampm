import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
import pydeck as pdk
import io

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="CRM Operacional AmPm",
    page_icon="⛽",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 2. ESTILIZAÇÃO CSS CUSTOMIZADA (DESIGN SYSTEM AMPM)
# ==========================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* Topbar Premium */
    .main-header {
        background: linear-gradient(135deg, #E27B00 0%, #FF9800 50%, #D32F2F 100%);
        padding: 24px 28px;
        border-radius: 16px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 8px 24px rgba(226, 123, 0, 0.25);
    }
    .main-header h1 {
        color: white !important;
        margin: 0 0 6px 0;
        font-weight: 700;
        font-size: 2.2rem;
    }
    .main-header p {
        margin: 0;
        font-size: 1.05rem;
        opacity: 0.95;
    }

    /* KPI Cards */
    .kpi-card {
        background-color: #1E222A;
        border-radius: 12px;
        padding: 20px;
        border: 1px solid #2D333F;
        border-left: 6px solid #E27B00;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
    }
    .kpi-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .kpi-title {
        font-size: 0.8rem;
        color: #A0AAB8;
        text-transform: uppercase;
        font-weight: 700;
        letter-spacing: 0.8px;
    }
    .kpi-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #FFFFFF;
        margin-top: 8px;
    }

    /* Estilização do Pipeline AmPm */
    .ampm-column {
        background-color: #14171D;
        border-radius: 12px;
        padding: 14px;
        border: 1px solid #2D333F;
        min-height: 500px;
    }
    .ampm-title {
        font-size: 0.95rem;
        font-weight: 700;
        margin-bottom: 12px;
        padding-bottom: 8px;
        border-bottom: 2px solid #2D333F;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }

    /* Cards de Informação PROCV e Call Center */
    .procv-card {
        background-color: #1A1D24;
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #2D333F;
        border-top: 4px solid #E27B00;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15);
        margin-bottom: 15px;
    }
    .procv-card h4 {
        margin-top: 0;
        margin-bottom: 12px;
        color: #FF9800;
        font-size: 1rem;
    }
    .procv-card p {
        margin: 4px 0;
        font-size: 0.9rem;
    }
    
    .top-instructor-card {
        background-color: #1A1D24;
        padding: 18px;
        border-radius: 12px;
        border: 1px solid #2D333F;
        border-left: 5px solid #4CAF50;
        margin-bottom: 14px;
    }

    /* Timeline de Atendimentos */
    .timeline-item {
        border-left: 3px solid #E27B00;
        padding-left: 15px;
        margin-bottom: 15px;
        position: relative;
    }

    /* Badges */
    .badge-info {
        background: rgba(226, 123, 0, 0.15);
        color: #FF9800;
        border: 1px solid #E27B00;
        padding: 3px 8px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.78rem;
    }

    /* Botões Customizados */
    .stButton>button {
        background: linear-gradient(90deg, #E27B00 0%, #FF9800 100%);
        color: #FFFFFF !important;
        font-weight: 600;
        border: none;
        border-radius: 8px;
        padding: 10px 24px;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        box-shadow: 0 4px 14px rgba(226, 123, 0, 0.5);
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. FUNÇÕES INTELIGENTES DE CARREGAMENTO E MAPEAMENTO
# ==========================================

KEYWORDS_CRM = {'pv', 'abadi', 'posto', 'razao', 'razão', 'loja', 'instrutor', 'treinamento', 'consultor', 'ampm', 'am_pm', 'uf'}

def normalizar_nome_coluna(col):
    return str(col).strip().lower().replace("_", " ").replace("-", " ")

def encontrar_coluna(df, candidatos):
    colunas_df = {normalizar_nome_coluna(c): c for c in df.columns}
    for cand in candidatos:
        cand_norm = normalizar_nome_coluna(cand)
        if cand_norm in colunas_df:
            return colunas_df[cand_norm]
    # Busca parcial por substring
    for col_norm, col_orig in colunas_df.items():
        for cand in candidatos:
            if normalizar_nome_coluna(cand) in col_norm:
                return col_orig
    return None

def validar_relevancia_crm(df):
    """Verifica se a planilha contém termos chave do CRM AmPm."""
    cols_norm = [normalizar_nome_coluna(c) for c in df.columns]
    for col in cols_norm:
        for kw in KEYWORDS_CRM:
            if kw in col:
                return True
    return False

def garantir_colunas_padrao(df):
    """Garante a existência de todas as colunas operacionais com valores default."""
    df = df.copy()
    
    # Mapear e padronizar nomes chave
    col_pv = encontrar_coluna(df, ['pv abadi', 'pv_abadi', 'pv', 'codigo', 'código', 'posto'])
    if col_pv and col_pv != 'PV Abadi':
        df['PV Abadi'] = df[col_pv]
    elif 'PV Abadi' not in df.columns:
        df['PV Abadi'] = range(1000, 1000 + len(df))

    col_rs = encontrar_coluna(df, ['razão social', 'razao social', 'razao_social', 'nome fantasia', 'posto'])
    if col_rs and col_rs != 'Razão Social':
        df['Razão Social'] = df[col_rs]
    elif 'Razão Social' not in df.columns:
        df['Razão Social'] = "Posto AmPm " + df['PV Abadi'].astype(str)

    col_uf = encontrar_coluna(df, ['uf', 'estado'])
    if col_uf and col_uf != 'UF':
        df['UF'] = df[col_uf]
    elif 'UF' not in df.columns:
        df['UF'] = "SP"

    col_mun = encontrar_coluna(df, ['municipio', 'município', 'cidade'])
    if col_mun and col_mun != 'Municipio':
        df['Municipio'] = df[col_mun]
    elif 'Municipio' not in df.columns:
        df['Municipio'] = "Não Informado"

    col_cf = encontrar_coluna(df, ['cf', 'consultor', 'consultor_cf'])
    if col_cf and col_cf != 'CF':
        df['CF'] = df[col_cf]
    elif 'CF' not in df.columns:
        df['CF'] = "Consultor Geral"

    col_gf = encontrar_coluna(df, ['gf', 'gerente', 'gerencia'])
    if col_gf and col_gf != 'GF':
        df['GF'] = df[col_gf]
    elif 'GF' not in df.columns:
        df['GF'] = "Gerência Regional"

    col_end = encontrar_coluna(df, ['endereço', 'endereco', 'logradouro'])
    if col_end and col_end != 'Endereço':
        df['Endereço'] = df[col_end]
    elif 'Endereço' not in df.columns:
        df['Endereço'] = "Endereço não cadastrado"

    col_status_loja = encontrar_coluna(df, ['status loja', 'status_loja', 'situacao'])
    if col_status_loja and col_status_loja != 'Status Loja':
        df['Status Loja'] = df[col_status_loja]
    elif 'Status Loja' not in df.columns:
        df['Status Loja'] = "Ativa"

    # Colunas de Gestão do Call Center
    campos_default = {
        'Status_Contato': 'A Contatar',
        'Tipo_Necessidade': 'Treinamento de Rede',
        'Instrutor_Sugerido': 'Pendente de Alocação',
        'Dias_desde_Ultimo_Treinamento': 0,
        'Previsão Inauguração': None,
        'Nome_Contato': '',
        'Telefone_Contato': '',
        'Qtd_Funcionarios': 0,
        'Material_Em_Loja': 'Não Informado',
        'Data_Agendada': None,
        'Observacoes': '',
        'Data_do_Contato': 'Sem registro'
    }

    for col, val in campos_default.items():
        c_existente = encontrar_coluna(df, [col])
        if c_existente and c_existente != col:
            df[col] = df[c_existente]
        elif col not in df.columns:
            df[col] = val

    df['Status_Contato'] = df['Status_Contato'].fillna('A Contatar')
    df['Tipo_Necessidade'] = df['Tipo_Necessidade'].fillna('Treinamento de Rede')
    df['Instrutor_Sugerido'] = df['Instrutor_Sugerido'].fillna('Pendente de Alocação')

    return df

@st.cache_data(show_spinner=False)
def carregar_bases_dinamicas(uploaded_file_or_path="Base_Unificada_AmPm.xlsx"):
    """
    Carrega dinamicamente qualquer arquivo Excel ou CSV, vasculhando abas
    e validando se pertence ao contexto do CRM AmPm.
    """
    dict_dfs = {}
    
    try:
        if isinstance(uploaded_file_or_path, str):
            if not os.path.exists(uploaded_file_or_path):
                return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
            if uploaded_file_or_path.endswith('.csv'):
                dict_dfs['CSV_Data'] = pd.read_csv(uploaded_file_or_path)
            else:
                xls = pd.ExcelFile(uploaded_file_or_path, engine='openpyxl')
                for sheet in xls.sheet_names:
                    dict_dfs[sheet] = pd.read_excel(xls, sheet_name=sheet)
        else:
            nome = uploaded_file_or_path.name.lower()
            if nome.endswith('.csv'):
                dict_dfs['CSV_Data'] = pd.read_csv(uploaded_file_or_path)
            else:
                xls = pd.ExcelFile(uploaded_file_or_path, engine='openpyxl')
                for sheet in xls.sheet_names:
                    dict_dfs[sheet] = pd.read_excel(xls, sheet_name=sheet)

        if not dict_dfs:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        # 1. Identificar a aba principal das Lojas / Base Operacional
        df_base_raw = None
        df_instrutores_raw = pd.DataFrame()
        df_rec_raw = pd.DataFrame()

        # Vasculhar abas procurando por Lojas, Instrutores e Recomendações
        for sheet_name, df_sheet in dict_dfs.items():
            df_sheet.columns = [str(c).strip() for c in df_sheet.columns]
            
            # Checar se é a aba de Instrutores
            if any(k in sheet_name.lower() for k in ['instrutor', 'equipe']) or encontrar_coluna(df_sheet, ['latitude', 'longitude', 'email']):
                df_instrutores_raw = df_sheet.copy()
            # Checar se é a aba de Recomendação / Distâncias
            elif any(k in sheet_name.lower() for k in ['recomendacao', 'deslocamento', 'ranking', 'distancia']):
                df_rec_raw = df_sheet.copy()
            # Se tiver colunas do CRM, define como base principal
            elif validar_relevancia_crm(df_sheet) and df_base_raw is None:
                df_base_raw = df_sheet.copy()

        # Se nenhuma aba específica foi filtrada, mas existe uma aba relevante
        if df_base_raw is None:
            for sheet_name, df_sheet in dict_dfs.items():
                if validar_relevancia_crm(df_sheet):
                    df_base_raw = df_sheet.copy()
                    break

        # Se ainda assim não houver relação com o CRM, rejeita
        if df_base_raw is None or df_base_raw.empty:
            st.error("⚠️ **Arquivo não reconhecido:** O arquivo enviado não possui estrutura ou colunas compatíveis com o CRM AmPm (ex: 'PV', 'Razão Social', 'UF' ou 'Instrutor').")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        # Garantir tratamento e padronização completa dos dados
        df_base_final = garantir_colunas_padrao(df_base_raw)

        # Tratar Instrutores se disponível
        if not df_instrutores_raw.empty:
            col_nome_inst = encontrar_coluna(df_instrutores_raw, ['nome', 'nome_completo', 'instrutor'])
            if col_nome_inst and col_nome_inst != 'NOME_COMPLETO':
                df_instrutores_raw['NOME_COMPLETO'] = df_instrutores_raw[col_nome_inst]
        else:
            df_instrutores_raw = pd.DataFrame({
                'NOME_COMPLETO': df_base_final['Instrutor_Sugerido'].unique(),
                'STATUS': 'Ativo',
                'TELEFONE': '(11) 99999-0000',
                'EMAIL': 'capacitacao@ampm.com.br',
                'Cidade': 'São Paulo',
                'UF': 'SP'
            })

        # Tratar Recomendações de Deslocamento se disponível
        if not df_rec_raw.empty:
            col_pv_rec = encontrar_coluna(df_rec_raw, ['pv_abadi', 'pv abadi', 'pv', 'codigo'])
            if col_pv_rec:
                df_rec_raw['PV_ABADI'] = pd.to_numeric(df_rec_raw[col_pv_rec], errors='coerce')
        else:
            # Gerar estrutura simulada de apoio
            df_rec_raw = pd.DataFrame({
                'PV_ABADI': df_base_final['PV Abadi'],
                'Razao_Social': df_base_final['Razão Social'],
                'Municipio_Loja': df_base_final['Municipio'],
                'UF_Loja': df_base_final['UF'],
                'Ranking_Proximidade': 1,
                'Instrutor_Sugerido': df_base_final['Instrutor_Sugerido'],
                'Cidade_Instrutor': 'São Paulo',
                'UF_Instrutor': 'SP',
                'Distancia_km_linha_reta': 120,
                'Dias_Treinamento_Necessarios': 3,
                'Lat_Loja': -23.5505,
                'Lon_Loja': -46.6333,
                'Lat_Instrutor': -23.5615,
                'Lon_Instrutor': -46.6559
            })

        return df_base_final, df_instrutores_raw, df_rec_raw

    except Exception as e:
        st.error(f"⚠️ Erro ao processar o arquivo enviado: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

def salvar_alteracoes_disco():
    caminho = "Base_Unificada_AmPm.xlsx"
    if 'df_base' in st.session_state and not st.session_state['df_base'].empty:
        try:
            with pd.ExcelWriter(caminho, engine='openpyxl', mode='w') as writer:
                st.session_state['df_base'].to_excel(writer, sheet_name='Base_CRM', index=False)
                if 'df_instrutores' in st.session_state and not st.session_state['df_instrutores'].empty:
                    st.session_state['df_instrutores'].to_excel(writer, sheet_name='Instrutores', index=False)
            st.toast("💾 Dados salvos no arquivo Excel com sucesso!", icon="✅")
        except Exception as e:
            st.toast("💾 Dados atualizados na sessão atual!", icon="ℹ️")

# ==========================================
# 4. INICIALIZAÇÃO DO ESTADO DA SESSÃO
# ==========================================
if 'df_base' not in st.session_state or st.session_state['df_base'].empty:
    b, i, r = carregar_bases_dinamicas("Base_Unificada_AmPm.xlsx")
    st.session_state['df_base'] = b
    st.session_state['df_instrutores'] = i
    st.session_state['df_rec'] = r

df_base_raw = st.session_state['df_base']
df_instrutores = st.session_state['df_instrutores']
df_rec = st.session_state['df_rec']

# ==========================================
# 5. SIDEBAR DE NAVEGAÇÃO E UPLOAD DINÂMICO
# ==========================================
with st.sidebar:
    st.markdown("## ⛽ **CRM AmPm**")
    st.caption("🌐 *Plataforma Integrada de Operações*")
    st.divider()
    
    modulo = st.radio(
        "📌 **Módulos do Sistema:**",
        [
            "📊 Dashboard Executivo", 
            "📋 Pipeline AmPm", 
            "🔍 PROCV & Filtros Avançados", 
            "📍 Calculadora & Otimizador de Custos", 
            "📞 Call Center & Timeline WhatsApp", 
            "👔 Equipe de Instrutores",
            "📂 Relatórios & Exportação"
        ]
    )
    
    st.divider()
    
    # UPLOAD INTELIGENTE
    st.markdown("📥 **Atualizar Banco de Dados**")
    uploaded_file = st.file_uploader(
        "Envie qualquer planilha (.xlsx, .xls ou .csv):", 
        type=["xlsx", "xls", "csv"], 
        help="O sistema analisa automaticamente a estrutura e valida os dados do CRM."
    )
    
    if uploaded_file is not None:
        b_new, i_new, r_new = carregar_bases_dinamicas(uploaded_file)
        if not b_new.empty:
            st.session_state['df_base'] = b_new
            st.session_state['df_instrutores'] = i_new
            st.session_state['df_rec'] = r_new
            st.success("✅ Base validada e carregada com sucesso!")
            st.rerun()

    st.divider()
    
    # FILTROS GLOBAIS
    st.markdown("🎯 **Filtros Globais**")
    uf_opcoes = ["Todas"] + sorted([str(x) for x in df_base_raw['UF'].dropna().unique()]) if 'UF' in df_base_raw.columns and not df_base_raw.empty else ["Todas"]
    filtro_uf = st.selectbox("Filtrar Estado (UF):", uf_opcoes)

    cf_opcoes = ["Todos"] + sorted([str(x) for x in df_base_raw['CF'].dropna().unique()]) if 'CF' in df_base_raw.columns and not df_base_raw.empty else ["Todos"]
    filtro_cf = st.selectbox("Filtrar Consultor (CF):", cf_opcoes)

    st.divider()
    st.markdown("📶 **Status:** `Operacional 🟢`")
    st.markdown(f"🏪 **Rede Total:** `{len(df_base_raw)} Unidades`")

# APLICAÇÃO DOS FILTROS GLOBAIS
df_base = df_base_raw.copy()
if not df_base.empty:
    if filtro_uf != "Todas" and 'UF' in df_base.columns:
        df_base = df_base[df_base['UF'] == filtro_uf]
    if filtro_cf != "Todos" and 'CF' in df_base.columns:
        df_base = df_base[df_base['CF'] == filtro_cf]

# HEADER GLOBAL
st.markdown("""
    <div class="main-header">
        <h1>⛽ CRM Operacional AmPm</h1>
        <p>🚀 Gestão Estratégica de Capacitação, Logística de Viagens e Atendimento da Rede</p>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# MÓDULO 1: DASHBOARD EXECUTIVO
# ==========================================
if modulo == "📊 Dashboard Executivo":
    if not df_base.empty:
        c1, c2, c3, c4 = st.columns(4)
        
        with c1:
            st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #E27B00;">
                    <div class="kpi-header"><span class="kpi-title">Rede Filtrada</span><span>🏪</span></div>
                    <div class="kpi-value">{len(df_base)}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with c2:
            pendentes = len(df_base[df_base['Status_Contato'] != 'Treinamento Realizado']) if 'Status_Contato' in df_base.columns else 0
            st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #FF9800;">
                    <div class="kpi-header"><span class="kpi-title">Fila Treinamento</span><span>🎓</span></div>
                    <div class="kpi-value">{pendentes}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with c3:
            a_contatar = len(df_base[df_base['Status_Contato'] == 'A Contatar']) if 'Status_Contato' in df_base.columns else 0
            st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #D32F2F;">
                    <div class="kpi-header"><span class="kpi-title">Pendentes Contato</span><span>📞</span></div>
                    <div class="kpi-value">{a_contatar}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with c4:
            inaug = len(df_base[df_base['Previsão Inauguração'].notna()]) if 'Previsão Inauguração' in df_base.columns else 0
            st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #0288D1;">
                    <div class="kpi-header"><span class="kpi-title">Inaugurações</span><span>🚀</span></div>
                    <div class="kpi-value">{inaug}</div>
                </div>
            """, unsafe_allow_html=True)

        st.write("")
        st.divider()
        col_A, col_B = st.columns(2)
        with col_A:
            st.subheader("🗺️ Concentração por Estado (UF)")
            if 'UF' in df_base.columns:
                st.bar_chart(df_base['UF'].value_counts().head(10), color="#E27B00")
        with col_B:
            st.subheader("📊 Situação dos Contatos no Call Center")
            if 'Status_Contato' in df_base.columns:
                st.bar_chart(df_base['Status_Contato'].value_counts(), color="#FF9800")
    else:
        st.info("💡 Por favor, envie uma planilha válida na barra lateral para visualizar o dashboard.")

# ==========================================
# MÓDULO 2: PIPELINE AMPM
# ==========================================
elif modulo == "📋 Pipeline AmPm":
    st.subheader("📋 Pipeline AmPm — Fluxo Operacional de Treinamentos")
    st.caption("Gerencie o fluxo de atendimento navegando entre os estágios de contato.")
    
    colunas_pipeline = ["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado"]
    cols_k = st.columns(len(colunas_pipeline))
    
    for idx, status in enumerate(colunas_pipeline):
        df_status = df_base[df_base['Status_Contato'] == status] if 'Status_Contato' in df_base.columns else pd.DataFrame()
        
        with cols_k[idx]:
            st.markdown(f"""
                <div class="ampm-column">
                    <div class="ampm-title">
                        <span>{status}</span>
                        <span style="background:#2D333F; padding:2px 8px; border-radius:10px; font-size:0.8rem;">{len(df_status)}</span>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            for _, item in df_status.head(6).iterrows():
                pv_val = item.get('PV Abadi', '-')
                raz_val = str(item.get('Razão Social', ''))[:14]
                with st.expander(f"📍 PV {pv_val} | {raz_val}..."):
                    st.write(f"**Cidade:** {item.get('Municipio', '-')}/{item.get('UF', '-')}")
                    st.write(f"**Necessidade:** {item.get('Tipo_Necessidade', '-')}")
                    st.write(f"**Treinandos:** {item.get('Qtd_Funcionarios', 0)} pessoas")
                    st.write(f"**Instrutor:** {item.get('Instrutor_Sugerido', 'Pendente')}")
                    
                    mudar_status = st.selectbox(
                        "Alterar Status:",
                        colunas_pipeline,
                        index=colunas_pipeline.index(status),
                        key=f"pipe_sel_{pv_val}"
                    )
                    
                    if mudar_status != status:
                        mask = st.session_state['df_base']['PV Abadi'] == item['PV Abadi']
                        st.session_state['df_base'].loc[mask, 'Status_Contato'] = mudar_status
                        salvar_alteracoes_disco()
                        st.success("Atualizado!")
                        st.rerun()

# ==========================================
# MÓDULO 3: PROCV & FILTROS AVANÇADOS
# ==========================================
elif modulo == "🔍 PROCV & Filtros Avançados":
    if not df_base.empty:
        with st.expander("🔎 **Pesquisa Avançada na Base Filtrada**", expanded=True):
            f1, f2 = st.columns(2)
            termo = f1.text_input("🔍 PV, Nome ou Município:", "")
            f_necessidade = f2.selectbox("🎯 Tipo de Necessidade:", ["Todas"] + sorted([str(x) for x in df_base['Tipo_Necessidade'].dropna().unique()])) if 'Tipo_Necessidade' in df_base.columns else ["Todas"]
            
        df_view = df_base.copy()
        if termo:
            df_view = df_view[
                df_view.get('Razão Social', pd.Series()).astype(str).str.contains(termo, case=False, na=False) |
                df_view.get('PV Abadi', pd.Series()).astype(str).str.contains(termo, na=False) |
                df_view.get('Municipio', pd.Series()).astype(str).str.contains(termo, case=False, na=False)
            ]
        if f_necessidade != "Todas" and 'Tipo_Necessidade' in df_view.columns:
            df_view = df_view[df_view['Tipo_Necessidade'] == f_necessidade]
            
        st.caption("👇 *Clique em uma linha para abrir a Ficha Detalhada PROCV:*")
        
        cols_mostrar = [c for c in ['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Status Loja', 'Tipo_Necessidade', 'Status_Contato'] if c in df_view.columns]
        evento = st.dataframe(
            df_view[cols_mostrar],
            use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun"
        )
        
        linhas = evento.selection.get("rows", [])
        if linhas:
            p = df_view.iloc[linhas[0]].to_dict()
            st.divider()
            st.markdown(f"### 📋 Ficha de Detalhes PROCV — **PV {p.get('PV Abadi', '-')} | {p.get('Razão Social', '-')}**")
            
            k1, k2, k3 = st.columns(3)
            with k1:
                st.markdown(f"""
                    <div class="procv-card">
                        <h4>🏪 Cadastro da Loja</h4>
                        <p>📍 <b>Endereço:</b> {p.get('Endereço', '-')}</p>
                        <p>🏙️ <b>Município/UF:</b> {p.get('Municipio', '-')}/{p.get('UF', '-')}</p>
                        <p>⚙️ <b>Status Loja:</b> {p.get('Status Loja', '-')}</p>
                    </div>
                """, unsafe_allow_html=True)
            with k2:
                st.markdown(f"""
                    <div class="procv-card">
                        <h4>👔 Gestão & Franquia</h4>
                        <p>👤 <b>Gerência (GF):</b> {p.get('GF', '-')}</p>
                        <p>👔 <b>Consultor (CF):</b> {p.get('CF', '-')}</p>
                        <p>📅 <b>Inauguração:</b> {p.get('Previsão Inauguração', 'N/A')}</p>
                    </div>
                """, unsafe_allow_html=True)
            with k3:
                st.markdown(f"""
                    <div class="procv-card">
                        <h4>📞 Status do Atendimento</h4>
                        <p>🎯 <b>Necessidade:</b> {p.get('Tipo_Necessidade', '-')}</p>
                        <p>👨‍🏫 <b>Instrutor Alocado:</b> {p.get('Instrutor_Sugerido', '-')}</p>
                        <p>🔄 <b>Status Contato:</b> {p.get('Status_Contato', '-')}</p>
                    </div>
                """, unsafe_allow_html=True)

# ==========================================
# MÓDULO 4: CALCULADORA & OTIMIZADOR DE CUSTOS
# ==========================================
elif modulo == "📍 Calculadora & Otimizador de Custos":
    st.subheader("📍 Análise Financeira e Otimização Logística")
    st.caption("Cálculo detalhado de custos com mapas interativos e simulação de rotas.")
    
    if not df_rec.empty:
        df_rec_filtrado = df_rec.copy()
        if filtro_uf != "Todas" and 'UF_Loja' in df_rec_filtrado.columns:
            df_rec_filtrado = df_rec_filtrado[df_rec_filtrado['UF_Loja'] == filtro_uf]

        postos_unicos = df_rec_filtrado[['PV_ABADI', 'Razao_Social', 'Municipio_Loja', 'UF_Loja']].drop_duplicates() if 'PV_ABADI' in df_rec_filtrado.columns else pd.DataFrame()
        
        if not postos_unicos.empty:
            postos_unicos['label'] = postos_unicos['PV_ABADI'].astype(str) + " - " + postos_unicos['Razao_Social'] + " (" + postos_unicos['Municipio_Loja'].astype(str) + "/" + postos_unicos['UF_Loja'].astype(str) + ")"
            
            posto_sel = st.selectbox("⛽ Selecione o Posto Alvo:", postos_unicos['label'].tolist())
            pv_sel = int(posto_sel.split(" - ")[0]) if posto_sel else 0
            
            top_3 = df_rec_filtrado[df_rec_filtrado['PV_ABADI'] == pv_sel].sort_values(by='Ranking_Proximidade').head(3)
            
            if not top_3.empty:
                st.divider()
                
                # --- MAPA 3D PYDECK ---
                primeira = top_3.iloc[0]
                if pd.notna(primeira.get('Lat_Loja')) and pd.notna(primeira.get('Lon_Loja')) and pd.notna(primeira.get('Lat_Instrutor')) and pd.notna(primeira.get('Lon_Instrutor')):
                    p_lat, p_lon = float(primeira['Lat_Loja']), float(primeira['Lon_Loja'])
                    i_lat, i_lon = float(primeira['Lat_Instrutor']), float(primeira['Lon_Instrutor'])
                    
                    df_mapa_pontos = pd.DataFrame([
                        {"name": f"Posto {primeira['PV_ABADI']}", "lat": p_lat, "lon": p_lon, "color": [226, 123, 0, 220]},
                        {"name": f"Instrutor {primeira['Instrutor_Sugerido']}", "lat": i_lat, "lon": i_lon, "color": [76, 175, 80, 220]}
                    ])
                    
                    df_mapa_arco = pd.DataFrame([{
                        "from_lat": i_lat, "from_lon": i_lon, "to_lat": p_lat, "to_lon": p_lon
                    }])
                    
                    layer_pontos = pdk.Layer(
                        "ScatterplotLayer",
                        df_mapa_pontos,
                        get_position="[lon, lat]",
                        get_color="color",
                        get_radius=20000,
                        pickable=True
                    )
                    
                    layer_arco = pdk.Layer(
                        "ArcLayer",
                        df_mapa_arco,
                        get_source_position="[from_lon, from_lat]",
                        get_target_position="[to_lon, to_lat]",
                        get_source_color=[76, 175, 80, 180],
                        get_target_color=[226, 123, 0, 180],
                        get_width=4
                    )
                    
                    view_state = pdk.ViewState(latitude=(p_lat + i_lat) / 2, longitude=(p_lon + i_lon) / 2, zoom=5, pitch=40)
                    
                    st.markdown("##### 🗺️ Visualização Geográfica do Deslocamento")
                    st.pydeck_chart(pdk.Deck(layers=[layer_pontos, layer_arco], initial_view_state=view_state, tooltip={"text": "{name}"}))

                with st.expander("⚙️ **Ajustar Parâmetros Financeiros de Viagem**", expanded=False):
                    ca1, ca2, ca3, ca4 = st.columns(4)
                    v_km = ca1.number_input("Valor KM (Terrestre R$):", value=2.10)
                    v_passagem = ca2.number_input("Passagem Aérea Média (R$):", value=1400.0)
                    v_diaria = ca3.number_input("Diária Instrutor (Hosp./Alimentação R$):", value=280.0)
                    v_traslado = ca4.number_input("Traslado/Uber Aeroporto (R$):", value=150.0)

                col1, col2, col3 = st.columns(3)
                cols = [col1, col2, col3]
                custos_calculados = []

                for idx, (_, row) in enumerate(top_3.iterrows()):
                    dist = row.get('Distancia_km_linha_reta', 100)
                    dias = row.get('Dias_Treinamento_Necessarios', 3)
                    
                    if dist <= 300:
                        modal = "Terrestre 🚗"
                        c_desloc = (dist * 2) * v_km
                        c_aereo = 0
                    else:
                        modal = "Aéreo ✈️"
                        c_desloc = v_traslado
                        c_aereo = v_passagem
                        
                    c_hospedagem = dias * v_diaria
                    custo_total = c_desloc + c_aereo + c_hospedagem
                    custos_calculados.append(custo_total)

                    if idx < 3:
                        with cols[idx]:
                            st.markdown(f"""
                                <div class="top-instructor-card">
                                    <h4 style="margin:0 0 8px 0; color:#E27B00;">#{row.get('Ranking_Proximidade', idx+1)}º {row.get('Instrutor_Sugerido', 'Instrutor')}</h4>
                                    <p style="margin:2px 0;">🏙️ <b>Origem:</b> {row.get('Cidade_Instrutor', 'SP')}/{row.get('UF_Instrutor', 'SP')}</p>
                                    <p style="margin:2px 0;">📏 <b>Distância:</b> <code>{dist} km</code></p>
                                    <p style="margin:2px 0;">✈️ <b>Modal:</b> {modal}</p>
                                    <hr style="border-color:#2D333F; margin:8px 0;">
                                    <p style="margin:2px 0; font-size:0.85rem;">• Deslocamento: R$ {c_desloc:.2f}</p>
                                    <p style="margin:2px 0; font-size:0.85rem;">• Passagem Aérea: R$ {c_aereo:.2f}</p>
                                    <p style="margin:2px 0; font-size:0.85rem;">• Diárias ({dias}d): R$ {c_hospedagem:.2f}</p>
                                    <h3 style="color:#4CAF50; margin:10px 0 0 0;">Total: R$ {custo_total:.2f}</h3>
                                </div>
                            """, unsafe_allow_html=True)

                if len(custos_calculados) >= 2:
                    economia = custos_calculados[1] - custos_calculados[0]
                    st.success(f"💡 **Economia Eficiente:** Optar pelo **1º Instrutor Recomendado** garante uma economia estimada de **R$ {economia:.2f}** nesta operação.")

# ==========================================
# MÓDULO 5: CALL CENTER & TIMELINE WHATSAPP
# ==========================================
elif modulo == "📞 Call Center & Timeline WhatsApp":
    if not df_base.empty:
        df_fila_view = df_base.copy()
        
        c_left, c_right = st.columns([1.2, 1.8])
        
        with c_left:
            st.subheader("📋 Fila de Atendimento")
            cols_call = [c for c in ['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Status_Contato'] if c in df_fila_view.columns]
            evento_call = st.dataframe(
                df_fila_view[cols_call],
                use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun"
            )
            selecionado = evento_call.selection.get("rows", [])
            
        with c_right:
            if selecionado:
                posto = df_fila_view.iloc[selecionado[0]]
                pv_alvo = posto.get('PV Abadi')
                tel_limpo = ''.join(filter(str.isdigit, str(posto.get('Telefone_Contato', ''))))
                
                st.markdown(f"### 📝 Ficha de Atendimento — **PV {posto.get('PV Abadi', '-')}**")
                
                st.markdown(f"""
                    <div class="procv-card">
                        <h4>🏪 Contexto do Posto (Consulta Rápida)</h4>
                        <div style="display: flex; justify-content: space-between; flex-wrap: wrap;">
                            <div style="flex: 1; min-width: 200px;">
                                <p>🏬 <b>Razão Social:</b> {posto.get('Razão Social', '-')}</p>
                                <p>📍 <b>Cidade/UF:</b> {posto.get('Municipio', '-')}/{posto.get('UF', '-')}</p>
                                <p>🏠 <b>Endereço:</b> {posto.get('Endereço', '-')}</p>
                            </div>
                            <div style="flex: 1; min-width: 200px;">
                                <p>👔 <b>Consultor (CF):</b> {posto.get('CF', '-')}</p>
                                <p>🎯 <b>Necessidade:</b> <span class="badge-info">{posto.get('Tipo_Necessidade', '-')}</span></p>
                                <p>⏱️ <b>Dias sem Treinamento:</b> {posto.get('Dias_desde_Ultimo_Treinamento', 'N/A')}</p>
                                <p>📅 <b>Inauguração Prevista:</b> {posto.get('Previsão Inauguração', 'N/A')}</p>
                            </div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                # INTEGRAÇÃO WHATSAPP
                if tel_limpo:
                    st.markdown("##### 📲 Envio de Mensagem WhatsApp")
                    opcao_wa = st.radio("Selecione o estilo do envio:", ["Link Direto Rápido", "Template Customizado"], horizontal=True)
                    
                    if opcao_wa == "Link Direto Rápido":
                        msg_final = f"Olá, equipe {posto.get('Razão Social', '')}! Aqui é da equipe de Capacitação AmPm. Gostaria de agendar o treinamento da loja."
                    else:
                        tmpl = st.selectbox("Escolha o Modelo de Mensagem:", [
                            "Agendamento de Treinamento",
                            "Cobrança / Verificação de Apostilas",
                            "Lembrete de Treinamento Agendado",
                            "Acompanhamento Pós-Treinamento"
                        ])
                        
                        if tmpl == "Agendamento de Treinamento":
                            msg_final = f"Olá! Aqui é da Capacitação AmPm. Gostaríamos de confirmar as datas disponíveis para o treinamento na loja {posto.get('Razão Social', '')} (PV {posto.get('PV Abadi', '')})."
                        elif tmpl == "Cobrança / Verificação de Apostilas":
                            msg_final = f"Olá, equipe {posto.get('Razão Social', '')}! Para darmos início ao treinamento, poderiam confirmar se o material de apoio e apostilas já chegaram na loja?"
                        elif tmpl == "Lembrete de Treinamento Agendado":
                            msg_final = f"Olá! Passando para lembrar que o treinamento AmPm da loja {posto.get('Razão Social', '')} está agendado para o dia {posto.get('Data_Agendada', 'em breve')}. Contamos com todos!"
                        else:
                            msg_final = f"Olá! Como foi o treinamento concluído na loja {posto.get('Razão Social', '')}? Estamos à disposição para dúvidas ou feedbacks."
                    
                    link_wa = f"https://wa.me/55{tel_limpo}?text={msg_final.replace(' ', '%20')}"
                    st.markdown(f"👉 **[Clique aqui para chamar no WhatsApp Direct]({link_wa})**")

                lista_instrutores = ["Pendente de Alocação"]
                if not df_instrutores.empty and 'NOME_COMPLETO' in df_instrutores.columns:
                    lista_instrutores += sorted(df_instrutores['NOME_COMPLETO'].dropna().unique().tolist())
                
                instrutor_atual = str(posto.get('Instrutor_Sugerido', 'Pendente de Alocação'))
                idx_instrutor = lista_instrutores.index(instrutor_atual) if instrutor_atual in lista_instrutores else 0

                val_data_agendada = posto.get('Data_Agendada')
                data_inicial = date.today()
                if isinstance(val_data_agendada, (date, datetime)):
                    data_inicial = val_data_agendada
                elif isinstance(val_data_agendada, str) and val_data_agendada:
                    try:
                        data_inicial = datetime.strptime(val_data_agendada, "%Y-%m-%d").date()
                    except ValueError:
                        try:
                            data_inicial = datetime.strptime(val_data_agendada, "%d/%m/%Y").date()
                        except ValueError:
                            data_inicial = date.today()

                # REGISTROS DA LIGAÇÃO
                with st.form("form_callcenter_editavel"):
                    st.markdown("#### ✍️ Registros Rápidos da Ligação")
                    
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        nome_c = st.text_input("👤 Nome do Responsável na Loja:", value=str(posto.get('Nome_Contato', '')))
                        tel_c = st.text_input("📞 Telefone de Contato:", value=str(posto.get('Telefone_Contato', '')))
                        qtd_func = st.number_input("👥 Qtd. de Funcionários para Treinar:", value=int(posto.get('Qtd_Funcionarios', 0)), min_value=0, step=1)
                        instrutor_escolhido = st.selectbox("👨‍🏫 Instrutor Designado:", lista_instrutores, index=idx_instrutor)
                        
                    with col_f2:
                        status_opcoes = ["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado"]
                        st_atual = posto.get('Status_Contato', 'A Contatar')
                        idx_st = status_opcoes.index(st_atual) if st_atual in status_opcoes else 0
                        novo_st = st.selectbox("🔄 Status do Atendimento:", status_opcoes, index=idx_st)
                        
                        mat_opcoes = ["Não Informado", "Sim", "Não"]
                        mat_atual = posto.get('Material_Em_Loja', 'Não Informado')
                        idx_mat = mat_opcoes.index(mat_atual) if mat_atual in mat_opcoes else 0
                        mat_loja = st.selectbox("📦 Possui Material/Apostilas na Loja?", mat_opcoes, index=idx_mat)
                        data_ag = st.date_input("📅 Data Agendada (Calendário):", value=data_inicial, format="DD/MM/YYYY")
                        
                    obs = st.text_area("💬 Observações e Alinhamentos:", value=str(posto.get('Observacoes', '')), height=80)
                    
                    if st.form_submit_button("💾 Salvar Registro do Atendimento"):
                        mask = st.session_state['df_base']['PV Abadi'] == pv_alvo
                        st.session_state['df_base'].loc[mask, 'Nome_Contato'] = nome_c
                        st.session_state['df_base'].loc[mask, 'Telefone_Contato'] = tel_c
                        st.session_state['df_base'].loc[mask, 'Qtd_Funcionarios'] = qtd_func
                        st.session_state['df_base'].loc[mask, 'Instrutor_Sugerido'] = instrutor_escolhido
                        st.session_state['df_base'].loc[mask, 'Material_Em_Loja'] = mat_loja
                        st.session_state['df_base'].loc[mask, 'Data_Agendada'] = data_ag.strftime("%d/%m/%Y")
                        st.session_state['df_base'].loc[mask, 'Status_Contato'] = novo_st
                        st.session_state['df_base'].loc[mask, 'Observacoes'] = obs
                        st.session_state['df_base'].loc[mask, 'Data_do_Contato'] = datetime.today().strftime('%d/%m/%Y %H:%M')
                        
                        salvar_alteracoes_disco()
                        st.success("✅ Atendimento registrado com sucesso!")
                        st.rerun()

                st.divider()
                st.markdown("#### ⏱️ Histórico de Interações")
                data_ct = posto.get('Data_do_Contato', 'Sem registro')
                data_agendada_str = posto.get('Data_Agendada', 'Não agendado')
                st.markdown(f"""
                    <div class="timeline-item">
                        <small style="color:#A0AAB8;"><b>Última Atualização:</b> {data_ct}</small><br>
                        <span><b>Status:</b> {posto.get('Status_Contato', '-')} | <b>Data Agendada:</b> {data_agendada_str} | <b>Instrutor:</b> {posto.get('Instrutor_Sugerido', '-')}</span><br>
                        <span style="color:#D1D5DB;"><i>"{posto.get('Observacoes', 'Sem observações registradas.')}"</i></span>
                    </div>
                """, unsafe_allow_html=True)

# ==========================================
# MÓDULO 6: EQUIPE DE INSTRUTORES
# ==========================================
elif modulo == "👔 Equipe de Instrutores":
    if not df_instrutores.empty:
        st.subheader("👔 Instrutores Credenciados na Rede")
        cols_inst = [c for c in ['NOME_COMPLETO', 'STATUS', 'TELEFONE', 'EMAIL', 'Cidade', 'UF'] if c in df_instrutores.columns]
        st.dataframe(df_instrutores[cols_inst], use_container_width=True, hide_index=True)
    else:
        st.info("ℹ️ Nenhuma lista separada de instrutores foi carregada. O sistema está utilizando os dados da base principal.")

# ==========================================
# MÓDULO 7: RELATÓRIOS & EXPORTAÇÃO
# ==========================================
elif modulo == "📂 Relatórios & Exportação":
    st.subheader("📂 Central de Exportação e Relatórios")
    st.caption("Faça o download dos dados operacionais atualizados em tempo real.")
    
    col_exp1, col_exp2 = st.columns(2)
    
    csv_buffer = df_base.to_csv(index=False).encode('utf-8')
    with col_exp1:
        st.download_button(
            label="📄 Baixar Base Filtrada em CSV",
            data=csv_buffer,
            file_name=f"Base_CRM_AmPm_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_base.to_excel(writer, sheet_name='Base_CRM_Atualizada', index=False)
    excel_data = output.getvalue()
    
    with col_exp2:
        st.download_button(
            label="📊 Baixar Base Filtrada em Excel",
            data=excel_data,
            file_name=f"Base_CRM_AmPm_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
