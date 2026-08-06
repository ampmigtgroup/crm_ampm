import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
import pydeck as pdk
import io

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="CRM Operacional AmPm",
    page_icon="⛽",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILIZAÇÃO CSS CUSTOMIZADA (DESIGN SYSTEM AMPM PREMIUM) ---
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

    /* Botão Customizado AmPm */
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

# --- CARREGAMENTO DE DADOS ---
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
            
            df_lojas['PV Abadi'] = pd.to_numeric(df_lojas['PV Abadi'], errors='coerce')
            df_fila['PV_Abadi'] = pd.to_numeric(df_fila['PV_Abadi'], errors='coerce')
            df_inaug['PV ABADI'] = pd.to_numeric(df_inaug['PV ABADI'], errors='coerce')
            df_rec['PV_ABADI'] = pd.to_numeric(df_rec['PV_ABADI'], errors='coerce')
            
            df_base = pd.merge(
                df_lojas,
                df_fila[['PV_Abadi', 'Tipo_Necessidade', 'Data_Ultimo_Treinamento', 
                         'Dias_desde_Ultimo_Treinamento', 'Instrutor_Sugerido', 
                         'Semana_Sugerida', 'Telefone_Contato', 'Status_Contato', 
                         'Data_do_Contato', 'Observacoes']],
                left_on='PV Abadi', right_on='PV_Abadi', how='left'
            )
            
            df_base = pd.merge(
                df_base,
                df_inaug[['PV ABADI', 'Previsão Inauguração', 'Pipeline', 'Consultor_Possivel_Instrutor']],
                left_on='PV Abadi', right_on='PV ABADI', how='left'
            )
            
            df_rec = pd.merge(
                df_rec,
                df_instrutores[['NOME_COMPLETO', 'Latitude', 'Longitude']],
                left_on='Instrutor_Sugerido', right_on='NOME_COMPLETO', how='left'
            ).rename(columns={'Latitude': 'Lat_Instrutor', 'Longitude': 'Lon_Instrutor'})
            
            df_rec = pd.merge(
                df_rec,
                df_lojas[['PV Abadi', 'Latitude', 'Longitude']],
                left_on='PV_ABADI', right_on='PV Abadi', how='left'
            ).rename(columns={'Latitude': 'Lat_Loja', 'Longitude': 'Lon_Loja'})

            df_base['Status_Contato'] = df_base['Status_Contato'].fillna('A Contatar')
            df_base['Tipo_Necessidade'] = df_base['Tipo_Necessidade'].fillna('Rede Ativa (Sem Pendência)')
            df_base['Instrutor_Sugerido'] = df_base['Instrutor_Sugerido'].fillna('Pendente de Alocação')
            df_base['Nome_Contato'] = ""
            df_base['Qtd_Funcionarios'] = 0
            df_base['Material_Em_Loja'] = "Não Informado"
            df_base['Data_Agendada'] = None
            
            return df_base, df_instrutores, df_rec
        except Exception as e:
            st.error(f"⚠️ Erro ao carregar planilhas: {e}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

def salvar_alteracoes_disco():
    caminho = "Base_Unificada_AmPm.xlsx"
    if os.path.exists(caminho) and 'df_base' in st.session_state:
        try:
            with pd.ExcelWriter(caminho, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                st.session_state['df_base'].to_excel(writer, sheet_name='Fila_CallCenter', index=False)
            st.toast("💾 Dados salvos no arquivo Excel local com sucesso!", icon="✅")
        except Exception as e:
            st.toast(f"⚠️ Erro ao salvar arquivo: {e}", icon="⚠️")

if 'df_base' not in st.session_state:
    b, i, r = carregar_bases_integradas()
    st.session_state['df_base'] = b
    st.session_state['df_instrutores'] = i
    st.session_state['df_rec'] = r

df_base_raw = st.session_state['df_base']
df_instrutores = st.session_state['df_instrutores']
df_rec = st.session_state['df_rec']

# --- SIDEBAR DE NAVEGAÇÃO, FILTROS GLOBAIS E UPLOAD ---
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
    
    # UPLOAD DE BANCO DE DADOS NA SIDEBAR
    st.markdown("📥 **Atualizar Banco de Dados**")
    uploaded_file = st.file_uploader("Envie a nova planilha (.xlsx):", type=["xlsx"], help="Carregue o arquivo Base_Unificada_AmPm.xlsx atualizado.")
    
    if uploaded_file is not None:
        try:
            with open("Base_Unificada_AmPm.xlsx", "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            st.cache_data.clear()
            b, i, r = carregar_bases_integradas()
            st.session_state['df_base'] = b
            st.session_state['df_instrutores'] = i
            st.session_state['df_rec'] = r
            
            st.success("✅ Banco de dados atualizado!")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Erro ao processar o arquivo: {e}")

    st.divider()
    
    # FILTROS GLOBAIS
    st.markdown("🎯 **Filtros Globais**")
    uf_opcoes = ["Todas"] + sorted([str(x) for x in df_base_raw['UF'].dropna().unique()]) if 'UF' in df_base_raw.columns else ["Todas"]
    filtro_uf = st.selectbox("Filtrar Estado (UF):", uf_opcoes)

    cf_opcoes = ["Todos"] + sorted([str(x) for x in df_base_raw['CF'].dropna().unique()]) if 'CF' in df_base_raw.columns else ["Todos"]
    filtro_cf = st.selectbox("Filtrar Consultor (CF):", cf_opcoes)

    st.divider()
    st.markdown("📶 **Status:** `Operacional 🟢`")
    st.markdown(f"🏪 **Rede Total:** `{len(df_base_raw)} Unidades`")

# APLICAÇÃO DOS FILTROS GLOBAIS
df_base = df_base_raw.copy()
if filtro_uf != "Todas":
    df_base = df_base[df_base['UF'] == filtro_uf]
if filtro_cf != "Todos":
    df_base = df_base[df_base['CF'] == filtro_cf]

# Header Global
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
            pendentes = len(df_base[df_base['Tipo_Necessidade'] != 'Rede Ativa (Sem Pendência)'])
            st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #FF9800;">
                    <div class="kpi-header"><span class="kpi-title">Fila Treinamento</span><span>🎓</span></div>
                    <div class="kpi-value">{pendentes}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with c3:
            a_contatar = len(df_base[df_base['Status_Contato'] == 'A Contatar'])
            st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #D32F2F;">
                    <div class="kpi-header"><span class="kpi-title">Pendentes Contato</span><span>📞</span></div>
                    <div class="kpi-value">{a_contatar}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with c4:
            inaug = len(df_base[df_base['Previsão Inauguração'].notna()])
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
            st.bar_chart(df_base['UF'].value_counts().head(10), color="#E27B00")
        with col_B:
            st.subheader("📊 Situação dos Contatos no Call Center")
            st.bar_chart(df_base['Status_Contato'].value_counts(), color="#FF9800")

# ==========================================
# MÓDULO 2: PIPELINE AMPM
# ==========================================
elif modulo == "📋 Pipeline AmPm":
    st.subheader("📋 Pipeline AmPm — Fluxo Operacional de Treinamentos")
    st.caption("Gerencie o fluxo de atendimento navegando entre os estágios de contato.")
    
    colunas_pipeline = ["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado"]
    cols_k = st.columns(len(colunas_pipeline))
    
    for idx, status in enumerate(colunas_pipeline):
        df_status = df_base[df_base['Status_Contato'] == status]
        
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
                with st.expander(f"📍 PV {item['PV Abadi']} | {str(item['Razão Social'])[:14]}..."):
                    st.write(f"**Cidade:** {item['Municipio']}/{item['UF']}")
                    st.write(f"**Necessidade:** {item['Tipo_Necessidade']}")
                    st.write(f"**Treinandos:** {item.get('Qtd_Funcionarios', 0)} pessoas")
                    st.write(f"**Instrutor:** {item.get('Instrutor_Sugerido', 'Pendente')}")
                    
                    mudar_status = st.selectbox(
                        "Alterar Status:",
                        colunas_pipeline,
                        index=colunas_pipeline.index(status),
                        key=f"pipe_sel_{item['PV Abadi']}"
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
            f_necessidade = f2.selectbox("🎯 Tipo de Necessidade:", ["Todas"] + sorted([str(x) for x in df_base['Tipo_Necessidade'].dropna().unique()]))
            
        df_view = df_base.copy()
        if termo:
            df_view = df_view[
                df_view['Razão Social'].astype(str).str.contains(termo, case=False, na=False) |
                df_view['PV Abadi'].astype(str).str.contains(termo, na=False) |
                df_view['Municipio'].astype(str).str.contains(termo, case=False, na=False)
            ]
        if f_necessidade != "Todas":
            df_view = df_view[df_view['Tipo_Necessidade'] == f_necessidade]
            
        st.caption("👇 *Clique em uma linha para abrir a Ficha Detalhada PROCV:*")
        
        evento = st.dataframe(
            df_view[['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Status Loja', 'Tipo_Necessidade', 'Status_Contato']],
            use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun"
        )
        
        linhas = evento.selection.get("rows", [])
        if linhas:
            p = df_view.iloc[linhas[0]].to_dict()
            st.divider()
            st.markdown(f"### 📋 Ficha de Detalhes PROCV — **PV {p['PV Abadi']} | {p['Razão Social']}**")
            
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
        if filtro_uf != "Todas":
            df_rec_filtrado = df_rec_filtrado[df_rec_filtrado['UF_Loja'] == filtro_uf]

        postos_unicos = df_rec_filtrado[['PV_ABADI', 'Razao_Social', 'Municipio_Loja', 'UF_Loja']].drop_duplicates()
        if not postos_unicos.empty:
            postos_unicos['label'] = postos_unicos['PV_ABADI'].astype(str) + " - " + postos_unicos['Razao_Social'] + " (" + postos_unicos['Municipio_Loja'] + "/" + postos_unicos['UF_Loja'] + ")"
            
            posto_sel = st.selectbox("⛽ Selecione o Posto Alvo:", postos_unicos['label'].tolist())
            pv_sel = int(posto_sel.split(" - ")[0])
            
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
                    dist = row['Distancia_km_linha_reta']
                    dias = row['Dias_Treinamento_Necessarios']
                    
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

                    with cols[idx]:
                        st.markdown(f"""
                            <div class="top-instructor-card">
                                <h4 style="margin:0 0 8px 0; color:#E27B00;">#{row['Ranking_Proximidade']}º {row['Instrutor_Sugerido']}</h4>
                                <p style="margin:2px 0;">🏙️ <b>Origem:</b> {row['Cidade_Instrutor']}/{row['UF_Instrutor']}</p>
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
        df_fila_view = df_base[df_base['Tipo_Necessidade'] != 'Rede Ativa (Sem Pendência)'].copy()
        
        c_left, c_right = st.columns([1.2, 1.8])
        
        with c_left:
            st.subheader("📋 Fila de Atendimento")
            evento_call = st.dataframe(
                df_fila_view[['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Status_Contato']],
                use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun"
            )
            selecionado = evento_call.selection.get("rows", [])
            
        with c_right:
            if selecionado:
                posto = df_fila_view.iloc[selecionado[0]]
                pv_alvo = posto['PV Abadi']
                tel_limpo = ''.join(filter(str.isdigit, str(posto.get('Telefone_Contato', ''))))
                
                st.markdown(f"### 📝 Ficha de Atendimento — **PV {posto['PV Abadi']}**")
                
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
                
                # --- INTEGRAÇÃO WHATSAPP FLEXÍVEL (LINK DIRETO + TEMPLATES) ---
                if tel_limpo:
                    st.markdown("##### 📲 Envio de Mensagem WhatsApp")
                    opcao_wa = st.radio("Selecione o estilo do envio:", ["Link Direto Rápido", "Template Customizado"], horizontal=True)
                    
                    if opcao_wa == "Link Direto Rápido":
                        msg_final = f"Olá, equipe {posto['Razão Social']}! Aqui é da equipe de Capacitação AmPm. Gostaria de agendar o treinamento da loja."
                    else:
                        tmpl = st.selectbox("Escolha o Modelo de Mensagem:", [
                            "Agendamento de Treinamento",
                            "Cobrança / Verificação de Apostilas",
                            "Lembrete de Treinamento Agendado",
                            "Acompanhamento Pós-Treinamento"
                        ])
                        
                        if tmpl == "Agendamento de Treinamento":
                            msg_final = f"Olá! Aqui é da Capacitação AmPm. Gostaríamos de confirmar as datas disponíveis para o treinamento na loja {posto['Razão Social']} (PV {posto['PV Abadi']})."
                        elif tmpl == "Cobrança / Verificação de Apostilas":
                            msg_final = f"Olá, equipe {posto['Razão Social']}! Para darmos início ao treinamento, poderiam confirmar se o material de apoio e apostilas já chegaram na loja?"
                        elif tmpl == "Lembrete de Treinamento Agendado":
                            msg_final = f"Olá! Passando para lembrar que o treinamento AmPm da loja {posto['Razão Social']} está agendado para o dia {posto.get('Data_Agendada', 'em breve')}. Contamos com todos!"
                        else:
                            msg_final = f"Olá! Como foi o treinamento concluído na loja {posto['Razão Social']}? Estamos à disposição para dúvidas ou feedbacks."
                    
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

                # --- REGISTROS RÁPIDOS DA LIGAÇÃO ---
                with st.form("form_callcenter_editavel"):
                    st.markdown("#### ✍️ Registros Rápidos da Ligação")
                    
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        nome_c = st.text_input("👤 Nome do Responsável na Loja:", value=str(posto.get('Nome_Contato', '')))
                        tel_c = st.text_input("📞 Telefone de Contato:", value=str(posto.get('Telefone_Contato', '')))
                        qtd_func = st.number_input("👥 Qtd. de Funcionários para Treinar:", value=int(posto.get('Qtd_Funcionarios', 0)), min_value=0, step=1)
                        instrutor_escolhido = st.selectbox("👨‍🏫 Instrutor Designado:", lista_instrutores, index=idx_instrutor)
                        
                    with col_f2:
                        novo_st = st.selectbox(
                            "🔄 Status do Atendimento:", 
                            ["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado"],
                            index=["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado"].index(posto.get('Status_Contato', 'A Contatar'))
                        )
                        mat_loja = st.selectbox("📦 Possui Material/Apostilas na Loja?", ["Não Informado", "Sim", "Não"], index=["Não Informado", "Sim", "Não"].index(posto.get('Material_Em_Loja', 'Não Informado')))
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
        st.dataframe(df_instrutores[['NOME_COMPLETO', 'STATUS', 'TELEFONE', 'EMAIL', 'Cidade', 'UF']], use_container_width=True, hide_index=True)

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

