import streamlit as st
import pandas as pd
import os
from datetime import datetime
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

    /* KPI Cards com Indicadores */
    .kpi-card {
        background-color: #1E222A;
        border-radius: 12px;
        padding: 20px;
        border: 1px solid #2D333F;
        border-left: 6px solid #E27B00;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        transition: transform 0.2s ease;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
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

    /* Estilização do Kanban */
    .kanban-column {
        background-color: #14171D;
        border-radius: 12px;
        padding: 14px;
        border: 1px solid #2D333F;
        min-height: 500px;
    }
    .kanban-title {
        font-size: 0.95rem;
        font-weight: 700;
        margin-bottom: 12px;
        padding-bottom: 8px;
        border-bottom: 2px solid #2D333F;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }

    /* Cards de Informação PROCV e Instrutores */
    .procv-card {
        background-color: #1A1D24;
        padding: 22px;
        border-radius: 12px;
        border: 1px solid #2D333F;
        border-top: 4px solid #E27B00;
        height: 100%;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    }
    .procv-card h4 {
        margin-top: 0;
        margin-bottom: 14px;
        color: #FF9800;
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

    /* Badges Logísticos */
    .badge-aereo {
        background: rgba(2, 136, 209, 0.15);
        color: #29B6F6;
        border: 1px solid #0288D1;
        padding: 4px 10px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.8rem;
    }
    .badge-terrestre {
        background: rgba(56, 142, 60, 0.15);
        color: #66BB6A;
        border: 1px solid #388E3C;
        padding: 4px 10px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.8rem;
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
            df_base['Material_Em_Loja'] = "N/A"
            
            return df_base, df_instrutores, df_rec
        except Exception as e:
            st.error(f"⚠️ Erro ao carregar planilhas: {e}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

if 'df_base' not in st.session_state:
    b, i, r = carregar_bases_integradas()
    st.session_state['df_base'] = b
    st.session_state['df_instrutores'] = i
    st.session_state['df_rec'] = r

df_base = st.session_state['df_base']
df_instrutores = st.session_state['df_instrutores']
df_rec = st.session_state['df_rec']

# --- SIDEBAR DE NAVEGAÇÃO ---
with st.sidebar:
    st.markdown("## ⛽ **CRM AmPm**")
    st.caption("🌐 *Plataforma Integrada de Operações*")
    st.divider()
    
    modulo = st.radio(
        "📌 **Módulos do Sistema:**",
        [
            "📊 Dashboard Executivo", 
            "📋 Pipeline Kanban", 
            "🔍 PROCV & Filtros Avançados", 
            "📍 Calculadora & Otimizador de Custos", 
            "📞 Call Center & Timeline WhatsApp", 
            "👔 Equipe de Instrutores",
            "📂 Relatórios & Exportação"
        ]
    )
    
    st.divider()
    st.markdown("📶 **Status:** `Operacional 🟢`")
    st.markdown(f"🏪 **Rede:** `{len(df_base)} Unidades`")

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
                    <div class="kpi-header"><span class="kpi-title">Rede Total</span><span>🏪</span></div>
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
# MÓDULO 2: PIPELINE KANBAN INTERATIVO
# ==========================================
elif modulo == "📋 Pipeline Kanban":
    st.subheader("📋 Pipeline Operacional de Treinamentos")
    st.caption("Gerencie o fluxo de atendimento da rede navegando entre os estágios de contato.")
    
    colunas_kanban = ["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado"]
    cols_k = st.columns(len(colunas_kanban))
    
    for idx, status in enumerate(colunas_kanban):
        df_status = df_base[df_base['Status_Contato'] == status]
        
        with cols_k[idx]:
            st.markdown(f"""
                <div class="kanban-column">
                    <div class="kanban-title">
                        <span>{status}</span>
                        <span style="background:#2D333F; padding:2px 8px; border-radius:10px; font-size:0.8rem;">{len(df_status)}</span>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            for _, item in df_status.head(6).iterrows():
                with st.expander(f"📍 PV {item['PV Abadi']} | {str(item['Razão Social'])[:14]}..."):
                    st.write(f"**Cidade:** {item['Municipio']}/{item['UF']}")
                    st.write(f"**Necessidade:** {item['Tipo_Necessidade']}")
                    
                    mudar_status = st.selectbox(
                        "Alterar Status:",
                        colunas_kanban,
                        index=colunas_kanban.index(status),
                        key=f"kan_sel_{item['PV Abadi']}"
                    )
                    
                    if mudar_status != status:
                        mask = st.session_state['df_base']['PV Abadi'] == item['PV Abadi']
                        st.session_state['df_base'].loc[mask, 'Status_Contato'] = mudar_status
                        st.success("Atualizado!")
                        st.rerun()

# ==========================================
# MÓDULO 3: PROCV & FILTROS AVANÇADOS
# ==========================================
elif modulo == "🔍 PROCV & Filtros Avançados":
    if not df_base.empty:
        with st.expander("🔎 **Filtros Avançados de Pesquisa**", expanded=True):
            f1, f2, f3 = st.columns(3)
            termo = f1.text_input("🔍 PV, Nome ou Município:", "")
            f_uf = f2.selectbox("📌 UF:", ["Todas"] + sorted([str(x) for x in df_base['UF'].dropna().unique()]))
            f_necessidade = f3.selectbox("🎯 Tipo de Necessidade:", ["Todas"] + sorted([str(x) for x in df_base['Tipo_Necessidade'].dropna().unique()]))
            
        df_view = df_base.copy()
        if termo:
            df_view = df_view[
                df_view['Razão Social'].astype(str).str.contains(termo, case=False, na=False) |
                df_view['PV Abadi'].astype(str).str.contains(termo, na=False) |
                df_view['Municipio'].astype(str).str.contains(termo, case=False, na=False)
            ]
        if f_uf != "Todas":
            df_view = df_view[df_view['UF'] == f_uf]
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
    st.caption("Cálculo detalhado de custos de viagens com indicador de economia por rota.")
    
    if not df_rec.empty:
        postos_unicos = df_rec[['PV_ABADI', 'Razao_Social', 'Municipio_Loja', 'UF_Loja']].drop_duplicates()
        postos_unicos['label'] = postos_unicos['PV_ABADI'].astype(str) + " - " + postos_unicos['Razao_Social'] + " (" + postos_unicos['Municipio_Loja'] + "/" + postos_unicos['UF_Loja'] + ")"
        
        posto_sel = st.selectbox("⛽ Selecione o Posto Alvo:", postos_unicos['label'].tolist())
        pv_sel = int(posto_sel.split(" - ")[0])
        
        top_3 = df_rec[df_rec['PV_ABADI'] == pv_sel].sort_values(by='Ranking_Proximidade').head(3)
        
        if not top_3.empty:
            st.divider()
            
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

            # Mapa Visual Pydeck
            st.divider()
            dados_rota = top_3.iloc[0]
            lat_loja, lon_loja = pd.to_numeric(dados_rota.get('Lat_Loja')), pd.to_numeric(dados_rota.get('Lon_Loja'))
            lat_inst, lon_inst = pd.to_numeric(dados_rota.get('Lat_Instrutor')), pd.to_numeric(dados_rota.get('Lon_Instrutor'))
            
            if pd.notna(lat_loja) and pd.notna(lat_inst):
                linha_data = [{'start_lon': float(lon_inst), 'start_lat': float(lat_inst), 'end_lon': float(lon_loja), 'end_lat': float(lat_loja)}]
                pontos_data = [
                    {"lon": float(lon_inst), "lat": float(lat_inst), "nome": f"Origem: {dados_rota['Cidade_Instrutor']}", "color": [76, 175, 80]},
                    {"lon": float(lon_loja), "lat": float(lat_loja), "nome": f"Destino: {dados_rota['Razao_Social']}", "color": [211, 47, 47]}
                ]
                
                layer_line = pdk.Layer("LineLayer", linha_data, get_source_position=["start_lon", "start_lat"], get_target_position=["end_lon", "end_lat"], get_color=[226, 123, 0, 255], get_width=6)
                layer_points = pdk.Layer("ScatterplotLayer", pontos_data, get_position=["lon", "lat"], get_color="color", get_radius=12000, pickable=True)
                view_state = pdk.ViewState(latitude=(float(lat_inst) + float(lat_loja)) / 2, longitude=(float(lon_inst) + float(lon_loja)) / 2, zoom=5)
                
                st.pydeck_chart(pdk.Deck(layers=[layer_line, layer_points], initial_view_state=view_state, tooltip={"text": "{nome}"}), use_container_width=True)

# ==========================================
# MÓDULO 5: CALL CENTER & TIMELINE WHATSAPP
# ==========================================
elif modulo == "📞 Call Center & Timeline WhatsApp":
    if not df_base.empty:
        df_fila_view = df_base[df_base['Tipo_Necessidade'] != 'Rede Ativa (Sem Pendência)'].copy()
        
        c_left, c_right = st.columns([1.5, 1.5])
        
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
                
                st.markdown(f"### 📝 Ficha de Contato — PV {posto['PV Abadi']}")
                
                if tel_limpo:
                    msg = f"Olá, equipe {posto['Razão Social']}! Aqui é da equipe de Treinamento AmPm."
                    link_wa = f"https://wa.me/55{tel_limpo}?text={msg.replace(' ', '%20')}"
                    st.markdown(f"📲 **[Abrir conversa no WhatsApp Web/App]( {link_wa} )**")

                with st.form("form_callcenter"):
                    nome_c = st.text_input("👤 Nome Responsável:", value=str(posto.get('Nome_Contato', '')))
                    tel_c = st.text_input("📞 Telefone:", value=str(posto.get('Telefone_Contato', '')))
                    novo_st = st.selectbox("🔄 Status Atendimento:", ["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado"])
                    obs = st.text_area("💬 Observações:", value=str(posto.get('Observacoes', '')))
                    
                    if st.form_submit_button("💾 Salvar Registro de Atendimento"):
                        mask = st.session_state['df_base']['PV Abadi'] == pv_alvo
                        st.session_state['df_base'].loc[mask, 'Nome_Contato'] = nome_c
                        st.session_state['df_base'].loc[mask, 'Telefone_Contato'] = tel_c
                        st.session_state['df_base'].loc[mask, 'Status_Contato'] = novo_st
                        st.session_state['df_base'].loc[mask, 'Observacoes'] = obs
                        st.session_state['df_base'].loc[mask, 'Data_do_Contato'] = datetime.today().strftime('%d/%m/%Y %H:%M')
                        
                        st.success("✅ Atendimento registrado com sucesso!")
                        st.rerun()

                # Histórico Cronológico / Timeline
                st.divider()
                st.markdown("#### ⏱️ Histórico Recente de Interações")
                data_ct = posto.get('Data_do_Contato', 'Sem registro')
                st.markdown(f"""
                    <div class="timeline-item">
                        <small style="color:#A0AAB8;"><b>Data do Contato:</b> {data_ct}</small><br>
                        <span><b>Status:</b> {posto.get('Status_Contato', '-')}</span><br>
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
            label="📄 Baixar Base Completa em CSV",
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
            label="📊 Baixar Base Completa em Excel",
            data=excel_data,
            file_name=f"Base_CRM_AmPm_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
