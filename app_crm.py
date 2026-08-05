import streamlit as st
import pandas as pd
import os
from datetime import datetime
import pydeck as pdk

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="CRM Operacional AmPm",
    page_icon="⛽",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILIZAÇÃO CSS CUSTOMIZADA & ÍCONES (DESIGN SYSTEM AMPM PREMIUM) ---
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
        display: flex;
        align-items: center;
        gap: 12px;
    }

    .main-header p {
        margin: 0;
        font-size: 1.05rem;
        opacity: 0.95;
    }

    /* Cards de Métricas (KPIs) com Destaque Visual */
    .kpi-card {
        background-color: #1E222A;
        border-radius: 12px;
        padding: 20px;
        border: 1px solid #2D333F;
        border-left: 6px solid #E27B00;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(0,0,0,0.3);
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
    .kpi-icon {
        font-size: 1.5rem;
    }
    .kpi-value {
        font-size: 2rem;
        font-weight: 700;
        color: #FFFFFF;
        margin-top: 8px;
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
        display: flex;
        align-items: center;
        gap: 8px;
        font-size: 1.1rem;
    }
    .procv-card p {
        margin: 8px 0;
        font-size: 0.95rem;
        color: #D1D5DB;
    }
    
    .top-instructor-card {
        background-color: #1A1D24;
        padding: 18px;
        border-radius: 12px;
        border: 1px solid #2D333F;
        border-left: 5px solid #4CAF50;
        margin-bottom: 14px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    }

    /* Badges Customizados */
    .badge-aereo {
        background: rgba(2, 136, 209, 0.15);
        color: #29B6F6;
        border: 1px solid #0288D1;
        padding: 5px 12px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.82rem;
        display: inline-flex;
        align-items: center;
        gap: 6px;
    }
    .badge-terrestre {
        background: rgba(56, 142, 60, 0.15);
        color: #66BB6A;
        border: 1px solid #388E3C;
        padding: 5px 12px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.82rem;
        display: inline-flex;
        align-items: center;
        gap: 6px;
    }

    /* Botão Customizado AmPm */
    .stButton>button {
        background: linear-gradient(90deg, #E27B00 0%, #FF9800 100%);
        color: #FFFFFF !important;
        font-weight: 600;
        border: none;
        border-radius: 8px;
        padding: 12px 28px;
        box-shadow: 0 4px 12px rgba(226, 123, 0, 0.3);
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        box-shadow: 0 6px 18px rgba(226, 123, 0, 0.5);
        transform: translateY(-2px);
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

# --- SIDEBAR DE NAVEGAÇÃO REFORÇADA COM ÍCONES ---
with st.sidebar:
    st.markdown("## ⛽ **CRM AmPm**")
    st.caption("🌐 *Plataforma Integrada de Operações*")
    st.divider()
    
    modulo = st.radio(
        "📌 **Módulos do Sistema:**",
        [
            "📊 Dashboard Executivo", 
            "🔍 PROCV & Consulta de Lojas", 
            "📍 Otimizador de Rotas (Top 3)", 
            "📞 Fila Call Center", 
            "👔 Equipe de Instrutores"
        ]
    )
    
    st.divider()
    st.markdown("📶 **Status do Servidor:** `Operacional 🟢`")
    st.markdown(f"🏪 **Rede de Lojas:** `{len(df_base)} Unidades`")
    st.markdown(f"👨‍🏫 **Instrutores Ativos:** `{len(df_instrutores)} Credenciados`")

# Header Global com Ícone de Destaque
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
                    <div class="kpi-header">
                        <span class="kpi-title">Rede Total</span>
                        <span class="kpi-icon">🏪</span>
                    </div>
                    <div class="kpi-value">{len(df_base)}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with c2:
            pendentes = len(df_base[df_base['Tipo_Necessidade'] != 'Rede Ativa (Sem Pendência)'])
            st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #FF9800;">
                    <div class="kpi-header">
                        <span class="kpi-title">Fila de Treinamento</span>
                        <span class="kpi-icon">🎓</span>
                    </div>
                    <div class="kpi-value">{pendentes}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with c3:
            a_contatar = len(df_base[df_base['Status_Contato'] == 'A Contatar'])
            st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #D32F2F;">
                    <div class="kpi-header">
                        <span class="kpi-title">Pendentes de Contato</span>
                        <span class="kpi-icon">📞</span>
                    </div>
                    <div class="kpi-value">{a_contatar}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with c4:
            inaug = len(df_base[df_base['Previsão Inauguração'].notna()])
            st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #0288D1;">
                    <div class="kpi-header">
                        <span class="kpi-title">Inaugurações Previstas</span>
                        <span class="kpi-icon">🚀</span>
                    </div>
                    <div class="kpi-value">{inaug}</div>
                </div>
            """, unsafe_allow_html=True)

        st.write("")
        st.divider()
        col_A, col_B = st.columns(2)
        with col_A:
            st.subheader("🗺️ Concentração Geográfica por Estado (UF)")
            st.bar_chart(df_base['UF'].value_counts().head(10), color="#E27B00")
        with col_B:
            st.subheader("📊 Situação dos Contatos no Call Center")
            st.bar_chart(df_base['Status_Contato'].value_counts(), color="#FF9800")

# ==========================================
# MÓDULO 2: PROCV & CONSULTA DE LOJAS
# ==========================================
elif modulo == "🔍 PROCV & Consulta de Lojas":
    if not df_base.empty:
        col_busca, col_uf = st.columns([3, 1])
        with col_busca:
            termo = st.text_input("🔍 Digite o PV Abadi, Nome da Loja ou Município:", "")
        with col_uf:
            f_uf = st.selectbox("📌 UF:", ["Todas"] + sorted([str(x) for x in df_base['UF'].dropna().unique()]))
            
        df_view = df_base.copy()
        if termo:
            df_view = df_view[
                df_view['Razão Social'].astype(str).str.contains(termo, case=False, na=False) |
                df_view['PV Abadi'].astype(str).str.contains(termo, na=False) |
                df_view['Municipio'].astype(str).str.contains(termo, case=False, na=False)
            ]
        if f_uf != "Todas":
            df_view = df_view[df_view['UF'] == f_uf]
            
        st.caption("👇 *Clique em uma linha da tabela abaixo para expandir o painel completo de detalhes:*")
        
        evento = st.dataframe(
            df_view[['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Status Loja', 'Tipo_Necessidade', 'Status_Contato']],
            use_container_width=True,
            hide_index=True,
            selection_mode="single-row",
            on_select="rerun"
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
                        <p>📅 <b>Previsão Inauguração:</b> {p.get('Previsão Inauguração', 'N/A')}</p>
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
# MÓDULO 3: OTIMIZADOR DE ROTAS (TOP 3)
# ==========================================
elif modulo == "📍 Otimizador de Rotas (Top 3)":
    if not df_rec.empty:
        postos_unicos = df_rec[['PV_ABADI', 'Razao_Social', 'Municipio_Loja', 'UF_Loja']].drop_duplicates()
        postos_unicos['label'] = postos_unicos['PV_ABADI'].astype(str) + " - " + postos_unicos['Razao_Social'] + " (" + postos_unicos['Municipio_Loja'] + "/" + postos_unicos['UF_Loja'] + ")"
        
        posto_sel = st.selectbox("⛽ Selecione o Posto/Cliente Destino:", postos_unicos['label'].tolist())
        
        pv_sel = int(posto_sel.split(" - ")[0])
        top_3 = df_rec[df_rec['PV_ABADI'] == pv_sel].sort_values(by='Ranking_Proximidade').head(3)
        
        if not top_3.empty:
            info_posto = top_3.iloc[0]
            st.caption(f"🎯 **Destino:** {info_posto['Municipio_Loja']}/{info_posto['UF_Loja']} | ⏱️ **Duração do Treinamento:** {info_posto['Dias_Treinamento_Necessarios']} dia(s)")
            
            st.write("")
            col1, col2, col3 = st.columns(3)
            cols = [col1, col2, col3]
            
            for idx, (_, row) in enumerate(top_3.iterrows()):
                dist = row['Distancia_km_linha_reta']
                modal_tag = "<span class='badge-aereo'>✈️ Passagem Aérea</span>" if dist > 300 else "<span class='badge-terrestre'>🚗 Deslocamento Terrestre</span>"
                
                with cols[idx]:
                    st.markdown(f"""
                        <div class="top-instructor-card">
                            <h4 style="margin:0 0 8px 0; color:#E27B00;">🥇 #{row['Ranking_Proximidade']}º {row['Instrutor_Sugerido']}</h4>
                            <p style="margin:4px 0;">🏙️ <b>Origem:</b> {row['Cidade_Instrutor']}/{row['UF_Instrutor']}</p>
                            <p style="margin:4px 0;">📏 <b>Distância:</b> <code>{dist} km</code></p>
                            <div style="margin-top:12px;">{modal_tag}</div>
                        </div>
                    """, unsafe_allow_html=True)
            
            st.divider()
            
            # --- MAPA DE ROTA COM PYDECK ---
            col_sel_mapa, col_mapa_view = st.columns([1, 2])
            
            with col_sel_mapa:
                st.markdown("### 🗺️ Seleção da Rota")
                inst_mapa = st.radio("Selecione o Instrutor para visualizar o trajeto:", top_3['Instrutor_Sugerido'].tolist())
                dados_rota = top_3[top_3['Instrutor_Sugerido'] == inst_mapa].iloc[0]
                
                dist_sel = dados_rota['Distancia_km_linha_reta']
                modal_text = "✈️ Aéreo (Avião + Conexão)" if dist_sel > 300 else "🚗 Terrestre (Carro / Uber / Ônibus)"
                
                st.info(f"""
                📌 **Resumo Logístico:**
                - 🛫 **Origem:** {dados_rota['Cidade_Instrutor']}/{dados_rota['UF_Instrutor']}
                - 🛬 **Destino:** {dados_rota['Municipio_Loja']}/{dados_rota['UF_Loja']}
                - 📏 **Distância:** `{dist_sel} km`
                - 💡 **Modal Indicado:** {modal_text}
                """)

            with col_mapa_view:
                lat_loja, lon_loja = pd.to_numeric(dados_rota.get('Lat_Loja')), pd.to_numeric(dados_rota.get('Lon_Loja'))
                lat_inst, lon_inst = pd.to_numeric(dados_rota.get('Lat_Instrutor')), pd.to_numeric(dados_rota.get('Lon_Instrutor'))
                
                if pd.notna(lat_loja) and pd.notna(lat_inst):
                    linha_data = [{'start_lon': float(lon_inst), 'start_lat': float(lat_inst), 'end_lon': float(lon_loja), 'end_lat': float(lat_loja)}]
                    pontos_data = [
                        {"lon": float(lon_inst), "lat": float(lat_inst), "nome": f"Origem: {dados_rota['Cidade_Instrutor']} (Instrutor)", "color": [76, 175, 80]},
                        {"lon": float(lon_loja), "lat": float(lat_loja), "nome": f"Destino: {dados_rota['Razao_Social']} (Posto)", "color": [211, 47, 47]}
                    ]
                    
                    layer_line = pdk.Layer("LineLayer", linha_data, get_source_position=["start_lon", "start_lat"], get_target_position=["end_lon", "end_lat"], get_color=[226, 123, 0, 255], get_width=6)
                    layer_points = pdk.Layer("ScatterplotLayer", pontos_data, get_position=["lon", "lat"], get_color="color", get_radius=12000, pickable=True)
                    
                    view_state = pdk.ViewState(latitude=(float(lat_inst) + float(lat_loja)) / 2, longitude=(float(lon_inst) + float(lon_loja)) / 2, zoom=5)
                    
                    st.pydeck_chart(pdk.Deck(layers=[layer_line, layer_points], initial_view_state=view_state, tooltip={"text": "{nome}"}), use_container_width=True)

# ==========================================
# MÓDULO 4: FILA CALL CENTER
# ==========================================
elif modulo == "📞 Fila Call Center":
    if not df_base.empty:
        df_fila_view = df_base[df_base['Tipo_Necessidade'] != 'Rede Ativa (Sem Pendência)'].copy()
        
        c_left, c_right = st.columns([1.6, 1.4])
        
        with c_left:
            st.subheader("📋 Fila de Lojas Pendentes / Agendadas")
            evento_call = st.dataframe(
                df_fila_view[['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Status_Contato']],
                use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun"
            )
            selecionado = evento_call.selection.get("rows", [])
            
        with c_right:
            if selecionado:
                posto = df_fila_view.iloc[selecionado[0]]
                pv_alvo = posto['PV Abadi']
                
                with st.form("form_callcenter"):
                    st.markdown(f"### 📝 Ficha de Contato — PV {posto['PV Abadi']}")
                    st.caption(f"🏪 **{posto['Razão Social']}** ({posto['Municipio']}/{posto['UF']})")
                    st.divider()
                    
                    nome_c = st.text_input("👤 Nome do Responsável / Gerente:", value=str(posto.get('Nome_Contato', '')))
                    tel_c = st.text_input("📞 Telefone para Contato:", value=str(posto.get('Telefone_Contato', '')))
                    qtd_f = st.number_input("👥 Nº de Funcionários na Loja:", min_value=0, value=int(posto.get('Qtd_Funcionarios', 0)))
                    
                    novo_st = st.selectbox("🔄 Status do Atendimento:", ["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado"])
                    obs = st.text_area("💬 Observações do Atendimento:", value=str(posto.get('Observacoes', '')))
                    
                    if st.form_submit_button("💾 Salvar Registro de Atendimento"):
                        mask = st.session_state['df_base']['PV Abadi'] == pv_alvo
                        st.session_state['df_base'].loc[mask, 'Nome_Contato'] = nome_c
                        st.session_state['df_base'].loc[mask, 'Telefone_Contato'] = tel_c
                        st.session_state['df_base'].loc[mask, 'Qtd_Funcionarios'] = qtd_f
                        st.session_state['df_base'].loc[mask, 'Status_Contato'] = novo_st
                        st.session_state['df_base'].loc[mask, 'Observacoes'] = obs
                        st.session_state['df_base'].loc[mask, 'Data_do_Contato'] = datetime.today().strftime('%d/%m/%Y %H:%M')
                        
                        st.success(f"✅ Atendimento do PV {pv_alvo} atualizado com sucesso!")
                        st.rerun()

# ==========================================
# MÓDULO 5: EQUIPE DE INSTRUTORES
# ==========================================
elif modulo == "👔 Equipe de Instrutores":
    if not df_instrutores.empty:
        st.subheader("👔 Instrutores Credenciados na Rede")
        st.dataframe(df_instrutores[['NOME_COMPLETO', 'STATUS', 'TELEFONE', 'EMAIL', 'Cidade', 'UF']], use_container_width=True, hide_index=True)
