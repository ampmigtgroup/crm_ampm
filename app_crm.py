import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from math import radians, cos, sin, asin, sqrt
from fpdf import FPDF
import io

# ==========================================
# CONFIGURAÇÃO DA PÁGINA (WIDE & DARK THEME)
# ==========================================
st.set_page_config(
    page_title="CRM Operacional AmPm - Gestão & Treinamentos",
    page_icon="⛽",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização CSS Customizada para visual Executivo/Dark
st.markdown("""
<style>
    .main { background-color: #0E1117; }
    .stMetric {
        background-color: #1E222D;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #2D333F;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
    }
    .top-instructor-card {
        background-color: #1E222D;
        border-left: 5px solid #FF9800;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 12px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    }
    .badge-info {
        background-color: #0083B0;
        color: white;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 0.8rem;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# FUNÇÕES AUXILIARES & CÁLCULOS LOGÍSTICOS
# ==========================================
def haversine(lon1, lat1, lon2, lat2):
    """Calcula a distância em quilômetros entre dois pontos geográficos."""
    try:
        lon1, lat1, lon2, lat2 = map(radians, [float(lon1), float(lat1), float(lon2), float(lat2)])
        dlon = lon2 - lon1 
        dlat = lat2 - lat1 
        a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
        c = 2 * asin(sqrt(a)) 
        r = 6371 # Raio da Terra em KM
        return c * r
    except:
        return np.nan

@st.cache_data
def carregar_dados_mock():
    """Gera dados realistas de teste da rede AmPm e Instrutores caso nenhuma planilha seja enviada."""
    lojas = []
    ufs = ['SP', 'RJ', 'PR', 'SC', 'RS', 'MG', 'BA', 'PE', 'GO', 'DF']
    cidades = ['São Paulo', 'Rio de Janeiro', 'Curitiba', 'Florianópolis', 'Porto Alegre', 'Belo Horizonte', 'Salvador', 'Recife', 'Goiânia', 'Brasília']
    lats = [-23.5505, -22.9068, -25.4284, -27.5954, -30.0346, -19.9167, -12.9714, -8.0476, -16.6869, -15.7975]
    lons = [-46.6333, -43.1729, -49.2733, -48.5480, -51.2177, -43.9345, -38.5014, -34.8770, -49.2648, -47.8919]
    
    # Criar 60 lojas fictícias
    for i in range(1, 61):
        idx_loc = i % len(ufs)
        lojas.append({
            'PV Abadi': 2600 + i,
            'Razão Social': f'Auto Posto Petrosolo Ltd {i}',
            'Nome Fantasia': f'Posto AmPm {i}',
            'UF': ufs[idx_loc],
            'Municipio': cidades[idx_loc],
            'Regional': f'Regional {ufs[idx_loc]}',
            'Consultor Negócios': f'Consultor {i%5 + 1}',
            'Status': np.random.choice(['Pendente', 'Em Agendamento', 'Treinado', 'Recusado'], p=[0.4, 0.3, 0.2, 0.1]),
            'SLA_Dias': np.random.randint(2, 25),
            'Latitude': lats[idx_loc] + np.random.uniform(-0.15, 0.15),
            'Longitude': lons[idx_loc] + np.random.uniform(-0.15, 0.15)
        })
    
    # Criar 6 instrutores
    instrutores = [
        {'NOME_COMPLETO': 'Carlos Eduardo Silva', 'UF': 'SP', 'Cidade': 'São Paulo', 'Latitude': -23.5505, 'Longitude': -46.6333, 'Especialidade': 'Loja & Pista'},
        {'NOME_COMPLETO': 'Mariana Souza Santos', 'UF': 'RJ', 'Cidade': 'Rio de Janeiro', 'Latitude': -22.9068, 'Longitude': -43.1729, 'Especialidade': 'Liderança'},
        {'NOME_COMPLETO': 'Roberto Alves', 'UF': 'PR', 'Cidade': 'Curitiba', 'Latitude': -25.4284, 'Longitude': -49.2733, 'Especialidade': 'Fast Food AmPm'},
        {'NOME_COMPLETO': 'Fernanda Lima', 'UF': 'RS', 'Cidade': 'Porto Alegre', 'Latitude': -30.0346, 'Longitude': -51.2177, 'Especialidade': 'Loja & Pista'},
        {'NOME_COMPLETO': 'Lucas Oliveira', 'UF': 'MG', 'Cidade': 'Belo Horizonte', 'Latitude': -19.9167, 'Longitude': -43.9345, 'Especialidade': 'Segurança & Processos'},
        {'NOME_COMPLETO': 'Juliana Costa', 'UF': 'BA', 'Cidade': 'Salvador', 'Latitude': -12.9714, 'Longitude': -38.5014, 'Especialidade': 'Excelência Operacional'}
    ]
    
    return pd.DataFrame(lojas), pd.DataFrame(instrutores)

# ==========================================
# GERADOR DE RELATÓRIO PDF (FPDF)
# ==========================================
class PDFReport(FPDF):
    def header(self):
        self.set_fill_color(30, 34, 45)
        self.rect(0, 0, 210, 25, 'F')
        self.set_font('Arial', 'B', 15)
        self.set_text_color(255, 152, 0)
        self.cell(0, 10, 'CRM Operacional AmPm - Ficha de Treinamento', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

def gerar_pdf_ficha(posto_info, instrutor_info, custo_total):
    pdf = PDFReport()
    pdf.add_page()
    pdf.set_font("Arial", size=11)
    
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, txt="1. Dados do Posto Alvo", ln=True)
    pdf.set_font("Arial", size=11)
    pdf.cell(200, 8, txt=f"PV ABADI: {posto_info['PV Abadi']}", ln=True)
    pdf.cell(200, 8, txt=f"Razao Social: {posto_info['Razão Social']}", ln=True)
    pdf.cell(200, 8, txt=f"Localizacao: {posto_info['Municipio']} / {posto_info['UF']}", ln=True)
    pdf.cell(200, 8, txt=f"Consultor de Negocios: {posto_info['Consultor Negócios']}", ln=True)
    
    pdf.ln(5)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, txt="2. Alocacao Logistica do Instrutor", ln=True)
    pdf.set_font("Arial", size=11)
    pdf.cell(200, 8, txt=f"Instrutor Selecionado: {instrutor_info['Instrutor_Sugerido']}", ln=True)
    pdf.cell(200, 8, txt=f"Base Origem: {instrutor_info['Cidade_Instrutor']} / {instrutor_info['UF_Instrutor']}", ln=True)
    pdf.cell(200, 8, txt=f"Distancia em Linha Reta: {instrutor_info['Distancia_km_linha_reta']:.1f} km", ln=True)
    
    pdf.ln(5)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, txt="3. Estimativa de Custos Logisticos", ln=True)
    pdf.set_font("Arial", size=11)
    pdf.cell(200, 8, txt=f"Custo Total Estimado de Deslocamento/Diarias: R$ {custo_total:.2f}", ln=True)
    
    return pdf.output(dest='S').encode('latin-1', errors='replace')

# ==========================================
# CARREGAMENTO E PROCESSAMENTO DE DADOS
# ==========================================
st.sidebar.title("⛽ CRM AmPm")
st.sidebar.caption("Plataforma Integrada de Operações")

file_upload = st.sidebar.file_uploader("Suba uma nova planilha (Excel/CSV):", type=['xlsx', 'csv'])

df_mock_lojas, df_mock_inst = carregar_dados_mock()

if file_upload is not None:
    try:
        if file_upload.name.endswith('.csv'):
            df_base = pd.read_csv(file_upload)
        else:
            df_base = pd.read_excel(file_upload)
        st.sidebar.success("Arquivo carregado com sucesso!")
    except Exception as e:
        st.sidebar.error("Erro ao ler o arquivo. Usando base demonstrativa.")
        df_base = df_mock_lojas
else:
    df_base = df_mock_lojas

df_instrutores = df_mock_inst

# Processamento do PROCV Geográfico / Haversine
lista_recomendacoes = []

for _, loja in df_base.iterrows():
    pv = loja.get('PV Abadi', 0)
    razao = loja.get('Razão Social', 'N/A')
    uf_l = loja.get('UF', 'N/A')
    mun_l = loja.get('Municipio', 'N/A')
    lat_l = loja.get('Latitude', np.nan)
    lon_l = loja.get('Longitude', np.nan)
    
    if pd.notnull(lat_l) and pd.notnull(lon_l):
        distancias = []
        for _, inst in df_instrutores.iterrows():
            d = haversine(lon_l, lat_l, inst['Longitude'], inst['Latitude'])
            distancias.append({
                'Instrutor_Sugerido': inst['NOME_COMPLETO'],
                'UF_Instrutor': inst['UF'],
                'Cidade_Instrutor': inst['Cidade'],
                'Lat_Instrutor': inst['Latitude'],
                'Lon_Instrutor': inst['Longitude'],
                'Distancia_km_linha_reta': d
            })
        
        distancias_df = pd.DataFrame(distancias).sort_values(by='Distancia_km_linha_reta')
        
        for rank, (_, row_inst) in enumerate(distancias_df.head(3).iterrows(), 1):
            lista_recomendacoes.append({
                'PV_ABADI': pv,
                'Razao_Social': razao,
                'UF_Loja': uf_l,
                'Municipio_Loja': mun_l,
                'Lat_Loja': lat_l,
                'Lon_Loja': lon_l,
                'Ranking_Proximidade': rank,
                'Instrutor_Sugerido': row_inst['Instrutor_Sugerido'],
                'UF_Instrutor': row_inst['UF_Instrutor'],
                'Cidade_Instrutor': row_inst['Cidade_Instrutor'],
                'Lat_Instrutor': row_inst['Lat_Instrutor'],
                'Lon_Instrutor': row_inst['Lon_Instrutor'],
                'Distancia_km_linha_reta': row_inst['Distancia_km_linha_reta'],
                'Dias_Treinamento_Necessarios': 3
            })

df_rec = pd.DataFrame(lista_recomendacoes)

# ==========================================
# NAVEGAÇÃO LATERAL (MÓDULOS)
# ==========================================
modulo = st.sidebar.radio(
    "📌 Módulos do Sistema:",
    [
        "📊 Dashboard Executivo & SLAs",
        "📋 Pipeline Kanban",
        "🔍 PROCV & Filtros Avançados",
        "📍 Calculadora & Otimizador de Custos",
        "📞 Call Center & Timeline WhatsApp",
        "👥 Equipe de Instrutores",
        "📂 Relatórios, Importação & PDF"
    ]
)

st.sidebar.divider()
st.sidebar.markdown("🟢 **Status:** Operacional")
st.sidebar.markdown(f"🏪 **Rede:** {len(df_base)} Unidades")

# ==========================================
# MÓDULO 1: DASHBOARD EXECUTIVO
# ==========================================
if modulo == "📊 Dashboard Executivo & SLAs":
    st.subheader("📊 Painel Executivo de Gestão da Rede AmPm")
    st.caption("Acompanhamento de SLAs, status de treinamentos e gargalos operacionais em tempo real.")
    
    c1, c2, c3, c4 = st.columns(4)
    total_lojas = len(df_base)
    pendentes = len(df_base[df_base['Status'] == 'Pendente'])
    treinados = len(df_base[df_base['Status'] == 'Treinado'])
    sla_medio = df_base['SLA_Dias'].mean() if 'SLA_Dias' in df_base.columns else 0
    
    c1.metric("Total de Postos", total_lojas)
    c2.metric("Postos Pendentes", pendentes, delta_color="inverse", delta=f"{pendentes/total_lojas*100:.1f}%")
    c3.metric("Postos Treinados", treinados, delta=f"{treinados/total_lojas*100:.1f}%")
    c4.metric("SLA Médio (Dias)", f"{sla_medio:.1f} dias")
    
    st.divider()
    
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.markdown("##### 🎯 Status das Soluções de Treinamento")
        fig_donut = px.pie(
            df_base, names='Status', hole=0.5,
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig_donut.update_layout(margin=dict(l=20, r=20, t=30, b=20), paper_bgcolor='rgba(0,0,0,0)', font=dict(color="white"))
        st.plotly_chart(fig_donut, use_container_width=True)
        
    with col_g2:
        st.markdown("##### 🔻 Funil de Conversão de Treinamentos")
        funil_data = dict(
            number=[total_lojas, pendentes + treinados, treinados],
            stage=["Solicitados", "Agendados/Em Curso", "Concluídos"]
        )
        fig_funil = px.funnel(funil_data, x='number', y='stage', color_discrete_sequence=['#FF9800'])
        fig_funil.update_layout(margin=dict(l=20, r=20, t=30, b=20), paper_bgcolor='rgba(0,0,0,0)', font=dict(color="white"))
        st.plotly_chart(fig_funil, use_container_width=True)

# ==========================================
# MÓDULO 2: PIPELINE KANBAN
# ==========================================
elif modulo == "📋 Pipeline Kanban":
    st.subheader("📋 Quadro Kanban Operacional de Atendimento")
    st.caption("Acompanhe e movimente o status dos postos ao longo do ciclo de treinamento.")
    
    cols_kanban = st.columns(4)
    status_list = ['Pendente', 'Em Agendamento', 'Treinado', 'Recusado']
    
    for idx, st_nome in enumerate(status_list):
        with cols_kanban[idx]:
            st.markdown(f"### {st_nome}")
            df_st = df_base[df_base['Status'] == st_nome]
            st.caption(f"{len(df_st)} postos nesta etapa")
            st.divider()
            
            for _, row in df_st.head(8).iterrows():
                st.markdown(f"""
                    <div style="background-color: #1E222D; border: 1px solid #2D333F; padding: 10px; border-radius: 6px; margin-bottom: 8px;">
                        <b>PV {row.get('PV Abadi', '')}</b><br>
                        <small>{row.get('Razão Social', '')[:25]}...</small><br>
                        <span class="badge-info">{row.get('UF', '')}</span>
                    </div>
                """, unsafe_allow_html=True)

# ==========================================
# MÓDULO 3: PROCV & FILTROS AVANÇADOS
# ==========================================
elif modulo == "🔍 PROCV & Filtros Avançados":
    st.subheader("🔍 Cruzamento Inteligente (PROCV) & Filtros")
    st.caption("Explore a recomendação de instrutores ordenada pela menor distância em linha reta.")
    
    col_f1, col_f2 = st.columns(2)
    uf_sel = col_f1.multiselect("Filtrar por UF da Loja:", df_rec['UF_Loja'].unique() if not df_rec.empty else [])
    rank_sel = col_f2.slider("Mostrar até o Ranking de Proximidade:", 1, 3, 3)
    
    df_filtered = df_rec.copy()
    if uf_sel:
        df_filtered = df_filtered[df_filtered['UF_Loja'].isin(uf_sel)]
    if not df_filtered.empty:
        df_filtered = df_filtered[df_filtered['Ranking_Proximidade'] <= rank_sel]
        
    st.dataframe(df_filtered, use_container_width=True)

# ==========================================
# MÓDULO 4: CALCULADORA & OTIMIZADOR DE CUSTOS
# ==========================================
elif modulo == "📍 Calculadora & Otimizador de Custos":
    st.subheader("📍 Centro Operacional de Logística e Rotas Avançadas")
    st.caption("Análise de rotas, clusterização de postos por proximidade e otimização de custos de deslocamento.")
    
    if not df_rec.empty:
        aba_rota, aba_rede = st.tabs(["🛣️ Análise de Rota Individual", "🌐 Mapa Geral da Rede & Circuitos"])
        
        postos_unicos = df_rec[['PV_ABADI', 'Razao_Social', 'Municipio_Loja', 'UF_Loja']].drop_duplicates()
        postos_unicos['label'] = postos_unicos['PV_ABADI'].astype(str) + " - " + postos_unicos['Razao_Social'] + " (" + postos_unicos['Municipio_Loja'] + "/" + postos_unicos['UF_Loja'] + ")"
        
        with aba_rota:
            posto_sel = st.selectbox("⛽ Selecione o Posto Alvo para Roteamento:", postos_unicos['label'].tolist(), key="sb_posto_logistica")
            pv_sel = int(posto_sel.split(" - ")[0])
            
            top_3 = df_rec[df_rec['PV_ABADI'] == pv_sel].sort_values(by='Ranking_Proximidade').head(3)
            
            if not top_3.empty:
                fig_mapa = go.Figure()
                loja_info = top_3.iloc[0]
                lat_loja = loja_info['Lat_Loja']
                lon_loja = loja_info['Lon_Loja']
                
                fig_mapa.add_trace(go.Scattermapbox(
                    lat=[lat_loja], lon=[lon_loja], mode='markers+text',
                    marker=dict(size=18, color='#FF5252'),
                    text=[f"PV {pv_sel}"], textposition="top center",
                    name="Posto AmPm Alvo", hoverinfo='text',
                    hovertext=f"<b>⛽ POSTO ALVO</b><br>PV: {pv_sel}<br>Razão: {loja_info['Razao_Social']}<br>Cidade: {loja_info['Municipio_Loja']}/{loja_info['UF_Loja']}"
                ))
                
                cores_inst = ['#4CAF50', '#81C784', '#A5D6A7']
                
                for idx, (_, row) in enumerate(top_3.iterrows()):
                    lat_inst = row['Lat_Instrutor']
                    lon_inst = row['Lon_Instrutor']
                    dist_km = row['Distancia_km_linha_reta']
                    nome_inst = row['Instrutor_Sugerido']
                    rank = row['Ranking_Proximidade']
                    
                    fig_mapa.add_trace(go.Scattermapbox(
                        lat=[lat_loja, lat_inst], lon=[lon_loja, lon_inst], mode='lines',
                        line=dict(width=3 if rank == 1 else 1.5, color=cores_inst[idx]),
                        name=f"Rota #{rank} ({dist_km:.0f} km)", hoverinfo='text',
                        hovertext=f"<b>Trajeto #{rank}:</b> {dist_km:.1f} km"
                    ))
                    
                    fig_mapa.add_trace(go.Scattermapbox(
                        lat=[lat_inst], lon=[lon_inst], mode='markers+text',
                        marker=dict(size=14 if rank == 1 else 10, color=cores_inst[idx]),
                        text=[f"#{rank} {nome_inst.split()[0]}"], textposition="bottom center",
                        name=f"#{rank} {nome_inst}", hoverinfo='text',
                        hovertext=f"<b>👨‍🏫 INSTRUTOR #{rank}</b><br>{nome_inst}<br>Origem: {row['Cidade_Instrutor']}/{row['UF_Instrutor']}<br>Distância: {dist_km:.1f} km"
                    ))
                
                fig_mapa.update_layout(
                    mapbox=dict(
                        style="carto-darkmatter",
                        center=dict(lat=lat_loja, lon=lon_loja),
                        zoom=6.5
                    ),
                    margin=dict(l=0, r=0, t=30, b=0), height=520,
                    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                    legend=dict(yanchor="top", y=0.98, xanchor="left", x=0.01, bgcolor="rgba(20, 23, 29, 0.85)", font=dict(color="white"))
                )
                
                st.markdown("#### 🗺️ Diagrama de Rotas e Proximidade Logística")
                st.plotly_chart(fig_mapa, use_container_width=True)
                
                st.divider()
                st.markdown("#### 💰 Estimador de Custos de Viagem e Alocação")
                
                with st.expander("⚙️ **Parâmetros e Premissas Financeiras de Viagem**", expanded=False):
                    ca1, ca2, ca3, ca4 = st.columns(4)
                    v_km = ca1.number_input("Custo Terrestre/KM (R$):", value=2.10)
                    v_passagem = ca2.number_input("Passagem Aérea (R$):", value=1400.0)
                    v_diaria = ca3.number_input("Diária Instrutor (R$):", value=280.0)
                    v_traslado = ca4.number_input("Traslado/Uber (R$):", value=150.0)

                col1, col2, col3 = st.columns(3)
                cols = [col1, col2, col3]
                custos_calculados = []

                for idx, (_, row) in enumerate(top_3.iterrows()):
                    dist = row['Distancia_km_linha_reta']
                    dias = row['Dias_Treinamento_Necessarios']
                    
                    if dist <= 300:
                        modal = "🚗 Terrestre"
                        c_desloc = (dist * 2) * v_km
                        c_aereo = 0
                    else:
                        modal = "✈️ Aéreo"
                        c_desloc = v_traslado
                        c_aereo = v_passagem
                        
                    c_hospedagem = dias * v_diaria
                    custo_total = c_desloc + c_aereo + c_hospedagem
                    custos_calculados.append(custo_total)

                    with cols[idx]:
                        borda_cor = "#4CAF50" if idx == 0 else "#2D333F"
                        st.markdown(f"""
                            <div class="top-instructor-card" style="border-left-color: {borda_cor};">
                                <span class="badge-info" style="float:right;">#{row['Ranking_Proximidade']}º Opção</span>
                                <h4 style="margin:0 0 8px 0; color:#FF9800;">{row['Instrutor_Sugerido']}</h4>
                                <p style="margin:2px 0;">🏙️ <b>Origem:</b> {row['Cidade_Instrutor']}/{row['UF_Instrutor']}</p>
                                <p style="margin:2px 0;">📏 <b>Distância Direta:</b> <code>{dist:.1f} km</code></p>
                                <p style="margin:2px 0;">🚀 <b>Modal Indicado:</b> {modal}</p>
                                <hr style="border-color:#2D333F; margin:10px 0;">
                                <p style="margin:2px 0; font-size:0.85rem;">• Deslocamento: R$ {c_desloc:.2f}</p>
                                <p style="margin:2px 0; font-size:0.85rem;">• Aéreo/Passagens: R$ {c_aereo:.2f}</p>
                                <p style="margin:2px 0; font-size:0.85rem;">• Hospedagem/Diárias ({dias}d): R$ {c_hospedagem:.2f}</p>
                                <h3 style="color:#4CAF50; margin:12px 0 0 0;">Total: R$ {custo_total:.2f}</h3>
                            </div>
                        """, unsafe_allow_html=True)

                if len(custos_calculados) >= 2:
                    economia = custos_calculados[1] - custos_calculados[0]
                    st.success(f"💡 **Decisão Inteligente de Logística:** Alocar a **1ª Opção ({top_3.iloc[0]['Instrutor_Sugerido']})** gera uma **economia imediata de R$ {economia:.2f}** frente à segunda alternativa.")

        with aba_rede:
            st.markdown("#### 🌐 Distribuição Geográfica Global (Rede AmPm & Instrutores)")
            st.caption("Identifique densidade de lojas por estado e agrupe visitas em circuitos regionais.")
            
            fig_global = go.Figure()
            
            if 'Latitude' in df_base.columns and 'Longitude' in df_base.columns:
                df_lojas_geo = df_base.dropna(subset=['Latitude', 'Longitude'])
                fig_global.add_trace(go.Scattermapbox(
                    lat=df_lojas_geo['Latitude'], lon=df_lojas_geo['Longitude'], mode='markers',
                    marker=dict(size=7, color='#E27B00', opacity=0.7), name="Postos AmPm",
                    hoverinfo='text', hovertext=df_lojas_geo['PV Abadi'].astype(str) + " - " + df_lojas_geo['Razão Social'] + "<br>" + df_lojas_geo['Municipio'] + "/" + df_lojas_geo['UF']
                ))
            
            if not df_instrutores.empty and 'Latitude' in df_instrutores.columns:
                df_inst_geo = df_instrutores.dropna(subset=['Latitude', 'Longitude'])
                fig_global.add_trace(go.Scattermapbox(
                    lat=df_inst_geo['Latitude'], lon=df_inst_geo['Longitude'], mode='markers+text',
                    marker=dict(size=14, color='#4CAF50'),
                    text=df_inst_geo['NOME_COMPLETO'].str.split().str[0], textposition="bottom center",
                    name="Base Instrutores", hoverinfo='text',
                    hovertext="<b>👨‍🏫 Instrutor:</b> " + df_inst_geo['NOME_COMPLETO'] + "<br>" + df_inst_geo['Cidade'] + "/" + df_inst_geo['UF']
                ))
                
            fig_global.update_layout(
                mapbox=dict(
                    style="carto-darkmatter", center=dict(lat=-14.2350, lon=-51.9253), zoom=3.8
                ),
                margin=dict(l=0, r=0, t=10, b=0), height=550,
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                legend=dict(bgcolor="rgba(20, 23, 29, 0.85)", font=dict(color="white"))
            )
            
            st.plotly_chart(fig_global, use_container_width=True)

# ==========================================
# MÓDULO 5: CALL CENTER & WHATSAPP
# ==========================================
elif modulo == "📞 Call Center & Timeline WhatsApp":
    st.subheader("📞 Central de Relacionamento & Registros")
    st.caption("Simulação de disparos de mensagens e histórico de interações.")
    
    col_w1, col_w2 = st.columns([1, 2])
    
    with col_w1:
        st.markdown("##### 💬 Disparo de Notificação")
        posto_contato = st.selectbox("Selecione o Posto:", df_base['PV Abadi'].unique() if 'PV Abadi' in df_base.columns else [])
        msg = st.text_area("Mensagem do WhatsApp:", "Olá! Confirmamos o agendamento do treinamento operacional para o seu posto AmPm.")
        if st.button("📲 Enviar Notificação via WhatsApp", use_container_width=True):
            st.success("Mensagem enviada para a fila de transmissão!")
            
    with col_w2:
        st.markdown("##### 🕒 Linha do Tempo de Interações")
        st.markdown("""
            * **05/08 14:30** - 📲 *WhatsApp*: Confirmação de data enviada ao Gerente.
            * **04/08 10:15** - 📞 *Ligação*: Agendamento prévio com Consultor de Negócios.
            * **01/08 09:00** - ✉️ *E-mail*: Solicitação inicial de treinamento registrada no sistema.
        """)

# ==========================================
# MÓDULO 6: EQUIPE DE INSTRUTORES
# ==========================================
elif modulo == "👥 Equipe de Instrutores":
    st.subheader("👥 Quadro de Capilaridade dos Instrutores")
    st.caption("Bases operacionais e áreas de cobertura dos instrutores cadastrados.")
    st.dataframe(df_instrutores, use_container_width=True)

# ==========================================
# MÓDULO 7: RELATÓRIOS, IMPORTAÇÃO & PDF
# ==========================================
elif modulo == "📂 Relatórios, Importação & PDF":
    st.subheader("📂 Emissão de Fichas e Exportação de Relatórios")
    st.caption("Gere arquivos PDF oficiais da Ficha de Atendimento para impressão e envio ao Consultor.")
    
    if not df_rec.empty:
        p_sel_pdf = st.selectbox("Selecione o Posto para Gerar a Ficha em PDF:", df_rec['PV_ABADI'].unique())
        top_pdf = df_rec[df_rec['PV_ABADI'] == p_sel_pdf].sort_values(by='Ranking_Proximidade').iloc[0]
        p_info = df_base[df_base['PV Abadi'] == p_sel_pdf].iloc[0] if 'PV Abadi' in df_base.columns else top_pdf
        
        dist_pdf = top_pdf['Distancia_km_linha_reta']
        custo_pdf = (dist_pdf * 2 * 2.10) + (3 * 280) if dist_pdf <= 300 else 1400 + 150 + (3 * 280)
        
        pdf_bytes = gerar_pdf_ficha(p_info, top_pdf, custo_pdf)
        
        st.download_button(
            label="📄 Baixar Ficha Operacional de Treinamento (PDF)",
            data=pdf_bytes,
            file_name=f"Ficha_Treinamento_PV_{p_sel_pdf}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
