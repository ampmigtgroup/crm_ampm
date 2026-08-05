import streamlit as st
import pandas as pd
import os
from datetime import datetime

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
    .form-box {
        background-color: #1e222a;
        padding: 20px;
        border-radius: 10px;
        border-top: 4px solid #e0a96d;
    }
    .modal-badge-aereo {
        background-color: #0288D1;
        color: white;
        padding: 4px 8px;
        border-radius: 4px;
        font-weight: bold;
        font-size: 0.85rem;
    }
    .modal-badge-terrestre {
        background-color: #388E3C;
        color: white;
        padding: 4px 8px;
        border-radius: 4px;
        font-weight: bold;
        font-size: 0.85rem;
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
                         'Semana_Sugerida', 'Telefone_Contato', 'Status_Contato', 
                         'Data_do_Contato', 'Observacoes']],
                left_on='PV Abadi', right_on='PV_Abadi', how='left'
            )
            
            # PROCV 2: Cruzando com Previsão de Inaugurações
            df_base = pd.merge(
                df_base,
                df_inaug[['PV ABADI', 'Previsão Inauguração', 'Pipeline', 'Consultor_Possivel_Instrutor']],
                left_on='PV Abadi', right_on='PV ABADI', how='left'
            )
            
            # PROCV 3: Cruzando Geocodificação dos Instrutores na tabela de Recomendação
            df_rec = pd.merge(
                df_rec,
                df_instrutores[['NOME_COMPLETO', 'Latitude', 'Longitude']],
                left_on='Instrutor_Sugerido', right_on='NOME_COMPLETO', how='left'
            ).rename(columns={'Latitude': 'Lat_Instrutor', 'Longitude': 'Lon_Instrutor'})
            
            # PROCV 4: Cruzando Geocodificação das Lojas na tabela de Recomendação
            df_rec = pd.merge(
                df_rec,
                df_lojas[['PV Abadi', 'Latitude', 'Longitude']],
                left_on='PV_ABADI', right_on='PV Abadi', how='left'
            ).rename(columns={'Latitude': 'Lat_Loja', 'Longitude': 'Lon_Loja'})

            # Tratamento de valores nulos
            df_base['Status_Contato'] = df_base['Status_Contato'].fillna('A Contatar')
            df_base['Tipo_Necessidade'] = df_base['Tipo_Necessidade'].fillna('Rede Ativa (Sem Pendência)')
            df_base['Instrutor_Sugerido'] = df_base['Instrutor_Sugerido'].fillna('Pendente de Alocação')
            df_base['Nome_Contato'] = ""
            df_base['Qtd_Funcionarios'] = 0
            df_base['Material_Em_Loja'] = "N/A"
            
            return df_base, df_instrutores, df_rec
        except Exception as e:
            st.error(f"Erro ao ler planilhas: {e}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    else:
        st.warning("⚠️ Arquivo 'Base_Unificada_AmPm.xlsx' não localizado.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# Inicialização da Session State
if 'df_base' not in st.session_state:
    b, i, r = carregar_bases_integradas()
    st.session_state['df_base'] = b
    st.session_state['df_instrutores'] = i
    st.session_state['df_rec'] = r

df_base = st.session_state['df_base']
df_instrutores = st.session_state['df_instrutores']
df_rec = st.session_state['df_rec']

# --- MENU LATERAL DE NAVEGAÇÃO ---
st.sidebar.title("⛽ Menu CRM AmPm")
modulo = st.sidebar.radio(
    "Selecione o Módulo:",
    [
        "📊 Dashboard Executivo", 
        "🔍 PROCV & Gestão de Lojas", 
        "📍 Menor Custo & Geodeslocamento (Top 3)", 
        "📞 Fila Call Center & Registro de Contatos", 
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
                st.markdown(f"**Contato:** {p.get('Nome_Contato', '-')}")
                st.markdown(f"**Nº Funcionários:** {p.get('Qtd_Funcionarios', 0)}")
                st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# MÓDULO 3: MENOR CUSTO & GEODESLOCAMENTO (TOP 3) + MINI MAPA
# ==========================================
elif modulo == "📍 Menor Custo & Geodeslocamento (Top 3)":
    st.title("📍 Otimizador de Deslocamento & Rotas — Top 3 Instrutores")
    st.markdown("Análise geográfica de rotas para redução de custos com passagens e estadias.")
    
    if not df_rec.empty:
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
            st.subheader("🥇 Top 3 Opções de Instrutores & Modal Recomendado")
            
            col1, col2, col3 = st.columns(3)
            cols = [col1, col2, col3]
            
            for idx, (_, row) in enumerate(top_3.iterrows()):
                dist = row['Distancia_km_linha_reta']
                if dist > 300:
                    modal = "✈️ Aéreo (Avião + Conexão)"
                    badge = "<span class='modal-badge-aereo'>Viagem Aérea</span>"
                else:
                    modal = "🚗 Terrestre (Carro / Uber / Ônibus)"
                    badge = "<span class='modal-badge-terrestre'>Deslocamento Terrestre</span>"
                    
                if idx < 3:
                    with cols[idx]:
                        st.markdown(f"<div class='top-instructor'>", unsafe_allow_html=True)
                        st.markdown(f"#### #{row['Ranking_Proximidade']}º — {row['Instrutor_Sugerido']}")
                        st.markdown(f"📍 **Origem:** {row['Cidade_Instrutor']} / {row['UF_Instrutor']}")
                        st.markdown(f"📏 **Distância:** `{dist} km`")
                        st.markdown(f"🚌 **Modal Indicado:** {modal}")
                        st.markdown(badge, unsafe_allow_html=True)
                        st.markdown("</div>", unsafe_allow_html=True)
            
            st.divider()
            
            # --- MINI MAPA DE VISUALIZAÇÃO DE VIAGEM DO INSTRUTOR ---
            st.subheader("🗺️ Mini Mapa de Viagem do Instrutor")
            
            col_selecao, col_mapa = st.columns([1, 2])
            
            with col_selecao:
                st.markdown("##### 🔍 Selecione o Instrutor:")
                instrutor_opcoes = top_3['Instrutor_Sugerido'].tolist()
                inst_selecionado = st.radio(
                    "Clique para gerar a rota:",
                    instrutor_opcoes
                )
                
                dados_rota = top_3[top_3['Instrutor_Sugerido'] == inst_selecionado].iloc[0]
                
                dist_sel = dados_rota['Distancia_km_linha_reta']
                modal_sel = "✈️ Aéreo (Avião)" if dist_sel > 300 else "🚗 Terrestre (Uber / Carro)"
                
                st.info(f"""
                **Resumo da Rota:**
                - **Origem:** {dados_rota['Cidade_Instrutor']}/{dados_rota['UF_Instrutor']}
                - **Destino:** {dados_rota['Municipio_Loja']}/{dados_rota['UF_Loja']}
                - **Distância:** `{dist_sel} km`
                - **Meio Sugerido:** {modal_sel}
                """)
                
            with col_mapa:
                # Conversão explícita para garantir floats válidos
                lat_loja = pd.to_numeric(dados_rota.get('Lat_Loja'), errors='coerce')
                lon_loja = pd.to_numeric(dados_rota.get('Lon_Loja'), errors='coerce')
                lat_inst = pd.to_numeric(dados_rota.get('Lat_Instrutor'), errors='coerce')
                lon_inst = pd.to_numeric(dados_rota.get('Lon_Instrutor'), errors='coerce')
                
                if pd.notna(lat_loja) and pd.notna(lon_loja) and pd.notna(lat_inst) and pd.notna(lon_inst):
                    df_mini_mapa = pd.DataFrame({
                        'lat': [float(lat_inst), float(lat_loja)],
                        'lon': [float(lon_inst), float(lon_loja)]
                    })
                    st.map(df_mini_mapa, zoom=5, use_container_width=True)
                    st.caption(f"🗺️ **Trajeto:** Partida de {dados_rota['Cidade_Instrutor']} até o Posto {dados_rota['Razao_Social']} ({dados_rota['Municipio_Loja']}).")
                else:
                    st.warning("⚠️ Coordenadas geográficas indisponíveis na base para este posto/instrutor.")

            st.divider()
            st.markdown("### 📊 Tabela de Comparação Logística")
            st.dataframe(
                top_3[['Ranking_Proximidade', 'Instrutor_Sugerido', 'Cidade_Instrutor', 'UF_Instrutor', 'Distancia_km_linha_reta']],
                use_container_width=True,
                hide_index=True
            )

# ==========================================
# MÓDULO 4: FILA CALL CENTER & REGISTRO DE CONTATO
# ==========================================
elif modulo == "📞 Fila Call Center & Registro de Contatos":
    st.title("📞 Fila de Atendimento do Call Center & Registro Pós-Contato")
    st.markdown("Selecione um posto da fila para abrir a ficha de atendimento e registrar as informações do contato.")
    
    if not df_base.empty:
        df_fila_view = df_base[df_base['Tipo_Necessidade'] != 'Rede Ativa (Sem Pendência)'].copy()
        
        col1_tabela, col2_form = st.columns([1.7, 1.3])
        
        with col1_tabela:
            st.subheader("📋 Lojas Pendentes / Agendadas")
            cols_exibicao = ['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Tipo_Necessidade', 'Status_Contato']
            
            evento_call = st.dataframe(
                df_fila_view[cols_exibicao],
                use_container_width=True,
                hide_index=True,
                selection_mode="single-row",
                on_select="rerun"
            )
            
            selecionado = evento_call.selection.get("rows", [])
            
        with col2_form:
            st.subheader("📝 Ficha Pós-Contato")
            
            if selecionado:
                idx_sel = selecionado[0]
                posto = df_fila_view.iloc[idx_sel]
                pv_alvo = posto['PV Abadi']
                tipo_nec = str(posto.get('Tipo_Necessidade', '')).lower()
                
                is_inauguracao = 'inaugura' in tipo_nec or pd.notna(posto.get('Previsão Inauguração'))
                
                with st.form(key="form_registro_callcenter"):
                    st.markdown(f"<div class='form-box'>", unsafe_allow_html=True)
                    st.markdown(f"### ⛽ PV {posto['PV Abadi']} — {posto['Razão Social']}")
                    st.markdown(f"**Cidade/UF:** {posto['Municipio']}/{posto['UF']}")
                    st.markdown(f"**Necessidade:** `{posto['Tipo_Necessidade']}`")
                    
                    st.divider()
                    
                    nome_contato = st.text_input(
                        "👤 Nome do Contato / Responsável:", 
                        value=str(posto.get('Nome_Contato', '') if pd.notna(posto.get('Nome_Contato')) else '')
                    )
                    
                    c_tel, c_func = st.columns(2)
                    with c_tel:
                        telefone_contato = st.text_input(
                            "📞 Telefone de Contato:", 
                            value=str(posto.get('Telefone_Contato', '') if pd.notna(posto.get('Telefone_Contato')) else '')
                        )
                    with c_func:
                        qtd_funcionarios = st.number_input(
                            "👥 Nº de Funcionários:", 
                            min_value=0, 
                            max_value=200, 
                            value=int(posto.get('Qtd_Funcionarios', 0)) if pd.notna(posto.get('Qtd_Funcionarios')) else 0
                        )
                    
                    material_loja = "N/A"
                    if is_inauguracao:
                        st.markdown("---")
                        st.markdown("📦 **Controle de Insumos (Exclusivo Inaugurações)**")
                        material_loja = st.radio(
                            "Todo o material de treinamento já está disponível na loja?",
                            ["Sim — Material Completo na Loja", "Não — Aguardando Chegada dos Materiais", "Parcial — Entregue Incompleto"],
                            index=0
                        )
                    
                    st.markdown("---")
                    
                    novo_status = st.selectbox(
                        "Status do Contato:",
                        ["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado / Indisponível"],
                        index=["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado / Indisponível"].index(posto.get('Status_Contato', 'A Contatar')) if posto.get('Status_Contato') in ["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado / Indisponível"] else 0
                    )
                    
                    lista_instrutores = ["Pendente de Alocação"]
                    if not df_instrutores.empty and 'NOME_COMPLETO' in df_instrutores.columns:
                        lista_instrutores += sorted(df_instrutores['NOME_COMPLETO'].dropna().unique().tolist())
                        
                    novo_instrutor = st.selectbox(
                        "Instrutor Sugerido / Confirmado:",
                        lista_instrutores,
                        index=lista_instrutores.index(posto.get('Instrutor_Sugerido')) if posto.get('Instrutor_Sugerido') in lista_instrutores else 0
                    )
                    
                    semana_agendada = st.text_input("Semana Agendada (Ex: Sem 35):", value=str(posto.get('Semana_Sugerida', '') if pd.notna(posto.get('Semana_Sugerida')) else ''))
                    observacoes_contato = st.text_area("Observações do Atendimento:", value=str(posto.get('Observacoes', '') if pd.notna(posto.get('Observacoes')) else ''))
                    
                    btn_salvar = st.form_submit_button("💾 Salvar Informações do Atendimento")
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                    if btn_salvar:
                        mask = st.session_state['df_base']['PV Abadi'] == pv_alvo
                        st.session_state['df_base'].loc[mask, 'Nome_Contato'] = nome_contato
                        st.session_state['df_base'].loc[mask, 'Qtd_Funcionarios'] = qtd_funcionarios
                        st.session_state['df_base'].loc[mask, 'Material_Em_Loja'] = material_loja
                        st.session_state['df_base'].loc[mask, 'Status_Contato'] = novo_status
                        st.session_state['df_base'].loc[mask, 'Instrutor_Sugerido'] = novo_instrutor
                        st.session_state['df_base'].loc[mask, 'Semana_Sugerida'] = semana_agendada
                        st.session_state['df_base'].loc[mask, 'Telefone_Contato'] = telefone_contato
                        st.session_state['df_base'].loc[mask, 'Observacoes'] = observacoes_contato
                        st.session_state['df_base'].loc[mask, 'Data_do_Contato'] = datetime.today().strftime('%d/%m/%Y %H:%M')
                        
                        st.success(f"✅ Atendimento do PV {pv_alvo} atualizado com sucesso!")
                        st.rerun()
            else:
                st.info("👈 Selecione um posto na tabela ao lado para carregar o formulário.")

# ==========================================
# MÓDULO 5: GESTÃO DE INSTRUTORES
# ==========================================
elif modulo == "👔 Gestão de Instrutores":
    st.title("👔 Relação de Instrutores")
    if not df_instrutores.empty:
        st.dataframe(df_instrutores[['NOME_COMPLETO', 'STATUS', 'TELEFONE', 'EMAIL', 'Cidade', 'UF']], use_container_width=True, hide_index=True)
