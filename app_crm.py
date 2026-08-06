import streamlit as st
import pandas as pd
import sqlite3
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

DB_PATH = "crm_ampm.db"

# --- INICIALIZAÇÃO DO BANCO DE DADOS SQLITE ---
def get_db_connection():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Tabela Operacional (Editável pelo CRM)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS fila_callcenter (
            pv_abadi INTEGER PRIMARY KEY,
            razao_social TEXT,
            municipio TEXT,
            uf TEXT,
            endereco TEXT,
            status_loja TEXT,
            gf TEXT,
            cf TEXT,
            tipo_necessidade TEXT,
            dias_desde_ultimo_treinamento TEXT,
            previsao_inauguracao TEXT,
            telefone_contato TEXT,
            nome_contato TEXT,
            qtd_funcionarios INTEGER,
            material_em_loja TEXT,
            status_contato TEXT,
            instrutor_sugerido TEXT,
            data_agendada TEXT,
            observacoes TEXT,
            data_do_contato TEXT
        )
    """)
    
    # Tabela de Instrutores (ESTÁTICA / APENAS LEITURA - NUNCA ALTERADA)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS instrutores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_completo TEXT UNIQUE,
            cidade TEXT,
            uf TEXT,
            latitude REAL,
            longitude REAL
        )
    """)
    
    conn.commit()
    conn.close()

init_db()

# --- CARREGAMENTO DE DADOS ---
@st.cache_data
def carregar_dados_sqlite():
    conn = get_db_connection()
    
    df_base = pd.read_sql_query("SELECT * FROM fila_callcenter", conn)
    df_instrutores = pd.read_sql_query("SELECT * FROM instrutores", conn)
    
    # Se a base local SQLite estiver vazia, tenta popular via Excel inicial
    caminho_excel = "Base_Unificada_AmPm.xlsx"
    if df_base.empty and os.path.exists(caminho_excel):
        try:
            xls = pd.ExcelFile(caminho_excel, engine='openpyxl')
            df_lojas = pd.read_excel(xls, sheet_name='Rede_de_Lojas')
            df_fila = pd.read_excel(xls, sheet_name='Fila_CallCenter')
            df_inst_excel = pd.read_excel(xls, sheet_name='Instrutores')
            
            df_lojas['PV Abadi'] = pd.to_numeric(df_lojas['PV Abadi'], errors='coerce')
            df_fila['PV_Abadi'] = pd.to_numeric(df_fila['PV_Abadi'], errors='coerce')
            
            df_merged = pd.merge(
                df_lojas,
                df_fila[['PV_Abadi', 'Tipo_Necessidade', 'Dias_desde_Ultimo_Treinamento', 
                         'Instrutor_Sugerido', 'Telefone_Contato', 'Status_Contato', 
                         'Data_do_Contato', 'Observacoes']],
                left_on='PV Abadi', right_on='PV_Abadi', how='left'
            )
            
            df_db = pd.DataFrame({
                'pv_abadi': df_merged['PV Abadi'],
                'razao_social': df_merged.get('Razão Social', ''),
                'municipio': df_merged.get('Municipio', ''),
                'uf': df_merged.get('UF', ''),
                'endereco': df_merged.get('Endereço', ''),
                'status_loja': df_merged.get('Status Loja', ''),
                'gf': df_merged.get('GF', ''),
                'cf': df_merged.get('CF', ''),
                'tipo_necessidade': df_merged.get('Tipo_Necessidade', 'Rede Ativa (Sem Pendência)'),
                'dias_desde_ultimo_treinamento': df_merged.get('Dias_desde_Ultimo_Treinamento', 'N/A'),
                'previsao_inauguracao': 'N/A',
                'telefone_contato': df_merged.get('Telefone_Contato', ''),
                'nome_contato': '',
                'qtd_funcionarios': 0,
                'material_em_loja': 'Não Informado',
                'status_contato': df_merged.get('Status_Contato', 'A Contatar').fillna('A Contatar'),
                'instrutor_sugerido': df_merged.get('Instrutor_Sugerido', 'Pendente de Alocação').fillna('Pendente de Alocação'),
                'data_agendada': None,
                'observacoes': df_merged.get('Observacoes', ''),
                'data_do_contato': df_merged.get('Data_do_Contato', '')
            }).drop_duplicates(subset=['pv_abadi'])
            
            df_db.to_sql('fila_callcenter', conn, if_exists='replace', index=False)
            
            # Popula instrutores (Sem permitir escrita posterior)
            if not df_inst_excel.empty:
                df_inst_db = pd.DataFrame({
                    'nome_completo': df_inst_excel.get('NOME_COMPLETO', ''),
                    'cidade': df_inst_excel.get('Cidade', ''),
                    'uf': df_inst_excel.get('UF', ''),
                    'latitude': df_inst_excel.get('Latitude', 0.0),
                    'longitude': df_inst_excel.get('Longitude', 0.0)
                }).dropna(subset=['nome_completo'])
                df_inst_db.to_sql('instrutores', conn, if_exists='replace', index=False)
            
            df_base = pd.read_sql_query("SELECT * FROM fila_callcenter", conn)
            df_instrutores = pd.read_sql_query("SELECT * FROM instrutores", conn)
        except Exception as e:
            st.error(f"⚠️ Erro na carga inicial do Excel para SQLite: {e}")
            
    conn.close()
    return df_base, df_instrutores

# SALVAMENTO EXCLUSIVO DA FILA OPERACIONAL (INSTRUTORES PERMANECEM INTACTOS)
def atualizar_atendimento_db(pv_abadi, dados_update):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    query = """
        UPDATE fila_callcenter 
        SET nome_contato = ?,
            telefone_contato = ?,
            qtd_funcionarios = ?,
            material_em_loja = ?,
            status_contato = ?,
            instrutor_sugerido = ?,
            data_agendada = ?,
            observacoes = ?,
            data_do_contato = ?
        WHERE pv_abadi = ?
    """
    
    cursor.execute(query, (
        dados_update['nome_contato'],
        dados_update['telefone_contato'],
        dados_update['qtd_funcionarios'],
        dados_update['material_em_loja'],
        dados_update['status_contato'],
        dados_update['instrutor_sugerido'],
        dados_update['data_agendada'],
        dados_update['observacoes'],
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        pv_abadi
    ))
    
    conn.commit()
    conn.close()
    st.cache_data.clear()

if 'df_base' not in st.session_state:
    b, i = carregar_dados_sqlite()
    st.session_state['df_base'] = b
    st.session_state['df_instrutores'] = i

df_base_raw = st.session_state['df_base']
df_instrutores = st.session_state['df_instrutores']

# --- ESTILIZAÇÃO CSS CUSTOMIZADA ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .main-header {
        background: linear-gradient(135deg, #E27B00 0%, #FF9800 50%, #D32F2F 100%);
        padding: 24px 28px; border-radius: 16px; color: white; margin-bottom: 25px;
        box-shadow: 0 8px 24px rgba(226, 123, 0, 0.25);
    }
    .main-header h1 { color: white !important; margin: 0 0 6px 0; font-weight: 700; font-size: 2.2rem; }
    .kpi-card {
        background-color: #1E222A; border-radius: 12px; padding: 20px;
        border: 1px solid #2D333F; border-left: 6px solid #E27B00;
    }
    .kpi-title { font-size: 0.8rem; color: #A0AAB8; text-transform: uppercase; font-weight: 700; }
    .kpi-value { font-size: 1.8rem; font-weight: 700; color: #FFFFFF; margin-top: 8px; }
    .procv-card {
        background-color: #1A1D24; padding: 20px; border-radius: 12px;
        border: 1px solid #2D333F; border-top: 4px solid #E27B00; margin-bottom: 15px;
    }
    .procv-card h4 { margin-top: 0; color: #FF9800; font-size: 1rem; }
    </style>
""", unsafe_allow_html=True)

# --- SIDEBAR DE NAVEGAÇÃO & FILTROS ---
with st.sidebar:
    st.markdown("## ⛽ **CRM AmPm (SQLite)**")
    st.caption("💾 *Dados Persistence Engine Ativado*")
    st.divider()
    
    modulo = st.radio(
        "📌 **Módulos do Sistema:**",
        [
            "📊 Dashboard Executivo", 
            "📋 Pipeline AmPm", 
            "🔍 PROCV & Filtros Avançados", 
            "📞 Call Center & Timeline WhatsApp", 
            "👔 Equipe de Instrutores (Leitura)",
            "📂 Relatórios & Exportação"
        ]
    )
    
    st.divider()
    
    st.markdown("🎯 **Filtros Globais**")
    uf_opcoes = ["Todas"] + sorted([str(x) for x in df_base_raw['uf'].dropna().unique()]) if 'uf' in df_base_raw.columns else ["Todas"]
    filtro_uf = st.selectbox("Filtrar Estado (UF):", uf_opcoes)

    cf_opcoes = ["Todos"] + sorted([str(x) for x in df_base_raw['cf'].dropna().unique()]) if 'cf' in df_base_raw.columns else ["Todos"]
    filtro_cf = st.selectbox("Filtrar Consultor (CF):", cf_opcoes)

    st.divider()
    st.markdown("📶 **Banco de Dados:** `SQLite Conectado 🟢`")
    st.markdown("🔒 **Tabela Instrutores:** `Protegida / Imutável 🛡️`")

# Aplicação dos Filtros Globais
df_base = df_base_raw.copy()
if filtro_uf != "Todas":
    df_base = df_base[df_base['uf'] == filtro_uf]
if filtro_cf != "Todos":
    df_base = df_base[df_base['cf'] == filtro_cf]

# Topbar
st.markdown("""
    <div class="main-header">
        <h1>⛽ CRM Operacional AmPm</h1>
        <p>🚀 Gestão Estratégica com Armazenamento Persistente em Banco de Dados SQLite</p>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# MÓDULO 1: DASHBOARD EXECUTIVO
# ==========================================
if modulo == "📊 Dashboard Executivo":
    if not df_base.empty:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f'<div class="kpi-card"><div class="kpi-title">Rede Filtrada</div><div class="kpi-value">{len(df_base)}</div></div>', unsafe_allow_html=True)
        with c2:
            pendentes = len(df_base[df_base['tipo_necessidade'] != 'Rede Ativa (Sem Pendência)']) if 'tipo_necessidade' in df_base.columns else 0
            st.markdown(f'<div class="kpi-card"><div class="kpi-title">Fila Treinamento</div><div class="kpi-value">{pendentes}</div></div>', unsafe_allow_html=True)
        with c3:
            a_contatar = len(df_base[df_base['status_contato'] == 'A Contatar']) if 'status_contato' in df_base.columns else 0
            st.markdown(f'<div class="kpi-card"><div class="kpi-title">Pendentes Contato</div><div class="kpi-value">{a_contatar}</div></div>', unsafe_allow_html=True)
        with c4:
            st.markdown(f'<div class="kpi-card"><div class="kpi-title">Instrutores Cadastrados</div><div class="kpi-value">{len(df_instrutores)}</div></div>', unsafe_allow_html=True)

        st.divider()
        col_A, col_B = st.columns(2)
        with col_A:
            st.subheader("🗺️ Rede por Estado (UF)")
            if 'uf' in df_base.columns:
                st.bar_chart(df_base['uf'].value_counts().head(10), color="#E27B00")
        with col_B:
            st.subheader("📊 Status dos Contatos")
            if 'status_contato' in df_base.columns:
                st.bar_chart(df_base['status_contato'].value_counts(), color="#FF9800")

# ==========================================
# MÓDULO 2: PIPELINE AMPM
# ==========================================
elif modulo == "📋 Pipeline AmPm":
    st.subheader("📋 Pipeline AmPm — Fluxo Operacional")
    colunas_pipeline = ["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado"]
    cols = st.columns(len(colunas_pipeline))
    
    for idx, status in enumerate(colunas_pipeline):
        df_status = df_base[df_base['status_contato'] == status] if 'status_contato' in df_base.columns else pd.DataFrame()
        with cols[idx]:
            st.markdown(f"### {status} ({len(df_status)})")
            for _, item in df_status.head(5).iterrows():
                with st.expander(f"PV {item.get('pv_abadi')} | {str(item.get('razao_social'))[:12]}..."):
                    st.write(f"**Cidade:** {item.get('municipio')}/{item.get('uf')}")
                    st.write(f"**Instrutor:** {item.get('instrutor_sugerido')}")
                    
                    novo_st = st.selectbox("Status:", colunas_pipeline, index=colunas_pipeline.index(status), key=f"p_{item.get('pv_abadi')}")
                    if novo_st != status:
                        atualizar_atendimento_db(item['pv_abadi'], {
                            'nome_contato': item.get('nome_contato', ''),
                            'telefone_contato': item.get('telefone_contato', ''),
                            'qtd_funcionarios': item.get('qtd_funcionarios', 0),
                            'material_em_loja': item.get('material_em_loja', 'Não Informado'),
                            'status_contato': novo_st,
                            'instrutor_sugerido': item.get('instrutor_sugerido', ''),
                            'data_agendada': item.get('data_agendada'),
                            'observacoes': item.get('observacoes', '')
                        })
                        st.toast("Status atualizado no SQLite!", icon="✅")
                        st.rerun()

# ==========================================
# MÓDULO 3: PROCV & FILTROS AVANÇADOS
# ==========================================
elif modulo == "🔍 PROCV & Filtros Avançados":
    termo = st.text_input("🔍 Pesquisar por PV, Razão Social ou Cidade:", "")
    df_v = df_base.copy()
    if termo:
        df_v = df_v[
            df_v['razao_social'].astype(str).str.contains(termo, case=False, na=False) |
            df_v['pv_abadi'].astype(str).str.contains(termo, na=False) |
            df_v['municipio'].astype(str).str.contains(termo, case=False, na=False)
        ]
    
    cols_m = [c for c in ['pv_abadi', 'razao_social', 'municipio', 'uf', 'status_loja', 'status_contato'] if c in df_v.columns]
    st.dataframe(df_v[cols_m], use_container_width=True, hide_index=True)

# ==========================================
# MÓDULO 4: CALL CENTER & TIMELINE WHATSAPP
# ==========================================
elif modulo == "📞 Call Center & Timeline WhatsApp":
    df_fila = df_base[df_base['tipo_necessidade'] != 'Rede Ativa (Sem Pendência)'].copy()
    c_left, c_right = st.columns([1.2, 1.8])
    
    with c_left:
        st.subheader("📋 Fila de Atendimento")
        evento = st.dataframe(
            df_fila[['pv_abadi', 'razao_social', 'municipio', 'uf', 'status_contato']],
            use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun"
        )
        sel = evento.selection.get("rows", [])
        
    with c_right:
        if sel:
            posto = df_fila.iloc[sel[0]]
            pv_target = posto['pv_abadi']
            
            st.markdown(f"### 📝 Atendimento — **PV {pv_target} | {posto.get('razao_social')}**")
            
            tel_digits = ''.join(filter(str.isdigit, str(posto.get('telefone_contato', ''))))
            if tel_digits:
                link_wa = f"https://wa.me/55{tel_digits}?text=Olá,%20equipe%20da%20loja%20{posto.get('razao_social')}!%20Falamos%20da%20Capacitação%20AmPm."
                st.markdown(f"📲 **[Abrir Conversa Direta no WhatsApp Direct]({link_wa})**")

            # LISTA DE INSTRUTORES CARREGADA APENAS PARA LEITURA/SELEÇÃO
            lista_instrutores = ["Pendente de Alocação"]
            if not df_instrutores.empty and 'nome_completo' in df_instrutores.columns:
                lista_instrutores += sorted(df_instrutores['nome_completo'].dropna().unique().tolist())
                
            inst_atual = str(posto.get('instrutor_sugerido', 'Pendente de Alocação'))
            idx_inst = lista_instrutores.index(inst_atual) if inst_atual in lista_instrutores else 0

            with st.form("form_callcenter_sqlite"):
                st.markdown("#### ✍️ Atualizar Registro no Banco de Dados")
                
                f1, f2 = st.columns(2)
                with f1:
                    nome_c = st.text_input("👤 Contato na Loja:", value=str(posto.get('nome_contato', '')) if pd.notna(posto.get('nome_contato')) else "")
                    tel_c = st.text_input("📞 Telefone:", value=str(posto.get('telefone_contato', '')) if pd.notna(posto.get('telefone_contato')) else "")
                    qtd_f = st.number_input("👥 Treinandos:", min_value=0, value=int(posto.get('qtd_funcionarios', 0)))
                
                with f2:
                    mat_options = ["Não Informado", "Entregue / Em Loja", "Pendente / Em Trânsito"]
                    mat_val = str(posto.get('material_em_loja', 'Não Informado'))
                    idx_m = mat_options.index(mat_val) if mat_val in mat_options else 0
                    mat_c = st.selectbox("📦 Material em Loja:", mat_options, index=idx_m)
                    
                    col_pipe = ["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado"]
                    st_val = str(posto.get('status_contato', 'A Contatar'))
                    idx_st = col_pipe.index(st_val) if st_val in col_pipe else 0
                    st_c = st.selectbox("🔄 Status Contato:", col_pipe, index=idx_st)
                    
                    instrutor_c = st.selectbox("👨‍🏫 Alocar Instrutor:", lista_instrutores, index=idx_inst)
                    dt_c = st.date_input("📅 Data Agendada:", value=date.today())

                obs_c = st.text_area("📝 Histórico / Observações:", value=str(posto.get('observacoes', '')) if pd.notna(posto.get('observacoes')) else "")
                
                if st.form_submit_button("💾 Salvar Atendimento no SQLite"):
                    atualizar_atendimento_db(pv_target, {
                        'nome_contato': nome_c,
                        'telefone_contato': tel_c,
                        'qtd_funcionarios': qtd_f,
                        'material_em_loja': mat_c,
                        'status_contato': st_c,
                        'instrutor_sugerido': instrutor_c,
                        'data_agendada': dt_c.strftime("%Y-%m-%d"),
                        'observacoes': obs_c
                    })
                    st.success("✅ Registrado com sucesso no banco de dados!")
                    st.rerun()

# ==========================================
# MÓDULO 5: EQUIPE DE INSTRUTORES (LEITURA)
# ==========================================
elif modulo == "👔 Equipe de Instrutores (Leitura)":
    st.subheader("👔 Cadastro de Instrutores (Modo Protegido / Somente Leitura)")
    st.info("🛡️ **Aviso de Integridade:** O cadastro da equipe de instrutores é protegido e não sofre modificações através das operações do CRM.")
    
    if not df_instrutores.empty:
        st.dataframe(df_instrutores, use_container_width=True, hide_index=True)
    else:
        st.warning("Nenhum instrutor cadastrado no banco de dados.")

# ==========================================
# MÓDULO 6: RELATÓRIOS & EXPORTAÇÃO
# ==========================================
elif modulo == "📂 Relatórios & Exportação":
    st.subheader("📂 Exportação de Relatórios Operacionais")
    
    buffer_excel = io.BytesIO()
    with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
        df_base.to_excel(writer, sheet_name='Atendimentos_SQLite', index=False)
        df_instrutores.to_excel(writer, sheet_name='Instrutores_Protegido', index=False)
        
    st.download_button(
        label="💾 Baixar Base Completa do SQLite (.xlsx)",
        data=buffer_excel.getvalue(),
        file_name=f"CRM_AmPm_DB_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
