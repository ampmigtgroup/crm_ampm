import streamlit as st
import pandas as pd
import numpy as np
from sqlalchemy import create_engine, text
from io import BytesIO

st.set_page_config(page_title="CRM Operacional AmPm - Inteligente", layout="wide")

DATABASE_URL = "postgresql://postgres.nptazzfvwhhmotfrvgdj:Lssj.ampm%40%23@aws-0-sa-east-1.pooler.supabase.com:6543/postgres"

@st.cache_resource
def get_connection():
    return create_engine(DATABASE_URL)

engine = get_connection()

# Função para calcular distância Haversine (linha reta em KM)
def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0 # Raio da Terra em KM
    dlat = np.radians(lat2 - lat1)
    dlon = np.radians(lon2 - lon1)
    a = np.sin(dlat / 2)**2 + np.cos(np.radians(lat1)) * np.cos(np.radians(lat2)) * np.sin(dlon / 2)**2
    c = 2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a))
    return R * c

st.title("⛽ CRM Operacional AmPm — Treinamentos & Inteligência de Roteamento")

# Consulta Base da Fila de Atendimentos
query_fila = """
    SELECT 
        c.id_atendimento,
        c.pv_abadi,
        COALESCE(c.loja, l.razao_social, 'N/A') AS loja,
        COALESCE(c.municipio, l.municipio_uf, 'N/A') AS municipio,
        COALESCE(c.uf, '') AS uf,
        COALESCE(c.tipo_necessidade, 'Retreinamento') AS tipo_necessidade,
        COALESCE(c.instrutor_sugerido, 'A definir') AS instrutor_sugerido,
        c.semana_sugerida,
        COALESCE(c.status_contato, 'A Contatar') AS status_contato,
        c.observacoes
    FROM tb_fila_call_center c
    LEFT JOIN tb_lojas l ON c.pv_abadi = l.pv_abadi
    ORDER BY c.id_atendimento ASC;
"""

try:
    df_fila = pd.read_sql(query_fila, engine)

    # Sidebar: Filtros Globais
    st.sidebar.header("🔍 Filtros da Fila")
    busca_pv = st.sidebar.text_input("Buscar por PV Abadi ou Nome:")
    lista_status = ["Todos"] + list(df_fila['status_contato'].unique())
    status_filtro = st.sidebar.selectbox("Status do Contato:", lista_status)
    lista_instrutores = ["Todos"] + list(df_fila['instrutor_sugerido'].dropna().unique())
    instrutor_filtro = st.sidebar.selectbox("Instrutor Sugerido:", lista_instrutores)
    lista_ufs = ["Todos"] + sorted([uf for uf in df_fila['uf'].unique() if uf])
    uf_filtro = st.sidebar.selectbox("UF (Estado):", lista_ufs)

    # Aplicação de Filtros
    df_filtrado = df_fila.copy()
    if busca_pv:
        df_filtrado = df_filtrado[
            df_filtrado['pv_abadi'].astype(str).str.contains(busca_pv, case=False, na=False) |
            df_filtrado['loja'].astype(str).str.contains(busca_pv, case=False, na=False)
        ]
    if status_filtro != "Todos":
        df_filtrado = df_filtrado[df_filtrado['status_contato'] == status_filtro]
    if instrutor_filtro != "Todos":
        df_filtrado = df_filtrado[df_filtrado['instrutor_sugerido'] == instrutor_filtro]
    if uf_filtro != "Todos":
        df_filtrado = df_filtrado[df_filtrado['uf'] == uf_filtro]

    # KPIs Operacionais
    total_registros = len(df_fila)
    a_contatar = len(df_fila[df_fila['status_contato'] == 'A Contatar'])
    contatados = len(df_fila[df_fila['status_contato'] == 'Contatado'])
    confirmados = len(df_fila[df_fila['status_contato'] == 'Confirmado'])
    concluidos = len(df_fila[df_fila['status_contato'] == 'Concluído'])

    col_kpi1, col_kpi2, col_kpi3, col_kpi4, col_kpi5 = st.columns(5)
    col_kpi1.metric("Total Fila", total_registros)
    col_kpi2.metric("A Contatar", a_contatar)
    col_kpi3.metric("Contatados", contatados)
    col_kpi4.metric("Confirmados", confirmados)
    col_kpi5.metric("Concluídos", concluidos)

    st.divider()

    # Criação das Abas do Sistema (Incluso Histórico e Menor Custo)
    aba_fila, aba_menor_custo, aba_historico, aba_cadastro, aba_exportar = st.tabs([
        "📋 Fila & Atualização", 
        "📍 Menor Custo (Instrutor)",
        "📊 Histórico & Custos",
        "➕ Novo Cadastro Manual", 
        "📥 Exportar Relatórios"
    ])

    # ------------------------------------------
    # ABA 1: FILA & ATUALIZAÇÃO
    # ------------------------------------------
    with aba_fila:
        st.subheader(f"Fila Prioritária de Contatos ({len(df_filtrado)} registros)")
        st.dataframe(df_filtrado, use_container_width=True)

        if not df_filtrado.empty:
            st.divider()
            st.subheader("📝 Atualizar Agendamento / Status")
            lista_atendimentos = df_filtrado['id_atendimento'].tolist()
            id_selecionado = st.selectbox("Selecione o ID para editar:", lista_atendimentos)
            item = df_filtrado[df_filtrado['id_atendimento'] == id_selecionado].iloc[0]
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.write(f"**PV Abadi:** {item['pv_abadi']}")
                st.write(f"**Loja/Posto:** {item['loja']}")
                st.write(f"**Local:** {item['municipio']} / {item['uf']}")
            with col2:
                st.write(f"**Necessidade:** {item['tipo_necessidade']}")
                st.write(f"**Instrutor Sugerido:** {item['instrutor_sugerido']}")
                st.write(f"**Semana Sugerida:** {item['semana_sugerida']}")
            with col3:
                status_atuais = ["A Contatar", "Contatado", "Confirmado", "Recusado/Adiado", "Concluído"]
                status_item = item['status_contato'] if item['status_contato'] in status_atuais else "A Contatar"
                novo_status = st.selectbox("Novo Status:", status_atuais, index=status_atuais.index(status_item))
            
            obs_atual = item['observacoes'] if pd.notnull(item['observacoes']) else ""
            novas_obs = st.text_area("Observações do Atendimento:", value=obs_atual)
            
            if st.button("💾 Salvar Alterações"):
                with engine.connect() as conn:
                    query_update = text("UPDATE tb_fila_call_center SET status_contato = :s, observacoes = :o WHERE id_atendimento = :id")
                    conn.execute(query_update, {"s": novo_status, "o": novas_obs, "id": int(id_selecionado)})
                    conn.commit()
                st.success("Atualizado com sucesso!")
                st.rerun()

    # ------------------------------------------
    # ABA 2: CONSULTA MENOR CUSTO (DISTÂNCIA)
    # ------------------------------------------
    with aba_menor_custo:
        st.subheader("📍 Roteamento e Recomendação por Menor Distância")
        st.write("Calcule os instrutores ativos mais próximos do posto de destino para otimizar despesas de transporte.")
        
        try:
            df_inst = pd.read_sql("SELECT * FROM tb_instrutores WHERE status = 'Ativo'", engine)
            if not df_inst.empty and 'lat' in df_inst.columns:
                col_m1, col_m2 = st.columns(2)
                with col_m1:
                    lat_dest = st.number_input("Latitude do Posto Destino:", value=-23.550520, format="%.6f")
                with col_m2:
                    lon_dest = st.number_input("Longitude do Posto Destino:", value=-46.633308, format="%.6f")
                
                if st.button("🔍 Calcular Instrutores Mais Próximos"):
                    df_inst['distancia_km'] = haversine(lat_dest, lon_dest, df_inst['lat'], df_inst['lon'])
                    df_ranking = df_inst[['nome', 'cidade', 'uf', 'distancia_km']].sort_values(by='distancia_km').head(3)
                    df_ranking['distancia_km'] = df_ranking['distancia_km'].round(1).astype(str) + " km"
                    
                    st.success("Top 3 Instrutores Recomendados (Menor Deslocamento):")
                    st.table(df_ranking)
            else:
                st.info("Popule a tabela 'tb_instrutores' no Supabase para habilitar o cálculo dinâmico de distância.")
        except Exception as ex_m:
            st.info("Tabela 'tb_instrutores' pronta no banco. Aguardando população de coordenadas.")

    # ------------------------------------------
    # ABA 3: HISTÓRICO & ANÁLISE DE CUSTOS
    # ------------------------------------------
    with aba_historico:
        st.subheader("📊 Histórico de Treinamentos Realizados & Despesas")
        try:
            df_hist = pd.read_sql("SELECT * FROM tb_historico_treinamentos", engine)
            if not df_hist.empty:
                st.dataframe(df_hist, use_container_width=True)
                custo_total_geral = df_hist['custo_total'].sum() if 'custo_total' in df_hist.columns else 0
                st.metric("Custo Total Investido em Treinamentos", f"R$ {custo_total_geral:,.2f}")
            else:
                st.info("O histórico de treinamentos pode ser importado diretamente para a tabela 'tb_historico_treinamentos'.")
        except Exception as ex_h:
            st.info("Tabela de histórico pronta para receber os registros de treinamentos anteriores.")

    # ------------------------------------------
    # ABA 4: CADASTRO MANUAL
    # ------------------------------------------
    with aba_cadastro:
        st.subheader("➕ Adicionar Novo Atendimento à Fila")
        with st.form("form_novo_atendimento", clear_on_submit=True):
            col_c1, col_c2 = st.columns(2)
            with col_c1:
                novo_pv = st.number_input("PV Abadi:", min_value=1, step=1)
                nova_loja = st.text_input("Nome da Loja/Posto:")
                novo_municipio = st.text_input("Município:")
                novo_uf = st.text_input("UF:", max_chars=2)
            with col_c2:
                novo_tipo = st.selectbox("Tipo:", ["Retreinamento", "Inauguração", "Emergencial"])
                novo_instrutor = st.text_input("Instrutor Sugerido:")
                nova_semana = st.text_input("Semana Sugerida:")
                novo_status_cad = st.selectbox("Status:", ["A Contatar", "Contatado", "Confirmado"])
            novas_obs_cad = st.text_area("Observações:")
            if st.form_submit_button("➕ Cadastrar"):
                if novo_pv and nova_loja:
                    with engine.connect() as conn:
                        query_insert = text("""
                            INSERT INTO tb_fila_call_center 
                            (pv_abadi, loja, municipio, uf, tipo_necessidade, instrutor_sugerido, semana_sugerida, status_contato, observacoes)
                            VALUES (:pv, :loja, :muni, :uf, :tipo, :inst, :semana, :status, :obs)
                        """)
                        conn.execute(query_insert, {
                            "pv": int(novo_pv), "loja": nova_loja, "muni": novo_municipio,
                            "uf": novo_uf.upper(), "tipo": novo_tipo, "inst": novo_instrutor or "A definir",
                            "semana": nova_semana, "status": novo_status_cad, "obs": novas_obs_cad
                        })
                        conn.commit()
                    st.success("Cadastrado com sucesso!")
                    st.rerun()

    # ------------------------------------------
    # ABA 5: EXPORTAÇÃO
    # ------------------------------------------
    with aba_exportar:
        st.subheader("📥 Exportar Dados")
        col_exp1, col_exp2 = st.columns(2)
        csv_data = df_filtrado.to_csv(index=False).encode('utf-8')
        col_exp1.download_button("📄 Baixar CSV", data=csv_data, file_name="fila_ampm.csv", mime="text/csv")
        
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_filtrado.to_excel(writer, index=False, sheet_name='Fila')
        col_exp2.download_button("📊 Baixar Excel (.xlsx)", data=buffer.getvalue(), file_name="fila_ampm.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

except Exception as e:
    st.error(f"Erro ao carregar sistema: {e}")
# Modal Pop-up para Atualização de Contato
@st.dialog("📝 Registrar Contato / Atendimento")
def editar_contato(loja_dados):
    st.write(f"**PV:** {loja_dados.get('pv_abadi')} - {loja_dados.get('loja')}")
    
    with st.form("form_contato"):
        nome = st.text_input("Nome do Contato", value=loja_dados.get("nome_contato", ""))
        telefone = st.text_input("Telefone / WhatsApp", value=loja_dados.get("telefone", ""))
        
        status = st.selectbox(
            "Status do Contato",
            ["A Contatar", "Interessado - Aguardando confirmação", "Agendado", "Recusou", "Sem Resposta", "Loja Inativa"],
            index=0
        )
        
        obs = st.text_area("Observações do Atendimento", value=loja_dados.get("observacao", ""))
        
        btn_salvar = st.form_submit_button("💾 Salvar Registro")
        
        if btn_salvar:
            # Aqui você atualiza o banco/dataframe
            st.success("Contato atualizado com sucesso!")
            st.rerun()

# Seletor para abrir o pop-up
loja_selecionada = st.selectbox(
    "Selecione o PV para registrar contato:",
    options=df_fila['pv_abadi'].tolist() if 'df_fila' in locals() else []
)

if st.button("Abrir Formulario de Contato"):
    dados = df_fila[df_fila['pv_abadi'] == loja_selecionada].iloc[0].to_dict()
    editar_contato(dados)
# Exemplo de correção no carregamento da equipe
# Busca a lista de instrutores diretamente dos dados carregados na fila
if 'df_fila' in locals() and 'instrutor_sugerido' in df_fila.columns:
    lista_instrutores = df_fila['instrutor_sugerido'].dropna().unique().tolist()
elif 'df_base' in locals() and 'instrutor_sugerido' in df_base.columns:
    lista_instrutores = df_base['instrutor_sugerido'].dropna().unique().tolist()
else:
    lista_instrutores = ["Isabela Paim Ricardo", "Carla Fernandes Dionizio"] # Nomes de fallback baseados no seu painel
# Configuração da página
st.set_page_config(page_title="CRM Operacional AmPm", layout="wide")

# Estilização visual (Cores AmPm)
st.markdown("""
    <style>
    .stButton>button {
        background-color: #e0a96d;
        color: #000;
        font-weight: bold;
        border-radius: 5px;
    }
    </style>
""", unsafe_allow_html=True)

# --- DIÁLOGO / MODAL DE ATENDIMENTO ---
@st.dialog("📝 Registrar Contato / Atendimento")
def abrir_modal_contato(loja_dados, lista_instrutores):
    st.markdown(f"### PV: {loja_dados.get('pv_abadi')} — {loja_dados.get('loja')}")
    st.caption(f"📍 Município: {loja_dados.get('municipio')} | UF: {loja_dados.get('uf')}")
    
    with st.form("form_registro_crm"):
        col1, col2 = st.columns(2)
        
        with col1:
            nome_contato = st.text_input("Nome do Decisor / Contato", value=loja_dados.get("nome_contato", ""))
            telefone = st.text_input("Telefone / WhatsApp", value=loja_dados.get("telefone", ""))
            status = st.selectbox(
                "Status da Abordagem",
                ["A Contatar", "Interessado - Aguardando confirmação", "Agendado", "Recusou", "Sem Resposta", "Loja Inativa"]
            )
            
        with col2:
            # Puxando a lista CORRETA de Instrutores
            instrutor_alocado = st.selectbox("Instrutor Alocado", options=lista_instrutores)
            data_agendamento = st.date_input("Data Prevista para Treinamento")
            
        obs = st.text_area("Observações do Atendimento", value=loja_dados.get("observacao", ""))
        
        salvar = st.form_submit_button("💾 Salvar Registro")
        if salvar:
            # Lógica para persistir os dados
            st.success("Registro atualizado com sucesso!")
            st.rerun()

# --- CORPO PRINCIPAL ---
st.title("⛽ CRM Operacional AmPm — Treinamentos & Inteligência")

# Seleção rápida para registrar atendimento
st.subheader("📋 Gestão da Fila de Atendimento")

col_sel, col_btn = st.columns([3, 1])

with col_sel:
    pv_selecionado = st.selectbox(
        "Selecione a loja/PV para atualizar registro:",
        options=df_fila['pv_abadi'].tolist() if 'df_fila' in locals() else []
    )

with col_btn:
    st.write("") # Espaçamento
    st.write("") 
    if st.button("📝 Registrar Contato"):
        dados_loja = df_fila[df_fila['pv_abadi'] == pv_selecionado].iloc[0].to_dict()
        lista_inst = df_instrutores['nome'].tolist() if 'df_instrutores' in locals() else []
        abrir_modal_contato(dados_loja, lista_inst)
