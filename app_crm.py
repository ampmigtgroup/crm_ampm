import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, text
from io import BytesIO

st.set_page_config(page_title="CRM Operacional AmPm", layout="wide")

DATABASE_URL = "postgresql://postgres.nptazzfvwhhmotfrvgdj:Lssj.ampm%40%23@aws-0-sa-east-1.pooler.supabase.com:6543/postgres"

@st.cache_resource
def get_connection():
    return create_engine(DATABASE_URL)

engine = get_connection()

st.title("⛽ CRM Operacional - Treinamentos & Agendamentos AmPm")

# Consulta Base no Supabase
query = """
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
    df_fila = pd.read_sql(query, engine)

    # ==========================================
    # 1. FILTROS AVANÇADOS NA BARRA LATERAL
    # ==========================================
    st.sidebar.header("🔍 Filtros da Fila")

    # Busca por texto (PV ou Nome)
    busca_pv = st.sidebar.text_input("Buscar por PV Abadi ou Nome:")

    # Filtro de Status
    lista_status = ["Todos"] + list(df_fila['status_contato'].unique())
    status_filtro = st.sidebar.selectbox("Status do Contato:", lista_status)

    # Filtro de Instrutor
    lista_instrutores = ["Todos"] + list(df_fila['instrutor_sugerido'].dropna().unique())
    instrutor_filtro = st.sidebar.selectbox("Instrutor Sugerido:", lista_instrutores)

    # Filtro de UF
    lista_ufs = ["Todos"] + sorted([uf for uf in df_fila['uf'].unique() if uf])
    uf_filtro = st.sidebar.selectbox("UF (Estado):", lista_ufs)

    # Aplicação dos Filtros
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

    # ==========================================
    # 2. DASHBOARD / INDICADORES (KPIs)
    # ==========================================
    total_registros = len(df_fila)
    a_contatar = len(df_fila[df_fila['status_contato'] == 'A Contatar'])
    contatados = len(df_fila[df_fila['status_contato'] == 'Contatado'])
    confirmados = len(df_fila[df_fila['status_contato'] == 'Confirmado'])
    concluidos = len(df_fila[df_fila['status_contato'] == 'Concluído'])

    col_kpi1, col_kpi2, col_kpi3, col_kpi4, col_kpi5 = st.columns(5)
    col_kpi1.metric("Total na Fila", total_registros)
    col_kpi2.metric("A Contatar", a_contatar, delta=f"{(a_contatar/total_registros*100):.1f}%" if total_registros > 0 else "0%")
    col_kpi3.metric("Contatados", contatados)
    col_kpi4.metric("Confirmados", confirmados)
    col_kpi5.metric("Concluídos", concluidos)

    st.divider()

    # Criação de Abas para Organizar o CRM
    aba_fila, aba_cadastro, aba_exportar = st.tabs(["📋 Fila & Atualização", "➕ Novo Cadastro Manual", "📥 Exportar Relatórios"])

    # ------------------------------------------
    # ABA 1: FILA DE ATENDIMENTOS E EDIÇÃO
    # ------------------------------------------
    with aba_fila:
        st.subheader(f"Fila Prioritária ({len(df_filtrado)} registros filtrados)")
        st.dataframe(df_filtrado, use_container_width=True)

        if not df_filtrado.empty:
            st.divider()
            st.subheader("📝 Atualizar Agendamento / Status")

            lista_atendimentos = df_filtrado['id_atendimento'].tolist()
            id_selecionado = st.selectbox("Selecione o ID do Atendimento para editar:", lista_atendimentos)
            
            item = df_filtrado[df_filtrado['id_atendimento'] == id_selecionado].iloc[0]
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.write(f"**PV Abadi:** {item['pv_abadi']}")
                st.write(f"**Loja/Posto:** {item['loja']}")
                st.write(f"**Localidade:** {item['municipio']} / {item['uf']}")
            with col2:
                st.write(f"**Necessidade:** {item['tipo_necessidade']}")
                st.write(f"**Instrutor Sugerido:** {item['instrutor_sugerido']}")
                st.write(f"**Semana Sugerida:** {item['semana_sugerida']}")
            with col3:
                status_atuais = ["A Contatar", "Contatado", "Confirmado", "Recusado/Adiado", "Concluído"]
                status_item = item['status_contato'] if item['status_contato'] in status_atuais else "A Contatar"
                idx_status = status_atuais.index(status_item)
                novo_status = st.selectbox("Novo Status de Contato:", status_atuais, index=idx_status)
            
            obs_atual = item['observacoes'] if pd.notnull(item['observacoes']) else ""
            novas_obs = st.text_area("Observações do Atendimento:", value=obs_atual)
            
            if st.button("💾 Salvar Alterações no Banco"):
                with engine.connect() as conn:
                    query_update = text("""
                        UPDATE tb_fila_call_center 
                        SET status_contato = :status, observacoes = :obs
                        WHERE id_atendimento = :id
                    """)
                    conn.execute(query_update, {"status": novo_status, "obs": novas_obs, "id": int(id_selecionado)})
                    conn.commit()
                st.success(f"Status do atendimento #{id_selecionado} atualizado com sucesso!")
                st.rerun()

    # ------------------------------------------
    # ABA 2: MÓDULO DE CADASTRO MANUAL
    # ------------------------------------------
    with aba_cadastro:
        st.subheader("➕ Adicionar Novo Atendimento à Fila")
        
        with st.form("form_novo_atendimento", clear_on_submit=True):
            col_c1, col_c2 = st.columns(2)
            with col_c1:
                novo_pv = st.number_input("PV Abadi:", min_value=1, step=1)
                nova_loja = st.text_input("Nome da Loja/Posto:")
                novo_municipio = st.text_input("Município:")
                novo_uf = st.text_input("UF (Estado - ex: SP):", max_chars=2)
            with col_c2:
                novo_tipo = st.selectbox("Tipo de Necessidade:", ["Retreinamento", "Inauguração", "Suporte Emergencial", "Treinamento Específico"])
                novo_instrutor = st.text_input("Instrutor Sugerido:")
                nova_semana = st.text_input("Semana Sugerida (ex: 15/08/2026 a 19/08/2026):")
                novo_status_cad = st.selectbox("Status Inicial:", ["A Contatar", "Contatado", "Confirmado"])
            
            novas_obs_cad = st.text_area("Observações Iniciais:")
            btn_cadastrar = st.form_submit_button("➕ Cadastrar Atendimento")
            
            if btn_cadastrar:
                if novo_pv and nova_loja:
                    with engine.connect() as conn:
                        query_insert = text("""
                            INSERT INTO tb_fila_call_center 
                            (pv_abadi, loja, municipio, uf, tipo_necessidade, instrutor_sugerido, semana_sugerida, status_contato, observacoes)
                            VALUES (:pv, :loja, :muni, :uf, :tipo, :inst, :semana, :status, :obs)
                        """)
                        conn.execute(query_insert, {
                            "pv": int(novo_pv),
                            "loja": nova_loja,
                            "muni": novo_municipio,
                            "uf": novo_uf.upper(),
                            "tipo": novo_tipo,
                            "inst": novo_instrutor if novo_instrutor else "A definir",
                            "semana": nova_semana,
                            "status": novo_status_cad,
                            "obs": novas_obs_cad
                        })
                        conn.commit()
                    st.success(f"Nova solicitação para o PV {novo_pv} cadastrada com sucesso!")
                    st.rerun()
                else:
                    st.warning("Preencha ao menos o PV Abadi e o Nome da Loja.")

    # ------------------------------------------
    # ABA 3: EXPORTAÇÃO DE RELATÓRIOS
    # ------------------------------------------
    with aba_exportar:
        st.subheader("📥 Baixar Relatórios")
        st.write("Exporte a base filtrada atualmente ou a base completa para análise externa.")

        col_exp1, col_exp2 = st.columns(2)

        # Exportar CSV
        csv_data = df_filtrado.to_csv(index=False).encode('utf-8')
        col_exp1.download_button(
            label="📄 Baixar Relatório (CSV)",
            data=csv_data,
            file_name="fila_call_center_ampm.csv",
            mime="text/csv"
        )

        # Exportar Excel (.xlsx)
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_filtrado.to_excel(writer, index=False, sheet_name='Fila_Filtrada')
        
        col_exp2.download_button(
            label="📊 Baixar Relatório (Excel .xlsx)",
            data=buffer.getvalue(),
            file_name="fila_call_center_ampm.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

except Exception as e:
    st.error(f"Erro ao carregar dados: {e}")
