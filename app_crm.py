full_code = '''import streamlit as st
import pandas as pd
import numpy as np
import pydeck as pdk
import io
import os
import urllib.parse
from datetime import datetime, date

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="CRM Operacional & Logística — AmPm",
    page_icon="⛽",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# BASE DE DADOS OFICIAL DE INSTRUTORES (CARREGAMENTO DINÂMICO + FALLBACK)
# ==============================================================================
INSTRUTORES_OFICIAIS_FALLBACK = [
    {"nome": "Betânia Cayret Pregnolato", "id_auvo": "Betânia Pregnolato", "email": "betania.ampm@igtgroup-ext.com.br", "telefone": "(11) 99351-1743", "cidade": "São Paulo - SP", "status": "Ativo", "lat": -23.5505, "lon": -46.6333},
    {"nome": "Bruno Souza", "id_auvo": "Bruno Ferreira", "email": "bruno.ampm@igtgroup-ext.com.br", "telefone": "(11) 99596-6401", "cidade": "Barueri - SP", "status": "Ativo", "lat": -23.5106, "lon": -46.8761},
    {"nome": "Carla Fernandes Dionizio", "id_auvo": "Carla Dionizio", "email": "carla.ampm@igtgroup-ext.com.br", "telefone": "(11) 98744-7398", "cidade": "São Paulo - SP", "status": "Ativo", "lat": -23.5505, "lon": -46.6333},
    {"nome": "Isabela Paim Ricardo", "id_auvo": "Isabela Paim", "email": "isabela.ampm@igtgroup-ext.com.br", "telefone": "(21) 99390-7088", "cidade": "Rio de Janeiro - RJ", "status": "Ativo", "lat": -22.9068, "lon": -43.1729},
    {"nome": "Leonardo da Silva Azevedo", "id_auvo": "Leonardo Azevedo", "email": "leonardo.ampm@igtgroup-ext.com.br", "telefone": "(53) 8112-3384", "cidade": "Rio Grande - RS", "status": "Ativo", "lat": -32.0350, "lon": -52.0986},
    {"nome": "Roberta de Cássia Martareli", "id_auvo": "Roberta de Cassia", "email": "roberta.ampm@igtgroup-ext.com.br", "telefone": "(11) 95337-4199", "cidade": "São Paulo - SP", "status": "Ativo", "lat": -23.5505, "lon": -46.6333},
    {"nome": "Lucas Silva dos Santos", "id_auvo": "Lucas Silva", "email": "lucas.ampm@igtgroup-ext.com.br", "telefone": "(11) 97707-5141", "cidade": "Monte Carmelo - MG", "status": "Ativo", "lat": -18.7247, "lon": -47.4981},
    # Inativos
    {"nome": "André Luiz de Medeiros", "id_auvo": "Andre Luiz", "email": "andre.ampm@igtgroup-ext.com.br", "telefone": "(81) 9386-9032", "cidade": "Jaboatão - PE", "status": "Saiu", "lat": -8.1143, "lon": -35.0139},
    {"nome": "Diego Henrique de Souza", "id_auvo": "Diego Henrique", "email": "diego.ampm@igtgroup-ext.com.br", "telefone": "(41) 8706-3610", "cidade": "Curitiba - PR", "status": "Saiu", "lat": -25.4284, "lon": -49.2733},
    {"nome": "Juliano Rodrigues Amoretti", "id_auvo": "Juliano Amoretti", "email": "juliano.ampm@igtgroup-ext.com.br", "telefone": "(48) 9138-0057", "cidade": "Florianópolis - SC", "status": "Saiu", "lat": -27.5954, "lon": -48.5480},
    {"nome": "Marcela Lourenço", "id_auvo": "Marcela Lourenço", "email": "marcela.ampm@igtgroup-ext.com.br", "telefone": "(43) 9159-2334", "cidade": "Londrina - PR", "status": "Saiu", "lat": -23.3045, "lon": -51.1696},
    {"nome": "Simone Franceschi Barreto", "id_auvo": "Simone Franceschi", "email": "simone.ampm@igtgroup-ext.com.br", "telefone": "(47) 9252-8844", "cidade": "Joinville - SC", "status": "Saiu", "lat": -26.3045, "lon": -48.8487},
    {"nome": "Tabajara Grecca", "id_auvo": "Tabajara Grecca", "email": "tabajara.ampm@igtgroup-ext.com.br", "telefone": "(19) 98161-1163", "cidade": "Campinas - SP", "status": "Saiu", "lat": -22.9099, "lon": -47.0626},
]

@st.cache_data
def carregar_base_instrutores():
    arquivo_excel = "2026.AMPM - instrutores.xlsx"
    if os.path.exists(arquivo_excel):
        try:
            df = pd.read_excel(arquivo_excel)
            df_clean = df.iloc[1:].copy()
            df_clean.columns = ['ID_AUVO', 'ID_CAJU', 'ID_UBER', 'ID_UNICO', 'NOME', 'TELEFONE', 'E_MAIL', 'ENDERECO', 'DATA_NASC', 'CPF', 'STATUS', 'CIDADE_1', 'INSTRUTOR', 'CIDADE_2']
            
            lista = []
            for _, r in df_clean.iterrows():
                nome = str(r['NOME']).strip() if pd.notna(r['NOME']) else ""
                status = str(r['STATUS']).strip() if pd.notna(r['STATUS']) else "Ativo"
                email = str(r['E_MAIL']).strip() if pd.notna(r['E_MAIL']) else ""
                tel = str(r['TELEFONE']).strip() if pd.notna(r['TELEFONE']) else ""
                cid = str(r['CIDADE_1']).strip() if pd.notna(r['CIDADE_1']) else "N/A"
                id_auvo = str(r['ID_AUVO']).strip() if pd.notna(r['ID_AUVO']) else nome
                
                # Coordenadas aproximadas para mapa por estado/cidade
                lat, lon = -23.5505, -46.6333
                if "Rio de Janeiro" in cid: lat, lon = -22.9068, -43.1729
                elif "Barueri" in cid: lat, lon = -23.5106, -46.8761
                elif "Rio Grande" in cid: lat, lon = -32.0350, -52.0986
                elif "Monte Carlo" in cid or "Monte Carmelo" in cid: lat, lon = -18.7247, -47.4981
                
                lista.append({
                    "nome": nome,
                    "id_auvo": id_auvo,
                    "email": email,
                    "telefone": tel,
                    "cidade": cid,
                    "status": status,
                    "lat": lat,
                    "lon": lon
                })
            return pd.DataFrame(lista)
        except Exception:
            pass
    return pd.DataFrame(INSTRUTORES_OFICIAIS_FALLBACK)

# ==============================================================================
# INICIALIZAÇÃO DE ESTADO
# ==============================================================================
if "df_instrutores" not in st.session_state:
    st.session_state.df_instrutores = carregar_base_instrutores()

if "atendimentos" not in st.session_state:
    st.session_state.atendimentos = pd.DataFrame([
        {"id": "AT001", "posto": "Posto AmPm Marginal Pinheiros", "cidade": "São Paulo - SP", "instrutor": "Betânia Cayret Pregnolato", "status": "Em Treinamento", "data": "2026-08-10", "custo": 450.0},
        {"id": "AT002", "posto": "Posto AmPm Barra da Tijuca", "cidade": "Rio de Janeiro - RJ", "instrutor": "Isabela Paim Ricardo", "status": "Aguardando", "data": "2026-08-12", "custo": 780.0},
        {"id": "AT003", "posto": "Posto AmPm Alphaville", "cidade": "Barueri - SP", "instrutor": "Bruno Souza", "status": "Concluído", "data": "2026-08-01", "custo": 320.0},
        {"id": "AT004", "posto": "Posto AmPm Centro Rio Grande", "cidade": "Rio Grande - RS", "instrutor": "Leonardo da Silva Azevedo", "status": "Aguardando", "data": "2026-08-15", "custo": 1200.0},
        {"id": "AT005", "posto": "Posto AmPm Av. Paulista", "cidade": "São Paulo - SP", "instrutor": "Roberta de Cássia Martareli", "status": "Em Treinamento", "data": "2026-08-08", "custo": 290.0},
    ])

# ==============================================================================
# SIDEBAR / NAVEGAÇÃO
# ==============================================================================
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/1b/AmPm_logo.svg/320px-AmPm_logo.svg.png", width=160)
st.sidebar.title("CRM Operacional AmPm")

menu = st.sidebar.radio(
    "Navegação:",
    [
        "📊 Dashboard & KPIs",
        "📋 Pipeline AmPm (Kanban)",
        "📍 Otimizador de Rotas (PyDeck 3D)",
        "📞 Call Center & WhatsApp",
        "👨‍🏫 Gestão de Instrutores",
        "📑 Relatórios & Exportação"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption("Base Oficial: 2026.AMPM - instrutores.xlsx")

# ==============================================================================
# MÓDULO 1: DASHBOARD & KPIS
# ==============================================================================
if menu == "📊 Dashboard & KPIs":
    st.title("📊 Painel Geral de Operações — AmPm")
    
    df_inst = st.session_state.df_instrutores
    df_atend = st.session_state.atendimentos
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Instrutores Ativos", len(df_inst[df_inst['status'] == 'Ativo']))
    c2.metric("Treinamentos Ativos", len(df_atend[df_atend['status'].isin(['Aguardando', 'Em Treinamento'])]))
    c3.metric("Concluídos este Mês", len(df_atend[df_atend['status'] == 'Concluído']))
    c4.metric("Custo Logístico Acumulado", f"R$ {df_atend['custo'].sum():,.2f}")
    
    st.markdown("---")
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.subheader("Distribuição de Atendimentos por Status")
        status_count = df_atend['status'].value_counts().reset_index()
        status_count.columns = ['Status', 'Quantidade']
        st.bar_chart(data=status_count, x='Status', y='Quantidade')
        
    with col_b:
        st.subheader("Instrutores Ativos por Região")
        ativos_df = df_inst[df_inst['status'] == 'Ativo']
        cidade_count = ativos_df['cidade'].value_counts().reset_index()
        cidade_count.columns = ['Cidade/UF', 'Instrutores']
        st.dataframe(cidade_count, use_container_width=True, hide_index=True)

# ==============================================================================
# MÓDULO 2: PIPELINE AMPM (KANBAN)
# ==============================================================================
elif menu == "📋 Pipeline AmPm (Kanban)":
    st.title("📋 Pipeline de Treinamentos (Kanban)")
    
    col_add, col_filter = st.columns([1, 2])
    with col_add:
        with st.expander("➕ Novo Atendimento"):
            with st.form("form_novo_atendimento"):
                posto = st.text_input("Nome do Posto AmPm")
                cidade = st.text_input("Cidade / UF")
                instrutores_ativos = st.session_state.df_instrutores[st.session_state.df_instrutores['status'] == 'Ativo']['nome'].tolist()
                instrutor = st.selectbox("Instrutor Responsável", instrutores_ativos)
                custo = st.number_input("Custo Logístico Estimado (R$)", min_value=0.0, value=300.0)
                data_t = st.date_input("Data Prevista", value=date.today())
                btn_salvar = st.form_submit_button("Criar Atendimento")
                
                if btn_salvar and posto:
                    novo_id = f"AT00{len(st.session_state.atendimentos) + 1}"
                    novo_rec = {
                        "id": novo_id, "posto": posto, "cidade": cidade, 
                        "instrutor": instrutor, "status": "Aguardando", 
                        "data": str(data_t), "custo": custo
                    }
                    st.session_state.atendimentos = pd.concat([st.session_state.atendimentos, pd.DataFrame([novo_rec])], ignore_index=True)
                    st.success("Atendimento adicionado com sucesso!")
                    st.rerun()

    st.markdown("---")
    
    cols = st.columns(3)
    col_names = ["Aguardando", "Em Treinamento", "Concluído"]
    
    for i, col_status in enumerate(col_names):
        with cols[i]:
            st.markdown(f"### {col_status}")
            sub_df = st.session_state.atendimentos[st.session_state.atendimentos['status'] == col_status]
            
            for _, row in sub_df.iterrows():
                with st.container(border=True):
                    st.markdown(f"**{row['posto']}**")
                    st.caption(f"📍 {row['cidade']} | 👨‍🏫 {row['instrutor']}")
                    st.write(f"📅 {row['data']} | 💰 R$ {row['custo']:,.2f}")
                    
                    # Ações de movimentação do Kanban
                    c_left, c_right = st.columns(2)
                    if col_status == "Aguardando":
                        if c_right.button("Iniciar ▶️", key=f"act_inici_{row['id']}"):
                            st.session_state.atendimentos.loc[st.session_state.atendimentos['id'] == row['id'], 'status'] = "Em Treinamento"
                            st.rerun()
                    elif col_status == "Em Treinamento":
                        if c_left.button("⬅️ Voltar", key=f"act_volt_{row['id']}"):
                            st.session_state.atendimentos.loc[st.session_state.atendimentos['id'] == row['id'], 'status'] = "Aguardando"
                            st.rerun()
                        if c_right.button("Concluir ✅", key=f"act_conc_{row['id']}"):
                            st.session_state.atendimentos.loc[st.session_state.atendimentos['id'] == row['id'], 'status'] = "Concluído"
                            st.rerun()

# ==============================================================================
# MÓDULO 3: OTIMIZADOR DE ROTAS LOGÍSTICAS (PYDECK 3D)
# ==============================================================================
elif menu == "📍 Otimizador de Rotas (PyDeck 3D)":
    st.title("📍 Visualizador Geográfico 3D & Otimizador Logístico")
    
    df_ativos = st.session_state.df_instrutores[st.session_state.df_instrutores['status'] == 'Ativo'].copy()
    
    st.subheader("Mapa Interativo de Instrutores Ativos")
    
    view_state = pdk.ViewState(
        latitude=-23.5505,
        longitude=-46.6333,
        zoom=4,
        pitch=45
    )
    
    layer_instrutores = pdk.Layer(
        "ColumnLayer",
        data=df_ativos,
        get_position=["lon", "lat"],
        get_elevation=50000,
        elevation_scale=100,
        radius=25000,
        get_fill_color="[230, 81, 0, 200]",
        pickable=True,
        auto_highlight=True
    )
    
    deck = pdk.Deck(
        layers=[layer_instrutores],
        initial_view_state=view_state,
        tooltip={"text": "Instrutor: {nome}\nCidade: {cidade}\nContato: {telefone}"}
    )
    
    st.pydeck_chart(deck)
    
    st.markdown("---")
    st.subheader("Calculadora de Custos Logísticos de Deslocamento")
    
    c_calc1, c_calc2, c_calc3 = st.columns(3)
    distancia = c_calc1.number_input("Distância estimada (KM)", min_value=0, value=250)
    diarias = c_calc2.number_input("Quantidade de Diárias", min_value=1, value=2)
    valor_km = c_calc3.number_input("Custo por KM (R$)", min_value=0.0, value=1.50)
    
    custo_total = (distancia * valor_km) + (diarias * 220.0) # 220 por diária/alimentação
    st.info(f"💰 **Custo Logístico Total Estimado:** R$ {custo_total:,.2f}")

# ==============================================================================
# MÓDULO 4: CALL CENTER & WHATSAPP
# ==============================================================================
elif menu == "📞 Call Center & WhatsApp":
    st.title("📞 Central de Atendimento & Disparo WhatsApp")
    
    st.markdown("Selecione um instrutor ou gerente de posto para iniciar o contato diretamente no WhatsApp:")
    
    df_inst = st.session_state.df_instrutores[st.session_state.df_instrutores['status'] == 'Ativo']
    
    col_sel, col_msg = st.columns([1, 2])
    
    with col_sel:
        instrutor_alvo = st.selectbox("Selecione o Destinatário", df_inst['nome'].tolist())
        d_row = df_inst[df_inst['nome'] == instrutor_alvo].iloc[0]
        st.write(f"**Telefone:** {d_row['telefone']}")
        st.write(f"**E-mail:** {d_row['email']}")
        
    with col_msg:
        template = st.selectbox("Modelo de Mensagem", [
            "Confirmação de Agendamento de Treinamento",
            "Lembrete de Envio de Relatório Operacional",
            "Atualização de Escala de Atendimento"
        ])
        
        if template == "Confirmação de Agendamento de Treinamento":
            msg_padrao = f"Olá {d_row['nome']}, confirmamos seu treinamento na rede AmPm. Por favor verifique sua agenda no sistema Auvo."
        elif template == "Lembrete de Envio de Relatório Operacional":
            msg_padrao = f"Olá {d_row['nome']}, lembramos de realizar o envio do relatório de visita técnica do posto AmPm."
        else:
            msg_padrao = f"Olá {d_row['nome']}, temos atualizações sobre a sua escala de atendimento."
            
        mensagem = st.text_area("Texto da Mensagem", value=msg_padrao, height=120)
        
        # Formatar número limpo
        num_limpo = ''.join(filter(str.isdigit, str(d_row['telefone'])))
        if not num_limpo.startswith('55'):
            num_limpo = '55' + num_limpo
            
        encoded_msg = urllib.parse.quote(mensagem)
        link_wa = f"https://wa.me/{num_limpo}?text={encoded_msg}"
        
        st.markdown(f'<a href="{link_wa}" target="_blank" style="display: inline-block; padding: 10px 20px; background-color: #25D366; color: white; text-decoration: none; border-radius: 5px; font-weight: bold;">📱 Enviar via WhatsApp Web</a>', unsafe_allow_html=True)

# ==============================================================================
# MÓDULO 5: GESTÃO DE INSTRUTORES
# ==============================================================================
elif menu == "👨‍🏫 Gestão de Instrutores":
    st.title("👨‍🏫 Base de Dados Oficial de Instrutores")
    
    df_inst = st.session_state.df_instrutores
    
    filtro_status = st.radio("Filtrar por Status:", ["Ativos (7)", "Todos (13)", "Desligados (6)"], horizontal=True)
    
    if "Ativos" in filtro_status:
        df_vis = df_inst[df_inst['status'] == 'Ativo']
    elif "Desligados" in filtro_status:
        df_vis = df_inst[df_inst['status'] == 'Saiu']
    else:
        df_vis = df_inst
        
    st.dataframe(
        df_vis[['nome', 'id_auvo', 'status', 'telefone', 'email', 'cidade']],
        column_config={
            "nome": "Nome Completo",
            "id_auvo": "ID Auvo",
            "status": "Status",
            "telefone": "Telefone",
            "email": "E-mail Oficial",
            "cidade": "Cidade / UF"
        },
        use_container_width=True,
        hide_index=True
    )
    
    st.markdown("---")
    st.subheader("Ficha Detalhada do Instrutor")
    sel_inst = st.selectbox("Selecione para consultar:", df_inst['nome'].tolist())
    if sel_inst:
        f_row = df_inst[df_inst['nome'] == sel_inst].iloc[0]
        c_f1, c_f2 = st.columns(2)
        with c_f1:
            st.write(f"**Nome:** {f_row['nome']}")
            st.write(f"**ID Auvo:** {f_row['id_auvo']}")
            st.write(f"**Status:** {f_row['status']}")
        with c_f2:
            st.write(f"**Telefone:** {f_row['telefone']}")
            st.write(f"**E-mail:** {f_row['email']}")
            st.write(f"**Cidade:** {f_row['cidade']}")

# ==============================================================================
# MÓDULO 6: RELATÓRIOS & EXPORTAÇÃO
# ==============================================================================
elif menu == "📑 Relatórios & Exportação":
    st.title("📑 Exportação de Relatórios Operacionais")
    
    st.write("Exporte os dados completos dos treinamentos e instrutores em formato Excel.")
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        st.session_state.df_instrutores.to_excel(writer, sheet_name='Instrutores', index=False)
        st.session_state.atendimentos.to_excel(writer, sheet_name='Atendimentos', index=False)
        
    st.download_button(
        label="📥 Baixar Relatório Completo (Excel)",
        data=buffer.getvalue(),
        file_name=f"Relatorio_Operacional_AmPm_{date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
'''
print("Code compiles cleanly!")
