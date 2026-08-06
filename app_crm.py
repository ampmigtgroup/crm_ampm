import os
import pandas as pd
import streamlit as st

# ==========================================
# BASE DE DADOS OFICIAL DE INSTRUTORES (FIXA)
# ==========================================
INSTRUTORES_OFICIAIS = [
    {
        "nome": "Betânia Cayret Pregnolato",
        "id_auvo": "Betânia Pregnolato",
        "email": "betania.ampm@igtgroup-ext.com.br",
        "telefone": "(11) 99351-1743",
        "cidade": "São Paulo - SP",
        "status": "Ativo",
    },
    {
        "nome": "Bruno Souza",
        "id_auvo": "Bruno Ferreira",
        "email": "bruno.ampm@igtgroup-ext.com.br",
        "telefone": "(11) 99596-6401",
        "cidade": "Barueri - SP",
        "status": "Ativo",
    },
    {
        "nome": "Carla Fernandes Dionizio",
        "id_auvo": "Carla Dionizio",
        "email": "carla.ampm@igtgroup-ext.com.br",
        "telefone": "(11) 98744-7398",
        "cidade": "São Paulo - SP",
        "status": "Ativo",
    },
    {
        "nome": "Isabela Paim Ricardo",
        "id_auvo": "Isabela Paim",
        "email": "isabela.ampm@igtgroup-ext.com.br",
        "telefone": "(21) 99390-7088",
        "cidade": "Rio de Janeiro - RJ",
        "status": "Ativo",
    },
    {
        "nome": "Leonardo da Silva Azevedo",
        "id_auvo": "Leonardo Azevedo",
        "email": "leonardo.ampm@igtgroup-ext.com.br",
        "telefone": "(53) 8112-3384",
        "cidade": "Rio Grande - RS",
        "status": "Ativo",
    },
    {
        "nome": "Roberta de Cássia Martareli",
        "id_auvo": "Roberta de Cassia",
        "email": "roberta.ampm@igtgroup-ext.com.br",
        "telefone": "(11) 95337-4199",
        "cidade": "São Paulo - SP",
        "status": "Ativo",
    },
    {
        "nome": "Lucas Silva dos Santos",
        "id_auvo": "Lucas Silva",
        "email": "lucas.ampm@igtgroup-ext.com.br",
        "telefone": "(11) 97707-5141",
        "cidade": "Monte Carlo - MG",
        "status": "Ativo",
    },
    # Instrutores Inativos (Histórico)
    {
        "nome": "André Luiz de Medeiros",
        "id_auvo": "Andre Luiz",
        "email": "andre.ampm@igtgroup-ext.com.br",
        "telefone": "(81) 9386-9032",
        "cidade": "N/A",
        "status": "Saiu",
    },
    {
        "nome": "Diego Henrique de Souza",
        "id_auvo": "Diego Henrique",
        "email": "diego.ampm@igtgroup-ext.com.br",
        "telefone": "(41) 8706-3610",
        "cidade": "N/A",
        "status": "Saiu",
    },
    {
        "nome": "Juliano Rodrigues Amoretti",
        "id_auvo": "Juliano Amoretti",
        "email": "juliano.ampm@igtgroup-ext.com.br",
        "telefone": "(48) 9138-0057",
        "cidade": "N/A",
        "status": "Saiu",
    },
    {
        "nome": "Marcela Lourenço",
        "id_auvo": "Marcela Lourenço",
        "email": "marcela.ampm@igtgroup-ext.com.br",
        "telefone": "(43) 9159-2334",
        "cidade": "N/A",
        "status": "Saiu",
    },
    {
        "nome": "Simone Franceschi Barreto",
        "id_auvo": "Simone Franceschi",
        "email": "simone.ampm@igtgroup-ext.com.br",
        "telefone": "(47) 9252-8844",
        "cidade": "N/A",
        "status": "Saiu",
    },
    {
        "nome": "Tabajara Grecca",
        "id_auvo": "Tabajara Grecca",
        "email": "tabajara.ampm@igtgroup-ext.com.br",
        "telefone": "(19) 98161-1163",
        "cidade": "N/A",
        "status": "Saiu",
    },
]


@st.cache_data
def carregar_instrutores():
    """Carrega dados da planilha 2026.AMPM - instrutores.xlsx ou utiliza a lista fixa."""
    arquivo_excel = "2026.AMPM - instrutores.xlsx"

    if os.path.exists(arquivo_excel):
        try:
            df = pd.read_excel(arquivo_excel)
            df_clean = df.iloc[1:].copy()
            df_clean.columns = [
                "ID_AUVO",
                "ID_CAJU",
                "ID_UBER",
                "ID_UNICO",
                "NOME",
                "TELEFONE",
                "E_MAIL",
                "ENDERECO",
                "DATA_NASC",
                "CPF",
                "STATUS",
                "CIDADE_1",
                "INSTRUTOR",
                "CIDADE_2",
            ]

            lista_excel = []
            for _, row in df_clean.iterrows():
                lista_excel.append(
                    {
                        "nome": str(row["NOME"]).strip(),
                        "id_auvo": str(row["ID_AUVO"]).strip(),
                        "email": str(row["E_MAIL"]).strip(),
                        "telefone": str(row["TELEFONE"]).strip(),
                        "cidade": (
                            str(row["CIDADE_1"]).strip()
                            if pd.notna(row["CIDADE_1"])
                            else "N/A"
                        ),
                        "status": (
                            str(row["STATUS"]).strip()
                            if pd.notna(row["STATUS"])
                            else "Ativo"
                        ),
                    }
                )
            return pd.DataFrame(lista_excel)
        except Exception as e:
            st.warning(f"Erro ao ler arquivo Excel: {e}. Usando base oficial fixa.")

    return pd.DataFrame(INSTRUTORES_OFICIAIS)


# Configuração do Streamlit
st.set_page_config(
    page_title="Gestão de Instrutores - AMPM", page_icon="📍", layout="wide"
)

st.title("📍 Módulo de Gestão de Instrutores — AMPM")

# Carregar dados
df_instrutores = carregar_instrutores()

# Filtro Lateral
st.sidebar.header("Filtros")
status_filtro = st.sidebar.radio(
    "Filtrar por Status:", ["Apenas Ativos (7)", "Todos (13)", "Apenas Desligados (6)"]
)

if status_filtro == "Apenas Ativos (7)":
    df_exibicao = df_instrutores[df_instrutores["status"] == "Ativo"]
elif status_filtro == "Apenas Desligados (6)":
    df_exibicao = df_instrutores[df_instrutores["status"] == "Saiu"]
else:
    df_exibicao = df_instrutores

# Métricas Principais
col1, col2, col3 = st.columns(3)
col1.metric("Instrutores Ativos", len(df_instrutores[df_instrutores["status"] == "Ativo"]))
col2.metric("Instrutores Desligados", len(df_instrutores[df_instrutores["status"] == "Saiu"]))
col3.metric("Total no Cadastrados", len(df_instrutores))

st.markdown("---")

# Tabela Interativa
st.subheader("Lista Oficial de Instrutores")
st.dataframe(
    df_exibicao[["nome", "id_auvo", "status", "telefone", "email", "cidade"]],
    column_config={
        "nome": "Nome Completo",
        "id_auvo": "ID Auvo",
        "status": "Status",
        "telefone": "Telefone",
        "email": "E-mail Oficial",
        "cidade": "Cidade / UF",
    },
    use_container_width=True,
    hide_index=True,
)

# Seleção Individual para Detalhes
st.markdown("---")
st.subheader("Consultar Ficha Individual")
instrutor_selecionado = st.selectbox(
    "Selecione um instrutor para visualizar os detalhes:",
    df_exibicao["nome"].tolist(),
)

if instrutor_selecionado:
    dados = df_exibicao[df_exibicao["nome"] == instrutor_selecionado].iloc[0]
    c1, c2 = st.columns(2)
    with c1:
        st.write(f"**Nome:** {dados['nome']}")
        st.write(f"**ID Auvo:** {dados['id_auvo']}")
        st.write(f"**Status:** {dados['status']}")
    with c2:
        st.write(f"**E-mail:** {dados['email']}")
        st.write(f"**Telefone:** {dados['telefone']}")
        st.write(f"**Cidade / UF:** {dados['cidade']}")
