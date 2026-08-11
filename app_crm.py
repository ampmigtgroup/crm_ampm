from __future__ import annotations

import os
from pathlib import Path
import pandas as pd
import pydeck as pdk
import streamlit as st

# ==========================================
# CONSTANTES E CONFIGURAÇÕES GERAIS
# ==========================================
CAMINHO_ARQUIVO = Path("dados_crm.xlsx")

COLUNAS_FILA = [
    "ID",
    "Cliente",
    "Contato",
    "Status",
    "Lat",
    "Lon",
    "Data_Criacao",
]

ENTIDADES = [
    "Lojas AmPm",
    "Postos de Combustível",
    "Franquias",
    "Parceiros Logísticos",
]


# ==========================================
# FUNÇÕES DE ESTADO E AUTENTICAÇÃO
# ==========================================
def exigir_login() -> bool:
    """Gerencia a autenticação e acesso de usuários."""
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = True  # Padrão temporário / integração
    return st.session_state.autenticado


def inicializar_estado() -> None:
    """Inicializa as variáveis de sessão necessárias do Streamlit."""
    if "fila_trabalho" not in st.session_state:
        st.session_state.fila_trabalho = construir_base_unificada()


def construir_base_unificada() -> pd.DataFrame:
    """Carrega a base de dados do disco ou gera uma estrutura padrão caso não exista."""
    if CAMINHO_ARQUIVO.exists():
        try:
            return pd.read_excel(CAMINHO_ARQUIVO)
        except Exception:
            pass

    # Base Padrão de Exemplo (Rio de Janeiro e Região)
    dados_iniciais = {
        "ID": [1, 2, 3],
        "Cliente": ["Unidade Centro RJ", "Unidade Ipanema", "Unidade Niterói"],
        "Contato": ["(21) 99999-0001", "(21) 99999-0002", "(21) 99999-0003"],
        "Status": ["Em Atendimento", "Pendente", "Concluído"],
        "Lat": [-22.9068, -22.9836, -22.8833],
        "Lon": [-43.1729, -43.2044, -43.1036],
        "Data_Criacao": ["2026-08-01", "2026-08-05", "2026-08-10"],
    }
    return pd.DataFrame(dados_iniciais)


def salvar_fila_no_disco(df: pd.DataFrame) -> None:
    """Salva a fila de trabalho atualizada em um arquivo Excel."""
    try:
        df.to_excel(CAMINHO_ARQUIVO, index=False)
        st.success("Dados salvos com sucesso!")
    except Exception as e:
        st.error(f"Erro ao salvar arquivo: {e}")


# ==========================================
# RENDERIZAÇÃO DO MAPA (PYDECK - SEM LIMITE DE COTA)
# ==========================================
def renderizar_mapa(df: pd.DataFrame) -> None:
    """Exibe o mapa interativo via Pydeck utilizando blocos abertos do CartoDB."""
    st.subheader("Mapa de Pontos de Referência")

    # Filtra registros com coordenadas válidas
    df_mapa = df.dropna(subset=["Lat", "Lon"]).copy()

    if df_mapa.empty:
        st.warning("Nenhum ponto com coordenadas válidas para exibir no mapa.")
        return

    # Ponto central para inicialização da câmera
    lat_centro = float(df_mapa["Lat"].mean())
    lon_centro = float(df_mapa["Lon"].mean())

    view_state = pdk.ViewState(
        latitude=lat_centro,
        longitude=lon_centro,
        zoom=10,
        pitch=0,
    )

    # Camada de marcadores (pontos vermelhos)
    layer = pdk.Layer(
        "ScatterplotLayer",
        data=df_mapa,
        get_position=["Lon", "Lat"],
        get_color=[230, 57, 70, 200],
        get_radius=300,
        pickable=True,
    )

    # Renderiza o mapa com estilo vetorial do CartoDB (livre de chave de API)
    st.pydeck_chart(
        pdk.Deck(
            layers=[layer],
            initial_view_state=view_state,
            map_style="https://basemaps.cartocdn.com/gl/positron-gl-style/style.json",
            tooltip={
                "html": "<b>Cliente:</b> {Cliente}<br/><b>Status:</b> {Status}<br/><b>Contato:</b> {Contato}",
                "style": {"backgroundColor": "steelblue", "color": "white"},
            },
        )
    )


# ==========================================
# FLUXO PRINCIPAL DA APLICAÇÃO
# ==========================================
def main() -> None:
    st.set_page_config(page_title="CRM AmPm", layout="wide")

    if not exigir_login():
        st.stop()

    inicializar_estado()

    st.title("Sistema CRM AmPm")

    # Sidebar para filtros e configurações
    st.sidebar.header("Navegação & Filtros")
    entidade_selecionada = st.sidebar.selectbox("Selecione a Entidade", ENTIDADES)
    st.sidebar.markdown(f"**Entidade Ativa:** {entidade_selecionada}")

    df_atual = st.session_state.fila_trabalho

    # Layout em colunas: Tabela à esquerda, Mapa à direita
    col_tabela, col_mapa = st.columns([1, 1])

    with col_tabela:
        st.subheader("Fila de Trabalho")
        st.dataframe(df_atual, use_container_width=True)

        if st.button("Salvar Alterações no Disco"):
            salvar_fila_no_disco(df_atual)

    with col_mapa:
        renderizar_mapa(df_atual)


if __name__ == "__main__":
    main()
