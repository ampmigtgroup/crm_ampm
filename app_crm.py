from __future__ import annotations

import io
import os
from pathlib import Path
import pandas as pd
import pydeck as pdk
import streamlit as st

# ==========================================
# CONSTANTES E CONFIGURAÇÕES GERAIS
# ==========================================
CAMINHO_ARQUIVO = Path("Base_Unificada_AmPm.xlsx")

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
        st.session_state.autenticado = True
    return st.session_state.autenticado


def inicializar_estado() -> None:
    """Inicializa as variáveis de sessão no Streamlit."""
    if "df_base" not in st.session_state:
        st.session_state.df_base = carregar_base_inicial()
    if "fila_trabalho" not in st.session_state:
        st.session_state.fila_trabalho = st.session_state.df_base


@st.cache_data
def carregar_base_inicial() -> pd.DataFrame:
    """Tenta carregar o arquivo local padrão do projeto."""
    if CAMINHO_ARQUIVO.exists():
        try:
            return pd.read_excel(CAMINHO_ARQUIVO)
        except Exception:
            pass

    # Estrutura inicial padrão de fallback
    return pd.DataFrame({
        "ID": [1, 2, 3],
        "Cliente": ["Unidade Centro RJ", "Unidade Ipanema", "Unidade Niterói"],
        "Contato": ["(21) 99999-0001", "(21) 99999-0002", "(21) 99999-0003"],
        "Status": ["Em Atendimento", "Pendente", "Concluído"],
        "Lat": [-22.9068, -22.9836, -22.8833],
        "Lon": [-43.1729, -43.2044, -43.1036],
        "Data_Criacao": ["2026-08-01", "2026-08-05", "2026-08-10"],
    })


def construir_base_unificada() -> pd.DataFrame:
    """Retorna a base unificada ativa na memória."""
    if "df_base" in st.session_state and not st.session_state.df_base.empty:
        return st.session_state.df_base
    return carregar_base_inicial()


def salvar_fila_no_disco(df: pd.DataFrame) -> None:
    """Atualiza os dados na sessão e tenta salvar no disco local."""
    st.session_state.df_base = df
    st.session_state.fila_trabalho = df
    try:
        df.to_excel(CAMINHO_ARQUIVO, index=False)
        st.toast("Banco de dados atualizado com sucesso!", icon="✅")
    except Exception as e:
        st.toast(f"Dados mantidos em memória (disco em modo leitura): {e}", icon="⚠️")


# ==========================================
# RENDERIZAÇÃO DO MAPA (PYDECK - SEM LIMITE DE COTA)
# ==========================================
def renderizar_mapa(df: pd.DataFrame) -> None:
    """Exibe o mapa interativo via Pydeck utilizando blocos abertos do CartoDB."""
    st.subheader("Visão Geográfica de Pontos")

    if df.empty or "Lat" not in df.columns or "Lon" not in df.columns:
        st.info("Nenhum dado geográfico disponível para exibição no mapa.")
        return

    df_mapa = df.dropna(subset=["Lat", "Lon"]).copy()

    if df_mapa.empty:
        st.warning("Nenhum ponto com coordenadas válidas para exibir no mapa.")
        return

    lat_centro = float(df_mapa["Lat"].mean())
    lon_centro = float(df_mapa["Lon"].mean())

    view_state = pdk.ViewState(
        latitude=lat_centro,
        longitude=lon_centro,
        zoom=10,
        pitch=0,
    )

    layer = pdk.Layer(
        "ScatterplotLayer",
        data=df_mapa,
        get_position=["Lon", "Lat"],
        get_color=[230, 57, 70, 200],
        get_radius=300,
        pickable=True,
    )

    st.pydeck_chart(
        pdk.Deck(
            layers=[layer],
            initial_view_state=view_state,
            map_style="https://basemaps.cartocdn.com/gl/positron-gl-style/style.json",
            tooltip={
                "html": "<b>Cliente:</b> {Cliente}<br/><b>Status:</b> {Status}<br/><b>Contato:</b> {Contato}",
                "style": {"backgroundColor": "#1E293B", "color": "white"},
            },
        )
    )


# ==========================================
# FLUXO PRINCIPAL DA APLICAÇÃO
# ==========================================
def main() -> None:
    st.set_page_config(page_title="CRM Operacional AmPm", layout="wide")

    if not exigir_login():
        st.stop()

    inicializar_estado()

    st.title("CRM Operacional AmPm")

    # BARRA LATERAL - UPLOAD E FILTROS
    st.sidebar.header("Atualizar Banco de Dados")

    arquivo_enviado = st.sidebar.file_uploader(
        "Envie a nova planilha (.xlsx ou .csv):",
        type=["xlsx", "csv"],
        key="uploader_base",
    )

    # Processamento direto do arquivo na memória
    if arquivo_enviado is not None:
        try:
            if arquivo_enviado.name.endswith(".csv"):
                df_carregado = pd.read_csv(arquivo_enviado)
            else:
                df_carregado = pd.read_excel(arquivo_enviado)

            if not df_carregado.empty:
                st.session_state.df_base = df_carregado
                st.session_state.fila_trabalho = df_carregado
                st.sidebar.success("Base de dados carregada com sucesso!")
        except Exception as err:
            st.sidebar.error(f"Erro ao ler arquivo enviado: {err}")

    st.sidebar.divider()
    st.sidebar.header("Módulos & Filtros")
    entidade_selecionada = st.sidebar.selectbox("Selecione a Entidade", ENTIDADES)

    df_atual = construir_base_unificada()

    # Módulo do Dashboard / Visualização
    if df_atual.empty:
        st.info("Nenhum dado carregado ainda. Envie o arquivo Base_Unificada_AmPm.xlsx na barra lateral.")
    else:
        col_tabela, col_mapa = st.columns([1, 1])

        with col_tabela:
            st.subheader("Fila de Trabalho / Registros")
            st.dataframe(df_atual, use_container_width=True)

            if st.button("Salvar Alterações na Sessão"):
                salvar_fila_no_disco(df_atual)

        with col_mapa:
            renderizar_mapa(df_atual)


if __name__ == "__main__":
    main()
