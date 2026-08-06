import streamlit as st
import pandas as pd
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

CAMINHO_ARQUIVO = "Base_Unificada_AmPm.xlsx"
CAMINHO_BACKUP = "Base_Unificada_AmPm.backup.xlsx"

# Colunas "editáveis" que vivem na aba Fila_CallCenter.
# Consolidamos aqui TODOS os campos que o Call Center e o Pipeline escrevem,
# para que nada seja perdido entre recarregamentos e para que salvar no disco
# grave exatamente o que deveria ser gravado (e não a base inteira mesclada).
COLUNAS_FILA = [
    "PV_Abadi", "Tipo_Necessidade", "Data_Ultimo_Treinamento",
    "Dias_desde_Ultimo_Treinamento", "Instrutor_Sugerido", "Semana_Sugerida",
    "Telefone_Contato", "Status_Contato", "Data_do_Contato", "Observacoes",
    "Nome_Contato", "Qtd_Funcionarios", "Material_Em_Loja", "Data_Agendada",
]

# --- ESTILIZAÇÃO CSS CUSTOMIZADA (DESIGN SYSTEM AMPM PREMIUM) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500&display=swap');

    /* ============ DESIGN TOKENS ============ */
    :root {
        --ampm-orange: #E27B00;
        --ampm-orange-light: #FF9800;
        --ampm-red: #D32F2F;
        --bg-app: #0E1116;
        --bg-surface: #161A22;
        --bg-surface-alt: #1B2029;
        --bg-surface-raised: #1F2530;
        --border-subtle: #262C38;
        --border-strong: #333B49;
        --text-primary: #F2F4F8;
        --text-secondary: #9AA4B4;
        --text-tertiary: #6B7688;
        --success: #22C55E;
        --success-bg: rgba(34, 197, 94, 0.12);
        --warning: #F5A524;
        --warning-bg: rgba(245, 165, 36, 0.12);
        --danger: #EF4444;
        --danger-bg: rgba(239, 68, 68, 0.12);
        --info: #3B9EFF;
        --info-bg: rgba(59, 158, 255, 0.12);
        --neutral-bg: rgba(154, 164, 180, 0.12);
        --radius-sm: 8px;
        --radius-md: 12px;
        --radius-lg: 16px;
        --shadow-sm: 0 1px 3px rgba(0,0,0,0.24);
        --shadow-md: 0 6px 16px rgba(0,0,0,0.28);
        --shadow-lg: 0 12px 32px rgba(0,0,0,0.34);
        --shadow-glow: 0 8px 24px rgba(226, 123, 0, 0.22);
    }

    /* ============ BASE ============ */
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        letter-spacing: -0.01em;
    }
    .stApp {
        background: radial-gradient(circle at 12% 0%, #171C25 0%, var(--bg-app) 42%);
    }
    code, .mono {
        font-family: 'JetBrains Mono', 'Courier New', monospace !important;
    }
    h1, h2, h3, h4, h5 { letter-spacing: -0.02em; }
    hr { border-color: var(--border-subtle) !important; }

    /* Compacta um pouco o padding padrão do container principal */
    .block-container { padding-top: 1.6rem; padding-bottom: 3rem; }

    /* ============ TOPBAR PREMIUM ============ */
    .main-header {
        background: linear-gradient(120deg, #B85E00 0%, var(--ampm-orange) 42%, var(--ampm-orange-light) 78%, var(--ampm-red) 130%);
        padding: 30px 34px;
        border-radius: var(--radius-lg);
        color: white;
        margin-bottom: 28px;
        box-shadow: var(--shadow-glow);
        position: relative;
        overflow: hidden;
        border: 1px solid rgba(255,255,255,0.08);
    }
    .main-header::after {
        content: "";
        position: absolute;
        top: -40%; right: -8%;
        width: 280px; height: 280px;
        background: radial-gradient(circle, rgba(255,255,255,0.16) 0%, rgba(255,255,255,0) 70%);
        border-radius: 50%;
        pointer-events: none;
    }
    .main-header-top {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        position: relative;
        z-index: 1;
    }
    .main-header h1 {
        color: #FFFFFF !important;
        margin: 0 0 6px 0;
        font-weight: 800;
        font-size: 2.05rem;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .main-header p {
        margin: 0;
        font-size: 0.98rem;
        color: rgba(255,255,255,0.92);
        font-weight: 500;
    }
    .header-status-chip {
        background: rgba(0,0,0,0.22);
        border: 1px solid rgba(255,255,255,0.22);
        color: #fff;
        padding: 7px 14px;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 700;
        letter-spacing: 0.3px;
        white-space: nowrap;
        backdrop-filter: blur(6px);
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .pulse-dot {
        width: 7px; height: 7px;
        border-radius: 50%;
        background: #7CFFA0;
        box-shadow: 0 0 0 0 rgba(124,255,160,0.7);
        animation: pulse-anim 2s infinite;
        display: inline-block;
    }
    @keyframes pulse-anim {
        0%   { box-shadow: 0 0 0 0 rgba(124,255,160,0.55); }
        70%  { box-shadow: 0 0 0 7px rgba(124,255,160,0); }
        100% { box-shadow: 0 0 0 0 rgba(124,255,160,0); }
    }

    /* ============ SECTION HEADER ============ */
    .section-header {
        display: flex;
        align-items: center;
        gap: 12px;
        margin: 4px 0 18px 0;
    }
    .section-header .icon-badge {
        width: 38px; height: 38px;
        min-width: 38px;
        border-radius: var(--radius-sm);
        background: linear-gradient(135deg, var(--ampm-orange) 0%, var(--ampm-orange-light) 100%);
        display: flex; align-items: center; justify-content: center;
        font-size: 1.1rem;
        box-shadow: var(--shadow-sm);
    }
    .section-header .titles h3 {
        margin: 0; font-size: 1.15rem; font-weight: 700; color: var(--text-primary);
    }
    .section-header .titles span {
        font-size: 0.82rem; color: var(--text-secondary);
    }

    /* ============ KPI CARDS ============ */
    .kpi-card {
        background: linear-gradient(160deg, var(--bg-surface-raised) 0%, var(--bg-surface) 100%);
        border-radius: var(--radius-md);
        padding: 20px 22px;
        border: 1px solid var(--border-subtle);
        box-shadow: var(--shadow-sm);
        transition: transform 0.18s ease, box-shadow 0.18s ease, border-color 0.18s ease;
        position: relative;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: var(--shadow-md);
        border-color: var(--border-strong);
    }
    .kpi-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .kpi-icon-circle {
        width: 34px; height: 34px;
        border-radius: 10px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1rem;
    }
    .kpi-title {
        font-size: 0.74rem;
        color: var(--text-secondary);
        text-transform: uppercase;
        font-weight: 700;
        letter-spacing: 0.9px;
    }
    .kpi-value {
        font-size: 2rem;
        font-weight: 800;
        color: var(--text-primary);
        margin-top: 10px;
        line-height: 1;
        font-family: 'JetBrains Mono', monospace;
        letter-spacing: -0.01em;
    }
    .kpi-footer {
        margin-top: 10px;
        font-size: 0.76rem;
        color: var(--text-tertiary);
    }

    /* ============ PIPELINE (KANBAN) ============ */
    .ampm-column {
        background: var(--bg-surface);
        border-radius: var(--radius-md);
        padding: 16px;
        border: 1px solid var(--border-subtle);
        min-height: 480px;
        box-shadow: var(--shadow-sm);
    }
    .ampm-title {
        font-size: 0.86rem;
        font-weight: 700;
        margin-bottom: 14px;
        padding-bottom: 10px;
        border-bottom: 1px solid var(--border-subtle);
        display: flex;
        justify-content: space-between;
        align-items: center;
        color: var(--text-primary);
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .ampm-title .pill-count {
        background: var(--bg-surface-raised);
        border: 1px solid var(--border-strong);
        padding: 2px 10px;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 700;
        font-family: 'JetBrains Mono', monospace;
    }
    .col-a-contatar   { border-top: 3px solid var(--text-tertiary); }
    .col-em-negociacao { border-top: 3px solid var(--warning); }
    .col-agendado      { border-top: 3px solid var(--info); }
    .col-treinamento-realizado { border-top: 3px solid var(--success); }
    .col-recusado      { border-top: 3px solid var(--danger); }

    /* ============ CARDS DE INFORMAÇÃO (PROCV / CALL CENTER) ============ */
    .procv-card {
        background: var(--bg-surface-alt);
        padding: 22px;
        border-radius: var(--radius-md);
        border: 1px solid var(--border-subtle);
        border-top: 3px solid var(--ampm-orange);
        box-shadow: var(--shadow-sm);
        margin-bottom: 16px;
    }
    .procv-card h4 {
        margin-top: 0;
        margin-bottom: 14px;
        color: var(--ampm-orange-light);
        font-size: 0.98rem;
        font-weight: 700;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .procv-card p {
        margin: 6px 0;
        font-size: 0.89rem;
        color: var(--text-primary);
        line-height: 1.5;
    }
    .procv-card p b { color: var(--text-secondary); font-weight: 600; }

    .top-instructor-card {
        background: var(--bg-surface-alt);
        padding: 20px;
        border-radius: var(--radius-md);
        border: 1px solid var(--border-subtle);
        border-left: 4px solid var(--success);
        margin-bottom: 14px;
        box-shadow: var(--shadow-sm);
        transition: transform 0.15s ease;
    }
    .top-instructor-card:hover { transform: translateY(-2px); }

    /* ============ TIMELINE ============ */
    .timeline-item {
        border-left: 3px solid var(--ampm-orange);
        padding: 4px 0 4px 16px;
        margin-bottom: 15px;
        position: relative;
    }
    .timeline-item::before {
        content: "";
        position: absolute;
        left: -7px; top: 8px;
        width: 11px; height: 11px;
        border-radius: 50%;
        background: var(--ampm-orange);
        border: 2px solid var(--bg-app);
    }

    /* ============ BADGES / STATUS PILLS ============ */
    .badge-info, .badge-pill {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        padding: 3px 10px;
        border-radius: 999px;
        font-weight: 700;
        font-size: 0.74rem;
        letter-spacing: 0.2px;
        border: 1px solid transparent;
    }
    .badge-info { background: var(--warning-bg); color: var(--ampm-orange-light); border-color: rgba(226,123,0,0.35); }
    .badge-neutral  { background: var(--neutral-bg); color: var(--text-secondary); border-color: var(--border-strong); }
    .badge-warning  { background: var(--warning-bg); color: var(--warning); border-color: rgba(245,165,36,0.35); }
    .badge-info-blue{ background: var(--info-bg); color: var(--info); border-color: rgba(59,158,255,0.35); }
    .badge-success  { background: var(--success-bg); color: var(--success); border-color: rgba(34,197,94,0.35); }
    .badge-danger   { background: var(--danger-bg); color: var(--danger); border-color: rgba(239,68,68,0.35); }

    /* ============ BOTÕES ============ */
    .stButton>button {
        background: linear-gradient(100deg, var(--ampm-orange) 0%, var(--ampm-orange-light) 100%);
        color: #FFFFFF !important;
        font-weight: 700;
        border: none;
        border-radius: var(--radius-sm);
        padding: 10px 22px;
        letter-spacing: 0.2px;
        transition: all 0.2s ease;
        box-shadow: var(--shadow-sm);
    }
    .stButton>button:hover {
        box-shadow: 0 6px 18px rgba(226, 123, 0, 0.45);
        transform: translateY(-1px);
    }
    .stButton>button:active { transform: translateY(0); }
    .stDownloadButton>button {
        border-radius: var(--radius-sm) !important;
        font-weight: 700 !important;
        border: 1px solid var(--border-strong) !important;
        transition: all 0.2s ease;
    }
    .stDownloadButton>button:hover {
        border-color: var(--ampm-orange) !important;
        color: var(--ampm-orange-light) !important;
    }

    /* ============ INPUTS / SELECTS / DATAFRAMES ============ */
    .stTextInput input, .stNumberInput input, .stTextArea textarea,
    .stDateInput input, div[data-baseweb="select"] > div {
        border-radius: var(--radius-sm) !important;
        border-color: var(--border-strong) !important;
    }
    .stDataFrame {
        border-radius: var(--radius-md);
        overflow: hidden;
        border: 1px solid var(--border-subtle);
    }
    div[data-testid="stExpander"] {
        border-radius: var(--radius-md) !important;
        border: 1px solid var(--border-subtle) !important;
        background: var(--bg-surface);
        overflow: hidden;
    }
    div[data-testid="stForm"] {
        border: 1px solid var(--border-subtle);
        border-radius: var(--radius-md);
        padding: 18px;
        background: var(--bg-surface-alt);
    }

    /* ============ SIDEBAR ============ */
    section[data-testid="stSidebar"] {
        border-right: 1px solid var(--border-subtle);
    }
    section[data-testid="stSidebar"] .stRadio label {
        font-size: 0.92rem;
    }
    .sidebar-brand {
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 2px;
    }
    .sidebar-brand .logo-chip {
        width: 34px; height: 34px;
        border-radius: 9px;
        background: linear-gradient(135deg, var(--ampm-orange) 0%, var(--ampm-red) 100%);
        display: flex; align-items: center; justify-content: center;
        font-size: 1rem;
        box-shadow: var(--shadow-sm);
    }
    .sidebar-metric {
        background: var(--bg-surface);
        border: 1px solid var(--border-subtle);
        border-radius: var(--radius-sm);
        padding: 10px 12px;
        font-size: 0.82rem;
        color: var(--text-secondary);
        margin-top: 4px;
    }
    .sidebar-metric b { color: var(--text-primary); }

    /* Scrollbar discreta */
    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: var(--border-strong); border-radius: 8px; }
    ::-webkit-scrollbar-thumb:hover { background: var(--text-tertiary); }
    </style>
""", unsafe_allow_html=True)


# ============ HELPERS DE APRESENTAÇÃO ============

def render_section_header(icone, titulo, subtitulo=""):
    """Renderiza um cabeçalho de seção padronizado (ícone + título + legenda)
    usado no topo de cada módulo, mantendo a identidade visual consistente."""
    st.markdown(f"""
        <div class="section-header">
            <div class="icon-badge">{icone}</div>
            <div class="titles">
                <h3>{titulo}</h3>
                {f'<span>{subtitulo}</span>' if subtitulo else ''}
            </div>
        </div>
    """, unsafe_allow_html=True)


STATUS_BADGE_MAP = {
    "A Contatar": ("badge-neutral", "⏳"),
    "Em Negociação": ("badge-warning", "🤝"),
    "Agendado": ("badge-info-blue", "📅"),
    "Treinamento Realizado": ("badge-success", "✅"),
    "Recusado": ("badge-danger", "🚫"),
}


def badge_status_html(status):
    """Retorna o HTML de um badge colorido de acordo com o status de
    atendimento, para uso consistente no Pipeline, PROCV e Call Center."""
    classe, emoji = STATUS_BADGE_MAP.get(str(status), ("badge-neutral", "•"))
    return f'<span class="badge-pill {classe}">{emoji} {status}</span>'


def status_css_class(status):
    """Classe CSS (kebab-case) usada para colorir a borda superior de cada
    coluna do Pipeline, de acordo com o status."""
    mapa = {
        "A Contatar": "col-a-contatar",
        "Em Negociação": "col-em-negociacao",
        "Agendado": "col-agendado",
        "Treinamento Realizado": "col-treinamento-realizado",
        "Recusado": "col-recusado",
    }
    return mapa.get(status, "")




# ==========================================
# CAMADA DE DADOS
# ==========================================

def parse_data_flexivel(valor):
    """Converte qualquer representação comum de data para um objeto date.
    Sempre tenta ISO (%Y-%m-%d) primeiro, pois é o formato em que passamos
    a gravar internamente; cai para BR (%d/%m/%Y) por compatibilidade com
    dados antigos. Retorna None se não for possível interpretar."""
    if valor is None:
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, pd.Timestamp):
        if pd.isna(valor):
            return None
        return valor.date()
    texto = str(valor).strip()
    if not texto or texto.lower() in ("nan", "nat", "none"):
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    return None


def _normalizar_nome(nome):
    """Normaliza um texto (nome de aba ou de coluna) para comparação
    tolerante a acentos, espaços/underscores extras e maiúsculas/minúsculas
    (ex.: 'PV_Abadi', 'pv abadi', ' PV ABADI ' → 'pv abadi')."""
    import unicodedata
    nome = str(nome).strip().lower()
    nome = unicodedata.normalize('NFKD', nome).encode('ascii', 'ignore').decode('ascii')
    nome = nome.replace('_', ' ').replace('-', ' ')
    nome = ' '.join(nome.split())
    return nome


# Para cada tipo de dado que o sistema precisa, definimos a coluna-chave
# (obrigatória para reconhecer a aba) e os "apelidos" aceitos para cada
# coluna. Isso permite que o app reconheça uma planilha pelo CONTEÚDO
# (quais colunas ela tem), não pelo nome da aba nem pela grafia exata —
# só entram no sistema as colunas que ele reconhece como compatíveis.
ENTIDADES = {
    "lojas": {
        "chave": "PV Abadi",
        "obrigatoria": True,
        "colunas": {
            "PV Abadi": ["pv abadi", "pv", "codigo pv", "cod pv", "id loja", "codigo loja", "numero pv", "n pv", "pv abadi rede"],
            "Razão Social": ["razao social", "nome loja", "loja", "unidade", "nome fantasia", "franquia", "nome da loja"],
            "Municipio": ["municipio", "cidade", "municipio loja"],
            "UF": ["uf", "estado", "uf loja"],
            "Endereço": ["endereco", "endereco completo", "logradouro", "endereço"],
            "Status Loja": ["status loja", "status", "situacao loja", "situacao"],
            "GF": ["gf", "gerente franquia", "gerente"],
            "CF": ["cf", "consultor", "consultor franquia", "consultor de franquia"],
            "Latitude": ["latitude", "lat"],
            "Longitude": ["longitude", "lon", "long", "lng"],
        },
    },
    "fila": {
        "chave": "PV_Abadi",
        "obrigatoria": False,
        "colunas": {
            "PV_Abadi": ["pv abadi", "pv", "codigo pv", "cod pv", "id loja", "codigo loja"],
            "Tipo_Necessidade": ["tipo necessidade", "necessidade", "tipo de necessidade", "tipo pendencia"],
            "Data_Ultimo_Treinamento": ["data ultimo treinamento", "ultimo treinamento", "data do ultimo treinamento"],
            "Dias_desde_Ultimo_Treinamento": ["dias desde ultimo treinamento", "dias sem treinamento", "dias desde treinamento"],
            "Instrutor_Sugerido": ["instrutor sugerido", "instrutor", "instrutor designado"],
            "Semana_Sugerida": ["semana sugerida", "semana"],
            "Telefone_Contato": ["telefone contato", "telefone", "contato telefone", "celular"],
            "Status_Contato": ["status contato", "status do contato", "status atendimento"],
            "Data_do_Contato": ["data do contato", "data contato", "ultima atualizacao"],
            "Observacoes": ["observacoes", "observacao", "obs", "comentarios"],
            "Nome_Contato": ["nome contato", "nome do contato", "responsavel loja", "responsavel"],
            "Qtd_Funcionarios": ["qtd funcionarios", "quantidade funcionarios", "qtd de funcionarios", "numero de funcionarios", "funcionarios"],
            "Material_Em_Loja": ["material em loja", "material na loja", "possui material", "apostilas"],
            "Data_Agendada": ["data agendada", "data do agendamento", "agendamento"],
        },
    },
    "inaug": {
        "chave": "PV ABADI",
        "obrigatoria": False,
        "colunas": {
            "PV ABADI": ["pv abadi", "pv", "codigo pv", "cod pv"],
            "Previsão Inauguração": ["previsao inauguracao", "data inauguracao", "previsao de inauguracao", "inauguracao"],
            "Pipeline": ["pipeline", "etapa pipeline", "fase"],
            "Consultor_Possivel_Instrutor": ["consultor possivel instrutor", "consultor instrutor", "possivel instrutor"],
        },
    },
    "instrutores": {
        "chave": "NOME_COMPLETO",
        "chave_numerica": False,
        "obrigatoria": False,
        "colunas": {
            "NOME_COMPLETO": ["nome completo", "nome", "instrutor", "nome do instrutor"],
            "STATUS": ["status", "situacao"],
            "TELEFONE": ["telefone", "celular", "contato telefone"],
            "EMAIL": ["email", "e mail"],
            "Cidade": ["cidade", "municipio"],
            "UF": ["uf", "estado"],
            "Latitude": ["latitude", "lat"],
            "Longitude": ["longitude", "lon", "long", "lng"],
        },
    },
    "rec": {
        "chave": "PV_ABADI",
        "obrigatoria": False,
        "colunas": {
            "PV_ABADI": ["pv abadi", "pv", "codigo pv", "cod pv"],
            "Razao_Social": ["razao social", "nome loja", "loja", "unidade"],
            "Municipio_Loja": ["municipio loja", "municipio", "cidade loja"],
            "UF_Loja": ["uf loja", "uf", "estado loja"],
            "Instrutor_Sugerido": ["instrutor sugerido", "instrutor"],
            "Cidade_Instrutor": ["cidade instrutor", "cidade do instrutor"],
            "UF_Instrutor": ["uf instrutor", "uf do instrutor"],
            "Ranking_Proximidade": ["ranking proximidade", "ranking", "posicao ranking"],
            "Distancia_km_linha_reta": ["distancia km linha reta", "distancia km", "distancia", "distancia linha reta"],
            "Dias_Treinamento_Necessarios": ["dias treinamento necessarios", "dias necessarios", "dias de treinamento"],
        },
    },
}

MIN_SCORE_CONFIANTE = 2  # chave + pelo menos 1 outra coluna reconhecida


def _construir_lookup(colunas_dict):
    """Constrói um dicionário {apelido_normalizado: nome_canonico} para uma
    entidade, incluindo o próprio nome canônico como apelido de si mesmo."""
    lookup = {}
    for canonico, apelidos in colunas_dict.items():
        for apelido in set(apelidos) | {canonico}:
            lookup[_normalizar_nome(apelido)] = canonico
    return lookup


def _mapear_colunas_compativeis(df, definicao_entidade):
    """Compara as colunas reais de um DataFrame com os apelidos aceitos de
    uma entidade. Retorna (rename_map, canonicas_encontradas, colunas_ignoradas).
    Só as colunas reconhecidas entram no rename_map — o resto é ignorado."""
    lookup = _construir_lookup(definicao_entidade["colunas"])
    rename_map = {}
    canonicas_encontradas = set()
    colunas_ignoradas = []
    for col in df.columns:
        chave_norm = _normalizar_nome(col)
        canonico = lookup.get(chave_norm)
        if canonico and canonico not in canonicas_encontradas:
            rename_map[col] = canonico
            canonicas_encontradas.add(canonico)
        else:
            colunas_ignoradas.append(str(col))
    return rename_map, canonicas_encontradas, colunas_ignoradas


def detectar_entidades_no_workbook(xls):
    """Examina TODAS as abas do arquivo e, para cada uma, mede a
    compatibilidade de colunas com cada tipo de dado esperado pelo sistema
    (lojas, fila de call center, inaugurações, instrutores, recomendação de
    deslocamento). Cada aba é atribuída ao tipo mais compatível, independente
    do nome da aba ou da grafia exata das colunas. Colunas não reconhecidas
    são descartadas silenciosamente (registradas no relatório, não usadas).
    Retorna (bases_dict, relatorio_list)."""
    dfs_brutos = {}
    candidatos = []  # (score, sheet_name, entidade, rename_map, canonicas)

    for sheet_name in xls.sheet_names:
        try:
            df_bruto = pd.read_excel(xls, sheet_name=sheet_name)
        except Exception:
            continue
        if df_bruto is None or len(df_bruto.columns) == 0:
            continue
        dfs_brutos[sheet_name] = df_bruto
        for entidade, definicao in ENTIDADES.items():
            rename_map, canonicas, _ = _mapear_colunas_compativeis(df_bruto, definicao)
            if definicao["chave"] not in canonicas:
                continue  # sem a coluna-chave, a aba não pode representar esta entidade
            score = len(canonicas)
            candidatos.append((score, sheet_name, entidade, rename_map, canonicas))

    candidatos.sort(key=lambda x: -x[0])

    entidade_atribuida = {}
    aba_usada = set()
    for score, sheet_name, entidade, rename_map, canonicas in candidatos:
        if score < MIN_SCORE_CONFIANTE:
            continue
        if entidade in entidade_atribuida or sheet_name in aba_usada:
            continue
        entidade_atribuida[entidade] = (sheet_name, rename_map, canonicas, score)
        aba_usada.add(sheet_name)

    # 'lojas' é obrigatória: se nada bateu com confiança, aceitamos o melhor
    # candidato disponível (mesmo que só tenha a coluna-chave), pois sem
    # rede de lojas o sistema não tem como funcionar.
    if "lojas" not in entidade_atribuida:
        for score, sheet_name, entidade, rename_map, canonicas in candidatos:
            if entidade == "lojas" and sheet_name not in aba_usada:
                entidade_atribuida["lojas"] = (sheet_name, rename_map, canonicas, score)
                aba_usada.add(sheet_name)
                break

    bases = {}
    relatorio = []
    for entidade, definicao in ENTIDADES.items():
        colunas_canonicas = list(definicao["colunas"].keys())
        if entidade in entidade_atribuida:
            sheet_name, rename_map, canonicas, score = entidade_atribuida[entidade]
            df_bruto = dfs_brutos[sheet_name]
            df_mapeado = df_bruto.rename(columns=rename_map)
            colunas_presentes = [c for c in colunas_canonicas if c in df_mapeado.columns]
            df_final = df_mapeado[colunas_presentes].copy()
            colunas_ignoradas = [c for c in df_bruto.columns if c not in rename_map]
            bases[entidade] = df_final
            relatorio.append({
                "entidade": entidade,
                "aba_origem": sheet_name,
                "confianca": "alta" if score >= MIN_SCORE_CONFIANTE else "baixa",
                "colunas_reconhecidas": colunas_presentes,
                "colunas_ignoradas": colunas_ignoradas,
            })
        else:
            bases[entidade] = pd.DataFrame(columns=colunas_canonicas)
            relatorio.append({
                "entidade": entidade,
                "aba_origem": None,
                "confianca": "n/a",
                "colunas_reconhecidas": [],
                "colunas_ignoradas": [],
            })

    return bases, relatorio


def _processar_excelfile(xls):
    """Detecta e monta as bases a partir de um pd.ExcelFile já aberto (seja
    de um caminho em disco ou de bytes em memória), aceitando qualquer nome
    de aba/coluna e usando apenas as informações compatíveis com o sistema.
    Lança ValueError se não for possível identificar nem a base de lojas."""
    bases, relatorio = detectar_entidades_no_workbook(xls)

    if bases["lojas"].empty or ENTIDADES["lojas"]["chave"] not in bases["lojas"].columns:
        abas_disponiveis = ", ".join(xls.sheet_names) if xls.sheet_names else "(nenhuma)"
        raise ValueError(
            "Não foi possível identificar, em nenhuma aba do arquivo, dados de rede de lojas "
            "(é necessário pelo menos uma coluna equivalente a 'PV Abadi'). "
            f"Abas presentes no arquivo enviado: {abas_disponiveis}."
        )

    for entidade, definicao in ENTIDADES.items():
        chave = definicao["chave"]
        df = bases[entidade]
        if definicao.get("chave_numerica", True) and chave in df.columns:
            df[chave] = pd.to_numeric(df[chave], errors="coerce")
        bases[entidade] = df

    for col in COLUNAS_FILA:
        if col not in bases["fila"].columns:
            bases["fila"][col] = pd.NA

    return bases, relatorio


def validar_bytes_excel(conteudo_bytes):
    """Valida (sem tocar no disco) se os bytes de um .xlsx enviado contêm ao
    menos dados de lojas reconhecíveis. Retorna (bases_dict, relatorio, None)
    em caso de sucesso, ou (None, None, mensagem_de_erro) em caso de falha —
    usado para checar o upload ANTES de sobrescrever o arquivo em produção."""
    try:
        xls = pd.ExcelFile(io.BytesIO(conteudo_bytes), engine='openpyxl')
        bases, relatorio = _processar_excelfile(xls)
        return bases, relatorio, None
    except Exception as e:
        return None, None, str(e)


@st.cache_data
def carregar_bases_do_disco(caminho, assinatura=None):
    """Lê e detecta as abas do Excel no disco. O parâmetro `assinatura`
    (ex.: data de modificação do arquivo) força o Streamlit a invalidar o
    cache quando o arquivo muda, mesmo com o mesmo caminho."""
    if not os.path.exists(caminho):
        return None, None
    xls = pd.ExcelFile(caminho, engine='openpyxl')
    return _processar_excelfile(xls)


def construir_base_unificada(df_lojas, df_fila, df_inaug):
    """Mescla as bases brutas em uma visão única para exibição.
    Esta função é chamada a cada execução (não é cacheada), pois os dados
    brutos podem ter sido editados em session_state. O custo é baixo para
    o volume de dados de um CRM de rede de lojas."""
    if df_lojas is None or df_lojas.empty:
        return pd.DataFrame()

    df_base = pd.merge(
        df_lojas,
        df_fila[[c for c in COLUNAS_FILA if c in df_fila.columns]],
        left_on='PV Abadi', right_on='PV_Abadi', how='left'
    )

    df_base = pd.merge(
        df_base,
        df_inaug[['PV ABADI', 'Previsão Inauguração', 'Pipeline', 'Consultor_Possivel_Instrutor']],
        left_on='PV Abadi', right_on='PV ABADI', how='left'
    )

    df_base['Status_Contato'] = df_base['Status_Contato'].fillna('A Contatar')
    df_base['Tipo_Necessidade'] = df_base['Tipo_Necessidade'].fillna('Rede Ativa (Sem Pendência)')
    df_base['Instrutor_Sugerido'] = df_base['Instrutor_Sugerido'].fillna('Pendente de Alocação')
    df_base['Nome_Contato'] = df_base['Nome_Contato'].fillna("")
    df_base['Qtd_Funcionarios'] = pd.to_numeric(df_base['Qtd_Funcionarios'], errors='coerce').fillna(0).astype(int)
    df_base['Material_Em_Loja'] = df_base['Material_Em_Loja'].fillna("Não Informado")
    # Data_Agendada e Observacoes permanecem como estão (podem ser NaN/"")

    return df_base


def salvar_fila_no_disco():
    """Grava a aba Fila_CallCenter (e SOMENTE ela) de volta no Excel,
    preservando as demais abas do arquivo. Corrige o bug original que
    sobrescrevia essa aba com a base inteira mesclada."""
    if not os.path.exists(CAMINHO_ARQUIVO):
        st.toast("⚠️ Arquivo local não encontrado — alterações mantidas apenas na sessão.", icon="⚠️")
        return
    try:
        df_fila = st.session_state['bases']['fila']
        colunas_saida = [c for c in COLUNAS_FILA if c in df_fila.columns]
        with pd.ExcelWriter(CAMINHO_ARQUIVO, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_fila[colunas_saida].to_excel(writer, sheet_name='Fila_CallCenter', index=False)
        st.toast("💾 Dados salvos no arquivo Excel local com sucesso!", icon="✅")
    except Exception as e:
        st.toast(f"⚠️ Erro ao salvar arquivo: {e}", icon="⚠️")


def atualizar_fila(pv_abadi, campos: dict):
    """Atualiza (ou cria, se ainda não existir) a linha correspondente ao PV
    na aba Fila_CallCenter em session_state, aplica os campos informados e
    persiste no disco. Usada tanto pelo Pipeline quanto pelo Call Center,
    garantindo que as duas telas nunca se dessincronizem."""
    df_fila = st.session_state['bases']['fila']
    pv_abadi = float(pv_abadi) if pd.notna(pv_abadi) else pv_abadi

    mask = df_fila['PV_Abadi'] == pv_abadi
    if not mask.any():
        nova_linha = {col: pd.NA for col in COLUNAS_FILA}
        nova_linha['PV_Abadi'] = pv_abadi
        df_fila = pd.concat([df_fila, pd.DataFrame([nova_linha])], ignore_index=True)
        mask = df_fila['PV_Abadi'] == pv_abadi

    for campo, valor in campos.items():
        if campo not in df_fila.columns:
            df_fila[campo] = pd.NA
        df_fila.loc[mask, campo] = valor

    st.session_state['bases']['fila'] = df_fila
    salvar_fila_no_disco()


def _bases_vazias():
    return {
        "lojas": pd.DataFrame(),
        "fila": pd.DataFrame(columns=COLUNAS_FILA),
        "inaug": pd.DataFrame(),
        "instrutores": pd.DataFrame(),
        "rec": pd.DataFrame(),
    }


def inicializar_estado():
    if 'bases' in st.session_state:
        return
    st.session_state.setdefault('erro_carga', None)
    st.session_state.setdefault('relatorio_importacao', None)
    if not os.path.exists(CAMINHO_ARQUIVO):
        st.session_state['bases'] = _bases_vazias()
        return
    try:
        assinatura = os.path.getmtime(CAMINHO_ARQUIVO)
        bases, relatorio = carregar_bases_do_disco(CAMINHO_ARQUIVO, assinatura)
        st.session_state['bases'] = bases if bases is not None else _bases_vazias()
        st.session_state['relatorio_importacao'] = relatorio
        st.session_state['erro_carga'] = None
    except Exception as e:
        # Nunca deixa uma falha de leitura derrubar o app inteiro: guarda o
        # erro para mostrar na tela e segue com bases vazias, permitindo que
        # o usuário faça um novo upload válido ou restaure o backup.
        st.session_state['erro_carga'] = str(e)
        st.session_state['bases'] = _bases_vazias()


inicializar_estado()

if st.session_state.get('erro_carga'):
    st.error(
        "⚠️ Não foi possível carregar `Base_Unificada_AmPm.xlsx`:\n\n"
        f"{st.session_state['erro_carga']}\n\n"
        "Envie um arquivo válido na barra lateral, ou restaure o último backup se houver um disponível."
    )

# --- SIDEBAR DE NAVEGAÇÃO, FILTROS GLOBAIS E UPLOAD ---
with st.sidebar:
    st.markdown("""
        <div class="sidebar-brand">
            <div class="logo-chip">⛽</div>
            <div>
                <div style="font-weight:800; font-size:1.05rem; line-height:1.1;">CRM AmPm</div>
                <div style="font-size:0.74rem; color:var(--text-tertiary);">Plataforma Integrada de Operações</div>
            </div>
        </div>
    """, unsafe_allow_html=True)
    st.divider()

    modulo = st.radio(
        "📌 **Módulos do Sistema:**",
        [
            "📊 Dashboard Executivo",
            "📋 Pipeline AmPm",
            "🔍 PROCV & Filtros Avançados",
            "📍 Calculadora & Otimizador de Custos",
            "📞 Call Center & Timeline WhatsApp",
            "👔 Equipe de Instrutores",
            "📂 Relatórios & Exportação"
        ]
    )

    st.divider()

    # UPLOAD DE BANCO DE DADOS NA SIDEBAR (.XLSX E .CSV)
    st.markdown("📥 **Atualizar Banco de Dados**")
    st.caption("Um .xlsx substitui o arquivo inteiro. Um .csv atualiza apenas UMA aba — escolha qual abaixo.")

    aba_destino_csv = st.selectbox(
        "Se o CSV não puder ser identificado automaticamente, use como destino:",
        ["Rede_de_Lojas", "Fila_CallCenter", "Previsao_Inauguracao", "Instrutores", "Recomendacao_Deslocamento"],
        help="O sistema tenta identificar sozinho pelo conteúdo das colunas. Esta opção só é usada como reserva, se a identificação automática falhar."
    )

    uploaded_file = st.file_uploader(
        "Envie a nova planilha (.xlsx ou .csv):",
        type=["xlsx", "csv"],
        help="O sistema aceita qualquer nome de aba/coluna — ele reconhece o conteúdo automaticamente e usa apenas o que for compatível."
    )

    def _exibir_relatorio_importacao(relatorio):
        nomes_entidade = {
            "lojas": "🏪 Rede de Lojas",
            "fila": "📞 Fila de Call Center",
            "inaug": "🚀 Previsão de Inauguração",
            "instrutores": "👔 Instrutores",
            "rec": "📍 Recomendação de Deslocamento",
        }
        with st.expander("🔎 Ver o que foi reconhecido no arquivo", expanded=False):
            for item in relatorio:
                nome = nomes_entidade.get(item["entidade"], item["entidade"])
                if item["aba_origem"]:
                    st.markdown(f"**{nome}** — encontrado na aba `{item['aba_origem']}`")
                    if item["colunas_reconhecidas"]:
                        st.caption("Colunas usadas: " + ", ".join(item["colunas_reconhecidas"]))
                    if item["colunas_ignoradas"]:
                        st.caption("Colunas ignoradas (não reconhecidas): " + ", ".join(item["colunas_ignoradas"]))
                else:
                    st.markdown(f"**{nome}** — não encontrado no arquivo (ficará vazio)")

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.xlsx'):
                conteudo = uploaded_file.getbuffer().tobytes()
                # Valida a estrutura ANTES de tocar no arquivo em produção —
                # isto evita sobrescrever uma base boa com um arquivo inválido
                # e deixar o app quebrado até um novo upload.
                bases_validadas, relatorio, erro = validar_bytes_excel(conteudo)
                if erro:
                    st.error(f"❌ Arquivo rejeitado — a base atual foi mantida intacta.\n\n{erro}")
                else:
                    if os.path.exists(CAMINHO_ARQUIVO):
                        with open(CAMINHO_ARQUIVO, "rb") as f_atual, open(CAMINHO_BACKUP, "wb") as f_bak:
                            f_bak.write(f_atual.read())
                    with open(CAMINHO_ARQUIVO, "wb") as f:
                        f.write(conteudo)
                    st.cache_data.clear()
                    st.session_state['bases'] = bases_validadas
                    st.session_state['relatorio_importacao'] = relatorio
                    st.session_state['erro_carga'] = None
                    # st.toast (e não st.success) porque sobrevive ao st.rerun() logo abaixo
                    st.toast("💾 Banco de dados atualizado! Backup do arquivo anterior guardado.", icon="✅")
                    st.rerun()
            elif uploaded_file.name.endswith('.csv'):
                df_csv = pd.read_csv(uploaded_file)
                chave_map = {
                    "Rede_de_Lojas": "lojas",
                    "Fila_CallCenter": "fila",
                    "Previsao_Inauguracao": "inaug",
                    "Instrutores": "instrutores",
                    "Recomendacao_Deslocamento": "rec",
                }

                # Tenta identificar automaticamente a que tipo de dado o CSV
                # corresponde, comparando suas colunas com as de cada entidade.
                melhor_entidade, melhor_score = None, 0
                for entidade, definicao in ENTIDADES.items():
                    _, canonicas, _ = _mapear_colunas_compativeis(df_csv, definicao)
                    if definicao["chave"] in canonicas and len(canonicas) > melhor_score:
                        melhor_entidade, melhor_score = entidade, len(canonicas)

                if melhor_entidade and melhor_score >= MIN_SCORE_CONFIANTE:
                    chave = melhor_entidade
                    origem = "identificação automática"
                else:
                    chave = chave_map[aba_destino_csv]
                    origem = "aba selecionada manualmente (não foi possível identificar automaticamente)"

                definicao = ENTIDADES[chave]
                rename_map, canonicas, ignoradas = _mapear_colunas_compativeis(df_csv, definicao)
                if definicao["chave"] not in canonicas:
                    st.error(
                        f"❌ O CSV não possui uma coluna equivalente a '{definicao['chave']}', "
                        "necessária para identificar os registros. Nada foi alterado."
                    )
                else:
                    df_mapeado = df_csv.rename(columns=rename_map)
                    colunas_presentes = [c for c in definicao["colunas"].keys() if c in df_mapeado.columns]
                    df_final = df_mapeado[colunas_presentes].copy()
                    df_final[definicao["chave"]] = pd.to_numeric(df_final[definicao["chave"]], errors='coerce')

                    if chave == "fila":
                        for col in COLUNAS_FILA:
                            if col not in df_final.columns:
                                df_final[col] = pd.NA

                    st.session_state['bases'][chave] = df_final
                    if chave == "fila":
                        salvar_fila_no_disco()
                    st.toast(f"✅ Dados de '{chave}' atualizados via {origem}!", icon="✅")
                    if ignoradas:
                        st.toast("ℹ️ Colunas ignoradas: " + ", ".join(ignoradas), icon="ℹ️")
                    st.rerun()
        except Exception as e:
            st.error(f"❌ Erro ao processar o arquivo: {e}")

    if os.path.exists(CAMINHO_BACKUP):
        if st.button("↩️ Restaurar último backup do Excel"):
            try:
                with open(CAMINHO_BACKUP, "rb") as f_bak, open(CAMINHO_ARQUIVO, "wb") as f_atual:
                    f_atual.write(f_bak.read())
                st.cache_data.clear()
                if 'bases' in st.session_state:
                    del st.session_state['bases']
                st.session_state['relatorio_importacao'] = None
                inicializar_estado()
                st.success("✅ Backup restaurado com sucesso!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Erro ao restaurar backup: {e}")

    if st.session_state.get('relatorio_importacao'):
        _exibir_relatorio_importacao(st.session_state['relatorio_importacao'])

    st.divider()

    # RECONSTRÓI A BASE UNIFICADA A PARTIR DAS BASES BRUTAS ATUAIS
    bases = st.session_state['bases']
    df_base_raw = construir_base_unificada(bases["lojas"], bases["fila"], bases["inaug"])
    df_instrutores = bases["instrutores"]
    df_rec_raw = bases["rec"]

    # Mescla de coordenadas para o módulo de otimização (feito aqui pois
    # depende de df_instrutores e df_lojas atuais)
    if not df_rec_raw.empty and not df_instrutores.empty:
        df_rec = pd.merge(
            df_rec_raw,
            df_instrutores[['NOME_COMPLETO', 'Latitude', 'Longitude']],
            left_on='Instrutor_Sugerido', right_on='NOME_COMPLETO', how='left'
        ).rename(columns={'Latitude': 'Lat_Instrutor', 'Longitude': 'Lon_Instrutor'})

        df_rec = pd.merge(
            df_rec,
            bases["lojas"][['PV Abadi', 'Latitude', 'Longitude']],
            left_on='PV_ABADI', right_on='PV Abadi', how='left'
        ).rename(columns={'Latitude': 'Lat_Loja', 'Longitude': 'Lon_Loja'})
    else:
        df_rec = df_rec_raw

    # FILTROS GLOBAIS
    st.markdown("🎯 **Filtros Globais**")
    uf_opcoes = ["Todas"] + sorted([str(x) for x in df_base_raw['UF'].dropna().unique()]) if 'UF' in df_base_raw.columns else ["Todas"]
    filtro_uf = st.selectbox("Filtrar Estado (UF):", uf_opcoes)
    cf_opcoes = ["Todos"] + sorted([str(x) for x in df_base_raw['CF'].dropna().unique()]) if 'CF' in df_base_raw.columns else ["Todos"]
    filtro_cf = st.selectbox("Filtrar Consultor (CF):", cf_opcoes)
    st.divider()
    st.markdown(f"""
        <div class="sidebar-metric">📶 Status do sistema: <b>Operacional</b> 🟢</div>
        <div class="sidebar-metric">🏪 Rede total: <b>{len(df_base_raw)} unidades</b></div>
    """, unsafe_allow_html=True)

# APLICAÇÃO DOS FILTROS GLOBAIS
df_base = df_base_raw.copy()
if filtro_uf != "Todas":
    df_base = df_base[df_base['UF'] == filtro_uf]
if filtro_cf != "Todos":
    df_base = df_base[df_base['CF'] == filtro_cf]

# Header Global
st.markdown(f"""
    <div class="main-header">
        <div class="main-header-top">
            <div>
                <h1>⛽ CRM Operacional AmPm</h1>
                <p>Gestão Estratégica de Capacitação, Logística de Viagens e Atendimento da Rede</p>
            </div>
            <div class="header-status-chip"><span class="pulse-dot"></span> SISTEMA ONLINE</div>
        </div>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# MÓDULO 1: DASHBOARD EXECUTIVO
# ==========================================
if modulo == "📊 Dashboard Executivo":
    render_section_header("📊", "Dashboard Executivo", "Panorama consolidado da operação em tempo real")
    if not df_base.empty:
        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-header">
                        <span class="kpi-title">Rede Filtrada</span>
                        <span class="kpi-icon-circle" style="background: var(--warning-bg); color: var(--ampm-orange-light);">🏪</span>
                    </div>
                    <div class="kpi-value">{len(df_base)}</div>
                    <div class="kpi-footer">unidades na seleção atual</div>
                </div>
            """, unsafe_allow_html=True)

        with c2:
            pendentes = len(df_base[df_base['Tipo_Necessidade'] != 'Rede Ativa (Sem Pendência)']) if 'Tipo_Necessidade' in df_base.columns else 0
            st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-header">
                        <span class="kpi-title">Fila Treinamento</span>
                        <span class="kpi-icon-circle" style="background: var(--warning-bg); color: var(--warning);">🎓</span>
                    </div>
                    <div class="kpi-value">{pendentes}</div>
                    <div class="kpi-footer">lojas com pendência de capacitação</div>
                </div>
            """, unsafe_allow_html=True)

        with c3:
            a_contatar = len(df_base[df_base['Status_Contato'] == 'A Contatar']) if 'Status_Contato' in df_base.columns else 0
            st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-header">
                        <span class="kpi-title">Pendentes Contato</span>
                        <span class="kpi-icon-circle" style="background: var(--danger-bg); color: var(--danger);">📞</span>
                    </div>
                    <div class="kpi-value">{a_contatar}</div>
                    <div class="kpi-footer">aguardando primeiro contato</div>
                </div>
            """, unsafe_allow_html=True)

        with c4:
            inaug = len(df_base[df_base['Previsão Inauguração'].notna()]) if 'Previsão Inauguração' in df_base.columns else 0
            st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-header">
                        <span class="kpi-title">Inaugurações</span>
                        <span class="kpi-icon-circle" style="background: var(--info-bg); color: var(--info);">🚀</span>
                    </div>
                    <div class="kpi-value">{inaug}</div>
                    <div class="kpi-footer">com previsão de abertura</div>
                </div>
            """, unsafe_allow_html=True)

        st.write("")
        st.divider()

        col_A, col_B = st.columns(2)
        with col_A:
            render_section_header("🗺️", "Concentração por Estado", "Top 10 UFs por número de lojas")
            if 'UF' in df_base.columns:
                st.bar_chart(df_base['UF'].value_counts().head(10), color="#FF9800")
        with col_B:
            render_section_header("📶", "Situação dos Contatos", "Distribuição por status no Call Center")
            if 'Status_Contato' in df_base.columns:
                st.bar_chart(df_base['Status_Contato'].value_counts(), color="#3B9EFF")
    else:
        st.info("📭 Nenhum dado carregado ainda. Envie o arquivo `Base_Unificada_AmPm.xlsx` na barra lateral.")

# ==========================================
# MÓDULO 2: PIPELINE AMPM
# ==========================================
elif modulo == "📋 Pipeline AmPm":
    render_section_header("📋", "Pipeline AmPm", "Fluxo operacional de treinamentos — navegue entre os estágios de contato")

    colunas_pipeline = ["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado"]
    cols_k = st.columns(len(colunas_pipeline))

    for idx, status in enumerate(colunas_pipeline):
        df_status = df_base[df_base['Status_Contato'] == status] if 'Status_Contato' in df_base.columns else pd.DataFrame()
        _, emoji_status = STATUS_BADGE_MAP.get(status, ("badge-neutral", "•"))

        with cols_k[idx]:
            st.markdown(f"""
                <div class="ampm-column {status_css_class(status)}">
                    <div class="ampm-title">
                        <span>{emoji_status} {status}</span>
                        <span class="pill-count">{len(df_status)}</span>
                    </div>
                </div>
            """, unsafe_allow_html=True)

            for _, item in df_status.head(6).iterrows():
                with st.expander(f"📍 PV {item.get('PV Abadi', '-')} | {str(item.get('Razão Social', ''))[:14]}..."):
                    st.write(f"**Cidade:** {item.get('Municipio', '-')}/{item.get('UF', '-')}")
                    st.write(f"**Necessidade:** {item.get('Tipo_Necessidade', '-')}")
                    st.write(f"**Treinandos:** {item.get('Qtd_Funcionarios', 0)} pessoas")
                    st.write(f"**Instrutor:** {item.get('Instrutor_Sugerido', 'Pendente')}")

                    mudar_status = st.selectbox(
                        "Alterar Status:",
                        colunas_pipeline,
                        index=colunas_pipeline.index(status),
                        key=f"pipe_sel_{item.get('PV Abadi')}"
                    )

                    if mudar_status != status:
                        atualizar_fila(item['PV Abadi'], {'Status_Contato': mudar_status})
                        st.success("Atualizado!")
                        st.rerun()

# ==========================================
# MÓDULO 3: PROCV & FILTROS AVANÇADOS
# ==========================================
elif modulo == "🔍 PROCV & Filtros Avançados":
    render_section_header("🔍", "PROCV & Filtros Avançados", "Consulta detalhada por PV, nome, município ou necessidade")
    if not df_base.empty:
        with st.expander("🔎 **Pesquisa Avançada na Base Filtrada**", expanded=True):
            f1, f2 = st.columns(2)
            termo = f1.text_input("🔍 PV, Nome ou Município:", "")
            f_necessidade = f2.selectbox("🎯 Tipo de Necessidade:", ["Todas"] + sorted([str(x) for x in df_base['Tipo_Necessidade'].dropna().unique()])) if 'Tipo_Necessidade' in df_base.columns else ["Todas"]

        df_view = df_base.copy()
        if termo:
            df_view = df_view[
                df_view.get('Razão Social', pd.Series(dtype=object)).astype(str).str.contains(termo, case=False, na=False) |
                df_view.get('PV Abadi', pd.Series(dtype=object)).astype(str).str.contains(termo, na=False) |
                df_view.get('Municipio', pd.Series(dtype=object)).astype(str).str.contains(termo, case=False, na=False)
            ]
        if f_necessidade != "Todas" and 'Tipo_Necessidade' in df_view.columns:
            df_view = df_view[df_view['Tipo_Necessidade'] == f_necessidade]

        st.caption("👇 *Clique em uma linha para abrir a Ficha Detalhada PROCV:*")

        cols_mostrar = [c for c in ['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Status Loja', 'Tipo_Necessidade', 'Status_Contato'] if c in df_view.columns]
        evento = st.dataframe(
            df_view[cols_mostrar],
            use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun"
        )

        linhas = evento.selection.get("rows", [])
        if linhas:
            p = df_view.iloc[linhas[0]].to_dict()
            st.divider()
            st.markdown(f"""
                <div style="display:flex; align-items:center; gap:10px; margin: 4px 0 16px 0;">
                    <span style="font-size:1.05rem; font-weight:700; color:var(--text-primary);">📋 Ficha PROCV — PV {p.get('PV Abadi', '-')} · {p.get('Razão Social', '-')}</span>
                </div>
            """, unsafe_allow_html=True)

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
                        <p>📅 <b>Inauguração:</b> {p.get('Previsão Inauguração', 'N/A')}</p>
                    </div>
                """, unsafe_allow_html=True)
            with k3:
                st.markdown(f"""
                    <div class="procv-card">
                        <h4>📞 Status do Atendimento</h4>
                        <p>🎯 <b>Necessidade:</b> {p.get('Tipo_Necessidade', '-')}</p>
                        <p>👨‍🏫 <b>Instrutor Alocado:</b> {p.get('Instrutor_Sugerido', '-')}</p>
                        <p>🔄 <b>Status Contato:</b> {badge_status_html(p.get('Status_Contato', '-'))}</p>
                    </div>
                """, unsafe_allow_html=True)
    else:
        st.info("📭 Nenhum dado carregado ainda.")

# ==========================================
# MÓDULO 4: CALCULADORA & OTIMIZADOR DE CUSTOS
# ==========================================
elif modulo == "📍 Calculadora & Otimizador de Custos":
    render_section_header("📍", "Calculadora & Otimizador de Custos", "Análise financeira e simulação de rotas com mapa interativo")

    if not df_rec.empty:
        df_rec_filtrado = df_rec.copy()
        if filtro_uf != "Todas" and 'UF_Loja' in df_rec_filtrado.columns:
            df_rec_filtrado = df_rec_filtrado[df_rec_filtrado['UF_Loja'] == filtro_uf]
        postos_unicos = df_rec_filtrado[['PV_ABADI', 'Razao_Social', 'Municipio_Loja', 'UF_Loja']].drop_duplicates()
        if not postos_unicos.empty:
            postos_unicos['label'] = postos_unicos['PV_ABADI'].astype(str) + " - " + postos_unicos['Razao_Social'] + " (" + postos_unicos['Municipio_Loja'] + "/" + postos_unicos['UF_Loja'] + ")"

            posto_sel = st.selectbox("⛽ Selecione o Posto Alvo:", postos_unicos['label'].tolist())
            pv_sel = int(posto_sel.split(" - ")[0])

            top_3 = df_rec_filtrado[df_rec_filtrado['PV_ABADI'] == pv_sel].sort_values(by='Ranking_Proximidade').head(3)

            if not top_3.empty:
                st.divider()

                # --- MAPA 3D PYDECK ---
                primeira = top_3.iloc[0]
                if pd.notna(primeira.get('Lat_Loja')) and pd.notna(primeira.get('Lon_Loja')) and pd.notna(primeira.get('Lat_Instrutor')) and pd.notna(primeira.get('Lon_Instrutor')):
                    p_lat, p_lon = float(primeira['Lat_Loja']), float(primeira['Lon_Loja'])
                    i_lat, i_lon = float(primeira['Lat_Instrutor']), float(primeira['Lon_Instrutor'])

                    df_mapa_pontos = pd.DataFrame([
                        {"name": f"Posto {primeira['PV_ABADI']}", "lat": p_lat, "lon": p_lon, "color": [226, 123, 0, 220]},
                        {"name": f"Instrutor {primeira['Instrutor_Sugerido']}", "lat": i_lat, "lon": i_lon, "color": [76, 175, 80, 220]}
                    ])

                    df_mapa_arco = pd.DataFrame([{
                        "from_lat": i_lat, "from_lon": i_lon, "to_lat": p_lat, "to_lon": p_lon
                    }])

                    layer_pontos = pdk.Layer(
                        "ScatterplotLayer",
                        df_mapa_pontos,
                        get_position="[lon, lat]",
                        get_color="color",
                        get_radius=20000,
                        pickable=True
                    )

                    layer_arco = pdk.Layer(
                        "ArcLayer",
                        df_mapa_arco,
                        get_source_position="[from_lon, from_lat]",
                        get_target_position="[to_lon, to_lat]",
                        get_source_color=[76, 175, 80, 180],
                        get_target_color=[226, 123, 0, 180],
                        get_width=4
                    )

                    view_state = pdk.ViewState(latitude=(p_lat + i_lat) / 2, longitude=(p_lon + i_lon) / 2, zoom=5, pitch=40)

                    st.markdown("##### 🗺️ Visualização Geográfica do Deslocamento")
                    st.pydeck_chart(pdk.Deck(layers=[layer_pontos, layer_arco], initial_view_state=view_state, tooltip={"text": "{name}"}))

                with st.expander("⚙️ **Ajustar Parâmetros Financeiros de Viagem**", expanded=False):
                    ca1, ca2, ca3, ca4 = st.columns(4)
                    v_km = ca1.number_input("Valor KM (Terrestre R$):", value=2.10)
                    v_passagem = ca2.number_input("Passagem Aérea Média (R$):", value=1400.0)
                    v_diaria = ca3.number_input("Diária Instrutor (Hosp./Alimentação R$):", value=280.0)
                    v_traslado = ca4.number_input("Traslado/Uber Aeroporto (R$):", value=150.0)

                col1, col2, col3 = st.columns(3)
                cols = [col1, col2, col3]
                custos_calculados = []
                for idx, (_, row) in enumerate(top_3.iterrows()):
                    dist = row['Distancia_km_linha_reta']
                    dias = row['Dias_Treinamento_Necessarios']

                    if dist <= 300:
                        modal = "Terrestre 🚗"
                        c_desloc = (dist * 2) * v_km
                        c_aereo = 0
                    else:
                        modal = "Aéreo ✈️"
                        c_desloc = v_traslado
                        c_aereo = v_passagem

                    c_hospedagem = dias * v_diaria
                    custo_total = c_desloc + c_aereo + c_hospedagem
                    custos_calculados.append(custo_total)
                    ranking_num = int(row['Ranking_Proximidade']) if pd.notna(row['Ranking_Proximidade']) else idx + 1
                    destaque = ' style="border-left-color: var(--ampm-orange);"' if ranking_num == 1 else ""
                    with cols[idx]:
                        st.markdown(f"""
                            <div class="top-instructor-card"{destaque}>
                                <h4 style="margin:0 0 8px 0; color:var(--ampm-orange-light);">#{ranking_num}º {row['Instrutor_Sugerido']}</h4>
                                <p style="margin:2px 0;">🏙️ <b>Origem:</b> {row['Cidade_Instrutor']}/{row['UF_Instrutor']}</p>
                                <p style="margin:2px 0;">📏 <b>Distância:</b> <code>{dist} km</code></p>
                                <p style="margin:2px 0;">✈️ <b>Modal:</b> {modal}</p>
                                <hr style="border-color:var(--border-subtle); margin:10px 0;">
                                <p style="margin:2px 0; font-size:0.85rem; color:var(--text-secondary);">• Deslocamento: R$ {c_desloc:.2f}</p>
                                <p style="margin:2px 0; font-size:0.85rem; color:var(--text-secondary);">• Passagem Aérea: R$ {c_aereo:.2f}</p>
                                <p style="margin:2px 0; font-size:0.85rem; color:var(--text-secondary);">• Diárias ({dias}d): R$ {c_hospedagem:.2f}</p>
                                <h3 style="color:var(--success); margin:12px 0 0 0; font-family:'JetBrains Mono',monospace;">Total: R$ {custo_total:.2f}</h3>
                            </div>
                        """, unsafe_allow_html=True)

                if len(custos_calculados) >= 2:
                    economia = custos_calculados[1] - custos_calculados[0]
                    st.success(f"💡 **Economia Eficiente:** Optar pelo **1º Instrutor Recomendado** garante uma economia estimada de **R$ {economia:.2f}** nesta operação.")
        else:
            st.info("Nenhum posto disponível para os filtros atuais.")
    else:
        st.info("Base de recomendação de deslocamento vazia ou não carregada.")

# ==========================================
# MÓDULO 5: CALL CENTER & TIMELINE WHATSAPP
# ==========================================
elif modulo == "📞 Call Center & Timeline WhatsApp":
    render_section_header("📞", "Call Center & Timeline WhatsApp", "Registro de atendimentos e disparo de mensagens")
    if not df_base.empty:
        df_fila_view = df_base[df_base['Tipo_Necessidade'] != 'Rede Ativa (Sem Pendência)'].copy() if 'Tipo_Necessidade' in df_base.columns else df_base.copy()

        c_left, c_right = st.columns([1.2, 1.8])

        with c_left:
            st.markdown("**📋 Fila de Atendimento**")
            cols_call = [c for c in ['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Status_Contato'] if c in df_fila_view.columns]
            evento_call = st.dataframe(
                df_fila_view[cols_call],
                use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun"
            )
            selecionado = evento_call.selection.get("rows", [])

        with c_right:
            if selecionado:
                posto = df_fila_view.iloc[selecionado[0]]
                pv_alvo = posto.get('PV Abadi')
                tel_limpo = ''.join(filter(str.isdigit, str(posto.get('Telefone_Contato', ''))))

                st.markdown(f"""
                    <div style="display:flex; align-items:center; gap:10px; margin: 4px 0 14px 0;">
                        <span style="font-size:1.05rem; font-weight:700; color:var(--text-primary);">📝 Ficha de Atendimento — PV {posto.get('PV Abadi', '-')}</span>
                        {badge_status_html(posto.get('Status_Contato', '-'))}
                    </div>
                """, unsafe_allow_html=True)

                st.markdown(f"""
                    <div class="procv-card">
                        <h4>🏪 Contexto do Posto (Consulta Rápida)</h4>
                        <div style="display: flex; justify-content: space-between; flex-wrap: wrap;">
                            <div style="flex: 1; min-width: 200px;">
                                <p>🏬 <b>Razão Social:</b> {posto.get('Razão Social', '-')}</p>
                                <p>📍 <b>Cidade/UF:</b> {posto.get('Municipio', '-')}/{posto.get('UF', '-')}</p>
                                <p>🏠 <b>Endereço:</b> {posto.get('Endereço', '-')}</p>
                            </div>
                            <div style="flex: 1; min-width: 200px;">
                                <p>👔 <b>Consultor (CF):</b> {posto.get('CF', '-')}</p>
                                <p>🎯 <b>Necessidade:</b> <span class="badge-info">{posto.get('Tipo_Necessidade', '-')}</span></p>
                                <p>⏱️ <b>Dias sem Treinamento:</b> {posto.get('Dias_desde_Ultimo_Treinamento', 'N/A')}</p>
                                <p>📅 <b>Inauguração Prevista:</b> {posto.get('Previsão Inauguração', 'N/A')}</p>
                            </div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)

                # --- INTEGRAÇÃO WHATSAPP FLEXÍVEL (LINK DIRETO + TEMPLATES) ---
                if tel_limpo:
                    st.markdown("##### 📲 Envio de Mensagem WhatsApp")
                    opcao_wa = st.radio("Selecione o estilo do envio:", ["Link Direto Rápido", "Template Customizado"], horizontal=True)

                    if opcao_wa == "Link Direto Rápido":
                        msg_final = f"Olá, equipe {posto.get('Razão Social', '')}! Aqui é da equipe de Capacitação AmPm. Gostaria de agendar o treinamento da loja."
                    else:
                        tmpl = st.selectbox("Escolha o Modelo de Mensagem:", [
                            "Agendamento de Treinamento",
                            "Cobrança / Verificação de Apostilas",
                            "Lembrete de Treinamento Agendado",
                            "Acompanhamento Pós-Treinamento"
                        ])

                        data_agendada_disp = posto.get('Data_Agendada')
                        data_agendada_fmt = parse_data_flexivel(data_agendada_disp)
                        data_agendada_fmt = data_agendada_fmt.strftime("%d/%m/%Y") if data_agendada_fmt else "em breve"

                        if tmpl == "Agendamento de Treinamento":
                            msg_final = f"Olá! Aqui é da Capacitação AmPm. Gostaríamos de confirmar as datas disponíveis para o treinamento na loja {posto.get('Razão Social', '')} (PV {posto.get('PV Abadi', '')})."
                        elif tmpl == "Cobrança / Verificação de Apostilas":
                            msg_final = f"Olá, equipe {posto.get('Razão Social', '')}! Para darmos início ao treinamento, poderiam confirmar se o material de apoio e apostilas já chegaram na loja?"
                        elif tmpl == "Lembrete de Treinamento Agendado":
                            msg_final = f"Olá! Passando para lembrar que o treinamento AmPm da loja {posto.get('Razão Social', '')} está agendado para o dia {data_agendada_fmt}. Contamos com todos!"
                        else:
                            msg_final = f"Olá! Como foi o treinamento concluído na loja {posto.get('Razão Social', '')}? Estamos à disposição para dúvidas ou feedbacks."

                    link_wa = f"https://wa.me/55{tel_limpo}?text={msg_final.replace(' ', '%20')}"
                    st.markdown(f"👉 **[Clique aqui para chamar no WhatsApp Direct]({link_wa})**")

                lista_instrutores = ["Pendente de Alocação"]
                if not df_instrutores.empty and 'NOME_COMPLETO' in df_instrutores.columns:
                    lista_instrutores += sorted(df_instrutores['NOME_COMPLETO'].dropna().unique().tolist())

                instrutor_atual = str(posto.get('Instrutor_Sugerido', 'Pendente de Alocação'))
                idx_instrutor = lista_instrutores.index(instrutor_atual) if instrutor_atual in lista_instrutores else 0

                data_inicial = parse_data_flexivel(posto.get('Data_Agendada')) or date.today()

                # --- REGISTROS RÁPIDOS DA LIGAÇÃO ---
                with st.form("form_callcenter_editavel"):
                    st.markdown("#### ✍️ Registros Rápidos da Ligação")

                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        nome_c = st.text_input("👤 Nome do Responsável na Loja:", value=str(posto.get('Nome_Contato', '') or ''))
                        tel_c = st.text_input("📞 Telefone de Contato:", value=str(posto.get('Telefone_Contato', '') or ''))
                        qtd_func = st.number_input("👥 Qtd. de Funcionários para Treinar:", value=int(posto.get('Qtd_Funcionarios', 0) or 0), min_value=0, step=1)
                        instrutor_escolhido = st.selectbox("👨‍🏫 Instrutor Designado:", lista_instrutores, index=idx_instrutor)

                    with col_f2:
                        status_opcoes = ["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado"]
                        st_atual = posto.get('Status_Contato', 'A Contatar')
                        idx_st = status_opcoes.index(st_atual) if st_atual in status_opcoes else 0
                        novo_st = st.selectbox("🔄 Status do Atendimento:", status_opcoes, index=idx_st)

                        mat_opcoes = ["Não Informado", "Sim", "Não"]
                        mat_atual = posto.get('Material_Em_Loja', 'Não Informado')
                        idx_mat = mat_opcoes.index(mat_atual) if mat_atual in mat_opcoes else 0
                        mat_loja = st.selectbox("📦 Possui Material/Apostilas na Loja?", mat_opcoes, index=idx_mat)
                        data_ag = st.date_input("📅 Data Agendada (Calendário):", value=data_inicial, format="DD/MM/YYYY")

                    obs = st.text_area("💬 Observações e Alinhamentos:", value=str(posto.get('Observacoes', '') or ''), height=80)

                    if st.form_submit_button("💾 Salvar Registro do Atendimento"):
                        atualizar_fila(pv_alvo, {
                            'Nome_Contato': nome_c,
                            'Telefone_Contato': tel_c,
                            'Qtd_Funcionarios': qtd_func,
                            'Instrutor_Sugerido': instrutor_escolhido,
                            'Material_Em_Loja': mat_loja,
                            'Data_Agendada': data_ag.strftime("%Y-%m-%d"),
                            'Status_Contato': novo_st,
                            'Observacoes': obs,
                            'Data_do_Contato': datetime.today().strftime('%d/%m/%Y %H:%M'),
                        })
                        st.success("✅ Atendimento registrado com sucesso!")
                        st.rerun()

                st.divider()
                st.markdown("**⏱️ Histórico de Interações**")
                data_ct = posto.get('Data_do_Contato', 'Sem registro')
                data_agendada_obj = parse_data_flexivel(posto.get('Data_Agendada'))
                data_agendada_str = data_agendada_obj.strftime("%d/%m/%Y") if data_agendada_obj else "Não agendado"
                st.markdown(f"""
                    <div class="timeline-item">
                        <small style="color:var(--text-tertiary);"><b>Última Atualização:</b> {data_ct}</small><br>
                        <span style="color:var(--text-primary);">{badge_status_html(posto.get('Status_Contato', '-'))} &nbsp;·&nbsp; <b>Agendado:</b> {data_agendada_str} &nbsp;·&nbsp; <b>Instrutor:</b> {posto.get('Instrutor_Sugerido', '-')}</span><br>
                        <span style="color:var(--text-secondary);"><i>"{posto.get('Observacoes', 'Sem observações registradas.')}"</i></span>
                    </div>
                """, unsafe_allow_html=True)
    else:
        st.info("📭 Nenhum dado carregado ainda.")

# ==========================================
# MÓDULO 6: EQUIPE DE INSTRUTORES
# ==========================================
elif modulo == "👔 Equipe de Instrutores":
    render_section_header("👔", "Equipe de Instrutores", "Instrutores credenciados na rede")
    if not df_instrutores.empty:
        cols_inst = [c for c in ['NOME_COMPLETO', 'STATUS', 'TELEFONE', 'EMAIL', 'Cidade', 'UF'] if c in df_instrutores.columns]
        st.dataframe(df_instrutores[cols_inst], use_container_width=True, hide_index=True)
    else:
        st.info("📭 Nenhum instrutor carregado ainda.")

# ==========================================
# MÓDULO 7: RELATÓRIOS & EXPORTAÇÃO
# ==========================================
elif modulo == "📂 Relatórios & Exportação":
    render_section_header("📂", "Relatórios & Exportação", "Baixe os dados operacionais atualizados em tempo real")

    col_exp1, col_exp2 = st.columns(2)

    csv_buffer = df_base.to_csv(index=False).encode('utf-8')
    with col_exp1:
        st.download_button(
            label="📄 Baixar Base Filtrada em CSV",
            data=csv_buffer,
            file_name=f"Base_CRM_AmPm_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_base.to_excel(writer, sheet_name='Base_CRM_Atualizada', index=False)
    excel_data = output.getvalue()

    with col_exp2:
        st.download_button(
            label="📊 Baixar Base Filtrada em Excel",
            data=excel_data,
            file_name=f"Base_CRM_AmPm_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
