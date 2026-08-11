import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
import pydeck as pdk
import io
import time
import requests
import json
import re
import unicodedata
from difflib import SequenceMatcher
import streamlit_authenticator as stauth

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="CRM Operacional AmPm",
    page_icon="⛽",
    layout="wide",
    initial_sidebar_state="expanded"
)

CAMINHO_ARQUIVO = "Base_Unificada_AmPm.xlsx"
CAMINHO_BACKUP = "Base_Unificada_AmPm.backup.xlsx"

COLUNAS_FILA = [
    "PV_Abadi", "Tipo_Necessidade", "Data_Ultimo_Treinamento",
    "Dias_desde_Ultimo_Treinamento", "Instrutor_Sugerido", "Semana_Sugerida",
    "Telefone_Contato", "Status_Contato", "Data_do_Contato", "Observacoes",
    "Nome_Contato", "Qtd_Funcionarios", "Material_Em_Loja", "Data_Agendada",
]

# --- ESTILIZAÇÃO CSS CUSTOMIZADA ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500&display=swap');

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

    .block-container { padding-top: 1.6rem; padding-bottom: 3rem; }

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

    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: var(--border-strong); border-radius: 8px; }
    ::-webkit-scrollbar-thumb:hover { background: var(--text-tertiary); }
    </style>
""", unsafe_allow_html=True)

# --- AUTENTICAÇÃO ---
CAMINHO_USUARIOS = "usuarios_ampm.json"

def _secrets_para_dict(obj):
    if hasattr(obj, "items"):
        return {chave: _secrets_para_dict(valor) for chave, valor in obj.items()}
    return obj

def carregar_usuarios_arquivo():
    if os.path.exists(CAMINHO_USUARIOS):
        try:
            with open(CAMINHO_USUARIOS, "r", encoding="utf-8") as f:
                dados = json.load(f)
            if isinstance(dados, dict) and "usernames" in dados:
                return dados
        except Exception:
            pass
    return {"usernames": {}}

def salvar_usuarios_arquivo(credenciais_completas):
    try:
        secrets_usernames = set()
        try:
            secrets_usernames = set(_secrets_para_dict(st.secrets["credentials"]).get("usernames", {}).keys())
        except Exception:
            pass
        usernames_para_salvar = {
            usuario: dados for usuario, dados in credenciais_completas.get("usernames", {}).items()
            if usuario not in secrets_usernames
        }
        with open(CAMINHO_USUARIOS, "w", encoding="utf-8") as f:
            json.dump({"usernames": usernames_para_salvar}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.toast(f"⚠️ Erro ao salvar cadastro: {e}", icon="⚠️")

def _tela_marca_login(subtitulo):
    st.markdown(f"""
        <div style="display:flex; justify-content:center; margin: 40px 0 24px 0;">
            <div style="display:flex; align-items:center; gap:14px; background:var(--bg-surface);
                        border:1px solid var(--border-subtle); border-radius:var(--radius-lg);
                        padding:18px 28px; box-shadow:var(--shadow-md);">
                <div class="logo-chip" style="width:46px; height:46px; font-size:1.4rem;">⛽</div>
                <div>
                    <div style="font-weight:800; font-size:1.35rem; color:var(--text-primary);">CRM AmPm</div>
                    <div style="font-size:0.82rem; color:var(--text-tertiary);">{subtitulo}</div>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

def exigir_login():
    try:
        cookie_key = st.secrets["COOKIE_KEY"]
    except Exception:
        _tela_marca_login("Configuração de acesso pendente")
        st.warning("🔒 O login ainda não foi configurado neste app.")
        st.markdown(
            "Configure a chave `COOKIE_KEY` em **⋮ → Settings → Secrets** no Streamlit Cloud "
            "para habilitar o acesso."
        )
        st.stop()

    credenciais_arquivo = carregar_usuarios_arquivo()
    try:
        credenciais_secrets = _secrets_para_dict(st.secrets["credentials"])
    except Exception:
        credenciais_secrets = {"usernames": {}}

    credenciais = {
        "usernames": {
            **credenciais_arquivo.get("usernames", {}),
            **credenciais_secrets.get("usernames", {}),
        }
    }

    try:
        dominios_permitidos_raw = st.secrets.get("ALLOWED_EMAIL_DOMAINS", "")
    except Exception:
        dominios_permitidos_raw = ""
    dominios_permitidos = [d.strip() for d in str(dominios_permitidos_raw).split(",") if d.strip()] or None

    autenticador = stauth.Authenticate(
        credenciais,
        cookie_name="crm_ampm_auth",
        cookie_key=cookie_key,
        cookie_expiry_days=7,
        auto_hash=True,
    )

    if not st.session_state.get("authentication_status"):
        _tela_marca_login("Acesso restrito — faça login ou crie sua conta")
        aba_login, aba_cadastro = st.tabs(["🔑 Entrar", "🆕 Criar conta"])

        with aba_login:
            autenticador.login(location="main", key="LoginPrincipal")

        with aba_cadastro:
            if dominios_permitidos:
                st.caption(f"✉️ Cadastro liberado apenas para e-mails: {', '.join(dominios_permitidos)}")
            try:
                email_novo, usuario_novo, nome_novo = autenticador.register_user(
                    location="main",
                    domains=dominios_permitidos,
                    password_hint=False,
                    fields={
                        "Form name": "Criar minha conta",
                        "First name": "Nome",
                        "Last name": "Sobrenome",
                        "Email": "E-mail",
                        "Username": "Usuário (para login)",
                        "Password": "Senha",
                        "Repeat password": "Repita a senha",
                        "Register": "Criar conta",
                    },
                    captcha=False,
                )
                if email_novo:
                    salvar_usuarios_arquivo(autenticador.authentication_controller.authentication_model.credentials)
                    st.success(f"✅ Conta criada para **{nome_novo}**! Vá até a aba '🔑 Entrar' e faça login.")
            except Exception as e:
                msg = str(e)
                if "domain" in msg.lower():
                    st.error(f"❌ Esse e-mail não pertence a um domínio autorizado ({', '.join(dominios_permitidos or [])}).")
                elif "already taken" in msg.lower() or "already exists" in msg.lower():
                    st.error("❌ Esse usuário ou e-mail já está cadastrado.")
                elif "match" in msg.lower():
                    st.error("❌ As senhas digitadas não coincidem.")
                else:
                    st.error(f"❌ Não foi possível criar a conta: {msg}")

    status_login = st.session_state.get("authentication_status")
    if status_login is False:
        st.error("❌ Usuário ou senha incorretos.")
        st.stop()
    elif status_login is None:
        st.stop()

    return autenticador

AUTENTICADOR = exigir_login()

# --- HELPERS DE APRESENTAÇÃO ---
def render_section_header(icone, titulo, subtitulo=""):
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
    classe, emoji = STATUS_BADGE_MAP.get(str(status), ("badge-neutral", "•"))
    return f'<span class="badge-pill {classe}">{emoji} {status}</span>'

def status_css_class(status):
    mapa = {
        "A Contatar": "col-a-contatar",
        "Em Negociação": "col-em-negociacao",
        "Agendado": "col-agendado",
        "Treinamento Realizado": "col-treinamento-realizado",
        "Recusado": "col-recusado",
    }
    return mapa.get(status, "")

# --- CAMADA DE DADOS ---
def parse_data_flexivel(valor):
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
    """Normaliza nomes de colunas para comparação robusta.
    Aceita acentos, pontuação, underscores, hífens e diferenças de caixa.
    """
    texto = "" if nome is None else str(nome).strip().lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return " ".join(texto.split())


def _similaridade_coluna(a, b):
    """Calcula similaridade entre dois nomes de coluna.
    Além da comparação textual, considera palavras em comum.
    """
    a_norm = _normalizar_nome(a)
    b_norm = _normalizar_nome(b)
    if not a_norm or not b_norm:
        return 0.0
    if a_norm == b_norm:
        return 1.0
    tokens_a = set(a_norm.split())
    tokens_b = set(b_norm.split())
    inter = len(tokens_a & tokens_b)
    cobertura = inter / max(len(tokens_a), len(tokens_b), 1)
    sequencia = SequenceMatcher(None, a_norm, b_norm).ratio()
    return max(sequencia, cobertura * 0.92)


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
            "UF_Instrutor": ["uf instrutor", "uf do instrutor", "estado instrutor"],
            "Ranking_Proximidade": ["ranking proximidade", "ranking", "posicao ranking"],
            "Distancia_km_linha_reta": ["distancia km linha reta", "distancia km", "distancia", "distancia linha reta"],
            "Dias_Treinamento_Necessarios": ["dias treinamento necessarios", "dias necessarios", "dias de treinamento"],
        },
    },
}

MIN_SCORE_CONFIANTE = 2
FUZZY_THRESHOLD = 0.82


def _construir_lookup(colunas_dict):
    lookup = {}
    for canonico, apelidos in colunas_dict.items():
        for apelido in set(apelidos) | {canonico}:
            lookup[_normalizar_nome(apelido)] = canonico
    return lookup


def _mapear_colunas_compativeis(df, definicao_entidade):
    """Reconhece colunas por nome exato, apelido e similaridade.
    Colunas novas que não existem no dicionário são preservadas como colunas dinâmicas.
    """
    lookup = _construir_lookup(definicao_entidade["colunas"])
    rename_map = {}
    canonicas_encontradas = set()
    colunas_ignoradas = []
    colunas_novas = []

    for col in df.columns:
        original = str(col)
        chave_norm = _normalizar_nome(original)
        canonico = lookup.get(chave_norm)

        if canonico is None and chave_norm:
            melhor = None
            melhor_score = 0.0
            for alias_norm, destino in lookup.items():
                score = _similaridade_coluna(chave_norm, alias_norm)
                if score > melhor_score and destino not in canonicas_encontradas:
                    melhor = destino
                    melhor_score = score
            if melhor is not None and melhor_score >= FUZZY_THRESHOLD:
                canonico = melhor

        if canonico and canonico not in canonicas_encontradas:
            rename_map[original] = canonico
            canonicas_encontradas.add(canonico)
        else:
            colunas_novas.append(original)

    return rename_map, canonicas_encontradas, colunas_novas


def _normalizar_chave_dataframe(df, chave, numerica=True):
    df = df.copy()
    if chave not in df.columns:
        return df
    if numerica:
        df[chave] = pd.to_numeric(df[chave], errors="coerce")
    else:
        df[chave] = df[chave].astype("string").str.strip()
    return df


def _coluna_dinamica_segura(nome, existentes):
    base = str(nome).strip()
    if not base:
        return None
    if base in existentes:
        return base
    candidato = base
    contador = 2
    while candidato in existentes:
        candidato = f"{base}_{contador}"
        contador += 1
    return candidato


def _preparar_dataframe_entidade(df_bruto, definicao):
    """Converte uma tabela externa para o modelo interno sem descartar informação nova."""
    rename_map, canonicas, colunas_novas = _mapear_colunas_compativeis(df_bruto, definicao)
    df_mapeado = df_bruto.rename(columns=rename_map).copy()
    existentes = set(df_mapeado.columns)

    # Mantém colunas novas em vez de descartá-las. Isso permite ao CRM aprender
    # campos adicionais sem exigir alteração de código para cada nova planilha.
    for coluna in list(colunas_novas):
        if coluna not in df_mapeado.columns:
            continue
        nova = _coluna_dinamica_segura(coluna, existentes)
        if nova and nova != coluna:
            df_mapeado.rename(columns={coluna: nova}, inplace=True)
            existentes.add(nova)

    return df_mapeado, rename_map, canonicas, colunas_novas


def _score_aba_para_entidade(df, sheet_name, entidade, definicao):
    _, canonicas, _ = _mapear_colunas_compativeis(df, definicao)
    chave = definicao["chave"]
    if chave not in canonicas:
        return -1, canonicas

    score = len(canonicas)
    nome_aba = _normalizar_nome(sheet_name)
    pistas = {
        "lojas": ["loja", "rede", "posto", "base", "cadastro"],
        "fila": ["fila", "call", "contato", "treinamento"],
        "inaug": ["inaug", "abertura", "pipeline"],
        "instrutores": ["instrutor", "equipe", "professor"],
        "rec": ["recomend", "desloc", "rota", "proximidade"],
    }
    score += sum(1 for pista in pistas.get(entidade, []) if pista in nome_aba)
    return score, canonicas


def detectar_entidades_no_workbook(xls):
    """Lê todas as abas e identifica o tipo pelo conteúdo, não pelo nome da aba."""
    dfs_brutos = {}
    candidatos = []

    for sheet_name in xls.sheet_names:
        try:
            df_bruto = pd.read_excel(xls, sheet_name=sheet_name)
        except Exception:
            continue
        if df_bruto is None or len(df_bruto.columns) == 0:
            continue
        df_bruto = df_bruto.dropna(axis=0, how="all").dropna(axis=1, how="all")
        if df_bruto.empty:
            continue
        dfs_brutos[sheet_name] = df_bruto

        for entidade, definicao in ENTIDADES.items():
            score, canonicas = _score_aba_para_entidade(df_bruto, sheet_name, entidade, definicao)
            if score >= 1 and definicao["chave"] in canonicas:
                candidatos.append((score, sheet_name, entidade, canonicas))

    prioridade = {"lojas": 5, "fila": 4, "inaug": 3, "instrutores": 2, "rec": 1}
    candidatos.sort(key=lambda x: (-x[0], -prioridade.get(x[2], 0), x[1]))
    entidade_atribuida = {}
    aba_usada = set()

    for score, sheet_name, entidade, _ in candidatos:
        if entidade in entidade_atribuida or sheet_name in aba_usada:
            continue
        entidade_atribuida[entidade] = sheet_name
        aba_usada.add(sheet_name)

    # Rede de lojas é a âncora do CRM. Se houver mais de uma candidata,
    # escolhemos a que tiver mais campos reconhecidos.
    if "lojas" not in entidade_atribuida:
        raise ValueError("Nenhuma aba com uma chave de loja/PV foi reconhecida.")

    bases = {}
    relatorio = []

    for entidade, definicao in ENTIDADES.items():
        colunas_canonicas = list(definicao["colunas"].keys())
        sheet_name = entidade_atribuida.get(entidade)

        if sheet_name:
            df_bruto = dfs_brutos[sheet_name]
            df_final, rename_map, canonicas, colunas_novas = _preparar_dataframe_entidade(df_bruto, definicao)
            df_final = _normalizar_chave_dataframe(
                df_final,
                definicao["chave"],
                definicao.get("chave_numerica", True),
            )
            bases[entidade] = df_final
            relatorio.append({
                "entidade": entidade,
                "aba_origem": sheet_name,
                "confianca": "alta" if len(canonicas) >= MIN_SCORE_CONFIANTE else "média",
                "colunas_reconhecidas": [c for c in colunas_canonicas if c in df_final.columns],
                "colunas_novas": [c for c in df_final.columns if c not in colunas_canonicas],
                "colunas_ignoradas": [],
                "linhas_lidas": len(df_final),
            })
        else:
            bases[entidade] = pd.DataFrame(columns=colunas_canonicas)
            relatorio.append({
                "entidade": entidade,
                "aba_origem": None,
                "confianca": "n/a",
                "colunas_reconhecidas": [],
                "colunas_novas": [],
                "colunas_ignoradas": [],
                "linhas_lidas": 0,
            })

    for col in COLUNAS_FILA:
        if col not in bases["fila"].columns:
            bases["fila"][col] = pd.NA

    return bases, relatorio


def _valor_preenchido(valor):
    if valor is None:
        return False
    try:
        if pd.isna(valor):
            return False
    except Exception:
        pass
    return str(valor).strip().lower() not in ("", "nan", "nat", "none", "null")


def mesclar_entidade_existente(df_atual, df_novo, definicao):
    """UPSERT inteligente: atualiza registros existentes e adiciona registros novos.
    Valores vazios do arquivo novo nunca apagam informação existente.
    """
    chave = definicao["chave"]
    numerica = definicao.get("chave_numerica", True)
    atual = _normalizar_chave_dataframe(df_atual if df_atual is not None else pd.DataFrame(), chave, numerica)
    novo = _normalizar_chave_dataframe(df_novo if df_novo is not None else pd.DataFrame(), chave, numerica)

    if novo.empty:
        return atual, 0, 0, []
    if chave not in novo.columns:
        return atual, 0, 0, []

    # Remove chaves vazias: não existe forma segura de fazer upsert sem identificador.
    novo = novo[novo[chave].notna()].copy()
    if novo.empty:
        return atual, 0, 0, []

    # Mantém todas as colunas existentes e acrescenta as novas.
    for coluna in novo.columns:
        if coluna not in atual.columns:
            atual[coluna] = pd.NA
    for coluna in atual.columns:
        if coluna not in novo.columns:
            novo[coluna] = pd.NA
    novo = novo[atual.columns.tolist() + [c for c in novo.columns if c not in atual.columns]]

    if atual.empty:
        resultado = novo.drop_duplicates(subset=[chave], keep="last").reset_index(drop=True)
        return resultado, 0, len(resultado), [c for c in novo.columns if c not in definicao["colunas"]]

    atual = atual.reset_index(drop=True)
    novo = novo.reset_index(drop=True)
    indice_atual = {}
    for idx, valor in atual[chave].items():
        if _valor_preenchido(valor):
            indice_atual[str(valor)] = idx

    atualizados = 0
    adicionados = 0
    novas_colunas = [c for c in novo.columns if c not in definicao["colunas"]]

    for _, linha in novo.iterrows():
        chave_linha = str(linha[chave])
        if chave_linha in indice_atual:
            idx = indice_atual[chave_linha]
            mudou = False
            for coluna, valor in linha.items():
                if coluna == chave:
                    continue
                if _valor_preenchido(valor):
                    valor_atual = atual.at[idx, coluna]
                    if not _valor_preenchido(valor_atual) or str(valor_atual) != str(valor):
                        atual.at[idx, coluna] = valor
                        mudou = True
            if mudou:
                atualizados += 1
        else:
            nova_linha = {col: pd.NA for col in atual.columns}
            for coluna, valor in linha.items():
                if coluna in nova_linha:
                    nova_linha[coluna] = valor
            atual = pd.concat([atual, pd.DataFrame([nova_linha])], ignore_index=True)
            indice_atual[chave_linha] = len(atual) - 1
            adicionados += 1

    return atual, atualizados, adicionados, novas_colunas


def mesclar_bases(bases_atuais, bases_novas):
    """Aplica UPSERT em todas as entidades reconhecidas."""
    resultado = dict(bases_atuais or _bases_vazias())
    estatisticas = []

    for entidade, definicao in ENTIDADES.items():
        atual = resultado.get(entidade, pd.DataFrame(columns=list(definicao["colunas"].keys())))
        novo = bases_novas.get(entidade, pd.DataFrame())
        combinado, atualizados, adicionados, novas_colunas = mesclar_entidade_existente(atual, novo, definicao)
        resultado[entidade] = combinado
        estatisticas.append({
            "entidade": entidade,
            "atualizados": atualizados,
            "adicionados": adicionados,
            "novas_colunas": novas_colunas,
        })

    for col in COLUNAS_FILA:
        if col not in resultado["fila"].columns:
            resultado["fila"][col] = pd.NA

    return resultado, estatisticas


def _processar_excelfile(xls, exigir_lojas=False):
    bases, relatorio = detectar_entidades_no_workbook(xls)
    if exigir_lojas and (bases["lojas"].empty or ENTIDADES["lojas"]["chave"] not in bases["lojas"].columns):
        abas_disponiveis = ", ".join(xls.sheet_names) if xls.sheet_names else "(nenhuma)"
        raise ValueError(
            "Não foi possível identificar uma base de lojas no banco principal. "
            f"Abas presentes: {abas_disponiveis}. "
            "A base principal precisa conter uma coluna equivalente a PV."
        )
    return bases, relatorio


def _ler_csv_flexivel(uploaded_file):
    """Lê CSV com diferentes codificações e separadores comuns no Brasil."""
    dados = uploaded_file.getvalue()
    ultimo_erro = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        for sep in (None, ";", ",", "\t", "|"):
            try:
                kwargs = {"encoding": encoding}
                if sep is None:
                    kwargs.update({"sep": None, "engine": "python"})
                else:
                    kwargs["sep"] = sep
                df = pd.read_csv(io.BytesIO(dados), **kwargs)
                if len(df.columns) > 1:
                    return df
            except Exception as exc:
                ultimo_erro = exc
    raise ValueError(f"Não foi possível interpretar o CSV: {ultimo_erro}")


def validar_bytes_excel(conteudo_bytes):
    try:
        xls = pd.ExcelFile(io.BytesIO(conteudo_bytes), engine="openpyxl")
        bases, relatorio = _processar_excelfile(xls, exigir_lojas=False)
        return bases, relatorio, None
    except Exception as e:
        return None, None, str(e)


@st.cache_data
def carregar_bases_do_disco(caminho, assinatura=None):
    if not os.path.exists(caminho):
        return None, None
    xls = pd.ExcelFile(caminho, engine="openpyxl")
    return _processar_excelfile(xls, exigir_lojas=True)


def construir_base_unificada(df_lojas, df_fila, df_inaug):
    if df_lojas is None or df_lojas.empty:
        return pd.DataFrame()

    df_base = df_lojas.copy()

    if df_fila is not None and not df_fila.empty and "PV_Abadi" in df_fila.columns:
        colunas_fila = [c for c in df_fila.columns if c != "PV_Abadi"]
        df_fila_merge = df_fila[["PV_Abadi"] + colunas_fila].copy()
        df_base = pd.merge(df_base, df_fila_merge, left_on="PV Abadi", right_on="PV_Abadi", how="left")

    if df_inaug is not None and not df_inaug.empty and "PV ABADI" in df_inaug.columns:
        colunas_inaug = [c for c in df_inaug.columns if c != "PV ABADI"]
        df_inaug_merge = df_inaug[["PV ABADI"] + colunas_inaug].copy()
        df_base = pd.merge(df_base, df_inaug_merge, left_on="PV Abadi", right_on="PV ABADI", how="left")

    defaults = {
        "Status_Contato": "A Contatar",
        "Tipo_Necessidade": "Rede Ativa (Sem Pendência)",
        "Instrutor_Sugerido": "Pendente de Alocação",
        "Nome_Contato": "",
        "Material_Em_Loja": "Não Informado",
    }
    for coluna, valor in defaults.items():
        if coluna in df_base.columns:
            df_base[coluna] = df_base[coluna].fillna(valor)

    if "Qtd_Funcionarios" in df_base.columns:
        df_base["Qtd_Funcionarios"] = pd.to_numeric(df_base["Qtd_Funcionarios"], errors="coerce").fillna(0).astype(int)

    return df_base


def salvar_bases_combinadas_no_disco(bases, caminho=CAMINHO_ARQUIVO):
    """Persiste as entidades no Excel preservando abas que não pertencem ao CRM."""
    if os.path.exists(caminho):
        with pd.ExcelFile(caminho, engine="openpyxl") as xls:
            abas_originais = {}
            for aba in xls.sheet_names:
                try:
                    abas_originais[aba] = pd.read_excel(xls, sheet_name=aba)
                except Exception:
                    pass
    else:
        abas_originais = {}

    nomes_entidades = {
        "lojas": "Rede_de_Lojas",
        "fila": "Fila_CallCenter",
        "inaug": "Previsao_Inauguracao",
        "instrutores": "Instrutores",
        "rec": "Recomendacao_Deslocamento",
    }

    for entidade, nome_aba in nomes_entidades.items():
        abas_originais[nome_aba] = bases.get(entidade, pd.DataFrame())

    with pd.ExcelWriter(caminho, engine="openpyxl", mode="w") as writer:
        for nome_aba, df in abas_originais.items():
            nome_seguro = str(nome_aba)[:31] or "Dados"
            df.to_excel(writer, sheet_name=nome_seguro, index=False)


def salvar_fila_no_disco():
    if not os.path.exists(CAMINHO_ARQUIVO):
        st.toast("⚠️ Arquivo local não encontrado — alterações mantidas apenas na sessão.", icon="⚠️")
        return
    try:
        bases = st.session_state["bases"]
        salvar_bases_combinadas_no_disco(bases)
        st.toast("💾 Banco de dados salvo com sucesso!", icon="✅")
    except Exception as e:
        st.toast(f"⚠️ Erro ao salvar banco de dados: {e}", icon="⚠️")


def salvar_lojas_no_disco():
    if not os.path.exists(CAMINHO_ARQUIVO):
        st.toast("⚠️ Arquivo local não encontrado — alterações mantidas apenas na sessão.", icon="⚠️")
        return
    try:
        salvar_bases_combinadas_no_disco(st.session_state["bases"])
        st.toast("💾 Rede de Lojas salva no banco de dados!", icon="✅")
    except Exception as e:
        st.toast(f"⚠️ Erro ao salvar banco de dados: {e}", icon="⚠️")

def buscar_telefone_google_places(endereco_completo, nome_loja, api_key, timeout=8):
    consulta = f"{nome_loja}, {endereco_completo}" if nome_loja else endereco_completo
    try:
        resp_busca = requests.get(
            "https://maps.googleapis.com/maps/api/place/findplacefromtext/json",
            params={
                "input": consulta,
                "inputtype": "textquery",
                "fields": "place_id",
                "language": "pt-BR",
                "key": api_key,
            },
            timeout=timeout,
        )
        dados_busca = resp_busca.json()
        status = dados_busca.get("status")
        if status != "OK" or not dados_busca.get("candidates"):
            return None, status or "SEM_RESULTADO"

        place_id = dados_busca["candidates"][0]["place_id"]

        resp_detalhes = requests.get(
            "https://maps.googleapis.com/maps/api/place/details/json",
            params={
                "place_id": place_id,
                "fields": "formatted_phone_number,international_phone_number",
                "language": "pt-BR",
                "key": api_key,
            },
            timeout=timeout,
        )
        dados_detalhes = resp_detalhes.json()
        if dados_detalhes.get("status") != "OK":
            return None, dados_detalhes.get("status", "ERRO_DETALHES")

        resultado = dados_detalhes.get("result", {})
        telefone = resultado.get("formatted_phone_number") or resultado.get("international_phone_number")
        if telefone:
            return telefone, "OK"
        return None, "SEM_TELEFONE_CADASTRADO"
    except requests.exceptions.RequestException as e:
        return None, f"ERRO_REDE: {e}"
    except Exception as e:
        return None, f"ERRO: {e}"

def atualizar_fila(pv_abadi, campos: dict):
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
        st.session_state['erro_carga'] = str(e)
        st.session_state['bases'] = _bases_vazias()

inicializar_estado()

if st.session_state.get('erro_carga'):
    st.error(
        "⚠️ Não foi possível carregar `Base_Unificada_AmPm.xlsx`:\n\n"
        f"{st.session_state['erro_carga']}\n\n"
        "Envie um arquivo válido na barra lateral."
    )

# --- SIDEBAR & NAVEGAÇÃO ---
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

    st.markdown(f"""
        <div class="sidebar-metric" style="display:flex; align-items:center; justify-content:space-between;">
            <span>👤 <b>{st.session_state.get('name', 'Usuário')}</b></span>
        </div>
    """, unsafe_allow_html=True)
    AUTENTICADOR.logout("🚪 Sair", "sidebar")

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
            "📇 Enriquecimento de Rede",
            "📂 Relatórios & Exportação"
        ]
    )

    st.divider()

    st.markdown("📥 **Atualizar Banco de Dados**")
    st.caption("O CRM interpreta o conteúdo, reconhece colunas parecidas e faz atualização incremental: registros novos entram e registros existentes são atualizados sem apagar dados válidos.")

    aba_destino_csv = st.selectbox(
        "Destino para CSV:",
        ["Rede_de_Lojas", "Fila_CallCenter", "Previsao_Inauguracao", "Instrutores", "Recomendacao_Deslocamento"]
    )

    uploaded_file = st.file_uploader(
        "Envie a nova planilha (.xlsx ou .csv):",
        type=["xlsx", "csv"]
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
                    if item.get("colunas_novas"):
                        st.caption("🆕 Colunas novas armazenadas: " + ", ".join(item["colunas_novas"]))
                    if item.get("linhas_lidas") is not None:
                        st.caption(f"Linhas lidas: {item['linhas_lidas']}")
                    if item["colunas_ignoradas"]:
                        st.caption("Colunas ignoradas: " + ", ".join(item["colunas_ignoradas"]))
                else:
                    st.markdown(f"**{nome}** — não encontrado no arquivo")

    if uploaded_file is not None:
        try:
            if uploaded_file.name.lower().endswith(".xlsx"):
                conteudo = uploaded_file.getbuffer().tobytes()
                bases_validadas, relatorio, erro = validar_bytes_excel(conteudo)
                if erro:
                    st.error(f"❌ Arquivo rejeitado:\n\n{erro}")
                else:
                    bases_atuais = st.session_state.get("bases", _bases_vazias())
                    bases_combinadas, estatisticas = mesclar_bases(bases_atuais, bases_validadas)

                    if os.path.exists(CAMINHO_ARQUIVO):
                        with open(CAMINHO_ARQUIVO, "rb") as f_atual, open(CAMINHO_BACKUP, "wb") as f_bak:
                            f_bak.write(f_atual.read())

                    salvar_bases_combinadas_no_disco(bases_combinadas)
                    st.cache_data.clear()
                    st.session_state["bases"] = bases_combinadas
                    st.session_state["relatorio_importacao"] = relatorio
                    st.session_state["estatisticas_importacao"] = estatisticas
                    st.session_state["erro_carga"] = None
                    total_novos = sum(x["adicionados"] for x in estatisticas)
                    total_atualizados = sum(x["atualizados"] for x in estatisticas)
                    st.toast(
                        f"💾 Importação concluída: {total_novos} novos e {total_atualizados} atualizados.",
                        icon="✅",
                    )
                    st.rerun()

            elif uploaded_file.name.lower().endswith(".csv"):
                df_csv = _ler_csv_flexivel(uploaded_file)
                chave_map = {
                    "Rede_de_Lojas": "lojas",
                    "Fila_CallCenter": "fila",
                    "Previsao_Inauguracao": "inaug",
                    "Instrutores": "instrutores",
                    "Recomendacao_Deslocamento": "rec",
                }

                candidatos_csv = []
                for entidade, definicao in ENTIDADES.items():
                    score, canonicas = _score_aba_para_entidade(df_csv, uploaded_file.name, entidade, definicao)
                    if definicao["chave"] in canonicas:
                        candidatos_csv.append((score, entidade))
                candidatos_csv.sort(reverse=True)

                if candidatos_csv and candidatos_csv[0][0] >= MIN_SCORE_CONFIANTE:
                    chave = candidatos_csv[0][1]
                    origem = "identificação automática"
                else:
                    chave = chave_map[aba_destino_csv]
                    origem = "destino selecionado"

                definicao = ENTIDADES[chave]
                df_preparado, _, canonicas, _ = _preparar_dataframe_entidade(df_csv, definicao)
                if definicao["chave"] not in canonicas:
                    st.error(f"❌ O arquivo não possui uma coluna equivalente a '{definicao['chave']}'.")
                else:
                    bases_atuais = st.session_state.get("bases", _bases_vazias())
                    bases_novas = _bases_vazias()
                    bases_novas[chave] = df_preparado
                    bases_combinadas, estatisticas = mesclar_bases(bases_atuais, bases_novas)

                    if os.path.exists(CAMINHO_ARQUIVO):
                        with open(CAMINHO_ARQUIVO, "rb") as f_atual, open(CAMINHO_BACKUP, "wb") as f_bak:
                            f_bak.write(f_atual.read())

                    salvar_bases_combinadas_no_disco(bases_combinadas)
                    st.cache_data.clear()
                    st.session_state["bases"] = bases_combinadas
                    st.session_state["relatorio_importacao"] = [{
                        "entidade": chave,
                        "aba_origem": uploaded_file.name,
                        "confianca": "alta" if origem == "identificação automática" else "manual",
                        "colunas_reconhecidas": [c for c in definicao["colunas"] if c in df_preparado.columns],
                        "colunas_novas": [c for c in df_preparado.columns if c not in definicao["colunas"]],
                        "colunas_ignoradas": [],
                        "linhas_lidas": len(df_preparado),
                    }]
                    st.session_state["estatisticas_importacao"] = estatisticas
                    st.toast(
                        f"✅ CSV incorporado ({origem}): {estatisticas[[x['entidade'] for x in estatisticas].index(chave)]['adicionados']} novos / "
                        f"{estatisticas[[x['entidade'] for x in estatisticas].index(chave)]['atualizados']} atualizados.",
                        icon="✅",
                    )
                    st.rerun()
        except Exception as e:
            st.error(f"❌ Erro ao interpretar/incorporar o arquivo: {e}")

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

    if st.session_state.get('estatisticas_importacao'):
        with st.expander("📊 Resumo da última integração", expanded=False):
            for item in st.session_state["estatisticas_importacao"]:
                if item["atualizados"] or item["adicionados"] or item["novas_colunas"]:
                    st.write(
                        f"**{item['entidade']}** — "
                        f"{item['adicionados']} novos, {item['atualizados']} atualizados"
                        + (f" | novas colunas: {', '.join(item['novas_colunas'])}" if item["novas_colunas"] else "")
                    )

    st.divider()

    bases = st.session_state['bases']
    df_base_raw = construir_base_unificada(bases["lojas"], bases["fila"], bases["inaug"])
    df_instrutores = bases["instrutores"]
    df_rec_raw = bases["rec"]

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

    st.markdown("🎯 **Filtros Globais**")
    uf_opcoes = ["Todas"] + sorted([str(x) for x in df_base_raw['UF'].dropna().unique()]) if 'UF' in df_base_raw.columns else ["Todas"]
    filtro_uf = st.selectbox("Filtrar Estado (UF):", uf_opcoes)
    cf_opcoes = ["Todos"] + sorted([str(x) for x in df_base_raw['CF'].dropna().unique()]) if 'CF' in df_base_raw.columns else ["Todos"]
    filtro_cf = st.selectbox("Filtrar Consultor (CF):", cf_opcoes)
    st.divider()
    st.markdown(f"""
        <div class="sidebar-metric">📶 Status: <b>Operacional</b> 🟢</div>
        <div class="sidebar-metric">🏪 Rede total: <b>{len(df_base_raw)} unidades</b></div>
    """, unsafe_allow_html=True)

# APLICAÇÃO DOS FILTROS GLOBAIS
df_base = df_base_raw.copy()
if filtro_uf != "Todas":
    df_base = df_base[df_base['UF'] == filtro_uf]
if filtro_cf != "Todos":
    df_base = df_base[df_base['CF'] == filtro_cf]

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

# --- MÓDULOS DA APLICAÇÃO ---

if modulo == "📊 Dashboard Executivo":
    render_section_header("📊", "Dashboard Executivo", "Panorama consolidado da operação")
    if not df_base.empty:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-header"><span class="kpi-title">Rede Filtrada</span><span class="kpi-icon-circle">🏪</span></div>
                    <div class="kpi-value">{len(df_base)}</div>
                    <div class="kpi-footer">unidades na seleção</div>
                </div>
            """, unsafe_allow_html=True)
        with c2:
            pendentes = len(df_base[df_base['Tipo_Necessidade'] != 'Rede Ativa (Sem Pendência)']) if 'Tipo_Necessidade' in df_base.columns else 0
            st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-header"><span class="kpi-title">Fila Treinamento</span><span class="kpi-icon-circle">🎓</span></div>
                    <div class="kpi-value">{pendentes}</div>
                    <div class="kpi-footer">lojas com pendência</div>
                </div>
            """, unsafe_allow_html=True)
        with c3:
            a_contatar = len(df_base[df_base['Status_Contato'] == 'A Contatar']) if 'Status_Contato' in df_base.columns else 0
            st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-header"><span class="kpi-title">Pendentes Contato</span><span class="kpi-icon-circle">📞</span></div>
                    <div class="kpi-value">{a_contatar}</div>
                    <div class="kpi-footer">aguardando contato</div>
                </div>
            """, unsafe_allow_html=True)
        with c4:
            inaug = len(df_base[df_base['Previsão Inauguração'].notna()]) if 'Previsão Inauguração' in df_base.columns else 0
            st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-header"><span class="kpi-title">Inaugurações</span><span class="kpi-icon-circle">🚀</span></div>
                    <div class="kpi-value">{inaug}</div>
                    <div class="kpi-footer">com previsão de abertura</div>
                </div>
            """, unsafe_allow_html=True)

        st.divider()
        col_A, col_B = st.columns(2)
        with col_A:
            render_section_header("🗺️", "Concentração por Estado", "Top 10 UFs")
            if 'UF' in df_base.columns:
                st.bar_chart(df_base['UF'].value_counts().head(10), color="#FF9800")
        with col_B:
            render_section_header("📶", "Situação dos Contatos", "Distribuição no Call Center")
            if 'Status_Contato' in df_base.columns:
                st.bar_chart(df_base['Status_Contato'].value_counts(), color="#3B9EFF")

elif modulo == "📋 Pipeline AmPm":
    render_section_header("📋", "Pipeline AmPm", "Fluxo operacional de treinamentos")
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
                with st.expander(f"📍 PV {item.get('PV Abadi', '-')} | {str(item.get('Razão Social', ''))[:12]}..."):
                    st.write(f"**Cidade:** {item.get('Municipio', '-')}/{item.get('UF', '-')}")
                    st.write(f"**Necessidade:** {item.get('Tipo_Necessidade', '-')}")
                    mudar_status = st.selectbox(
                        "Alterar Status:", colunas_pipeline,
                        index=colunas_pipeline.index(status),
                        key=f"pipe_sel_{item.get('PV Abadi')}"
                    )
                    if mudar_status != status:
                        atualizar_fila(item['PV Abadi'], {'Status_Contato': mudar_status})
                        st.success("Atualizado!")
                        st.rerun()

elif modulo == "🔍 PROCV & Filtros Avançados":
    render_section_header("🔍", "PROCV & Filtros Avançados", "Consulta detalhada na rede")
    if not df_base.empty:
        with st.expander("🔎 **Pesquisa Avançada**", expanded=True):
            f1, f2 = st.columns(2)
            termo = f1.text_input("🔍 PV, Nome ou Município:", "")
            f_necessidade = f2.selectbox("🎯 Necessidade:", ["Todas"] + sorted([str(x) for x in df_base['Tipo_Necessidade'].dropna().unique()])) if 'Tipo_Necessidade' in df_base.columns else ["Todas"]

        df_view = df_base.copy()
        if termo:
            df_view = df_view[
                df_view.get('Razão Social', pd.Series(dtype=object)).astype(str).str.contains(termo, case=False, na=False) |
                df_view.get('PV Abadi', pd.Series(dtype=object)).astype(str).str.contains(termo, na=False) |
                df_view.get('Municipio', pd.Series(dtype=object)).astype(str).str.contains(termo, case=False, na=False)
            ]
        if f_necessidade != "Todas" and 'Tipo_Necessidade' in df_view.columns:
            df_view = df_view[df_view['Tipo_Necessidade'] == f_necessidade]

        cols_mostrar = [c for c in ['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Status Loja', 'Tipo_Necessidade', 'Status_Contato'] if c in df_view.columns]
        evento = st.dataframe(df_view[cols_mostrar], use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun")

        linhas = evento.selection.get("rows", [])
        if linhas:
            p = df_view.iloc[linhas[0]].to_dict()
            st.divider()
            k1, k2, k3 = st.columns(3)
            with k1:
                st.markdown(f"""
                    <div class="procv-card">
                        <h4>🏪 Cadastro</h4>
                        <p>📍 <b>Endereço:</b> {p.get('Endereço', '-')}</p>
                        <p>🏙️ <b>Cidade/UF:</b> {p.get('Municipio', '-')}/{p.get('UF', '-')}</p>
                    </div>
                """, unsafe_allow_html=True)
            with k2:
                st.markdown(f"""
                    <div class="procv-card">
                        <h4>👔 Gestão</h4>
                        <p>👤 <b>Gerente:</b> {p.get('GF', '-')}</p>
                        <p>👔 <b>Consultor:</b> {p.get('CF', '-')}</p>
                    </div>
                """, unsafe_allow_html=True)
            with k3:
                st.markdown(f"""
                    <div class="procv-card">
                        <h4>📞 Atendimento</h4>
                        <p>🎯 <b>Necessidade:</b> {p.get('Tipo_Necessidade', '-')}</p>
                        <p>🔄 <b>Status:</b> {badge_status_html(p.get('Status_Contato', '-'))}</p>
                    </div>
                """, unsafe_allow_html=True)

elif modulo == "📍 Calculadora & Otimizador de Custos":
    render_section_header("📍", "Calculadora & Otimizador de Custos", "Simulação de rotas e análise de custos")
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
                primeira = top_3.iloc[0]
                if pd.notna(primeira.get('Lat_Loja')) and pd.notna(primeira.get('Lon_Loja')) and pd.notna(primeira.get('Lat_Instrutor')) and pd.notna(primeira.get('Lon_Instrutor')):
                    p_lat, p_lon = float(primeira['Lat_Loja']), float(primeira['Lon_Loja'])
                    i_lat, i_lon = float(primeira['Lat_Instrutor']), float(primeira['Lon_Instrutor'])

                    df_mapa_pontos = pd.DataFrame([
                        {"name": f"Posto {primeira['PV_ABADI']}", "lat": p_lat, "lon": p_lon, "color": [226, 123, 0, 220]},
                        {"name": f"Instrutor {primeira['Instrutor_Sugerido']}", "lat": i_lat, "lon": i_lon, "color": [76, 175, 80, 220]}
                    ])
                    df_mapa_arco = pd.DataFrame([{"from_lat": i_lat, "from_lon": i_lon, "to_lat": p_lat, "to_lon": p_lon}])

                    layer_pontos = pdk.Layer("ScatterplotLayer", df_mapa_pontos, get_position="[lon, lat]", get_color="color", get_radius=20000, pickable=True)
                    layer_arco = pdk.Layer("ArcLayer", df_mapa_arco, get_source_position="[from_lon, from_lat]", get_target_position="[to_lon, to_lat]", get_source_color=[76, 175, 80, 180], get_target_color=[226, 123, 0, 180], get_width=4)
                    view_state = pdk.ViewState(latitude=(p_lat + i_lat) / 2, longitude=(p_lon + i_lon) / 2, zoom=5, pitch=40)

                    st.pydeck_chart(pdk.Deck(layers=[layer_pontos, layer_arco], initial_view_state=view_state, tooltip={"text": "{name}"}))

                col1, col2, col3 = st.columns(3)
                cols = [col1, col2, col3]
                for idx, (_, row) in enumerate(top_3.iterrows()):
                    dist = row['Distancia_km_linha_reta']
                    dias = row['Dias_Treinamento_Necessarios']
                    custo_total = (dist * 2 * 2.10) + (dias * 280.0)
                    with cols[idx]:
                        st.markdown(f"""
                            <div class="top-instructor-card">
                                <h4>#{idx+1} {row['Instrutor_Sugerido']}</h4>
                                <p>Origem: {row['Cidade_Instrutor']}/{row['UF_Instrutor']}</p>
                                <p>Distância: {dist} km</p>
                                <h3>Total Est.: R$ {custo_total:.2f}</h3>
                            </div>
                        """, unsafe_allow_html=True)

elif modulo == "📞 Call Center & Timeline WhatsApp":
    render_section_header("📞", "Call Center & Timeline WhatsApp", "Atendimentos e disparos de mensagem")
    if not df_base.empty:
        df_fila_view = df_base[df_base['Tipo_Necessidade'] != 'Rede Ativa (Sem Pendência)'].copy() if 'Tipo_Necessidade' in df_base.columns else df_base.copy()
        c_left, c_right = st.columns([1.2, 1.8])

        with c_left:
            cols_call = [c for c in ['PV Abadi', 'Razão Social', 'Municipio', 'UF', 'Status_Contato'] if c in df_fila_view.columns]
            evento_call = st.dataframe(df_fila_view[cols_call], use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun")
            selecionado = evento_call.selection.get("rows", [])

        with c_right:
            if selecionado:
                posto = df_fila_view.iloc[selecionado[0]]
                pv_alvo = posto.get('PV Abadi')
                tel_limpo = ''.join(filter(str.isdigit, str(posto.get('Telefone_Contato', ''))))

                st.markdown(f"### Ficha PV {pv_alvo} - {posto.get('Razão Social', '')}")
                if tel_limpo:
                    msg = f"Olá! Aqui é da Capacitação AmPm referente ao posto {posto.get('Razão Social', '')}."
                    link_wa = f"https://wa.me/55{tel_limpo}?text={msg.replace(' ', '%20')}"
                    st.markdown(f"👉 **[Chamar no WhatsApp]({link_wa})**")

                with st.form("form_callcenter_editavel"):
                    nome_c = st.text_input("Responsável:", value=str(posto.get('Nome_Contato', '') or ''))
                    tel_c = st.text_input("Telefone:", value=str(posto.get('Telefone_Contato', '') or ''))
                    obs = st.text_area("Observações:", value=str(posto.get('Observacoes', '') or ''))
                    novo_st = st.selectbox("Status:", ["A Contatar", "Em Negociação", "Agendado", "Treinamento Realizado", "Recusado"])

                    if st.form_submit_button("💾 Salvar Registro"):
                        atualizar_fila(pv_alvo, {
                            'Nome_Contato': nome_c,
                            'Telefone_Contato': tel_c,
                            'Status_Contato': novo_st,
                            'Observacoes': obs,
                            'Data_do_Contato': datetime.today().strftime('%d/%m/%Y %H:%M'),
                        })
                        st.success("✅ Salvo!")
                        st.rerun()

elif modulo == "👔 Equipe de Instrutores":
    render_section_header("👔", "Equipe de Instrutores", "Instrutores credenciados")
    if not df_instrutores.empty:
        st.dataframe(df_instrutores, use_container_width=True, hide_index=True)

elif modulo == "📇 Enriquecimento de Rede":
    render_section_header("📇", "Enriquecimento de Rede", "Atualizações de lojas e telefones")
    st.info("Utilize a barra lateral para fazer upload de novas bases ou enriquecer os dados existentes.")

elif modulo == "📂 Relatórios & Exportação":
    render_section_header("📂", "Relatórios & Exportação", "Download das bases atualizadas")
    csv_buffer = df_base.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📄 Baixar Base em CSV",
        data=csv_buffer,
        file_name=f"Base_CRM_AmPm_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv"
    )

