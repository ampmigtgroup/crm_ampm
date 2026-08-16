import streamlit as st
from datetime import date, datetime, timezone
from uuid import UUID, uuid4
from supabase import create_client
import re

# ============================================================
# FORMULÁRIO CLIENTE AMPM / IGT — V4 SUPABASE
# Formulário operacional com token individual, Supabase e fotos privadas.
# ============================================================

st.set_page_config(
    page_title="Solicitação de Treinamento | AmPm",
    page_icon="📝",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ------------------------------------------------------------
# ESTILO — MOBILE FIRST / PREMIUM / COMPACTO
# ------------------------------------------------------------
st.markdown(
    """
    <style>
        #MainMenu, footer, header, [data-testid="stToolbar"],
        [data-testid="stDecoration"], [data-testid="stStatusWidget"] {
            display: none !important;
        }

        html, body, [class*="css"] {
            font-family: Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        }

        .stApp {
            background:
                radial-gradient(circle at top right, rgba(255,171,0,.08), transparent 28%),
                #f6f7f9;
        }

        .block-container {
            max-width: 760px;
            padding-top: .7rem;
            padding-bottom: 2.5rem;
        }

        .brandbar {
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:12px;
            margin: 4px 2px 10px 2px;
            color:#343434;
            font-weight:800;
            letter-spacing:.02em;
        }

        .brand-ampm {
            font-size:1.15rem;
            color:#f57c00;
        }

        .brand-igt {
            font-size:.78rem;
            color:#555;
            text-transform:uppercase;
            letter-spacing:.13em;
        }

        .hero {
            position:relative;
            overflow:hidden;
            background:linear-gradient(135deg,#f36f21 0%,#ff9800 48%,#ffc107 100%);
            border-radius:24px;
            padding:24px 24px 22px 24px;
            color:#fff;
            box-shadow:0 16px 38px rgba(65,55,40,.16);
            margin-bottom:12px;
        }

        .hero:after {
            content:"";
            position:absolute;
            width:170px;
            height:170px;
            border-radius:50%;
            right:-65px;
            top:-75px;
            background:rgba(255,255,255,.14);
        }

        .hero-kicker {
            font-size:.72rem;
            font-weight:800;
            letter-spacing:.12em;
            text-transform:uppercase;
            opacity:.94;
            margin-bottom:7px;
        }

        .hero h1 {
            margin:0;
            font-size:1.7rem;
            line-height:1.08;
            font-weight:850;
        }

        .hero p {
            margin:8px 0 0 0;
            max-width:560px;
            font-size:.93rem;
            line-height:1.45;
            opacity:.96;
        }

        .progress-wrap {
            background:#fff;
            border:1px solid #eceff2;
            border-radius:16px;
            padding:12px 14px;
            margin:0 0 12px 0;
            box-shadow:0 5px 18px rgba(0,0,0,.035);
        }

        .progress-label {
            display:flex;
            justify-content:space-between;
            align-items:center;
            color:#5e6268;
            font-size:.75rem;
            margin-bottom:7px;
        }

        .progress-track {
            width:100%;
            height:7px;
            border-radius:99px;
            background:#eceff2;
            overflow:hidden;
        }

        .progress-fill {
            width:100%;
            height:100%;
            border-radius:99px;
            background:linear-gradient(90deg,#ff7a00,#ffc107);
        }

        .section-head {
            display:flex;
            align-items:flex-start;
            gap:10px;
            margin:2px 0 11px 0;
        }

        .section-number {
            width:30px;
            height:30px;
            min-width:30px;
            display:flex;
            align-items:center;
            justify-content:center;
            border-radius:10px;
            background:#fff2e1;
            color:#e96b00;
            font-weight:850;
            font-size:.82rem;
            border:1px solid #ffe1b8;
        }

        .section-title {
            color:#25272a;
            font-size:1.02rem;
            font-weight:850;
            line-height:1.15;
            margin:0;
        }

        .section-subtitle {
            color:#747980;
            font-size:.80rem;
            line-height:1.35;
            margin-top:3px;
        }

        .section-card {
            background:#fff;
            border:1px solid #eaedf0;
            border-radius:18px;
            padding:16px 17px 12px 17px;
            margin:0 0 11px 0;
            box-shadow:0 7px 22px rgba(0,0,0,.045);
        }

        .station-card {
            background:linear-gradient(180deg,#fff9ef,#fff);
            border:1px solid #ffe0b2;
            border-radius:14px;
            padding:13px 14px;
            margin:2px 0 12px 0;
            line-height:1.45;
            color:#4b4b4b;
            font-size:.88rem;
        }

        .station-card .name {
            font-weight:850;
            color:#2b2b2b;
            font-size:.94rem;
            margin-bottom:2px;
        }

        .chip {
            display:inline-block;
            padding:4px 8px;
            background:#f1f4f8;
            color:#53606e;
            border-radius:99px;
            font-size:.70rem;
            font-weight:750;
            margin:0 5px 5px 0;
        }

        .info-box {
            background:#eef6ff;
            border:1px solid #d9eaff;
            border-radius:13px;
            padding:10px 12px;
            color:#36526d;
            font-size:.82rem;
            line-height:1.4;
            margin:8px 0 4px 0;
        }

        .success-box {
            background:#effaf3;
            border:1px solid #cdebd7;
            border-radius:16px;
            padding:16px;
            color:#28563a;
            margin:12px 0;
        }

        .required-note {
            color:#777d84;
            font-size:.76rem;
            margin:3px 2px 12px 2px;
        }

        div[data-baseweb="input"] > div,
        div[data-baseweb="select"] > div,
        textarea {
            border-radius:12px !important;
        }

        div[data-testid="stFileUploader"] section {
            border-radius:14px;
            border-style:dashed;
            background:#fbfbfb;
        }

        div[data-testid="stButton"] > button,
        div[data-testid="stDownloadButton"] > button {
            border-radius:14px !important;
            min-height:49px;
            font-weight:800;
            box-shadow:none !important;
        }

        button[kind="primary"] {
            background:linear-gradient(90deg,#f36f21,#ff9800) !important;
            border:none !important;
        }

        .stRadio > label,
        .stSelectbox > label,
        .stTextInput > label,
        .stTextArea > label,
        .stNumberInput > label,
        .stDateInput > label,
        .stFileUploader > label {
            font-size:.86rem !important;
            font-weight:700 !important;
            color:#404449 !important;
        }

        div[data-testid="stMarkdownContainer"] p {
            line-height:1.4;
        }

        @media (max-width: 640px) {
            .block-container {
                padding-left:.72rem;
                padding-right:.72rem;
                padding-top:.45rem;
            }

            .hero {
                border-radius:18px;
                padding:19px 17px 18px 17px;
            }

            .hero h1 {
                font-size:1.38rem;
            }

            .hero p {
                font-size:.86rem;
            }

            .section-card {
                border-radius:15px;
                padding:14px 13px 10px 13px;
            }

            .brandbar {
                margin-left:2px;
                margin-right:2px;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def qp(nome, padrao=""):
    try:
        valor = st.query_params.get(nome, padrao)
        if isinstance(valor, list):
            return valor[0] if valor else padrao
        return str(valor or padrao)
    except Exception:
        return padrao


def somente_digitos(valor):
    return re.sub(r"\D+", "", str(valor or ""))


def validar_email(valor):
    valor = str(valor or "").strip()
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", valor))


def validar_telefone(valor):
    return len(somente_digitos(valor)) >= 10


def validar_cnpj_basico(valor):
    digitos = somente_digitos(valor)
    return len(digitos) == 14


def preenchido(valor):
    return bool(str(valor or "").strip())


def secao(numero, titulo, subtitulo):
    st.markdown(
        f"""
        <div class="section-head">
            <div class="section-number">{numero}</div>
            <div>
                <div class="section-title">{titulo}</div>
                <div class="section-subtitle">{subtitulo}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def resetar_formulario():
    preservar = {"token"}
    for chave in list(st.session_state.keys()):
        if chave not in preservar:
            del st.session_state[chave]


# ------------------------------------------------------------
# SUPABASE + TOKEN DO CLIENTE
# ------------------------------------------------------------
SUPABASE_PROJECT_URL_PADRAO = "https://nptazzfvwhhmotfrvgdj.supabase.co"
BUCKET_FOTOS = "crm-form-fotos"


@st.cache_resource(show_spinner=False)
def supabase_client():
    """
    Usa os mesmos nomes de Secrets aceitos pelo CRM principal.
    Compatível com:
    - SUPABASE_SERVICE_ROLE_KEY
    - SUPABASE_SECRET_KEY
    """
    try:
        url = str(
            st.secrets.get("SUPABASE_URL", SUPABASE_PROJECT_URL_PADRAO)
            or SUPABASE_PROJECT_URL_PADRAO
        ).strip()

        chave = (
            st.secrets.get("SUPABASE_SERVICE_ROLE_KEY")
            or st.secrets.get("SUPABASE_SECRET_KEY")
        )

        if not chave:
            return None

        return create_client(url, str(chave).strip())
    except Exception:
        return None


def token_uuid_valido(valor):
    try:
        UUID(str(valor))
        return True
    except Exception:
        return False


def carregar_token_formulario(token):
    client = supabase_client()
    if client is None:
        return None, "A conexão segura com o formulário ainda não foi configurada."

    if not token or not token_uuid_valido(token):
        return None, "Link de formulário inválido."

    try:
        resp = (
            client.table("crm_form_tokens")
            .select("*")
            .eq("token", str(token))
            .limit(1)
            .execute()
        )
        rows = resp.data or []
        if not rows:
            return None, "Este link não foi encontrado ou não é mais válido."

        row = rows[0]
        status = str(row.get("status", "") or "").strip().lower()

        if status == "respondido":
            return row, "respondido"
        if status == "cancelado":
            return None, "Este formulário foi cancelado."
        if status == "expirado":
            return None, "Este link expirou."

        expires_at = row.get("expires_at")
        if expires_at:
            try:
                dt_exp = datetime.fromisoformat(str(expires_at).replace("Z", "+00:00"))
                if dt_exp.tzinfo is None:
                    dt_exp = dt_exp.replace(tzinfo=timezone.utc)
                if datetime.now(timezone.utc) > dt_exp:
                    try:
                        (
                            client.table("crm_form_tokens")
                            .update({"status": "expirado", "updated_at": datetime.now(timezone.utc).isoformat()})
                            .eq("id", row["id"])
                            .execute()
                        )
                    except Exception:
                        pass
                    return None, "Este link expirou."
            except Exception:
                pass

        if status != "ativo":
            return None, "Este formulário não está disponível."

        return row, None
    except Exception:
        return None, "Não foi possível validar o formulário neste momento."


def nome_arquivo_seguro(nome):
    nome = re.sub(r"[^A-Za-z0-9._-]+", "_", str(nome or "foto"))
    return nome[:120] or "foto"


def upload_fotos_resposta(client, token_id, resposta_id, arquivos):
    enviados = []

    for foto in arquivos or []:
        if getattr(foto, "size", 0) > 6 * 1024 * 1024:
            raise ValueError(f"A foto '{foto.name}' excede o limite de 6 MB.")

        ext_nome = nome_arquivo_seguro(foto.name)
        caminho = f"{token_id}/{resposta_id}/{uuid4().hex}_{ext_nome}"
        tipo = str(getattr(foto, "type", "") or "image/jpeg")

        client.storage.from_(BUCKET_FOTOS).upload(
            path=caminho,
            file=foto.getvalue(),
            file_options={
                "content-type": tipo,
                "upsert": "false",
            },
        )

        enviados.append({
            "nome": foto.name,
            "tipo": tipo,
            "tamanho_bytes": int(getattr(foto, "size", 0) or 0),
            "storage_path": caminho,
        })

    return enviados


token_interno = qp("token").strip()
token_registro, token_erro = carregar_token_formulario(token_interno)

# Um formulário público nunca deve funcionar sem um token válido.
if token_erro == "respondido":
    st.markdown(
        """
        <div class="success-box">
            <strong>✓ Formulário já enviado</strong><br>
            Esta solicitação já foi respondida. Não é necessário preencher novamente.
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

if token_registro is None:
    st.error(token_erro or "Link inválido.")
    st.info("Abra o formulário pelo link recebido da equipe AmPm/IGT.")
    st.stop()

rs_link = str(token_registro.get("razao_social", "") or "")
cnpj_link = str(token_registro.get("cnpj", "") or "")
municipio_link = str(token_registro.get("municipio", "") or "")
uf_link = str(token_registro.get("uf", "") or "").upper()
modelo_link = str(token_registro.get("modelo", "") or "")
pv_interno = str(token_registro.get("pv_abadi", "") or "")
ufs = [
    "AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG",
    "PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"
]

# ------------------------------------------------------------
# TOPO
# ------------------------------------------------------------
st.markdown(
    """
    <div class="brandbar">
        <div class="brand-ampm">ampm☀</div>
        <div class="brand-igt">IGT GROUP</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
        <div class="hero-kicker">Solicitação de treinamento</div>
        <h1>Vamos preparar sua unidade.</h1>
        <p>Leva poucos minutos. As informações abaixo ajudam nossa equipe a planejar o treinamento com mais precisão.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="progress-wrap">
        <div class="progress-label">
            <span>Formulário completo</span>
            <span>7 etapas</span>
        </div>
        <div class="progress-track"><div class="progress-fill"></div></div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="required-note">* Campos obrigatórios</div>', unsafe_allow_html=True)

# ============================================================
# 1. IDENTIFICAÇÃO
# ============================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
secao("01", "Identificação do posto", "Confira os dados da unidade antes de continuar.")

if rs_link or cnpj_link or municipio_link or uf_link:
    st.markdown(
        f"""
        <div class="station-card">
            <div class="name">{rs_link or "Razão Social não informada"}</div>
            <span class="chip">CNPJ {cnpj_link or "—"}</span>
            <span class="chip">{municipio_link or "Município"} {("- " + uf_link) if uf_link else ""}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

cadastro_correto = st.radio(
    "Essas informações estão corretas? *",
    ["Sim", "Não"],
    horizontal=True,
    key="cadastro_correto",
)

editar_cadastro = cadastro_correto == "Não" or not (
    rs_link and cnpj_link and municipio_link and uf_link
)

razao_social = st.text_input(
    "Razão Social *",
    value=rs_link,
    disabled=not editar_cadastro,
    key="razao_social_form",
)

cnpj = st.text_input(
    "CNPJ *",
    value=cnpj_link,
    disabled=not editar_cadastro,
    placeholder="00.000.000/0000-00",
    key="cnpj_form",
)

c1, c2 = st.columns([2, 1])
with c1:
    municipio = st.text_input(
        "Município *",
        value=municipio_link,
        disabled=not editar_cadastro,
        key="municipio_form",
    )
with c2:
    if editar_cadastro:
        uf_index = ufs.index(uf_link) if uf_link in ufs else 0
        uf = st.selectbox("UF *", ufs, index=uf_index, key="uf_form_editavel")
    else:
        uf = st.text_input("UF *", value=uf_link, disabled=True, key="uf_form_bloqueada")

modelos = ["Cafeteria", "Padaria", "Pizza Hut"]
modelo_default = modelo_link if modelo_link in modelos else modelos[0]
tipo_modelo = st.selectbox(
    "Tipo/modelo do posto *",
    modelos,
    index=modelos.index(modelo_default),
    key="tipo_modelo_form",
)

if cadastro_correto == "Não":
    st.markdown(
        '<div class="info-box">As correções ficarão identificadas na resposta para futura atualização do cadastro.</div>',
        unsafe_allow_html=True,
    )

st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# 2. CONTATO
# ============================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
secao("02", "Contato", "Informe quem está preenchendo e será nossa referência para este atendimento.")

nome_responsavel = st.text_input(
    "Nome do responsável *",
    placeholder="Nome e sobrenome",
    key="nome_responsavel",
)
telefone = st.text_input(
    "Telefone / WhatsApp *",
    placeholder="(00) 00000-0000",
    key="telefone_responsavel",
)
email = st.text_input(
    "E-mail *",
    placeholder="nome@empresa.com",
    key="email_responsavel",
)
cargo = st.text_input(
    "Cargo / Função *",
    placeholder="Ex.: Gerente, franqueado, supervisor",
    key="cargo_responsavel",
)
st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# 3. EQUIPE
# ============================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
secao("03", "Equipe para treinamento", "Precisamos entender o tamanho e o perfil básico da equipe.")

tem_funcionarios = st.radio(
    "Há funcionários para treinar? *",
    ["Sim", "Não"],
    horizontal=True,
    key="tem_funcionarios",
)

qtd_funcionarios = 0
turno = None

if tem_funcionarios == "Sim":
    c1, c2 = st.columns([1, 1.4])
    with c1:
        qtd_funcionarios = st.number_input(
            "Quantidade *",
            min_value=1,
            step=1,
            value=1,
            key="qtd_funcionarios",
        )
    with c2:
        turno = st.selectbox(
            "Turno predominante *",
            ["Manhã", "Tarde", "Noite", "Misto"],
            key="turno_predominante",
        )
else:
    st.markdown(
        '<div class="info-box">Sem funcionários para treinar: a quantidade será registrada como 0.</div>',
        unsafe_allow_html=True,
    )

treinamento_anterior = st.radio(
    "A equipe já recebeu treinamento anteriormente? *",
    ["Sim", "Não", "Não sabe"],
    horizontal=True,
    key="treinamento_anterior",
)
st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# 4. TREINAMENTO
# ============================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
secao("04", "Treinamento", "Informe o motivo, a modalidade e sua preferência de data.")

c1, c2 = st.columns(2)
with c1:
    motivo_solicitacao = st.selectbox(
        "Motivo da solicitação *",
        ["Inauguração", "Retreinamento"],
        key="motivo_solicitacao",
    )
with c2:
    tipo_treinamento = st.selectbox(
        "Tipo de treinamento *",
        ["Cafeteria", "Padaria", "Pizza Hut"],
        key="tipo_treinamento",
    )

hoje = date.today()
data_desejada = st.date_input(
    "Data desejada *",
    min_value=hoje,
    value=hoje,
    key="data_desejada",
    help="A data será considerada como preferência e ficará sujeita à disponibilidade da equipe.",
)

deseja_segunda_data = st.checkbox(
    "Quero informar uma segunda opção de data",
    key="usar_segunda_data",
)

segunda_data = None
if deseja_segunda_data:
    segunda_data = st.date_input(
        "Segunda opção de data",
        min_value=hoje,
        value=hoje,
        key="segunda_data",
    )

melhor_periodo = st.radio(
    "Melhor período *",
    ["Manhã", "Tarde", "Noite"],
    horizontal=True,
    key="melhor_periodo",
)

motivo_necessidade = st.text_area(
    "Motivo da necessidade de treinamento *",
    placeholder="Ex.: nova equipe, baixa performance, atualização de processos...",
    height=105,
    max_chars=500,
    key="motivo_necessidade",
)
st.caption(f"{len(motivo_necessidade or '')}/500 caracteres")
st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# 5. ESTRUTURA
# ============================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
secao("05", "Estrutura da loja", "Confirme se o ambiente já está preparado para o treinamento.")

material_loja = st.radio(
    "O material necessário já está na loja? *",
    ["Sim", "Não", "Parcialmente"],
    horizontal=True,
    key="material_loja",
)

equipamentos_instalados = st.radio(
    "Os equipamentos estão instalados? *",
    ["Sim", "Não", "Parcialmente"],
    horizontal=True,
    key="equipamentos_instalados",
)

equipamentos_funcionando = st.radio(
    "Os equipamentos estão funcionando corretamente? *",
    ["Sim", "Não", "Não sabe"],
    horizontal=True,
    key="equipamentos_funcionando",
)

loja_pronta = st.radio(
    "A loja está pronta para receber o treinamento? *",
    ["Sim", "Não"],
    horizontal=True,
    key="loja_pronta",
)

tem_pendencia = st.radio(
    "Existe alguma pendência que possa impedir o treinamento? *",
    ["Não", "Sim"],
    horizontal=True,
    key="tem_pendencia",
)

descricao_pendencia = ""
if tem_pendencia == "Sim":
    descricao_pendencia = st.text_area(
        "Qual pendência? *",
        placeholder="Explique brevemente o que ainda precisa ser resolvido.",
        height=95,
        max_chars=500,
        key="descricao_pendencia",
    )
    st.caption(f"{len(descricao_pendencia or '')}/500 caracteres")

st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# 6. FOTOS
# ============================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
secao("06", "Fotos da estrutura/equipamentos", "Opcional. As imagens ajudam nossa equipe a entender melhor o cenário.")

fotos = st.file_uploader(
    "Adicionar fotos",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
    key="fotos_estrutura",
    help="Formatos aceitos: JPG, PNG e WEBP.",
)

if fotos:
    st.success(f"{len(fotos)} foto(s) selecionada(s).")
    preview = fotos[:4]
    cols = st.columns(min(len(preview), 2))
    for i, foto in enumerate(preview):
        with cols[i % len(cols)]:
            st.image(foto, caption=foto.name, use_container_width=True)

    if len(fotos) > 4:
        st.caption(f"+ {len(fotos) - 4} foto(s) selecionada(s) sem prévia.")

st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# 7. CONFIRMAÇÃO
# ============================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
secao("07", "Confirmação", "Revise as informações e confirme o envio.")

confirmo = st.checkbox(
    "Confirmo que as informações fornecidas estão corretas. *",
    key="confirmo_dados",
)
autorizo = st.checkbox(
    "Autorizo o uso dessas informações para organização do treinamento. *",
    key="autorizo_uso",
)

if nome_responsavel:
    st.markdown(
        f'<div class="info-box"><strong>Responsável:</strong> {nome_responsavel}<br>'
        'A data e a hora do envio serão registradas automaticamente.</div>',
        unsafe_allow_html=True,
    )

st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# VALIDAÇÃO
# ============================================================
erros = []

for rotulo, valor in [
    ("Razão Social", razao_social),
    ("CNPJ", cnpj),
    ("Município", municipio),
    ("UF", uf),
    ("Nome do responsável", nome_responsavel),
    ("Telefone", telefone),
    ("E-mail", email),
    ("Cargo/Função", cargo),
    ("Motivo da necessidade de treinamento", motivo_necessidade),
]:
    if not preenchido(valor):
        erros.append(f"Preencha: {rotulo}.")

if cnpj and not validar_cnpj_basico(cnpj):
    erros.append("Informe um CNPJ com 14 dígitos.")

if not validar_telefone(telefone):
    erros.append("Informe um telefone válido com DDD.")

if email and not validar_email(email):
    erros.append("Informe um e-mail válido.")

if str(uf or "").strip().upper() not in ufs:
    erros.append("Selecione uma UF válida.")

if tem_funcionarios == "Sim" and int(qtd_funcionarios or 0) < 1:
    erros.append("Informe a quantidade de funcionários.")

if tem_pendencia == "Sim" and not preenchido(descricao_pendencia):
    erros.append("Descreva a pendência que pode impedir o treinamento.")

if segunda_data and segunda_data == data_desejada:
    erros.append("A segunda opção de data deve ser diferente da data desejada.")

if not confirmo:
    erros.append("Confirme que as informações estão corretas.")

if not autorizo:
    erros.append("Autorize o uso das informações para organização do treinamento.")

# ------------------------------------------------------------
# AVISO DE PROTÓTIPO
# ------------------------------------------------------------
st.markdown(
    """
    <div class="info-box">
        <strong>Envio seguro.</strong> As respostas serão registradas para organização do treinamento.
        O link é individual e será encerrado depois do envio.
    </div>
    """,
    unsafe_allow_html=True,
)

# ------------------------------------------------------------
# ENVIO
# ------------------------------------------------------------
enviar = st.button(
    "Enviar solicitação",
    type="primary",
    use_container_width=True,
    key="enviar_formulario",
)

if enviar:
    if erros:
        st.error("Revise os campos destacados abaixo:")
        for erro in erros:
            st.write(f"• {erro}")
    else:
        client = supabase_client()

        if client is None:
            st.error("Não foi possível conectar ao serviço de envio. Tente novamente em alguns instantes.")
            st.stop()

        # Revalida o token imediatamente antes da gravação.
        token_atual, erro_token_atual = carregar_token_formulario(token_interno)
        if token_atual is None or erro_token_atual:
            if erro_token_atual == "respondido":
                st.info("Esta solicitação já foi enviada.")
            else:
                st.error(erro_token_atual or "Este link não está mais disponível.")
            st.stop()

        resposta_id = str(uuid4())
        enviado_em = datetime.now(timezone.utc).isoformat()

        try:
            with st.spinner("Enviando sua solicitação..."):
                fotos_enviadas = upload_fotos_resposta(
                    client,
                    token_atual["id"],
                    resposta_id,
                    fotos,
                )

                payload = {
                    "token": token_interno,
                    "pv_abadi": pv_interno or None,
                    "enviado_em": enviado_em,
                    "identificacao": {
                        "cadastro_confirmado": cadastro_correto == "Sim",
                        "razao_social": razao_social.strip(),
                        "cnpj": cnpj.strip(),
                        "municipio": municipio.strip(),
                        "uf": str(uf).strip().upper(),
                        "tipo_modelo": tipo_modelo,
                    },
                    "contato": {
                        "nome_responsavel": nome_responsavel.strip(),
                        "telefone": telefone.strip(),
                        "email": email.strip(),
                        "cargo_funcao": cargo.strip(),
                    },
                    "equipe": {
                        "tem_funcionarios": tem_funcionarios == "Sim",
                        "qtd_funcionarios": int(qtd_funcionarios or 0),
                        "turno_predominante": turno,
                        "treinamento_anterior": treinamento_anterior,
                    },
                    "treinamento": {
                        "motivo_solicitacao": motivo_solicitacao,
                        "tipo_treinamento": tipo_treinamento,
                        "data_desejada": data_desejada.isoformat(),
                        "segunda_opcao_data": segunda_data.isoformat() if segunda_data else None,
                        "melhor_periodo": melhor_periodo,
                        "motivo_necessidade": motivo_necessidade.strip(),
                    },
                    "estrutura": {
                        "material_loja": material_loja,
                        "equipamentos_instalados": equipamentos_instalados,
                        "equipamentos_funcionando": equipamentos_funcionando,
                        "loja_pronta": loja_pronta,
                        "tem_pendencia": tem_pendencia == "Sim",
                        "descricao_pendencia": descricao_pendencia.strip() if descricao_pendencia else "",
                    },
                    "fotos": fotos_enviadas,
                    "confirmacao": {
                        "confirmou_dados": bool(confirmo),
                        "autorizou_uso": bool(autorizo),
                        "nome_confirmacao": nome_responsavel.strip(),
                    },
                }

                registro = {
                    "id": resposta_id,
                    "token_id": token_atual["id"],
                    "pv_abadi": pv_interno or None,
                    "cadastro_confirmado": cadastro_correto == "Sim",
                    "razao_social": razao_social.strip(),
                    "cnpj": cnpj.strip(),
                    "municipio": municipio.strip(),
                    "uf": str(uf).strip().upper(),
                    "tipo_modelo": tipo_modelo,
                    "nome_responsavel": nome_responsavel.strip(),
                    "telefone": telefone.strip(),
                    "email": email.strip(),
                    "cargo_funcao": cargo.strip(),
                    "tem_funcionarios": tem_funcionarios == "Sim",
                    "qtd_funcionarios": int(qtd_funcionarios or 0),
                    "turno_predominante": turno,
                    "treinamento_anterior": treinamento_anterior,
                    "motivo_solicitacao": motivo_solicitacao,
                    "tipo_treinamento": tipo_treinamento,
                    "data_desejada": data_desejada.isoformat(),
                    "segunda_opcao_data": segunda_data.isoformat() if segunda_data else None,
                    "melhor_periodo": melhor_periodo,
                    "motivo_necessidade": motivo_necessidade.strip(),
                    "material_loja": material_loja,
                    "equipamentos_instalados": equipamentos_instalados,
                    "equipamentos_funcionando": equipamentos_funcionando,
                    "loja_pronta": loja_pronta == "Sim",
                    "tem_pendencia": tem_pendencia == "Sim",
                    "descricao_pendencia": descricao_pendencia.strip() if descricao_pendencia else None,
                    "fotos": fotos_enviadas,
                    "confirmou_dados": bool(confirmo),
                    "autorizou_uso": bool(autorizo),
                    "payload": payload,
                    "enviado_em": enviado_em,
                }

                client.table("crm_form_respostas").insert(registro).execute()

                (
                    client.table("crm_form_tokens")
                    .update({
                        "status": "respondido",
                        "respondido_em": enviado_em,
                        "updated_at": enviado_em,
                    })
                    .eq("id", token_atual["id"])
                    .eq("status", "ativo")
                    .execute()
                )

            st.session_state["form_envio_concluido"] = True
            st.markdown(
                f"""
                <div class="success-box">
                    <strong>✓ Solicitação enviada com sucesso</strong><br>
                    Obrigado, {nome_responsavel}. As informações foram registradas e serão utilizadas
                    para organização do treinamento.
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.info("Você já pode fechar esta página.")

        except Exception as e:
            mensagem = str(e)
            if "uq_crm_form_respostas_token_id" in mensagem or "duplicate key" in mensagem.lower():
                st.info("Esta solicitação já foi enviada anteriormente.")
            elif "excede o limite de 6 MB" in mensagem:
                st.error(mensagem)
            else:
                st.error(
                    "Não foi possível concluir o envio. Nenhuma alteração foi feita no CRM principal. "
                    "Tente novamente em alguns instantes."
                )

st.caption("ampm☀  •  IGT Group  |  Solicitação de treinamento")
