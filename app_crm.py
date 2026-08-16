import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
import pydeck as pdk
import io
import zipfile
import uuid
import time
import requests
import json
import hashlib
import re
import unicodedata
import html
from difflib import SequenceMatcher
import streamlit_authenticator as stauth
from supabase import create_client, Client

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="CRM Operacional AmPm",
    page_icon="⛽",
    layout="wide",
    initial_sidebar_state="expanded"
)

# V53 — interface limpa: oculta permanentemente controles nativos do
# Streamlit Cloud ligados a GitHub/source/deploy, inclusive na tela de login.
st.markdown(
    """
    <style>
        [data-testid="stToolbar"],
        [data-testid="stAppDeployButton"],
        [data-testid="stDecoration"],
        [data-testid="stStatusWidget"],
        a[href*="github.com"],
        button[aria-label*="GitHub"],
        button[title*="GitHub"],
        a[aria-label*="GitHub"],
        a[title*="GitHub"],
        #MainMenu {
            display: none !important;
            visibility: hidden !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

CAMINHO_ARQUIVO = "Base_Unificada_AmPm.xlsx"
CAMINHO_BACKUP = "Base_Unificada_AmPm.backup.xlsx"

# Banco central online. O Excel permanece apenas como fonte de importação/fallback.
SUPABASE_PROJECT_URL_PADRAO = "https://nptazzfvwhhmotfrvgdj.supabase.co"
SUPABASE_TABLES = {
    "lojas": "crm_lojas",
    "fila": "crm_fila_callcenter",
    "inaug": "crm_inauguracoes",
    "instrutores": "crm_instrutores",
    "rec": "crm_recomendacao_deslocamento",
}

CAMINHO_ORCAMENTOS = "orcamentos_crm.json"
PASTA_DOCUMENTOS_ORCAMENTO = "documentos_orcamentos"


def _carregar_orcamentos():
    if not os.path.exists(CAMINHO_ORCAMENTOS):
        return {}
    try:
        with open(CAMINHO_ORCAMENTOS, "r", encoding="utf-8") as arquivo:
            dados = json.load(arquivo)
            return dados if isinstance(dados, dict) else {}
    except Exception:
        return {}


def _salvar_orcamentos(dados):
    with open(CAMINHO_ORCAMENTOS, "w", encoding="utf-8") as arquivo:
        json.dump(dados, arquivo, ensure_ascii=False, indent=2, default=str)


def _chave_orcamento(pv):
    return str(pv).strip()


PRODUTOS_TREINAMENTO = [
    "Treinamento de Cafeteria",
    "Treinamento em Pizzaria Pizza Hut",
    "Treinamento em Padaria",
    "Deslocamento Aéreo",
    "Deslocamento Rodoviário",
    "Hospedagem",
    "Aluguel de Carro",
    "Transporte por Aplicativo",
]

UNIDADE_PADRAO_ORCAMENTO = {
    "Treinamento de Cafeteria": "Dia",
    "Treinamento em Pizzaria Pizza Hut": "Dia",
    "Treinamento em Padaria": "Dia",
    "Deslocamento Aéreo": "Trecho",
    "Deslocamento Rodoviário": "Trecho",
    "Hospedagem": "Diária",
    "Aluguel de Carro": "Diária",
    "Transporte por Aplicativo": "Corrida",
}

UNIDADES_ORCAMENTO = [
    "Dia",
    "Diária",
    "Trecho",
    "Corrida",
    "Unidade",
]


def _orcamento_vazio(posto):
    agora = datetime.now()
    return {
        "numero": f"ORC-{agora.strftime('%Y%m%d')}-{uuid.uuid4().hex[:6].upper()}",
        "titulo": f"Orçamento - {posto.get('Razão Social', '')}",
        "responsavel": "",
        "validade_dias": 7,
        "condicao_pagamento": "A definir",
        "observacoes": "",
        "desconto": 0.0,
        "itens": [],
        "documentos": [],
        "atualizado_em": agora.strftime("%d/%m/%Y %H:%M"),
    }


def _normalizar_itens_orcamento(itens):
    """
    Normaliza itens antigos e novos.
    Orçamentos antigos com Dias/Valor por Dia continuam funcionando.
    Novos itens usam Quantidade/Unidade/Valor Unitário.
    """
    linhas = []
    for item in itens or []:
        if not isinstance(item, dict):
            continue

        produto = str(
            item.get(
                "Produto",
                item.get("Descrição", item.get("descricao", ""))
            ) or ""
        ).strip()
        if not produto:
            continue

        try:
            quantidade = float(
                item.get(
                    "Quantidade",
                    item.get("Dias", item.get("Qtd", item.get("qtd", 1)))
                ) or 0
            )
        except (TypeError, ValueError):
            quantidade = 0.0

        try:
            valor_unitario = float(
                item.get(
                    "Valor Unitário",
                    item.get(
                        "Valor por Dia",
                        item.get("valor_unitario", 0)
                    )
                ) or 0
            )
        except (TypeError, ValueError):
            valor_unitario = 0.0

        unidade = str(
            item.get(
                "Unidade",
                UNIDADE_PADRAO_ORCAMENTO.get(produto, "Unidade")
            ) or UNIDADE_PADRAO_ORCAMENTO.get(produto, "Unidade")
        ).strip()

        total = quantidade * valor_unitario

        linhas.append({
            "Item": str(item.get("Item", len(linhas) + 1)),
            "Produto": produto,
            "Quantidade": quantidade,
            "Unidade": unidade,
            "Valor Unitário": valor_unitario,
            "Total": total,
        })

    return linhas


def _gerar_excel_orcamento(orcamento, posto, pv):
    itens = _normalizar_itens_orcamento(orcamento.get("itens", []))
    colunas = ["Item", "Produto", "Quantidade", "Unidade", "Valor Unitário", "Total"]
    df_itens = pd.DataFrame(itens, columns=colunas)
    if df_itens.empty:
        df_itens = pd.DataFrame(columns=colunas)

    subtotal = float(sum(float(x.get("Total", 0) or 0) for x in itens))
    desconto = float(orcamento.get("desconto", 0) or 0)
    desconto_valor = subtotal * (desconto / 100)
    total = subtotal - desconto_valor

    resumo = pd.DataFrame([
        ["Orçamento", orcamento.get("numero", "")],
        ["PV", pv],
        ["Razão Social", posto.get("Razão Social", "")],
        ["Cidade/UF", f"{posto.get('Municipio', '')}/{posto.get('UF', '')}"],
        ["Responsável", orcamento.get("responsavel", "")],
        ["Validade", f"{orcamento.get('validade_dias', 7)} dias"],
        ["Condição de pagamento", orcamento.get("condicao_pagamento", "")],
        ["Subtotal", subtotal],
        ["Desconto (%)", desconto],
        ["Desconto (R$)", desconto_valor],
        ["TOTAL", total],
        ["Observações", orcamento.get("observacoes", "")],
        ["Atualizado em", orcamento.get("atualizado_em", "")],
    ], columns=["Campo", "Informação"])

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        resumo.to_excel(writer, index=False, sheet_name="Orçamento")
        df_itens.to_excel(writer, index=False, sheet_name="Itens Orçamento")

        ws = writer.sheets["Orçamento"]
        ws.freeze_panes = "A2"
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 65

        ws2 = writer.sheets["Itens Orçamento"]
        ws2.freeze_panes = "A2"
        ws2.auto_filter.ref = ws2.dimensions
        for col in ws2.columns:
            letra = col[0].column_letter
            maior = max([len(str(c.value or "")) for c in col[:100]] + [10])
            ws2.column_dimensions[letra].width = min(maior + 2, 45)
        for linha in range(2, ws2.max_row + 1):
            ws2.cell(linha, 5).number_format = 'R$ #,##0.00'
            ws2.cell(linha, 6).number_format = 'R$ #,##0.00'

        for linha in range(1, ws.max_row + 1):
            if ws.cell(linha, 1).value in ("Subtotal", "Desconto (R$)", "TOTAL"):
                ws.cell(linha, 2).number_format = 'R$ #,##0.00'

    buffer.seek(0)
    return buffer.getvalue()


def _salvar_documentos_orcamento(pv, arquivos):
    pasta = os.path.join(PASTA_DOCUMENTOS_ORCAMENTO, _chave_orcamento(pv).replace("/", "_"))
    os.makedirs(pasta, exist_ok=True)
    salvos = []

    for arquivo in arquivos or []:
        nome_original = os.path.basename(arquivo.name)
        nome_seguro = re.sub(r"[^A-Za-z0-9._ -]", "_", nome_original).strip() or "documento"
        nome_final = f"{uuid.uuid4().hex[:8]}_{nome_seguro}"
        caminho = os.path.join(pasta, nome_final)
        with open(caminho, "wb") as destino:
            destino.write(arquivo.getbuffer())
        salvos.append({
            "nome": nome_original,
            "arquivo": caminho,
            "tamanho": os.path.getsize(caminho),
        })
    return salvos


def _criar_pacote_orcamento(orcamento, posto, pv):
    excel = _gerar_excel_orcamento(orcamento, posto, pv)
    pacote = io.BytesIO()
    nome_base = re.sub(r"[^A-Za-z0-9_-]", "_", str(orcamento.get("numero", "orcamento")))

    with zipfile.ZipFile(pacote, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"{nome_base}/Orcamento.xlsx", excel)

        resumo_txt = (
            f"ORÇAMENTO {orcamento.get('numero', '')}\n"
            f"PV: {pv}\n"
            f"Razão Social: {posto.get('Razão Social', '')}\n"
            f"Total: R$ {float(orcamento.get('_total_calculado', 0) or 0):,.2f}\n"
            f"Atualizado em: {orcamento.get('atualizado_em', '')}\n"
        )
        zf.writestr(f"{nome_base}/LEIA-ME.txt", resumo_txt)

        for doc in orcamento.get("documentos", []):
            caminho = doc.get("arquivo", "")
            if caminho and os.path.exists(caminho):
                zf.write(caminho, arcname=f"{nome_base}/Documentos/{doc.get('nome', os.path.basename(caminho))}")

    pacote.seek(0)
    return pacote.getvalue()


def _renderizar_mini_orcamento(posto, pv):
    """Mini aplicativo de orçamento vinculado ao cliente selecionado no Call Center."""
    st.markdown("### 💰 Gerador de Orçamento")
    st.caption(
        "Monte o orçamento, altere os itens, anexe documentos e exporte tudo em um pacote único."
    )

    if "orcamentos_crm" not in st.session_state:
        st.session_state["orcamentos_crm"] = _carregar_orcamentos()

    chave = _chave_orcamento(pv)
    orcamentos = st.session_state["orcamentos_crm"]
    if chave not in orcamentos:
        orcamentos[chave] = _orcamento_vazio(posto)

    orcamento = orcamentos[chave]

    # Cabeçalho do orçamento.
    c1, c2, c3 = st.columns([1, 1.6, 1])
    with c1:
        numero = st.text_input("Nº do orçamento", value=str(orcamento.get("numero", "")))
    with c2:
        titulo = st.text_input("Título", value=str(orcamento.get("titulo", "")))
    with c3:
        validade = st.number_input(
            "Validade (dias)",
            min_value=1,
            max_value=365,
            value=int(orcamento.get("validade_dias", 7) or 7),
        )

    c4, c5 = st.columns(2)
    with c4:
        responsavel = st.text_input("Responsável pelo orçamento", value=str(orcamento.get("responsavel", "")))
    with c5:
        condicao = st.selectbox(
            "Condição de pagamento",
            ["A definir", "À vista", "50% entrada + 50% na conclusão", "Parcelado", "Outro"],
            index=(
                ["A definir", "À vista", "50% entrada + 50% na conclusão", "Parcelado", "Outro"]
                .index(orcamento.get("condicao_pagamento", "A definir"))
                if orcamento.get("condicao_pagamento", "A definir") in
                ["A definir", "À vista", "50% entrada + 50% na conclusão", "Parcelado", "Outro"]
                else 0
            ),
        )

    st.markdown("#### 🧾 Treinamentos e Logística")
    st.caption(
        "Inclua treinamentos e os custos necessários à execução: deslocamento, "
        "hospedagem, aluguel de carro e transporte por aplicativo."
    )

    itens_iniciais = _normalizar_itens_orcamento(
        orcamento.get("itens", [])
    )
    df_itens_inicial = pd.DataFrame(
        itens_iniciais,
        columns=[
            "Item", "Produto", "Quantidade", "Unidade",
            "Valor Unitário", "Total"
        ],
    )

    if df_itens_inicial.empty:
        df_itens_inicial = pd.DataFrame([{
            "Item": 1,
            "Produto": PRODUTOS_TREINAMENTO[0],
            "Quantidade": 1.0,
            "Unidade": "Dia",
            "Valor Unitário": 0.0,
            "Total": 0.0,
        }])

    df_editado = st.data_editor(
        df_itens_inicial,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "Item": st.column_config.TextColumn(
                "Item",
                disabled=True
            ),
            "Produto": st.column_config.SelectboxColumn(
                "Produto / Serviço",
                options=PRODUTOS_TREINAMENTO,
                required=True
            ),
            "Quantidade": st.column_config.NumberColumn(
                "Quantidade",
                min_value=0.0,
                step=0.5
            ),
            "Unidade": st.column_config.SelectboxColumn(
                "Unidade",
                options=UNIDADES_ORCAMENTO,
                required=True
            ),
            "Valor Unitário": st.column_config.NumberColumn(
                "Valor Unitário (R$)",
                min_value=0.0,
                step=0.01,
                format="R$ %.2f"
            ),
            "Total": st.column_config.NumberColumn(
                "Total (R$)",
                disabled=True,
                format="R$ %.2f"
            ),
        },
        key=f"orcamento_editor_{chave}",
    ).copy()

    if not df_editado.empty:
        # Ajusta automaticamente a unidade padrão para linhas novas ou vazias.
        for idx, linha in df_editado.iterrows():
            produto = str(linha.get("Produto", "") or "").strip()
            unidade = str(linha.get("Unidade", "") or "").strip()
            if not unidade:
                df_editado.at[idx, "Unidade"] = UNIDADE_PADRAO_ORCAMENTO.get(
                    produto, "Unidade"
                )

        df_editado["Total"] = (
            pd.to_numeric(
                df_editado["Quantidade"], errors="coerce"
            ).fillna(0)
            * pd.to_numeric(
                df_editado["Valor Unitário"], errors="coerce"
            ).fillna(0)
        )

    st.markdown("#### 🧮 Fechamento")
    c6, c7 = st.columns(2)
    with c6:
        desconto = st.number_input(
            "Desconto (%)", min_value=0.0, max_value=100.0,
            value=float(orcamento.get("desconto", 0) or 0), step=0.5
        )
    with c7:
        subtotal = float(pd.to_numeric(df_editado["Total"], errors="coerce").fillna(0).sum()) if not df_editado.empty else 0.0
        desconto_valor = subtotal * desconto / 100
        total = subtotal - desconto_valor
        st.metric("💰 Total do orçamento", f"R$ {total:,.2f}")

    observacoes = st.text_area(
        "Observações / condições comerciais",
        value=str(orcamento.get("observacoes", "")),
        height=100,
    )

    st.markdown("#### 📎 Documentos relacionados")
    documentos = orcamento.get("documentos", [])
    uploads = st.file_uploader(
        "Anexe propostas, contratos, memoriais, imagens ou outros documentos",
        type=["pdf", "doc", "docx", "xlsx", "xls", "csv", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key=f"orcamento_upload_{chave}",
    )

    if uploads:
        if st.button("📥 Salvar documentos anexados", key=f"salvar_docs_{chave}"):
            novos = _salvar_documentos_orcamento(pv, uploads)
            documentos.extend(novos)
            orcamento["documentos"] = documentos
            orcamentos[chave] = orcamento
            _salvar_orcamentos(orcamentos)
            st.session_state["orcamentos_crm"] = orcamentos
            st.success(f"✅ {len(novos)} documento(s) anexado(s).")
            st.rerun()

    if documentos:
        for doc in documentos:
            caminho = doc.get("arquivo", "")
            if caminho and os.path.exists(caminho):
                with open(caminho, "rb") as arquivo:
                    st.download_button(
                        f"⬇️ {doc.get('nome', 'Documento')}",
                        data=arquivo.read(),
                        file_name=doc.get("nome", "Documento"),
                        key=f"download_doc_{chave}_{doc.get('arquivo')}",
                    )
            else:
                st.caption(f"⚠️ Arquivo não encontrado: {doc.get('nome', 'Documento')}")

    st.divider()
    b1, b2 = st.columns(2)
    with b1:
        salvar = st.button("💾 Salvar Orçamento", use_container_width=True, type="primary", key=f"salvar_orc_{chave}")
    with b2:
        limpar = st.button("🆕 Novo Orçamento", use_container_width=True, key=f"novo_orc_{chave}")

    if salvar:
        itens_salvos = []
        for _, linha in df_editado.iterrows():
            produto = str(linha.get("Produto", "") or "").strip()
            if not produto:
                continue
            itens_salvos.append({
                "Item": str(linha.get("Item", len(itens_salvos) + 1)),
                "Produto": produto,
                "Quantidade": float(linha.get("Quantidade", 0) or 0),
                "Unidade": str(
                    linha.get(
                        "Unidade",
                        UNIDADE_PADRAO_ORCAMENTO.get(produto, "Unidade")
                    ) or ""
                ),
                "Valor Unitário": float(
                    linha.get("Valor Unitário", 0) or 0
                ),
                "Total": float(linha.get("Total", 0) or 0),
            })

        orcamento.update({
            "numero": numero,
            "titulo": titulo,
            "validade_dias": validade,
            "responsavel": responsavel,
            "condicao_pagamento": condicao,
            "desconto": desconto,
            "observacoes": observacoes,
            "itens": itens_salvos,
            "documentos": documentos,
            "atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "_total_calculado": total,
        })
        # Migração silenciosa: qualquer orçamento antigo que possuía frete deixa de usá-lo.
        orcamento.pop("frete", None)
        orcamentos[chave] = orcamento
        st.session_state["orcamentos_crm"] = orcamentos
        _salvar_orcamentos(orcamentos)
        st.success("✅ Orçamento de treinamento e logística salvo com sucesso!")

    if limpar:
        orcamentos[chave] = _orcamento_vazio(posto)
        _salvar_orcamentos(orcamentos)
        st.session_state["orcamentos_crm"] = orcamentos
        st.rerun()

    # Exportação somente depois de salvo.
    orcamento_atual = st.session_state["orcamentos_crm"].get(chave, orcamento)
    if orcamento_atual.get("itens"):
        st.markdown("#### 📤 Exportar orçamento pronto")
        excel_orc = _gerar_excel_orcamento(orcamento_atual, posto, pv)
        pacote = _criar_pacote_orcamento(orcamento_atual, posto, pv)
        nome_base = re.sub(r"[^A-Za-z0-9_-]", "_", str(orcamento_atual.get("numero", "orcamento")))

        e1, e2 = st.columns(2)
        with e1:
            st.download_button(
                "📊 Exportar Orçamento em Excel",
                data=excel_orc,
                file_name=f"{nome_base}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"export_excel_orc_{chave}",
            )
        with e2:
            st.download_button(
                "📦 Exportar Orçamento + Documentos",
                data=pacote,
                file_name=f"{nome_base}_com_documentos.zip",
                mime="application/zip",
                use_container_width=True,
                key=f"export_zip_orc_{chave}",
            )


COLUNAS_FILA = [
    "PV_Abadi", "Tipo_Necessidade", "Data_Ultimo_Treinamento",
    "Dias_desde_Ultimo_Treinamento", "Instrutor_Sugerido", "Semana_Sugerida",
    "Telefone_Contato", "Email_Contato", "Status_Contato", "Data_do_Contato", "Observacoes",
    "Nome_Contato", "Tem_Funcionarios", "Qtd_Funcionarios", "Material_Em_Loja", "Data_Agendada",
    "Tipo_Pagamento", "Data_Pagamento", "Data_Liberacao_Treinamento",
]

# --- ESTILIZAÇÃO CSS CUSTOMIZADA ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600&display=swap');

:root {
    --ampm-orange: #FF4D00;
    --ampm-orange-2: #FF8A00;
    --ampm-orange-soft: #FFB38A;
    --ampm-yellow: #FFD300;
    --ampm-yellow-soft: #FFF4B0;
    --ampm-blue: #0D47A1;
    --ampm-blue-2: #1976D2;
    --ampm-blue-soft: #E3F2FD;

    /* Carvão propositalmente restrito a navegação, títulos e pequenos contrastes. */
    --charcoal: #1A1A1A;
    --charcoal-soft: #2B2B2B;

    --bg-app: #F7F8FA;
    --bg-surface: #FFFFFF;
    --bg-surface-alt: #FAFAFB;
    --bg-soft: #F5F5F5;
    --border-subtle: #E8E8E8;
    --border-strong: #D7D7D7;

    --text-primary: #202124;
    --text-secondary: #5F6368;
    --text-tertiary: #80868B;

    --success: #149B55;
    --success-bg: #E9F8F0;
    --warning: #D99000;
    --warning-bg: #FFF7DB;
    --danger: #D93025;
    --danger-bg: #FDECEA;
    --info: #1976D2;
    --info-bg: #EAF3FD;
    --neutral-bg: #F1F3F4;

    --radius-sm: 9px;
    --radius-md: 14px;
    --radius-lg: 20px;
    --shadow-sm: 0 1px 3px rgba(26,26,26,.06), 0 1px 2px rgba(26,26,26,.04);
    --shadow-md: 0 8px 24px rgba(26,26,26,.08);
    --shadow-lg: 0 18px 45px rgba(26,26,26,.10);
    --shadow-brand: 0 10px 28px rgba(255,77,0,.18);
}

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    letter-spacing: -0.01em;
}
.stApp {
    background:
        radial-gradient(circle at 100% 0%, rgba(255,211,0,.08), transparent 28rem),
        radial-gradient(circle at 0% 8%, rgba(255,77,0,.05), transparent 26rem),
        var(--bg-app);
    color: var(--text-primary);
}
code, .mono { font-family: 'JetBrains Mono', monospace !important; }
h1,h2,h3,h4,h5,h6 { color: var(--text-primary); letter-spacing:-.025em; }
p, label, .stCaption { color: var(--text-secondary); }
hr { border-color: var(--border-subtle) !important; }
.block-container { padding-top: 1.45rem; padding-bottom: 3rem; max-width: 1600px; }

/* Cabeçalho: marca forte, mas sem transformar a aplicação inteira em laranja. */
.main-header {
    background: linear-gradient(105deg, var(--ampm-orange) 0%, #FF6500 58%, var(--ampm-orange-2) 100%);
    padding: 27px 32px;
    border-radius: var(--radius-lg);
    color: #fff;
    margin-bottom: 26px;
    box-shadow: var(--shadow-brand);
    position: relative;
    overflow: hidden;
    border: 1px solid rgba(255,255,255,.35);
}
.main-header::before {
    content:"";
    position:absolute;
    width:220px; height:220px;
    right:-55px; top:-95px;
    border:28px solid rgba(255,211,0,.25);
    border-radius:50%;
}
.main-header::after {
    content:"";
    position:absolute;
    left:0; right:0; bottom:0; height:5px;
    background:var(--ampm-yellow);
}
.main-header-top {
    display:flex; justify-content:space-between; align-items:flex-start;
    position:relative; z-index:2;
}
.main-header h1 {
    color:#fff !important; margin:0 0 6px; font-weight:800;
    font-size:2rem; display:flex; align-items:center; gap:10px;
}
.main-header p { margin:0; color:rgba(255,255,255,.94); font-weight:500; }
.header-status-chip {
    background:rgba(255,255,255,.16);
    border:1px solid rgba(255,255,255,.38);
    color:#fff; padding:7px 13px; border-radius:999px;
    font-size:.76rem; font-weight:700; backdrop-filter:blur(8px);
}
.pulse-dot {
    width:7px; height:7px; border-radius:50%; background:#D7FF68;
    display:inline-block; margin-right:5px;
}

/* Seções */
.section-header { display:flex; align-items:center; gap:12px; margin:4px 0 17px; }
.section-header .icon-badge {
    width:38px; height:38px; min-width:38px; border-radius:11px;
    background:linear-gradient(135deg,var(--ampm-orange),var(--ampm-orange-2));
    color:#fff; display:flex; align-items:center; justify-content:center;
    box-shadow:var(--shadow-sm);
}
.section-header .titles h3 { margin:0; font-size:1.12rem; font-weight:750; color:var(--text-primary); }
.section-header .titles span { font-size:.81rem; color:var(--text-secondary); }

/* Cards claros */
.kpi-card, .ampm-column, .procv-card, .top-instructor-card {
    background:var(--bg-surface);
    border:1px solid var(--border-subtle);
    box-shadow:var(--shadow-sm);
}
.kpi-card {
    border-radius:var(--radius-md); padding:20px 21px;
    transition:transform .16s ease, box-shadow .16s ease, border-color .16s ease;
}
.kpi-card:hover { transform:translateY(-2px); box-shadow:var(--shadow-md); border-color:#FFD7C5; }
.kpi-header { display:flex; justify-content:space-between; align-items:center; }
.kpi-icon-circle {
    width:34px; height:34px; border-radius:10px;
    display:flex; align-items:center; justify-content:center;
    background:var(--ampm-blue-soft);
}
.kpi-title {
    font-size:.72rem; color:var(--text-secondary); text-transform:uppercase;
    font-weight:750; letter-spacing:.75px;
}
.kpi-value {
    font-size:1.95rem; font-weight:800; color:var(--text-primary);
    margin-top:10px; line-height:1; font-family:'JetBrains Mono',monospace;
}
.kpi-footer { margin-top:9px; font-size:.75rem; color:var(--text-tertiary); }

.ampm-column { border-radius:var(--radius-md); padding:15px; min-height:480px; }
.ampm-title {
    font-size:.83rem; font-weight:750; margin-bottom:13px; padding-bottom:10px;
    border-bottom:1px solid var(--border-subtle); display:flex;
    justify-content:space-between; color:var(--text-primary);
    text-transform:uppercase; letter-spacing:.45px;
}
.ampm-title .pill-count {
    background:var(--neutral-bg); border:1px solid var(--border-subtle);
    padding:2px 9px; border-radius:999px; font-size:.76rem; font-weight:700;
}
.col-a-contatar { border-top:3px solid #9AA0A6; }
.col-em-negociacao { border-top:3px solid var(--ampm-yellow); }
.col-agendado { border-top:3px solid var(--ampm-blue-2); }
.col-treinamento-realizado { border-top:3px solid var(--success); }
.col-recusado { border-top:3px solid var(--danger); }

.procv-card {
    padding:21px; border-radius:var(--radius-md);
    border-top:3px solid var(--ampm-orange); margin-bottom:15px;
}
.procv-card h4 { margin:0 0 13px; color:var(--ampm-blue); font-size:.98rem; font-weight:750; }
.procv-card p { margin:6px 0; font-size:.88rem; color:var(--text-primary); line-height:1.5; }
.procv-card p b { color:var(--text-secondary); font-weight:650; }

.top-instructor-card {
    padding:19px; border-radius:var(--radius-md);
    border-left:4px solid var(--ampm-blue-2); margin-bottom:13px;
}
.timeline-item {
    border-left:3px solid var(--ampm-orange); padding:4px 0 4px 16px;
    margin-bottom:15px; position:relative;
}
.timeline-item::before {
    content:""; position:absolute; left:-7px; top:8px; width:11px; height:11px;
    border-radius:50%; background:var(--ampm-yellow);
    border:2px solid var(--ampm-orange);
}

/* Badges */
.badge-info,.badge-pill {
    display:inline-flex; align-items:center; gap:4px; padding:3px 10px;
    border-radius:999px; font-weight:700; font-size:.73rem; border:1px solid transparent;
}
.badge-info { background:#FFF0E9; color:#D94200; border-color:#FFD3C0; }
.badge-neutral { background:var(--neutral-bg); color:var(--text-secondary); border-color:#E2E5E7; }
.badge-warning { background:var(--warning-bg); color:#8A6200; border-color:#F6DF8A; }
.badge-info-blue { background:var(--info-bg); color:var(--ampm-blue); border-color:#C8DFF7; }
.badge-success { background:var(--success-bg); color:#087A3D; border-color:#BFE9D2; }
.badge-danger { background:var(--danger-bg); color:#B3261E; border-color:#F3C5C1; }

/* Botões */
.stButton>button[kind="primary"], .stFormSubmitButton>button[kind="primary"] {
    background:linear-gradient(100deg,var(--ampm-orange),var(--ampm-orange-2)) !important;
    color:#fff !important; border:none !important;
    box-shadow:0 5px 14px rgba(255,77,0,.18) !important;
}
.stButton>button, .stFormSubmitButton>button, .stDownloadButton>button,
[data-testid="stLinkButton"] a {
    border-radius:var(--radius-sm) !important;
    font-weight:700 !important;
    transition:all .16s ease !important;
}
.stButton>button:not([kind="primary"]), .stFormSubmitButton>button:not([kind="primary"]),
.stDownloadButton>button, [data-testid="stLinkButton"] a {
    background:#fff !important;
    color:var(--ampm-orange) !important;
    border:1px solid #FFB999 !important;
}
.stButton>button:hover, .stDownloadButton>button:hover, [data-testid="stLinkButton"] a:hover {
    transform:translateY(-1px);
    box-shadow:var(--shadow-sm) !important;
    border-color:var(--ampm-orange) !important;
}

/* Inputs, tabelas, forms */
.stTextInput input,.stNumberInput input,.stTextArea textarea,.stDateInput input,
div[data-baseweb="select"]>div {
    background:#fff !important; color:var(--text-primary) !important;
    border-radius:var(--radius-sm) !important; border-color:var(--border-strong) !important;
}
.stTextInput input:focus,.stNumberInput input:focus,.stTextArea textarea:focus {
    border-color:var(--ampm-blue-2) !important;
    box-shadow:0 0 0 1px var(--ampm-blue-2) !important;
}
.stDataFrame {
    border-radius:var(--radius-md); overflow:hidden;
    border:1px solid var(--border-subtle); background:#fff;
}
div[data-testid="stExpander"] {
    border-radius:var(--radius-md) !important;
    border:1px solid var(--border-subtle) !important;
    background:#fff !important; overflow:hidden;
}
div[data-testid="stForm"] {
    border:1px solid var(--border-subtle);
    border-radius:var(--radius-md); padding:18px; background:#fff;
}
div[data-testid="stMetric"] {
    background:#fff; border:1px solid var(--border-subtle);
    border-radius:var(--radius-md); padding:14px 16px; box-shadow:var(--shadow-sm);
}

/* Tabs */
button[data-baseweb="tab"] { font-weight:650; }
button[data-baseweb="tab"][aria-selected="true"] {
    color:var(--ampm-orange) !important;
    border-bottom-color:var(--ampm-orange) !important;
}

/* Sidebar = principal uso do carvão. */
section[data-testid="stSidebar"] {
    background:linear-gradient(180deg,var(--charcoal) 0%, #222325 100%) !important;
    border-right:1px solid #333 !important;
}
section[data-testid="stSidebar"] * { color:#F7F7F7; }
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
section[data-testid="stSidebar"] .stCaption { color:#C9CDD2 !important; }
section[data-testid="stSidebar"] .stRadio label {
    font-size:.9rem; padding:4px 2px;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
    background:rgba(255,77,0,.10); border-radius:8px;
}
section[data-testid="stSidebar"] input { accent-color:var(--ampm-orange); }
.sidebar-brand { display:flex; align-items:center; gap:10px; margin-bottom:2px; }
.sidebar-brand .logo-chip,.logo-chip {
    width:34px; height:34px; border-radius:9px;
    background:linear-gradient(135deg,var(--ampm-orange),var(--ampm-yellow));
    display:flex; align-items:center; justify-content:center; font-size:1rem;
    box-shadow:var(--shadow-sm);
}
.sidebar-metric {
    background:rgba(255,255,255,.06);
    border:1px solid rgba(255,255,255,.10);
    border-radius:var(--radius-sm); padding:10px 12px;
    font-size:.81rem; color:#D2D5D9; margin-top:4px;
}
.sidebar-metric b { color:#fff; }

/* Alerts */
div[data-testid="stAlert"] {
    border-radius:var(--radius-md);
    border-width:1px;
}

/* Scroll */
::-webkit-scrollbar { width:8px; height:8px; }
::-webkit-scrollbar-track { background:transparent; }
::-webkit-scrollbar-thumb { background:#C9CDD2; border-radius:8px; }
::-webkit-scrollbar-thumb:hover { background:#AEB4BA; }

/* ==========================================================
   V49 — APRESENTAÇÃO APROVADA: DASHBOARD / SISTEMA
   ========================================================== */

/* Barra superior AmPm + IGT */
.brand-topbar {
    margin: -0.15rem 0 1.25rem 0;
    min-height: 82px;
    border-radius: 0 0 26px 26px;
    background:
        linear-gradient(105deg, #FF3D00 0%, #FF4D00 46%, #FF8A00 61%,
        #FFD300 72%, #FFFFFF 72.2%, #FFFFFF 100%);
    box-shadow: 0 7px 24px rgba(26,26,26,.10);
    display:flex;
    align-items:center;
    justify-content:space-between;
    padding: 0 28px 0 30px;
    overflow:hidden;
    border-bottom:1px solid #ECECEC;
}
.brand-top-left {
    display:flex; align-items:center; gap:22px; color:#fff;
}
.ampm-wordmark {
    color:#FFD300;
    font-size:2.45rem;
    font-weight:900;
    font-style:italic;
    letter-spacing:-.13rem;
    line-height:1;
    text-shadow:0 1px 1px rgba(0,0,0,.05);
}
.brand-divider {
    width:1px; height:46px; background:rgba(255,255,255,.7);
}
.brand-product {
    color:#fff; line-height:1.08;
}
.brand-product strong {
    display:block; font-size:1.03rem; font-weight:800; letter-spacing:.02rem;
}
.brand-product span {
    display:block; font-size:1.24rem; font-weight:800; margin-top:3px;
}
.brand-top-right {
    display:flex; align-items:center; gap:18px;
}
.igt-wordmark {
    color:#16181C; line-height:.8; text-align:center;
    font-weight:900; font-size:2.2rem; letter-spacing:-.12rem;
}
.igt-wordmark small {
    display:block; font-size:.48rem; letter-spacing:.16rem;
    font-weight:750; margin-top:9px;
}
.top-user-chip {
    background:#F5F6F8;
    border:1px solid #ECEEF1;
    border-radius:999px;
    color:#181A1F;
    padding:9px 16px;
    font-size:.79rem;
    font-weight:700;
    box-shadow:0 2px 8px rgba(0,0,0,.04);
}

/* Sidebar aprovada */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #171A20 0%, #20242C 100%) !important;
}
section[data-testid="stSidebar"] .sidebar-brand {
    padding: 8px 4px 14px;
}
.sidebar-ampm {
    color:#FFD300 !important;
    font-size:2.02rem;
    font-weight:900;
    font-style:italic;
    letter-spacing:-.11rem;
}
.sidebar-igt {
    margin-top:20px;
    padding:18px 14px;
    border-top:1px solid rgba(255,255,255,.12);
    color:#fff;
}
.sidebar-igt-logo {
    font-size:2rem; font-weight:900; letter-spacing:-.08rem;
    line-height:1; color:#fff;
}
.sidebar-igt-logo span {
    font-size:.67rem; letter-spacing:.10rem; font-weight:650;
}
.sidebar-igt p {
    font-size:.72rem !important;
    color:#C8CCD3 !important;
    line-height:1.45;
    margin:.75rem 0 0;
}

/* Dashboard */
.dashboard-title {
    display:flex; align-items:center; gap:13px; margin: 3px 0 5px;
}
.dashboard-title-icon {
    width:44px; height:44px; border-radius:12px;
    display:flex; align-items:center; justify-content:center;
    background:linear-gradient(135deg,#FF4D00,#FFD300);
    color:#fff; font-size:1.25rem;
    box-shadow:0 5px 14px rgba(255,77,0,.18);
}
.dashboard-title h1 {
    margin:0; font-size:1.75rem; font-weight:850; color:#17191E;
}
.dashboard-subtitle {
    color:#686D75; margin:0 0 17px 57px; font-size:.88rem;
}

.dashboard-kpi {
    background:#fff;
    border:1px solid #E7E9ED;
    border-radius:17px;
    padding:17px 17px 15px;
    min-height:120px;
    box-shadow:0 4px 14px rgba(23,25,30,.055);
}
.dashboard-kpi-top {
    display:flex; align-items:center; gap:10px;
}
.dashboard-kpi-icon {
    width:39px; height:39px; border-radius:11px;
    display:flex; align-items:center; justify-content:center;
    color:#fff; font-size:1.05rem; font-weight:800;
}
.dk-orange { background:linear-gradient(135deg,#FF3D00,#FF7200); }
.dk-yellow { background:linear-gradient(135deg,#FFB800,#FFD300); color:#222; }
.dk-blue { background:linear-gradient(135deg,#0D47A1,#1976D2); }
.dk-cyan { background:linear-gradient(135deg,#25AFC5,#64D3E3); }
.dk-purple { background:linear-gradient(135deg,#6F31D7,#A742F4); }
.dashboard-kpi-label {
    color:#333840; font-size:.66rem; text-transform:uppercase;
    font-weight:800; letter-spacing:.035rem; line-height:1.25;
}
.dashboard-kpi-value {
    font-family:'JetBrains Mono', monospace;
    font-size:1.67rem; font-weight:800; color:#15171B;
    margin:9px 0 2px;
}
.dashboard-kpi-note {
    font-size:.68rem; color:#777C84; margin-left:49px;
}

.dashboard-panel {
    background:#fff;
    border:1px solid #E6E8EC;
    border-radius:18px;
    padding:17px 18px;
    box-shadow:0 4px 15px rgba(23,25,30,.05);
    height:100%;
}
.dashboard-panel-title {
    font-size:.94rem; font-weight:800; color:#20232A;
    margin-bottom:14px;
}
.dashboard-panel-link {
    color:#FF4D00; text-align:center; font-size:.72rem;
    font-weight:750; margin-top:12px;
}
.dashboard-mini-row {
    padding:10px 2px;
    display:flex; align-items:center; justify-content:space-between;
    border-bottom:1px solid #F0F1F3;
    gap:8px;
}
.dashboard-mini-row:last-child { border-bottom:none; }
.dashboard-mini-main { font-size:.76rem; font-weight:700; color:#24272D; }
.dashboard-mini-sub { font-size:.66rem; color:#757B84; margin-top:2px; }
.dashboard-mini-badge {
    width:30px; height:30px; border-radius:9px;
    display:flex; align-items:center; justify-content:center;
    background:#FFF0E9;
}
.dashboard-model-row {
    display:grid; grid-template-columns:90px 1fr 45px;
    gap:10px; align-items:center; margin:17px 0;
    font-size:.77rem; color:#2C3036;
}
.dashboard-progress {
    height:12px; background:#ECEFF2; border-radius:999px; overflow:hidden;
}
.dashboard-progress > span {
    display:block; height:100%; border-radius:999px;
}
.dashboard-pipeline-strip {
    display:grid;
    grid-template-columns:repeat(6,minmax(90px,1fr));
    gap:8px;
}
.pipeline-stat {
    border-radius:13px;
    padding:11px 8px;
    text-align:center;
    font-size:.67rem;
    line-height:1.2;
}
.pipeline-stat strong {
    display:block; font-size:1.25rem; color:#181A1F; margin-top:4px;
}
.ps-blue {background:#E8F1FF;color:#0D47A1}
.ps-yellow {background:#FFF5CF;color:#8D6A00}
.ps-orange {background:#FFE6D9;color:#D94700}
.ps-green {background:#DCF6E7;color:#087A3D}
.ps-gray {background:#EEF0F2;color:#5F6368}
.ps-red {background:#FDE7E4;color:#B3261E}

.donut-wrap {
    display:flex; align-items:center; gap:22px; min-height:180px;
}
.dashboard-donut {
    width:150px; height:150px; min-width:150px;
    border-radius:50%;
    position:relative;
}
.dashboard-donut::after {
    content:"";
    position:absolute; inset:34px;
    background:#fff; border-radius:50%;
    box-shadow:inset 0 0 0 1px #ECEEF1;
}
.donut-center {
    position:absolute; inset:0; z-index:2;
    display:flex; align-items:center; justify-content:center;
    flex-direction:column; pointer-events:none;
}
.donut-center strong { font-size:1.3rem; color:#202329; }
.donut-center span { font-size:.62rem; color:#777D85; }
.donut-legend { flex:1; }
.donut-legend-row {
    display:grid; grid-template-columns:10px 1fr auto;
    gap:8px; align-items:center; margin:7px 0;
    font-size:.69rem;
}
.legend-dot { width:9px; height:9px; border-radius:50%; }

.dashboard-html-table {
    width:100%; border-collapse:separate; border-spacing:0;
    font-size:.69rem;
    overflow:hidden; border-radius:10px;
    border:1px solid #ECEEF1;
}
.dashboard-html-table th {
    background:linear-gradient(90deg,#FF4D00,#FF7200);
    color:#fff; padding:8px 9px; text-align:left; font-weight:750;
}
.dashboard-html-table td {
    padding:8px 9px; border-top:1px solid #EEF0F2; color:#34383F;
}
.dashboard-html-table tr:nth-child(even) td { background:#FCFCFD; }

.quick-actions-grid {
    display:grid; grid-template-columns:repeat(6,1fr);
    gap:11px;
}
.quick-action-card {
    min-height:66px;
    border-radius:13px;
    border:1px solid #E6E8EC;
    background:#fff;
    box-shadow:0 3px 10px rgba(0,0,0,.04);
    display:flex; align-items:center; justify-content:center;
    gap:9px; padding:9px;
    font-size:.72rem; font-weight:700; color:#24272D;
}
.quick-action-card span { font-size:1.2rem; }

/* Deixa os controles superiores com aparência do mockup */
div[data-baseweb="select"] > div,
.stDateInput input,
.stTextInput input {
    min-height:40px;
    border-radius:10px !important;
}

/* responsivo */
@media (max-width: 1100px) {
    .brand-topbar { padding:0 15px; }
    .brand-top-right .top-user-chip { display:none; }
    .dashboard-pipeline-strip { grid-template-columns:repeat(3,1fr); }
    .quick-actions-grid { grid-template-columns:repeat(3,1fr); }
}
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


def _usuario_atual():
    """Retorna o username/nome atualmente autenticado, de forma tolerante."""
    return (
        st.session_state.get("username")
        or st.session_state.get("user")
        or st.session_state.get("name")
        or ""
    )


CAMINHO_PERMISSOES = "permissoes_usuarios_ampm.json"

MODULOS_PERMISSOES = {
    "dashboard": "📊 Dashboard Executivo",
    "pipeline": "📋 Pipeline AmPm",
    "procv": "🔍 PROCV Gestão e Franquia AMPM",
    "custos": "📍 Calculadora & Otimizador de Custos",
    "callcenter": "📞 Call Center & Timeline WhatsApp",
    "instrutores": "👔 Equipe de Instrutores",
    "relatorios": "📂 Relatórios & Exportação",
}


def _usuario_atual():
    return (
        st.session_state.get("username")
        or st.session_state.get("user")
        or st.session_state.get("name")
        or ""
    )


def _lista_admins_configurada():
    """Administradores definidos no Streamlit Secrets, nunca por senha no código."""
    try:
        bruto = st.secrets.get("ADMIN_USERNAMES", "admin")
    except Exception:
        bruto = "admin"

    if isinstance(bruto, str):
        return {
            item.strip().lower()
            for item in bruto.replace(";", ",").split(",")
            if item.strip()
        }
    if isinstance(bruto, (list, tuple, set)):
        return {str(item).strip().lower() for item in bruto if str(item).strip()}
    return {"admin"}


def usuario_e_admin():
    return str(_usuario_atual()).strip().lower() in _lista_admins_configurada()


def carregar_permissoes_usuarios():
    padrao = {"admins": {}, "usuarios": {}}
    if not os.path.exists(CAMINHO_PERMISSOES):
        return padrao
    try:
        with open(CAMINHO_PERMISSOES, "r", encoding="utf-8") as f:
            dados = json.load(f)
        if not isinstance(dados, dict):
            return padrao
        dados.setdefault("admins", {})
        dados.setdefault("usuarios", {})
        return dados
    except Exception:
        return padrao


def salvar_permissoes_usuarios(permissoes):
    """Persiste apenas permissões; senhas nunca são armazenadas aqui."""
    with open(CAMINHO_PERMISSOES, "w", encoding="utf-8") as f:
        json.dump(permissoes, f, ensure_ascii=False, indent=2)


def permissoes_do_usuario(username=None):
    username = str(username or _usuario_atual()).strip().lower()
    if not username:
        return set()
    if usuario_e_admin() and username == str(_usuario_atual()).strip().lower():
        return set(MODULOS_PERMISSOES.keys())

    dados = carregar_permissoes_usuarios()
    registro = dados.get("usuarios", {}).get(username, {})
    if isinstance(registro, dict):
        return {
            chave for chave, permitido in registro.items()
            if chave in MODULOS_PERMISSOES and bool(permitido)
        }
    return set()


def usuario_tem_permissao(chave):
    return usuario_e_admin() or chave in permissoes_do_usuario()


def garantir_usuario_no_controle(username):
    username = str(username or "").strip().lower()
    if not username:
        return
    dados = carregar_permissoes_usuarios()
    usuarios = dados.setdefault("usuarios", {})
    if username not in usuarios:
        usuarios[username] = {chave: False for chave in MODULOS_PERMISSOES}
        salvar_permissoes_usuarios(dados)


def listar_usuarios_cadastrados():
    """Lê os usernames do arquivo local e dos Secrets, sem retornar senhas."""
    encontrados = set()
    arquivo = carregar_usuarios_arquivo()
    encontrados.update(arquivo.get("usernames", {}).keys())
    try:
        secrets_cred = _secrets_para_dict(st.secrets.get("credentials", {}))
        encontrados.update(secrets_cred.get("usernames", {}).keys())
    except Exception:
        pass
    return sorted(str(u).strip().lower() for u in encontrados if str(u).strip())


def salvar_permissoes_admin(username, novas_permissoes):
    if not usuario_e_admin():
        raise PermissionError("Somente administradores podem alterar permissões.")

    username = str(username or "").strip().lower()
    if not username:
        raise ValueError("Usuário inválido.")

    if username in _lista_admins_configurada():
        # Admins continuam com acesso total e não podem ser bloqueados por flags.
        return

    dados = carregar_permissoes_usuarios()
    dados.setdefault("usuarios", {})[username] = {
        chave: bool(novas_permissoes.get(chave, False))
        for chave in MODULOS_PERMISSOES
    }
    salvar_permissoes_usuarios(dados)


def render_administracao():
    if not usuario_e_admin():
        st.error("🚫 Acesso negado. Esta área é exclusiva para administradores.")
        st.stop()

    render_section_header(
        "🛡️",
        "Administração",
        "Controle de usuários e permissões do CRM"
    )

    st.info(
        "🔐 Os administradores possuem acesso total. "
        "Para usuários comuns, marque nas caixas abaixo exatamente quais módulos "
        "eles podem acessar. Senhas não são exibidas nem armazenadas neste arquivo."
    )

    usuarios = listar_usuarios_cadastrados()
    if not usuarios:
        st.warning("Nenhum usuário cadastrado foi encontrado.")
        return

    usuario_selecionado = st.selectbox(
        "👤 Usuário para configurar",
        usuarios,
        format_func=lambda u: f"{u} {'(ADMIN)' if u in _lista_admins_configurada() else ''}"
    )

    if usuario_selecionado in _lista_admins_configurada():
        st.success("🛡️ Este usuário é administrador e possui acesso total.")
        st.caption(
            "Acesso de administrador é definido em ADMIN_USERNAMES nos Secrets."
        )
        return

    garantir_usuario_no_controle(usuario_selecionado)
    atuais = permissoes_do_usuario(usuario_selecionado)

    st.markdown("### 🔑 Permissões de acesso")

    with st.form(f"form_permissoes_{usuario_selecionado}"):
        novas = {}
        cols = st.columns(2)
        for i, (chave, nome) in enumerate(MODULOS_PERMISSOES.items()):
            with cols[i % 2]:
                novas[chave] = st.checkbox(
                    nome,
                    value=chave in atuais,
                    key=f"perm_{usuario_selecionado}_{chave}"
                )

        salvar = st.form_submit_button(
            "💾 Salvar permissões",
            use_container_width=True
        )

    if salvar:
        try:
            salvar_permissoes_admin(usuario_selecionado, novas)
            st.success(f"✅ Permissões de **{usuario_selecionado}** atualizadas.")
            st.rerun()
        except Exception as exc:
            st.error(f"❌ Não foi possível salvar as permissões: {exc}")


def _texto_seguro_instrutor(valor):
    if valor is None:
        return ""
    try:
        resultado = pd.isna(valor)
        if isinstance(resultado, bool) and resultado:
            return ""
        if not isinstance(resultado, bool) and bool(resultado.all()):
            return ""
    except (TypeError, ValueError):
        pass
    return str(valor)


def filtrar_instrutores_ativos(df):
    """
    Exibe somente instrutores cujo STATUS esteja marcado como ativo.
    Mantém o banco completo, incluindo históricos de instrutores que saíram.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=list(df.columns) if df is not None else [])

    if "STATUS" not in df.columns:
        return df.copy()

    status_normalizado = (
        df["STATUS"]
        .map(_texto_seguro_instrutor)
        .str.strip()
        .str.casefold()
    )

    # Apenas status explicitamente ativo entra na equipe operacional.
    return df.loc[status_normalizado.eq("ativo")].copy()


def adicionar_instrutor_admin(nome, telefone="", email="", cidade="", uf=""):
    """Adiciona ou atualiza um instrutor como ATIVO no banco, somente após validação de admin."""
    if not usuario_e_admin():
        raise PermissionError("Somente administradores podem cadastrar instrutores.")

    nome = str(nome or "").strip()
    if not nome:
        raise ValueError("Informe o nome completo do instrutor.")

    bases = st.session_state["bases"]
    df = bases.get("instrutores", pd.DataFrame()).copy()

    colunas_padrao = ["NOME_COMPLETO", "STATUS", "TELEFONE", "EMAIL", "Cidade", "UF"]
    for coluna in colunas_padrao:
        if coluna not in df.columns:
            df[coluna] = ""

    # Garante que novos cadastros entrem como ativos.
    dados_novo = {
        "NOME_COMPLETO": nome,
        "STATUS": "Ativo",
        "TELEFONE": str(telefone or "").strip(),
        "EMAIL": str(email or "").strip(),
        "Cidade": str(cidade or "").strip(),
        "UF": str(uf or "").strip().upper(),
    }

    nomes_existentes = (
        df["NOME_COMPLETO"]
        .map(_texto_seguro_instrutor)
        .str.strip()
        .str.casefold()
    )
    mask = nomes_existentes.eq(nome.casefold())

    if mask.any():
        # Se já existe, atualiza o cadastro e reativa o instrutor.
        idx = df.index[mask][0]
        for coluna, valor in dados_novo.items():
            df.at[idx, coluna] = valor
        acao = "atualizado e reativado"
    else:
        df = pd.concat([df, pd.DataFrame([dados_novo])], ignore_index=True)
        acao = "adicionado"

    bases["instrutores"] = df
    st.session_state["bases"] = bases
    salvar_bases_combinadas_no_disco(bases)
    _sincronizar_modulos_apos_mutacao("instrutores")

    return acao

def atualizar_status_instrutor_admin(nome, ativo):
    """Ativa/inativa um instrutor no Supabase e sincroniza a sessão."""
    if not usuario_e_admin():
        raise PermissionError("Somente administradores podem alterar o status de instrutores.")

    nome = str(nome or "").strip()
    if not nome:
        raise ValueError("Instrutor inválido.")

    novo_status = "Ativo" if bool(ativo) else "Saiu"

    # Fonte oficial: Supabase.
    (
        _supabase_client()
        .table("crm_instrutores")
        .update({"status": novo_status})
        .eq("nome_completo", nome)
        .execute()
    )

    # Sincroniza a cópia em memória imediatamente.
    bases = st.session_state.get("bases", {})
    df = bases.get("instrutores", pd.DataFrame()).copy()

    if not df.empty and "NOME_COMPLETO" in df.columns:
        mask = (
            df["NOME_COMPLETO"]
            .map(_texto_seguro_instrutor)
            .str.strip()
            .str.casefold()
            .eq(nome.casefold())
        )
        if mask.any():
            if "STATUS" not in df.columns:
                df["STATUS"] = ""
            df.loc[mask, "STATUS"] = novo_status
            bases["instrutores"] = df
            st.session_state["bases"] = bases

    _sincronizar_modulos_apos_mutacao("instrutores")
    return novo_status


def salvar_status_instrutores_admin(df_editor, df_original):
    """Aplica somente as mudanças feitas nas caixas de seleção."""
    if not usuario_e_admin():
        raise PermissionError("Somente administradores podem alterar o status de instrutores.")

    if df_editor is None or df_editor.empty:
        return 0

    alteracoes = 0

    orig_map = {}
    for _, row in df_original.iterrows():
        nome = str(row.get("NOME_COMPLETO", "") or "").strip()
        if nome:
            orig_map[nome.casefold()] = (
                str(row.get("STATUS", "") or "").strip().casefold() == "ativo"
            )

    for _, row in df_editor.iterrows():
        nome = str(row.get("NOME_COMPLETO", "") or "").strip()
        if not nome:
            continue

        ativo_novo = bool(row.get("Em atividade", False))
        ativo_antigo = orig_map.get(nome.casefold())

        if ativo_antigo is None or ativo_novo == ativo_antigo:
            continue

        atualizar_status_instrutor_admin(nome, ativo_novo)
        alteracoes += 1

    return alteracoes




def _nome_arquivo_seguro(texto):
    texto = unicodedata.normalize("NFKD", str(texto or "exportacao"))
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"[^A-Za-z0-9_-]+", "_", texto).strip("_")
    return texto or "exportacao"


def _excel_dataframe_bytes(df, nome_aba="Dados"):
    if df is None:
        df = pd.DataFrame()
    buffer = io.BytesIO()
    aba = str(nome_aba or "Dados")[:31]
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.copy().to_excel(writer, index=False, sheet_name=aba)
        ws = writer.sheets[aba]
        ws.freeze_panes = "A2"
        if ws.max_column:
            ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            letra = col[0].column_letter
            maior = max([len(str(c.value or "")) for c in list(col)[:300]] + [10])
            ws.column_dimensions[letra].width = min(maior + 2, 45)
    buffer.seek(0)
    return buffer.getvalue()


def render_exportacao_modulo(df, nome_modulo, nome_aba=None, legenda=None):
    if df is None:
        df = pd.DataFrame()
    df_export = df.copy()

    st.divider()
    st.markdown("### 📤 Exportar dados deste módulo")
    st.caption(
        legenda or "Exporta exatamente os dados disponíveis nesta visão do módulo."
    )

    if df_export.empty:
        st.info("Não há dados disponíveis nesta visão para exportar.")
        return

    nome_base = _nome_arquivo_seguro(nome_modulo)
    csv_bytes = df_export.to_csv(index=False).encode("utf-8-sig")
    excel_bytes = _excel_dataframe_bytes(df_export, nome_aba or nome_modulo)

    c_csv, c_xlsx = st.columns(2)
    with c_csv:
        st.download_button(
            "📄 Exportar CSV",
            data=csv_bytes,
            file_name=f"{nome_base}.csv",
            mime="text/csv",
            use_container_width=True,
            key=f"export_csv_mod_{nome_base}",
        )
    with c_xlsx:
        st.download_button(
            "📊 Exportar Excel",
            data=excel_bytes,
            file_name=f"{nome_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"export_xlsx_mod_{nome_base}",
        )



MODELOS_WHATSAPP_PADRAO = {
    "primeiro_contato": {
        "nome": "Primeiro contato",
        "mensagem": (
            "Olá, {nome_contato}! Tudo bem? Aqui é {atendente}, da equipe de Capacitação AmPm. "
            "Estou entrando em contato sobre o PV {pv} - {razao_social}. "
            "Identificamos a necessidade de {necessidade}. Podemos alinhar os próximos passos?"
        ),
    },
    "retreinamento": {
        "nome": "Retreinamento",
        "mensagem": (
            "Olá, {nome_contato}! Aqui é {atendente}, da equipe de Capacitação AmPm. "
            "Estamos entrando em contato sobre o PV {pv} - {razao_social} para alinharmos "
            "um retreinamento da equipe. Hoje temos {qtd_funcionarios} funcionário(s) para treinar. "
            "Podemos verificar uma data?"
        ),
    },
    "inauguracao": {
        "nome": "Inauguração",
        "mensagem": (
            "Olá, {nome_contato}! Aqui é {atendente}, da equipe de Capacitação AmPm. "
            "Estamos acompanhando a inauguração do PV {pv} - {razao_social}, prevista para "
            "{data_inauguracao}. Gostaríamos de alinhar o treinamento de {tipo_modelo} "
            "e a preparação da equipe."
        ),
    },
    "confirmacao_treinamento": {
        "nome": "Confirmação de treinamento",
        "mensagem": (
            "Olá, {nome_contato}! Confirmamos o treinamento do PV {pv} - {razao_social} "
            "para {data_agendada}, com o instrutor {instrutor}. Qualquer alteração, por favor nos avise."
        ),
    },
    "reagendamento": {
        "nome": "Reagendamento",
        "mensagem": (
            "Olá, {nome_contato}! Precisamos alinhar uma nova data para o treinamento do "
            "PV {pv} - {razao_social}. Podemos verificar a melhor disponibilidade para a equipe?"
        ),
    },
    "pos_treinamento": {
        "nome": "Pós-treinamento",
        "mensagem": (
            "Olá, {nome_contato}! Tudo bem? Gostaríamos de saber como foi o treinamento realizado "
            "no PV {pv} - {razao_social}. Ficamos à disposição para dúvidas, ajustes ou novos alinhamentos."
        ),
    },
}


def carregar_modelos_whatsapp():
    """Lê os modelos centrais do Supabase; usa defaults locais apenas como fallback."""
    try:
        resposta = (
            _supabase_client()
            .table("crm_modelos_whatsapp")
            .select("chave,nome,mensagem,ativo,atualizado_por,atualizado_em")
            .eq("ativo", True)
            .order("nome")
            .execute()
        )
        dados = resposta.data or []
        if dados:
            return {item["chave"]: item for item in dados}
    except Exception:
        pass

    return {
        chave: {
            "chave": chave,
            "nome": valor["nome"],
            "mensagem": valor["mensagem"],
            "ativo": True,
        }
        for chave, valor in MODELOS_WHATSAPP_PADRAO.items()
    }


def salvar_modelo_whatsapp(chave, nome, mensagem):
    """Salva edição como novo padrão global do CRM."""
    chave = str(chave or "").strip()
    nome = str(nome or "").strip()
    mensagem = str(mensagem or "").strip()

    if not chave or not nome or not mensagem:
        raise ValueError("Nome e mensagem do modelo são obrigatórios.")

    registro = {
        "chave": chave,
        "nome": nome,
        "mensagem": mensagem,
        "ativo": True,
        "atualizado_por": str(_usuario_atual() or ""),
        "atualizado_em": datetime.now().isoformat(),
    }

    (
        _supabase_client()
        .table("crm_modelos_whatsapp")
        .upsert(registro, on_conflict="chave")
        .execute()
    )

    st.session_state.pop("modelos_whatsapp_cache", None)
    return registro


def _texto_modelo_whatsapp(template, posto):
    """Substitui variáveis conhecidas do CRM sem quebrar se algum campo estiver vazio."""
    nome_contato = _procv_valor_flexivel(
        posto,
        "Nome_Contato",
        ["Nome do Contato", "Contato", "Responsável", "Responsavel"],
    )
    if nome_contato == "Não informado":
        nome_contato = "equipe"

    tipo_modelo = _procv_valor_flexivel(
        posto,
        "Tipo de Modelo",
        ["Tipo Modelo", "Modelo", "Modelo da Loja", "Modelo Loja"],
    )

    dados = {
        "nome_contato": nome_contato,
        "atendente": str(st.session_state.get("name") or _usuario_atual() or "equipe de Capacitação AmPm"),
        "pv": str(posto.get("PV Abadi", "") or ""),
        "razao_social": str(posto.get("Razão Social", "") or ""),
        "necessidade": str(posto.get("Tipo_Necessidade", "") or ""),
        "qtd_funcionarios": str(posto.get("Qtd_Funcionarios", "") or ""),
        "data_inauguracao": str(
            posto.get("Previsão Inauguração", "")
            or posto.get("Data Inauguração Atual", "")
            or ""
        ),
        "tipo_modelo": tipo_modelo if tipo_modelo != "Não informado" else "",
        "data_agendada": str(posto.get("Data_Agendada", "") or ""),
        "instrutor": str(posto.get("Instrutor_Sugerido", "") or ""),
        "municipio": str(posto.get("Municipio", "") or ""),
        "uf": str(posto.get("UF", "") or ""),
    }

    class _SafeDict(dict):
        def __missing__(self, key):
            return "{" + key + "}"

    try:
        return str(template or "").format_map(_SafeDict(dados))
    except Exception:
        return str(template or "")


def _render_editor_modelos_whatsapp():
    """Editor de padrões; qualquer usuário do Call Center pode ajustar e salvar o padrão global."""
    modelos = carregar_modelos_whatsapp()
    if not modelos:
        st.info("Nenhum modelo de WhatsApp disponível.")
        return

    opcoes = {v["nome"]: k for k, v in modelos.items()}
    nome_sel = st.selectbox(
        "Modelo para editar",
        list(opcoes.keys()),
        key="wa_editor_modelo_sel_v46",
    )
    chave_sel = opcoes[nome_sel]
    atual = modelos[chave_sel]

    novo_nome = st.text_input(
        "Nome do modelo",
        value=str(atual.get("nome", nome_sel)),
        key=f"wa_editor_nome_{chave_sel}",
    )
    nova_msg = st.text_area(
        "Mensagem padrão",
        value=str(atual.get("mensagem", "")),
        height=180,
        key=f"wa_editor_msg_{chave_sel}",
    )

    st.caption(
        "Variáveis disponíveis: {nome_contato}, {atendente}, {pv}, {razao_social}, "
        "{necessidade}, {qtd_funcionarios}, {data_inauguracao}, {tipo_modelo}, "
        "{data_agendada}, {instrutor}, {municipio}, {uf}."
    )

    if st.button(
        "💾 Salvar como modelo padrão",
        type="primary",
        use_container_width=True,
        key=f"wa_salvar_modelo_{chave_sel}",
    ):
        try:
            salvar_modelo_whatsapp(chave_sel, novo_nome, nova_msg)
            st.success("✅ Modelo atualizado. A partir de agora ele é o padrão para todos os usuários.")
            st.rerun()
        except Exception as exc:
            st.error(f"❌ Não foi possível salvar o modelo: {exc}")


# --- HELPERS DE APRESENTAÇÃO ---
def _render_html_dashboard(conteudo):
    """
    Renderiza HTML do Dashboard sem o Markdown interpretar blocos indentados
    como código. Isso é essencial para fragmentos HTML gerados dinamicamente.
    """
    texto = str(conteudo or "").strip()
    texto = re.sub(r">\s+<", "><", texto)
    st.markdown(texto, unsafe_allow_html=True)


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
    "Recusado": ("badge-danger", "🚫"),
    "Aguardando Pagamento": ("badge-warning", "💳"),
    "Em Negociação": ("badge-warning", "🤝"),
    "Agendado": ("badge-info-blue", "📅"),
    "Treinamento Realizado": ("badge-success", "✅"),
}

def badge_status_html(status):
    classe, emoji = STATUS_BADGE_MAP.get(str(status), ("badge-neutral", "•"))
    return f'<span class="badge-pill {classe}">{emoji} {status}</span>'

def status_css_class(status):
    mapa = {
        "A Contatar": "col-a-contatar",
        "Recusado": "col-recusado",
        "Aguardando Pagamento": "col-em-negociacao",
        "Em Negociação": "col-em-negociacao",
        "Agendado": "col-agendado",
        "Treinamento Realizado": "col-treinamento-realizado",
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
            "CNPJ": ["cnpj", "cnpj loja", "documento cnpj", "cnpj da loja", "cnpj posto", "cnpj completo", "cnpj c/ dv"],
            "Telefone_Contato": ["telefone", "telefone contato", "telefone da loja", "telefone loja", "celular", "whatsapp", "fone", "telefone comercial", "telefone contato loja"],
            "Email_Contato": ["email", "e mail", "e-mail", "email contato", "email da loja", "email loja", "correio eletronico", "correio eletrônico", "e-mail contato"],
            "Nome_Contato": ["nome contato", "nome do contato", "responsavel", "responsável", "responsavel loja", "responsável loja", "contato", "nome contato loja"],
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
            "Instrutor_Treinamento": ["instrutor treinamento", "instrutor do treinamento", "instrutor que treinou", "instrutor responsavel treinamento", "instrutor responsavel pelo treinamento"],
            "Instrutor_Inauguracao": ["instrutor inauguracao", "instrutor da inauguracao", "instrutor que inaugurou", "instrutor responsavel inauguracao"],
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



def _extrair_lojas_de_todas_as_abas(xls):
    """
    Consolida todas as abas que contenham uma chave compatível com PV.
    Isso evita depender da classificação da aba para CNPJ/telefone/e-mail.
    """
    definicao = ENTIDADES["lojas"]
    partes = []

    for sheet_name in xls.sheet_names:
        try:
            bruto = pd.read_excel(xls, sheet_name=sheet_name)
        except Exception:
            continue

        if bruto is None or bruto.empty:
            continue

        bruto = bruto.dropna(axis=0, how="all").dropna(axis=1, how="all")
        if bruto.empty:
            continue

        preparado, _, canonicas, _ = _preparar_dataframe_entidade(
            bruto, definicao
        )

        if "PV Abadi" not in canonicas or "PV Abadi" not in preparado.columns:
            continue

        preparado = _normalizar_chave_dataframe(
            preparado, "PV Abadi", True
        )
        preparado = preparado[preparado["PV Abadi"].notna()].copy()
        if not preparado.empty:
            partes.append(preparado)

    if not partes:
        return pd.DataFrame(columns=list(definicao["colunas"].keys()))

    combinado = pd.concat(partes, ignore_index=True, sort=False)

    # Mantém a última ocorrência de cada PV, mas antes agrega valores preenchidos
    # para não perder contato presente numa aba e cadastro presente em outra.
    colunas = [c for c in combinado.columns if c != "PV Abadi"]
    registros = []
    for pv, grupo in combinado.groupby("PV Abadi", dropna=True, sort=False):
        reg = {"PV Abadi": pv}
        for col in colunas:
            valor_final = pd.NA
            for valor in grupo[col].tolist():
                if _valor_preenchido(valor):
                    valor_final = valor
            reg[col] = valor_final
        registros.append(reg)

    return pd.DataFrame(registros)



def _extrair_instrutores_explicitos(xls):
    """
    Procura primeiro abas cujo nome indique claramente instrutores/equipe.
    Isso evita confundir consultores/GF/CF da base gerencial com instrutores.
    """
    definicao = ENTIDADES["instrutores"]
    candidatos = []

    for sheet_name in xls.sheet_names:
        nome_norm = _normalizar_nome(sheet_name)
        if not any(p in nome_norm for p in ("instrutor", "instrutores", "equipe instrutor", "equipe treinamento")):
            continue

        try:
            bruto = pd.read_excel(xls, sheet_name=sheet_name)
        except Exception:
            continue

        if bruto is None or bruto.empty:
            continue

        bruto = bruto.dropna(axis=0, how="all").dropna(axis=1, how="all")
        if bruto.empty:
            continue

        preparado, _, canonicas, _ = _preparar_dataframe_entidade(
            bruto, definicao
        )

        if "NOME_COMPLETO" not in canonicas or "NOME_COMPLETO" not in preparado.columns:
            continue

        preparado = _normalizar_chave_dataframe(
            preparado, "NOME_COMPLETO", False
        )
        preparado = preparado[
            preparado["NOME_COMPLETO"].map(_valor_preenchido)
        ].copy()

        if preparado.empty:
            continue

        # Normaliza status sem inventar informação.
        if "STATUS" not in preparado.columns:
            preparado["STATUS"] = "Ativo"

        # Mantém uma linha por instrutor pelo nome.
        preparado = preparado.drop_duplicates(
            subset=["NOME_COMPLETO"], keep="last"
        ).reset_index(drop=True)

        candidatos.append((sheet_name, preparado, len(canonicas)))

    if not candidatos:
        return None, pd.DataFrame(columns=list(definicao["colunas"].keys()))

    # Prioriza a aba com mais colunas reconhecidas e depois mais linhas.
    candidatos.sort(
        key=lambda x: (x[2], len(x[1])),
        reverse=True
    )
    return candidatos[0][0], candidatos[0][1]


def _substituir_instrutores_supabase(df_instrutores, tamanho_lote=300):
    """
    Quando uma aba explícita de Instrutores é importada, ela passa a ser
    a fonte oficial da equipe. Remove a lista antiga e grava a lista real.
    """
    if df_instrutores is None or df_instrutores.empty:
        return 0

    client = _supabase_client()

    # Limpa a tabela atual somente neste fluxo explícito de substituição.
    try:
        client.table("crm_instrutores").delete().neq("nome_completo", "").execute()
    except Exception:
        # Fallback para tabelas onde haja nulos/eventuais registros estranhos.
        try:
            existentes = _supabase_fetch_all("crm_instrutores")
            for item in existentes:
                nome = item.get("nome_completo")
                if nome:
                    client.table("crm_instrutores").delete().eq("nome_completo", nome).execute()
        except Exception as exc:
            raise RuntimeError(f"Não foi possível substituir a equipe de instrutores: {exc}")

    _upsert_dataframe_supabase(
        "instrutores",
        df_instrutores,
        tamanho_lote=tamanho_lote
    )
    return len(df_instrutores)


def detectar_entidades_no_workbook(xls):
    """
    Lê todas as abas e permite que uma mesma aba alimente várias entidades.

    Isso é importante para planilhas gerenciais/contatos: a mesma linha com PV
    pode conter dados de cadastro (lojas), contato (fila) e histórico. A versão
    anterior atribuía uma aba a apenas uma entidade e podia mandar uma planilha
    de contatos para a fila, deixando a Rede de Lojas sem CNPJ/telefone/e-mail.
    """
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
            score, canonicas = _score_aba_para_entidade(
                df_bruto, sheet_name, entidade, definicao
            )
            if definicao["chave"] in canonicas:
                # Uma aba pode alimentar mais de uma entidade.
                # Exigimos a chave e pelo menos uma coluna adicional reconhecida
                # para evitar importar abas irrelevantes.
                quantidade_reconhecida = len(canonicas)
                if quantidade_reconhecida >= 2 or entidade == "lojas":
                    candidatos.append((score, sheet_name, entidade, canonicas))

    # Para cada entidade, escolhe a melhor aba. A mesma aba pode ser escolhida
    # para várias entidades quando ela realmente contém os respectivos campos.
    prioridade = {"lojas": 5, "fila": 4, "inaug": 3, "instrutores": 2, "rec": 1}
    candidatos.sort(key=lambda x: (-x[0], -prioridade.get(x[2], 0), x[1]))

    entidade_atribuida = {}
    for score, sheet_name, entidade, canonicas in candidatos:
        if entidade not in entidade_atribuida:
            entidade_atribuida[entidade] = sheet_name

    if "lojas" not in entidade_atribuida:
        raise ValueError(
            "Nenhuma aba com uma chave de loja/PV foi reconhecida. "
            "Verifique se existe PV/Código da Loja."
        )

    bases = {}
    relatorio = []

    for entidade, definicao in ENTIDADES.items():
        colunas_canonicas = list(definicao["colunas"].keys())
        sheet_name = entidade_atribuida.get(entidade)

        if sheet_name:
            df_bruto = dfs_brutos[sheet_name]
            df_final, rename_map, canonicas, colunas_novas = _preparar_dataframe_entidade(
                df_bruto, definicao
            )
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
                "colunas_reconhecidas": [
                    c for c in colunas_canonicas if c in df_final.columns
                ],
                "colunas_novas": [
                    c for c in df_final.columns if c not in colunas_canonicas
                ],
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

    # Instrutores: usa preferencialmente a aba explícita "Instrutores"/"Equipe".
    aba_instrutores_exp, instrutores_exp = _extrair_instrutores_explicitos(xls)
    if instrutores_exp is not None and not instrutores_exp.empty:
        bases["instrutores"] = instrutores_exp
        for item in relatorio:
            if item.get("entidade") == "instrutores":
                item["aba_origem"] = aba_instrutores_exp
                item["linhas_lidas"] = len(instrutores_exp)
                item["confianca"] = "alta"
                item["colunas_reconhecidas"] = [
                    c for c in ENTIDADES["instrutores"]["colunas"].keys()
                    if c in instrutores_exp.columns
                ]
                break

    # Rede de Lojas é consolidada de TODAS as abas com PV.
    # Assim CNPJ/telefone/e-mail nunca dependem do nome/classificação da aba.
    lojas_consolidadas = _extrair_lojas_de_todas_as_abas(xls)
    if lojas_consolidadas is not None and not lojas_consolidadas.empty:
        bases["lojas"] = lojas_consolidadas
        for item in relatorio:
            if item.get("entidade") == "lojas":
                item["linhas_lidas"] = len(lojas_consolidadas)
                item["colunas_reconhecidas"] = [
                    c for c in ENTIDADES["lojas"]["colunas"].keys()
                    if c in lojas_consolidadas.columns
                ]
                item["aba_origem"] = "Consolidação de abas com PV"
                break

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

    # Pandas 3 pode inferir dtype "str" para colunas vindas do Supabase/Excel.
    # Uma planilha pode então trazer Timestamp, int, bool etc. para a mesma coluna,
    # e a atribuição falha com "Invalid value ... for dtype str".
    # Como o CRM é um integrador heterogêneo, colunas não-chave precisam aceitar
    # tipos mistos durante o merge. A conversão para tipos de banco ocorre depois.
    for _df in (atual, novo):
        for _col in _df.columns:
            if _col != chave:
                try:
                    _df[_col] = _df[_col].astype("object")
                except Exception:
                    pass

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


def _fingerprint_upload(conteudo):
    return hashlib.sha256(conteudo).hexdigest()


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


def _abrir_excel_resiliente(conteudo_bytes):
    """Abre Excel recebido pelo Streamlit com validação e fallbacks.

    Alguns arquivos chegam com extensão .xlsx mas podem ter sido gerados por
    sistemas diferentes. Primeiro validamos o pacote OOXML (ZIP), depois
    tentamos o openpyxl por bytes e, por fim, o leitor padrão do pandas.
    """
    if not conteudo_bytes:
        raise ValueError("O arquivo recebido está vazio.")

    dados = bytes(conteudo_bytes)
    if len(dados) < 4:
        raise ValueError("O arquivo recebido é pequeno demais para ser um Excel válido.")

    tentativas = []

    # XLSX/XLSM são pacotes ZIP. Se não for ZIP, o nome pode estar errado.
    if zipfile.is_zipfile(io.BytesIO(dados)):
        for engine in ("openpyxl", None):
            try:
                if engine:
                    return pd.ExcelFile(io.BytesIO(dados), engine=engine)
                return pd.ExcelFile(io.BytesIO(dados))
            except Exception as exc:
                tentativas.append(f"{engine or 'padrão'}: {exc}")
    else:
        # Evita o erro genérico "File is not a zip file" e explica a causa.
        assinatura = dados[:16].hex(" ")
        raise ValueError(
            "O arquivo possui extensão Excel, mas o conteúdo recebido não é um pacote XLSX válido. "
            f"Assinatura recebida: {assinatura}. Verifique se o arquivo não foi renomeado de CSV/HTML "
            "para .xlsx ou exporte novamente como .xlsx pelo Excel/LibreOffice."
        )

    raise ValueError("Não foi possível abrir o Excel. Tentativas: " + " | ".join(tentativas))


def validar_bytes_excel(conteudo_bytes):
    try:
        xls = _abrir_excel_resiliente(conteudo_bytes)
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


def _valor_util(valor):
    """Retorna True quando a célula contém informação real."""
    if valor is None:
        return False
    try:
        if pd.isna(valor):
            return False
    except (TypeError, ValueError):
        pass
    texto = str(valor).strip().lower()
    return texto not in {"", "nan", "nat", "none", "null", "<na>"}


def _coalescer_colunas(df, base, candidatos):
    """Escolhe o primeiro valor preenchido entre colunas equivalentes."""
    existentes = [c for c in candidatos if c in df.columns]
    if not existentes:
        return df
    principal = base
    if principal not in df.columns:
        df[principal] = pd.Series([pd.NA] * len(df), index=df.index, dtype="object")
    else:
        # Necessário para aceitar valores heterogêneos de planilhas (ex.: Timestamp).
        try:
            df[principal] = df[principal].astype("object")
        except Exception:
            pass
    for col in existentes:
        if col == principal:
            continue
        mask = ~df[principal].map(_valor_util)
        if mask.any():
            df.loc[mask, principal] = df.loc[mask, col]
    extras = [c for c in existentes if c != principal]
    if extras:
        df.drop(columns=extras, inplace=True, errors='ignore')
    return df


def _merge_base_por_pv(df_base, df_extra, chave_base, chave_extra):
    """
    Merge seguro por PV, tolerante a dtypes rígidos do pandas/pyarrow.

    Dados vindos do Supabase podem chegar como string[pyarrow]. Antes de
    coalescer campos de fila/inauguração sobre a rede de lojas, convertemos
    colunas não-chave para object. Isso permite misturar string, número,
    Timestamp e valores ausentes sem TypeError.
    """
    if df_extra is None or df_extra.empty or chave_extra not in df_extra.columns:
        return df_base

    base = df_base.copy()
    extra = df_extra.copy()

    if chave_base not in base.columns:
        return base

    base[chave_base] = pd.to_numeric(base[chave_base], errors='coerce')
    extra[chave_extra] = pd.to_numeric(extra[chave_extra], errors='coerce')

    # Neutraliza ExtensionArrays rígidos antes do merge.
    for _df, _chave in ((base, chave_base), (extra, chave_extra)):
        for _col in _df.columns:
            if _col == _chave:
                continue
            try:
                _df[_col] = _df[_col].astype("object")
            except Exception:
                pass

    extra = extra.dropna(subset=[chave_extra]).copy()
    if extra.empty:
        return base

    extra = extra.drop_duplicates(subset=[chave_extra], keep='last')

    extra_cols = [c for c in extra.columns if c != chave_extra]
    temp_names = {}
    for c in extra_cols:
        if c in base.columns:
            temp_names[c] = f"__extra__{c}"

    extra = extra.rename(columns=temp_names)

    merged = base.merge(
        extra,
        left_on=chave_base,
        right_on=chave_extra,
        how='left',
        suffixes=('', '__duplicata')
    )

    if chave_extra != chave_base and chave_extra in merged.columns:
        merged.drop(columns=[chave_extra], inplace=True, errors='ignore')

    # O merge pode recriar dtypes extension; converte novamente antes do coalesce.
    for original, tmp in temp_names.items():
        if original in merged.columns:
            try:
                merged[original] = merged[original].astype("object")
            except Exception:
                pass
        if tmp in merged.columns:
            try:
                merged[tmp] = merged[tmp].astype("object")
            except Exception:
                pass

    for original, tmp in temp_names.items():
        if tmp not in merged.columns:
            continue

        if original not in merged.columns:
            merged[original] = merged[tmp].astype("object")
        else:
            destino = merged[original].astype("object").copy()
            origem = merged[tmp].astype("object")
            mask = ~destino.map(_valor_util)
            if mask.any():
                destino.loc[mask] = origem.loc[mask].tolist()
            merged[original] = destino

        merged.drop(columns=[tmp], inplace=True, errors='ignore')

    for c in list(merged.columns):
        if not c.endswith('__duplicata'):
            continue

        original = c[:-11]
        if original in merged.columns:
            destino = merged[original].astype("object").copy()
            origem = merged[c].astype("object")
            mask = ~destino.map(_valor_util)
            if mask.any():
                destino.loc[mask] = origem.loc[mask].tolist()
            merged[original] = destino

        merged.drop(columns=[c], inplace=True, errors='ignore')

    return merged


def construir_base_unificada(df_lojas, df_fila, df_inaug):
    """Constrói a visão do PROCV sem perder dados de nenhuma entidade.
    Campos que existem em mais de uma origem são coalescidos por PV.
    """
    if df_lojas is None or df_lojas.empty:
        return pd.DataFrame()

    df_base = df_lojas.copy()

    if df_fila is not None and not df_fila.empty and "PV_Abadi" in df_fila.columns:
        df_base = _merge_base_por_pv(df_base, df_fila, "PV Abadi", "PV_Abadi")

    if df_inaug is not None and not df_inaug.empty and "PV ABADI" in df_inaug.columns:
        df_base = _merge_base_por_pv(df_base, df_inaug, "PV Abadi", "PV ABADI")

    # Limpa sufixos já existentes de bases antigas (_x/_y), mantendo o primeiro valor preenchido.
    for coluna in list(df_base.columns):
        if coluna.endswith('_x') and f"{coluna[:-2]}_y" in df_base.columns:
            base_col = coluna[:-2]
            y_col = f"{base_col}_y"
            if base_col not in df_base.columns:
                df_base[base_col] = pd.Series(
                    [pd.NA] * len(df_base),
                    index=df_base.index,
                    dtype="object"
                )
            else:
                df_base[base_col] = df_base[base_col].astype("object")

            origem_x = df_base[coluna].astype("object")
            origem_y = df_base[y_col].astype("object")

            destino = df_base[base_col].astype("object").copy()
            mask = ~destino.map(_valor_util)
            if mask.any():
                destino.loc[mask] = origem_x.loc[mask].tolist()

            mask2 = ~destino.map(_valor_util)
            if mask2.any():
                destino.loc[mask2] = origem_y.loc[mask2].tolist()

            df_base[base_col] = destino

    defaults = {
        "Status_Contato": "A Contatar",
        "Tipo_Necessidade": "Rede Ativa (Sem Pendência)",
        "Instrutor_Sugerido": "Pendente de Alocação",
        "Nome_Contato": "",
        "Material_Em_Loja": "Não Informado",
        "Tem_Funcionarios": "Não",
    }
    for coluna, valor in defaults.items():
        if coluna in df_base.columns:
            df_base[coluna] = df_base[coluna].fillna(valor)

    if "Qtd_Funcionarios" in df_base.columns:
        qtd = pd.to_numeric(df_base["Qtd_Funcionarios"], errors="coerce").fillna(0).clip(lower=0).astype(int)
        df_base["Qtd_Funcionarios"] = qtd
        df_base["Tem_Funcionarios"] = qtd.gt(0).map({True: "Sim", False: "Não"})

    return df_base



@st.cache_resource
def _supabase_client():
    """Cliente server-side. A chave secreta fica somente nos Secrets do Streamlit."""
    try:
        url = st.secrets.get("SUPABASE_URL", SUPABASE_PROJECT_URL_PADRAO)
        key = (
            st.secrets.get("SUPABASE_SERVICE_ROLE_KEY")
            or st.secrets.get("SUPABASE_SECRET_KEY")
        )
    except Exception:
        url = SUPABASE_PROJECT_URL_PADRAO
        key = None

    if not key:
        raise RuntimeError(
            "SUPABASE_SERVICE_ROLE_KEY não configurada nos Secrets do Streamlit."
        )
    return create_client(str(url), str(key))


def _valor_json_seguro(valor):
    if valor is None:
        return None
    try:
        vazio = pd.isna(valor)
        if isinstance(vazio, bool) and vazio:
            return None
    except Exception:
        pass
    if isinstance(valor, (pd.Timestamp, datetime, date)):
        return valor.isoformat()
    # numpy.datetime64 and similar pandas date scalars
    try:
        if pd.api.types.is_datetime64_any_dtype(type(valor)):
            return pd.Timestamp(valor).isoformat()
    except Exception:
        pass
    if hasattr(valor, "item"):
        try:
            valor = valor.item()
        except Exception:
            pass
    if isinstance(valor, float):
        if pd.isna(valor):
            return None
        if valor.is_integer():
            return int(valor)
    return valor


def _supabase_fetch_all(tabela, page_size=1000):
    client = _supabase_client()
    dados = []
    inicio = 0
    while True:
        resposta = (
            client.table(tabela)
            .select("*")
            .range(inicio, inicio + page_size - 1)
            .execute()
        )
        lote = resposta.data or []
        dados.extend(lote)
        if len(lote) < page_size:
            break
        inicio += page_size
    return dados


def _expandir_raw_data(registros):
    saida = []
    for item in registros or []:
        item = dict(item)
        raw = item.pop("raw_data", None)
        if isinstance(raw, dict):
            for chave, valor in raw.items():
                item.setdefault(chave, valor)
        saida.append(item)
    return saida


def _renomear_db_para_dataframe(registros, mapa):
    registros = _expandir_raw_data(registros)
    if not registros:
        return pd.DataFrame()
    df = pd.DataFrame(registros)
    return df.rename(columns=mapa)


MAPA_DB_PARA_DF = {
    "lojas": {
        "pv_abadi": "PV Abadi",
        "cnpj": "CNPJ",
        "razao_social": "Razão Social",
        "status_loja": "Status Loja",
        "grupo_economico": "Grupo Econômico",
        "gf": "GF",
        "cf": "CF",
        "endereco": "Endereço",
        "bairro": "Bairro",
        "municipio": "Municipio",
        "uf": "UF",
        "cep": "CEP",
        "nome_contato": "Nome_Contato",
        "telefone_contato": "Telefone_Contato",
        "email_contato": "Email_Contato",
        "data_inauguracao": "Data Inauguração Atual",
        "tipo_modelo": "Tipo de Modelo",
    },
    "fila": {
        "pv_abadi": "PV_Abadi",
        "tipo_necessidade": "Tipo_Necessidade",
        "data_ultimo_treinamento": "Data_Ultimo_Treinamento",
        "dias_desde_ultimo_treinamento": "Dias_desde_Ultimo_Treinamento",
        "instrutor_sugerido": "Instrutor_Sugerido",
        "semana_sugerida": "Semana_Sugerida",
        "telefone_contato": "Telefone_Contato",
        "email_contato": "Email_Contato",
        "status_contato": "Status_Contato",
        "data_do_contato": "Data_do_Contato",
        "observacoes": "Observacoes",
        "nome_contato": "Nome_Contato",
        "tem_funcionarios": "Tem_Funcionarios",
        "qtd_funcionarios": "Qtd_Funcionarios",
        "material_em_loja": "Material_Em_Loja",
        "data_agendada": "Data_Agendada",
        "tipo_pagamento": "Tipo_Pagamento",
        "data_pagamento": "Data_Pagamento",
        "data_liberacao_treinamento": "Data_Liberacao_Treinamento",
    },
    "inaug": {
        "pv_abadi": "PV ABADI",
        "previsao_inauguracao": "Previsão Inauguração",
        "instrutor_inauguracao": "Instrutor_Inauguracao",
    },
    "instrutores": {
        "id": "ID",
        "nome_completo": "NOME_COMPLETO",
        "status": "STATUS",
        "telefone": "TELEFONE",
        "email": "EMAIL",
        "cidade": "Cidade",
        "uf": "UF",
    },
    "rec": {
        "id": "ID",
        "pv_abadi": "PV_ABADI",
        "instrutor_sugerido": "Instrutor_Sugerido",
        "ranking_proximidade": "Ranking_Proximidade",
        "distancia_km_linha_reta": "Distancia_km_linha_reta",
        "municipio_loja": "Municipio_Loja",
        "uf_loja": "UF_Loja",
        "lat_loja": "Lat_Loja",
        "lon_loja": "Lon_Loja",
        "lat_instrutor": "Lat_Instrutor",
        "lon_instrutor": "Lon_Instrutor",
    },
}

MAPA_DF_PARA_DB = {
    entidade: {v: k for k, v in mapa.items()}
    for entidade, mapa in MAPA_DB_PARA_DF.items()
}


def carregar_bases_supabase():
    """Carrega todas as entidades do banco central em DataFrames compatíveis com o CRM."""
    bases = _bases_vazias()
    for entidade, tabela in SUPABASE_TABLES.items():
        registros = _supabase_fetch_all(tabela)
        bases[entidade] = _renomear_db_para_dataframe(
            registros,
            MAPA_DB_PARA_DF[entidade]
        )

    # Colunas operacionais sempre existem, mesmo em uma base ainda vazia.
    for col in COLUNAS_FILA:
        if col not in bases["fila"].columns:
            bases["fila"][col] = pd.NA

    return bases


def _registro_dataframe_para_db(linha, entidade):
    mapa = MAPA_DF_PARA_DB[entidade]
    registro = {}
    raw = {}

    for coluna, valor in linha.items():
        valor = _valor_json_seguro(valor)
        if coluna in mapa:
            campo_db = mapa[coluna]
            # IDs identity não são necessários para novos registros.
            if campo_db == "id" and valor is None:
                continue
            registro[campo_db] = valor
        else:
            # Guarda campos ainda não modelados sem perdê-los.
            if valor is not None:
                raw[str(coluna)] = valor

    if entidade in {"lojas", "fila", "inaug", "instrutores", "rec"}:
        registro["raw_data"] = raw

    return registro


def _upsert_dataframe_supabase(entidade, df, tamanho_lote=400):
    if df is None or df.empty:
        return

    tabela = SUPABASE_TABLES[entidade]
    client = _supabase_client()

    conflitos = {
        "lojas": "pv_abadi",
        "fila": "pv_abadi",
        "inaug": "pv_abadi",
        "instrutores": "nome_completo",
        "rec": "pv_abadi,instrutor_sugerido",
    }

    registros = []
    for _, linha in df.iterrows():
        registro = _registro_dataframe_para_db(linha.to_dict(), entidade)

        chave = conflitos[entidade].split(",")[0]
        if not registro.get(chave):
            continue

        # Em recomendações, instrutor também é parte da chave.
        if entidade == "rec" and not registro.get("instrutor_sugerido"):
            continue

        registros.append(registro)

    for inicio in range(0, len(registros), tamanho_lote):
        lote = registros[inicio:inicio + tamanho_lote]
        (
            client.table(tabela)
            .upsert(lote, on_conflict=conflitos[entidade])
            .execute()
        )



def salvar_importacao_supabase(bases_resultado, bases_novas):
    """
    Persiste somente as entidades efetivamente trazidas pela planilha.
    Evita regravar tabelas não relacionadas durante um upload.
    """
    salvas = []
    for entidade in ("lojas", "fila", "inaug", "instrutores", "rec"):
        novo = bases_novas.get(entidade, pd.DataFrame())
        if novo is None or novo.empty:
            continue

        resultado = bases_resultado.get(entidade, pd.DataFrame())
        if resultado is None or resultado.empty:
            continue

        # Para lojas, garante uma linha por PV antes do upsert.
        if entidade == "lojas" and "PV Abadi" in resultado.columns:
            resultado = (
                resultado
                .dropna(subset=["PV Abadi"])
                .drop_duplicates(subset=["PV Abadi"], keep="last")
                .copy()
            )

        # Para fila/inaug, também garante uma linha por PV.
        if entidade == "fila" and "PV_Abadi" in resultado.columns:
            resultado = (
                resultado
                .dropna(subset=["PV_Abadi"])
                .drop_duplicates(subset=["PV_Abadi"], keep="last")
                .copy()
            )
        if entidade == "inaug" and "PV ABADI" in resultado.columns:
            resultado = (
                resultado
                .dropna(subset=["PV ABADI"])
                .drop_duplicates(subset=["PV ABADI"], keep="last")
                .copy()
            )

        _upsert_dataframe_supabase(entidade, resultado)
        salvas.append(entidade)

    return salvas



def _extrair_contatos_dataframe(df_bruto):
    """Extrai TODOS os contatos de uma tabela com PV, sem colapsar um PV em uma única linha."""
    if df_bruto is None or df_bruto.empty:
        return pd.DataFrame(columns=["PV Abadi", "Nome_Contato", "Telefone_Contato", "Email_Contato"])

    preparado, _, canonicas, _ = _preparar_dataframe_entidade(
        df_bruto,
        ENTIDADES["lojas"],
    )

    if "PV Abadi" not in canonicas or "PV Abadi" not in preparado.columns:
        return pd.DataFrame(columns=["PV Abadi", "Nome_Contato", "Telefone_Contato", "Email_Contato"])

    for col in ["Nome_Contato", "Telefone_Contato", "Email_Contato"]:
        if col not in preparado.columns:
            preparado[col] = pd.NA

    contatos = preparado[
        ["PV Abadi", "Nome_Contato", "Telefone_Contato", "Email_Contato"]
    ].copy()

    contatos["PV Abadi"] = pd.to_numeric(
        contatos["PV Abadi"], errors="coerce"
    ).astype("Int64")

    contatos = contatos[contatos["PV Abadi"].notna()].copy()

    tem_contato = (
        contatos["Nome_Contato"].map(_valor_preenchido)
        | contatos["Telefone_Contato"].map(_valor_preenchido)
        | contatos["Email_Contato"].map(_valor_preenchido)
    )
    contatos = contatos[tem_contato].copy()

    # Remove apenas duplicatas EXATAS; múltiplos contatos diferentes do mesmo PV são preservados.
    contatos = contatos.drop_duplicates(
        subset=["PV Abadi", "Nome_Contato", "Telefone_Contato", "Email_Contato"],
        keep="last",
    ).reset_index(drop=True)

    return contatos


def _extrair_contatos_workbook(xls):
    partes = []
    for sheet_name in xls.sheet_names:
        try:
            bruto = pd.read_excel(xls, sheet_name=sheet_name)
        except Exception:
            continue
        parte = _extrair_contatos_dataframe(bruto)
        if parte is not None and not parte.empty:
            partes.append(parte)

    if not partes:
        return pd.DataFrame(columns=["PV Abadi", "Nome_Contato", "Telefone_Contato", "Email_Contato"])

    return (
        pd.concat(partes, ignore_index=True, sort=False)
        .drop_duplicates(
            subset=["PV Abadi", "Nome_Contato", "Telefone_Contato", "Email_Contato"],
            keep="last",
        )
        .reset_index(drop=True)
    )


def _chave_contato(pv, nome, telefone, email):
    texto = "|".join([
        str(pv or "").strip(),
        str(nome or "").strip().lower(),
        str(telefone or "").strip().lower(),
        str(email or "").strip().lower(),
    ])
    return hashlib.sha256(texto.encode("utf-8")).hexdigest()


def _upsert_contatos_supabase(df_contatos, tamanho_lote=500):
    if df_contatos is None or df_contatos.empty:
        return 0

    registros = []
    for _, row in df_contatos.iterrows():
        pv = row.get("PV Abadi")
        if not _valor_preenchido(pv):
            continue

        try:
            pv_txt = str(int(float(pv)))
        except Exception:
            pv_txt = str(pv).strip()

        nome = _valor_json_seguro(row.get("Nome_Contato"))
        telefone = _valor_json_seguro(row.get("Telefone_Contato"))
        email = _valor_json_seguro(row.get("Email_Contato"))

        if not any(_valor_preenchido(v) for v in [nome, telefone, email]):
            continue

        registros.append({
            "pv_abadi": pv_txt,
            "nome_contato": None if not _valor_preenchido(nome) else str(nome),
            "telefone": None if not _valor_preenchido(telefone) else str(telefone),
            "email": None if not _valor_preenchido(email) else str(email),
            "origem": "importador_inteligente",
            "dedupe_key": _chave_contato(pv_txt, nome, telefone, email),
            "raw_data": {},
        })

    client = _supabase_client()
    total = 0
    for inicio in range(0, len(registros), tamanho_lote):
        lote = registros[inicio:inicio + tamanho_lote]
        if not lote:
            continue
        (
            client.table("crm_contatos")
            .upsert(lote, on_conflict="dedupe_key")
            .execute()
        )
        total += len(lote)
    return total


def _contar_contatos_supabase():
    """Conta contatos únicos persistidos. Falha silenciosamente para não derrubar dashboard."""
    try:
        resposta = (
            _supabase_client()
            .table("crm_contatos")
            .select("id", count="exact")
            .limit(1)
            .execute()
        )
        return int(resposta.count or 0)
    except Exception:
        return 0


@st.cache_data(ttl=86400, show_spinner=False)
def _geocodificar_municipio(municipio, uf):
    """
    Geocodificação sob demanda da cidade do posto.
    Uma consulta por município/UF fica em cache por 24 horas.
    """
    municipio = str(municipio or "").strip()
    uf = str(uf or "").strip()
    if not municipio:
        return None, None

    try:
        resposta = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={
                "city": municipio,
                "state": uf,
                "country": "Brazil",
                "format": "jsonv2",
                "limit": 1,
            },
            headers={
                "User-Agent": "crm-ampm-operacional/1.0 (geocodificacao de municipios)"
            },
            timeout=8,
        )
        resposta.raise_for_status()
        dados = resposta.json()
        if dados:
            return float(dados[0]["lat"]), float(dados[0]["lon"])
    except Exception:
        pass

    # Fallback com consulta livre, útil quando o município está sem acentos.
    try:
        resposta = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={
                "q": f"{municipio}, {uf}, Brasil",
                "format": "jsonv2",
                "limit": 1,
            },
            headers={
                "User-Agent": "crm-ampm-operacional/1.0 (geocodificacao de municipios)"
            },
            timeout=8,
        )
        resposta.raise_for_status()
        dados = resposta.json()
        if dados:
            return float(dados[0]["lat"]), float(dados[0]["lon"])
    except Exception:
        pass

    return None, None


@st.cache_data(ttl=21600, show_spinner=False)
def _obter_rota_rodoviaria(lat_origem, lon_origem, lat_destino, lon_destino):
    """Rota rodoviária real via OSRM; retorna geometria, km e minutos."""
    try:
        url = (
            "https://router.project-osrm.org/route/v1/driving/"
            f"{float(lon_origem)},{float(lat_origem)};"
            f"{float(lon_destino)},{float(lat_destino)}"
        )
        r = requests.get(
            url,
            params={"overview":"full","geometries":"geojson","steps":"false"},
            timeout=12,
        )
        r.raise_for_status()
        rotas=(r.json().get("routes") or [])
        if not rotas:
            return None, None, None
        rota=rotas[0]
        coords=rota.get("geometry",{}).get("coordinates") or []
        if len(coords)<2:
            return None, None, None
        return coords, float(rota.get("distance",0))/1000, float(rota.get("duration",0))/60
    except Exception:
        return None, None, None




def _amostrar_pontos_rota(path, quantidade=14):
    """Retorna poucos pontos distribuídos ao longo da geometria da rota."""
    if not path or len(path) < 2:
        return []

    quantidade = max(2, int(quantidade))
    if len(path) <= quantidade:
        return path

    indices = {
        round(i * (len(path) - 1) / (quantidade - 1))
        for i in range(quantidade)
    }
    return [path[i] for i in sorted(indices)]


def _zoom_para_pontos(pontos):
    """Calcula um zoom aproximado para enquadrar todos os pontos no mapa."""
    if not pontos:
        return 4.0

    lats = [float(p[0]) for p in pontos]
    lons = [float(p[1]) for p in pontos]
    lat_span = max(lats) - min(lats)
    lon_span = max(lons) - min(lons)
    span = max(lat_span, lon_span)

    if span <= 0.10:
        return 10.5
    if span <= 0.30:
        return 9.0
    if span <= 0.70:
        return 7.5
    if span <= 1.50:
        return 6.5
    if span <= 3.00:
        return 5.5
    if span <= 6.00:
        return 4.5
    if span <= 12.00:
        return 3.8
    if span <= 25.00:
        return 3.1
    return 2.5


def _montar_rotas_mapa(top_instrutores):
    """
    Monta as rotas dos 3 instrutores para o posto.
    Usa OSRM quando disponível; se falhar, mantém uma linha reta visível.
    """
    if top_instrutores is None or top_instrutores.empty:
        return [], [], []

    cores = [
        [0, 150, 90, 230],
        [245, 140, 0, 230],
        [55, 105, 220, 230],
    ]

    rotas = []
    pontos = []
    resumo = []

    primeira = top_instrutores.iloc[0]
    p_lat = float(primeira["Lat_Loja"])
    p_lon = float(primeira["Lon_Loja"])

    pontos.append({
        "name": f"Posto {primeira.get('PV_ABADI', '')}",
        "tipo": "Posto",
        "lat": p_lat,
        "lon": p_lon,
        "cor": [220, 40, 40, 255],
    })

    for idx, (_, row) in enumerate(top_instrutores.iterrows()):
        try:
            i_lat = float(row["Lat_Instrutor"])
            i_lon = float(row["Lon_Instrutor"])
        except Exception:
            continue

        nome = str(row.get("Instrutor_Sugerido", f"Instrutor {idx + 1}"))
        cor = cores[min(idx, len(cores) - 1)]

        pontos.append({
            "name": nome,
            "tipo": f"Instrutor #{idx + 1}",
            "lat": i_lat,
            "lon": i_lon,
            "cor": cor,
        })

        rota_coords, rota_km, rota_min = _obter_rota_rodoviaria(
            i_lat, i_lon, p_lat, p_lon
        )

        if rota_coords:
            path = rota_coords
            tipo_rota = "Rodoviária"
            km_exibicao = rota_km
            min_exibicao = rota_min
        else:
            # A rota continua visível no mapa mesmo se o servidor OSRM estiver indisponível.
            path = [[i_lon, i_lat], [p_lon, p_lat]]
            tipo_rota = "Linha reta (fallback)"
            km_exibicao = float(row.get("Distancia_km_linha_reta", 0) or 0)
            min_exibicao = None

        rotas.append({
            "Instrutor": nome,
            "Ranking": idx + 1,
            "path": path,
            "cor": cor,
            "largura": max(5, 9 - idx * 2),
            "Tipo": tipo_rota,
            "Distancia_km": km_exibicao,
        })

        resumo.append({
            "Ranking": idx + 1,
            "Instrutor": nome,
            "Rota": tipo_rota,
            "Distância": km_exibicao,
            "Tempo_min": min_exibicao,
        })

    return rotas, pontos, resumo


def _haversine_km(lat1, lon1, lat2, lon2):
    import math
    r = 6371.0088
    p1 = math.radians(float(lat1))
    p2 = math.radians(float(lat2))
    dp = math.radians(float(lat2) - float(lat1))
    dl = math.radians(float(lon2) - float(lon1))
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def _top_instrutores_proximos(posto, df_instrutores, limite=3):
    if posto is None or df_instrutores is None or df_instrutores.empty:
        return pd.DataFrame()

    municipio = posto.get("Municipio")
    uf_loja = posto.get("UF")
    lat_loja, lon_loja = _geocodificar_municipio(municipio, uf_loja)

    if lat_loja is None or lon_loja is None:
        return pd.DataFrame()

    instr = df_instrutores.copy()

    # Somente instrutores atualmente ativos.
    status_col = "STATUS" if "STATUS" in instr.columns else ("status" if "status" in instr.columns else None)
    if status_col:
        mask_ativos = instr[status_col].astype(str).str.strip().str.lower().isin(
            ["ativo", "ativa", "em atividade", "active", "sim"]
        )
        instr = instr[mask_ativos].copy()

    candidatos = []
    for _, row in instr.iterrows():
        nome = (
            row.get("NOME_COMPLETO")
            or row.get("Nome Completo")
            or row.get("nome_completo")
            or row.get("nome")
        )
        cidade = row.get("Cidade") or row.get("cidade")
        uf = row.get("UF") or row.get("uf")

        lat = row.get("lat")
        lon = row.get("lon")
        if not _valor_preenchido(lat):
            lat = row.get("Latitude")
        if not _valor_preenchido(lon):
            lon = row.get("Longitude")

        if not (_valor_preenchido(nome) and _valor_preenchido(lat) and _valor_preenchido(lon)):
            continue

        try:
            dist = _haversine_km(lat_loja, lon_loja, float(lat), float(lon))
        except Exception:
            continue

        candidatos.append({
            "PV_ABADI": posto.get("PV Abadi"),
            "Razao_Social": posto.get("Razão Social"),
            "Municipio_Loja": municipio,
            "UF_Loja": uf_loja,
            "Lat_Loja": lat_loja,
            "Lon_Loja": lon_loja,
            "Instrutor_Sugerido": str(nome),
            "Cidade_Instrutor": cidade,
            "UF_Instrutor": uf,
            "Lat_Instrutor": float(lat),
            "Lon_Instrutor": float(lon),
            "Distancia_km_linha_reta": float(dist),
        })

    if not candidatos:
        return pd.DataFrame()

    resultado = pd.DataFrame(candidatos).sort_values(
        "Distancia_km_linha_reta"
    ).head(int(limite)).reset_index(drop=True)
    resultado["Ranking_Proximidade"] = range(1, len(resultado) + 1)
    return resultado

def salvar_bases_combinadas_no_disco(bases, caminho=CAMINHO_ARQUIVO):
    """
    Compatibilidade com o código legado: o nome da função foi mantido,
    mas a persistência agora é feita no Supabase/PostgreSQL.
    """
    for entidade in ("lojas", "fila", "inaug", "instrutores", "rec"):
        _upsert_dataframe_supabase(entidade, bases.get(entidade, pd.DataFrame()))


def salvar_fila_no_disco():
    try:
        _upsert_dataframe_supabase(
            "fila",
            st.session_state["bases"].get("fila", pd.DataFrame())
        )
        st.toast("💾 Call Center salvo no banco central!", icon="✅")
    except Exception as e:
        st.toast(f"⚠️ Erro ao salvar no banco central: {e}", icon="⚠️")


def salvar_lojas_no_disco():
    try:
        _upsert_dataframe_supabase(
            "lojas",
            st.session_state["bases"].get("lojas", pd.DataFrame())
        )
        st.toast("💾 Rede de Lojas salva no banco central!", icon="✅")
    except Exception as e:
        st.toast(f"⚠️ Erro ao salvar no banco central: {e}", icon="⚠️")


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

def calcular_liberacao_treinamento(data_pagamento):
    """Calcula a data mínima para liberar o treinamento: pagamento + 7 dias."""
    data = parse_data_flexivel(data_pagamento)
    if data is None:
        return None
    return data + pd.Timedelta(days=7).date()


def _data_exibicao_segura(valor):
    data = parse_data_flexivel(valor)
    return data.strftime("%d/%m/%Y") if data else ""


def treinamento_liberado(data_pagamento):
    """True somente a partir de 7 dias após o pagamento informado."""
    data_liberacao = calcular_liberacao_treinamento(data_pagamento)
    return bool(data_liberacao and date.today() >= data_liberacao)



def _sincronizar_modulos_apos_mutacao(entidade=None):
    """
    Mantém todos os módulos usando o mesmo estado após qualquer gravação.
    Dashboard, Pipeline, PROCV, Call Center e Calculadora são reconstruídos
    a partir de st.session_state['bases'] no rerun seguinte.
    """
    st.cache_data.clear()
    st.session_state["ultima_mutacao_entidade"] = entidade
    st.session_state["revisao_dados"] = int(st.session_state.get("revisao_dados", 0)) + 1


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
    _sincronizar_modulos_apos_mutacao("fila")

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

    # Produção: Supabase é a fonte oficial.
    try:
        st.session_state['bases'] = carregar_bases_supabase()
        st.session_state['erro_carga'] = None
        st.session_state['fonte_dados'] = "Supabase"
        return
    except Exception as e_supabase:
        # Fallback temporário para não derrubar o app enquanto os Secrets
        # ainda não foram configurados no Streamlit Cloud.
        st.session_state['erro_carga'] = (
            "Banco central indisponível: " + str(e_supabase)
        )

    if os.path.exists(CAMINHO_ARQUIVO):
        try:
            assinatura = os.path.getmtime(CAMINHO_ARQUIVO)
            bases, relatorio = carregar_bases_do_disco(
                CAMINHO_ARQUIVO,
                assinatura
            )
            st.session_state['bases'] = (
                bases if bases is not None else _bases_vazias()
            )
            st.session_state['relatorio_importacao'] = relatorio
            st.session_state['fonte_dados'] = "Excel fallback"
            return
        except Exception as e_excel:
            st.session_state['erro_carga'] += " | Fallback Excel: " + str(e_excel)

    st.session_state['bases'] = _bases_vazias()
    st.session_state['fonte_dados'] = "Nenhuma"

inicializar_estado()

# ------------------------------------------------------------
# VISOES DERIVADAS: mantem todos os modulos usando a mesma base
# consolidada. Esta etapa precisa existir antes do despacho dos
# modulos (Dashboard, Pipeline, PROCV, Call Center etc.).
# ------------------------------------------------------------
_bases_atuais_ui = st.session_state.get("bases", _bases_vazias())
df_lojas = _bases_atuais_ui.get("lojas", pd.DataFrame()).copy()
df_fila = _bases_atuais_ui.get("fila", pd.DataFrame(columns=COLUNAS_FILA)).copy()
df_inaug = _bases_atuais_ui.get("inaug", pd.DataFrame()).copy()
df_instrutores = _bases_atuais_ui.get("instrutores", pd.DataFrame()).copy()
df_rec = _bases_atuais_ui.get("rec", pd.DataFrame()).copy()

df_base = construir_base_unificada(df_lojas, df_fila, df_inaug)
# O calculador usa este filtro opcional; por padrao mostra toda a rede.
filtro_uf = "Todas"

if st.session_state.get('erro_carga'):
    if st.session_state.get("fonte_dados") == "Excel fallback":
        st.warning(
            "⚠️ O CRM está usando temporariamente o Excel porque o banco central "
            "ainda não está configurado nos Secrets.\n\n"
            f"{st.session_state['erro_carga']}"
        )
    else:
        st.error(
            "⚠️ Não foi possível carregar o banco central.\n\n"
            f"{st.session_state['erro_carga']}"
        )


# ============================================================
# v21 — IMPORTADOR INTELIGENTE DE PLANILHAS
# ============================================================
IMPORTADOR_CAMPOS = {
    "PV_Abadi": ["pv","pv abadi","pv_abadi","codigo pv","codigo loja","numero pv","nº pv","n pv","posto","codigo posto"],
    "CNPJ": ["cnpj","cnpj loja","cnpj da loja","cnpj posto","documento cnpj","cnpj do posto"],
    "Razão Social": ["razao social","razão social","empresa","nome empresa","nome da empresa","razao_social"],
    "Nome_Contato": ["nome contato","nome do contato","contato","responsavel","responsável","nome responsavel","nome responsável"],
    "Telefone_Contato": ["telefone","telefone contato","telefone do contato","celular","whatsapp","fone","telefone celular"],
    "Email_Contato": ["email","e mail","e-mail","email contato","e mail contato","e-mail contato","correio eletronico","correio eletrônico"],
    "Qtd_Funcionarios": ["qtd funcionarios","qtd funcionários","quantidade funcionarios","quantidade funcionários","numero funcionarios","n funcionarios","funcionarios","funcionários","qtd de funcionarios"],
    "Tem_Funcionarios": ["tem funcionarios","tem funcionários","ha funcionarios","há funcionários","possui funcionarios","possui funcionários","funcionarios para treinar","funcionários para treinar"],
    "Instrutor_Treinamento": ["instrutor","instrutor treinamento","instrutor do treinamento","instrutor treinou","instrutor responsavel","instrutor responsável"],
    "Instrutor_Inauguracao": ["instrutor inauguracao","instrutor inauguração","instrutor da inauguracao","instrutor da inauguração","instrutor inaugurou","responsavel inauguracao"],
    "Municipio": ["municipio","município","cidade","cidade loja","localidade"],
    "UF": ["uf","estado","sigla estado","estado uf"],
    "Tipo_Necessidade": ["tipo necessidade","necessidade","tipo de necessidade","motivo","tipo"],
    "Data_Ultimo_Treinamento": ["data ultimo treinamento","data último treinamento","ultimo treinamento","último treinamento","data treinamento"],
    "Status_Contato": ["status contato","status do contato","status atendimento","status","situacao contato","situação contato"],
    "Data_do_Contato": ["data contato","data do contato","ultimo contato","último contato"],
    "Observacoes": ["observacoes","observações","observacao","observação","comentarios","comentários"],
    "Material_Em_Loja": ["material em loja","material loja","apostila","materiais"],
    "Data_Agendada": ["data agendada","data agendamento","agendamento","data treinamento agendado"],
}

def _imp_norm(valor):
    import unicodedata
    s = "" if valor is None else str(valor).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def _imp_similarity(a,b):
    a,b=_imp_norm(a),_imp_norm(b)
    if not a or not b:return 0.0
    if a==b:return 1.0
    if a in b or b in a:return 0.90
    aa,bb=set(a.split()),set(b.split())
    if aa and bb:
        u=len(aa|bb); i=len(aa&bb)
        if u:return 0.55+0.35*(i/u)
    prev=list(range(len(b)+1))
    for i,ca in enumerate(a,1):
        cur=[i]
        for j,cb in enumerate(b,1):cur.append(min(cur[-1]+1,prev[j]+1,prev[j-1]+(ca!=cb)))
        prev=cur
    return max(0.0,1.0-prev[-1]/max(len(a),len(b),1))

def _imp_sugerir_campo(coluna):
    melhor=(None,0.0)
    for campo,aliases in IMPORTADOR_CAMPOS.items():
        for alias in [campo]+aliases:
            score=_imp_similarity(coluna,alias)
            if score>melhor[1]:melhor=(campo,score)
    return melhor

def _imp_ler_planilha(arquivo):
    xls=pd.ExcelFile(arquivo); abas={}
    for aba in xls.sheet_names:
        try:
            df=pd.read_excel(arquivo,sheet_name=aba).dropna(axis=0,how="all").dropna(axis=1,how="all")
            if not df.empty:abas[aba]=df
        except Exception:pass
    return abas

def _imp_normalizar_valor(campo,valor):
    if pd.isna(valor):return pd.NA
    if campo=="Qtd_Funcionarios":
        try:return max(0,int(float(valor)))
        except Exception:return 0
    if campo=="Tem_Funcionarios":
        s=_imp_norm(valor)
        if s in {"sim","s","yes","y","true","1","tem","possui"}:return "Sim"
        if s in {"nao","não","n","no","false","0","nao possui","não possui"}:return "Não"
    return valor

def _imp_chave_linha(linha,mapeamento):
    pv_col=next((c for c,f in mapeamento.items() if f=="PV_Abadi"),None)
    cnpj_col=next((c for c,f in mapeamento.items() if f=="CNPJ"),None)
    if pv_col is not None and not pd.isna(linha.get(pv_col)):
        v=str(linha.get(pv_col)).strip()
        if v and v.lower() not in {"nan","none"}:return "PV",_imp_norm(v)
    if cnpj_col is not None and not pd.isna(linha.get(cnpj_col)):
        v=re.sub(r"\D","",str(linha.get(cnpj_col)))
        if v:return "CNPJ",v
    return None,None

def _imp_processar_df(df,mapeamento):
    registros=[]
    for _,linha in df.iterrows():
        kt,k=_imp_chave_linha(linha,mapeamento); dados={}
        for col,campo in mapeamento.items():
            if not campo or col not in df.columns:continue
            v=_imp_normalizar_valor(campo,linha.get(col))
            if campo=="Qtd_Funcionarios":dados[campo]=v;dados["Tem_Funcionarios"]="Sim" if int(v or 0)>0 else "Não"
            elif campo=="Tem_Funcionarios" and "Qtd_Funcionarios" not in dados:dados[campo]=v
            else:dados[campo]=v
        registros.append({"chave_tipo":kt,"chave":k,"dados":dados})
    return registros

def _imp_merge_com_base(base_df,registros):
    df=base_df.copy() if isinstance(base_df,pd.DataFrame) else pd.DataFrame(); ins=atu=0
    def localizar(tipo,chave):
        if not chave:return None
        if tipo=="PV" and "PV_Abadi" in df.columns:
            s=df["PV_Abadi"].astype(str).map(_imp_norm); x=s[s==chave].index
            if len(x):return x[0]
        if tipo=="CNPJ" and "CNPJ" in df.columns:
            s=df["CNPJ"].astype(str).map(lambda x:re.sub(r"\D","",x)); x=s[s==chave].index
            if len(x):return x[0]
        return None
    for reg in registros:
        if not reg["chave"]:continue
        idx=localizar(reg["chave_tipo"],reg["chave"]); dados=reg["dados"]
        if idx is None:
            novo={c:pd.NA for c in df.columns}
            for c,v in dados.items():
                if c not in df.columns:df[c]=pd.NA
                novo[c]=v
            df=pd.concat([df,pd.DataFrame([novo])],ignore_index=True);ins+=1
        else:
            for c,v in dados.items():
                if c not in df.columns:df[c]=pd.NA
                if not pd.isna(v) and str(v).strip() not in {"","nan","None"}:df.at[idx,c]=v
            atu+=1
    if "Qtd_Funcionarios" in df.columns:
        q=pd.to_numeric(df["Qtd_Funcionarios"],errors="coerce").fillna(0).clip(lower=0).astype(int)
        df["Qtd_Funcionarios"]=q;df["Tem_Funcionarios"]=q.gt(0).map({True:"Sim",False:"Não"})
    return df,ins,atu


def _formatar_conf_importador(valor):
    return str(valor or "").strip().lower()


def _resumo_importacao_integrada(estatisticas):
    nomes = {
        "lojas": "🏪 Rede de Lojas",
        "fila": "📞 Fila Call Center",
        "inaug": "🚀 Previsão de Inauguração",
        "instrutores": "👔 Instrutores",
        "rec": "📍 Recomendação de Deslocamento",
    }
    totais = {
        "novos": sum(int(x.get("adicionados", 0) or 0) for x in estatisticas),
        "atualizados": sum(int(x.get("atualizados", 0) or 0) for x in estatisticas),
    }
    return nomes, totais


def render_importador_inteligente():
    st.markdown("## 📥 Importador Inteligente de Planilhas")
    st.caption(
        "Um único ponto de entrada para planilhas novas. "
        "O CRM identifica abas, reconhece colunas, mescla por PV/CNPJ e atualiza o banco central."
    )

    arquivo = st.file_uploader(
        "📤 Envie a planilha",
        type=["xlsx", "xls", "csv"],
        key="importador_inteligente_principal",
        help="O importador pode receber planilhas gerenciais, contatos, treinamentos e outras bases."
    )

    if not arquivo:
        st.info(
            "Envie uma planilha para começar. "
            "Nada será alterado até você revisar o diagnóstico e clicar em Confirmar."
        )
        return

    conteudo = arquivo.getvalue()
    assinatura = _fingerprint_upload(conteudo) if "_fingerprint_upload" in globals() else str(len(conteudo))

    try:
        contatos_novos = pd.DataFrame(
            columns=["PV Abadi", "Nome_Contato", "Telefone_Contato", "Email_Contato"]
        )

        aba_instrutores_importada = None
        instrutores_importados = pd.DataFrame()

        if arquivo.name.lower().endswith(".csv"):
            df_csv = _ler_csv_flexivel(arquivo)
            contatos_novos = _extrair_contatos_dataframe(df_csv)

            candidatos = []
            for entidade, definicao in ENTIDADES.items():
                score, canonicas = _score_aba_para_entidade(
                    df_csv, arquivo.name, entidade, definicao
                )
                if definicao["chave"] in canonicas:
                    candidatos.append((score, entidade))

            candidatos.sort(reverse=True)

            if candidatos and candidatos[0][0] >= MIN_SCORE_CONFIANTE:
                entidade_escolhida = candidatos[0][1]
            else:
                entidade_escolhida = st.selectbox(
                    "Destino da planilha CSV",
                    list(ENTIDADES.keys()),
                    format_func=lambda x: {
                        "lojas": "🏪 Rede de Lojas",
                        "fila": "📞 Fila Call Center",
                        "inaug": "🚀 Previsão de Inauguração",
                        "instrutores": "👔 Instrutores",
                        "rec": "📍 Recomendação de Deslocamento",
                    }.get(x, x)
                )

            definicao = ENTIDADES[entidade_escolhida]
            preparado, _, canonicas, colunas_novas = _preparar_dataframe_entidade(
                df_csv, definicao
            )

            if definicao["chave"] not in canonicas:
                st.error(
                    f"❌ Não encontrei a chave '{definicao['chave']}' na planilha."
                )
                return

            bases_novas = _bases_vazias()
            bases_novas[entidade_escolhida] = preparado

            st.markdown("### 🔎 Diagnóstico da importação")
            st.write(
                f"**Arquivo:** `{arquivo.name}`  \n"
                f"**Linhas:** {len(df_csv):,}  \n"
                f"**Colunas reconhecidas:** {len(canonicas)}  \n"
                f"**Colunas novas preservadas:** {len(colunas_novas)}"
            )

        else:
            xls = _abrir_excel_resiliente(conteudo)
            contatos_novos = _extrair_contatos_workbook(xls)
            aba_instrutores_importada, instrutores_importados = _extrair_instrutores_explicitos(xls)
            bases_novas, relatorio = _processar_excelfile(
                xls,
                exigir_lojas=False
            )

            # Rejeita apenas um workbook que não tenha nenhuma entidade útil.
            entidades_com_dados = [
                item for item in relatorio
                if item.get("aba_origem") and item.get("linhas_lidas", 0) > 0
            ]

            if not entidades_com_dados:
                st.error("❌ Nenhuma aba útil foi reconhecida nessa planilha.")
                return

            st.markdown("### 🔎 Diagnóstico da importação")
            for item in relatorio:
                nome = {
                    "lojas": "🏪 Rede de Lojas",
                    "fila": "📞 Fila Call Center",
                    "inaug": "🚀 Previsão de Inauguração",
                    "instrutores": "👔 Instrutores",
                    "rec": "📍 Recomendação de Deslocamento",
                }.get(item["entidade"], item["entidade"])

                if item.get("aba_origem"):
                    col_rec = item.get("colunas_reconhecidas", [])
                    col_new = item.get("colunas_novas", [])
                    conf = item.get("confianca", "n/a")
                    st.markdown(
                        f"**{nome}** → aba `{item['aba_origem']}` · "
                        f"confiança **{conf}** · "
                        f"{item.get('linhas_lidas', 0):,} linhas"
                    )
                    if col_rec:
                        st.caption("Colunas reconhecidas: " + ", ".join(col_rec))
                    if col_new:
                        st.caption("🆕 Colunas novas preservadas: " + ", ".join(map(str, col_new)))

        # Prévia do impacto ANTES de gravar.
        bases_atuais = st.session_state.get("bases", _bases_vazias())
        bases_teste, estatisticas = mesclar_bases(bases_atuais, bases_novas)

        total_novos = sum(int(x.get("adicionados", 0) or 0) for x in estatisticas)
        total_atualizados = sum(int(x.get("atualizados", 0) or 0) for x in estatisticas)

        st.markdown("### 📊 Impacto previsto")
        c1, c2 = st.columns(2)
        c1.metric("🆕 Novos registros", f"{total_novos:,}")
        c2.metric("🔄 Registros atualizados", f"{total_atualizados:,}")

        st.info(
            "Nenhum dado será salvo até a confirmação. "
            "O merge usa PV como chave principal, preserva valores existentes quando a nova célula está vazia "
            "e mantém colunas adicionais."
        )

        confirmar = st.checkbox(
            "✅ Revisei o diagnóstico e autorizo a integração desta planilha.",
            key=f"confirmar_importacao_{assinatura}"
        )

        if not confirmar:
            return

        if st.button(
            "🚀 Confirmar e integrar ao banco",
            type="primary",
            use_container_width=True,
            key=f"confirmar_integracao_v32_{assinatura}"
        ):
            with st.spinner("Gravando no Supabase e verificando os dados..."):
                # Salva apenas entidades realmente presentes no arquivo.
                entidades_salvas = salvar_importacao_supabase(
                    bases_teste,
                    bases_novas
                )

                contatos_gravados = _upsert_contatos_supabase(contatos_novos)

                instrutores_substituidos = 0
                if (
                    aba_instrutores_importada
                    and instrutores_importados is not None
                    and not instrutores_importados.empty
                ):
                    instrutores_substituidos = _substituir_instrutores_supabase(
                        instrutores_importados
                    )
                    if "instrutores" not in entidades_salvas:
                        entidades_salvas.append("instrutores")

                if not entidades_salvas:
                    raise RuntimeError(
                        "Nenhuma entidade com dados foi encontrada para gravar."
                    )

                # Releitura obrigatória do banco central.
                st.cache_data.clear()
                bases_recarregadas = carregar_bases_supabase()
                lojas_db = bases_recarregadas.get("lojas", pd.DataFrame())

                def _contar_preenchidos(df, coluna):
                    if df is None or df.empty or coluna not in df.columns:
                        return 0
                    return int(df[coluna].map(_valor_preenchido).sum())

                total_db = len(lojas_db)
                cnpj_db = _contar_preenchidos(lojas_db, "CNPJ")
                tel_db = _contar_preenchidos(lojas_db, "Telefone_Contato")
                email_db = _contar_preenchidos(lojas_db, "Email_Contato")
                contato_db = _contar_preenchidos(lojas_db, "Nome_Contato")

                # Teste adicional com os PVs do arquivo enviado.
                pvs_arquivo = set()
                lojas_novas = bases_novas.get("lojas", pd.DataFrame())
                if (
                    lojas_novas is not None
                    and not lojas_novas.empty
                    and "PV Abadi" in lojas_novas.columns
                ):
                    pvs_arquivo = {
                        str(v).strip()
                        for v in lojas_novas["PV Abadi"].dropna().tolist()
                        if _valor_preenchido(v)
                    }

                pvs_banco = set()
                if not lojas_db.empty and "PV Abadi" in lojas_db.columns:
                    pvs_banco = {
                        str(v).strip()
                        for v in lojas_db["PV Abadi"].dropna().tolist()
                        if _valor_preenchido(v)
                    }

                pvs_confirmados = len(pvs_arquivo & pvs_banco)

                st.session_state["bases"] = bases_recarregadas
                st.session_state["ultimo_importador_assinatura"] = assinatura
                st.session_state["ultimo_importador_relatorio"] = estatisticas
                st.session_state["erro_carga"] = None
                st.session_state["fonte_dados"] = "Supabase"

                st.success("✅ Importação gravada e conferida diretamente no Supabase.")
                total_contatos_db = _contar_contatos_supabase()
                st.info(
                    f"🗄️ Banco central: {total_db:,} lojas únicas · "
                    f"CNPJ: {cnpj_db:,} · Telefones principais: {tel_db:,} · "
                    f"E-mails principais: {email_db:,} · "
                    f"Contatos únicos: {total_contatos_db:,}"
                )
                st.caption(
                    f"Linhas de contato reconhecidas neste arquivo: {len(contatos_novos):,} · "
                    f"enviadas ao Supabase nesta operação: {contatos_gravados:,}"
                )
                st.caption(
                    f"Entidades gravadas: {', '.join(entidades_salvas)} · "
                    f"PVs do arquivo confirmados no banco: "
                    f"{pvs_confirmados:,}/{len(pvs_arquivo):,}"
                )
                if instrutores_substituidos:
                    st.success(
                        f"👔 Equipe oficial atualizada pela aba "
                        f"`{aba_instrutores_importada}`: "
                        f"{instrutores_substituidos:,} instrutor(es)."
                    )

                # Auditoria da importação.
                try:
                    _supabase_client().table("crm_importacoes").insert({
                        "arquivo": arquivo.name,
                        "aba": None,
                        "registros_lidos": int(
                            sum(
                                int(x.get("linhas_lidas", 0) or 0)
                                for x in relatorio
                            )
                        ) if "relatorio" in locals() else int(len(lojas_novas)),
                        "novos": int(total_novos),
                        "atualizados": int(total_atualizados),
                        "status": "concluida",
                        "detalhes": {
                            "entidades_salvas": entidades_salvas,
                            "pvs_arquivo": len(pvs_arquivo),
                            "pvs_confirmados": pvs_confirmados,
                            "total_lojas_db": total_db,
                            "cnpj_db": cnpj_db,
                            "telefones_db": tel_db,
                            "emails_db": email_db,
                        },
                    }).execute()
                except Exception:
                    pass

                # Não faz rerun automático: mantém os números visíveis.

    except Exception as exc:
        st.error(f"❌ Falha na importação: {exc}")


# --- CONTROLE DO MENU SUPERIOR DO STREAMLIT ---
# V53: GitHub/source/deploy ficam ocultos para todos os perfis.
# A regra é aplicada no início do arquivo para funcionar também no login.

# --- SIDEBAR & NAVEGAÇÃO ---
with st.sidebar:
    st.caption(f"🗄️ Fonte de dados: {st.session_state.get('fonte_dados', '-')}")
    if st.button("🔄 Recarregar dados do Supabase", use_container_width=True):
        try:
            st.cache_data.clear()
            st.session_state["bases"] = carregar_bases_supabase()
            st.session_state["fonte_dados"] = "Supabase"
            st.session_state["erro_carga"] = None
            st.rerun()
        except Exception as _e:
            st.error(f"Falha ao recarregar Supabase: {_e}")
    st.markdown("""
        <div class="sidebar-brand">
            <div>
                <div class="sidebar-ampm">ampm☀</div>
                <div style="font-weight:750; font-size:.76rem; color:#F0F1F3;">CRM OPERACIONAL</div>
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

    opcoes_modulos = [
        (chave, nome) for chave, nome in MODULOS_PERMISSOES.items()
        if usuario_tem_permissao(chave)
    ]
    if usuario_e_admin() or usuario_tem_permissao("procv") or usuario_tem_permissao("administracao"):
        opcoes_modulos.append(("importador_inteligente", "📥 Importador Inteligente"))

    if usuario_e_admin():
        opcoes_modulos.append(("administracao", "🛡️ Administração"))

    if not opcoes_modulos:
        st.warning("🔒 Seu usuário ainda não possui módulos liberados.")
        st.stop()

    modulo = st.radio(
        "📌 **Módulos do Sistema:**",
        [nome for _, nome in opcoes_modulos]
    )

    st.markdown("""
        <div class="sidebar-igt">
            <div class="sidebar-igt-logo">igt <span>IGT GROUP</span></div>
            <p>Excelência em Treinamentos.<br>Resultados que transformam.</p>
        </div>
    """, unsafe_allow_html=True)

    st.divider()


def _procv_valor_flexivel(posto, campo, aliases=()):
    """Busca valor canônico ou equivalente sem quebrar com NA/NaN."""
    normalizados = {}
    for chave in posto.keys():
        try:
            normalizados[_normalizar_nome(str(chave))] = chave
        except Exception:
            pass

    for candidato in [campo] + list(aliases):
        if candidato in posto:
            valor = posto.get(candidato)
            if _valor_preenchido(valor):
                return valor

        real = normalizados.get(_normalizar_nome(candidato))
        if real is not None:
            valor = posto.get(real)
            if _valor_preenchido(valor):
                return valor

    return "Não informado"


def _valor_historico_instrutor(posto, *campos):
    for campo in campos:
        valor = _procv_valor_flexivel(posto, campo)
        if valor != "Não informado":
            return str(valor).strip()
    return "Não informado"


def _historico_instrutor_procv(posto):
    """Histórico informativo, inclusive instrutores inativos."""
    treinamento = _valor_historico_instrutor(
        posto, "Instrutor_Treinamento", "Instrutor_Sugerido"
    )
    inauguracao = _valor_historico_instrutor(
        posto, "Instrutor_Inauguracao"
    )
    return treinamento, inauguracao


# --- CABEÇALHO VISUAL GLOBAL AMPM + IGT ---
_nome_topo = str(st.session_state.get("name") or _usuario_atual() or "Usuário")
st.markdown(
    f"""
    <div class="brand-topbar">
        <div class="brand-top-left">
            <div class="ampm-wordmark">ampm☀</div>
            <div class="brand-divider"></div>
            <div class="brand-product">
                <strong>CRM OPERACIONAL</strong>
                <span>AmPm</span>
            </div>
        </div>
        <div class="brand-top-right">
            <div class="igt-wordmark">igt<small>IGT GROUP</small></div>
            <div class="top-user-chip">👤 {_nome_topo}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

if modulo == "🛡️ Administração":
    render_administracao()

elif modulo == "📊 Dashboard Executivo":
    if df_base.empty:
        st.info("📭 Nenhum dado carregado ainda.")
    else:
        # --------------------------------------------------------
        # TÍTULO E FILTROS
        # --------------------------------------------------------
        _render_html_dashboard(
            """
            <div class="dashboard-title">
                <div class="dashboard-title-icon">📊</div>
                <h1>Dashboard Executivo</h1>
            </div>
            <div class="dashboard-subtitle">
                Visão geral do desempenho operacional dos treinamentos
            </div>
            """,
        )

        df_dash = df_base.copy()

        # Modelo real da loja, quando disponível.
        coluna_modelo = next(
            (
                c for c in [
                    "Tipo de Modelo", "Tipo Modelo", "Modelo da Loja",
                    "Modelo Loja", "Modelo"
                ]
                if c in df_dash.columns
            ),
            None,
        )

        filtros = st.columns([1.1, 1.1, 1.1, .85])
        with filtros[0]:
            ufs_dash = ["Todos os Estados"]
            if "UF" in df_dash.columns:
                ufs_dash += sorted(
                    [
                        str(x) for x in df_dash["UF"].dropna().unique()
                        if str(x).strip()
                    ]
                )
            uf_dash = st.selectbox(
                "Estado",
                ufs_dash,
                label_visibility="collapsed",
                key="dash_filtro_uf_v49",
            )

        with filtros[1]:
            modelos_dash = ["Todos os Modelos"]
            if coluna_modelo:
                modelos_dash += sorted(
                    [
                        str(x) for x in df_dash[coluna_modelo].dropna().unique()
                        if str(x).strip()
                    ]
                )
            modelo_dash = st.selectbox(
                "Modelo",
                modelos_dash,
                label_visibility="collapsed",
                key="dash_filtro_modelo_v49",
            )

        with filtros[2]:
            necessidades_dash = ["Todas as Necessidades"]
            if "Tipo_Necessidade" in df_dash.columns:
                necessidades_dash += sorted(
                    [
                        str(x)
                        for x in df_dash["Tipo_Necessidade"].dropna().unique()
                        if str(x).strip()
                    ]
                )
            necessidade_dash = st.selectbox(
                "Necessidade",
                necessidades_dash,
                label_visibility="collapsed",
                key="dash_filtro_nec_v49",
            )

        with filtros[3]:
            st.markdown(
                "<div style='height:2px'></div>",
                unsafe_allow_html=True,
            )
            st.caption("📡 Dados em tempo real")

        if uf_dash != "Todos os Estados" and "UF" in df_dash.columns:
            df_dash = df_dash[df_dash["UF"].astype(str) == uf_dash].copy()
        if (
            modelo_dash != "Todos os Modelos"
            and coluna_modelo
        ):
            df_dash = df_dash[
                df_dash[coluna_modelo].astype(str) == modelo_dash
            ].copy()
        if (
            necessidade_dash != "Todas as Necessidades"
            and "Tipo_Necessidade" in df_dash.columns
        ):
            df_dash = df_dash[
                df_dash["Tipo_Necessidade"].astype(str) == necessidade_dash
            ].copy()

        # --------------------------------------------------------
        # KPIs
        # --------------------------------------------------------
        total_postos = len(df_dash)
        agendados = (
            int((df_dash["Status_Contato"].astype(str) == "Agendado").sum())
            if "Status_Contato" in df_dash.columns else 0
        )
        concluidos = (
            int(
                df_dash["Status_Contato"].astype(str).isin(
                    ["Treinamento Realizado", "Concluído", "Concluido"]
                ).sum()
            )
            if "Status_Contato" in df_dash.columns else 0
        )
        taxa_conclusao = (
            (concluidos / total_postos * 100.0)
            if total_postos else 0.0
        )
        clientes_ativos = (
            int(
                df_dash["Status Loja"]
                .astype(str)
                .str.contains("ativ", case=False, na=False)
                .sum()
            )
            if "Status Loja" in df_dash.columns
            else total_postos
        )

        k1, k2, k3, k4, k5 = st.columns(5)
        kpis = [
            (k1, "🏪", "dk-orange", "TOTAL DE POSTOS", total_postos, "PVs na visão atual"),
            (k2, "📅", "dk-yellow", "TREINAMENTOS AGENDADOS", agendados, "agenda operacional"),
            (k3, "✅", "dk-blue", "TREINAMENTOS CONCLUÍDOS", concluidos, "treinamentos realizados"),
            (k4, "🎯", "dk-cyan", "TAXA DE CONCLUSÃO", f"{taxa_conclusao:.1f}%", "concluídos / total"),
            (k5, "👥", "dk-purple", "CLIENTES ATIVOS", clientes_ativos, "status ativo na rede"),
        ]
        for col, icon, cls, label, value, note in kpis:
            with col:
                _render_html_dashboard(
                    f"""
                    <div class="dashboard-kpi">
                        <div class="dashboard-kpi-top">
                            <div class="dashboard-kpi-icon {cls}">{icon}</div>
                            <div class="dashboard-kpi-label">{label}</div>
                        </div>
                        <div class="dashboard-kpi-value">{value}</div>
                        <div class="dashboard-kpi-note">{note}</div>
                    </div>
                    """,
                )

        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

        # --------------------------------------------------------
        # STATUS / MODELO / ATIVIDADES
        # --------------------------------------------------------
        row_a = st.columns([1.25, 1.05, 1.15])

        status_cores = {
            "A Contatar": "#1976D2",
            "Aguardando Pagamento": "#FFD300",
            "Em Negociação": "#FF8A00",
            "Agendado": "#FF4D00",
            "Treinamento Realizado": "#149B55",
            "Recusado": "#9AA0A6",
        }
        status_counts = {}
        if "Status_Contato" in df_dash.columns:
            for status in status_cores:
                status_counts[status] = int(
                    (df_dash["Status_Contato"].astype(str) == status).sum()
                )
        else:
            status_counts = {s: 0 for s in status_cores}

        total_status = sum(status_counts.values()) or 1
        inicio = 0.0
        partes_gradiente = []
        legenda_html = []
        for status, cor in status_cores.items():
            valor = status_counts.get(status, 0)
            pct = valor / total_status * 100
            fim = inicio + pct
            partes_gradiente.append(
                f"{cor} {inicio:.2f}% {fim:.2f}%"
            )
            legenda_html.append(
                f"""
                <div class="donut-legend-row">
                    <span class="legend-dot" style="background:{cor}"></span>
                    <span>{status}</span>
                    <strong>{valor}</strong>
                </div>
                """
            )
            inicio = fim

        donut_gradient = ", ".join(partes_gradiente) if partes_gradiente else "#E6E8EC 0 100%"

        with row_a[0]:
            _render_html_dashboard(
                f"""
                <div class="dashboard-panel">
                    <div class="dashboard-panel-title">Treinamentos por Status (Pipeline)</div>
                    <div class="donut-wrap">
                        <div class="dashboard-donut"
                             style="background:conic-gradient({donut_gradient});">
                            <div class="donut-center">
                                <strong>{sum(status_counts.values())}</strong>
                                <span>pipeline</span>
                            </div>
                        </div>
                        <div class="donut-legend">
                            {''.join(legenda_html)}
                        </div>
                    </div>
                </div>
                """,
            )

        with row_a[1]:
            if coluna_modelo:
                modelo_counts = (
                    df_dash[coluna_modelo]
                    .fillna("Não informado")
                    .astype(str)
                    .value_counts()
                    .head(5)
                )
            else:
                modelo_counts = pd.Series(dtype=int)

            max_modelo = int(modelo_counts.max()) if not modelo_counts.empty else 1
            barras = []
            cores_modelo = ["#FF4D00", "#FFD300", "#1976D2", "#149B55", "#8E44E8"]
            for idx, (nome_modelo, qtd) in enumerate(modelo_counts.items()):
                largura = max(3.0, float(qtd) / max_modelo * 100)
                cor = cores_modelo[idx % len(cores_modelo)]
                barras.append(
                    f'<div class="dashboard-model-row">'
                    f'<span>{html.escape(str(nome_modelo))}</span>'
                    f'<div class="dashboard-progress">'
                    f'<span style="width:{largura:.1f}%;background:{cor};"></span>'
                    f'</div>'
                    f'<strong>{int(qtd)}</strong>'
                    f'</div>'
                )
            if not barras:
                barras.append(
                    "<div class='dashboard-mini-sub'>Modelo da loja ainda não disponível na visão filtrada.</div>"
                )

            _render_html_dashboard(
                f"""
                <div class="dashboard-panel">
                    <div class="dashboard-panel-title">Treinamentos por Modelo</div>
                    {''.join(barras)}
                </div>
                """,
            )

        with row_a[2]:
            atividades_html = []
            df_ativ = df_dash.copy()
            if "Data_do_Contato" in df_ativ.columns:
                df_ativ["_dt_dash"] = pd.to_datetime(
                    df_ativ["Data_do_Contato"],
                    errors="coerce",
                    dayfirst=True,
                )
                df_ativ = df_ativ.sort_values(
                    "_dt_dash", ascending=False
                )
            for _, row in df_ativ.head(4).iterrows():
                status = str(row.get("Status_Contato", "") or "Atualização")
                pv = str(row.get("PV Abadi", "") or "")
                razao = html.escape(str(row.get("Razão Social", "") or "Cliente"))
                atividades_html.append(
                    f'<div class="dashboard-mini-row">'
                    f'<div class="dashboard-mini-badge">📌</div>'
                    f'<div style="flex:1">'
                    f'<div class="dashboard-mini-main">{html.escape(status)}</div>'
                    f'<div class="dashboard-mini-sub">{razao} · PV {html.escape(pv)}</div>'
                    f'</div>'
                    f'</div>'
                )
            _render_html_dashboard(
                f"""
                <div class="dashboard-panel">
                    <div class="dashboard-panel-title">Atividades Recentes</div>
                    {''.join(atividades_html) if atividades_html else '<div class="dashboard-mini-sub">Sem atividades recentes.</div>'}
                    <div class="dashboard-panel-link">Acompanhamento central do CRM →</div>
                </div>
                """,
            )

        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

        # --------------------------------------------------------
        # PIPELINE RESUMIDO
        # --------------------------------------------------------
        ordem_pipeline_dash = [
            ("A Contatar", "ps-blue"),
            ("Aguardando Pagamento", "ps-yellow"),
            ("Em Negociação", "ps-orange"),
            ("Agendado", "ps-orange"),
            ("Treinamento Realizado", "ps-green"),
            ("Recusado", "ps-gray"),
        ]
        pipeline_cards = []
        for nome_status, classe in ordem_pipeline_dash:
            pipeline_cards.append(
                f'<div class="pipeline-stat {classe}">'
                f'{html.escape(nome_status)}'
                f'<strong>{status_counts.get(nome_status, 0)}</strong>'
                f'</div>'
            )

        _render_html_dashboard(
            f"""
            <div class="dashboard-panel">
                <div class="dashboard-panel-title">Resumo do Pipeline</div>
                <div class="dashboard-pipeline-strip">
                    {''.join(pipeline_cards)}
                </div>
            </div>
            """,
        )

        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

        # --------------------------------------------------------
        # ESTADOS + INSTRUTORES
        # --------------------------------------------------------
        row_b = st.columns([1.15, 1])

        with row_b[0]:
            if "UF" in df_dash.columns:
                estados = []
                for uf, grupo in df_dash.groupby("UF", dropna=True):
                    total_uf = len(grupo)
                    ag_uf = (
                        int((grupo["Status_Contato"].astype(str) == "Agendado").sum())
                        if "Status_Contato" in grupo.columns else 0
                    )
                    conc_uf = (
                        int(
                            grupo["Status_Contato"].astype(str).isin(
                                ["Treinamento Realizado", "Concluído", "Concluido"]
                            ).sum()
                        )
                        if "Status_Contato" in grupo.columns else 0
                    )
                    taxa_uf = conc_uf / total_uf * 100 if total_uf else 0
                    estados.append((str(uf), total_uf, ag_uf, conc_uf, taxa_uf))
                estados = sorted(
                    estados,
                    key=lambda x: x[1],
                    reverse=True
                )[:7]
            else:
                estados = []

            linhas_estado = "".join(
                f"<tr><td><b>{html.escape(uf)}</b></td><td>{total}</td><td>{ag}</td><td>{conc}</td><td>{taxa:.1f}%</td></tr>"
                for uf, total, ag, conc, taxa in estados
            )
            _render_html_dashboard(
                f"""
                <div class="dashboard-panel">
                    <div class="dashboard-panel-title">Treinamentos por Estado</div>
                    <table class="dashboard-html-table">
                        <thead>
                            <tr><th>Estado</th><th>Postos</th><th>Agendados</th><th>Concluídos</th><th>Taxa</th></tr>
                        </thead>
                        <tbody>{linhas_estado}</tbody>
                    </table>
                </div>
                """,
            )

        with row_b[1]:
            top_instrutores = pd.DataFrame()
            if (
                "Instrutor_Sugerido" in df_dash.columns
                and "Status_Contato" in df_dash.columns
            ):
                df_conc = df_dash[
                    df_dash["Status_Contato"].astype(str).isin(
                        ["Treinamento Realizado", "Concluído", "Concluido"]
                    )
                ].copy()
                top_instrutores = (
                    df_conc["Instrutor_Sugerido"]
                    .dropna()
                    .astype(str)
                    .loc[lambda s: ~s.str.contains("Pendente", case=False, na=False)]
                    .value_counts()
                    .head(7)
                    .rename_axis("Instrutor")
                    .reset_index(name="Concluídos")
                )

            linhas_instr = "".join(
                f"<tr><td><b>{html.escape(str(r['Instrutor']))}</b></td><td>{int(r['Concluídos'])}</td><td>⭐</td></tr>"
                for _, r in top_instrutores.iterrows()
            )
            if not linhas_instr:
                linhas_instr = "<tr><td colspan='3'>Ainda não há treinamentos concluídos com instrutor identificado.</td></tr>"

            _render_html_dashboard(
                f"""
                <div class="dashboard-panel">
                    <div class="dashboard-panel-title">Top Instrutores (Concluídos)</div>
                    <table class="dashboard-html-table">
                        <thead>
                            <tr><th>Instrutor</th><th>Concluídos</th><th>Status</th></tr>
                        </thead>
                        <tbody>{linhas_instr}</tbody>
                    </table>
                </div>
                """,
            )

        # --------------------------------------------------------
        # AÇÕES RÁPIDAS VISUAIS
        # --------------------------------------------------------
        st.markdown("<div style='height:13px'></div>", unsafe_allow_html=True)
        _render_html_dashboard(
            """
            <div class="dashboard-panel-title">Ações Rápidas</div>
            <div class="quick-actions-grid">
                <div class="quick-action-card"><span>📞</span>Registrar<br>Novo Contato</div>
                <div class="quick-action-card"><span>📅</span>Agendar<br>Treinamento</div>
                <div class="quick-action-card"><span>📄</span>Nova Solicitação<br>de Orçamento</div>
                <div class="quick-action-card"><span>🔎</span>Consultar<br>PROCV</div>
                <div class="quick-action-card"><span>🧮</span>Calculadora<br>de Custos</div>
                <div class="quick-action-card"><span>📊</span>Relatórios<br>Gerenciais</div>
            </div>
            """,
        )

        render_exportacao_modulo(
            df_dash,
            "Dashboard_Executivo",
            nome_aba="Dashboard",
            legenda="Exporta a visão atual do Dashboard após os filtros aplicados."
        )

elif modulo == "📋 Pipeline AmPm":
    render_section_header(
        "📋",
        "Pipeline AmPm",
        "Fluxo operacional de treinamentos e liberação financeira"
    )

    # Ordem operacional solicitada: Recusado imediatamente ao lado de A Contatar.
    colunas_pipeline = [
        "A Contatar",
        "Recusado",
        "Aguardando Pagamento",
        "Em Negociação",
        "Agendado",
        "Treinamento Realizado",
    ]
    cols_k = st.columns(len(colunas_pipeline))

    df_pipeline_export = df_base.copy()
    if "Status_Contato" in df_pipeline_export.columns:
        df_pipeline_export = df_pipeline_export[
            df_pipeline_export["Status_Contato"].isin(colunas_pipeline)
        ].copy()

    for idx, status in enumerate(colunas_pipeline):
        df_status = (
            df_base[df_base["Status_Contato"] == status]
            if "Status_Contato" in df_base.columns
            else pd.DataFrame()
        )
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

            for _, item in df_status.head(10).iterrows():
                pv = item.get("PV Abadi")
                if pd.isna(pv):
                    pv = item.get("PV_Abadi", "-")

                razao = str(item.get("Razão Social", "") or item.get("Razao_Social", ""))
                with st.expander(
                    f"📍 PV {pv} | {razao[:18]}{'...' if len(razao) > 18 else ''}"
                ):
                    st.write(
                        f"**Cidade:** {item.get('Municipio', '-')}/"
                        f"{item.get('UF', '-')}"
                    )
                    st.write(
                        f"**Necessidade:** {item.get('Tipo_Necessidade', '-')}"
                    )

                    # -------------------------------
                    # CONTROLE FINANCEIRO
                    # -------------------------------
                    if status == "Aguardando Pagamento":
                        st.markdown("#### 💳 Liberação do treinamento")

                        tipo_atual = str(
                            item.get("Tipo_Pagamento", "") or ""
                        ).strip()

                        opcoes_pagamento = [
                            "Primeira parcela",
                            "Pagamento integral",
                        ]
                        tipo_idx = (
                            opcoes_pagamento.index(tipo_atual)
                            if tipo_atual in opcoes_pagamento else 0
                        )

                        tipo_pagamento = st.selectbox(
                            "Pagamento considerado para a contagem:",
                            opcoes_pagamento,
                            index=tipo_idx,
                            key=f"tipo_pag_{pv}",
                        )

                        data_pagamento_atual = parse_data_flexivel(
                            item.get("Data_Pagamento")
                        )
                        data_pagamento = st.date_input(
                            "Data do pagamento:",
                            value=data_pagamento_atual or date.today(),
                            key=f"data_pag_{pv}",
                        )

                        data_liberacao = calcular_liberacao_treinamento(
                            data_pagamento
                        )

                        st.info(
                            f"🗓️ **Treinamento liberado a partir de "
                            f"{data_liberacao.strftime('%d/%m/%Y')}** "
                            f"(7 dias após o pagamento)."
                        )

                        dias_restantes = (data_liberacao - date.today()).days
                        if dias_restantes > 0:
                            st.warning(
                                f"⏳ Ainda faltam **{dias_restantes} dia(s)** "
                                f"para liberar o treinamento."
                            )
                        else:
                            st.success(
                                "✅ Prazo financeiro cumprido. "
                                "O treinamento está liberado para agendamento."
                            )

                        if st.button(
                            "💾 Salvar pagamento",
                            key=f"salvar_pag_{pv}",
                            use_container_width=True,
                        ):
                            atualizar_fila(
                                pv,
                                {
                                    "Tipo_Pagamento": tipo_pagamento,
                                    "Data_Pagamento": data_pagamento,
                                    "Data_Liberacao_Treinamento": data_liberacao,
                                },
                            )
                            st.success("Pagamento e data de liberação salvos.")
                            st.rerun()

                    # -------------------------------
                    # ALTERAÇÃO DE STATUS
                    # -------------------------------
                    indice_status = colunas_pipeline.index(status)
                    mudar_status = st.selectbox(
                        "Alterar Status:",
                        colunas_pipeline,
                        index=indice_status,
                        key=f"pipe_sel_{pv}",
                    )

                    if mudar_status != status:
                        # Não permite liberar/agendar treinamento antes
                        # de completar os 7 dias após o pagamento.
                        if mudar_status == "Agendado":
                            data_pagamento = item.get("Data_Pagamento")
                            data_liberacao = calcular_liberacao_treinamento(
                                data_pagamento
                            )

                            if not data_liberacao:
                                st.error(
                                    "🚫 Não é possível agendar ainda. "
                                    "Informe a data da primeira parcela ou "
                                    "do pagamento integral na coluna "
                                    "**Aguardando Pagamento**."
                                )
                            elif date.today() < data_liberacao:
                                faltam = (data_liberacao - date.today()).days
                                st.error(
                                    f"🚫 Agendamento bloqueado. "
                                    f"O treinamento só poderá ser liberado "
                                    f"a partir de **"
                                    f"{data_liberacao.strftime('%d/%m/%Y')}** "
                                    f"({faltam} dia(s) restantes)."
                                )
                            else:
                                atualizar_fila(
                                    pv,
                                    {
                                        "Status_Contato": mudar_status,
                                        "Data_Liberacao_Treinamento": data_liberacao,
                                    },
                                )
                                st.rerun()
                        else:
                            atualizar_fila(
                                pv,
                                {"Status_Contato": mudar_status}
                            )
                            st.rerun()
    render_exportacao_modulo(
        df_pipeline_export,
        "Pipeline_AmPm",
        nome_aba="Pipeline",
        legenda="Exporta os clientes e seus estágios atuais no Pipeline."
    )

elif modulo == "🔍 PROCV Gestão e Franquia AMPM":
    render_section_header("🔍", "PROCV Gestão e Franquia AMPM", "Consulta detalhada da loja, gestão, franquia e histórico de treinamento")
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
        cols_mostrar = [c for c in ['PV Abadi', 'Razão Social', 'CNPJ', 'Telefone_Contato', 'Email_Contato', 'Nome_Contato', 'Municipio', 'UF', 'Status Loja', 'Tipo_Necessidade', 'Instrutor_Treinamento', 'Instrutor_Inauguracao', 'Instrutor_Sugerido', 'Status_Contato'] if c in df_view.columns]
        evento = st.dataframe(df_view[cols_mostrar], use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun")
        linhas = evento.selection.get("rows", [])
        if linhas:
            p = df_view.iloc[linhas[0]].to_dict()
            instrutor_treinamento, instrutor_inauguracao = _historico_instrutor_procv(p)
            st.divider()
            st.markdown(f"**📋 Gestão e Franquia AMPM — PV {p.get('PV Abadi', '-')} · {p.get('Razão Social', '-')}**")
            k1, k2, k3 = st.columns(3)
            cnpj_procv = _procv_valor_flexivel(
                p, "CNPJ", ["CNPJ Completo", "CNPJ Loja", "CNPJ da Loja", "Documento CNPJ", "CNPJ Posto"]
            )
            telefone_procv = _procv_valor_flexivel(
                p, "Telefone_Contato", ["Telefone", "Telefone da Loja", "Telefone Loja", "Celular", "WhatsApp", "Fone", "Telefone Comercial"]
            )
            email_procv = _procv_valor_flexivel(
                p, "Email_Contato", ["Email", "E-mail", "E Mail", "Email da Loja", "Email Loja", "Correio Eletrônico"]
            )
            nome_contato_procv = _procv_valor_flexivel(
                p, "Nome_Contato", ["Nome do Contato", "Responsável", "Responsavel", "Responsável Loja", "Contato"]
            )
            qtd_func_procv = _procv_valor_flexivel(
                p, "Qtd_Funcionarios", ["Qtd Funcionários", "Quantidade de Funcionários", "Funcionários"]
            )

            with k1:
                st.markdown(f"""<div class="procv-card"><h4>🏪 Cadastro da Loja</h4><p>🧾 <b>CNPJ:</b> {cnpj_procv}</p><p>📍 <b>Endereço:</b> {p.get('Endereço', '-')}</p><p>🏙️ <b>Cidade/UF:</b> {p.get('Municipio', '-')}/{p.get('UF', '-')}</p><p>⚙️ <b>Status da loja:</b> {p.get('Status Loja', '-')}</p></div>""", unsafe_allow_html=True)
            with k2:
                st.markdown(f"""<div class="procv-card"><h4>📞 Contatos da Loja</h4><p>👤 <b>Responsável:</b> {nome_contato_procv}</p><p>📞 <b>Telefone:</b> {telefone_procv}</p><p>✉️ <b>E-mail:</b> {email_procv}</p><p>👥 <b>Funcionários:</b> {qtd_func_procv}</p></div>""", unsafe_allow_html=True)
            with k3:
                st.markdown(f"""<div class="procv-card"><h4>👔 Gestão & Histórico</h4><p>👤 <b>Gerente (GF):</b> {p.get('GF', 'Não informado')}</p><p>👔 <b>Consultor (CF):</b> {p.get('CF', 'Não informado')}</p><p>📅 <b>Inauguração:</b> {p.get('Previsão Inauguração', 'Não informado')}</p><hr><p>🎓 <b>Instrutor do treinamento:</b> {instrutor_treinamento}</p><p>🚀 <b>Instrutor da inauguração:</b> {instrutor_inauguracao}</p><p>📅 <b>Último treinamento:</b> {p.get('Data_Ultimo_Treinamento', 'Não informado')}</p><p>🎯 <b>Necessidade atual:</b> {p.get('Tipo_Necessidade', '-')}</p><p>🔄 <b>Status do atendimento:</b> {badge_status_html(p.get('Status_Contato', '-'))}</p></div>""", unsafe_allow_html=True)
        render_exportacao_modulo(
            df_view,
            "PROCV_Gestao_Franquia_AMPM",
            nome_aba="PROCV",
            legenda="Exporta o resultado atual da pesquisa e dos filtros aplicados no PROCV."
        )

    else:
        st.info("📭 Nenhum dado carregado ainda.")

elif modulo == "📍 Calculadora & Otimizador de Custos":
    render_section_header(
        "📍",
        "Calculadora & Otimizador de Custos",
        "Top 3 instrutores ativos da equipe oficial por proximidade e simulação de custos"
    )

    if df_lojas is None or df_lojas.empty:
        st.info("📭 Nenhuma loja disponível na base central.")
    elif df_instrutores is None or df_instrutores.empty:
        st.info("📭 Nenhum instrutor disponível na base central.")
    else:
        lojas_calc = df_lojas.copy()

        if filtro_uf != "Todas" and "UF" in lojas_calc.columns:
            lojas_calc = lojas_calc[lojas_calc["UF"] == filtro_uf].copy()

        cols_postos = [
            c for c in ["PV Abadi", "Razão Social", "Municipio", "UF"]
            if c in lojas_calc.columns
        ]
        postos_unicos = lojas_calc[cols_postos].drop_duplicates(subset=["PV Abadi"])

        if postos_unicos.empty:
            st.info("📭 Nenhum posto disponível para cálculo.")
        else:
            postos_unicos["label"] = (
                postos_unicos["PV Abadi"].astype(str)
                + " - "
                + postos_unicos.get("Razão Social", pd.Series("", index=postos_unicos.index)).fillna("").astype(str)
                + " ("
                + postos_unicos.get("Municipio", pd.Series("", index=postos_unicos.index)).fillna("").astype(str)
                + "/"
                + postos_unicos.get("UF", pd.Series("", index=postos_unicos.index)).fillna("").astype(str)
                + ")"
            )

            posto_sel = st.selectbox(
                "⛽ Selecione o Posto Alvo:",
                postos_unicos["label"].tolist()
            )
            pv_txt = posto_sel.split(" - ")[0].strip()

            posto_match = lojas_calc[
                lojas_calc["PV Abadi"].astype(str).str.replace(".0", "", regex=False) == pv_txt
            ]
            if posto_match.empty:
                st.warning("Não foi possível localizar o PV selecionado.")
            else:
                posto = posto_match.iloc[0].to_dict()

                with st.spinner("Calculando os instrutores ativos mais próximos..."):
                    top_3 = _top_instrutores_proximos(
                        posto,
                        df_instrutores,
                        limite=3
                    )

                if top_3.empty:
                    st.warning(
                        "Não foi possível calcular a proximidade deste posto. "
                        "Verifique Município/UF da loja e a localização dos instrutores."
                    )
                else:
                    st.success(
                        f"✅ {len(top_3)} instrutor(es) ativo(s) mais próximo(s) calculado(s) em tempo real."
                    )

                    # =====================================================
                    # MAPA — UMA ROTA POR VEZ, RESTANTE DA V37 PRESERVADO
                    # =====================================================
                    campos_geo = ["Lat_Loja", "Lon_Loja", "Lat_Instrutor", "Lon_Instrutor"]
                    top_geo = top_3[
                        top_3.apply(
                            lambda r: all(_valor_preenchido(r.get(c)) for c in campos_geo),
                            axis=1
                        )
                    ].copy()

                    if top_geo.empty:
                        st.warning("Não há coordenadas suficientes para montar o mapa de rotas.")
                    else:
                        opcoes_mapa = []
                        linhas_mapa = {}

                        for idx, (_, row_geo) in enumerate(top_geo.iterrows()):
                            nome_geo = str(row_geo.get("Instrutor_Sugerido", f"Instrutor {idx + 1}"))
                            dist_geo = float(row_geo.get("Distancia_km_linha_reta", 0) or 0)
                            label_geo = f"#{idx + 1} {nome_geo} · {dist_geo:.1f} km"
                            opcoes_mapa.append(label_geo)
                            linhas_mapa[label_geo] = row_geo

                        st.markdown("### 🗺️ Rotas dos instrutores")
                        instrutor_mapa_sel = st.radio(
                            "Selecione qual rota deseja visualizar:",
                            opcoes_mapa,
                            horizontal=True,
                            key=f"v39_rota_instrutor_{pv_txt}"
                        )

                        row_mapa = linhas_mapa[instrutor_mapa_sel]

                        p_lat = float(row_mapa["Lat_Loja"])
                        p_lon = float(row_mapa["Lon_Loja"])
                        i_lat = float(row_mapa["Lat_Instrutor"])
                        i_lon = float(row_mapa["Lon_Instrutor"])
                        nome_mapa = str(row_mapa.get("Instrutor_Sugerido", "Instrutor"))

                        with st.spinner("Montando rota selecionada..."):
                            rota_coords, rota_km, rota_min = _obter_rota_rodoviaria(
                                i_lat, i_lon, p_lat, p_lon
                            )

                        if rota_coords:
                            path = rota_coords
                            tipo_rota = "Rodoviária"
                            distancia_rota = rota_km
                            duracao_rota = rota_min
                        else:
                            path = [[i_lon, i_lat], [p_lon, p_lat]]
                            tipo_rota = "Linha reta (fallback)"
                            distancia_rota = float(
                                row_mapa.get("Distancia_km_linha_reta", 0) or 0
                            )
                            duracao_rota = None

                        # Cor da rota acompanha o instrutor selecionado.
                        ranking_mapa = int(row_mapa.get("Ranking_Proximidade", 1) or 1)
                        cores_rota = {
                            1: [255, 43, 32, 245],    # vermelho AmPm
                            2: [255, 196, 0, 245],    # amarelo
                            3: [28, 35, 43, 245],     # carvão
                        }
                        cor_rota = cores_rota.get(ranking_mapa, [255, 43, 32, 245])

                        df_rota_mapa = pd.DataFrame([{
                            "Instrutor": nome_mapa,
                            "Tipo": tipo_rota,
                            "path": path,
                            "cor": [cor_rota],
                        }])

                        # Marcadores discretos ao longo da rota, como no layout aprovado.
                        pontos_amostrados = _amostrar_pontos_rota(path, quantidade=14)
                        registros_pontos_rota = []
                        for lon_pt, lat_pt in pontos_amostrados:
                            registros_pontos_rota.append({
                                "name": nome_mapa,
                                "tipo": "Rota",
                                "lat": float(lat_pt),
                                "lon": float(lon_pt),
                                "cor": cor_rota,
                            })
                        df_pontos_rota = pd.DataFrame(registros_pontos_rota)

                        df_extremos_mapa = pd.DataFrame([
                            {
                                "name": nome_mapa,
                                "tipo": "Origem do instrutor",
                                "lat": i_lat,
                                "lon": i_lon,
                                "cor": cor_rota,
                            },
                            {
                                "name": f"Posto {row_mapa.get('PV_ABADI', '')}",
                                "tipo": "Destino",
                                "lat": p_lat,
                                "lon": p_lon,
                                "cor": [255, 138, 0, 255],
                            },
                        ])

                        # Halo branco fino para separar a rota do mapa sem criar borrão.
                        layer_rota_halo = pdk.Layer(
                            "PathLayer",
                            df_rota_mapa,
                            get_path="path",
                            get_color=[255, 255, 255, 235],
                            get_width=6,
                            width_units="pixels",
                            width_min_pixels=5,
                            width_max_pixels=7,
                            joint_rounded=True,
                            cap_rounded=True,
                            pickable=False,
                        )

                        # Linha principal: fina e nítida.
                        layer_rota = pdk.Layer(
                            "PathLayer",
                            df_rota_mapa,
                            get_path="path",
                            get_color="cor",
                            get_width=2.8,
                            width_units="pixels",
                            width_min_pixels=2,
                            width_max_pixels=4,
                            joint_rounded=True,
                            cap_rounded=True,
                            pickable=True,
                        )

                        # Pequenos pontos de referência na rota.
                        layer_pontos_rota = pdk.Layer(
                            "ScatterplotLayer",
                            df_pontos_rota,
                            get_position="[lon, lat]",
                            get_fill_color="cor",
                            get_line_color=[255, 255, 255, 255],
                            stroked=True,
                            filled=True,
                            radius_units="pixels",
                            get_radius=4.2,
                            radius_min_pixels=3,
                            radius_max_pixels=5,
                            line_width_units="pixels",
                            get_line_width=1.5,
                            pickable=False,
                        )

                        # Origem e destino um pouco maiores.
                        layer_extremos = pdk.Layer(
                            "ScatterplotLayer",
                            df_extremos_mapa,
                            get_position="[lon, lat]",
                            get_fill_color="cor",
                            get_line_color=[255, 255, 255, 255],
                            stroked=True,
                            filled=True,
                            radius_units="pixels",
                            get_radius=7,
                            radius_min_pixels=6,
                            radius_max_pixels=9,
                            line_width_units="pixels",
                            get_line_width=2,
                            pickable=True,
                        )

                        coords_enquadramento = [(p_lat, p_lon), (i_lat, i_lon)]
                        centro_lat = (p_lat + i_lat) / 2
                        centro_lon = (p_lon + i_lon) / 2
                        zoom_mapa = _zoom_para_pontos(coords_enquadramento)

                        st.pydeck_chart(
                            pdk.Deck(
                                map_style="https://basemaps.cartocdn.com/gl/positron-gl-style/style.json",
                                layers=[
                                    layer_rota_halo,
                                    layer_rota,
                                    layer_pontos_rota,
                                    layer_extremos,
                                ],
                                initial_view_state=pdk.ViewState(
                                    latitude=centro_lat,
                                    longitude=centro_lon,
                                    zoom=zoom_mapa,
                                    pitch=0,
                                    bearing=0,
                                ),
                                tooltip={"text": "{name}\n{tipo}"},
                            ),
                            use_container_width=True,
                        )

                        if duracao_rota is not None:
                            st.caption(
                                f"🛣️ {nome_mapa}: {distancia_rota:.1f} km · "
                                f"{duracao_rota/60:.1f} h · {tipo_rota}"
                            )
                        else:
                            st.caption(
                                f"🧭 {nome_mapa}: {distancia_rota:.1f} km · {tipo_rota}"
                            )

                    st.markdown("### 💰 Composição estimada de custos")
                    st.caption(
                        "Os três candidatos abaixo são ordenados pela distância em linha reta. "
                        "Agenda e disponibilidade continuam podendo alterar a escolha final."
                    )

                    resultados_custos = []
                    cols = st.columns(len(top_3))

                    for idx, (_, row) in enumerate(top_3.iterrows()):
                        nome_instrutor = str(row.get("Instrutor_Sugerido", "Instrutor"))
                        dist = float(row.get("Distancia_km_linha_reta", 0) or 0)

                        with cols[idx]:
                            st.markdown(
                                f"""<div class="top-instructor-card">
                                <h4>#{idx + 1} {nome_instrutor}</h4>
                                <p>Origem: {row.get('Cidade_Instrutor', '-')} / {row.get('UF_Instrutor', '-')}</p>
                                <p>Distância estimada: {dist:.1f} km</p>
                                </div>""",
                                unsafe_allow_html=True
                            )

                            with st.expander("⚙️ Configurar custos", expanded=True):
                                dias = st.number_input(
                                    "📅 Dias de treinamento",
                                    min_value=0.5,
                                    value=1.0,
                                    step=0.5,
                                    key=f"v35_dias_{pv_txt}_{idx}"
                                )

                                usar_diaria = st.checkbox("Diárias", True, key=f"v35_diaria_{pv_txt}_{idx}")
                                valor_diaria = st.number_input(
                                    "Diária (R$/dia)", min_value=0.0, value=280.0, step=10.0,
                                    key=f"v35_vdiaria_{pv_txt}_{idx}"
                                )

                                usar_hosp = st.checkbox("Hospedagem", True, key=f"v35_hosp_{pv_txt}_{idx}")
                                valor_hosp = st.number_input(
                                    "Hospedagem (R$/dia)", min_value=0.0, value=250.0, step=10.0,
                                    key=f"v35_vhosp_{pv_txt}_{idx}"
                                )

                                usar_carro = st.checkbox("Aluguel de carro", False, key=f"v35_carro_{pv_txt}_{idx}")
                                valor_carro = st.number_input(
                                    "Carro (R$/dia)", min_value=0.0, value=180.0, step=10.0,
                                    key=f"v35_vcarro_{pv_txt}_{idx}"
                                )

                                usar_rod = st.checkbox("Deslocamento rodoviário", False, key=f"v35_rod_{pv_txt}_{idx}")
                                valor_rod = st.number_input(
                                    "Rodoviário (R$/viagem)",
                                    min_value=0.0,
                                    value=float(max(0.0, dist * 2 * 2.10)),
                                    step=10.0,
                                    key=f"v35_vrod_{pv_txt}_{idx}"
                                )

                                usar_aviao = st.checkbox("Deslocamento de avião", False, key=f"v35_aviao_{pv_txt}_{idx}")
                                valor_aviao = st.number_input(
                                    "Avião (R$/viagem)", min_value=0.0, value=800.0, step=50.0,
                                    key=f"v35_vaviao_{pv_txt}_{idx}"
                                )

                                usar_treino = st.checkbox("Valor do treinamento", True, key=f"v35_treino_{pv_txt}_{idx}")
                                valor_treino = st.number_input(
                                    "Treinamento (R$/dia)", min_value=0.0, value=280.0, step=10.0,
                                    key=f"v35_vtreino_{pv_txt}_{idx}"
                                )

                                total = (
                                    (valor_diaria * dias if usar_diaria else 0)
                                    + (valor_hosp * dias if usar_hosp else 0)
                                    + (valor_carro * dias if usar_carro else 0)
                                    + (valor_rod if usar_rod else 0)
                                    + (valor_aviao if usar_aviao else 0)
                                    + (valor_treino * dias if usar_treino else 0)
                                )

                                st.metric("💰 Total estimado", f"R$ {total:,.2f}")

                                resultados_custos.append({
                                    "Ranking": idx + 1,
                                    "Instrutor": nome_instrutor,
                                    "Origem": f"{row.get('Cidade_Instrutor', '-')}/{row.get('UF_Instrutor', '-')}",
                                    "Distância (km)": dist,
                                    "Total Estimado": total,
                                })

                    if resultados_custos:
                        st.markdown("### 📊 Comparativo")
                        df_comparativo = pd.DataFrame(resultados_custos).sort_values(
                            "Total Estimado"
                        )
                        st.dataframe(
                            df_comparativo,
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "Distância (km)": st.column_config.NumberColumn(format="%.1f km"),
                                "Total Estimado": st.column_config.NumberColumn(format="R$ %.2f"),
                            }
                        )
                        df_export_calc = df_comparativo.copy()
                        df_export_calc.insert(0, "PV", pv_txt)
                        df_export_calc.insert(1, "Posto", str(posto.get("Razão Social", "") or ""))
                        df_export_calc.insert(
                            2,
                            "Destino",
                            f"{posto.get('Municipio', '')}/{posto.get('UF', '')}"
                        )
                        render_exportacao_modulo(
                            df_export_calc,
                            f"Calculadora_Custos_PV_{pv_txt}",
                            nome_aba="Comparativo Custos",
                            legenda="Exporta o comparativo calculado para o posto selecionado."
                        )

elif modulo == "📥 Importador Inteligente":
    render_importador_inteligente()

elif modulo == "📞 Call Center & Timeline WhatsApp":
    def _texto_seguro_callcenter(valor):
        if valor is None:
            return ""
        try:
            resultado = pd.isna(valor)
            if isinstance(resultado, bool) and resultado:
                return ""
            if not isinstance(resultado, bool) and bool(resultado.all()):
                return ""
        except (TypeError, ValueError):
            pass
        return str(valor)
    render_section_header("📞", "Call Center & Timeline WhatsApp", "Registro de atendimentos e disparo de mensagens")
    if not df_base.empty:
        # O Call Center parte de TODOS os clientes; filtros apenas reduzem a visão.
        df_fila_view = df_base.copy()

        st.markdown("#### 🎛️ Filtros combináveis do Call Center")
        busca_call = st.text_input(
            "🔍 Buscar PV, cliente, município, telefone, e-mail ou contato:",
            key="call_busca_v36"
        )

        f1, f2, f3 = st.columns(3)
        with f1:
            ufs_opts = sorted(df_fila_view["UF"].dropna().astype(str).unique().tolist()) if "UF" in df_fila_view.columns else []
            ufs_call = st.multiselect("🗺️ UF", ufs_opts, key="call_uf_v36")
        with f2:
            nec_opts = sorted(df_fila_view["Tipo_Necessidade"].dropna().astype(str).unique().tolist()) if "Tipo_Necessidade" in df_fila_view.columns else []
            necessidades_call = st.multiselect(
                "🎯 Situação / necessidade",
                nec_opts,
                help="Marque uma ou várias opções: a contratar, retreinamento, já treinados etc.",
                key="call_nec_v36"
            )
        with f3:
            status_opts = sorted(df_fila_view["Status_Contato"].dropna().astype(str).unique().tolist()) if "Status_Contato" in df_fila_view.columns else []
            status_call = st.multiselect("📞 Status do contato", status_opts, key="call_status_v36")

        if busca_call:
            campos = [c for c in ["PV Abadi","Razão Social","Municipio","Telefone_Contato","Email_Contato","Nome_Contato"] if c in df_fila_view.columns]
            mascara = pd.Series(False, index=df_fila_view.index)
            for c in campos:
                mascara |= df_fila_view[c].astype(str).str.contains(busca_call, case=False, na=False, regex=False)
            df_fila_view = df_fila_view[mascara]
        if ufs_call and "UF" in df_fila_view.columns:
            df_fila_view = df_fila_view[df_fila_view["UF"].astype(str).isin(ufs_call)]
        if necessidades_call and "Tipo_Necessidade" in df_fila_view.columns:
            df_fila_view = df_fila_view[df_fila_view["Tipo_Necessidade"].astype(str).isin(necessidades_call)]
        if status_call and "Status_Contato" in df_fila_view.columns:
            df_fila_view = df_fila_view[df_fila_view["Status_Contato"].astype(str).isin(status_call)]

        st.caption(f"👥 Exibindo {len(df_fila_view):,} de {len(df_base):,} clientes.")

        render_exportacao_modulo(
            df_fila_view,
            "Call_Center",
            nome_aba="Call Center",
            legenda="Exporta os clientes atualmente exibidos após busca e filtros do Call Center."
        )

        c_left, c_right = st.columns([1.2, 1.8])

        with c_left:
            st.markdown("**📋 Todos os Clientes / Fila de Atendimento**")
            cols_call = [c for c in ['PV Abadi','Razão Social','Municipio','UF','Tipo_Necessidade','Status_Contato'] if c in df_fila_view.columns]
            tabela_call = df_fila_view[cols_call].copy().reset_index(drop=True)
            tabela_call.insert(0, "Selecionar", False)
            tabela_call_editada = st.data_editor(
                tabela_call,
                use_container_width=True,
                hide_index=True,
                disabled=[c for c in tabela_call.columns if c != "Selecionar"],
                column_config={
                    "Selecionar": st.column_config.CheckboxColumn("☑️", help="Marque o cliente")
                },
                key="call_tabela_v36"
            )
            marcados = tabela_call_editada.index[tabela_call_editada["Selecionar"] == True].tolist()
            selecionado = marcados[:1]
            if len(marcados) > 1:
                st.caption(f"☑️ {len(marcados)} clientes marcados. A ficha à direita mostra o primeiro.")

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
                                <p>🏷️ <b>Tipo de Modelo:</b> {_procv_valor_flexivel(
                                    posto,
                                    "Tipo de Modelo",
                                    ["Tipo Modelo", "Modelo", "Modelo da Loja", "Modelo Loja"]
                                )}</p>
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

                # --- WHATSAPP: MODELOS CENTRAIS EDITÁVEIS + MENSAGEM PRÉ-PREENCHIDA ---
                # Chave segura do PV precisa existir antes dos widgets do WhatsApp.
                chave_pv_widget = re.sub(
                    r"[^A-Za-z0-9_-]+",
                    "_",
                    str(pv_alvo if pv_alvo is not None else "sem_pv")
                ).strip("_") or "sem_pv"

                if tel_limpo:
                    st.markdown("##### 📲 WhatsApp")

                    modelos_wa = carregar_modelos_whatsapp()
                    nomes_modelos_wa = {
                        dados.get("nome", chave): chave
                        for chave, dados in modelos_wa.items()
                        if dados.get("ativo", True)
                    }

                    if nomes_modelos_wa:
                        modelo_nome = st.selectbox(
                            "Modelo de mensagem:",
                            list(nomes_modelos_wa.keys()),
                            key=f"wa_modelo_{chave_pv_widget}",
                        )
                        modelo_chave = nomes_modelos_wa[modelo_nome]
                        modelo_dados = modelos_wa[modelo_chave]

                        mensagem_gerada = _texto_modelo_whatsapp(
                            modelo_dados.get("mensagem", ""),
                            posto,
                        )

                        mensagem_envio = st.text_area(
                            "Mensagem para este cliente:",
                            value=mensagem_gerada,
                            height=150,
                            key=f"wa_msg_cliente_{chave_pv_widget}_{modelo_chave}",
                            help=(
                                "Você pode ajustar esta mensagem somente para este envio. "
                                "Para alterar o padrão, use 'Editar modelos padrão'."
                            ),
                        )

                        from urllib.parse import quote
                        link_wa = (
                            f"https://wa.me/55{tel_limpo}"
                            f"?text={quote(str(mensagem_envio))}"
                        )

                        st.link_button(
                            "💬 Abrir conversa no WhatsApp",
                            link_wa,
                            use_container_width=True,
                        )

                        with st.expander("✏️ Editar modelos padrão", expanded=False):
                            _render_editor_modelos_whatsapp()
                    else:
                        st.info("Nenhum modelo de WhatsApp ativo foi encontrado.")
                else:
                    st.caption(
                        "📵 Este cliente não possui telefone cadastrado para abrir o WhatsApp."
                    )

                lista_instrutores = ["Pendente de Alocação"]
                df_instrutores_ativos_call = filtrar_instrutores_ativos(df_instrutores)
                if (
                    df_instrutores_ativos_call is not None
                    and not df_instrutores_ativos_call.empty
                    and 'NOME_COMPLETO' in df_instrutores_ativos_call.columns
                ):
                    lista_instrutores += sorted(
                        df_instrutores_ativos_call['NOME_COMPLETO']
                        .dropna()
                        .astype(str)
                        .unique()
                        .tolist()
                    )

                instrutor_atual = str(posto.get('Instrutor_Sugerido', 'Pendente de Alocação'))
                idx_instrutor = lista_instrutores.index(instrutor_atual) if instrutor_atual in lista_instrutores else 0

                data_inicial = parse_data_flexivel(posto.get('Data_Agendada')) or date.today()

                # --- REGISTROS RÁPIDOS DA LIGAÇÃO ---
                st.markdown("#### ✍️ Registros Rápidos da Ligação")

                # Estes dois controles ficam FORA do st.form porque widgets dentro
                # de formulários só atualizam após Submit. Assim "Sim" habilita
                # imediatamente a quantidade de funcionários.
                tem_func_atual = _texto_seguro_callcenter(
                    posto.get('Tem_Funcionarios', 'Sim')
                ) or 'Sim'
                tem_func_opcoes = ["Sim", "Não"]
                qtd_func_atual = _texto_seguro_callcenter(
                    posto.get('Qtd_Funcionarios', 0)
                )
                try:
                    qtd_func_padrao = int(float(qtd_func_atual or 0))
                except (TypeError, ValueError):
                    qtd_func_padrao = 0
                qtd_func_padrao = max(0, qtd_func_padrao)

                if qtd_func_padrao == 0 and tem_func_atual not in tem_func_opcoes:
                    tem_func_atual = "Não"
                elif tem_func_atual not in tem_func_opcoes:
                    tem_func_atual = "Sim"

                idx_tem_func = tem_func_opcoes.index(tem_func_atual)

                cf1, cf2 = st.columns(2)
                with cf1:
                    tem_funcionarios = st.selectbox(
                        "👥 Há funcionários para treinar?",
                        tem_func_opcoes,
                        index=idx_tem_func,
                        key=f"tem_func_call_{chave_pv_widget}",
                    )
                with cf2:
                    qtd_func = st.number_input(
                        "🔢 Qtd. de Funcionários para Treinar:",
                        value=(qtd_func_padrao if tem_funcionarios == "Sim" else 0),
                        min_value=0,
                        step=1,
                        disabled=(tem_funcionarios == "Não"),
                        key=f"qtd_func_call_{chave_pv_widget}",
                    )

                if tem_funcionarios == "Não":
                    qtd_func = 0

                with st.form("form_callcenter_editavel"):
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        nome_c = st.text_input("👤 Nome do Responsável na Loja:", value=_texto_seguro_callcenter(posto.get('Nome_Contato', '')))
                        tel_c = st.text_input("📞 Telefone de Contato:", value=_texto_seguro_callcenter(posto.get('Telefone_Contato', '')))
                        email_c = st.text_input("✉️ E-mail do Contato:", value=_texto_seguro_callcenter(posto.get('Email_Contato', '')))
                        instrutor_escolhido = st.selectbox(
                            "👨‍🏫 Instrutor Designado:",
                            lista_instrutores,
                            index=idx_instrutor
                        )
                        st.caption(
                            f"✅ {max(len(lista_instrutores)-1, 0)} instrutor(es) ativo(s) disponível(is)."
                        )

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

                    obs = st.text_area("💬 Observações e Alinhamentos:", value=_texto_seguro_callcenter(posto.get('Observacoes', '')), height=80)

                    if st.form_submit_button("💾 Salvar Registro do Atendimento"):
                        atualizar_fila(pv_alvo, {
                            'Nome_Contato': nome_c,
                            'Telefone_Contato': tel_c,
                            'Email_Contato': email_c,
                            'Tem_Funcionarios': 'Sim' if tem_funcionarios == 'Sim' and int(qtd_func) > 0 else 'Não',
                            'Qtd_Funcionarios': int(qtd_func),
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


                st.divider()
                aba_orc = st.tabs(["💰 Orçamento do Cliente"])
                with aba_orc[0]:
                    _renderizar_mini_orcamento(posto, pv_alvo)
    else:
        st.info("📭 Nenhum dado carregado ainda.")

elif modulo == "👔 Equipe de Instrutores":
    render_section_header(
        "👔",
        "Equipe de Instrutores",
        "Equipe oficial — status operacional controlado pelo administrador"
    )

    if df_instrutores is None or df_instrutores.empty:
        st.info("📭 Nenhum instrutor cadastrado no banco central.")
    else:
        df_instrutores_view = df_instrutores.copy()

        if "STATUS" not in df_instrutores_view.columns:
            df_instrutores_view["STATUS"] = "Ativo"

        # A Base_Unificada enviada possui 13 registros; somente STATUS=Ativo
        # participa de cálculo de proximidade, agenda e sugestões.
        instrutores_ativos = filtrar_instrutores_ativos(df_instrutores_view)
        total_base = len(df_instrutores_view)
        total_ativos = len(instrutores_ativos)
        total_inativos = max(total_base - total_ativos, 0)

        m1, m2, m3 = st.columns(3)
        m1.metric("✅ Em atividade", total_ativos)
        m2.metric("⏸️ Fora de atividade", total_inativos)
        m3.metric("📚 Total no histórico", total_base)

        colunas_seguras = [
            c for c in [
                "NOME_COMPLETO", "STATUS", "Cidade", "UF", "TELEFONE", "EMAIL"
            ]
            if c in df_instrutores_view.columns
        ]

        render_exportacao_modulo(
            df_instrutores_view[colunas_seguras].copy(),
            "Equipe_de_Instrutores",
            nome_aba="Instrutores",
            legenda="Exporta a equipe completa, incluindo ativos e histórico de quem saiu."
        )

        if usuario_e_admin():
            st.markdown("### ☑️ Controle de atividade")
            st.caption(
                "Marque somente quem está trabalhando atualmente. "
                "Instrutores desmarcados continuam no histórico, mas deixam de aparecer "
                "na Calculadora, sugestões de proximidade e novas alocações."
            )

            tabela_admin = df_instrutores_view[colunas_seguras].copy()
            tabela_admin.insert(
                0,
                "Em atividade",
                tabela_admin["STATUS"]
                .map(_texto_seguro_instrutor)
                .str.strip()
                .str.casefold()
                .eq("ativo")
            )

            tabela_editada = st.data_editor(
                tabela_admin,
                use_container_width=True,
                hide_index=True,
                disabled=[c for c in tabela_admin.columns if c != "Em atividade"],
                column_config={
                    "Em atividade": st.column_config.CheckboxColumn(
                        "✅ Em atividade",
                        help="Marcado = Ativo. Desmarcado = Saiu/inativo.",
                    ),
                    "NOME_COMPLETO": st.column_config.TextColumn("Instrutor"),
                    "STATUS": st.column_config.TextColumn("Status atual"),
                },
                key="editor_status_instrutores_v41",
            )

            b1, b2 = st.columns([1, 2])
            with b1:
                salvar_status = st.button(
                    "💾 Salvar status",
                    type="primary",
                    use_container_width=True,
                    key="salvar_status_instrutores_v41",
                )

            if salvar_status:
                try:
                    alteracoes = salvar_status_instrutores_admin(
                        tabela_editada,
                        tabela_admin,
                    )
                    if alteracoes:
                        # Recarrega a fonte oficial depois das alterações.
                        st.session_state["bases"] = carregar_bases_supabase()
                        st.success(
                            f"✅ {alteracoes} alteração(ões) de status salva(s) no Supabase."
                        )
                        st.rerun()
                    else:
                        st.info("Nenhuma alteração de status para salvar.")
                except PermissionError as exc:
                    st.error(f"🚫 {exc}")
                except Exception as exc:
                    st.error(f"❌ Falha ao salvar status: {exc}")

            st.divider()
            st.markdown("### 👥 Equipe ativa")
            if not instrutores_ativos.empty:
                st.dataframe(
                    instrutores_ativos[
                        [c for c in colunas_seguras if c in instrutores_ativos.columns]
                    ],
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info("📭 Nenhum instrutor marcado como Ativo.")

        else:
            st.markdown("### 👥 Instrutores em atividade")
            if not instrutores_ativos.empty:
                st.dataframe(
                    instrutores_ativos[
                        [c for c in colunas_seguras if c in instrutores_ativos.columns]
                    ],
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info("📭 Nenhum instrutor marcado como Ativo.")

        st.divider()

        if usuario_e_admin():
            st.markdown("### ➕ Cadastrar novo instrutor")
            st.caption(
                "Novos instrutores entram inicialmente como Ativo e podem ser "
                "desativados depois nas caixas acima."
            )

            with st.form("form_novo_instrutor", clear_on_submit=True):
                col1, col2 = st.columns(2)

                with col1:
                    novo_nome = st.text_input(
                        "Nome completo *",
                        placeholder="Ex.: João da Silva"
                    )
                    novo_telefone = st.text_input(
                        "Telefone",
                        placeholder="(00) 00000-0000"
                    )
                    novo_email = st.text_input(
                        "E-mail",
                        placeholder="instrutor@empresa.com"
                    )

                with col2:
                    nova_cidade = st.text_input(
                        "Cidade",
                        placeholder="Ex.: Rio de Janeiro"
                    )
                    nova_uf = st.text_input(
                        "UF",
                        max_chars=2,
                        placeholder="RJ"
                    )
                    st.text_input(
                        "Status inicial",
                        value="Ativo",
                        disabled=True
                    )

                cadastrar = st.form_submit_button(
                    "💾 Cadastrar Instrutor",
                    use_container_width=True
                )

            if cadastrar:
                try:
                    acao = adicionar_instrutor_admin(
                        novo_nome,
                        novo_telefone,
                        novo_email,
                        nova_cidade,
                        nova_uf,
                    )
                    st.session_state["bases"] = carregar_bases_supabase()
                    st.success(f"✅ Instrutor {acao} com STATUS = Ativo.")
                    st.rerun()
                except PermissionError as exc:
                    st.error(f"🚫 {exc}")
                except Exception as exc:
                    st.error(f"❌ Não foi possível cadastrar o instrutor: {exc}")
        else:
            st.caption(
                "🔒 Alteração de atividade e cadastro de instrutores disponíveis "
                "somente para administradores."
            )


elif modulo == "📂 Relatórios & Exportação":
    render_section_header(
        "📂",
        "Relatórios & Exportação",
        "Download das bases atualizadas"
    )

    st.markdown("### 📥 Exportar base do CRM")

    # CSV: formato simples e compatível com praticamente qualquer sistema.
    csv_buffer = df_base.to_csv(index=False).encode("utf-8-sig")

    # Excel: gera um .xlsx real em memória, sem criar arquivo temporário no servidor.
    # O openpyxl é usado pelo pandas para gravar o arquivo Excel.
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df_base.to_excel(writer, index=False, sheet_name="Base CRM")

        # Ajustes simples de apresentação da planilha.
        worksheet = writer.sheets["Base CRM"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        # Limita a largura automática para evitar colunas gigantes.
        for coluna in worksheet.columns:
            max_len = 0
            letra = coluna[0].column_letter
            for celula in coluna[:200]:
                valor = "" if celula.value is None else str(celula.value)
                max_len = max(max_len, len(valor))
            worksheet.column_dimensions[letra].width = min(max(max_len + 2, 10), 45)

    excel_buffer.seek(0)

    col_csv, col_excel = st.columns(2)

    with col_csv:
        st.download_button(
            label="📄 Baixar Base em CSV",
            data=csv_buffer,
            file_name=f"Base_CRM_AmPm_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with col_excel:
        st.download_button(
            label="📊 Baixar Base em Excel",
            data=excel_buffer.getvalue(),
            file_name=f"Base_CRM_AmPm_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.caption(
        f"Base pronta para exportação: {len(df_base):,} registros e "
        f"{len(df_base.columns):,} colunas."
    )
