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
    linhas = []
    for item in itens or []:
        if not isinstance(item, dict):
            continue
        produto = str(item.get("Produto", item.get("Descrição", item.get("descricao", ""))) or "").strip()
        if not produto:
            continue
        if produto not in PRODUTOS_TREINAMENTO:
            produto = str(item.get("Descrição", produto) or produto).strip()
        try:
            dias = float(item.get("Dias", item.get("Qtd", item.get("qtd", 1))) or 0)
        except (TypeError, ValueError):
            dias = 0.0
        try:
            valor_dia = float(item.get("Valor por Dia", item.get("Valor Unitário", item.get("valor_unitario", 0))) or 0)
        except (TypeError, ValueError):
            valor_dia = 0.0
        total = dias * valor_dia
        linhas.append({
            "Item": str(item.get("Item", len(linhas) + 1)),
            "Produto": produto,
            "Dias": dias,
            "Valor por Dia": valor_dia,
            "Total": total,
        })
    return linhas


def _gerar_excel_orcamento(orcamento, posto, pv):
    itens = _normalizar_itens_orcamento(orcamento.get("itens", []))
    colunas = ["Item", "Produto", "Dias", "Valor por Dia", "Total"]
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
        df_itens.to_excel(writer, index=False, sheet_name="Treinamentos")

        ws = writer.sheets["Orçamento"]
        ws.freeze_panes = "A2"
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 65

        ws2 = writer.sheets["Treinamentos"]
        ws2.freeze_panes = "A2"
        ws2.auto_filter.ref = ws2.dimensions
        for col in ws2.columns:
            letra = col[0].column_letter
            maior = max([len(str(c.value or "")) for c in col[:100]] + [10])
            ws2.column_dimensions[letra].width = min(maior + 2, 45)
        for linha in range(2, ws2.max_row + 1):
            ws2.cell(linha, 4).number_format = 'R$ #,##0.00'
            ws2.cell(linha, 5).number_format = 'R$ #,##0.00'

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

    st.markdown("#### 🧾 Treinamentos")
    st.caption("Os três produtos abaixo são os tipos de treinamento comercializados. O cálculo é feito por dias de treinamento.")
    itens_iniciais = _normalizar_itens_orcamento(orcamento.get("itens", []))
    df_itens_inicial = pd.DataFrame(
        itens_iniciais,
        columns=["Item", "Produto", "Dias", "Valor por Dia", "Total"],
    )
    if df_itens_inicial.empty:
        df_itens_inicial = pd.DataFrame([
            {"Item": 1, "Produto": PRODUTOS_TREINAMENTO[0], "Dias": 1.0, "Valor por Dia": 0.0, "Total": 0.0}
        ])

    df_editado = st.data_editor(
        df_itens_inicial,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "Item": st.column_config.TextColumn("Item", disabled=True),
            "Produto": st.column_config.SelectboxColumn("Produto", options=PRODUTOS_TREINAMENTO, required=True),
            "Dias": st.column_config.NumberColumn("Dias de treinamento", min_value=0.5, step=0.5),
            "Valor por Dia": st.column_config.NumberColumn("Valor por Dia (R$)", min_value=0.0, step=0.01, format="R$ %.2f"),
            "Total": st.column_config.NumberColumn("Total (R$)", disabled=True, format="R$ %.2f"),
        },
        key=f"orcamento_editor_{chave}",
    ).copy()

    if not df_editado.empty:
        df_editado["Total"] = (
            pd.to_numeric(df_editado["Dias"], errors="coerce").fillna(0)
            * pd.to_numeric(df_editado["Valor por Dia"], errors="coerce").fillna(0)
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
        st.metric("💰 Total do treinamento", f"R$ {total:,.2f}")

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
                "Dias": float(linha.get("Dias", 0) or 0),
                "Valor por Dia": float(linha.get("Valor por Dia", 0) or 0),
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
        st.success("✅ Orçamento de treinamento salvo com sucesso!")

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
    "enriquecimento": "📇 Enriquecimento de Rede",
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
    st.cache_data.clear()

    return acao

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
            "CNPJ": ["cnpj", "cnpj loja", "documento cnpj", "cnpj da loja", "cnpj posto", "cnpj completo"],
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

def render_importador_inteligente():
    st.markdown("## 📥 Importador Inteligente de Planilhas")
    st.caption("Upload de Excel, leitura de múltiplas abas, sugestão automática de campos, revisão e merge por PV/CNPJ.")
    arquivo=st.file_uploader("Escolha um arquivo Excel",type=["xlsx","xls"],key="importador_inteligente_upload")
    if not arquivo:
        st.info("Envie uma planilha para iniciar a análise automática.");return
    try:abas=_imp_ler_planilha(arquivo)
    except Exception as e:st.error(f"Não foi possível ler a planilha: {e}");return
    if not abas:st.warning("Nenhuma aba com dados utilizáveis foi encontrada.");return
    aba=st.selectbox("📑 Aba para importar",list(abas.keys()));df_imp=abas[aba]
    st.markdown(f"**Prévia:** {len(df_imp)} registros × {len(df_imp.columns)} colunas")
    st.dataframe(df_imp.head(10),use_container_width=True,hide_index=True)
    opcoes=["— Não importar esta coluna —"]+list(IMPORTADOR_CAMPOS.keys());mapa={}
    for coluna in df_imp.columns:
        sug,score=_imp_sugerir_campo(str(coluna));idx=opcoes.index(sug) if sug in opcoes else 0
        escolha=st.selectbox(f"`{coluna}`",opcoes,index=idx,key=f"imp_map_{_imp_norm(aba)}_{_imp_norm(coluna)}");mapa[coluna]=None if escolha.startswith("—") else escolha
        if mapa[coluna]:
            pct=round(score*100);st.caption(("🟢" if pct>=85 else "🟡" if pct>=65 else "🔴")+f" Confiança sugerida: {pct}%")
    vals=[v for v in mapa.values() if v]
    if len(vals)!=len(set(vals)):st.warning("⚠️ Há campos do CRM recebendo mais de uma coluna. Revise o mapeamento.")
    registros=_imp_processar_df(df_imp,mapa);validos=[r for r in registros if r["chave"]];sem=len(registros)-len(validos)
    c1,c2,c3=st.columns(3);c1.metric("Registros lidos",len(registros));c2.metric("Com PV/CNPJ",len(validos));c3.metric("Sem chave",sem)
    if sem:st.warning("Registros sem PV e sem CNPJ não entram no merge automático.")
    if not validos:st.error("Nenhum registro possui PV ou CNPJ utilizável.");return
    confirmar=st.checkbox("Revisei o mapeamento e autorizo atualizar a base do CRM.",key="imp_confirmacao")
    if st.button("🚀 Importar para o CRM",type="primary",use_container_width=True):
        if not confirmar:st.error("Marque a confirmação antes de importar.");return
        base=st.session_state["bases"].get("lojas",pd.DataFrame()).copy();nova,ins,atu=_imp_merge_com_base(base,validos);st.session_state["bases"]["lojas"]=nova
        rel={"arquivo":arquivo.name,"aba":aba,"lidos":len(registros),"validos":len(validos),"inseridos":ins,"atualizados":atu,"sem_chave":sem,"data":datetime.now().strftime("%d/%m/%Y %H:%M:%S")};st.session_state["relatorio_importacao"]=rel
        try:
            if "salvar_bases_no_disco" in globals():salvar_bases_no_disco()
            elif "salvar_base_unificada" in globals():salvar_base_unificada()
        except Exception as e:st.warning(f"Dados aplicados na sessão; persistência precisa ser verificada: {e}")
        st.success(f"✅ Importação concluída: {ins} novo(s) e {atu} atualizado(s).");st.json(rel)


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

    st.divider()

    st.markdown("📥 **Atualizar Banco de Dados**")
    st.caption("O CRM interpreta o conteúdo, reconhece colunas parecidas e faz atualização incremental: registros novos entram e registros existentes são atualizados sem apagar dados válidos.")

    aba_destino_csv = st.selectbox(
        "Destino para CSV:",
        ["Rede_de_Lojas", "Fila_CallCenter", "Previsao_Inauguracao", "Instrutores", "Recomendacao_Deslocamento"]
    )

    uploaded_file = st.file_uploader(
        "Envie a nova planilha (.xlsx ou .csv):",
        type=["xlsx", "xls", "csv"]
    )
    assinatura_upload = None
    if uploaded_file is not None and not (
        assinatura_upload
        and assinatura_upload == st.session_state.get("upload_processado_assinatura")
        and st.session_state.get("upload_processado_ok")
    ):
        try:
            assinatura_upload = _fingerprint_upload(uploaded_file.getvalue())
        except Exception:
            assinatura_upload = None

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
            if uploaded_file.name.lower().endswith((".xlsx", ".xls")):
                conteudo = uploaded_file.getvalue()
                st.session_state["ultimo_upload_info"] = {
                    "arquivo": uploaded_file.name,
                    "bytes": len(conteudo),
                    "zip_valido": zipfile.is_zipfile(io.BytesIO(conteudo)),
                }
                bases_validadas, relatorio, erro = validar_bytes_excel(conteudo)
                if erro:
                    info = st.session_state.get("ultimo_upload_info", {})
                    st.error(f"❌ Arquivo rejeitado:\n\n{erro}")
                    if info:
                        st.caption(
                            f"Diagnóstico: {info.get('arquivo', 'arquivo')} — "
                            f"{info.get('bytes', 0):,} bytes — "
                            f"pacote XLSX/ZIP: {'sim' if info.get('zip_valido') else 'não'}"
                        )
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
                    st.session_state["upload_processado_assinatura"] = assinatura_upload
                    st.session_state["upload_processado_ok"] = True
                    st.success(
                        f"✅ Importação concluída: {total_novos} novos e {total_atualizados} atualizados."
                    )
                    _exibir_relatorio_importacao(relatorio)

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

    elif uploaded_file is not None and st.session_state.get("upload_processado_ok"):
        st.info(
            f"✅ A planilha **{uploaded_file.name}** já foi processada nesta sessão. "
            "Remova o arquivo e envie novamente para importar uma nova versão."
        )

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

def _valor_historico_instrutor(posto, *campos):
    """Retorna o primeiro nome histórico preenchido, sem verificar se está ativo."""
    for campo in campos:
        if campo not in posto:
            continue
        valor = posto.get(campo)
        if valor is None:
            continue
        try:
            if pd.isna(valor):
                continue
        except (TypeError, ValueError):
            pass
        texto = str(valor).strip()
        if texto and texto.casefold() not in {"nan", "none", "nat", "-"}:
            return texto
    return "Não informado"

def _historico_instrutor_procv(posto):
    """Mostra o histórico do instrutor mesmo quando ele não está mais ativo."""
    treinamento = _valor_historico_instrutor(posto, "Instrutor_Treinamento", "Instrutor_Sugerido")
    inauguracao = _valor_historico_instrutor(posto, "Instrutor_Inauguracao")
    return treinamento, inauguracao


if modulo == "🛡️ Administração":
    render_administracao()

elif modulo == "📊 Dashboard Executivo":
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
        cols_mostrar = [c for c in ['PV Abadi', 'Razão Social', 'CNPJ', 'Municipio', 'UF', 'Status Loja', 'Tipo_Necessidade', 'Instrutor_Treinamento', 'Instrutor_Sugerido', 'Status_Contato'] if c in df_view.columns]
        evento = st.dataframe(df_view[cols_mostrar], use_container_width=True, hide_index=True, selection_mode="single-row", on_select="rerun")
        linhas = evento.selection.get("rows", [])
        if linhas:
            p = df_view.iloc[linhas[0]].to_dict()
            instrutor_treinamento, instrutor_inauguracao = _historico_instrutor_procv(p)
            st.divider()
            st.markdown(f"**📋 Gestão e Franquia AMPM — PV {p.get('PV Abadi', '-')} · {p.get('Razão Social', '-')}**")
            k1, k2, k3 = st.columns(3)
            with k1:
                st.markdown(f"""<div class="procv-card"><h4>🏪 Cadastro da Loja</h4><p>🧾 <b>CNPJ:</b> {p.get('CNPJ', 'Não informado')}</p><p>📍 <b>Endereço:</b> {p.get('Endereço', '-')}</p><p>🏙️ <b>Cidade/UF:</b> {p.get('Municipio', '-')}/{p.get('UF', '-')}</p><p>⚙️ <b>Status da loja:</b> {p.get('Status Loja', '-')}</p></div>""", unsafe_allow_html=True)
            with k2:
                st.markdown(f"""<div class="procv-card"><h4>👔 Gestão & Franquia</h4><p>👤 <b>Gerente (GF):</b> {p.get('GF', '-')}</p><p>👔 <b>Consultor (CF):</b> {p.get('CF', '-')}</p><p>📅 <b>Inauguração:</b> {p.get('Previsão Inauguração', 'N/A')}</p></div>""", unsafe_allow_html=True)
            with k3:
                st.markdown(f"""<div class="procv-card"><h4>👨‍🏫 Histórico de Treinamento</h4><p>🎓 <b>Instrutor do treinamento:</b> {instrutor_treinamento}</p><p>🚀 <b>Instrutor da inauguração:</b> {instrutor_inauguracao}</p><p>📅 <b>Último treinamento:</b> {p.get('Data_Ultimo_Treinamento', 'Não informado')}</p><p>🎯 <b>Necessidade atual:</b> {p.get('Tipo_Necessidade', '-')}</p><p>🔄 <b>Status do atendimento:</b> {badge_status_html(p.get('Status_Contato', '-'))}</p></div>""", unsafe_allow_html=True)
    else:
        st.info("📭 Nenhum dado carregado ainda.")

elif modulo == "📍 Calculadora & Otimizador de Custos":
    render_section_header(
        "📍",
        "Calculadora & Otimizador de Custos",
        "Simulação de rotas, custos estimados e comparação entre instrutores"
    )

    if not df_rec.empty:
        df_rec_filtrado = df_rec.copy()
        if filtro_uf != "Todas" and 'UF_Loja' in df_rec_filtrado.columns:
            df_rec_filtrado = df_rec_filtrado[df_rec_filtrado['UF_Loja'] == filtro_uf]

        postos_unicos = df_rec_filtrado[
            ['PV_ABADI', 'Razao_Social', 'Municipio_Loja', 'UF_Loja']
        ].drop_duplicates()

        if not postos_unicos.empty:
            postos_unicos['label'] = (
                postos_unicos['PV_ABADI'].astype(str)
                + " - " + postos_unicos['Razao_Social'].astype(str)
                + " (" + postos_unicos['Municipio_Loja'].astype(str)
                + "/" + postos_unicos['UF_Loja'].astype(str) + ")"
            )

            posto_sel = st.selectbox(
                "⛽ Selecione o Posto Alvo:",
                postos_unicos['label'].tolist()
            )
            pv_sel = int(posto_sel.split(" - ")[0])

            top_3 = (
                df_rec_filtrado[df_rec_filtrado['PV_ABADI'] == pv_sel]
                .sort_values(by='Ranking_Proximidade')
                .head(3)
            )

            if not top_3.empty:
                primeira = top_3.iloc[0]

                # Mapa continua mostrando o instrutor mais próximo.
                if (
                    pd.notna(primeira.get('Lat_Loja'))
                    and pd.notna(primeira.get('Lon_Loja'))
                    and pd.notna(primeira.get('Lat_Instrutor'))
                    and pd.notna(primeira.get('Lon_Instrutor'))
                ):
                    p_lat, p_lon = float(primeira['Lat_Loja']), float(primeira['Lon_Loja'])
                    i_lat, i_lon = float(primeira['Lat_Instrutor']), float(primeira['Lon_Instrutor'])

                    df_mapa_pontos = pd.DataFrame([
                        {
                            "name": f"Posto {primeira['PV_ABADI']}",
                            "lat": p_lat,
                            "lon": p_lon,
                            "color": [226, 123, 0, 220]
                        },
                        {
                            "name": f"Instrutor {primeira['Instrutor_Sugerido']}",
                            "lat": i_lat,
                            "lon": i_lon,
                            "color": [76, 175, 80, 220]
                        }
                    ])
                    df_mapa_arco = pd.DataFrame([{
                        "from_lat": i_lat,
                        "from_lon": i_lon,
                        "to_lat": p_lat,
                        "to_lon": p_lon
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
                    view_state = pdk.ViewState(
                        latitude=(p_lat + i_lat) / 2,
                        longitude=(p_lon + i_lon) / 2,
                        zoom=5,
                        pitch=40
                    )
                    st.pydeck_chart(
                        pdk.Deck(
                            layers=[layer_pontos, layer_arco],
                            initial_view_state=view_state,
                            tooltip={"text": "{name}"}
                        )
                    )

                st.markdown("### 💰 Composição estimada de custos")
                st.caption(
                    "Os valores abaixo são estimativas e podem ser alterados para cada instrutor. "
                    "As caixas de seleção definem quais custos entram no cálculo."
                )

                # Valores iniciais de referência. Não são valores fixos da empresa.
                # No futuro poderão ser substituídos por uma tabela de valores fixos.
                st.markdown(
                    "**Valores iniciais de referência:** "
                    "diária R$ 280/dia · hospedagem R$ 250/dia · carro R$ 180/dia · "
                    "rodoviário calculado pela distância · avião R$ 800/deslocamento · "
                    "treinamento R$ 280/dia."
                )

                resultados_custos = []

                col1, col2, col3 = st.columns(3)
                cols = [col1, col2, col3]

                for idx, (_, row) in enumerate(top_3.iterrows()):
                    nome_instrutor = str(row.get('Instrutor_Sugerido', 'Instrutor não informado'))
                    dist = float(row.get('Distancia_km_linha_reta', 0) or 0)

                    try:
                        dias_base = float(row.get('Dias_Treinamento_Necessarios', 1) or 1)
                    except (TypeError, ValueError):
                        dias_base = 1.0
                    dias_base = max(0.5, dias_base)

                    # Cada instrutor possui suas próprias flags e valores.
                    with cols[idx]:
                        st.markdown(
                            f"""
                            <div class="top-instructor-card">
                                <h4>#{idx+1} {nome_instrutor}</h4>
                                <p>Origem: {row.get('Cidade_Instrutor', '-')} / {row.get('UF_Instrutor', '-')}</p>
                                <p>Distância: {dist:.1f} km</p>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )

                        with st.expander("⚙️ Configurar custos deste instrutor", expanded=True):
                            dias = st.number_input(
                                "📅 Dias de treinamento",
                                min_value=0.5,
                                value=float(dias_base),
                                step=0.5,
                                key=f"custos_dias_{pv_sel}_{idx}"
                            )

                            st.markdown("**Selecione os custos que serão considerados:**")

                            usar_diaria = st.checkbox(
                                "☑ Diárias",
                                value=True,
                                key=f"usar_diaria_{pv_sel}_{idx}"
                            )
                            valor_diaria = st.number_input(
                                "Valor da diária (R$/dia)",
                                min_value=0.0,
                                value=280.0,
                                step=10.0,
                                key=f"valor_diaria_{pv_sel}_{idx}"
                            )

                            usar_hospedagem = st.checkbox(
                                "☑ Hospedagem",
                                value=True,
                                key=f"usar_hospedagem_{pv_sel}_{idx}"
                            )
                            valor_hospedagem = st.number_input(
                                "Hospedagem (R$/dia)",
                                min_value=0.0,
                                value=250.0,
                                step=10.0,
                                key=f"valor_hospedagem_{pv_sel}_{idx}"
                            )

                            usar_carro = st.checkbox(
                                "☑ Aluguel de carro",
                                value=False,
                                key=f"usar_carro_{pv_sel}_{idx}"
                            )
                            valor_carro = st.number_input(
                                "Aluguel de carro (R$/dia)",
                                min_value=0.0,
                                value=180.0,
                                step=10.0,
                                key=f"valor_carro_{pv_sel}_{idx}"
                            )

                            usar_rodoviario = st.checkbox(
                                "☐ Deslocamento rodoviário",
                                value=False,
                                key=f"usar_rodoviario_{pv_sel}_{idx}"
                            )
                            valor_rodoviario = st.number_input(
                                "Deslocamento rodoviário (R$/viagem)",
                                min_value=0.0,
                                value=float(max(0.0, dist * 2 * 2.10)),
                                step=10.0,
                                key=f"valor_rodoviario_{pv_sel}_{idx}"
                            )

                            usar_aviao = st.checkbox(
                                "☐ Deslocamento de avião",
                                value=False,
                                key=f"usar_aviao_{pv_sel}_{idx}"
                            )
                            valor_aviao = st.number_input(
                                "Deslocamento de avião (R$/viagem)",
                                min_value=0.0,
                                value=800.0,
                                step=50.0,
                                key=f"valor_aviao_{pv_sel}_{idx}"
                            )

                            usar_treinamento = st.checkbox(
                                "☑ Valor do treinamento",
                                value=True,
                                key=f"usar_treinamento_{pv_sel}_{idx}"
                            )
                            valor_treinamento = st.number_input(
                                "Valor do treinamento (R$/dia)",
                                min_value=0.0,
                                value=280.0,
                                step=10.0,
                                key=f"valor_treinamento_{pv_sel}_{idx}"
                            )

                            subtotal_diaria = valor_diaria * dias if usar_diaria else 0.0
                            subtotal_hospedagem = valor_hospedagem * dias if usar_hospedagem else 0.0
                            subtotal_carro = valor_carro * dias if usar_carro else 0.0
                            subtotal_rodoviario = valor_rodoviario if usar_rodoviario else 0.0
                            subtotal_aviao = valor_aviao if usar_aviao else 0.0
                            subtotal_treinamento = valor_treinamento * dias if usar_treinamento else 0.0

                            custo_total = (
                                subtotal_diaria
                                + subtotal_hospedagem
                                + subtotal_carro
                                + subtotal_rodoviario
                                + subtotal_aviao
                                + subtotal_treinamento
                            )

                            if usar_rodoviario and usar_aviao:
                                st.warning(
                                    "⚠️ Rodoviário e avião estão selecionados ao mesmo tempo. "
                                    "O sistema soma os dois. Se for uma alternativa de transporte, "
                                    "desmarque um deles."
                                )

                            st.markdown("**Resumo deste instrutor**")
                            st.write(f"Diárias: R$ {subtotal_diaria:,.2f}")
                            st.write(f"Hospedagem: R$ {subtotal_hospedagem:,.2f}")
                            st.write(f"Carro: R$ {subtotal_carro:,.2f}")
                            st.write(f"Rodoviário: R$ {subtotal_rodoviario:,.2f}")
                            st.write(f"Avião: R$ {subtotal_aviao:,.2f}")
                            st.write(f"Treinamento: R$ {subtotal_treinamento:,.2f}")
                            st.metric("💰 Custo total estimado", f"R$ {custo_total:,.2f}")

                            resultados_custos.append({
                                "Instrutor": nome_instrutor,
                                "Dias": dias,
                                "Diárias": subtotal_diaria,
                                "Hospedagem": subtotal_hospedagem,
                                "Carro": subtotal_carro,
                                "Rodoviário": subtotal_rodoviario,
                                "Avião": subtotal_aviao,
                                "Treinamento": subtotal_treinamento,
                                "Total Estimado": custo_total,
                            })

                if resultados_custos:
                    st.divider()
                    st.markdown("### 📊 Comparativo dos instrutores sugeridos")
                    df_custos = pd.DataFrame(resultados_custos).sort_values(
                        "Total Estimado", ascending=True
                    )
                    st.dataframe(
                        df_custos,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Diárias": st.column_config.NumberColumn("Diárias", format="R$ %.2f"),
                            "Hospedagem": st.column_config.NumberColumn("Hospedagem", format="R$ %.2f"),
                            "Carro": st.column_config.NumberColumn("Carro", format="R$ %.2f"),
                            "Rodoviário": st.column_config.NumberColumn("Rodoviário", format="R$ %.2f"),
                            "Avião": st.column_config.NumberColumn("Avião", format="R$ %.2f"),
                            "Treinamento": st.column_config.NumberColumn("Treinamento", format="R$ %.2f"),
                            "Total Estimado": st.column_config.NumberColumn("Total Estimado", format="R$ %.2f"),
                        }
                    )

                    melhor = df_custos.iloc[0]
                    st.success(
                        f"🏆 Menor custo estimado entre os instrutores selecionados: "
                        f"{melhor['Instrutor']} — R$ {melhor['Total Estimado']:,.2f}"
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
                        nome_c = st.text_input("👤 Nome do Responsável na Loja:", value=_texto_seguro_callcenter(posto.get('Nome_Contato', '')))
                        tel_c = st.text_input("📞 Telefone de Contato:", value=_texto_seguro_callcenter(posto.get('Telefone_Contato', '')))
                        email_c = st.text_input("✉️ E-mail do Contato:", value=_texto_seguro_callcenter(posto.get('Email_Contato', '')))
                        tem_func_atual = _texto_seguro_callcenter(posto.get('Tem_Funcionarios', 'Sim')) or 'Sim'
                        tem_func_opcoes = ["Sim", "Não"]
                        qtd_func_atual = _texto_seguro_callcenter(posto.get('Qtd_Funcionarios', 0))
                        try:
                            qtd_func_padrao = int(float(qtd_func_atual or 0))
                        except (TypeError, ValueError):
                            qtd_func_padrao = 0
                        qtd_func_padrao = max(0, qtd_func_padrao)

                        # Regra de negócio: zero funcionários significa automaticamente "Não".
                        # Isso também corrige registros antigos que estejam inconsistentes.
                        if qtd_func_padrao == 0:
                            tem_func_atual = "Não"
                        elif tem_func_atual not in tem_func_opcoes:
                            tem_func_atual = "Sim"

                        idx_tem_func = tem_func_opcoes.index(tem_func_atual)
                        tem_funcionarios = st.selectbox("👥 Há funcionários para treinar?", tem_func_opcoes, index=idx_tem_func)
                        qtd_func = st.number_input(
                            "🔢 Qtd. de Funcionários para Treinar:",
                            value=qtd_func_padrao, min_value=0, step=1,
                            disabled=(tem_funcionarios == "Não")
                        )
                        # Se a quantidade for zero, a regra prevalece sobre a seleção anterior.
                        if qtd_func == 0:
                            tem_funcionarios = "Não"
                        elif tem_funcionarios == "Não":
                            qtd_func = 0
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

                    obs = st.text_area("💬 Observações e Alinhamentos:", value=_texto_seguro_callcenter(posto.get('Observacoes', '')), height=80)

                    if st.form_submit_button("💾 Salvar Registro do Atendimento"):
                        atualizar_fila(pv_alvo, {
                            'Nome_Contato': nome_c,
                            'Telefone_Contato': tel_c,
                            'Email_Contato': email_c,
                            'Tem_Funcionarios': 'Não' if int(qtd_func) == 0 else 'Sim',
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
        "Instrutores atualmente em atividade"
    )

    instrutores_ativos = filtrar_instrutores_ativos(df_instrutores)

    # O banco continua contendo também os instrutores que já saíram.
    # Aqui exibimos somente quem está explicitamente com STATUS = Ativo.
    if "STATUS" not in df_instrutores.columns and not df_instrutores.empty:
        st.warning(
            "⚠️ A base de instrutores não possui a coluna `STATUS`. "
            "Por segurança, os registros estão sendo exibidos, mas "
            "adicione essa coluna com `Ativo` ou o status correspondente."
        )

    total_base = len(df_instrutores) if df_instrutores is not None else 0
    total_ativos = len(instrutores_ativos)

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.metric("👔 Instrutores em atividade", total_ativos)
    with col_m2:
        st.metric("📚 Registros no banco", total_base)

    if not instrutores_ativos.empty:
        st.dataframe(
            instrutores_ativos,
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("📭 Nenhum instrutor com STATUS = Ativo foi encontrado.")

    st.divider()

    # Somente usuários explicitamente definidos em ADMIN_USERNAMES
    # conseguem abrir o formulário de cadastro.
    if usuario_e_admin():
        st.markdown("### ➕ Cadastrar novo instrutor")
        st.caption(
            "Somente administradores podem adicionar instrutores. "
            "O novo cadastro será gravado com STATUS = Ativo."
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
                    "Status",
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
                st.success(f"✅ Instrutor {acao} com STATUS = Ativo.")
                st.rerun()
            except PermissionError as exc:
                st.error(f"🚫 {exc}")
            except Exception as exc:
                st.error(f"❌ Não foi possível cadastrar o instrutor: {exc}")
    else:
        st.caption("🔒 Cadastro de novos instrutores disponível somente para administradores.")

elif modulo == "📇 Enriquecimento de Rede":
    render_section_header("📇", "Enriquecimento de Rede", "Atualizações de lojas e telefones")
    st.info("Utilize a barra lateral para fazer upload de novas bases ou enriquecer os dados existentes.")

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

