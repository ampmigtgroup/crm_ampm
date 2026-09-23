"""
Camada central de dados do CRM AmPm.

Regra: PostgreSQL/Supabase é a fonte de verdade.
Streamlit session_state é somente cache de tela.
Excel é somente importação/exportação.
Arquivos de orçamento ficam no Supabase Storage.
"""

import os
import re
import uuid
from datetime import datetime
from typing import Any, Dict, Iterable, Optional

import streamlit as st
from supabase import create_client


PROJECT_URL_DEFAULT = "https://nptazzfvwhhmotfrvgdj.supabase.co"
BUCKET_ORCAMENTOS = "crm-documentos-orcamentos"


def _client():
    url = str(st.secrets.get("SUPABASE_URL", PROJECT_URL_DEFAULT) or PROJECT_URL_DEFAULT).strip()
    key = str(
        st.secrets.get("SUPABASE_SECRET_KEY", "")
        or st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", "")
        or ""
    ).strip()
    if not key:
        raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    return create_client(url, key)


def _safe(value):
    if value is None:
        return None
    try:
        if hasattr(value, "item"):
            value = value.item()
    except Exception:
        pass
    return value


def _as_float(value, default=0.0):
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return default


def _budget_row_from_legacy(pv: str, budget: Dict[str, Any]) -> Dict[str, Any]:
    subtotal = sum(
        _as_float(item.get("Total"))
        for item in (budget.get("itens") or [])
        if isinstance(item, dict)
    )
    desconto = _as_float(budget.get("desconto"))
    total = _as_float(budget.get("_total_calculado"), subtotal - subtotal * desconto / 100)
    budget_id = str(budget.get("id") or uuid.uuid4())
    return {
        "id": budget_id,
        "pv_abadi": str(pv).strip(),
        "numero": str(budget.get("numero") or f"ORC-{datetime.now():%Y%m%d}-{uuid.uuid4().hex[:6].upper()}"),
        "titulo": str(budget.get("titulo") or ""),
        "responsavel": str(budget.get("responsavel") or ""),
        "validade_dias": int(budget.get("validade_dias") or 7),
        "condicao_pagamento": str(budget.get("condicao_pagamento") or "A definir"),
        "observacoes": str(budget.get("observacoes") or ""),
        "desconto": desconto,
        "subtotal": subtotal,
        "total": total,
        "status": str(budget.get("status") or "rascunho"),
        "created_by": str(st.session_state.get("username") or "admin"),
        "updated_at": datetime.now().isoformat(),
    }


def _items_from_budget(budget_id: str, items: Iterable[Dict[str, Any]]):
    rows = []
    for index, item in enumerate(items or [], start=1):
        if not isinstance(item, dict):
            continue
        produto = str(item.get("Produto") or item.get("Descrição") or "").strip()
        if not produto:
            continue
        rows.append({
            "orcamento_id": budget_id,
            "item": index,
            "produto": produto,
            "quantidade": _as_float(item.get("Quantidade", item.get("Dias", 1))),
            "unidade": str(item.get("Unidade") or "Unidade"),
            "valor_unitario": _as_float(item.get("Valor Unitário", item.get("Valor por Dia"))),
            "total": _as_float(item.get("Total")),
            "raw_data": {str(k): _safe(v) for k, v in item.items()},
        })
    return rows


def load_budgets(pv: Optional[str] = None) -> Dict[str, Dict[str, Any]]:
    client = _client()
    query = client.table("crm_orcamentos").select(
        "id,pv_abadi,numero,titulo,responsavel,validade_dias,condicao_pagamento,"
        "observacoes,desconto,subtotal,total,status,created_by,created_at,updated_at"
    )
    if pv:
        query = query.eq("pv_abadi", str(pv).strip())
    rows = (query.order("updated_at", desc=True).execute()).data or []

    result: Dict[str, Dict[str, Any]] = {}
    for row in rows:
        key = str(row.get("pv_abadi") or "").strip()
        if not key:
            continue
        budget = {
            "id": row.get("id"),
            "numero": row.get("numero", ""),
            "titulo": row.get("titulo", ""),
            "responsavel": row.get("responsavel", ""),
            "validade_dias": row.get("validade_dias", 7),
            "condicao_pagamento": row.get("condicao_pagamento", "A definir"),
            "observacoes": row.get("observacoes", ""),
            "desconto": _as_float(row.get("desconto")),
            "_total_calculado": _as_float(row.get("total")),
            "status": row.get("status", "rascunho"),
            "atualizado_em": row.get("updated_at") or row.get("created_at") or "",
            "itens": [],
            "documentos": [],
        }
        items = (
            client.table("crm_orcamento_itens")
            .select("item,produto,quantidade,unidade,valor_unitario,total")
            .eq("orcamento_id", row["id"])
            .order("item")
            .execute()
        ).data or []
        budget["itens"] = [
            {
                "Item": str(x.get("item", i + 1)),
                "Produto": x.get("produto", ""),
                "Quantidade": _as_float(x.get("quantidade")),
                "Unidade": x.get("unidade", "Unidade"),
                "Valor Unitário": _as_float(x.get("valor_unitario")),
                "Total": _as_float(x.get("total")),
            }
            for i, x in enumerate(items)
        ]
        docs = (
            client.table("crm_orcamento_documentos")
            .select("id,nome,storage_path,mime_type,tamanho_bytes,created_at")
            .eq("orcamento_id", row["id"])
            .order("created_at")
            .execute()
        ).data or []
        budget["documentos"] = [
            {
                "id": d.get("id"),
                "nome": d.get("nome", "Documento"),
                "storage_path": d.get("storage_path", ""),
                "mime_type": d.get("mime_type"),
                "tamanho": d.get("tamanho_bytes"),
                "arquivo": "",
            }
            for d in docs
        ]
        result[key] = budget
    return result


def save_budgets(data: Dict[str, Dict[str, Any]]) -> None:
    client = _client()
    for pv, budget in (data or {}).items():
        if not isinstance(budget, dict):
            continue
        pv_txt = str(pv).strip()
        if not pv_txt:
            continue

        row = _budget_row_from_legacy(pv_txt, budget)
        client.table("crm_orcamentos").upsert(row, on_conflict="id").execute()
        budget["id"] = row["id"]
        budget["atualizado_em"] = row["updated_at"]
        budget["_total_calculado"] = row["total"]

        client.table("crm_orcamento_itens").delete().eq("orcamento_id", row["id"]).execute()
        item_rows = _items_from_budget(row["id"], budget.get("itens") or [])
        if item_rows:
            client.table("crm_orcamento_itens").insert(item_rows).execute()

        _log_event(
            pv_txt,
            "orcamento_salvo",
            "crm_orcamento",
            {
                "orcamento_id": row["id"],
                "numero": row["numero"],
                "total": row["total"],
            },
        )


def upload_budget_documents(pv: str, budget_id: str, files) -> list:
    client = _client()
    saved = []
    pv_txt = re.sub(r"[^A-Za-z0-9_-]", "_", str(pv).strip()) or "pv"
    budget_txt = re.sub(r"[^A-Za-z0-9_-]", "_", str(budget_id).strip())
    for file in files or []:
        original = os.path.basename(getattr(file, "name", "documento"))
        safe_name = re.sub(r"[^A-Za-z0-9._-]", "_", original).strip(".") or "documento"
        path = f"{pv_txt}/{budget_txt}/{uuid.uuid4().hex[:10]}_{safe_name}"
        content = file.getvalue()
        mime = getattr(file, "type", None) or "application/octet-stream"
        client.storage.from_(BUCKET_ORCAMENTOS).upload(
            path,
            content,
            file_options={"content-type": mime, "upsert": "false"},
        )
        row = {
            "orcamento_id": budget_id,
            "nome": original,
            "storage_path": path,
            "mime_type": mime,
            "tamanho_bytes": len(content),
            "created_by": str(st.session_state.get("username") or "admin"),
        }
        inserted = client.table("crm_orcamento_documentos").insert(row).execute().data or []
        saved.append({
            "id": inserted[0].get("id") if inserted else None,
            "nome": original,
            "storage_path": path,
            "mime_type": mime,
            "tamanho": len(content),
            "arquivo": "",
        })
    _log_event(pv_txt, "documentos_orcamento_anexados", "crm_orcamento", {"quantidade": len(saved), "orcamento_id": budget_id})
    return saved


def download_budget_document(storage_path: str) -> bytes:
    if not storage_path:
        return b""
    return _client().storage.from_(BUCKET_ORCAMENTOS).download(storage_path)


def _log_event(pv: Optional[str], event_type: str, origin: str, payload: Dict[str, Any]):
    try:
        _client().table("crm_eventos").insert({
            "pv_abadi": str(pv).strip() if pv else None,
            "tipo": event_type,
            "origem": origin,
            "usuario": str(st.session_state.get("username") or "admin"),
            "payload": payload or {},
        }).execute()
    except Exception:
        pass


def registrar_evento(pv: Optional[str], event_type: str, origin: str, payload: Optional[Dict[str, Any]] = None):
    _log_event(pv, event_type, origin, payload or {})


def load_unified_view(limit: Optional[int] = None):
    query = _client().table("crm_visao_unificada").select("*")
    if limit:
        query = query.limit(int(limit))
    return query.execute().data or []
