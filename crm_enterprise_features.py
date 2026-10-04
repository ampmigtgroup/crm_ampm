import math
from datetime import date, datetime, timedelta
from urllib.parse import quote

import pandas as pd
import streamlit as st
from supabase import create_client

MODULE_KEY = "gestao360"
MODULE_LABEL = "🧭 Gestão 360 / Enterprise"


def _client():
    try:
        url = st.secrets.get("SUPABASE_URL", "") or st.secrets.get("SUPABASE_PROJECT_URL", "")
        key = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", "") or st.secrets.get("SUPABASE_SECRET_KEY", "")
        return create_client(url, key) if url and key else None
    except Exception:
        return None


def _user():
    return str(st.session_state.get("username") or st.session_state.get("user") or st.session_state.get("name") or "sistema")


def _now():
    return datetime.now().isoformat(timespec="seconds")


def audit(entidade, entidade_id, acao, antes=None, depois=None):
    c = _client()
    if not c:
        return False
    try:
        a = antes or {}
        d = depois or {}
        diff = {k: {"de": a.get(k), "para": d.get(k)} for k in set(a) | set(d) if str(a.get(k)) != str(d.get(k))}
        c.table("crm_eventos").insert({
            "pv_abadi": str(entidade_id) if str(entidade) in {"loja","contato","orcamento","oportunidade"} else None,
            "tipo": str(acao),
            "origem": "gestao360",
            "usuario": _user(),
            "payload": {"entidade": entidade, "entidade_id": str(entidade_id), "antes": a, "depois": d, "diff": diff},
        }).execute()
        return True
    except Exception:
        return False


def config(chave, padrao):
    c = _client()
    if not c:
        return padrao
    try:
        r = c.table("crm_configuracoes").select("valor").eq("chave", chave).limit(1).execute()
        return r.data[0]["valor"] if r.data else padrao
    except Exception:
        return padrao


def save_config(chave, valor):
    c = _client()
    if not c:
        return False
    try:
        c.table("crm_configuracoes").upsert(
            {"chave": chave, "valor": valor, "updated_by": _user(), "updated_at": _now()},
            on_conflict="chave",
        ).execute()
        return True
    except Exception:
        return False


def _dias(row):
    for k in ("dias_desde_ultimo_treinamento", "Dias_Desde_Ultimo_Treinamento"):
        if k in row and pd.notna(row.get(k)):
            try:
                return int(float(row.get(k)))
            except Exception:
                pass
    for k in ("data_ultimo_treinamento", "Data_Ultimo_Treinamento"):
        if k in row:
            d = pd.to_datetime(row.get(k), errors="coerce")
            if pd.notna(d):
                return max(0, (pd.Timestamp(date.today()) - d.normalize()).days)
    return None


def _sla(df, limite):
    if df is None or df.empty:
        return pd.DataFrame()
    out = []
    for _, r in df.iterrows():
        dias = _dias(r)
        status = "NUNCA TREINOU" if dias is None else ("ESTOURADO" if dias > limite else ("A VENCER" if dias >= limite - 30 else "DENTRO"))
        out.append({
            "PV": r.get("PV Abadi", r.get("pv_abadi", "")),
            "Loja": r.get("Razão Social", r.get("razao_social", "")),
            "Cidade": r.get("Municipio", r.get("municipio", "")),
            "UF": r.get("UF", r.get("uf", "")),
            "Consultor": r.get("CF", r.get("cf", "")),
            "Dias sem treinamento": dias if dias is not None else "-",
            "SLA": status,
        })
    return pd.DataFrame(out)


def _carteira(df):
    if df is None or df.empty:
        return pd.DataFrame()
    c = "CF" if "CF" in df.columns else ("cf" if "cf" in df.columns else None)
    if not c:
        return pd.DataFrame()
    x = df.copy()
    x["_sem"] = x.apply(lambda r: _dias(r) is None, axis=1)
    out = x.groupby(c, dropna=False).agg(lojas=(c, "size"), sem_treinamento=("_sem", "sum")).reset_index()
    out = out.rename(columns={c: "Consultor"})
    out["cobertura_%"] = ((out["lojas"] - out["sem_treinamento"]) / out["lojas"].clip(lower=1) * 100).round(1)
    return out.sort_values(["sem_treinamento", "lojas"], ascending=False)


def _km(a,b,c,d):
    try:
        p = math.pi/180
        x = 0.5-math.cos((c-a)*p)/2+math.cos(a*p)*math.cos(c*p)*(1-math.cos((d-b)*p))/2
        return 12742*math.asin(math.sqrt(max(0,x)))
    except Exception:
        return None


def _agenda(df_lojas, df_instrutores, inicio):
    if df_lojas is None or df_lojas.empty or df_instrutores is None or df_instrutores.empty:
        return pd.DataFrame()
    lojas, instr = df_lojas.copy(), df_instrutores.copy()
    if "Status Loja" in lojas:
        lojas = lojas[lojas["Status Loja"].astype(str).str.contains("ativ|futur", case=False, na=False)]
    if "Status" in instr:
        instr = instr[instr["Status"].astype(str).str.contains("ativ", case=False, na=False)]
    rows = []
    for _, i in instr.iterrows():
        ilat, ilon = i.get("Lat_Instrutor", i.get("Lat", i.get("lat"))), i.get("Lon_Instrutor", i.get("Lon", i.get("lon")))
        if pd.isna(ilat) or pd.isna(ilon):
            continue
        candidatos = []
        for _, l in lojas.head(500).iterrows():
            llat, llon = l.get("Lat_Loja", l.get("Lat", l.get("lat"))), l.get("Lon_Loja", l.get("Lon", l.get("lon")))
            if pd.isna(llat) or pd.isna(llon):
                continue
            km = _km(float(ilat), float(ilon), float(llat), float(llon))
            if km is not None:
                candidatos.append((km,l))
        candidatos.sort(key=lambda x:x[0])
        for n,(km,l) in enumerate(candidatos[:5]):
            rows.append({
                "Data": inicio + timedelta(days=n % 5),
                "Instrutor": i.get("nome_completo", i.get("Nome", "")),
                "PV": l.get("PV Abadi", l.get("pv_abadi", "")),
                "Loja": l.get("Razão Social", l.get("razao_social", "")),
                "Cidade": l.get("Municipio", l.get("municipio", "")),
                "UF": l.get("UF", l.get("uf", "")),
                "Distância km": round(km,1),
            })
    return pd.DataFrame(rows)


def render(df_base, df_lojas=None, df_instrutores=None):
    st.markdown("# 🧭 Gestão 360 / Enterprise")
    st.caption("Requisitos da nova especificação incorporados à arquitetura original Streamlit + Supabase.")
    tabs = st.tabs(["Executivo","Agenda","Carteira","SLA","Relatórios","Auditoria","Mala direta","WhatsApp"])

    with tabs[0]:
        total = len(df_base) if df_base is not None else 0
        ag = int((df_base["Status_Contato"].astype(str)=="Agendado").sum()) if df_base is not None and "Status_Contato" in df_base else 0
        fin = int(df_base["Status_Contato"].astype(str).str.contains("Realizado|Concluído|Concluido",case=False,na=False).sum()) if df_base is not None and "Status_Contato" in df_base else 0
        a,b,c,d = st.columns(4)
        a.metric("Registros", total); b.metric("Agendados", ag); c.metric("Concluídos", fin)
        d.metric("Conversão", f"{(fin/total*100 if total else 0):.1f}%")
        st.info("Camada incorporada: dashboard executivo, SLA, carteira, agenda inteligente, auditoria, mala direta, WhatsApp e configuração operacional.")

    with tabs[1]:
        inicio = st.date_input("Semana", value=date.today()-timedelta(days=date.today().weekday()), key="enterprise_semana")
        if st.button("🧠 Gerar agenda inteligente", type="primary", use_container_width=True):
            st.session_state["enterprise_agenda"] = _agenda(df_lojas if df_lojas is not None else df_base, df_instrutores, inicio)
        agenda = st.session_state.get("enterprise_agenda", pd.DataFrame())
        if isinstance(agenda,pd.DataFrame) and not agenda.empty:
            st.dataframe(agenda,use_container_width=True,hide_index=True)
            if st.button("💾 Gravar agenda sugerida",use_container_width=True):
                c=_client()
                if c:
                    rows=[{"pv_abadi":str(r["PV"]),"data_agenda":pd.Timestamp(r["Data"]).date().isoformat(),"instrutor":str(r["Instrutor"]),"ordem":int(i),"distancia_km":float(r["Distância km"]),"status":"sugerido","usuario":_user()} for i,(_,r) in enumerate(agenda.iterrows())]
                    c.table("crm_agenda").upsert(rows,on_conflict="pv_abadi,data_agenda").execute()
                    audit("agenda",inicio.isoformat(),"criar_sugestao",depois={"registros":len(rows)})
                    st.success("Agenda gravada no Supabase.")
        else:
            st.info("Gere uma sugestão para montar a semana.")

    with tabs[2]:
        carteira=_carteira(df_base)
        if carteira.empty: st.info("Consultor/CF não disponível na base.")
        else:
            st.dataframe(carteira,use_container_width=True,hide_index=True)
            escolhido=st.selectbox("Consultor",carteira["Consultor"].tolist())
            if st.button("📞 Abrir fila sem treinamento"):
                st.session_state["enterprise_consultor_filtro"]=escolhido
                st.session_state["modulo_navegacao"]="📞 Call Center & Timeline WhatsApp"
                st.rerun()

    with tabs[3]:
        limite=int(config("sla_dias",180) or 180)
        novo=st.number_input("SLA em dias",1,720,limite,1)
        if novo!=limite and st.button("Salvar SLA"):
            save_config("sla_dias",int(novo)); st.rerun()
        q=_sla(df_base,limite)
        if not q.empty:
            a,b,c,d=st.columns(4)
            a.metric("Nunca treinou",int((q.SLA=="NUNCA TREINOU").sum()))
            b.metric("Estourado",int((q.SLA=="ESTOURADO").sum()))
            c.metric("A vencer",int((q.SLA=="A VENCER").sum()))
            d.metric("Dentro",int((q.SLA=="DENTRO").sum()))
            st.dataframe(q,use_container_width=True,hide_index=True)

    with tabs[4]:
        if df_base is None or df_base.empty:
            st.info("Sem dados.")
        else:
            x=df_base.copy()
            dc=next((c for c in ["Data_Ultimo_Treinamento","data_ultimo_treinamento","Data_Contato","data_do_contato"] if c in x.columns),None)
            if dc:
                x["_mes"]=pd.to_datetime(x[dc],errors="coerce").dt.to_period("M").astype(str)
                st.dataframe(x.groupby("_mes",dropna=True).size().reset_index(name="registros").tail(24),use_container_width=True,hide_index=True)
            st.download_button("⬇️ Exportar CSV",x.to_csv(index=False).encode("utf-8-sig"),"relatorio_crm_ampm.csv","text/csv",use_container_width=True)

    with tabs[5]:
        c=_client()
        if not c: st.info("Supabase indisponível.")
        else:
            try:
                data=c.table("crm_eventos").select("*").order("created_at",desc=True).limit(500).execute().data or []
                if data:
                    st.dataframe(pd.DataFrame(data)[[k for k in ["created_at","usuario","tipo","origem","pv_abadi"] if k in data[0]]],use_container_width=True,hide_index=True)
                    st.json(data[0].get("payload", {}))
                else: st.info("Nenhuma alteração auditada ainda.")
            except Exception as e: st.error(f"Falha na auditoria: {e}")

    with tabs[6]:
        st.markdown("### ✉️ Mala direta")
        st.caption("Editor/prévia. O disparo real pode ser ligado ao provedor de e-mail já configurado no ambiente, sem colocar credenciais no código.")
        assunto=st.text_input("Assunto", "Treinamento AmPm — {{loja}}")
        corpo=st.text_area("Mensagem","Olá {{nome}},\n\nPodemos entrar em contato para acertar os detalhes do treinamento da loja {{loja}}?\n\nAtenciosamente,\nCRM AmPm · IGT Group",height=180)
        if df_base is not None and not df_base.empty:
            r=df_base.iloc[0]
            msg=corpo.replace("{{nome}}",str(r.get("Nome_Contato",r.get("nome_contato","")))).replace("{{loja}}",str(r.get("Razão Social",r.get("razao_social",""))))
            st.markdown("#### Prévia"); st.info(f"**{assunto.replace('{{loja}}',str(r.get('Razão Social','')))}**\n\n{msg}")

    with tabs[7]:
        st.markdown("### 💬 WhatsApp")
        msg=st.text_area("Mensagem","Olá {{nome}}! Aqui é da IGT Group. Podemos alinhar os detalhes do treinamento da loja {{loja}}?",height=120)
        if df_base is not None and not df_base.empty:
            r=df_base.iloc[0]
            nome=str(r.get("Nome_Contato",r.get("nome_contato",""))); loja=str(r.get("Razão Social",r.get("razao_social",""))); tel=str(r.get("Telefone_Contato",r.get("telefone_contato","")) or "").replace(" ","").replace("-","")
            final=msg.replace("{{nome}}",nome).replace("{{loja}}",loja)
            if tel: st.link_button("Abrir WhatsApp",f"https://wa.me/{quote(tel)}?text={quote(final)}",use_container_width=True)
            st.code(final)
