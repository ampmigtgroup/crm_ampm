from datetime import datetime, timedelta, timezone
from typing import Any
from bson import ObjectId
from fastapi import FastAPI, HTTPException, Depends, Header, Query
from fastapi.middleware.cors import CORSMiddleware
from passlib.context import CryptContext
from jose import jwt, JWTError
from pydantic import BaseModel, Field
from .config import settings
from .db import db, ensure_indexes

app = FastAPI(title=settings.app_name, version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"] if settings.frontend_origin == "*" else [settings.frontend_origin],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

pwd = CryptContext(schemes=["bcrypt"], deprecated="auto")

class LoginIn(BaseModel):
    username: str
    password: str

class UserOut(BaseModel):
    username: str
    name: str
    email: str | None = None
    role: str = "user"

class LoginOut(BaseModel):
    token: str
    user: UserOut

class GenericDoc(BaseModel):
    data: dict[str, Any] = Field(default_factory=dict)

def oid(v: str):
    try:
        return ObjectId(v)
    except Exception:
        return v

def clean(doc):
    if not doc:
        return None
    doc = dict(doc)
    if "_id" in doc:
        doc["id"] = str(doc.pop("_id"))
    return doc

async def current_user(authorization: str | None = Header(default=None)):
    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(401, "Não autenticado")
    token = authorization.split(" ", 1)[1]
    try:
        payload = jwt.decode(token, settings.jwt_secret, algorithms=["HS256"])
        username = payload.get("sub")
    except JWTError:
        raise HTTPException(401, "Sessão inválida")
    user = await db.users.find_one({"username": username})
    if not user:
        raise HTTPException(401, "Usuário não encontrado")
    return clean(user)

@app.on_event("startup")
async def startup():
    await ensure_indexes()

@app.get("/api/health")
async def health():
    await db.command("ping")
    return {"status": "ok", "service": settings.app_name}

@app.post("/api/auth/login", response_model=LoginOut)
async def login(body: LoginIn):
    user = await db.users.find_one({"username": body.username.strip().lower()})
    if not user or not pwd.verify(body.password, user.get("password_hash", "")):
        raise HTTPException(401, "Usuário ou senha inválidos")
    exp = datetime.now(timezone.utc) + timedelta(minutes=settings.jwt_expire_minutes)
    token = jwt.encode({"sub": user["username"], "exp": exp}, settings.jwt_secret, algorithm="HS256")
    await db.sessions.insert_one({"token": token, "username": user["username"], "created_at": datetime.now(timezone.utc)})
    return {"token": token, "user": UserOut(**clean({k:v for k,v in user.items() if k != "password_hash"}))}

@app.get("/api/auth/me", response_model=UserOut)
async def me(user=Depends(current_user)):
    return UserOut(**{k:v for k,v in user.items() if k != "password_hash"})

@app.get("/api/dashboard")
async def dashboard(user=Depends(current_user)):
    collections = ["lojas", "contatos", "instrutores", "oportunidades", "orcamentos"]
    counts = {name: await db[name].count_documents({}) for name in collections}
    pipeline = {}
    for status in ["novo", "contato", "proposta", "a_contratar", "aguardando_pagamento", "treinamento", "concluido", "recusado"]:
        pipeline[status] = await db.oportunidades.count_documents({"status": status})
    return {"counts": counts, "pipeline": pipeline, "generated_at": datetime.now(timezone.utc)}

MODULES = {
    "lojas": "lojas",
    "contatos": "contatos",
    "instrutores": "instrutores",
    "agenda": "agenda",
    "pipeline": "oportunidades",
    "oportunidades": "oportunidades",
    "orcamentos": "orcamentos",
    "campanhas": "campanhas",
    "auditoria": "auditoria",
    "relatorios": "relatorios",
    "usuarios": "users",
    "sla": "sla",
    "carteira": "carteira",
    "resumo-semanal": "resumo_semanal",
}

@app.get("/api/{module}")
async def list_module(module: str, skip: int = 0, limit: int = Query(100, le=500), user=Depends(current_user)):
    collection = MODULES.get(module)
    if not collection:
        raise HTTPException(404, "Módulo não encontrado")
    docs = await db[collection].find({}).sort("_id", -1).skip(skip).limit(limit).to_list(limit)
    return [clean(x) for x in docs]

@app.post("/api/{module}")
async def create_module(module: str, body: GenericDoc, user=Depends(current_user)):
    collection = MODULES.get(module)
    if not collection:
        raise HTTPException(404, "Módulo não encontrado")
    data = dict(body.data)
    data["created_at"] = datetime.now(timezone.utc)
    data["created_by"] = user["username"]
    result = await db[collection].insert_one(data)
    await db.auditoria.insert_one({"acao":"criar","modulo":module,"registro_id":str(result.inserted_id),"usuario":user["username"],"created_at":datetime.now(timezone.utc)})
    return clean(await db[collection].find_one({"_id": result.inserted_id}))

@app.patch("/api/{module}/{item_id}")
async def update_module(module: str, item_id: str, body: GenericDoc, user=Depends(current_user)):
    collection = MODULES.get(module)
    if not collection:
        raise HTTPException(404, "Módulo não encontrado")
    data = dict(body.data)
    data["updated_at"] = datetime.now(timezone.utc)
    result = await db[collection].update_one({"_id": oid(item_id)}, {"$set": data})
    if not result.matched_count:
        raise HTTPException(404, "Registro não encontrado")
    await db.auditoria.insert_one({"acao":"editar","modulo":module,"registro_id":item_id,"usuario":user["username"],"created_at":datetime.now(timezone.utc)})
    return clean(await db[collection].find_one({"_id": oid(item_id)}))

@app.delete("/api/{module}/{item_id}")
async def delete_module(module: str, item_id: str, user=Depends(current_user)):
    collection = MODULES.get(module)
    if not collection:
        raise HTTPException(404, "Módulo não encontrado")
    result = await db[collection].delete_one({"_id": oid(item_id)})
    if not result.deleted_count:
        raise HTTPException(404, "Registro não encontrado")
    await db.auditoria.insert_one({"acao":"excluir","modulo":module,"registro_id":item_id,"usuario":user["username"],"created_at":datetime.now(timezone.utc)})
    return {"ok": True}
