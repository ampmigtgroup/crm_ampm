from motor.motor_asyncio import AsyncIOMotorClient
from .config import settings

client = AsyncIOMotorClient(settings.mongo_uri)
db = client[settings.mongo_db]

async def ensure_indexes():
    await db.users.create_index("username", unique=True)
    await db.users.create_index("email", unique=True, sparse=True)
    await db.sessions.create_index("token", unique=True)
    await db.lojas.create_index("cnpj", sparse=True)
    await db.lojas.create_index([("nome", "text"), ("cidade", "text"), ("uf", "text")])
    await db.contatos.create_index("loja_id")
    await db.instrutores.create_index("ativo")
    await db.oportunidades.create_index("status")
    await db.oportunidades.create_index("consultor")
    await db.orcamentos.create_index("created_at")
    await db.auditoria.create_index("created_at")
