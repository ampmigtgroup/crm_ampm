import os
import pandas as pd
from sqlalchemy import create_engine, text

NOME_EXCEL = 'Base_Unificada_AmPm.xlsx'
DATABASE_URL = "postgresql://postgres.nptazzfvwhhmotfrvgdj:Lssj.ampm%40%23@aws-0-sa-east-1.pooler.supabase.com:6543/postgres"

def popular_banco():
    if not os.path.exists(NOME_EXCEL):
        print(f"Coloque o arquivo {NOME_EXCEL} na pasta do projeto!")
        return

    engine = create_engine(DATABASE_URL)

    # 1. Migrar Lojas
    print("Migrando Lojas...")
    try:
        df_lojas = pd.read_excel(NOME_EXCEL, sheet_name='Rede_de_Lojas')
        mapeamento_lojas = {
            'PV Abadi': 'pv_abadi',
            'Razão Social': 'razao_social',
            'Status Loja': 'status_loja',
            'Municipio': 'municipio_uf',
            'Endereço': 'endereco',
            'CEP': 'cep',
            'Data Inauguracao': 'data_inauguracao'
        }
        cols_existentes = [c for c in mapeamento_lojas.keys() if c in df_lojas.columns]
        df_lojas = df_lojas[cols_existentes].rename(columns=mapeamento_lojas)
        
        # Usar IF_EXISTS='REPLACE' para recriar com os tipos corretos
        df_lojas.to_sql('tb_lojas', engine, if_exists='replace', index=False)
        
        # Adicionar Chave Primária via SQL
        with engine.connect() as conn:
            conn.execute(text("ALTER TABLE tb_lojas ADD PRIMARY KEY (pv_abadi);"))
            conn.commit()
        print("✓ Lojas migradas com sucesso!")
    except Exception as e:
        print(f"Aviso Lojas: {e}")

    # 2. Migrar Instrutores
    print("Migrando Instrutores...")
    try:
        df_instrutores = pd.read_excel(NOME_EXCEL, sheet_name='Instrutores')
        df_instrutores.columns = [c.lower().strip().replace(' ', '_') for c in df_instrutores.columns]
        df_instrutores.to_sql('tb_instrutores', engine, if_exists='replace', index=False)
        print("✓ Instrutores migrados com sucesso!")
    except Exception as e:
        print(f"Aviso Instrutores: {e}")

    # 3. Migrar Fila do Call Center
    print("Migrando Fila do Call Center...")
    try:
        df_fila = pd.read_excel(NOME_EXCEL, sheet_name='Fila_CallCenter')
        
        # Mapear e padronizar os nomes das colunas para minúsculas
        df_fila.columns = [c.lower().strip().replace(' ', '_') for c in df_fila.columns]
        
        # Se semana_sugerida for texto/data, garantir formato string
        if 'semana_sugerida' in df_fila.columns:
            df_fila['semana_sugerida'] = df_fila['semana_sugerida'].astype(str)
            
        df_fila.to_sql('tb_fila_call_center', engine, if_exists='replace', index=False)
        
        # Adicionar coluna de ID sequencial no PostgreSQL para o CRM
        with engine.connect() as conn:
            conn.execute(text("ALTER TABLE tb_fila_call_center ADD COLUMN IF NOT EXISTS id_atendimento SERIAL PRIMARY KEY;"))
            conn.commit()
        print("✓ Fila do Call Center migrada com sucesso!")
    except Exception as e:
        print(f"Erro na Fila: {e}")

    print("\n🎉 Todas as tabelas foram criadas e povoadas no Supabase!")

if __name__ == '__main__':
    popular_banco()
