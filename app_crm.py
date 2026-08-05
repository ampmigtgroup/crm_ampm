import streamlit as st
import pandas as pd

# Configuração da Página
st.set_page_config(page_title="Sistema de Gestão de Treinamentos - ampm", layout="wide")

# Styling CSS para identidade ampm
st.markdown("""
    <style>
    .stApp { background-color: #f8fafc; }
    .main-header { font-size: 24px; font-weight: bold; color: #0f172a; margin-bottom: 20px; }
    .ampm-badge { background-color: #f59e0b; color: #0f172a; padding: 4px 8px; border-radius: 4px; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header"><span class="ampm-badge">ampm</span> Gestão Integrada de Treinamentos</div>', unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# BASES DE DADOS REAIS & RESTAURADAS
# -----------------------------------------------------------------------------

# Base Completa de Clientes Restaurada
df_clientes = pd.DataFrame([
    {"Código": "CLI-101", "Razão Social": "Rede Rota Sul de Combustíveis Ltd.", "Postos": 14, "Região": "São Paulo - SP", "Módulo ampm": "Completo (Loja + Pista)", "Ficha": "EXP-2026-8841"},
    {"Código": "CLI-102", "Razão Social": "Grupo Nova Era Derivados de Petróleo", "Postos": 8, "Região": "Campinas / Jundiaí - SP", "Módulo ampm": "Gestão de Loja ampm", "Ficha": "EXP-2026-8842"},
    {"Código": "CLI-103", "Razão Social": "Rede Anhanguera de Postos de Serviços", "Postos": 22, "Região": "Capital & RMC - SP", "Módulo ampm": "Módulo Básico ampm", "Ficha": "EXP-2026-8843"},
    {"Código": "CLI-104", "Razão Social": "Comercial de Combustíveis Fluminense", "Postos": 11, "Região": "Rio de Janeiro - RJ", "Módulo ampm": "Completo (Loja + Pista)", "Ficha": "EXP-2026-8844"},
    {"Código": "CLI-105", "Razão Social": "Auto Posto Central de Belo Horizonte", "Postos": 5, "Região": "Belo Horizonte - MG", "Módulo ampm": "Gestão de Loja ampm", "Ficha": "EXP-2026-8845"},
    {"Código": "CLI-106", "Razão Social": "Rede Serrana de Derivados Petrópolis", "Postos": 7, "Região": "Serrana - RJ", "Módulo ampm": "Atendimento & Padrão ampm", "Ficha": "EXP-2026-8846"},
    {"Código": "CLI-107", "Razão Social": "Grupo Triângulo do Sol Combustíveis", "Postos": 18, "Região": "Uberlândia - MG", "Módulo ampm": "Completo (Loja + Pista)", "Ficha": "EXP-2026-8847"},
    {"Código": "CLI-108", "Razão Social": "Rede Sul Pistas de Serviços Rodoviários", "Postos": 30, "Região": "Curitiba & RMC - PR", "Módulo ampm": "Módulo Avançado ampm", "Ficha": "EXP-2026-8848"}
])

# Quadro de Capilaridade de Instrutores Completo (Restaurado)
df_instrutores = pd.DataFrame([
    {"Código": "INS-001", "Instrutor": "Carlos Eduardo Santos", "Região / Base": "São Paulo - Capital / ABC", "Especialidade": "Atendimento Pista & Lojas ampm", "Jornada": "Segunda a Sexta", "Status": "Em Treinamento"},
    {"Código": "INS-002", "Instrutor": "Mariana Silva Oliveira", "Região / Base": "São Paulo - Campinas / Interior", "Especialidade": "Gestão de Estoque & Caixa ampm", "Jornada": "Segunda a Sexta", "Status": "Disponível"},
    {"Código": "INS-003", "Instrutor": "Roberto Lima Junior", "Região / Base": "Rio de Janeiro - Niterói / Capital", "Especialidade": "Segurança Operacional & NR-20", "Jornada": "Segunda a Sexta", "Status": "Disponível"},
    {"Código": "INS-004", "Instrutor": "Amanda Costa Ramos", "Região / Base": "Minas Gerais - BH / Contagem", "Especialidade": "Treinamento Avançado ampm", "Jornada": "Segunda a Sexta", "Status": "Em Treinamento"},
    {"Código": "INS-005", "Instrutor": "Fernando Martins Prado", "Região / Base": "Paraná - Curitiba e RMC", "Especialidade": "Operações de Pista e Lojas ampm", "Jornada": "Segunda a Sexta", "Status": "Disponível"},
    {"Código": "INS-006", "Instrutor": "Juliana Barbosa Mendes", "Região / Base": "Rio Grande do Sul - Porto Alegre", "Especialidade": "Padrões de Qualidade & Atendimento", "Jornada": "Segunda a Sexta", "Status": "Disponível"},
    {"Código": "INS-007", "Instrutor": "Lucas Gabriel Albuquerque", "Região / Base": "Bahia - Salvador / Feira de Santana", "Especialidade": "Módulos de Gestão ampm", "Jornada": "Segunda a Sexta", "Status": "Disponível"},
    {"Código": "INS-008", "Instrutor": "Patricia Vasconcelos", "Região / Base": "Ceará - Fortaleza e Região MET", "Especialidade": "Capacitação de Gerentes ampm", "Jornada": "Segunda a Sexta", "Status": "Disponível"}
])

# -----------------------------------------------------------------------------
# ABA DE NAVEGAÇÃO
# -----------------------------------------------------------------------------
aba1, aba2, aba3, aba4, aba5, aba6 = st.tabs([
    "Pipeline ampm", 
    "Capilaridade Instrutores", 
    "Call Center / WhatsApp", 
    "Cruzamento PROCV & Rotas", 
    "Base de Clientes", 
    "Ficha de Exportação"
])

# 1. PIPELINE KANBAN AMPM
with aba1:
    st.subheader("Pipeline de Treinamento: ampm")
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.info("**1. Diagnóstico ampm**")
        st.caption("Posto Estrela do Sul\nRede Sul Pistas\n*Módulo Loja*")
    with col2:
        st.warning("**2. Agendamento**")
        st.caption("Posto Barra Funda\nInstrutora: Mariana Silva\n*Seg a Sex (10 a 14/Ago)*")
    with col3:
        st.success("**3. Em Treinamento**")
        st.caption("Posto Bandeirantes\nInstrutor: Carlos Eduardo\n*Em Andamento (Seg-Sex)*")
    with col4:
        st.error("**4. Avaliação & Homolog.**")
        st.caption("Posto Castelo Branco\nInstrutor: Roberto Lima\n*Nota: 9.8 / 10*")
    with col5:
        st.write("**5. Ficha Emitida**")
        st.caption("Posto Linha Verde\nFicha: EXP-2026-8841\n*Concluído*")

# 2. CAPILARIDADE DOS INSTRUTORES (COMPLETA RESTAURADA)
with aba2:
    st.subheader("Quadro Completo de Capilaridade dos Instrutores")
    st.dataframe(df_instrutores, use_container_width=True)

# 3. CALL CENTER & TIMELINE WHATSAPP INTEGRADOS
with aba3:
    st.subheader("Atendimento Unificado: Call Center + WhatsApp")
    
    modo = st.radio("Selecione a visualização dos canais de atendimento:", 
                    ["Ambos Integrados", "Apenas Voz (Call Center)", "Apenas WhatsApp"], horizontal=True)
    
    col_call, col_wp = st.columns(2)
    
    if modo in ["Ambos Integrados", "Apenas Voz (Call Center)"]:
        with col_call:
            st.markdown("### 📞 Call Center (Voz)")
            st.info("<strong>Chamada Ativa:</strong> (11) 98765-4321<br><strong>Cliente:</strong> Posto Marginal Tietê", unsafe_allow_html=True)
            st.text_area("Script / Anotações da Ligação", "Confirmar agendamento de treinamento ampm para a próxima segunda-feira.")
            if st.button("Finalizar e Registrar Chamada"):
                st.success("Chamada registrada com sucesso!")

    if modo in ["Ambos Integrados", "Apenas WhatsApp"]:
        with col_wp:
            st.markdown("### 💬 Timeline WhatsApp")
            st.chat_message("user").write("Bom dia! Confirmado o treinamento ampm para segunda-feira?")
            st.chat_message("assistant").write("Bom dia! Sim, o instrutor Carlos estará no posto na segunda-feira às 08h00.")
            
            mensagem = st.text_input("Enviar mensagem WhatsApp:")
            if st.button("Enviar"):
                if mensagem:
                    st.chat_message("assistant").write(mensagem)

# 4. CRUZAMENTO INTELIGENTE PROCV + ROTAS E REGRA FDS PAGO
with aba4:
    st.subheader("Cruzamento Inteligente (PROCV) & Otimização de Treinamento")
    
    st.warning("⚠️ **Regra Operacional de Jornada:** Os instrutores trabalham de **Segunda a Sexta-feira**. Caso haja necessidade emergencial no final de semana, o pagamento do Adicional FDS é ativado automaticamente.")
    
    posto_selecionado = st.selectbox("Selecione o Posto Alvo para Cruzamento:", [
        "Posto Bandeirantes (São Paulo - SP)",
        "Posto Barra Funda (São Paulo - SP)",
        "Posto Castelo Branco (Jundiaí - SP)"
    ])
    
    st.markdown("#### Postos Próximos Encontrados para Otimização de Rota:")
    
    if "Bandeirantes" in posto_selecionado:
        df_procv = pd.DataFrame([
            {"Posto Mapeado": "Posto Bandeirantes (ALVO)", "Distância": "0.0 km", "Rede": "Rede Rota Sul", "Instrutor": "Carlos Eduardo Santos", "Período": "Segunda a Sexta (10 a 14/Ago)", "Regime": "Regular (Seg-Sex)"},
            {"Posto Mapeado": "Posto Anhanguera Km 18", "Distância": "3.4 km", "Rede": "Rede Rota Sul", "Instrutor": "Carlos Eduardo Santos", "Período": "Segunda a Sexta (17 a 21/Ago)", "Regime": "Regular (Seg-Sex)"},
            {"Posto Mapeado": "Posto Castelo Branco Km 22", "Distância": "7.1 km", "Rede": "Grupo Nova Era", "Instrutor": "Carlos Eduardo Santos", "Período": "Sábado e Domingo (22 e 23/Ago)", "Regime": "Adicional FDS Pago"}
        ])
    else:
        df_procv = pd.DataFrame([
            {"Posto Mapeado": posto_selecionado, "Distância": "0.0 km", "Rede": "Rede Anhanguera", "Instrutor": "Mariana Silva Oliveira", "Período": "Segunda a Sexta (10 a 14/Ago)", "Regime": "Regular (Seg-Sex)"},
            {"Posto Mapeado": "Posto Auxiliar Próximo", "Distância": "4.2 km", "Rede": "Rede Anhanguera", "Instrutor": "Mariana Silva Oliveira", "Período": "Sábado (15/Ago)", "Regime": "Adicional FDS Pago"}
        ])
        
    st.dataframe(df_procv, use_container_width=True)

# 5. BASE DE CLIENTES COMPLETA RESTAURADA
with aba5:
    st.subheader("Base Completa de Clientes")
    st.dataframe(df_clientes, use_container_width=True)

# 6. FICHA DE EXPORTAÇÃO (PADRÃO ORIGINAL MANTIDO)
with aba6:
    st.subheader("Ficha de Exportação (Modelo Padrão Original)")
    
    st.markdown("""
    ---
    ### **FICHA DE EXPORTAÇÃO DE HOMOLOGAÇÃO DE TREINAMENTO AMPM**
    **Código do Documento:** EXP-2026-8841  
    **Autenticação:** HASH-AMP-99201-2026  
    
    * **Cliente:** Rede Rota Sul de Combustíveis Ltd.
    * **CNPJ:** 12.345.678/0001-90
    * **Unidade / Posto:** Posto Bandeirantes - Unidade 01
    * **Módulo Realizado:** Módulo Completo Loja ampm & Pista
    * **Instrutor Responsável:** Carlos Eduardo Santos
    * **Status da Homologação:** APROVADO (100%)
    ---
    """)
    st.button("📥 Baixar Ficha de Exportação (PDF)")
