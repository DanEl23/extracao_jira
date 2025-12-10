import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from base_dados.calculos import (
    calcular_metas_atuais,
    atualizar_historico_com_ano_atual,
    calcular_metas_por_superintendencia,
    atualizar_historico_macrodesafio_com_ano_atual,
    calcular_metas_atuais_por_macrodesafio
)
from base_dados.mapeamentos import ORDEM_SUPERINTENDENCIAS

# Configuração da Página
st.set_page_config(
    page_title="Relatório Técnico de Metas Estratégicas - TJMG",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- CSS Personalizado para aspecto formal de documento técnico ---
st.markdown("""
<style>
    /* Paleta institucional */
    :root {
        --cor-principal: #E36C0A;
        --cor-texto: #2c3e50;
        --cor-cinza: #7f8c8d;
    }
    
    /* Remover padding padrão para parecer documento */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1200px;
    }
    
    /* Títulos formatados como documento formal */
    h1 {
        color: var(--cor-texto);
        font-family: 'Times New Roman', serif;
        font-size: 26px;
        font-weight: bold;
        text-align: center;
        margin-bottom: 0.5rem;
        border-bottom: 2px solid var(--cor-principal);
        padding-bottom: 10px;
    }
    
    h2 {
        color: var(--cor-texto);
        font-family: 'Times New Roman', serif;
        font-size: 18px;
        font-weight: bold;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    
    h3 {
        color: var(--cor-texto);
        font-family: 'Times New Roman', serif;
        font-size: 14px;
        font-weight: bold;
        margin-top: 1.5rem;
        margin-bottom: 0.8rem;
    }
    
    /* Métricas com destaque institucional */
    [data-testid="stMetricValue"] {
        font-size: 28px;
        color: var(--cor-principal);
        font-weight: bold;
    }
    
    [data-testid="stMetricLabel"] {
        color: var(--cor-texto);
        font-size: 14px;
        font-weight: 600;
    }
    
    /* Tabelas formatadas profissionalmente */
    .dataframe {
        font-size: 11px;
        font-family: 'Times New Roman', serif;
    }
    
    /* Texto do corpo como documento formal */
    p, div, span {
        font-family: 'Times New Roman', serif;
        font-size: 12px;
        color: var(--cor-texto);
        text-align: justify;
        line-height: 1.6;
    }
    
    /* Destaque para informações importantes */
    .stAlert {
        font-family: 'Times New Roman', serif;
        border-left: 4px solid var(--cor-principal);
    }
    
    /* Rodapé institucional */
    footer {
        font-family: 'Times New Roman', serif;
        font-size: 10px;
        color: var(--cor-cinza);
        text-align: center;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #ddd;
    }
    
    /* Expander como seção do documento */
    .streamlit-expanderHeader {
        font-family: 'Times New Roman', serif;
        font-weight: bold;
        color: var(--cor-texto);
        font-size: 13px;
    }
</style>
""", unsafe_allow_html=True)

# --- Cabeçalho Institucional ---
st.markdown("""
<div style='text-align: center; margin-bottom: 2rem;'>
    <h1 style='color: #2c3e50; margin-bottom: 0.2rem;'>TRIBUNAL DE JUSTIÇA DO ESTADO DE MINAS GERAIS</h1>
    <p style='font-size: 14px; color: #7f8c8d; margin-top: 0;'>Assessoria Técnica e Jurídica de Planejamento, Gestão Estratégica e Qualidade - ASPLAG</p>
    <p style='font-size: 14px; color: #7f8c8d; margin-top: 0;'>Diretoria Executiva de Planejamento, Gestão Estratégica e Qualidade - DEPLAG</p>
</div>
""", unsafe_allow_html=True)

st.title("📊 RELATÓRIO TÉCNICO DE METAS ESTRATÉGICAS")
st.markdown(f"""
<div style='text-align: center; margin-bottom: 2rem;'>
    <p style='font-size: 13px; font-weight: bold;'>Relatório ao Comitê de Governança e Gestão Estratégica</p>
    <p style='font-size: 12px; color: #7f8c8d;'>Data de Referência: {datetime.now().strftime('%d de %B de %Y')}</p>
</div>
""", unsafe_allow_html=True)

st.divider()

# --- Carregamento de Dados (Cache para performance) ---
@st.cache_data
def carregar_dados_gerais():
    return calcular_metas_atuais()

@st.cache_data
def carregar_historico():
    return atualizar_historico_com_ano_atual()

@st.cache_data
def carregar_dados_superintendencia():
    return calcular_metas_por_superintendencia()

@st.cache_data
def carregar_dados_macrodesafio():
    return atualizar_historico_macrodesafio_com_ano_atual()

@st.cache_data
def carregar_contagem_macrodesafio_atual():
    return calcular_metas_atuais_por_macrodesafio()

try:
    dados_atuais = carregar_dados_gerais()
    dados_hist = carregar_historico()
    dados_super = carregar_dados_superintendencia()
    dados_macro = carregar_dados_macrodesafio()
    contagem_macro_atual = carregar_contagem_macrodesafio_atual()
except Exception as e:
    st.error(f"❌ Erro ao carregar dados. Certifique-se de ter executado a extração primeiro.")
    st.error(f"Detalhe técnico: {e}")
    st.stop()

# Variáveis do relatório
ano_atual = dados_atuais['ano']
total_metas = dados_atuais['total']
total_cnj = dados_atuais['nacionais']
total_tjmg = dados_atuais['institucionais']
total_superintendencias = len(dados_super)
total_macrodesafios = len(contagem_macro_atual)

# Variável para filtros (usando todas as superintendências por padrão)
super_filtro = ORDEM_SUPERINTENDENCIAS

# --- TAB 1: VISÃO GERAL (Substitui o Sumário Executivo) ---
tab1, tab2, tab3 = st.tabs(["📈 Visão Geral", "🏛️ Por Superintendência", "⏳ Evolução Histórica"])

with tab1:
    st.header(f"Panorama Geral - {ano_atual}")
    
    # KPIs (Métricas Principais)
    col1, col2, col3, col4 = st.columns(4)
    
    variacao = dados_hist['variacao'] # Pega a variação calculada no seu script
    
    with col1:
        st.metric("Total de Metas", dados_atuais['total'], delta=f"{variacao['total']}% (vs 2021)")
    with col2:
        st.metric("Metas Nacionais (CNJ)", dados_atuais['nacionais'], delta=f"{variacao['nacionais']}%")
    with col3:
        st.metric("Metas Institucionais", dados_atuais['institucionais'], delta=f"{variacao['institucionais']}%")
    with col4:
        st.metric("Percentual CNJ", f"{(dados_atuais['nacionais']/dados_atuais['total'])*100:.1f}%")

    # Gráfico de Macrodesafios (Substitui a tabela de macrodesafios)
    st.subheader("Distribuição por Macrodesafio")
    
    # Usar contagem atual de macrodesafios
    df_macro = pd.DataFrame(list(contagem_macro_atual.items()), columns=['Macrodesafio', 'Qtd'])
    # Limpeza do nome longo do macrodesafio para o gráfico
    df_macro['Macro_Curto'] = df_macro['Macrodesafio'].apply(lambda x: x.split('-')[0].strip() if '-' in x else x[:40])
    
    fig_macro = px.bar(
        df_macro, 
        x='Qtd', 
        y='Macro_Curto', 
        orientation='h',
        text='Qtd',
        title="Quantidade de Metas por Macrodesafio",
        color_discrete_sequence=['#E36C0A'] # Sua cor laranja
    )
    st.plotly_chart(fig_macro, use_container_width=True)

# --- TAB 2: POR SUPERINTENDÊNCIA (Substitui as tabelas detalhadas) ---
with tab2:
    st.header("Detalhamento por Área")
    
    for super_nome in super_filtro:
        if super_nome in dados_super:
            info = dados_super[super_nome]
            
            with st.expander(f"{super_nome} (Total: {info['total']})", expanded=True):
                c1, c2 = st.columns([1, 3])
                
                with c1:
                    st.info(f"""
                    **Resumo:**
                    - 🏛️ Institucionais: {info['institucionais']}
                    - ⚖️ Nacionais: {info['nacionais']}
                    """)
                
                with c2:
                    # Transforma a lista de metas em um DataFrame para visualização melhor
                    df_metas = pd.DataFrame(info['metas'], columns=['Código da Meta'])
                    # Adiciona uma coluna fictícia de status (já que não temos isso no extraído, mas mostra o potencial)
                    st.dataframe(df_metas, use_container_width=True, hide_index=True)

# --- TAB 3: HISTÓRICO (Substitui a tabela de evolução 2021-2024) ---
with tab3:
    st.header("Evolução das Metas Aprovadas (2021-2024)")
    
    hist_dict = dados_hist['historico_recente']
    
    # Converter dicionário de dicionários para DataFrame
    df_hist = pd.DataFrame.from_dict(hist_dict, orient='index').reset_index()
    df_hist.rename(columns={'index': 'Ano'}, inplace=True)
    
    # Gráfico de Linha Interativo
    fig_evolucao = px.line(
        df_hist, 
        x='Ano', 
        y=['nacionais', 'institucionais', 'total'],
        markers=True,
        title="Evolução do Número de Metas",
        labels={'value': 'Quantidade', 'variable': 'Tipo de Meta'}
    )
    st.plotly_chart(fig_evolucao, use_container_width=True)
    
    st.markdown("### Dados Tabulares")
    st.dataframe(df_hist, hide_index=True)

# Rodapé
st.sidebar.markdown("---")
st.sidebar.caption(f"Gerado via Extrator Unificado v2.0")