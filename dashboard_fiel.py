import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# Importando suas funções de negócio existentes
from base_dados.calculos import (
    calcular_metas_atuais,
    atualizar_historico_com_ano_atual,
    calcular_metas_atuais_por_macrodesafio,
    calcular_cumprimento_metas_cnj,
    calcular_metas_por_superintendencia
)
from resultado_monitoramento import calcular_resultado_monitoramento
from base_dados.mapeamentos import ORDEM_SUPERINTENDENCIAS

# --- Configuração da Página ---
st.set_page_config(
    page_title="Relatório de Metas Estratégicas - TJMG",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Estilos CSS (Tema Escuro) ---
st.markdown("""
<style>
    /* Fundo escuro global */
    .stApp {
        background-color: #1e1e1e;
        color: #ffffff;
    }
    
    /* Sidebar escura */
    [data-testid="stSidebar"] {
        background-color: #252526;
    }
    
    [data-testid="stSidebar"] * {
        color: #ffffff !important;
    }
    
    /* Títulos e textos brancos */
    h1, h2, h3, h4, h5, h6, p, span, div, label {
        color: #ffffff !important;
    }
    
    /* Métricas com fundo escuro e texto branco */
    [data-testid="metric-container"] {
        background-color: #2d2d2d;
        border-left: 5px solid #E36C0A;
        padding: 15px;
        border-radius: 5px;
        box-shadow: 2px 2px 8px rgba(0,0,0,0.4);
    }
    
    .stMetricValue {
        font-size: 28px !important;
        color: #E36C0A !important;
        font-weight: bold !important;
    }
    
    .stMetricLabel {
        color: #ffffff !important;
        font-size: 14px !important;
    }
    
    .stMetricDelta {
        color: #a0a0a0 !important;
    }
    
    /* Dataframes com fundo escuro */
    .dataframe {
        background-color: #2d2d2d !important;
        color: #ffffff !important;
    }
    
    .dataframe th {
        background-color: #3d3d3d !important;
        color: #E36C0A !important;
        font-weight: bold !important;
    }
    
    .dataframe td {
        background-color: #2d2d2d !important;
        color: #ffffff !important;
    }
    
    /* Expanders escuros */
    .streamlit-expanderHeader {
        background-color: #2d2d2d !important;
        color: #ffffff !important;
        border: 1px solid #3d3d3d !important;
    }
    
    .streamlit-expanderContent {
        background-color: #252526 !important;
        border: 1px solid #3d3d3d !important;
    }
    
    /* Dividers */
    hr {
        border-color: #3d3d3d !important;
    }
    
    /* Inputs e selects escuros */
    input, select, textarea {
        background-color: #2d2d2d !important;
        color: #ffffff !important;
        border: 1px solid #3d3d3d !important;
    }
    
    /* Botões */
    .stButton button {
        background-color: #E36C0A !important;
        color: #ffffff !important;
        border: none !important;
    }
    
    .stButton button:hover {
        background-color: #c55a08 !important;
    }
    
    /* Markdown tables escuras */
    table {
        color: #ffffff !important;
        background-color: #2d2d2d !important;
    }
    
    table th {
        background-color: #3d3d3d !important;
        color: #E36C0A !important;
    }
    
    table td {
        border-color: #3d3d3d !important;
    }
    
    /* Gráficos plotly com fundo escuro */
    .js-plotly-plot .plotly {
        background-color: #2d2d2d !important;
    }
    
    /* Alertas e mensagens */
    .stAlert {
        background-color: #2d2d2d !important;
        color: #ffffff !important;
        border-left: 4px solid #E36C0A !important;
    }
</style>
""", unsafe_allow_html=True)

# --- Carregamento de Dados (Cache) ---
@st.cache_data
def load_data():
    # 1. Dados Gerais e Histórico
    historico = atualizar_historico_com_ano_atual()
    
    # 2. Macrodesafios
    macros = calcular_metas_atuais_por_macrodesafio()
    
    # 3. Resultado do Monitoramento (Faixas de cumprimento)
    monitoramento = calcular_resultado_monitoramento()
    
    # 4. Metas Nacionais (Percentuais)
    nacionais = calcular_cumprimento_metas_cnj()
    
    return historico, macros, monitoramento, nacionais

try:
    dados_hist, dados_macro, dados_monit, dados_nac = load_data()
except Exception as e:
    st.error(f"Erro ao carregar dados. Verifique se os arquivos Excel existem na pasta 'exports'. Detalhe: {e}")
    st.stop()

# --- Cabeçalho do Relatório ---
col_logo, col_titulo = st.columns([1, 5])
with col_titulo:
    st.title("Monitoramento de Metas Estratégicas")
    st.markdown("**Relatório Técnico ao Comitê de Governança e Gestão Estratégica**")

st.divider()

# --- Filtros Globais (Sidebar) ---
st.sidebar.header("🎛️ Filtros do Relatório")
ano_ref = st.sidebar.selectbox("Ano de Referência", [2025, 2024], index=0)

# ==============================================================================
# PAINEL 1: QUANTIDADE E HISTÓRICO (O "Sumário Executivo")
# ==============================================================================
st.header("1. Visão Geral e Histórico")

# Métricas Principais (Topo)
col1, col2, col3, col4 = st.columns(4)
metas_atuais = dados_hist['historico'][dados_hist['anos_recentes'][-1]] # Pega último ano
variacao = dados_hist['variacao']

col1.metric("Total de Metas", metas_atuais['total'], f"{variacao['total']}% (Var. Histórica)")
col2.metric("Metas Nacionais", metas_atuais['nacionais'], f"{variacao['nacionais']}%")
col3.metric("Metas Institucionais", metas_atuais['institucionais'], f"{variacao['institucionais']}%")
col4.metric("Índice de Monitoramento", f"{(dados_monit['total'] - dados_monit['sem_apuracao']) / dados_monit['total'] * 100:.1f}%", help="Metas com apuração / Total")

# Gráfico de Evolução Histórica
st.subheader("Evolução Histórica (2021-2025)")
df_hist = pd.DataFrame.from_dict(dados_hist['historico_recente'], orient='index').reset_index()
df_hist.rename(columns={'index': 'Ano'}, inplace=True)

fig_hist = px.line(df_hist, x='Ano', y=['total', 'institucionais', 'nacionais'], 
                   markers=True, title='Crescimento do Portfólio de Metas',
                   color_discrete_map={'total': '#E36C0A', 'institucionais': '#ffffff', 'nacionais': '#a0a0a0'})
fig_hist.update_layout(
    xaxis_type='category',
    plot_bgcolor='#2d2d2d',
    paper_bgcolor='#2d2d2d',
    font=dict(color='#ffffff'),
    title_font=dict(color='#ffffff'),
    xaxis=dict(gridcolor='#3d3d3d', color='#ffffff'),
    yaxis=dict(gridcolor='#3d3d3d', color='#ffffff'),
    legend=dict(font=dict(color='#ffffff'))
)
st.plotly_chart(fig_hist, use_container_width=True)

# ==============================================================================
# PAINEL 2: RESULTADO DO MONITORAMENTO (O Gráfico de Pizza/Donut)
# ==============================================================================
st.header("2. Resultado do Monitoramento")
st.markdown("Distribuição das metas de acordo com a faixa de cumprimento (Nacionais + Institucionais).")

c1, c2 = st.columns([2, 1])

with c1:
    # Preparar dados para o gráfico
    labels = ['Maior ou igual a 100%', 'Entre 70% e 100%', 'Abaixo de 70%', 'Sem Apuração']
    values = [dados_monit['maior_100'], dados_monit['entre_70_100'], dados_monit['abaixo_70'], dados_monit['sem_apuracao']]
    colors = ['#2ecc71', '#f1c40f', '#e74c3c', '#95a5a6'] # Verde, Amarelo, Vermelho, Cinza
    
    fig_pie = go.Figure(data=[go.Pie(labels=labels, values=values, hole=.4, marker_colors=colors)])
    fig_pie.update_layout(
        margin=dict(t=0, b=0, l=0, r=0),
        plot_bgcolor='#2d2d2d',
        paper_bgcolor='#2d2d2d',
        font=dict(color='#ffffff'),
        showlegend=True,
        legend=dict(font=dict(color='#ffffff'))
    )
    st.plotly_chart(fig_pie, use_container_width=True)

with c2:
    # Tabela resumida ao lado
    st.write("")
    st.write("")
    st.markdown(f"""
    | Faixa de Cumprimento | Qtd | % |
    | :--- | :---: | :---: |
    | **≥ 100%** | {dados_monit['maior_100']} | {dados_monit['maior_100']/dados_monit['total']:.1%} |
    | **70% - 100%** | {dados_monit['entre_70_100']} | {dados_monit['entre_70_100']/dados_monit['total']:.1%} |
    | **< 70%** | {dados_monit['abaixo_70']} | {dados_monit['abaixo_70']/dados_monit['total']:.1%} |
    | **Sem Apuração** | {dados_monit['sem_apuracao']} | {dados_monit['sem_apuracao']/dados_monit['total']:.1%} |
    | **TOTAL** | **{dados_monit['total']}** | **100%** |
    """)

# ==============================================================================
# PAINEL 3: MACRODESAFIOS (Gráfico de Barras Horizontal)
# ==============================================================================
st.header("3. Distribuição por Macrodesafio")

df_macro = pd.DataFrame(list(dados_macro.items()), columns=['Macrodesafio', 'Quantidade'])
df_macro.sort_values('Quantidade', ascending=True, inplace=True) # Ordenar para o gráfico

fig_bar = px.bar(df_macro, x='Quantidade', y='Macrodesafio', orientation='h', text='Quantidade',
                 title="Quantidade de Metas Estratégicas por Macrodesafio",
                 color_discrete_sequence=['#E36C0A'])
fig_bar.update_traces(textposition='outside', textfont=dict(color='#ffffff'))
fig_bar.update_layout(
    height=600,
    plot_bgcolor='#2d2d2d',
    paper_bgcolor='#2d2d2d',
    font=dict(color='#ffffff'),
    title_font=dict(color='#ffffff'),
    xaxis=dict(gridcolor='#3d3d3d', color='#ffffff'),
    yaxis=dict(gridcolor='#3d3d3d', color='#ffffff')
)
st.plotly_chart(fig_bar, use_container_width=True)

# ==============================================================================
# PAINEL 4: CUMPRIMENTO DAS METAS NACIONAIS (Barra de Progresso)
# ==============================================================================
st.header("4. Cumprimento das Metas Nacionais (CNJ)")

# Transformar dicionário em DataFrame
data_nacionais = []
for meta, dados in dados_nac.items():
    data_nacionais.append({
        "Meta": meta,
        "Cumprimento": dados['cumprimento'] / 100, # Streamlit espera 0.0 a 1.0 para barra
        "Valor (%)": f"{dados['cumprimento']:.2f}%"
    })

df_nacionais = pd.DataFrame(data_nacionais).sort_values("Meta")

# Usar Dataframe com Column Config para mostrar barra de progresso visual
st.dataframe(
    df_nacionais,
    column_config={
        "Meta": st.column_config.TextColumn("Meta Nacional", width="medium"),
        "Cumprimento": st.column_config.ProgressColumn(
            "Progresso",
            help="Percentual de cumprimento da meta",
            format="%.2f%%",
            min_value=0,
            max_value=1,
        ),
        "Valor (%)": st.column_config.TextColumn("Valor Real")
    },
    hide_index=True,
    use_container_width=True
)

# ==============================================================================
# PAINEL 5: DETALHAMENTO COMPLETO DE METAS (Com Expanders)
# ==============================================================================
st.header("5. Detalhamento Completo de Metas")

# Carregar o Excel bruto para permitir filtragem detalhada
try:
    df_bruto = pd.read_excel('exports/teste_integração.xlsx')
    
    # Filtros na Interface
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        super_selecionada = st.selectbox("Filtrar por Área/Superintendência:", 
                                       ["Todas"] + ORDEM_SUPERINTENDENCIAS)
    with col_f2:
        busca_texto = st.text_input("Buscar por palavra-chave (ex: 'corregedoria', 'índice'):")

    # Aplicar Filtros
    df_filtrado = df_bruto.copy()
    
    # 1. Filtro de Superintendência (Lógica baseada no seu mapeamento)
    if super_selecionada != "Todas":
        # Precisamos reconstruir a lógica de qual meta pertence a qual super
        from base_dados.mapeamentos import META_SUPERINTENDENCIA
        
        def pertence_super(meta_key):
            # Tenta extrair código "TJMG 123"
            import re
            match = re.search(r'(TJMG\s+\d+|CNJ\s+\d+)', str(meta_key))
            if match:
                codigo = match.group(1)
                return META_SUPERINTENDENCIA.get(codigo) == super_selecionada
            return False
            
        df_filtrado = df_filtrado[df_filtrado['MetaKey'].apply(pertence_super)]

    # 2. Filtro de Texto
    if busca_texto:
        df_filtrado = df_filtrado[
            df_filtrado.astype(str).apply(lambda x: x.str.contains(busca_texto, case=False)).any(axis=1)
        ]

    st.info(f"📊 **{len(df_filtrado)} metas encontradas.** Clique em cada meta para ver detalhes completos.")
    
    # ==============================================================================
    # EXIBIR METAS COM EXPANDERS (Detalhamento completo)
    # ==============================================================================
    
    if len(df_filtrado) > 0:
        # Ordenar por MetaKey para melhor visualização
        df_filtrado = df_filtrado.sort_values('MetaKey')
        
        for idx, row in df_filtrado.iterrows():
            # Calcular percentual de cumprimento
            try:
                valor_meta = float(row.get('Valor da Meta', 0)) if pd.notna(row.get('Valor da Meta')) else 0
                valor_apurado = float(row.get('Valor Apurado', 0)) if pd.notna(row.get('Valor Apurado')) else 0
                
                if valor_meta > 0:
                    percentual = (valor_apurado / valor_meta) * 100
                else:
                    percentual = 0
                
                # Determinar cor da faixa
                if percentual >= 100:
                    cor_faixa = "🟢"
                    faixa_desc = "≥ 100% (Meta Cumprida)"
                elif percentual >= 70:
                    cor_faixa = "🟡"
                    faixa_desc = "70-100% (Em Cumprimento)"
                elif percentual > 0:
                    cor_faixa = "🔴"
                    faixa_desc = "< 70% (Abaixo do Esperado)"
                else:
                    cor_faixa = "⚪"
                    faixa_desc = "Sem Apuração"
                    percentual = 0
                
            except Exception as e:
                # Garantir que variáveis existam mesmo em caso de erro
                valor_meta = 0
                valor_apurado = 0
                percentual = 0
                cor_faixa = "⚪"
                faixa_desc = "Sem Apuração"
            
            # Título do expander com código, macrodesafio e status
            meta_key = row.get('MetaKey', 'Sem código')
            macrodesafio = row.get('Macrodesafio', 'Não informado')
            macro_curto = macrodesafio.split('-')[0].strip() if '-' in str(macrodesafio) else str(macrodesafio)[:30]
            
            titulo_expander = f"{cor_faixa} **{meta_key}** | {macro_curto} | {percentual:.1f}%"
            
            with st.expander(titulo_expander, expanded=False):
                # Container com informações completas da meta
                
                # Linha 1: Informações Básicas
                st.markdown("### 📋 Informações Básicas")
                col_a, col_b = st.columns(2)
                
                with col_a:
                    st.markdown(f"""
                    **Código da Meta:** {meta_key}  
                    **Ano:** {row.get('Ano da Meta', 'N/A')}  
                    **Macrodesafio:** {macrodesafio}
                    """)
                
                with col_b:
                    st.markdown(f"""
                    **Unidade Gestora:** {row.get('Unidade Gestora', 'Não informado')}  
                    **Polaridade:** {row.get('Polaridade', 'N/A')}  
                    **Situação:** {cor_faixa} {faixa_desc}
                    """)
                
                st.divider()
                
                # Linha 2: Objetivo/Resumo da Meta
                st.markdown("### 🎯 Objetivo Estratégico")
                resumo = row.get('Resumo', 'Não informado')
                if pd.notna(resumo) and resumo != 'Não informado':
                    # Processar formatação do resumo
                    import re
                    texto_resumo = str(resumo)
                    
                    # Aplicar negrito e itálico
                    texto_resumo = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', texto_resumo)
                    texto_resumo = re.sub(r'(?<!\*)\*(?!\*)([^*]+)\*(?!\*)', r'<em>\1</em>', texto_resumo)
                    
                    # Processar listas
                    linhas = texto_resumo.split('\n')
                    linhas_formatadas = []
                    em_lista = False
                    
                    for linha in linhas:
                        linha_stripped = linha.strip()
                        
                        if re.match(r'^[•\-\*]\s+', linha_stripped):
                            if not em_lista:
                                linhas_formatadas.append('<ul style="list-style-type: disc; padding-left: 25px; margin: 8px 0;">')
                                em_lista = True
                            texto_item = re.sub(r'^[•\-\*]\s+', '', linha_stripped)
                            linhas_formatadas.append(f'<li style="margin: 4px 0; color: #ffffff;">{texto_item}</li>')
                        
                        elif re.match(r'^\d+\.\s+', linha_stripped):
                            if em_lista and linhas_formatadas[-1].startswith('<ul'):
                                linhas_formatadas.append('</ul>')
                                em_lista = False
                            if not em_lista:
                                linhas_formatadas.append('<ol style="padding-left: 25px; margin: 8px 0;">')
                                em_lista = True
                            texto_item = re.sub(r'^\d+\.\s+', '', linha_stripped)
                            linhas_formatadas.append(f'<li style="margin: 4px 0; color: #ffffff;">{texto_item}</li>')
                        
                        else:
                            if em_lista:
                                if linhas_formatadas[-1].startswith('<ol'):
                                    linhas_formatadas.append('</ol>')
                                else:
                                    linhas_formatadas.append('</ul>')
                                em_lista = False
                            
                            if linha_stripped:
                                linhas_formatadas.append(f'<p style="margin: 8px 0;">{linha_stripped}</p>')
                    
                    if em_lista:
                        if linhas_formatadas and linhas_formatadas[-1].startswith('<li'):
                            if '<ol' in '\n'.join(linhas_formatadas):
                                linhas_formatadas.append('</ol>')
                            else:
                                linhas_formatadas.append('</ul>')
                    
                    texto_resumo = '\n'.join(linhas_formatadas)
                    
                    st.markdown(f"""
                    <div style='
                        text-align: justify; 
                        padding: 12px; 
                        background-color: #2d2d2d; 
                        border-left: 3px solid #E36C0A; 
                        border-radius: 3px;
                        font-size: 13px;
                        color: #ffffff;
                        line-height: 1.7;
                    '>
                        {texto_resumo}
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.info("Objetivo não disponível")
                
                st.divider()
                
                # Linha 3: Indicador
                st.markdown("### 📊 Indicador")
                indicador = row.get('Indicador', 'Não informado')
                if pd.notna(indicador) and indicador != 'Não informado':
                    st.markdown(f"**{indicador}**")
                else:
                    st.info("Indicador não disponível")
                
                # Linha 4: Valores e Performance
                st.markdown("### 📈 Valores e Desempenho")
                col_m1, col_m2, col_m3 = st.columns(3)
                
                with col_m1:
                    st.metric("Valor da Meta", f"{valor_meta:.2f}" if valor_meta > 0 else "N/A")
                
                with col_m2:
                    st.metric("Valor Apurado", f"{valor_apurado:.2f}" if valor_apurado > 0 else "N/A")
                
                with col_m3:
                    st.metric("Percentual de Cumprimento", f"{percentual:.1f}%", 
                             delta=f"{percentual - 100:.1f}% vs meta" if percentual > 0 else None)
                
                # Barra de progresso visual
                if percentual > 0:
                    st.progress(min(percentual / 100, 1.0))
                
                st.divider()
                
                # Linha 5: Iniciativas
                st.markdown("### 🚀 Iniciativas Relacionadas")
                iniciativa = row.get('Iniciativa', 'Não informado')
                if pd.notna(iniciativa) and iniciativa != 'Não informado':
                    # Processar formatação da iniciativa
                    import re
                    texto_iniciativa = str(iniciativa)
                    
                    # Aplicar negrito e itálico
                    texto_iniciativa = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', texto_iniciativa)
                    texto_iniciativa = re.sub(r'(?<!\*)\*(?!\*)([^*]+)\*(?!\*)', r'<em>\1</em>', texto_iniciativa)
                    
                    # Processar listas
                    linhas = texto_iniciativa.split('\n')
                    linhas_formatadas = []
                    em_lista = False
                    
                    for linha in linhas:
                        linha_stripped = linha.strip()
                        
                        if re.match(r'^[•\-\*]\s+', linha_stripped):
                            if not em_lista:
                                linhas_formatadas.append('<ul style="list-style-type: disc; padding-left: 25px; margin: 8px 0;">')
                                em_lista = True
                            texto_item = re.sub(r'^[•\-\*]\s+', '', linha_stripped)
                            linhas_formatadas.append(f'<li style="margin: 4px 0; color: #ffffff;">{texto_item}</li>')
                        
                        elif re.match(r'^\d+\.\s+', linha_stripped):
                            if em_lista and linhas_formatadas[-1].startswith('<ul'):
                                linhas_formatadas.append('</ul>')
                                em_lista = False
                            if not em_lista:
                                linhas_formatadas.append('<ol style="padding-left: 25px; margin: 8px 0;">')
                                em_lista = True
                            texto_item = re.sub(r'^\d+\.\s+', '', linha_stripped)
                            linhas_formatadas.append(f'<li style="margin: 4px 0; color: #ffffff;">{texto_item}</li>')
                        
                        else:
                            if em_lista:
                                if linhas_formatadas[-1].startswith('<ol'):
                                    linhas_formatadas.append('</ol>')
                                else:
                                    linhas_formatadas.append('</ul>')
                                em_lista = False
                            
                            if linha_stripped:
                                linhas_formatadas.append(f'<p style="margin: 8px 0;">{linha_stripped}</p>')
                    
                    if em_lista:
                        if linhas_formatadas and linhas_formatadas[-1].startswith('<li'):
                            if '<ol' in '\n'.join(linhas_formatadas):
                                linhas_formatadas.append('</ol>')
                            else:
                                linhas_formatadas.append('</ul>')
                    
                    texto_iniciativa = '\n'.join(linhas_formatadas)
                    
                    st.markdown(f"""
                    <div style='
                        text-align: justify; 
                        padding: 12px; 
                        background-color: #2d2d2d; 
                        border-radius: 3px;
                        font-size: 13px;
                        color: #ffffff;
                        line-height: 1.7;
                    '>
                        {texto_iniciativa}
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.info("Nenhuma iniciativa cadastrada")
                
                # Linha 6: Informações Complementares
                info_complementar = row.get('Informação complementar texto', '')
                if pd.notna(info_complementar) and str(info_complementar).strip() != '':
                    st.divider()
                    st.markdown("### 📝 Informações Complementares")
                    
                    # Processar texto para aplicar formatação
                    texto_formatado = str(info_complementar)
                    
                    # Importar re aqui
                    import re
                    
                    # Aplicar negrito para **texto**
                    texto_formatado = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', texto_formatado)
                    
                    # Aplicar itálico para *texto* (cuidado para não conflitar com **)
                    texto_formatado = re.sub(r'(?<!\*)\*(?!\*)([^*]+)\*(?!\*)', r'<em>\1</em>', texto_formatado)
                    
                    # Detectar e formatar listas com marcadores
                    linhas = texto_formatado.split('\n')
                    linhas_formatadas = []
                    em_lista = False
                    
                    for linha in linhas:
                        linha_stripped = linha.strip()
                        
                        # Detectar marcadores no início da linha
                        if re.match(r'^[•\-\*]\s+', linha_stripped):
                            if not em_lista:
                                linhas_formatadas.append('<ul style="list-style-type: disc; padding-left: 25px; margin: 8px 0;">')
                                em_lista = True
                            # Remover o marcador e adicionar como item de lista
                            texto_item = re.sub(r'^[•\-\*]\s+', '', linha_stripped)
                            linhas_formatadas.append(f'<li style="margin: 4px 0; color: #e0e0e0;">{texto_item}</li>')
                        
                        # Detectar numeração no início da linha
                        elif re.match(r'^\d+\.\s+', linha_stripped):
                            if em_lista and linhas_formatadas[-1].startswith('<ul'):
                                linhas_formatadas.append('</ul>')
                                em_lista = False
                            if not em_lista:
                                linhas_formatadas.append('<ol style="padding-left: 25px; margin: 8px 0;">')
                                em_lista = True
                            # Remover a numeração e adicionar como item de lista
                            texto_item = re.sub(r'^\d+\.\s+', '', linha_stripped)
                            linhas_formatadas.append(f'<li style="margin: 4px 0; color: #e0e0e0;">{texto_item}</li>')
                        
                        else:
                            # Se estava em lista, fechar
                            if em_lista:
                                if linhas_formatadas[-1].startswith('<ol'):
                                    linhas_formatadas.append('</ol>')
                                else:
                                    linhas_formatadas.append('</ul>')
                                em_lista = False
                            
                            # Adicionar linha normal (se não estiver vazia)
                            if linha_stripped:
                                linhas_formatadas.append(f'<p style="margin: 8px 0;">{linha_stripped}</p>')
                    
                    # Fechar lista se terminou em lista
                    if em_lista:
                        if linhas_formatadas[-1].startswith('<li'):
                            if '<ol' in '\n'.join(linhas_formatadas):
                                linhas_formatadas.append('</ol>')
                            else:
                                linhas_formatadas.append('</ul>')
                    
                    texto_formatado = '\n'.join(linhas_formatadas)
                    
                    st.markdown(f"""
                    <div style='
                        text-align: justify; 
                        padding: 15px; 
                        background-color: #252526; 
                        border-radius: 5px; 
                        font-size: 13px; 
                        color: #e0e0e0; 
                        line-height: 1.8;
                        border-left: 3px solid #E36C0A;
                    '>
                        {texto_formatado}
                    </div>
                    """, unsafe_allow_html=True)
    
    else:
        st.warning("Nenhuma meta encontrada com os filtros aplicados.")

except FileNotFoundError:
    st.error("❌ Arquivo 'teste_integração.xlsx' não encontrado. Execute o extrator primeiro para gerar os dados.")
except Exception as e:
    st.error(f"❌ Erro ao carregar dados: {e}")
    import traceback
    st.code(traceback.format_exc())

# Rodapé
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #a0a0a0; font-size: 11px; padding: 20px 0;'>
    <p style='margin: 0;'>Sistema de Monitoramento Estratégico - TJMG</p>
    <p style='margin: 0;'>Versão Dashboard Interativo 1.0</p>
    <p style='margin: 5px 0 0 0; font-size: 9px;'>ASPLAG • DEPLAG</p>
</div>
""", unsafe_allow_html=True)