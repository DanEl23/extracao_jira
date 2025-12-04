"""
Funções de Cálculo e Agregação de Dados
Contém todas as funções que processam dados das planilhas e atualizam históricos
"""

import pandas as pd
import re
from pathlib import Path
from .mapeamentos import META_SUPERINTENDENCIA, ORDEM_SUPERINTENDENCIAS
from .metas_cnj import METAS_CNJ_ALVOS, METAS_CNJ_PARA_MACRODESAFIO
from .historicos import HISTORICO_METAS_APROVADAS, HISTORICO_METAS_POR_MACRODESAFIO


def calcular_metas_atuais():
    """
    Calcula o total de metas nacionais e institucionais das planilhas atuais.
    
    Returns:
        dict: {'nacionais': int, 'institucionais': int, 'total': int, 'ano': int}
    """
    try:
        # Metas Nacionais (CNJ)
        df_cnj = pd.read_excel('exports/resultados_cnj.xlsx')
        metas_cnj = df_cnj['Meta'].str.extract(r'Meta\s+(\d+)', expand=False).dropna().unique()
        total_nacionais = len(metas_cnj)
        
        # Metas Institucionais (TJMG)
        df_tjmg = pd.read_excel('exports/teste_integração.xlsx')
        metas_tjmg = df_tjmg['MetaKey'].str.extract(r'^(TJMG\s+\d+)', expand=False).dropna().unique()
        total_institucionais = len(metas_tjmg)
        
        # Ano (pegar da coluna Ano da Meta)
        ano = df_tjmg['Ano da Meta'].mode()[0] if 'Ano da Meta' in df_tjmg.columns else 2025
        
        return {
            'nacionais': total_nacionais,
            'institucionais': total_institucionais,
            'total': total_nacionais + total_institucionais,
            'ano': int(ano)
        }
    except Exception as e:
        print(f"⚠️  Erro ao calcular metas atuais: {e}")
        return {'nacionais': 0, 'institucionais': 0, 'total': 0, 'ano': 2025}


def atualizar_historico_com_ano_atual():
    """
    Atualiza o dicionário HISTORICO_METAS_APROVADAS com os dados do ano atual
    extraídos das planilhas, e recalcula as variações.
    
    Returns:
        dict: Histórico atualizado com o ano atual
    """
    import copy
    
    # Copiar histórico existente
    historico_atualizado = copy.deepcopy(HISTORICO_METAS_APROVADAS)
    
    # Obter dados atuais das planilhas
    metas_atuais = calcular_metas_atuais()
    ano_atual = metas_atuais['ano']
    
    # Adicionar ou atualizar ano atual no histórico
    historico_atualizado[ano_atual] = {
        'nacionais': metas_atuais['nacionais'],
        'institucionais': metas_atuais['institucionais'],
        'total': metas_atuais['total']
    }
    
    # Recalcular variação apenas nos últimos 4 anos
    anos_ordenados = sorted(historico_atualizado.keys())
    
    # Pegar apenas os 4 anos mais recentes
    anos_recentes = anos_ordenados[-4:] if len(anos_ordenados) >= 4 else anos_ordenados
    primeiro_ano = anos_recentes[0]
    ultimo_ano = anos_recentes[-1]
    
    variacao_atualizada = {
        'nacionais': round(((historico_atualizado[ultimo_ano]['nacionais'] - 
                             historico_atualizado[primeiro_ano]['nacionais']) / 
                            historico_atualizado[primeiro_ano]['nacionais']) * 100),
        'institucionais': round(((historico_atualizado[ultimo_ano]['institucionais'] - 
                                  historico_atualizado[primeiro_ano]['institucionais']) / 
                                 historico_atualizado[primeiro_ano]['institucionais']) * 100),
        'total': round(((historico_atualizado[ultimo_ano]['total'] - 
                         historico_atualizado[primeiro_ano]['total']) / 
                        historico_atualizado[primeiro_ano]['total']) * 100)
    }
    
    return {
        'historico': historico_atualizado,
        'historico_recente': {ano: historico_atualizado[ano] for ano in anos_recentes},
        'anos_recentes': anos_recentes,
        'variacao': variacao_atualizada,
        'periodo': f"{primeiro_ano} - {ultimo_ano}"
    }


def calcular_metas_por_superintendencia():
    """
    Calcula quantas metas cada superintendência tem baseado na planilha atual.
    
    Returns:
        dict: {superintendencia: {'nacionais': int, 'institucionais': int, 'total': int}}
    """
    try:
        df_tjmg = pd.read_excel('exports/teste_integração.xlsx')
        
        resultado = {}
        for superintendencia in ORDEM_SUPERINTENDENCIAS:
            # Extrair código da meta
            codigos_meta = df_tjmg['MetaKey'].str.extract(r'^([A-Z]+\s+\d+)', expand=False)
            
            # Filtrar metas desta superintendência
            metas_super = []
            for codigo in codigos_meta.dropna().unique():
                if META_SUPERINTENDENCIA.get(codigo, '') == superintendencia:
                    metas_super.append(codigo)
            
            # Separar CNJ e TJMG
            nacionais = len([m for m in metas_super if m.startswith('CNJ')])
            institucionais = len([m for m in metas_super if m.startswith('TJMG')])
            
            if nacionais > 0 or institucionais > 0:
                resultado[superintendencia] = {
                    'nacionais': nacionais,
                    'institucionais': institucionais,
                    'total': nacionais + institucionais,
                    'metas': sorted(metas_super)
                }
        
        return resultado
    except Exception as e:
        print(f"⚠️  Erro ao calcular metas por superintendência: {e}")
        return {}


def calcular_metas_atuais_por_macrodesafio():
    """
    Calcula quantas metas existem por macrodesafio no ano atual.
    Inclui tanto metas institucionais (TJMG) quanto metas nacionais (CNJ).
    Normaliza os macrodesafios pelo número inicial para evitar duplicatas.
    
    Returns:
        dict: {macrodesafio: quantidade}
    """
    contagem = {}
    
    # === PROCESSAR METAS INSTITUCIONAIS (TJMG) ===
    try:
        df_tjmg = pd.read_excel('exports/teste_integração.xlsx')
        
        # Verificar se coluna existe
        if 'Macrodesafio' in df_tjmg.columns:
            # Extrair apenas o número do macrodesafio
            def extrair_numero(texto):
                if pd.isna(texto):
                    return None
                match = re.match(r'^(\d+)', str(texto).strip())
                return int(match.group(1)) if match else None
            
            # Adicionar coluna com número do macrodesafio
            df_tjmg['Numero_Macro'] = df_tjmg['Macrodesafio'].apply(extrair_numero)
            
            # Remover valores None
            df_tjmg = df_tjmg[df_tjmg['Numero_Macro'].notna()]
            
            # Agrupar por número e contar
            contagem_por_numero = df_tjmg.groupby('Numero_Macro').size().to_dict()
            
            # Para cada número, pegar o primeiro texto encontrado
            for numero in sorted(contagem_por_numero.keys()):
                primeiro_texto = df_tjmg[df_tjmg['Numero_Macro'] == numero]['Macrodesafio'].iloc[0]
                contagem[primeiro_texto] = contagem_por_numero[numero]
        else:
            print("⚠️  Coluna 'Macrodesafio' não encontrada em teste_integração.xlsx")
            
    except Exception as e:
        print(f"⚠️  Erro ao processar metas TJMG: {e}")
    
    # === PROCESSAR METAS CNJ ===
    try:
        df_cnj = pd.read_excel('exports/resultados_cnj.xlsx')
        
        # Extrair código da meta (CNJ X)
        def extrair_codigo_meta(texto):
            if pd.isna(texto):
                return None
            match = re.search(r'Meta\s+(\d+)', str(texto), re.IGNORECASE)
            if match:
                return f"CNJ {match.group(1)}"
            return None
        
        df_cnj['Meta_CNJ'] = df_cnj['Meta'].apply(extrair_codigo_meta)
        
        # Remover duplicatas (cada meta pode ter múltiplas categorias)
        metas_unicas = df_cnj['Meta_CNJ'].dropna().unique()
        
        # Contar quantas metas CNJ existem por macrodesafio
        for meta_cnj in metas_unicas:
            if meta_cnj in METAS_CNJ_PARA_MACRODESAFIO:
                numero_macro = METAS_CNJ_PARA_MACRODESAFIO[meta_cnj]
                
                # Encontrar o texto do macrodesafio correspondente
                macro_texto = None
                for texto_macro in contagem.keys():
                    if str(texto_macro).startswith(str(numero_macro)):
                        macro_texto = texto_macro
                        break
                
                # Se não encontrou, criar entrada nova
                if macro_texto is None:
                    # Buscar nome completo do macrodesafio na planilha TJMG
                    try:
                        df_ref = pd.read_excel('exports/teste_integração.xlsx')
                        macros_com_numero = df_ref[df_ref['Macrodesafio'].str.startswith(str(numero_macro), na=False)]
                        if not macros_com_numero.empty:
                            macro_texto = macros_com_numero['Macrodesafio'].iloc[0]
                        else:
                            macro_texto = f"{numero_macro} - Macrodesafio {numero_macro}"
                    except:
                        macro_texto = f"{numero_macro} - Macrodesafio {numero_macro}"
                    
                    contagem[macro_texto] = 0
                
                # Incrementar contador
                contagem[macro_texto] += 1
        
        print(f"✅ Processadas {len(metas_unicas)} metas CNJ")
        
    except Exception as e:
        print(f"⚠️  Erro ao processar metas CNJ: {e}")
    
    return contagem


def atualizar_historico_macrodesafio_com_ano_atual():
    """
    Atualiza o dicionário HISTORICO_METAS_POR_MACRODESAFIO com os dados do ano atual.
    Consolida macrodesafios pelo número para evitar duplicatas por diferenças textuais.
    
    Returns:
        dict: {'historico': dict, 'historico_recente': dict, 'anos_recentes': list}
    """
    import copy
    
    def extrair_numero_macro(macro):
        """Extrai o número do macrodesafio"""
        match = re.match(r'^(\d+)', str(macro))
        return int(match.group(1)) if match else 999
    
    def consolidar_macros_por_numero(dados_ano, macros_referencia=None):
        """
        Consolida macrodesafios com mesmo número, somando suas quantidades.
        Se macros_referencia for fornecido, usa esses nomes como padrão.
        """
        consolidado = {}
        macros_por_numero = {}
        
        # Se há referência, usar os nomes dela
        if macros_referencia:
            for macro_ref in macros_referencia:
                numero = extrair_numero_macro(macro_ref)
                macros_por_numero[numero] = macro_ref
                consolidado[macro_ref] = 0
        
        # Processar dados do ano
        for macro, quantidade in dados_ano.items():
            numero = extrair_numero_macro(macro)
            
            if numero in macros_por_numero:
                # Usar nome de referência
                macro_ref = macros_por_numero[numero]
                consolidado[macro_ref] += quantidade
            else:
                # Novo macrodesafio
                macros_por_numero[numero] = macro
                consolidado[macro] = quantidade
        
        # Remover macrodesafios com quantidade 0
        return {k: v for k, v in consolidado.items() if v > 0}
        
    # Criar lista de referência a partir do primeiro ano do histórico
    primeiro_ano = min(HISTORICO_METAS_POR_MACRODESAFIO.keys())
    macros_referencia = list(HISTORICO_METAS_POR_MACRODESAFIO[primeiro_ano].keys())
    
    # Copiar histórico existente e consolidar cada ano usando os nomes de referência
    historico_atualizado = {}
    for ano, dados in HISTORICO_METAS_POR_MACRODESAFIO.items():
        historico_atualizado[ano] = consolidar_macros_por_numero(dados, macros_referencia)
    
    # Obter dados atuais das planilhas
    metas_atuais = calcular_metas_atuais()
    ano_atual = metas_atuais['ano']
    
    # Obter contagem por macrodesafio e consolidar usando os nomes de referência
    contagem_macro = calcular_metas_atuais_por_macrodesafio()
    
    # Adicionar ou atualizar ano atual no histórico, usando nomes de referência
    historico_atualizado[ano_atual] = consolidar_macros_por_numero(contagem_macro, macros_referencia)
    
    # Pegar apenas os 4 anos mais recentes
    anos_ordenados = sorted(historico_atualizado.keys())
    anos_recentes = anos_ordenados[-4:] if len(anos_ordenados) >= 4 else anos_ordenados
    
    historico_recente = {ano: historico_atualizado[ano] for ano in anos_recentes}
    
    # Obter lista única de todos os macrodesafios (ordenada por número)
    todos_macros = set()
    for ano_data in historico_recente.values():
        todos_macros.update(ano_data.keys())
    
    macros_ordenados = sorted(todos_macros, key=extrair_numero_macro)
    
    return {
        'historico': historico_atualizado,
        'historico_recente': historico_recente,
        'anos_recentes': anos_recentes,
        'macrodesafios': macros_ordenados
    }


def calcular_cumprimento_metas_cnj():
    """
    Calcula o cumprimento agregado das metas CNJ considerando múltiplas categorias.
    
    Lógica:
    1. Para cada categoria, ajustar cumprimento (máximo 100%)
    2. Calcular valor apurado = (cumprimento_ajustado * alvo) / 100
    3. Calcular média aritmética dos valores apurados por meta
    4. Retornar cumprimento final de cada meta CNJ
    
    Returns:
        dict: {
            'CNJ 1': {'cumprimento': float, 'categorias': int},
            'CNJ 2': {'cumprimento': float, 'categorias': int},
            ...
        }
    """
    try:
        # Carregar dados CNJ
        df_cnj = pd.read_excel('exports/resultados_cnj.xlsx')
        
        # Extrair número da meta para agrupar
        df_cnj['Meta_Numero'] = df_cnj['Meta'].str.extract(r'(Meta\s+\d+)', expand=False)
        df_cnj['Meta_CNJ'] = df_cnj['Meta_Numero'].str.replace('Meta', 'CNJ')
        
        # Remover linhas sem meta identificada
        df_cnj = df_cnj[df_cnj['Meta_CNJ'].notna()].copy()
        
        # Converter resultado para float (remover % se necessário)
        def converter_percentual(val):
            if pd.isna(val):
                return 0.0
            val_str = str(val).replace('%', '').replace(',', '.').strip()
            try:
                return float(val_str)
            except:
                return 0.0
        
        df_cnj['Resultado_Float'] = df_cnj['Resultado'].apply(converter_percentual)
        
        # Calcular cumprimento por meta
        resultado_metas = {}
        
        for meta_cnj in df_cnj['Meta_CNJ'].unique():
            df_meta = df_cnj[df_cnj['Meta_CNJ'] == meta_cnj]
            
            valores_apurados = []
            
            for _, row in df_meta.iterrows():
                categoria = row['Categoria']
                cumprimento_real = row['Resultado_Float']
                
                # Buscar alvo da meta
                chave_alvo = (meta_cnj, categoria)
                alvo = METAS_CNJ_ALVOS.get(chave_alvo)
                
                if alvo is None:
                    print(f"⚠️  Alvo não encontrado para {meta_cnj} - {categoria}")
                    continue
                
                # Ajustar cumprimento (máximo 100%)
                cumprimento_ajustado = min(cumprimento_real, 100.0)
                
                # Calcular valor apurado
                valor_apurado = (cumprimento_ajustado * alvo) / 100.0
                valores_apurados.append(valor_apurado)
            
            # Calcular média aritmética
            if valores_apurados:
                cumprimento_final = sum(valores_apurados) / len(valores_apurados)
                resultado_metas[meta_cnj] = {
                    'cumprimento': round(cumprimento_final, 2),
                    'categorias': len(valores_apurados)
                }
        
        return resultado_metas
        
    except Exception as e:
        print(f"❌ Erro ao calcular cumprimento metas CNJ: {e}")
        import traceback
        traceback.print_exc()
        return {}


def gerar_relatorio_historico():
    """
    Gera relatório comparativo incluindo dados históricos e atuais.
    Mostra apenas os 4 anos mais recentes.
    Atualiza automaticamente o histórico com o ano atual.
    
    Returns:
        str: Relatório formatado em texto
    """
    # Atualizar histórico com ano atual
    dados_atualizados = atualizar_historico_com_ano_atual()
    historico = dados_atualizados['historico_recente']  # Apenas 4 anos mais recentes
    anos = dados_atualizados['anos_recentes']
    variacao = dados_atualizados['variacao']
    periodo = dados_atualizados['periodo']
    
    relatorio = []
    relatorio.append("=" * 70)
    relatorio.append("TOTAL DE METAS APROVADAS PARA COMPOSIÇÃO DO PLANEJAMENTO")
    relatorio.append("ESTRATÉGICO INSTITUCIONAL DO TJMG")
    relatorio.append("=" * 70)
    relatorio.append("")
    
    # Cabeçalho dinâmico com os 4 anos mais recentes
    linha_header = f"{'Ano':<20}"
    for ano in anos:
        linha_header += f"{ano:<10}"
    relatorio.append(linha_header)
    relatorio.append("-" * 70)
    
    # Metas Nacionais
    linha_nac = f"{'Metas Nacionais':<20}"
    for ano in anos:
        linha_nac += f"{historico[ano]['nacionais']:<10}"
    relatorio.append(linha_nac)
    
    # Metas Institucionais
    linha_inst = f"{'Metas Institucionais':<20}"
    for ano in anos:
        linha_inst += f"{historico[ano]['institucionais']:<10}"
    relatorio.append(linha_inst)
    
    # Total
    linha_total = f"{'Total':<20}"
    for ano in anos:
        linha_total += f"{historico[ano]['total']:<10}"
    relatorio.append(linha_total)
    
    relatorio.append("=" * 70)
    
    # Variação automática (calculada apenas nos 4 anos mais recentes)
    relatorio.append(f"\nVariação {periodo}:")
    relatorio.append(f"  Metas Nacionais: {variacao['nacionais']:+}%")
    relatorio.append(f"  Metas Institucionais: {variacao['institucionais']:+}%")
    relatorio.append(f"  Total: {variacao['total']:+}%")
    
    return "\n".join(relatorio)
