"""
Funções para carregamento de dados
"""

import pandas as pd
import os
import re
from .config import Config


def criar_pasta_saida():
    """Cria pasta para salvar relatórios se não existir"""
    if not os.path.exists(Config.PASTA_SAIDA):
        os.makedirs(Config.PASTA_SAIDA)
        print(f"✅ Pasta '{Config.PASTA_SAIDA}' criada.")


def carregar_dados():
    """Carrega dados do Excel"""
    print(f"📂 Carregando dados de '{Config.ARQUIVO_EXCEL}'...")
    
    try:
        if Config.NOME_ABA:
            df = pd.read_excel(Config.ARQUIVO_EXCEL, sheet_name=Config.NOME_ABA)
        else:
            df = pd.read_excel(Config.ARQUIVO_EXCEL)
        
        df = df.dropna(how='all')
        
        print(f"✅ {len(df)} registros carregados.")
        return df
    
    except FileNotFoundError:
        print(f"❌ ERRO: Arquivo '{Config.ARQUIVO_EXCEL}' não encontrado!")
        print(f"   Certifique-se de que o arquivo está na mesma pasta que este script.")
        return None
    except Exception as e:
        print(f"❌ ERRO ao carregar dados: {e}")
        return None


def carregar_mapeamento_superintendencias():
    """Carrega mapeamento de metas para superintendências"""
    try:
        from base_dados_fixos import META_SUPERINTENDENCIA
        print(f"✅ Mapeamento de superintendências carregado ({len(META_SUPERINTENDENCIA)} metas).")
        return META_SUPERINTENDENCIA
    except ImportError:
        print("❌ ERRO: Módulo 'base_dados_fixos.py' não encontrado!")
        return {}
    except Exception as e:
        print(f"❌ ERRO ao carregar mapeamento: {e}")
        return {}


def extrair_codigo_meta(meta_texto):
    """Extrai código da meta (ex: 'TJMG 111 - Texto' -> 'TJMG 111')"""
    match = re.match(r'^([A-Z]+\s+\d+)', str(meta_texto).strip())
    if match:
        return match.group(1)
    return None


def adicionar_coluna_superintendencia(df, mapeamento):
    """Adiciona coluna de superintendência ao DataFrame baseado no mapeamento"""
    print("\n🔍 Verificando mapeamento de superintendências...")
    
    def obter_superintendencia(meta):
        codigo = extrair_codigo_meta(meta)
        
        if hasattr(obter_superintendencia, 'contador'):
            obter_superintendencia.contador += 1
        else:
            obter_superintendencia.contador = 1
        
        if obter_superintendencia.contador <= 10:
            resultado = mapeamento.get(codigo, 'SEM CLASSIFICAÇÃO') if codigo else 'SEM CLASSIFICAÇÃO'
            print(f"   Linha {obter_superintendencia.contador}: Meta='{str(meta)[:50]}...' | Código='{codigo}' | Superintendência='{resultado}'")
        
        if codigo:
            superintendencia = mapeamento.get(codigo, 'SEM CLASSIFICAÇÃO')
            return superintendencia.upper() if superintendencia != 'SEM CLASSIFICAÇÃO' else superintendencia
        return 'SEM CLASSIFICAÇÃO'
    
    df['Superintendência'] = df[Config.COLUNAS['METAKEY']].apply(obter_superintendencia)
    
    classificacao_counts = df['Superintendência'].value_counts()
    print(f"\n📊 Estatísticas de classificação:")
    for super_nome, count in classificacao_counts.items():
        print(f"   {super_nome}: {count} registros")
    
    return df


def agrupar_por_superintendencia_e_macro(df):
    """Agrupa dados primeiro por Superintendência, depois por Macrodesafio"""
    print("📊 Agrupando dados por Superintendência e Macrodesafio...")
    
    from base_dados_fixos import ORDEM_SUPERINTENDENCIAS
    
    import re
    df['_ordem_macro'] = df[Config.COLUNAS['MACRODESAFIO']].apply(
        lambda x: int(re.match(r'^(\d+)', str(x)).group(1)) if pd.notna(x) and re.match(r'^(\d+)', str(x)) else 999
    )
    
    # Criar dicionário com ordem das superintendências
    ordem_dict = {super: idx for idx, super in enumerate(ORDEM_SUPERINTENDENCIAS)}
    df['_ordem_super'] = df['Superintendência'].map(lambda x: ordem_dict.get(x, 999))
    
    # Ordenar por superintendência e depois por macrodesafio
    df = df.sort_values(['_ordem_super', '_ordem_macro'])
    
    # Agrupar por superintendência
    grupos_super = {}
    for superintendencia in ORDEM_SUPERINTENDENCIAS:
        df_super = df[df['Superintendência'] == superintendencia]
        if len(df_super) > 0:
            # Dentro de cada superintendência, agrupar por macrodesafio
            grupos_macro = df_super.groupby(Config.COLUNAS['MACRODESAFIO'], sort=False)
            grupos_super[superintendencia] = list(grupos_macro)
    
    print(f"✅ {len(grupos_super)} Superintendências com dados encontradas.")
    return grupos_super


def agrupar_por_macrodesafio(df):
    """Agrupa dados por Macrodesafio"""
    print("📊 Agrupando dados por Macrodesafio...")
    
    col_macro = Config.COLUNAS['MACRODESAFIO']
    
    import re
    df['_ordem_macro'] = df[col_macro].apply(
        lambda x: int(re.match(r'^(\d+)', str(x)).group(1)) if pd.notna(x) and re.match(r'^(\d+)', str(x)) else 999
    )
    
    df = df.sort_values('_ordem_macro')
    grupos = df.groupby(col_macro, sort=False)
    
    print(f"✅ {len(grupos)} Macrodesafios encontrados.")
    return grupos
