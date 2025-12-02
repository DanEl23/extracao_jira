"""
Mapeamento de Metas para Superintendências - TJMG
Versão: 3.0 - Incluindo estatísticas históricas e cálculos automáticos
"""

import pandas as pd
import re
from pathlib import Path


# ============================================
# MAPEAMENTO PRINCIPAL
# ============================================

# Mapeamento principal: código da meta -> superintendência responsável
META_SUPERINTENDENCIA = {
    # Metas CNJ
    "CNJ 1": "PRESIDÊNCIA",
    "CNJ 10": "PRESIDÊNCIA",
    "CNJ 2": "PRESIDÊNCIA",
    "CNJ 3": "PRESIDÊNCIA",
    "CNJ 4": "PRESIDÊNCIA",
    "CNJ 5": "PRESIDÊNCIA",
    "CNJ 6": "PRESIDÊNCIA",
    "CNJ 7": "PRESIDÊNCIA",
    "CNJ 8": "PRESIDÊNCIA",
    "CNJ 9": "PRESIDÊNCIA",
    
    # Metas TJMG - PRESIDÊNCIA
    "TJMG 104": "PRESIDÊNCIA",
    "TJMG 109": "PRESIDÊNCIA",
    "TJMG 111": "PRESIDÊNCIA",
    "TJMG 123": "PRESIDÊNCIA",
    "TJMG 124": "PRESIDÊNCIA",
    "TJMG 126": "PRESIDÊNCIA",
    "TJMG 128": "PRESIDÊNCIA",
    "TJMG 130": "PRESIDÊNCIA",
    "TJMG 131": "PRESIDÊNCIA",
    "TJMG 133": "PRESIDÊNCIA",  # Nota: estava "Presidência" (minúscula) no JSON
    "TJMG 136": "PRESIDÊNCIA",
    "TJMG 137": "PRESIDÊNCIA",
    "TJMG 143": "PRESIDÊNCIA",
    "TJMG 144": "PRESIDÊNCIA",
    "TJMG 145": "PRESIDÊNCIA",
    "TJMG 150": "PRESIDÊNCIA",
    "TJMG 151": "PRESIDÊNCIA",
    "TJMG 152": "PRESIDÊNCIA",
    "TJMG 153": "PRESIDÊNCIA",
    "TJMG 154": "PRESIDÊNCIA",
    "TJMG 156": "PRESIDÊNCIA",
    "TJMG 17": "PRESIDÊNCIA",
    "TJMG 62": "PRESIDÊNCIA",
    "TJMG 69": "PRESIDÊNCIA",
    "TJMG 80": "PRESIDÊNCIA",
    "TJMG 85": "PRESIDÊNCIA",
    
    # Metas TJMG - 1ª VICE-PRESIDÊNCIA
    "TJMG 115": "1ª VICE-PRESIDÊNCIA",
    "TJMG 132": "1ª VICE-PRESIDÊNCIA",
    "TJMG 134": "1ª VICE-PRESIDÊNCIA",
    "TJMG 135": "1ª VICE-PRESIDÊNCIA",
    "TJMG 155": "1ª VICE-PRESIDÊNCIA",
    "TJMG 29": "1ª VICE-PRESIDÊNCIA",
    "TJMG 59": "1ª VICE-PRESIDÊNCIA",
    "TJMG 5": "1ª VICE-PRESIDÊNCIA",
    "TJMG 6": "1ª VICE-PRESIDÊNCIA",
    "TJMG 7": "1ª VICE-PRESIDÊNCIA",
    "TJMG 91": "1ª VICE-PRESIDÊNCIA",
    "TJMG 95": "1ª VICE-PRESIDÊNCIA",
    "TJMG 96": "1ª VICE-PRESIDÊNCIA",
    
    # Metas TJMG - 2ª VICE PRESIDÊNCIA
    "TJMG 67": "2ª VICE PRESIDÊNCIA",
    
    # Metas TJMG - 3ª VICE - PRESIDÊNCIA
    "TJMG 100": "3ª VICE - PRESIDÊNCIA",
    "TJMG 138": "3ª VICE - PRESIDÊNCIA",
    "TJMG 139": "3ª VICE - PRESIDÊNCIA",
    "TJMG 140": "3ª VICE - PRESIDÊNCIA",
    "TJMG 97": "3ª VICE - PRESIDÊNCIA",
    
    # Metas TJMG - CORREGEDORIA
    "TJMG 10": "CORREGEDORIA",
    "TJMG 129": "CORREGEDORIA",
    "TJMG 141": "CORREGEDORIA",
    "TJMG 142": "CORREGEDORIA",
    "TJMG 146": "CORREGEDORIA",
    "TJMG 147": "CORREGEDORIA",
    "TJMG 148": "CORREGEDORIA",
    "TJMG 149": "CORREGEDORIA",
    "TJMG 40": "CORREGEDORIA",
}


# Ordem de exibição das superintendências no relatório
ORDEM_SUPERINTENDENCIAS = [
    'PRESIDÊNCIA',
    '1ª VICE-PRESIDÊNCIA',
    '2ª VICE PRESIDÊNCIA',
    '3ª VICE - PRESIDÊNCIA',
    'CORREGEDORIA',
    'SEM CLASSIFICAÇÃO'
]


# ============================================
# ESTATÍSTICAS HISTÓRICAS
# ============================================

# Dados da tabela: Total de Metas Aprovadas (2021-2024)
HISTORICO_METAS_APROVADAS = {
    2021: {'nacionais': 10, 'institucionais': 56, 'total': 66},
    2022: {'nacionais': 10, 'institucionais': 59, 'total': 69},
    2023: {'nacionais': 9, 'institucionais': 67, 'total': 76},
    2024: {'nacionais': 9, 'institucionais': 74, 'total': 83},
}

# Cálculo de variação 2021-2024 (será recalculado automaticamente)
VARIACAO_2021_2024 = {
    'nacionais': -10,  # % (9 vs 10)
    'institucionais': 32,  # % (74 vs 56)
    'total': 24  # % (83 vs 66)
}


# Dados históricos: Total de Metas por Macrodesafio (2021-2024)
HISTORICO_METAS_POR_MACRODESAFIO = {
    2021: {
        '1 - Garantia dos Direitos Fundamentais e do Estado Democrático de Direito': 1,
        '2 - Ampliação da relação institucional do Judiciário com a Sociedade': 3,
        '3 - Agilidade e Produtividade na Prestação Jurisdicional': 22,
        '4 - Enfrentamento à Corrupção e à Improbidade Administrativa': 4,
        '5 - Prevenção de Litígios e Adoção de Soluções Consensuais para os Conflitos': 7,
        '6 - Consolidação do Sistema de Precedentes Obrigatórios': 2,
        '7 - Promoção da Sustentabilidade': 1,
        '8 - Aperfeiçoamento da Administração do Sistema de Justiça': 8,
        '9 - Aprimoramento da Gestão Administrativa e da Governança Judiciária': 9,
        '10 - Otimização da Gestão de Pessoas': 4,
        '11 - Modernização da Gestão Orçamentária e Financeira': 1,
        '12 - Fortalecimento da Estratégia de Tecnologias da Informação e Comunicação - TIC e de Proteção de Dados': 4,
    },
    2022: {
        '1 - Garantia dos Direitos Fundamentais e do Estado Democrático de Direito': 8,
        '2 - Ampliação da relação institucional do Judiciário com a Sociedade': 3,
        '3 - Agilidade e Produtividade na Prestação Jurisdicional': 23,
        '4 - Enfrentamento à Corrupção e à Improbidade Administrativa': 2,
        '5 - Prevenção de Litígios e Adoção de Soluções Consensuais para os Conflitos': 6,
        '6 - Consolidação do Sistema de Precedentes Obrigatórios': 3,
        '7 - Promoção da Sustentabilidade': 1,
        '8 - Aperfeiçoamento da Administração do Sistema de Justiça': 1,
        '9 - Aprimoramento da Gestão Administrativa e da Governança Judiciária': 13,
        '10 - Otimização da Gestão de Pessoas': 4,
        '11 - Modernização da Gestão Orçamentária e Financeira': 3,
        '12 - Fortalecimento da Estratégia de Tecnologias da Informação e Comunicação - TIC e de Proteção de Dados': 2,
    },
    2023: {
        '1 - Garantia dos Direitos Fundamentais e do Estado Democrático de Direito': 7,
        '2 - Ampliação da relação institucional do Judiciário com a Sociedade': 1,
        '3 - Agilidade e Produtividade na Prestação Jurisdicional': 35,
        '4 - Enfrentamento à Corrupção e à Improbidade Administrativa': 2,
        '5 - Prevenção de Litígios e Adoção de Soluções Consensuais para os Conflitos': 4,
        '6 - Consolidação do Sistema de Precedentes Obrigatórios': 3,
        '7 - Promoção da Sustentabilidade': 1,
        '8 - Aperfeiçoamento da Administração do Sistema de Justiça': 1,
        '9 - Aprimoramento da Gestão Administrativa e da Governança Judiciária': 13,
        '10 - Otimização da Gestão de Pessoas': 4,
        '11 - Modernização da Gestão Orçamentária e Financeira': 3,
        '12 - Fortalecimento da Estratégia de Tecnologias da Informação e Comunicação - TIC e de Proteção de Dados': 2,
    },
    2024: {
        '1 - Garantia dos Direitos Fundamentais e do Estado Democrático de Direito': 9,
        '2 - Ampliação da relação institucional do Judiciário com a Sociedade': 1,
        '3 - Agilidade e Produtividade na Prestação Jurisdicional': 36,
        '4 - Enfrentamento à Corrupção e à Improbidade Administrativa': 2,
        '5 - Prevenção de Litígios e Adoção de Soluções Consensuais para os Conflitos': 4,
        '6 - Consolidação do Sistema de Precedentes Obrigatórios': 4,
        '7 - Promoção da Sustentabilidade': 1,
        '8 - Aperfeiçoamento da Administração do Sistema de Justiça': 2,
        '9 - Aprimoramento da Gestão Administrativa e da Governança Judiciária': 15,
        '10 - Otimização da Gestão de Pessoas': 3,
        '11 - Modernização da Gestão Orçamentária e Financeira': 4,
        '12 - Fortalecimento da Estratégia de Tecnologias da Informação e Comunicação - TIC e de Proteção de Dados': 2,
    },
}


# ============================================
# FUNÇÕES DE CÁLCULO AUTOMÁTICO
# ============================================

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

def calcular_metas_atuais():
    """
    Calcula o total de metas nacionais e institucionais das planilhas atuais.
    
    Returns:
        dict: {'nacionais': int, 'institucionais': int, 'total': int, 'ano': int}
    """
    try:
        # Metas Nacionais (CNJ)
        df_cnj = pd.read_excel('exports/resultados_cnj.xlsx')
        # Extrair número da meta - padrão simplificado para capturar "Meta 1" até "Meta 10"
        metas_cnj = df_cnj['Meta'].str.extract(r'Meta\s+(\d+)', expand=False).dropna().unique()
        total_nacionais = len(metas_cnj)
        
        # Metas Institucionais (TJMG)
        df_tjmg = pd.read_excel('exports/teste_integração.xlsx')
        # Extrair código (ex: "TJMG 111 - Descrição" -> "TJMG 111")
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
    Normaliza os macrodesafios pelo número inicial para evitar duplicatas.
    
    Returns:
        dict: {macrodesafio: quantidade}
    """
    import re
    
    try:
        df_tjmg = pd.read_excel('exports/teste_integração.xlsx')
        
        # Verificar se coluna existe
        if 'Macrodesafio' not in df_tjmg.columns:
            print("⚠️  Coluna 'Macrodesafio' não encontrada")
            return {}
        
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
        contagem = {}
        for numero in sorted(contagem_por_numero.keys()):
            primeiro_texto = df_tjmg[df_tjmg['Numero_Macro'] == numero]['Macrodesafio'].iloc[0]
            contagem[primeiro_texto] = contagem_por_numero[numero]
        
        return contagem
        
    except Exception as e:
        print(f"⚠️  Erro ao calcular metas por macrodesafio: {e}")
        return {}


def atualizar_historico_macrodesafio_com_ano_atual():
    """
    Atualiza o dicionário HISTORICO_METAS_POR_MACRODESAFIO com os dados do ano atual.
    Consolida macrodesafios pelo número para evitar duplicatas por diferenças textuais.
    
    Returns:
        dict: {'historico': dict, 'historico_recente': dict, 'anos_recentes': list}
    """
    import copy
    import re
    
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


# ============================================
# FUNÇÃO AUXILIAR (ORIGINAL)
# ============================================

# Função auxiliar para obter superintendência
def obter_superintendencia(codigo_meta):
    """
    Retorna a superintendência responsável por uma meta.
    
    Args:
        codigo_meta (str): Código da meta (ex: "TJMG 111")
    
    Returns:
        str: Nome da superintendência ou "SEM CLASSIFICAÇÃO"
    """
    return META_SUPERINTENDENCIA.get(codigo_meta, "SEM CLASSIFICAÇÃO")


def salvar_historico_atualizado(arquivo='meta_por_superintendencia.py'):
    """
    Atualiza permanentemente o dicionário HISTORICO_METAS_APROVADAS no arquivo Python
    com os dados do ano atual extraídos das planilhas.
    
    ATENÇÃO: Esta função modifica o código-fonte do módulo!
    
    Args:
        arquivo (str): Caminho do arquivo a ser atualizado (padrão: este mesmo arquivo)
    
    Returns:
        bool: True se atualizado com sucesso, False caso contrário
    """
    try:
        # Obter dados atualizados
        dados = atualizar_historico_com_ano_atual()
        historico = dados['historico']
        variacao = dados['variacao']
        
        # Ler arquivo atual
        with open(arquivo, 'r', encoding='utf-8') as f:
            linhas = f.readlines()
        
        # Encontrar início e fim do dicionário HISTORICO_METAS_APROVADAS
        inicio = None
        fim = None
        for i, linha in enumerate(linhas):
            if 'HISTORICO_METAS_APROVADAS = {' in linha:
                inicio = i
            if inicio is not None and linha.strip() == '}':
                fim = i
                break
        
        if inicio is None or fim is None:
            print("❌ Não foi possível localizar o dicionário HISTORICO_METAS_APROVADAS")
            return False
        
        # Gerar novo conteúdo do dicionário
        novo_dict = ["HISTORICO_METAS_APROVADAS = {\n"]
        for ano in sorted(historico.keys()):
            novo_dict.append(f"    {ano}: {{'nacionais': {historico[ano]['nacionais']}, "
                           f"'institucionais': {historico[ano]['institucionais']}, "
                           f"'total': {historico[ano]['total']}}},\n")
        novo_dict.append("}\n")
        
        # Substituir no arquivo
        novas_linhas = linhas[:inicio] + novo_dict + linhas[fim+1:]
        
        # Salvar arquivo
        with open(arquivo, 'w', encoding='utf-8') as f:
            f.writelines(novas_linhas)
        
        print(f"✅ Histórico atualizado com sucesso em {arquivo}")
        print(f"   Ano adicionado/atualizado: {max(historico.keys())}")
        return True
        
    except Exception as e:
        print(f"❌ Erro ao salvar histórico atualizado: {e}")
        return False


# Estatísticas do mapeamento
def estatisticas():
    """Retorna estatísticas sobre o mapeamento"""
    from collections import Counter
    
    total = len(META_SUPERINTENDENCIA)
    por_super = Counter(META_SUPERINTENDENCIA.values())
    
    return {
        'total_metas': total,
        'por_superintendencia': dict(por_super),
        'metas_cnj': len([k for k in META_SUPERINTENDENCIA.keys() if k.startswith('CNJ')]),
        'metas_tjmg': len([k for k in META_SUPERINTENDENCIA.keys() if k.startswith('TJMG')])
    }


if __name__ == "__main__":
    # Executar este arquivo diretamente mostra as estatísticas
    print("📊 ANÁLISE COMPLETA DO PLANEJAMENTO ESTRATÉGICO\n")
    
    # Estatísticas básicas
    stats = estatisticas()
    print("=" * 70)
    print("ESTATÍSTICAS DO MAPEAMENTO")
    print("=" * 70)
    print(f"Total de metas mapeadas: {stats['total_metas']}")
    print(f"  - Metas CNJ: {stats['metas_cnj']}")
    print(f"  - Metas TJMG: {stats['metas_tjmg']}")
    print("\nDistribuição por Superintendência:")
    for super_nome, count in sorted(stats['por_superintendencia'].items()):
        print(f"  {super_nome}: {count} metas")
    
    print("\n")
    
    # Metas atuais das planilhas
    print("=" * 70)
    print("METAS ATUAIS (DAS PLANILHAS)")
    print("=" * 70)
    metas_atuais = calcular_metas_atuais()
    print(f"Ano: {metas_atuais['ano']}")
    print(f"Metas Nacionais: {metas_atuais['nacionais']}")
    print(f"Metas Institucionais: {metas_atuais['institucionais']}")
    print(f"Total: {metas_atuais['total']}")
    
    print("\n")
    
    # Relatório histórico
    print(gerar_relatorio_historico())
    
    print("\n")
    
    # Distribuição por superintendência (atual)
    print("=" * 70)
    print("DISTRIBUIÇÃO ATUAL POR SUPERINTENDÊNCIA")
    print("=" * 70)
    dist = calcular_metas_por_superintendencia()
    for super_nome in ORDEM_SUPERINTENDENCIAS:
        if super_nome in dist:
            info = dist[super_nome]
            print(f"\n{super_nome}:")
            print(f"  Nacionais: {info['nacionais']}")
            print(f"  Institucionais: {info['institucionais']}")
            print(f"  Total: {info['total']}")
            print(f"  Metas: {', '.join(info['metas'][:5])}{'...' if len(info['metas']) > 5 else ''}")

