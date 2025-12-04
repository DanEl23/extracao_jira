"""
Mapeamentos de Metas para Superintendências - TJMG
Contém o dicionário principal META_SUPERINTENDENCIA e ordenações
"""

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
    "TJMG 133": "PRESIDÊNCIA",
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
    'CORREGEDORIA'
]


def estatisticas():
    """Retorna estatísticas sobre o mapeamento"""
    metas_cnj = sum(1 for k in META_SUPERINTENDENCIA.keys() if k.startswith('CNJ'))
    metas_tjmg = sum(1 for k in META_SUPERINTENDENCIA.keys() if k.startswith('TJMG'))
    
    por_superintendencia = {}
    for superintendencia in META_SUPERINTENDENCIA.values():
        por_superintendencia[superintendencia] = por_superintendencia.get(superintendencia, 0) + 1
    
    return {
        'total_metas': len(META_SUPERINTENDENCIA),
        'metas_cnj': metas_cnj,
        'metas_tjmg': metas_tjmg,
        'por_superintendencia': por_superintendencia
    }
