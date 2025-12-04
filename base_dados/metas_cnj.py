"""
Dados e Configurações das Metas CNJ
Contém alvos por categoria e mapeamento para macrodesafios
"""

# ============================================
# METAS CNJ - ALVOS POR CATEGORIA
# ============================================

# Mapeamento dos valores alvo para cada meta CNJ e sua categoria
METAS_CNJ_ALVOS = {
    ("CNJ 1", "1º Grau"): 100.0,
    ("CNJ 1", "2º Grau"): 100.0,
    ("CNJ 1", "Turma Recursal"): 100.0,
    ("CNJ 1", "Juizado Especial"): 100.0,
    ("CNJ 1", "Total"): 100.0,  # Linha agregadora da Meta 1
    ("CNJ 2", "1º Grau"): 80.0,
    ("CNJ 2", "2º Grau"): 90.0,
    ("CNJ 2", "Juizados e Turmas"): 95.0,
    ("CNJ 2", "Processos mais Antigos"): 100.0,
    ("CNJ 3", "Total"): 1.0,
    ("CNJ 4", "Crimes contra a administração pública"): 65.0,
    ("CNJ 4", "Improbidade administrativa"): 100.0,
    ("CNJ 5", "Total"): 100.0,
    ("CNJ 6", "Total"): 50.0,
    ("CNJ 7", "Total Indígenas"): 50.0,
    ("CNJ 7", "Total Quilombola"): 50.0,
    ("CNJ 8", "Total Feminicídio"): 75.0,
    ("CNJ 8", "Total Violência Doméstica"): 90.0,
    ("CNJ 9", "Total"): 100.0,
    ("CNJ 10", "1º Grau"): 90.0,
    ("CNJ 10", "2º Grau"): 100.0,
}

# Mapeamento de Metas CNJ para Macrodesafios
METAS_CNJ_PARA_MACRODESAFIO = {
    "CNJ 1": 3,   # Macro 3 - Agilidade e Produtividade na Prestação Jurisdicional
    "CNJ 2": 3,   # Macro 3 - Agilidade e Produtividade na Prestação Jurisdicional
    "CNJ 3": 5,   # Macro 5 - Prevenção de Litígios e Adoção de Soluções Consensuais
    "CNJ 4": 4,   # Macro 4 - Integridade, Segurança Institucional e Prevenção à Corrupção
    "CNJ 5": 3,   # Macro 3 - Agilidade e Produtividade na Prestação Jurisdicional
    "CNJ 6": 3,   # Macro 3 - Agilidade e Produtividade na Prestação Jurisdicional
    "CNJ 7": 3,   # Macro 3 - Agilidade e Produtividade na Prestação Jurisdicional
    "CNJ 8": 3,   # Macro 3 - Agilidade e Produtividade na Prestação Jurisdicional
    "CNJ 9": 9,   # Macro 9 - Aprimoramento da Gestão Administrativa e da Governança Judiciária
    "CNJ 10": 3,  # Macro 3 - Agilidade e Produtividade na Prestação Jurisdicional
}
