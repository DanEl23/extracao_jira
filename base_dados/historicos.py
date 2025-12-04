"""
Dados Históricos de Metas Estratégicas
Contém históricos de metas aprovadas e por macrodesafio (2021-2024)
"""

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
