"""
Base de Dados Fixos - Package Modularizado
Contém mapeamentos, históricos e funções de cálculo para o projeto TJMG
"""

# Importar mapeamentos
from .mapeamentos import (
    META_SUPERINTENDENCIA,
    ORDEM_SUPERINTENDENCIAS,
    estatisticas
)

# Importar metas CNJ
from .metas_cnj import (
    METAS_CNJ_ALVOS,
    METAS_CNJ_PARA_MACRODESAFIO
)

# Importar históricos
from .historicos import (
    HISTORICO_METAS_APROVADAS,
    HISTORICO_METAS_POR_MACRODESAFIO,
    VARIACAO_2021_2024
)

# Importar funções de cálculo
from .calculos import (
    calcular_metas_atuais,
    atualizar_historico_com_ano_atual,
    calcular_metas_por_superintendencia,
    calcular_metas_atuais_por_macrodesafio,
    atualizar_historico_macrodesafio_com_ano_atual,
    calcular_cumprimento_metas_cnj,
    gerar_relatorio_historico
)

__all__ = [
    # Mapeamentos
    'META_SUPERINTENDENCIA',
    'ORDEM_SUPERINTENDENCIAS',
    'estatisticas',
    # Metas CNJ
    'METAS_CNJ_ALVOS',
    'METAS_CNJ_PARA_MACRODESAFIO',
    # Históricos
    'HISTORICO_METAS_APROVADAS',
    'HISTORICO_METAS_POR_MACRODESAFIO',
    'VARIACAO_2021_2024',
    # Cálculos
    'calcular_metas_atuais',
    'atualizar_historico_com_ano_atual',
    'calcular_metas_por_superintendencia',
    'calcular_metas_atuais_por_macrodesafio',
    'atualizar_historico_macrodesafio_com_ano_atual',
    'calcular_cumprimento_metas_cnj',
    'gerar_relatorio_historico'
]
