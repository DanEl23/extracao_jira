"""
Gerador de Relatórios - Package Modularizado

Este package contém todos os módulos necessários para gerar
relatórios de metas estratégicas do TJMG.
"""

# Importar configurações
from .config import Config

# Importar funções de dados
from .data_loader import (
    criar_pasta_saida,
    carregar_dados,
    carregar_mapeamento_superintendencias,
    extrair_codigo_meta,
    adicionar_coluna_superintendencia,
    agrupar_por_superintendencia_e_macro,
    agrupar_por_macrodesafio
)

# Importar formatadores
from .formatters import (
    formatar_valor,
    rgb_to_hex
)

# Importar estilos
from .styles import (
    set_cell_background,
    set_cell_border,
    set_paragraph_background,
    set_keep_together,
    set_keep_with_next
)

# Importar document builder
from .document_builder import (
    criar_documento,
    criar_secao_paisagem_inicial
)

# Importar tabelas históricas
from .table_historico import (
    adicionar_tabela_historica,
    adicionar_tabela_macrodesafio
)

# Importar tabela CNJ
from .table_cnj import adicionar_tabela_metas_nacionais

# Importar tabela de monitoramento
from .table_monitoramento import adicionar_tabela_resultado_monitoramento

# Importar tabelas de superintendência
from .table_superintendencia import (
    adicionar_nova_secao_superintendencia,
    adicionar_secao_macrodesafio,
    adicionar_tabela_indicador
)

__all__ = [
    # Config
    'Config',
    # Data
    'criar_pasta_saida',
    'carregar_dados',
    'carregar_mapeamento_superintendencias',
    'extrair_codigo_meta',
    'adicionar_coluna_superintendencia',
    'agrupar_por_superintendencia_e_macro',
    'agrupar_por_macrodesafio',
    # Formatters
    'formatar_valor',
    'rgb_to_hex',
    # Styles
    'set_cell_background',
    'set_cell_border',
    'set_paragraph_background',
    'set_keep_together',
    'set_keep_with_next',
    # Document Builder
    'criar_documento',
    'criar_secao_paisagem_inicial',
    # Table Historico
    'adicionar_tabela_historica',
    'adicionar_tabela_macrodesafio',
    # Table CNJ
    'adicionar_tabela_metas_nacionais',
    # Table Monitoramento
    'adicionar_tabela_resultado_monitoramento',
    # Table Superintendencia
    'adicionar_nova_secao_superintendencia',
    'adicionar_secao_macrodesafio',
    'adicionar_tabela_indicador'
]

