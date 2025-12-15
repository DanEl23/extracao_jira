"""
Configurações do Gerador de Relatórios
"""


class Config:
    """Configurações do relatório"""
    
    # Arquivo de entrada
    ARQUIVO_EXCEL = 'exports/teste_integração.xlsx'
    NOME_ABA = None  # Nome da aba (ou deixe None para a primeira)
    CAMINHO_IMAGEM_CAPA = 'templates/Capa.jpg'
    # Arquivo de saída
    PASTA_SAIDA = 'relatorios_gerados'
    NOME_RELATORIO = 'Relatorio_Metas_Estrategicas'
    
    # Colunas do Excel (ajuste conforme seus dados)
    COLUNAS = {
        'MACRODESAFIO': 'Macrodesafio',
        'METAKEY': 'MetaKey',
        'ANO_META': 'Ano da Meta',
        'MES_APURACAO': 'mes_apuracao',
        'INDICADOR': 'Indicador',
        'UNIDADE_GESTORA': 'Unidade Gestora',
        'POLARIDADE': 'Polaridade',
        'VALOR_APURADO': 'Valor Apurado',
        'VALOR_META': 'Valor da Meta',
        'INICIATIVA': 'Iniciativa',
        'INFO_COMPLEMENTAR': 'Informação complementar texto'
    }
    
    # Cores (RGB) - Formato correto: (R, G, B)
    CORES = {
        'CABECALHO_MACRO': (239, 108, 33),      # Laranja forte
        'CABECALHO_TABELA': (250, 191, 143),    # Laranja claro
        'TEXTO_BRANCO': (255, 255, 255),
        'TEXTO_PRETO': (0, 0, 0),
        'BORDA_TABELA': (166, 166, 166),        # Cinza para bordas
        'INICIATIVA_FUNDO': (250, 191, 143)     # Laranja claro para iniciativa
    }
    
    # Formatação
    FONTE_PADRAO = 'Calibri'
    TAMANHO_TITULO = 18
    TAMANHO_SUBTITULO = 14
    TAMANHO_TEXTO = 11
    TAMANHO_TABELA = 9
