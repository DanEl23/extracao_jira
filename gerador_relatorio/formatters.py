"""
Funções de formatação de valores e cores
"""

import pandas as pd


def formatar_valor(valor):
    """Formata valor para exibição"""
    if pd.isna(valor) or valor == '':
        return '-'
    elif isinstance(valor, (int, float)):
        return f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    else:
        return str(valor)


def rgb_to_hex(rgb_tuple):
    """Converte tupla RGB para hexadecimal"""
    return f"{rgb_tuple[0]:02x}{rgb_tuple[1]:02x}{rgb_tuple[2]:02x}"
