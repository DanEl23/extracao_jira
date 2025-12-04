"""
Funções de estilização de células e elementos do documento
"""

from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from .config import Config
from .formatters import rgb_to_hex


def set_cell_background(cell, rgb_tuple):
    """Define cor de fundo de uma célula da tabela"""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), rgb_to_hex(rgb_tuple))
    cell._element.get_or_add_tcPr().append(shading_elm)


def set_cell_border(cell, **kwargs):
    """Define bordas para células da tabela
    
    Args:
        cell: Célula da tabela
        **kwargs: Bordas a serem definidas (top, left, bottom, right)
                 Pode ser bool (True/False) ou dict com 'color' e 'size'
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    # Remover bordas antigas se existirem
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is not None:
        tcPr.remove(tcBorders)
    
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_config = kwargs.get(edge)
        if edge_config:
            edge_el = OxmlElement(f'w:{edge}')
            edge_el.set(qn('w:val'), 'single')
            
            # Se for dict, permite customizar cor e tamanho
            if isinstance(edge_config, dict):
                size = edge_config.get('size', 2)  # Default 2 = 1/4pt
                color = edge_config.get('color', Config.CORES['BORDA_TABELA'])
                edge_el.set(qn('w:sz'), str(size * 8))  # Converter pt para eighths of a point
                edge_el.set(qn('w:color'), rgb_to_hex(color))
            else:
                # Se for bool True, usar valores padrão
                edge_el.set(qn('w:sz'), '2')  # 2 = 1/4pt (sz é em oitavos de ponto)
                edge_el.set(qn('w:color'), rgb_to_hex(Config.CORES['BORDA_TABELA']))
            
            edge_el.set(qn('w:space'), '0')
            tcBorders.append(edge_el)
    
    tcPr.append(tcBorders)


def set_paragraph_background(paragraph, rgb_tuple):
    """Define cor de fundo de um parágrafo"""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), rgb_to_hex(rgb_tuple))
    paragraph._element.get_or_add_pPr().append(shading_elm)


def set_keep_together(table):
    """Força que a tabela não quebre entre linhas"""
    tbl = table._element
    for row in tbl.findall('.//w:tr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
        trPr = row.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr')
            row.insert(0, trPr)
        
        # Não permitir quebra dentro desta linha
        cantSplit = OxmlElement('w:cantSplit')
        trPr.append(cantSplit)


def set_keep_with_next(paragraph):
    """Força que o parágrafo/tabela permanece com o próximo elemento"""
    pPr = paragraph._element.get_or_add_pPr()
    keepNext = OxmlElement('w:keepNext')
    pPr.append(keepNext)
