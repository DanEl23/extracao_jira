"""
Funções para criação e configuração do documento base
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT, WD_SECTION
from .config import Config


def criar_documento(superintendencia='Presidência'):
    """Cria documento Word base com primeira página em retrato"""
    doc = Document()
    
    # === PRIMEIRA SEÇÃO: RETRATO (para tabela histórica) ===
    section_retrato = doc.sections[0]
    
    # Orientação retrato
    section_retrato.orientation = WD_ORIENT.PORTRAIT
    
    # Tamanho do papel A4 em retrato
    section_retrato.page_width = Cm(21.0)
    section_retrato.page_height = Cm(29.7)
    
    # Margens
    section_retrato.top_margin = Cm(2.5)
    section_retrato.bottom_margin = Cm(2.5)
    section_retrato.left_margin = Cm(2.0)
    section_retrato.right_margin = Cm(2.0)
    
    # Medianiz (gutter)
    section_retrato.gutter = Cm(0)
    
    # Cabeçalho e rodapé - distâncias reduzidas para ficarem mais próximos das margens
    section_retrato.header_distance = Cm(0.5)
    section_retrato.footer_distance = Cm(0.5)
    
    # === CABEÇALHO DA PRIMEIRA PÁGINA (RETRATO) ===
    header_retrato = section_retrato.header
    
    # Limpar cabeçalho padrão
    for paragraph in list(header_retrato.paragraphs):
        p_element = paragraph._element
        p_element.getparent().remove(p_element)
    
    # Linha 1: MONITORAMENTO DE METAS ESTRATÉGICAS - 2024 (negrito, tamanho 11, centralizado)
    p1 = header_retrato.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = p1.add_run('MONITORAMENTO DE METAS ESTRATÉGICAS - 2024')
    run1.font.size = Pt(11)
    run1.font.bold = True
    run1.font.name = Config.FONTE_PADRAO
    p1.paragraph_format.space_after = Pt(0)
    
    # Linha 2: Relatório Técnico ao Comitê de Governança e Gestão Estratégica (normal, tamanho 11, centralizado)
    p2 = header_retrato.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run('Relatório Técnico ao Comitê de Governança e Gestão Estratégica')
    run2.font.size = Pt(11)
    run2.font.bold = False
    run2.font.name = Config.FONTE_PADRAO
    p2.paragraph_format.space_after = Pt(6)
    
    # === RODAPÉ DA PRIMEIRA PÁGINA (RETRATO) ===
    footer_retrato = section_retrato.footer
    
    # Limpar rodapé padrão
    for paragraph in list(footer_retrato.paragraphs):
        p_element = paragraph._element
        p_element.getparent().remove(p_element)
    
    # Linha 1: Assessoria Técnica e Jurídica ao Planejamento e à Gestão Institucional - ASPLAG
    p_footer1 = footer_retrato.add_paragraph()
    p_footer1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_footer1 = p_footer1.add_run('Assessoria Técnica e Jurídica ao Planejamento e à Gestão Institucional - ASPLAG')
    run_footer1.font.size = Pt(9)
    run_footer1.font.name = Config.FONTE_PADRAO
    p_footer1.paragraph_format.space_after = Pt(0)
    
    # Linha 2: Diretoria Executiva de Planejamento Orçamentário e Qualidade na Gestão Institucional - DEPLAG
    p_footer2 = footer_retrato.add_paragraph()
    p_footer2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_footer2 = p_footer2.add_run('Diretoria Executiva de Planejamento Orçamentário e Qualidade na Gestão Institucional - DEPLAG')
    run_footer2.font.size = Pt(9)
    run_footer2.font.name = Config.FONTE_PADRAO
    p_footer2.paragraph_format.space_after = Pt(0)
    
    # === RETORNAR DOCUMENTO (segunda seção paisagem será criada depois) ===
    return doc


def criar_secao_paisagem_inicial(doc, superintendencia='Presidência'):
    """Cria segunda seção em paisagem para as tabelas de metas"""
    
    # Adicionar quebra de seção para paisagem
    section_paisagem = doc.add_section(WD_SECTION.NEW_PAGE)
    
    # Configurar orientação paisagem
    section_paisagem.orientation = WD_ORIENT.LANDSCAPE
    
    # Tamanho do papel A4 em paisagem
    section_paisagem.page_width = Cm(29.7)
    section_paisagem.page_height = Cm(21.0)
    
    # Margens
    section_paisagem.top_margin = Cm(2.5)
    section_paisagem.bottom_margin = Cm(2.5)
    section_paisagem.left_margin = Cm(2.0)
    section_paisagem.right_margin = Cm(1.5)
    
    # Medianiz (gutter)
    section_paisagem.gutter = Cm(0)
    
    # Cabeçalho e rodapé
    section_paisagem.header_distance = Cm(1.05)
    section_paisagem.footer_distance = Cm(1.05)
    
    # Desvincular cabeçalho da seção anterior
    header = section_paisagem.header
    header.is_linked_to_previous = False
    
    # Limpar cabeçalho padrão
    for paragraph in list(header.paragraphs):
        p_element = paragraph._element
        p_element.getparent().remove(p_element)
    
    # Título principal - negrito, 12pt, centralizado
    titulo = header.add_paragraph('Resultados do Monitoramento de Metas Estratégicas 2025')
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    titulo_format = titulo.runs[0]
    titulo_format.font.size = Pt(12)
    titulo_format.font.bold = True
    titulo_format.font.name = Config.FONTE_PADRAO
    titulo.paragraph_format.space_after = Pt(0)
    titulo.paragraph_format.space_before = Pt(0)
    
    # Subtítulo - normal, 11pt, centralizado
    subtitulo = header.add_paragraph('Relatório Técnico ao Comitê de Governança e Gestão Estratégica')
    subtitulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitulo_format = subtitulo.runs[0]
    subtitulo_format.font.size = Pt(11)
    subtitulo_format.font.name = Config.FONTE_PADRAO
    subtitulo.paragraph_format.space_after = Pt(0)
    
    # Órgão - negrito, 12pt, laranja, centralizado (DINÂMICO)
    orgao = header.add_paragraph(superintendencia)
    orgao.alignment = WD_ALIGN_PARAGRAPH.CENTER
    orgao_format = orgao.runs[0]
    orgao_format.font.size = Pt(12)
    orgao_format.font.bold = True
    orgao_format.font.name = Config.FONTE_PADRAO
    orgao_format.font.color.rgb = RGBColor(227, 108, 10)
    orgao.paragraph_format.space_after = Pt(6)
    
    # Adicionar conteúdo no rodapé
    footer = section_paisagem.footer
    footer.is_linked_to_previous = False
    
    # Limpar rodapé padrão
    for paragraph in list(footer.paragraphs):
        p_element = paragraph._element
        p_element.getparent().remove(p_element)
    
    # Primeira linha do rodapé
    linha1 = footer.add_paragraph('Assessoria Técnica e Jurídica ao Planejamento e à Gestão Institucional - ASPLAG')
    linha1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    linha1_format = linha1.runs[0]
    linha1_format.font.size = Pt(9)
    linha1_format.font.name = Config.FONTE_PADRAO
    linha1.paragraph_format.space_after = Pt(0)
    
    # Segunda linha do rodapé
    linha2 = footer.add_paragraph('Diretoria Executiva de Planejamento Orçamentário e Qualidade na Gestão Institucional - DEPLAG')
    linha2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    linha2_format = linha2.runs[0]
    linha2_format.font.size = Pt(9)
    linha2_format.font.name = Config.FONTE_PADRAO
    
    return section_paisagem
