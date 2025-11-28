"""
GERADOR DE RELATÓRIO DE METAS ESTRATÉGICAS - TJMG
Versão: 1.1 (Corrigido)
"""

import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
import os

# ============================================
# CONFIGURAÇÕES
# ============================================

class Config:
    """Configurações do relatório"""
    
    # Arquivo de entrada
    ARQUIVO_EXCEL = 'exports/teste_integração.xlsx'  # Nome do seu arquivo Excel
    NOME_ABA = None  # Nome da aba (ou deixe None para a primeira)
    
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


# ============================================
# FUNÇÕES AUXILIARES
# ============================================

def criar_pasta_saida():
    """Cria pasta para salvar relatórios se não existir"""
    if not os.path.exists(Config.PASTA_SAIDA):
        os.makedirs(Config.PASTA_SAIDA)
        print(f"✅ Pasta '{Config.PASTA_SAIDA}' criada.")


def carregar_dados():
    """Carrega dados do Excel"""
    print(f"📂 Carregando dados de '{Config.ARQUIVO_EXCEL}'...")
    
    try:
        # Ler Excel
        if Config.NOME_ABA:
            df = pd.read_excel(Config.ARQUIVO_EXCEL, sheet_name=Config.NOME_ABA)
        else:
            df = pd.read_excel(Config.ARQUIVO_EXCEL)
        
        # Remover linhas completamente vazias
        df = df.dropna(how='all')
        
        print(f"✅ {len(df)} registros carregados.")
        return df
    
    except FileNotFoundError:
        print(f"❌ ERRO: Arquivo '{Config.ARQUIVO_EXCEL}' não encontrado!")
        print(f"   Certifique-se de que o arquivo está na mesma pasta que este script.")
        return None
    except Exception as e:
        print(f"❌ ERRO ao carregar dados: {e}")
        return None


def agrupar_por_macrodesafio(df):
    """Agrupa dados por Macrodesafio"""
    print("📊 Agrupando dados por Macrodesafio...")
    
    col_macro = Config.COLUNAS['MACRODESAFIO']
    grupos = df.groupby(col_macro, sort=True)
    
    print(f"✅ {len(grupos)} Macrodesafios encontrados.")
    return grupos


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


def set_cell_background(cell, rgb_tuple):
    """Define cor de fundo de uma célula da tabela"""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), rgb_to_hex(rgb_tuple))
    cell._element.get_or_add_tcPr().append(shading_elm)


def set_cell_border(cell, **kwargs):
    """Define bordas para células da tabela"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    # Remover bordas antigas se existirem
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is not None:
        tcPr.remove(tcBorders)
    
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        if kwargs.get(edge):
            edge_el = OxmlElement(f'w:{edge}')
            edge_el.set(qn('w:val'), 'single')
            edge_el.set(qn('w:sz'), '2')  # 2 = 1/4pt (sz é em oitavos de ponto)
            edge_el.set(qn('w:space'), '0')
            edge_el.set(qn('w:color'), rgb_to_hex(Config.CORES['BORDA_TABELA']))
            tcBorders.append(edge_el)
    
    tcPr.append(tcBorders)


def set_paragraph_background(paragraph, rgb_tuple):
    """Define cor de fundo de um parágrafo"""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), rgb_to_hex(rgb_tuple))
    paragraph._element.get_or_add_pPr().append(shading_elm)


# ============================================
# FUNÇÕES DE CRIAÇÃO DO DOCUMENTO
# ============================================

def criar_documento():
    """Cria documento Word base"""
    doc = Document()
    
    # Configurar margens e orientação
    sections = doc.sections
    for section in sections:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Inches(11)
        section.page_height = Inches(8.5)
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
    
    return doc


def adicionar_cabecalho_relatorio(doc):
    """Adiciona cabeçalho do relatório"""
    
    # Título principal
    titulo = doc.add_heading('Resultados do Monitoramento de Metas Estratégicas 2024', level=0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    titulo_format = titulo.runs[0]
    titulo_format.font.size = Pt(Config.TAMANHO_TITULO)
    titulo_format.font.color.rgb = RGBColor(0, 0, 0)
    titulo_format.font.name = Config.FONTE_PADRAO
    
    # Subtítulo
    subtitulo = doc.add_paragraph('Relatório Técnico ao Comitê de Governança e Gestão Estratégica')
    subtitulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitulo_format = subtitulo.runs[0]
    subtitulo_format.font.size = Pt(Config.TAMANHO_SUBTITULO)
    subtitulo_format.font.name = Config.FONTE_PADRAO
    subtitulo_format.italic = True
    
    # Órgão
    orgao = doc.add_paragraph('Presidência')
    orgao.alignment = WD_ALIGN_PARAGRAPH.CENTER
    orgao_format = orgao.runs[0]
    orgao_format.font.size = Pt(Config.TAMANHO_SUBTITULO)
    orgao_format.font.name = Config.FONTE_PADRAO
    orgao_format.font.color.rgb = RGBColor(*Config.CORES['CABECALHO_MACRO'])
    orgao_format.bold = True
    
    # Data de geração
    data_hora = datetime.now().strftime('%d/%m/%Y às %H:%M')
    data = doc.add_paragraph(f'Gerado em: {data_hora}')
    data.alignment = WD_ALIGN_PARAGRAPH.CENTER
    data_format = data.runs[0]
    data_format.font.size = Pt(9)
    data_format.font.name = Config.FONTE_PADRAO
    data_format.italic = True
    
    # Linha separadora
    doc.add_paragraph('_' * 80)
    doc.add_paragraph()


def adicionar_secao_macrodesafio(doc, macrodesafio, df_grupo, primeira_secao=False):
    """Adiciona seção de um Macrodesafio"""
    
    # Quebra de página entre macrodesafios (exceto no primeiro)
    if not primeira_secao:
        doc.add_page_break()
    
    # Título do Macrodesafio usando tabela para controlar largura
    titulo_tabela = doc.add_table(rows=1, cols=1)
    titulo_tabela.style = None
    titulo_cell = titulo_tabela.rows[0].cells[0]
    titulo_cell.text = macrodesafio.upper()
    
    # Formatação do título
    for paragraph in titulo_cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_before = Pt(5)
        paragraph.paragraph_format.space_after = Pt(5)
        for run in paragraph.runs:
            run.font.size = Pt(10)
            run.font.name = Config.FONTE_PADRAO
            run.font.color.rgb = RGBColor(*Config.CORES['TEXTO_BRANCO'])
            run.font.bold = True
    
    # Cor de fundo e alinhamento vertical
    set_cell_background(titulo_cell, Config.CORES['CABECALHO_MACRO'])
    from docx.enum.table import WD_ALIGN_VERTICAL
    titulo_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    
    # Remover bordas
    set_cell_border(titulo_cell, top=False, bottom=False, left=False, right=False)
    
    # Definir largura igual à soma das colunas da tabela de dados
    # Calculando a soma total das larguras (será calculada dinamicamente)
    largura_total_tabela = Cm(9.09 + 2.28 + 2.1 + 1.97 + 2.18 + 0.72)  # Sem indicador
    # Adicionar espaço para indicador (será ajustado automaticamente)
    titulo_tabela.rows[0].cells[0].width = largura_total_tabela + Cm(7)  # Aproximação para indicador
    
    # Processar cada indicador individualmente
    for idx, (_, row) in enumerate(df_grupo.iterrows()):
        # Criar tabela para este indicador
        adicionar_tabela_indicador(doc, row)
        
        # Adicionar informação complementar se existir
        info_complementar = row.get(Config.COLUNAS['INFO_COMPLEMENTAR'], '')
        if pd.notna(info_complementar) and str(info_complementar).strip() != '':
            situacao = doc.add_paragraph()
            situacao_run = situacao.add_run(f'Situação: ')
            situacao_run.font.bold = True
            situacao_run.font.size = Pt(10)
            situacao_run.font.name = Config.FONTE_PADRAO
            
            texto_run = situacao.add_run(str(info_complementar))
            texto_run.font.size = Pt(10)
            texto_run.font.name = Config.FONTE_PADRAO
            situacao.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        doc.add_paragraph()  # Espaçamento entre indicadores


def adicionar_tabela_indicador(doc, row):
    """Adiciona tabela com um único indicador"""
    
    # Criar tabela (1 cabeçalho + 1 linha de dados)
    num_linhas = 2
    num_colunas = 7  # Incluindo coluna de indicador visual
    tabela = doc.add_table(rows=num_linhas, cols=num_colunas)
    tabela.style = None  # Remover estilo padrão para controle total
    
    # Cabeçalhos
    headers = ['INDICADOR', 'META', 'UNIDADE RESPONSÁVEL', 'POLARIDADE', 'RESULTADO APURADO', 'INICIATIVA(S)', '●']
    header_cells = tabela.rows[0].cells
    
    for i, header in enumerate(headers):
        cell = header_cells[i]
        cell.text = header
        
        # Formatação do cabeçalho
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER  # Centro horizontal
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(9)
                run.font.name = Config.FONTE_PADRAO
                run.font.color.rgb = RGBColor(*Config.CORES['TEXTO_PRETO'])
        
        from docx.enum.table import WD_ALIGN_VERTICAL
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER  # Centro vertical
        
        # Cor de fundo
        set_cell_background(cell, Config.CORES['CABECALHO_TABELA'])
        
        # Adicionar bordas brancas
        set_cell_border(cell, top=True, bottom=True, left=True, right=True)
    
    # Preencher dados (apenas uma linha)
    cells = tabela.rows[1].cells
    
    # Dados
    dados = [
        formatar_valor(row.get(Config.COLUNAS['INDICADOR'], '-')),
        formatar_valor(row.get(Config.COLUNAS['METAKEY'], '-')),
        formatar_valor(row.get(Config.COLUNAS['UNIDADE_GESTORA'], '-')),
        formatar_valor(row.get(Config.COLUNAS['POLARIDADE'], '-')),
        formatar_valor(row.get(Config.COLUNAS['VALOR_APURADO'], '-')),
        formatar_valor(row.get(Config.COLUNAS['INICIATIVA'], '-')),
        '●'  # Indicador visual
    ]
    
    for i, dado in enumerate(dados):
        cell = cells[i]
        cell.text = dado
        
        # Formatação do texto
        for paragraph in cell.paragraphs:
            # Alinhamento
            if i in [3, 4, 6]:  # Polaridade, Resultado, Indicador visual
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            for run in paragraph.runs:
                run.font.size = Pt(Config.TAMANHO_TABELA)
                run.font.name = Config.FONTE_PADRAO
                
                # Destacar resultado apurado
                if i == 4:  # Coluna de Resultado Apurado
                    run.font.bold = True
                
                # Bolinha verde no indicador visual
                if i == 6:
                    run.font.color.rgb = RGBColor(0, 255, 0)  # Verde
                    run.font.size = Pt(16)
        
        # Cor de fundo para coluna Iniciativa
        if i == 5:
            set_cell_background(cell, Config.CORES['INICIATIVA_FUNDO'])
            # Texto preto no fundo laranja claro
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.color.rgb = RGBColor(*Config.CORES['TEXTO_PRETO'])
        
        from docx.enum.table import WD_ALIGN_VERTICAL
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER  # Centro vertical
        
        # Adicionar bordas brancas
        set_cell_border(cell, top=True, bottom=True, left=True, right=True)
    
    # Ajustar largura das colunas (em centímetros)
    # Largura mínima de 1,3cm para todas
    larguras = [
        None,        # INDICADOR - não alterável (automática)
        Cm(9.09),    # META
        Cm(2.28),    # UNIDADE RESPONSÁVEL
        Cm(2.1),     # POLARIDADE
        Cm(1.97),    # RESULTADO APURADO
        Cm(2.18),    # INICIATIVA(S)
        Cm(0.72)     # ●
    ]
    
    # Calcular largura do Indicador baseado no espaço restante
    # Largura total da página em paisagem (11" - margens) ≈ 9" = 22.86 cm
    largura_total = Cm(25.4)  # Largura disponível aproximada
    largura_outras = sum([l for l in larguras if l is not None])
    largura_indicador = largura_total - largura_outras
    larguras[0] = max(largura_indicador, Cm(1.3))  # Mínimo de 1,3cm
    
    for row in tabela.rows:
        for idx, width in enumerate(larguras):
            if width:
                row.cells[idx].width = width


# ============================================
# FUNÇÃO PRINCIPAL
# ============================================

def gerar_relatorio():
    """Função principal para gerar o relatório"""
    
    print("\n" + "="*60)
    print("📊 GERADOR DE RELATÓRIO DE METAS ESTRATÉGICAS - TJMG")
    print("="*60 + "\n")
    
    # 1. Criar pasta de saída
    criar_pasta_saida()
    
    # 2. Carregar dados
    df = carregar_dados()
    if df is None:
        return
    
    # 3. Agrupar dados
    grupos = agrupar_por_macrodesafio(df)
    
    # 4. Criar documento
    print("📝 Criando documento Word...")
    doc = criar_documento()
    
    # 5. Adicionar cabeçalho do relatório
    adicionar_cabecalho_relatorio(doc)
    
    # 6. Adicionar cada Macrodesafio (APENAS PRIMEIRO PARA TESTE)
    print("✍️  Gerando seções do relatório...")
    for idx, (macrodesafio, df_grupo) in enumerate(grupos):
        print(f"   → {macrodesafio} ({len(df_grupo)} registros)")
        adicionar_secao_macrodesafio(doc, macrodesafio, df_grupo, primeira_secao=(idx==0))
        break  # TESTAR APENAS PRIMEIRO MACRODESAFIO
    
    # 7. Salvar documento
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome_arquivo = f"{Config.NOME_RELATORIO}_{timestamp}.docx"
    caminho_completo = os.path.join(Config.PASTA_SAIDA, nome_arquivo)
    
    doc.save(caminho_completo)
    
    print(f"\n✅ RELATÓRIO GERADO COM SUCESSO!")
    print(f"📁 Localização: {os.path.abspath(caminho_completo)}")
    print(f"📄 Total de Macrodesafios: {len(grupos)}")
    print(f"📊 Total de registros: {len(df)}")
    print("\n" + "="*60 + "\n")


# ============================================
# EXECUÇÃO
# ============================================

if __name__ == "__main__":
    try:
        gerar_relatorio()
    except KeyboardInterrupt:
        print("\n\n⚠️  Geração cancelada pelo usuário.")
    except Exception as e:
        print(f"\n❌ ERRO INESPERADO: {e}")
        import traceback
        traceback.print_exc()