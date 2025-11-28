"""
GERADOR DE RELATÓRIO DE METAS ESTRATÉGICAS - TJMG
Versão: 1.1 (Corrigido)
"""

import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
    ARQUIVO_EXCEL = 'teste_integração.xlsx'  # Nome do seu arquivo Excel
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
        'CABECALHO_MACRO': (255, 107, 53),      # Laranja
        'CABECALHO_TABELA': (78, 205, 196),     # Azul claro
        'TEXTO_BRANCO': (255, 255, 255),
        'BORDA_TABELA': (149, 165, 166)         # Cinza
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
    
    # Configurar margens
    sections = doc.sections
    for section in sections:
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
    
    # Título do Macrodesafio
    titulo = doc.add_paragraph()
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    titulo_run = titulo.add_run(macrodesafio.upper())
    titulo_run.font.size = Pt(14)
    titulo_run.font.name = Config.FONTE_PADRAO
    titulo_run.font.color.rgb = RGBColor(*Config.CORES['TEXTO_BRANCO'])
    titulo_run.bold = True
    
    # Cor de fundo do título
    set_paragraph_background(titulo, Config.CORES['CABECALHO_MACRO'])
    
    doc.add_paragraph()
    
    # Criar tabela
    adicionar_tabela_indicadores(doc, df_grupo)
    
    doc.add_paragraph()


def adicionar_tabela_indicadores(doc, df_grupo):
    """Adiciona tabela com indicadores"""
    
    # Criar tabela (1 cabeçalho + linhas de dados)
    num_linhas = len(df_grupo) + 1
    num_colunas = 6
    tabela = doc.add_table(rows=num_linhas, cols=num_colunas)
    tabela.style = 'Light Grid Accent 1'
    
    # Cabeçalhos
    headers = ['INDICADOR', 'META', 'UNIDADE RESPONSÁVEL', 'POLARIDADE', 'RESULTADO APURADO', 'INICIATIVA(S)']
    header_cells = tabela.rows[0].cells
    
    for i, header in enumerate(headers):
        cell = header_cells[i]
        cell.text = header
        
        # Formatação do cabeçalho
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(10)
                run.font.name = Config.FONTE_PADRAO
                run.font.color.rgb = RGBColor(*Config.CORES['TEXTO_BRANCO'])
        
        # Cor de fundo
        set_cell_background(cell, Config.CORES['CABECALHO_TABELA'])
    
    # Preencher dados
    for idx, (_, row) in enumerate(df_grupo.iterrows(), start=1):
        cells = tabela.rows[idx].cells
        
        # Dados
        dados = [
            formatar_valor(row.get(Config.COLUNAS['INDICADOR'], '-')),
            formatar_valor(row.get(Config.COLUNAS['METAKEY'], '-')),
            formatar_valor(row.get(Config.COLUNAS['UNIDADE_GESTORA'], '-')),
            formatar_valor(row.get(Config.COLUNAS['POLARIDADE'], '-')),
            formatar_valor(row.get(Config.COLUNAS['VALOR_APURADO'], '-')),
            formatar_valor(row.get(Config.COLUNAS['INICIATIVA'], '-'))
        ]
        
        for i, dado in enumerate(dados):
            cell = cells[i]
            cell.text = dado
            
            # Formatação do texto
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(Config.TAMANHO_TABELA)
                    run.font.name = Config.FONTE_PADRAO
                    
                    # Destacar resultado apurado
                    if i == 4:  # Coluna de Resultado Apurado
                        run.font.bold = True
    
    # Ajustar largura das colunas
    larguras = [Inches(2.2), Inches(1.8), Inches(1.2), Inches(1.0), Inches(1.0), Inches(1.0)]
    for row in tabela.rows:
        for idx, width in enumerate(larguras):
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
    
    # 6. Adicionar cada Macrodesafio
    print("✍️  Gerando seções do relatório...")
    for idx, (macrodesafio, df_grupo) in enumerate(grupos):
        print(f"   → {macrodesafio} ({len(df_grupo)} registros)")
        adicionar_secao_macrodesafio(doc, macrodesafio, df_grupo, primeira_secao=(idx==0))
    
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