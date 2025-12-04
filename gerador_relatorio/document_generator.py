"""
Módulo para geração do documento final integrando templates e dados
Responsável por criar o relatório Word combinando estrutura, conteúdo e dados
"""

from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import Dict, List, Any

from .config import Config
from .document_builder import criar_documento, criar_secao_paisagem_inicial
from .table_historico import adicionar_tabela_historica, adicionar_tabela_macrodesafio
from .table_cnj import adicionar_tabela_metas_nacionais
from .table_monitoramento import adicionar_tabela_resultado_monitoramento
from .table_superintendencia import (
    adicionar_nova_secao_superintendencia,
    adicionar_secao_macrodesafio
)


class DocumentGenerator:
    """Gerador de documento integrando templates e dados"""
    
    def __init__(self, template_data: Dict, df, grupos_super: Dict, variaveis: Dict):
        """
        Inicializa o gerador
        
        Args:
            template_data: Dados processados dos templates (estrutura + conteúdo)
            df: DataFrame com dados das metas
            grupos_super: Dicionário com grupos por superintendência
            variaveis: Dicionário com variáveis calculadas
        """
        self.estrutura = template_data['estrutura']
        self.conteudo = template_data['conteudo']
        self.df = df
        self.grupos_super = grupos_super
        self.variaveis = variaveis
        self.doc = None
        self.secao_paisagem_criada = False
    
    def gerar_documento(self):
        """
        Gera documento final integrando templates e dados
        
        Returns:
            Documento Word gerado
        """
        # Criar documento base (primeira página em retrato)
        print("📝 Criando documento base...")
        primeira_super = list(self.grupos_super.keys())[0] if self.grupos_super else 'Presidência'
        self.doc = criar_documento(primeira_super)
        
        # Adicionar sumário
        print("📑 Gerando sumário...")
        self._adicionar_sumario()
        
        # Adicionar sumário detalhado das metas
        print("📋 Gerando sumário detalhado das metas...")
        self._adicionar_sumario_metas()
        
        # Adicionar quebra de página após o sumário
        self._adicionar_quebra_pagina()
        
        # Iterar pela estrutura do sumário
        print("\n✍️  Processando estrutura do template...")
        total = len(self.estrutura)
        
        for idx, item in enumerate(self.estrutura):
            chave = item['chave']
            level = item['level']
            texto = item['texto']
            prefixo = item['prefixo']
            
            print(f"   [{idx+1}/{total}] {chave}")
            
            # Adicionar título formatado
            self._adicionar_titulo(texto, prefixo, level)
            
            # Buscar e processar conteúdo correspondente
            if chave in self.conteudo:
                self._processar_conteudo(chave)
            else:
                # Título sem conteúdo (apenas título fica no doc)
                pass
        
        print("\n✅ Documento gerado com sucesso!")
        return self.doc
    
    def _adicionar_titulo(self, texto: str, prefixo: str, level: int):
        """
        Adiciona título formatado ao documento
        
        Args:
            texto: Texto do título
            prefixo: Prefixo numérico (1, 1.1, etc.)
            level: Nível hierárquico
        """
        para = self.doc.add_paragraph()
        
        # Formatar título baseado no level
        if level == 1:
            # Nível 1: Adicionar espaço após número
            texto_completo = f"{prefixo}. {texto}"
            run = para.add_run(texto_completo)
            run.font.size = Pt(14)
            run.font.bold = True
            run.font.color.rgb = RGBColor(227, 108, 10)  # Laranja
            para.paragraph_format.space_before = Pt(18)
            para.paragraph_format.space_after = Pt(12)
            para.paragraph_format.left_indent = Cm(0.5)  # Recuo de 0,5cm
        elif level == 2:
            texto_completo = f"{prefixo} {texto}"
            run = para.add_run(texto_completo)
            run.font.size = Pt(14)  # Mesmo tamanho do nível 1
            run.font.bold = True
            run.font.color.rgb = RGBColor(227, 108, 10)  # Mesma cor laranja dos títulos
            para.paragraph_format.space_before = Pt(12)
            para.paragraph_format.space_after = Pt(6)
            para.paragraph_format.left_indent = Cm(0.75)  # Recuo maior que nível 1
        else:
            texto_completo = f"{prefixo} {texto}"
            run = para.add_run(texto_completo)
            run.font.size = Pt(11)
            run.font.bold = True
            para.paragraph_format.space_before = Pt(6)
            para.paragraph_format.space_after = Pt(3)
        
        run.font.name = Config.FONTE_PADRAO
    
    def _processar_conteudo(self, chave: str):
        """
        Processa todos os blocos de conteúdo de uma seção
        
        Args:
            chave: Chave do título (ex: "1 INTRODUÇÃO")
        """
        blocos = self.conteudo[chave]
        
        for bloco in blocos:
            tipo = bloco['tipo']
            
            if tipo == 'PARAGRAFO':
                self._adicionar_paragrafo(bloco)
            
            elif tipo == 'PARAGRAFO_COM_VARIAVEL':
                self._adicionar_paragrafo_variavel(bloco)
            
            elif tipo == 'TEXTO_DESTAQUE':
                self._adicionar_texto_destaque(bloco)
            
            elif tipo == 'LISTA_NUMERADA':
                self._adicionar_lista_numerada(bloco)
            
            elif tipo == 'LISTA_MARCADORES':
                self._adicionar_lista_marcadores(bloco)
            
            elif tipo == 'TABELA_HISTORICA':
                adicionar_tabela_historica(self.doc)
            
            elif tipo == 'TABELA_MACRODESAFIO':
                adicionar_tabela_macrodesafio(self.doc)
            
            elif tipo == 'TABELA_MONITORAMENTO':
                adicionar_tabela_resultado_monitoramento(self.doc)
            
            elif tipo == 'TABELA_CNJ':
                adicionar_tabela_metas_nacionais(self.doc)
            
            elif tipo == 'SECAO_SUPERINTENDENCIAS':
                self._gerar_secoes_superintendencias()
            
            elif tipo == 'QUEBRA_PAGINA':
                self._adicionar_quebra_pagina()
    
    def _adicionar_paragrafo(self, bloco: Dict):
        """Adiciona parágrafo normal"""
        para = self.doc.add_paragraph()
        self._processar_texto_com_formatacao_inline(para, bloco['texto'])
        
        # Aplicar alinhamento
        alinhamento = bloco.get('alinhamento', 'JUSTIFY')
        if alinhamento == 'CENTER':
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif alinhamento == 'RIGHT':
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif alinhamento == 'JUSTIFY':
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        else:
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        para.paragraph_format.space_after = Pt(6)
        para.paragraph_format.line_spacing = 1.5
        para.paragraph_format.left_indent = Cm(0.5)  # Recuo do bloco
        para.paragraph_format.first_line_indent = Cm(1.27)  # Recuo de primeira linha
    
    def _adicionar_paragrafo_variavel(self, bloco: Dict):
        """Adiciona parágrafo substituindo variáveis"""
        texto = bloco['texto']
        
        # Substituir cada variável encontrada
        for marcador in bloco['variaveis']:
            valor = self._obter_valor_variavel(marcador)
            texto = texto.replace(marcador, str(valor))
        
        # Criar parágrafo com texto substituído
        para = self.doc.add_paragraph()
        self._processar_texto_com_formatacao_inline(para, texto)
        
        # Aplicar alinhamento
        alinhamento = bloco.get('alinhamento', 'JUSTIFY')
        if alinhamento == 'CENTER':
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif alinhamento == 'RIGHT':
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif alinhamento == 'JUSTIFY':
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        else:
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        para.paragraph_format.space_after = Pt(6)
        para.paragraph_format.line_spacing = 1.5
        para.paragraph_format.left_indent = Cm(0.5)  # Recuo do bloco
        para.paragraph_format.first_line_indent = Cm(1.27)  # Recuo de primeira linha
    
    def _processar_texto_com_formatacao_inline(self, para, texto: str):
        """Processa texto com marcadores inline **texto** para negrito
        
        Args:
            para: Parágrafo do documento
            texto: Texto com possíveis marcadores **texto**
        """
        import re
        
        # Dividir texto em partes: normal e negrito
        # Padrão: captura texto fora de ** e dentro de **
        partes = re.split(r'(\*\*.*?\*\*)', texto)
        
        for parte in partes:
            if not parte:  # Ignorar strings vazias
                continue
                
            if parte.startswith('**') and parte.endswith('**'):
                # Texto em negrito (remover os **)
                texto_negrito = parte[2:-2]
                run = para.add_run(texto_negrito)
                run.font.bold = True
                run.font.size = Pt(12)
                run.font.name = Config.FONTE_PADRAO
            else:
                # Texto normal
                run = para.add_run(parte)
                run.font.size = Pt(12)
                run.font.name = Config.FONTE_PADRAO
    
    def _adicionar_texto_destaque(self, bloco: Dict):
        """Adiciona texto destacado (apenas negrito)"""
        para = self.doc.add_paragraph()
        run = para.add_run(bloco['texto'])
        run.font.bold = True
        run.font.size = Pt(12)
        run.font.name = Config.FONTE_PADRAO
        
        para.paragraph_format.space_after = Pt(6)
        para.paragraph_format.line_spacing = 1.5
        para.paragraph_format.left_indent = Cm(0.5)  # Recuo de 0,5cm
    
    def _adicionar_lista_numerada(self, bloco: Dict):
        """Adiciona lista numerada"""
        for idx, item in enumerate(bloco['itens'], 1):
            para = self.doc.add_paragraph(f"{idx}. {item}")
            para.paragraph_format.left_indent = Cm(0.5)
            para.paragraph_format.space_after = Pt(3)
            
            for run in para.runs:
                run.font.size = Pt(12)
                run.font.name = Config.FONTE_PADRAO
    
    def _adicionar_lista_marcadores(self, bloco: Dict):
        """Adiciona lista com marcadores"""
        for item in bloco['itens']:
            para = self.doc.add_paragraph(item, style='List Bullet')
            para.paragraph_format.space_after = Pt(3)
            
            for run in para.runs:
                run.font.size = Pt(12)
                run.font.name = Config.FONTE_PADRAO
    
    def _adicionar_quebra_pagina(self):
        """Adiciona quebra de página sem criar linha vazia"""
        from docx.enum.text import WD_BREAK
        from docx.shared import Pt
        
        # Criar parágrafo com espaçamento zero
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        
        # Adicionar quebra de página
        p.add_run().add_break(WD_BREAK.PAGE)
    
    def _gerar_secoes_superintendencias(self):
        """Gera seções detalhadas por superintendência (tabelas dinâmicas em paisagem)"""
        print("\n📋 Gerando seções de superintendências...")
        
        # Criar seção paisagem se ainda não foi criada
        if not self.secao_paisagem_criada:
            primeira_super = list(self.grupos_super.keys())[0] if self.grupos_super else 'Presidência'
            criar_secao_paisagem_inicial(self.doc, primeira_super)
            self.secao_paisagem_criada = True
        
        primeira_super = True
        
        for superintendencia, grupos_macro in self.grupos_super.items():
            print(f"   → {superintendencia}")
            
            if not primeira_super:
                adicionar_nova_secao_superintendencia(self.doc, superintendencia, False)
            
            primeira_secao = True
            for macrodesafio, df_grupo in grupos_macro:
                adicionar_secao_macrodesafio(self.doc, macrodesafio, df_grupo, primeira_secao)
                primeira_secao = False
            
            primeira_super = False
    
    def _obter_valor_variavel(self, marcador: str) -> Any:
        """
        Obtém valor de uma variável pelo marcador
        
        Args:
            marcador: Marcador da variável (ex: [NUMERO_METAS_2025])
            
        Returns:
            Valor da variável ou '???' se não encontrada
        """
        from .content_mapper import ContentMapper
        
        # Mapear marcador para nome de variável
        nome_var = ContentMapper.MARCADORES_VARIAVEIS.get(marcador)
        
        if nome_var and nome_var in self.variaveis:
            valor = self.variaveis[nome_var]
            
            # Formatar decimais
            if isinstance(valor, float):
                return f"{valor:.1f}"
            
            return valor
        
        return '???'
    
    def _adicionar_sumario(self):
        """Adiciona sumário ao documento baseado na estrutura extraída"""
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        
        # Título "SUMÁRIO"
        titulo_sumario = self.doc.add_paragraph()
        titulo_sumario.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_titulo = titulo_sumario.add_run('SUMÁRIO')
        run_titulo.font.size = Pt(14)
        run_titulo.font.bold = True
        run_titulo.font.name = Config.FONTE_PADRAO
        titulo_sumario.paragraph_format.space_after = Pt(12)
        
        # Adicionar itens do sumário
        for item in self.estrutura:
            prefixo = item['prefixo']
            texto = item['texto']
            level = item['level']
            
            # Criar linha do sumário
            para = self.doc.add_paragraph()
            
            # Configurar recuo baseado no nível
            if level == 1:
                para.paragraph_format.left_indent = Cm(0)
            elif level == 2:
                para.paragraph_format.left_indent = Cm(0.5)
            else:
                para.paragraph_format.left_indent = Cm(1.0)
            
            # Adicionar texto do sumário (prefixo + título)
            texto_sumario = f"{prefixo}   {texto}"
            run_texto = para.add_run(texto_sumario)
            run_texto.font.size = Pt(10)
            run_texto.font.name = Config.FONTE_PADRAO
            
            # Negrito apenas para nível 1
            if level == 1:
                run_texto.font.bold = True
            
            # Adicionar TAB com pontos de preenchimento
            # Criar elemento de tabulação
            pPr = para._element.get_or_add_pPr()
            tabs = pPr.find(qn('w:tabs'))
            if tabs is None:
                tabs = OxmlElement('w:tabs')
                pPr.append(tabs)
            
            # Adicionar tab stop na direita (16cm da margem esquerda) com leader dots
            tab = OxmlElement('w:tab')
            tab.set(qn('w:val'), 'right')
            tab.set(qn('w:pos'), '9070')  # 16cm em twips (16 * 567 = 9072)
            tab.set(qn('w:leader'), 'dot')  # Pontos de preenchimento
            tabs.append(tab)
            
            # Adicionar TAB e número de página
            run_tab = para.add_run('\t')
            run_numero = para.add_run('1')  # Placeholder - Word atualizará
            run_numero.font.size = Pt(10)
            run_numero.font.name = Config.FONTE_PADRAO
            
            para.paragraph_format.space_after = Pt(0)
    
    def _adicionar_sumario_metas(self):
        """Adiciona sumário detalhado das metas por superintendência"""
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        import pandas as pd
        
        # Espaçamento antes do sumário de metas
        self.doc.add_paragraph().paragraph_format.space_after = Pt(12)
        
        # Iterar por cada superintendência
        for super_nome in sorted(self.grupos_super.keys()):
            lista_metas = self.grupos_super[super_nome]
            
            if not lista_metas:
                continue
            
            # Converter lista para DataFrame se necessário
            if isinstance(lista_metas, list):
                df_super = pd.DataFrame(lista_metas)
            else:
                df_super = lista_metas
            
            # Título da superintendência em negrito
            para_super = self.doc.add_paragraph()
            run_super = para_super.add_run(super_nome.upper())
            run_super.font.size = Pt(10)
            run_super.font.bold = True
            run_super.font.name = Config.FONTE_PADRAO
            para_super.paragraph_format.space_before = Pt(6)
            para_super.paragraph_format.space_after = Pt(3)
            
            # Listar cada meta da superintendência
            for _, row in df_super.iterrows():
                # Usar diretamente a coluna META que já contém o texto completo
                texto_meta = row.get('META', 'N/A')
                
                # Criar linha da meta
                para_meta = self.doc.add_paragraph()
                para_meta.paragraph_format.left_indent = Cm(0.5)
                
                # Adicionar texto da meta
                run_meta = para_meta.add_run(texto_meta)
                run_meta.font.size = Pt(10)
                run_meta.font.name = Config.FONTE_PADRAO
                
                # Configurar tabulação com pontos
                pPr = para_meta._element.get_or_add_pPr()
                tabs = pPr.find(qn('w:tabs'))
                if tabs is None:
                    tabs = OxmlElement('w:tabs')
                    pPr.append(tabs)
                
                tab = OxmlElement('w:tab')
                tab.set(qn('w:val'), 'right')
                tab.set(qn('w:pos'), '9070')
                tab.set(qn('w:leader'), 'dot')
                tabs.append(tab)
                
                # Adicionar TAB e número de página
                para_meta.add_run('\t')
                run_pag = para_meta.add_run('1')  # Placeholder
                run_pag.font.size = Pt(10)
                run_pag.font.name = Config.FONTE_PADRAO
                
                para_meta.paragraph_format.space_after = Pt(0)
