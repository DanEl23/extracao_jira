# 📋 PROPOSTA DE IMPLEMENTAÇÃO - SISTEMA DE TEMPLATES DOCX

## 🎯 OBJETIVO

Adaptar a estratégia de leitura de arquivos `.docx` para o projeto de Extração Jira, permitindo:
1. **Definir estrutura do relatório** através de um arquivo de sumário
2. **Gerenciar conteúdo textual** através de um arquivo de conteúdo fonte
3. **Integrar dados extraídos** (tabelas dinâmicas) com textos estáticos
4. **Manter ordem e hierarquia** do documento conforme template

---

## 📊 ANÁLISE DO ESTADO ATUAL

### **Como Funciona Hoje:**

```
1. Extração de Dados
   ├─> extracao_cnj.py → exports/resultados_cnj.xlsx
   └─> extrator_unificado.py → exports/teste_integração.xlsx

2. Geração do Relatório (TOTALMENTE PROGRAMÁTICO)
   └─> gerador_relatorio.py
       ├─> Cria documento do zero
       ├─> Gera tabelas a partir dos dados
       ├─> Estrutura hardcoded no código
       └─> Output: relatorios_gerados/*.docx
```

### **Limitações Atuais:**

❌ Estrutura do documento está "presa" no código  
❌ Textos explicativos não existem (só tabelas)  
❌ Ordem das seções é fixa no código  
❌ Não há sumário navegável  
❌ Difícil adicionar/remover seções sem alterar código  

---

## 🎨 ARQUITETURA PROPOSTA

### **Nova Estrutura de Arquivos:**

```
Extracao_Jira/
├─> templates/                          # 🆕 Nova pasta
│   ├─> Sumario_Modelo.docx            # Define estrutura hierárquica
│   ├─> Conteudo_Fonte.docx            # Contém textos e marcadores
│   └─> readme_templates.md            # Guia de uso dos templates
│
├─> gerador_relatorio/
│   ├─> config.py
│   ├─> data_loader.py
│   ├─> document_builder.py
│   ├─> template_reader.py             # 🆕 Leitor de templates
│   ├─> content_mapper.py              # 🆕 Mapeador de conteúdo
│   ├─> structure_parser.py            # 🆕 Parser de estrutura
│   └─> ... (módulos existentes)
│
└─> gerador_relatorio.py                # Fluxo principal atualizado
```

---

## 📝 FASE 1: DEFINIÇÃO DOS TEMPLATES

### **1.1 - Template de Sumário (Sumario_Modelo.docx)**

**Estrutura Proposta:**

```
SUMÁRIO

1 INTRODUÇÃO
1.1 Contextualização
1.2 Objetivos do Relatório

2 METAS APROVADAS
2.1 Total de Metas Estratégicas
2.2 Distribuição por Macrodesafio

3 RESULTADO DO MONITORAMENTO
3.1 Indicadores de Cumprimento
3.2 Análise por Faixa de Desempenho

4 METAS NACIONAIS DO CNJ
4.1 Visão Geral
4.2 Detalhamento por Meta

5 SUPERINTENDÊNCIAS
5.1 Presidência
5.2 Superintendência Administrativa
5.3 Superintendência de Gestão Estratégica
... (demais superintendências)

6 CONSIDERAÇÕES FINAIS
```

**Características:**
- Numeração hierárquica (1, 1.1, 1.2, etc.)
- Sem páginas (serão geradas automaticamente)
- Títulos exatos que serão mapeados no conteúdo

---

### **1.2 - Template de Conteúdo (Conteudo_Fonte.docx)**

**Estrutura com Marcadores Especiais:**

```docx
1 INTRODUÇÃO

Este relatório apresenta os resultados do monitoramento das metas estratégicas do 
TJMG para o ano de 2025, conforme deliberações do Comitê de Governança e Gestão 
Estratégica.

1.1 Contextualização

O processo de monitoramento tem como objetivo acompanhar o desempenho institucional 
e fornecer subsídios para tomada de decisão pelos gestores.

1.2 Objetivos do Relatório

#Os principais objetivos deste documento são:
[INICIAR_LISTA_MARCADORES]
Apresentar o panorama geral das metas aprovadas
Analisar os resultados apurados no período
Identificar oportunidades de melhoria
Subsidiar decisões estratégicas
[FINALIZAR_LISTA_MARCADORES]

2 METAS APROVADAS

[INSERIR_TABELA_HISTORICA]

O quadro acima demonstra a evolução histórica das metas estratégicas aprovadas pelo 
TJMG nos últimos 4 anos.

2.1 Total de Metas Estratégicas

Conforme apresentado, foram aprovadas [NUMERO_METAS_2025] metas estratégicas para 
o ano de 2025, distribuídas em [NUMERO_MACRODESAFIOS] macrodesafios institucionais.

2.2 Distribuição por Macrodesafio

[INSERIR_TABELA_MACRODESAFIO]

3 RESULTADO DO MONITORAMENTO

A seguir são apresentados os resultados consolidados do monitoramento realizado no 
período de referência.

3.1 Indicadores de Cumprimento

[INSERIR_TABELA_MONITORAMENTO]

3.2 Análise por Faixa de Desempenho

#Destaque: [PERCENTUAL_VERDE]% das metas atingiram resultado satisfatório (faixa verde).

4 METAS NACIONAIS DO CNJ

[INSERIR_TABELA_CNJ]

As metas nacionais estabelecidas pelo Conselho Nacional de Justiça representam 
compromissos do Poder Judiciário brasileiro com a sociedade.

5 SUPERINTENDÊNCIAS

A partir desta seção são apresentados os resultados detalhados por superintendência 
e macrodesafio.

[INICIAR_SECAO_SUPERINTENDENCIAS]
```

---

## 🔧 FASE 2: IMPLEMENTAÇÃO DOS MÓDULOS

### **2.1 - Módulo: `structure_parser.py`**

**Responsabilidade:** Extrair estrutura hierárquica do sumário

```python
"""
Módulo para parsing da estrutura do sumário
"""

import re
from docx import Document
from typing import List, Dict


class StructureParser:
    """Parser para extrair estrutura hierárquica do sumário"""
    
    # Regex para detectar títulos numerados
    PATTERN_SUMARIO = r'^\s*(\d+(?:\.\d+)*\.?)\s+(.+?)(?:\s*\.{2,}\s*\d+\s*)?$'
    
    def __init__(self, template_path: str):
        self.template_path = template_path
        self.estrutura = []
    
    def extrair_estrutura(self) -> List[Dict]:
        """
        Extrai estrutura hierárquica do documento de sumário
        
        Returns:
            Lista de dicionários com estrutura:
            [
                {
                    "tipo": "TITULO",
                    "level": 1,
                    "prefixo": "1",
                    "texto": "INTRODUÇÃO",
                    "chave": "1 INTRODUÇÃO"
                },
                ...
            ]
        """
        doc = Document(self.template_path)
        estrutura = []
        
        for para in doc.paragraphs:
            texto_limpo = self._limpar_texto(para.text)
            
            if not texto_limpo or texto_limpo == "SUMÁRIO":
                continue
            
            match = re.match(self.PATTERN_SUMARIO, texto_limpo)
            
            if match:
                prefixo_bruto = match.group(1)
                titulo = match.group(2).strip()
                
                # Normalizar prefixo (remover ponto final)
                prefixo = prefixo_bruto.rstrip('.')
                
                # Calcular level pela contagem de pontos
                level = prefixo.count('.') + 1
                
                # Criar chave normalizada
                chave = f"{prefixo} {titulo}"
                
                estrutura.append({
                    "tipo": "TITULO",
                    "level": level,
                    "prefixo": prefixo,
                    "texto": titulo,
                    "chave": chave
                })
        
        self.estrutura = estrutura
        return estrutura
    
    def _limpar_texto(self, texto: str) -> str:
        """Remove caracteres especiais e espaços extras"""
        return texto.replace('\xa0', ' ').strip()
    
    def obter_hierarquia(self) -> Dict:
        """Retorna estrutura hierárquica aninhada"""
        # Implementar lógica para criar árvore hierárquica
        pass
```

---

### **2.2 - Módulo: `content_mapper.py`**

**Responsabilidade:** Mapear conteúdo do arquivo fonte aos títulos

```python
"""
Módulo para mapeamento de conteúdo do template
"""

import re
from docx import Document
from typing import Dict, List, Any


class ContentMapper:
    """Mapeador de conteúdo do arquivo fonte"""
    
    # Regex para detectar títulos no conteúdo
    PATTERN_TITULO = r'^\s*(\d+(?:\.\d{1,2})*\.?)\s+([A-ZÁÀÂÃÉÊÍÓÔÕÚÇ].*)$'
    
    # Marcadores especiais
    MARCADORES_TABELAS = {
        '[INSERIR_TABELA_HISTORICA]': 'TABELA_HISTORICA',
        '[INSERIR_TABELA_MACRODESAFIO]': 'TABELA_MACRODESAFIO',
        '[INSERIR_TABELA_MONITORAMENTO]': 'TABELA_MONITORAMENTO',
        '[INSERIR_TABELA_CNJ]': 'TABELA_CNJ',
        '[INICIAR_SECAO_SUPERINTENDENCIAS]': 'SECAO_SUPERINTENDENCIAS'
    }
    
    MARCADORES_VARIAVEIS = {
        '[NUMERO_METAS_2025]': 'var_total_metas',
        '[NUMERO_MACRODESAFIOS]': 'var_total_macros',
        '[PERCENTUAL_VERDE]': 'var_percentual_verde'
    }
    
    def __init__(self, template_path: str):
        self.template_path = template_path
        self.conteudo_mapeado = {}
        self.modo_lista = None
        self.itens_lista_temp = []
    
    def mapear_conteudo(self) -> Dict[str, List[Dict]]:
        """
        Mapeia todo o conteúdo do documento fonte
        
        Returns:
            Dicionário onde:
            - Chave = "1 INTRODUÇÃO", "1.1 Contextualização", etc.
            - Valor = Lista de blocos de conteúdo
        """
        doc = Document(self.template_path)
        chave_titulo_atual = None
        
        for para in doc.paragraphs:
            texto = self._limpar_texto(para.text)
            
            if not texto:
                continue
            
            # 1. PROCESSAR CONTROLES DE LISTA
            if self._processar_controle_lista(texto, chave_titulo_atual):
                continue
            
            # 2. SE ESTÁ EM MODO LISTA, acumular item
            if self.modo_lista:
                self.itens_lista_temp.append(texto)
                continue
            
            # 3. DETECTAR TÍTULO
            match = re.match(self.PATTERN_TITULO, texto)
            if match and self._validar_titulo(match):
                prefixo = match.group(1).rstrip('.')
                titulo = match.group(2).strip()
                chave_titulo_atual = f"{prefixo} {titulo}"
                self.conteudo_mapeado[chave_titulo_atual] = []
                continue
            
            # 4. PROCESSAR CONTEÚDO (se temos título ativo)
            if chave_titulo_atual:
                bloco = self._processar_bloco_conteudo(texto, para)
                if bloco:
                    self.conteudo_mapeado[chave_titulo_atual].append(bloco)
        
        return self.conteudo_mapeado
    
    def _processar_controle_lista(self, texto: str, chave_atual: str) -> bool:
        """Processa marcadores de início/fim de lista"""
        if texto == '[INICIAR_LISTA_NUMERADA]':
            self.modo_lista = 'NUMERADA'
            self.itens_lista_temp = []
            return True
        
        if texto == '[FINALIZAR_LISTA_NUMERADA]':
            if chave_atual:
                self.conteudo_mapeado[chave_atual].append({
                    'tipo': 'LISTA_NUMERADA',
                    'itens': self.itens_lista_temp.copy()
                })
            self.modo_lista = None
            self.itens_lista_temp = []
            return True
        
        if texto == '[INICIAR_LISTA_MARCADORES]':
            self.modo_lista = 'MARCADORES'
            self.itens_lista_temp = []
            return True
        
        if texto == '[FINALIZAR_LISTA_MARCADORES]':
            if chave_atual:
                self.conteudo_mapeado[chave_atual].append({
                    'tipo': 'LISTA_MARCADORES',
                    'itens': self.itens_lista_temp.copy()
                })
            self.modo_lista = None
            self.itens_lista_temp = []
            return True
        
        return False
    
    def _processar_bloco_conteudo(self, texto: str, para) -> Dict:
        """Identifica tipo de bloco e retorna estrutura apropriada"""
        
        # Verificar marcadores de tabelas
        for marcador, tipo in self.MARCADORES_TABELAS.items():
            if texto == marcador:
                return {'tipo': tipo, 'dados': None}
        
        # Verificar texto com destaque (começa com #)
        if texto.startswith('#'):
            return {
                'tipo': 'TEXTO_DESTAQUE',
                'texto': texto[1:].strip(),
                'cor': (162, 22, 18)  # Vermelho
            }
        
        # Verificar se contém variáveis
        for marcador, var_name in self.MARCADORES_VARIAVEIS.items():
            if marcador in texto:
                return {
                    'tipo': 'PARAGRAFO_COM_VARIAVEL',
                    'texto': texto,
                    'variaveis': [var_name]
                }
        
        # Parágrafo normal
        return {
            'tipo': 'PARAGRAFO',
            'texto': texto,
            'alinhamento': self._detectar_alinhamento(para)
        }
    
    def _validar_titulo(self, match) -> bool:
        """Valida se é realmente um título (evitar falsos positivos)"""
        prefixo = match.group(1)
        
        # Deve ter pelo menos um ponto
        if '.' not in prefixo.rstrip('.'):
            # Títulos de nível 1 são OK sem ponto interno
            return True
        
        # Segmentos devem ter no máximo 2 dígitos
        segmentos = prefixo.rstrip('.').split('.')
        for seg in segmentos:
            if len(seg) > 2:
                return False
        
        return True
    
    def _limpar_texto(self, texto: str) -> str:
        """Remove caracteres especiais"""
        return texto.replace('\xa0', ' ').strip()
    
    def _detectar_alinhamento(self, para):
        """Detecta alinhamento do parágrafo"""
        # Implementar detecção de alinhamento
        return 'JUSTIFY'
```

---

### **2.3 - Módulo: `template_reader.py`**

**Responsabilidade:** Orquestrar leitura e integração

```python
"""
Módulo principal para leitura de templates
"""

from .structure_parser import StructureParser
from .content_mapper import ContentMapper
from typing import Dict, List, Any


class TemplateReader:
    """Leitor e integrador de templates"""
    
    def __init__(self, sumario_path: str, conteudo_path: str):
        self.sumario_path = sumario_path
        self.conteudo_path = conteudo_path
        
        self.parser = StructureParser(sumario_path)
        self.mapper = ContentMapper(conteudo_path)
        
        self.estrutura = []
        self.conteudo = {}
    
    def processar_templates(self) -> Dict[str, Any]:
        """
        Processa ambos os templates e retorna estrutura integrada
        
        Returns:
            {
                'estrutura': [...],  # Lista hierárquica de títulos
                'conteudo': {...},   # Dicionário de conteúdo por chave
                'metadados': {...}   # Informações extras
            }
        """
        print("📄 Processando template de sumário...")
        self.estrutura = self.parser.extrair_estrutura()
        print(f"   ✅ {len(self.estrutura)} títulos identificados")
        
        print("📝 Processando template de conteúdo...")
        self.conteudo = self.mapper.mapear_conteudo()
        print(f"   ✅ {len(self.conteudo)} seções mapeadas")
        
        # Validar correspondência
        self._validar_correspondencia()
        
        return {
            'estrutura': self.estrutura,
            'conteudo': self.conteudo,
            'metadados': {
                'total_titulos': len(self.estrutura),
                'total_secoes': len(self.conteudo),
                'sem_conteudo': self._identificar_secoes_vazias()
            }
        }
    
    def _validar_correspondencia(self):
        """Valida se todos os títulos do sumário têm conteúdo"""
        titulos_sumario = {item['chave'] for item in self.estrutura}
        titulos_conteudo = set(self.conteudo.keys())
        
        # Títulos sem conteúdo
        sem_conteudo = titulos_sumario - titulos_conteudo
        if sem_conteudo:
            print(f"\n⚠️  Atenção: {len(sem_conteudo)} título(s) sem conteúdo:")
            for titulo in sorted(sem_conteudo):
                print(f"   - {titulo}")
        
        # Conteúdo sem título no sumário
        extras = titulos_conteudo - titulos_sumario
        if extras:
            print(f"\n⚠️  Atenção: {len(extras)} seção(ões) sem título no sumário:")
            for titulo in sorted(extras):
                print(f"   - {titulo}")
    
    def _identificar_secoes_vazias(self) -> List[str]:
        """Retorna lista de títulos sem conteúdo"""
        titulos_sumario = {item['chave'] for item in self.estrutura}
        titulos_conteudo = set(self.conteudo.keys())
        return list(titulos_sumario - titulos_conteudo)
```

---

## 🏗️ FASE 3: INTEGRAÇÃO COM GERADOR ATUAL

### **3.1 - Atualizar `gerador_relatorio.py`**

**Novo Fluxo Principal:**

```python
"""
GERADOR DE RELATÓRIO DE METAS ESTRATÉGICAS - TJMG
Versão: 4.0 (Com Sistema de Templates)
"""

import os
from datetime import datetime
from pathlib import Path

# Importações existentes
from gerador_relatorio import (
    Config,
    criar_pasta_saida,
    carregar_dados,
    # ... demais importações
)

# 🆕 Novas importações
from gerador_relatorio.template_reader import TemplateReader
from gerador_relatorio.document_generator import DocumentGenerator


# Caminhos dos templates
TEMPLATE_DIR = Path(__file__).parent / 'templates'
SUMARIO_PATH = TEMPLATE_DIR / 'Sumario_Modelo.docx'
CONTEUDO_PATH = TEMPLATE_DIR / 'Conteudo_Fonte.docx'


def gerar_relatorio_com_template():
    """Função principal NOVA - com templates"""
    
    print("\n" + "="*70)
    print("📊 GERADOR DE RELATÓRIO - VERSÃO 4.0 (COM TEMPLATES)")
    print("="*70 + "\n")
    
    # FASE 1: PROCESSAR TEMPLATES
    print("="*70)
    print("FASE 1: PROCESSAMENTO DOS TEMPLATES")
    print("="*70 + "\n")
    
    if not SUMARIO_PATH.exists():
        print(f"❌ Template de sumário não encontrado: {SUMARIO_PATH}")
        return
    
    if not CONTEUDO_PATH.exists():
        print(f"❌ Template de conteúdo não encontrado: {CONTEUDO_PATH}")
        return
    
    reader = TemplateReader(str(SUMARIO_PATH), str(CONTEUDO_PATH))
    template_data = reader.processar_templates()
    
    # FASE 2: CARREGAR DADOS
    print("\n" + "="*70)
    print("FASE 2: CARREGAMENTO DOS DADOS")
    print("="*70 + "\n")
    
    criar_pasta_saida()
    df = carregar_dados()
    if df is None:
        return
    
    mapeamento = carregar_mapeamento_superintendencias()
    df = adicionar_coluna_superintendencia(df, mapeamento)
    grupos_super = agrupar_por_superintendencia_e_macro(df)
    
    # FASE 3: CALCULAR VARIÁVEIS DINÂMICAS
    print("\n" + "="*70)
    print("FASE 3: CÁLCULO DE VARIÁVEIS")
    print("="*70 + "\n")
    
    variaveis = {
        'var_total_metas': len(df),
        'var_total_macros': df['MACRODESAFIO'].nunique(),
        'var_percentual_verde': calcular_percentual_verde(df),
        'var_ano_atual': datetime.now().year
    }
    
    print(f"   Total de Metas: {variaveis['var_total_metas']}")
    print(f"   Total de Macrodesafios: {variaveis['var_total_macros']}")
    print(f"   % Faixa Verde: {variaveis['var_percentual_verde']}%")
    
    # FASE 4: GERAR DOCUMENTO
    print("\n" + "="*70)
    print("FASE 4: GERAÇÃO DO DOCUMENTO")
    print("="*70 + "\n")
    
    generator = DocumentGenerator(template_data, df, grupos_super, variaveis)
    doc = generator.gerar_documento()
    
    # FASE 5: SALVAR
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome_arquivo = f"Relatorio_Metas_Estrategicas_{timestamp}.docx"
    caminho_completo = os.path.join(Config.PASTA_SAIDA, nome_arquivo)
    
    doc.save(caminho_completo)
    
    print(f"\n✅ RELATÓRIO GERADO COM SUCESSO!")
    print(f"📁 Localização: {os.path.abspath(caminho_completo)}")
    print("\n" + "="*70 + "\n")


def calcular_percentual_verde(df):
    """Calcula percentual de metas na faixa verde"""
    # Implementar lógica de cálculo
    return 75.5


if __name__ == "__main__":
    try:
        gerar_relatorio_com_template()
    except Exception as e:
        print(f"\n❌ ERRO: {e}")
        import traceback
        traceback.print_exc()
```

---

### **3.2 - Criar `document_generator.py`**

**Responsabilidade:** Gerar documento final integrando template + dados

```python
"""
Módulo para geração do documento final integrando templates e dados
"""

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import Dict, List, Any

from .config import Config
from .document_builder import criar_documento
from .table_historico import adicionar_tabela_historica
from .table_macrodesafio import adicionar_tabela_macrodesafio
from .table_cnj import adicionar_tabela_metas_nacionais
from .table_monitoramento import adicionar_tabela_resultado_monitoramento
from .table_superintendencia import (
    adicionar_nova_secao_superintendencia,
    adicionar_secao_macrodesafio
)


class DocumentGenerator:
    """Gerador de documento integrando templates e dados"""
    
    def __init__(self, template_data: Dict, df, grupos_super: Dict, variaveis: Dict):
        self.estrutura = template_data['estrutura']
        self.conteudo = template_data['conteudo']
        self.df = df
        self.grupos_super = grupos_super
        self.variaveis = variaveis
        self.doc = None
    
    def gerar_documento(self):
        """Gera documento final"""
        
        # Criar documento base
        print("📝 Criando documento base...")
        primeira_super = list(self.grupos_super.keys())[0]
        self.doc = criar_documento(primeira_super)
        
        # Iterar pela estrutura do sumário
        print("✍️  Processando estrutura do template...")
        
        for idx, item in enumerate(self.estrutura):
            chave = item['chave']
            level = item['level']
            texto = item['texto']
            
            print(f"   [{idx+1}/{len(self.estrutura)}] {chave}")
            
            # Adicionar título
            self._adicionar_titulo(texto, level)
            
            # Buscar e processar conteúdo correspondente
            if chave in self.conteudo:
                self._processar_conteudo(chave)
            else:
                # Título sem conteúdo (apenas título fica no doc)
                pass
        
        return self.doc
    
    def _adicionar_titulo(self, texto: str, level: int):
        """Adiciona título formatado ao documento"""
        para = self.doc.add_paragraph()
        run = para.add_run(texto)
        
        # Formatação baseada no level
        if level == 1:
            run.font.size = Pt(14)
            run.font.bold = True
            run.font.color.rgb = RGBColor(227, 108, 10)  # Laranja
            para.paragraph_format.space_before = Pt(18)
            para.paragraph_format.space_after = Pt(12)
        elif level == 2:
            run.font.size = Pt(12)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)
            para.paragraph_format.space_before = Pt(12)
            para.paragraph_format.space_after = Pt(6)
        else:
            run.font.size = Pt(11)
            run.font.bold = True
            para.paragraph_format.space_before = Pt(6)
            para.paragraph_format.space_after = Pt(3)
        
        run.font.name = Config.FONTE_PADRAO
    
    def _processar_conteudo(self, chave: str):
        """Processa todos os blocos de conteúdo de uma seção"""
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
    
    def _adicionar_paragrafo(self, bloco: Dict):
        """Adiciona parágrafo normal"""
        para = self.doc.add_paragraph(bloco['texto'])
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        para.paragraph_format.space_after = Pt(6)
        para.paragraph_format.line_spacing = 1.5
        
        for run in para.runs:
            run.font.size = Pt(12)
            run.font.name = Config.FONTE_PADRAO
    
    def _adicionar_paragrafo_variavel(self, bloco: Dict):
        """Adiciona parágrafo substituindo variáveis"""
        texto = bloco['texto']
        
        # Substituir variáveis
        for var_name in bloco['variaveis']:
            marcador = self._get_marcador_by_var(var_name)
            valor = self.variaveis.get(var_name, '???')
            texto = texto.replace(marcador, str(valor))
        
        para = self.doc.add_paragraph(texto)
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        para.paragraph_format.space_after = Pt(6)
        
        for run in para.runs:
            run.font.size = Pt(12)
            run.font.name = Config.FONTE_PADRAO
    
    def _adicionar_texto_destaque(self, bloco: Dict):
        """Adiciona texto destacado (negrito, cor)"""
        para = self.doc.add_paragraph()
        run = para.add_run(bloco['texto'])
        run.font.bold = True
        run.font.size = Pt(12)
        run.font.name = Config.FONTE_PADRAO
        
        if 'cor' in bloco:
            run.font.color.rgb = RGBColor(*bloco['cor'])
        
        para.paragraph_format.space_after = Pt(6)
    
    def _adicionar_lista_numerada(self, bloco: Dict):
        """Adiciona lista numerada"""
        for idx, item in enumerate(bloco['itens'], 1):
            para = self.doc.add_paragraph(f"{idx}. {item}")
            para.paragraph_format.left_indent = Pt(36)
            para.paragraph_format.space_after = Pt(3)
    
    def _adicionar_lista_marcadores(self, bloco: Dict):
        """Adiciona lista com marcadores"""
        for item in bloco['itens']:
            para = self.doc.add_paragraph(item, style='List Bullet')
            para.paragraph_format.space_after = Pt(3)
    
    def _gerar_secoes_superintendencias(self):
        """Gera seções detalhadas por superintendência (tabelas dinâmicas)"""
        primeira_super = True
        
        for superintendencia, grupos_macro in self.grupos_super.items():
            if not primeira_super:
                adicionar_nova_secao_superintendencia(self.doc, superintendencia, False)
            
            primeira_secao = True
            for macrodesafio, df_grupo in grupos_macro:
                adicionar_secao_macrodesafio(self.doc, macrodesafio, df_grupo, primeira_secao)
                primeira_secao = False
            
            primeira_super = False
    
    def _get_marcador_by_var(self, var_name: str) -> str:
        """Retorna marcador original dado o nome da variável"""
        from .content_mapper import ContentMapper
        for marcador, nome in ContentMapper.MARCADORES_VARIAVEIS.items():
            if nome == var_name:
                return marcador
        return var_name
```

---

## 📚 FASE 4: GUIA DE USO DOS TEMPLATES

### **4.1 - Criar `templates/readme_templates.md`**

```markdown
# 📘 GUIA DE USO DOS TEMPLATES

## 🎯 Visão Geral

Os templates permitem controlar a estrutura e conteúdo do relatório sem modificar código.

### **Arquivos:**

1. **Sumario_Modelo.docx** - Define estrutura hierárquica
2. **Conteudo_Fonte.docx** - Contém textos e marcadores

---

## 📝 SUMARIO_MODELO.DOCX

### **Formato:**

```
1 TÍTULO NÍVEL 1
1.1 Título Nível 2
1.2 Outro Título Nível 2
2 OUTRO TÍTULO NÍVEL 1
2.1 Subtítulo
2.1.1 Nível 3
```

### **Regras:**

✅ Use numeração hierárquica (1, 1.1, 1.1.1)  
✅ Ponto final é opcional (1. ou 1)  
✅ Títulos EXATOS serão mapeados no conteúdo  
❌ Não adicione números de página  

---

## 📄 CONTEUDO_FONTE.DOCX

### **Marcadores de Tabelas Dinâmicas:**

| Marcador | Descrição |
|----------|-----------|
| `[INSERIR_TABELA_HISTORICA]` | Tabela de metas aprovadas (4 anos) |
| `[INSERIR_TABELA_MACRODESAFIO]` | Distribuição por macrodesafio |
| `[INSERIR_TABELA_MONITORAMENTO]` | Resultado do monitoramento |
| `[INSERIR_TABELA_CNJ]` | Metas nacionais do CNJ |
| `[INICIAR_SECAO_SUPERINTENDENCIAS]` | Seções detalhadas (tabelas paisagem) |

### **Marcadores de Variáveis:**

| Marcador | Descrição |
|----------|-----------|
| `[NUMERO_METAS_2025]` | Total de metas do ano |
| `[NUMERO_MACRODESAFIOS]` | Total de macrodesafios |
| `[PERCENTUAL_VERDE]` | % de metas na faixa verde |

### **Formatação Especial:**

- **Texto com `#`**: `#Este texto será negrito e vermelho`
- **Listas numeradas**:
  ```
  [INICIAR_LISTA_NUMERADA]
  Item 1
  Item 2
  [FINALIZAR_LISTA_NUMERADA]
  ```
- **Listas com marcadores**:
  ```
  [INICIAR_LISTA_MARCADORES]
  Item A
  Item B
  [FINALIZAR_LISTA_MARCADORES]
  ```

---

## 🔄 WORKFLOW

1. Edite **Sumario_Modelo.docx** para ajustar estrutura
2. Edite **Conteudo_Fonte.docx** para ajustar textos
3. Execute `python gerador_relatorio.py`
4. Relatório final gerado em `relatorios_gerados/`
```

---

## ✅ FASE 5: CHECKLIST DE IMPLEMENTAÇÃO

### **Prioridade 1 (Core):**

- [ ] Criar pasta `templates/`
- [ ] Criar `structure_parser.py`
- [ ] Criar `content_mapper.py`
- [ ] Criar `template_reader.py`
- [ ] Criar `document_generator.py`
- [ ] Atualizar `gerador_relatorio.py` com novo fluxo
- [ ] Criar templates de exemplo (Sumario_Modelo.docx e Conteudo_Fonte.docx)

### **Prioridade 2 (Validação):**

- [ ] Adicionar testes unitários para parsers
- [ ] Validar correspondência sumário ↔ conteúdo
- [ ] Tratamento de erros (templates não encontrados)
- [ ] Logs detalhados de processamento

### **Prioridade 3 (Melhorias):**

- [ ] Interface GUI para edição de templates
- [ ] Preview do relatório antes de gerar
- [ ] Validador de templates (sintaxe correta)
- [ ] Versionamento de templates

---

## 🚀 BENEFÍCIOS DA IMPLEMENTAÇÃO

| Antes | Depois |
|-------|--------|
| ❌ Estrutura hardcoded no código | ✅ Estrutura editável via Word |
| ❌ Sem textos explicativos | ✅ Textos completos nos templates |
| ❌ Difícil adicionar/remover seções | ✅ Edição fácil sem código |
| ❌ Ordem fixa | ✅ Ordem personalizável |
| ❌ Manutenção por programador | ✅ Manutenção por analista |

---

## 📊 COMPARAÇÃO DE ARQUITETURA

### **ANTES:**

```
Dados (Excel) → Código Python → Relatório DOCX
                    ↑
              (tudo hardcoded)
```

### **DEPOIS:**

```
┌─ Sumario_Modelo.docx (estrutura)
│
├─ Conteudo_Fonte.docx (textos) ─→ Template Reader ─┐
│                                                     │
└─ Dados (Excel) ───────────────→ Data Loader ──────┼─→ Document Generator ─→ Relatório Final
                                                      │
                                    Variáveis ────────┘
                                    (cálculos)
```

---

## 🎓 PRÓXIMOS PASSOS SUGERIDOS

1. **Criar templates de exemplo** com conteúdo real do TJMG
2. **Implementar módulos core** (parsers e readers)
3. **Testar com relatório simplificado** (apenas seção de introdução)
4. **Expandir gradualmente** para todas as seções
5. **Validar com stakeholders** (analistas e gestores)
6. **Deploy em produção**

---

## 📞 DÚVIDAS E SUPORTE

Para dúvidas sobre implementação:
- Consulte exemplos em `docs/exemplos_templates/`
- Revise logs de processamento
- Valide templates com `validador_templates.py`

---

**Versão do Documento:** 1.0  
**Data:** Dezembro/2025  
**Autor:** Sistema de Documentação Automatizada
