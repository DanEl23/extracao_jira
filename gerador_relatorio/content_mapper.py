"""
Módulo para mapeamento de conteúdo do template
Mapeia conteúdo do Conteudo_Fonte.docx aos títulos correspondentes
"""

import re
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import Dict, List, Any, Optional


class ContentMapper:
    """Mapeador de conteúdo do arquivo fonte"""
    
    # Regex para detectar títulos no conteúdo
    PATTERN_TITULO = r'^\s*(\d+(?:\.\d{1,2})*\.?)\s+([A-ZÁÀÂÃÉÊÍÓÔÕÚÇ].*)$'
    
    # Marcadores especiais para tabelas dinâmicas
    MARCADORES_TABELAS = {
        '[INSERIR_TABELA_HISTORICA]': 'TABELA_HISTORICA',
        '[INSERIR_TABELA_MACRODESAFIO]': 'TABELA_MACRODESAFIO',
        '[INSERIR_TABELA_MONITORAMENTO]': 'TABELA_MONITORAMENTO',
        '[INSERIR_TABELA_CNJ]': 'TABELA_CNJ',
        '[INICIAR_SECAO_SUPERINTENDENCIAS]': 'SECAO_SUPERINTENDENCIAS',
        '[QUEBRA_PAGINA]': 'QUEBRA_PAGINA'
    }
    
    # Marcadores de variáveis dinâmicas
    MARCADORES_VARIAVEIS = {
        '[NUMERO_METAS_2025]': 'var_total_metas',
        '[NUMERO_METAS_CNJ]': 'var_total_metas_cnj',
        '[NUMERO_METAS_TJMG]': 'var_total_metas_tjmg',
        '[NUMERO_MACRODESAFIOS]': 'var_total_macros',
        '[PERCENTUAL_VERDE]': 'var_percentual_verde',
        '[PERCENTUAL_AMARELO]': 'var_percentual_amarelo',
        '[PERCENTUAL_VERMELHO]': 'var_percentual_vermelho',
        '[VAR_ANO_ATUAL]': 'var_ano_atual',
        '[TOTAL_SUPERINTENDENCIAS]': 'var_total_superintendencias'
    }
    
    def __init__(self, template_path: str):
        """
        Inicializa o mapeador
        
        Args:
            template_path: Caminho para o arquivo Conteudo_Fonte.docx
        """
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
            
            # 3. DETECTAR TÍTULO (apenas se não estiver em modo lista)
            match = re.match(self.PATTERN_TITULO, texto)
            if match and self._validar_titulo(match):
                prefixo_bruto = match.group(1)
                titulo = match.group(2).strip()
                
                # Normalizar prefixo (remover ponto final)
                prefixo = prefixo_bruto.rstrip('.')
                
                # Criar chave normalizada
                chave_titulo_atual = f"{prefixo} {titulo}"
                
                # Inicializar lista de conteúdo para este título
                if chave_titulo_atual not in self.conteudo_mapeado:
                    self.conteudo_mapeado[chave_titulo_atual] = []
                
                continue
            
            # 4. PROCESSAR CONTEÚDO (se temos título ativo)
            if chave_titulo_atual:
                bloco = self._processar_bloco_conteudo(texto, para)
                if bloco:
                    self.conteudo_mapeado[chave_titulo_atual].append(bloco)
        
        return self.conteudo_mapeado
    
    def _processar_controle_lista(self, texto: str, chave_atual: Optional[str]) -> bool:
        """
        Processa marcadores de início/fim de lista
        
        Args:
            texto: Texto a verificar
            chave_atual: Chave do título atual (pode ser None)
            
        Returns:
            True se foi processado um marcador de controle
        """
        if texto == '[INICIAR_LISTA_NUMERADA]':
            self.modo_lista = 'NUMERADA'
            self.itens_lista_temp = []
            return True
        
        if texto == '[FINALIZAR_LISTA_NUMERADA]':
            if chave_atual and self.itens_lista_temp:
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
            if chave_atual and self.itens_lista_temp:
                self.conteudo_mapeado[chave_atual].append({
                    'tipo': 'LISTA_MARCADORES',
                    'itens': self.itens_lista_temp.copy()
                })
            self.modo_lista = None
            self.itens_lista_temp = []
            return True
        
        return False
    
    def _processar_bloco_conteudo(self, texto: str, para) -> Optional[Dict]:
        """
        Identifica tipo de bloco e retorna estrutura apropriada
        
        Args:
            texto: Texto do parágrafo
            para: Objeto parágrafo do python-docx
            
        Returns:
            Dicionário com informações do bloco ou None
        """
        # Verificar marcadores de tabelas
        for marcador, tipo in self.MARCADORES_TABELAS.items():
            if texto == marcador:
                return {'tipo': tipo, 'dados': None}
        
        # Verificar texto com destaque (começa com #)
        if texto.startswith('#'):
            return {
                'tipo': 'TEXTO_DESTAQUE',
                'texto': texto[1:].strip()
            }
        
        # Verificar se contém variáveis
        variaveis_encontradas = []
        for marcador in self.MARCADORES_VARIAVEIS.keys():
            if marcador in texto:
                variaveis_encontradas.append(marcador)
        
        if variaveis_encontradas:
            return {
                'tipo': 'PARAGRAFO_COM_VARIAVEL',
                'texto': texto,
                'variaveis': variaveis_encontradas,
                'alinhamento': self._detectar_alinhamento(para)
            }
        
        # Parágrafo normal
        return {
            'tipo': 'PARAGRAFO',
            'texto': texto,
            'alinhamento': self._detectar_alinhamento(para)
        }
    
    def _validar_titulo(self, match) -> bool:
        """
        Valida se é realmente um título (evitar falsos positivos como "436 magistrados")
        
        Args:
            match: Objeto match do regex
            
        Returns:
            True se é um título válido
        """
        prefixo = match.group(1)
        
        # Remover ponto final para análise
        prefixo_limpo = prefixo.rstrip('.')
        
        # Se não tem ponto interno, aceitar apenas se for um dígito (nível 1)
        if '.' not in prefixo_limpo:
            return len(prefixo_limpo) <= 2  # Aceita "1" ou "10", mas não "436"
        
        # Se tem pontos, validar cada segmento
        segmentos = prefixo_limpo.split('.')
        for seg in segmentos:
            if len(seg) > 2:  # Segmentos devem ter no máximo 2 dígitos
                return False
        
        return True
    
    def _limpar_texto(self, texto: str) -> str:
        """
        Remove caracteres especiais
        
        Args:
            texto: Texto a ser limpo
            
        Returns:
            Texto limpo
        """
        texto = texto.replace('\xa0', ' ')
        texto = ' '.join(texto.split())
        return texto.strip()
    
    def _detectar_alinhamento(self, para) -> str:
        """
        Detecta alinhamento do parágrafo
        
        Args:
            para: Objeto parágrafo do python-docx
            
        Returns:
            String com alinhamento ('LEFT', 'CENTER', 'RIGHT', 'JUSTIFY')
        """
        try:
            alignment = para.alignment
            if alignment == WD_ALIGN_PARAGRAPH.CENTER:
                return 'CENTER'
            elif alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                return 'RIGHT'
            elif alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                return 'JUSTIFY'
            else:
                return 'LEFT'
        except:
            return 'JUSTIFY'  # Default
    
    def obter_estatisticas(self) -> Dict:
        """
        Retorna estatísticas do conteúdo mapeado
        
        Returns:
            Dicionário com estatísticas
        """
        total_secoes = len(self.conteudo_mapeado)
        total_blocos = sum(len(blocos) for blocos in self.conteudo_mapeado.values())
        
        tipos_blocos = {}
        for blocos in self.conteudo_mapeado.values():
            for bloco in blocos:
                tipo = bloco['tipo']
                tipos_blocos[tipo] = tipos_blocos.get(tipo, 0) + 1
        
        return {
            "total_secoes": total_secoes,
            "total_blocos": total_blocos,
            "tipos_blocos": tipos_blocos
        }
