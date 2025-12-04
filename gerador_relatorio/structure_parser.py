"""
Módulo para parsing da estrutura do sumário
Extrai hierarquia de títulos do arquivo Sumario_Modelo.docx
"""

import re
from docx import Document
from typing import List, Dict


class StructureParser:
    """Parser para extrair estrutura hierárquica do sumário"""
    
    # Regex para detectar títulos numerados no sumário
    PATTERN_SUMARIO = r'^\s*(\d+(?:\.\d+)*\.?)\s+(.+?)(?:\s*\.{2,}\s*\d+\s*)?$'
    
    def __init__(self, template_path: str):
        """
        Inicializa o parser
        
        Args:
            template_path: Caminho para o arquivo Sumario_Modelo.docx
        """
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
                {
                    "tipo": "TITULO",
                    "level": 2,
                    "prefixo": "1.1",
                    "texto": "Contextualização",
                    "chave": "1.1 Contextualização"
                },
                ...
            ]
        """
        doc = Document(self.template_path)
        estrutura = []
        
        for para in doc.paragraphs:
            texto_limpo = self._limpar_texto(para.text)
            
            # Ignorar linhas vazias e palavra "SUMÁRIO"
            if not texto_limpo or texto_limpo.upper() == "SUMÁRIO":
                continue
            
            # Tentar fazer match com padrão de título numerado
            match = re.match(self.PATTERN_SUMARIO, texto_limpo)
            
            if match:
                prefixo_bruto = match.group(1)
                titulo = match.group(2).strip()
                
                # Normalizar prefixo (remover ponto final se existir)
                prefixo = prefixo_bruto.rstrip('.')
                
                # Calcular level pela contagem de pontos
                # "1" = level 1, "1.1" = level 2, "1.1.1" = level 3
                level = prefixo.count('.') + 1
                
                # Criar chave normalizada para correspondência
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
        """
        Remove caracteres especiais e espaços extras
        
        Args:
            texto: Texto a ser limpo
            
        Returns:
            Texto limpo
        """
        # Substituir non-breaking spaces e remover espaços extras
        texto = texto.replace('\xa0', ' ')
        texto = ' '.join(texto.split())
        return texto.strip()
    
    def obter_estatisticas(self) -> Dict:
        """
        Retorna estatísticas da estrutura extraída
        
        Returns:
            Dicionário com estatísticas
        """
        if not self.estrutura:
            return {
                "total_titulos": 0,
                "por_nivel": {},
                "max_nivel": 0
            }
        
        por_nivel = {}
        for item in self.estrutura:
            level = item['level']
            por_nivel[level] = por_nivel.get(level, 0) + 1
        
        return {
            "total_titulos": len(self.estrutura),
            "por_nivel": por_nivel,
            "max_nivel": max(item['level'] for item in self.estrutura) if self.estrutura else 0
        }
    
    def obter_titulos_por_nivel(self, level: int) -> List[Dict]:
        """
        Retorna todos os títulos de um nível específico
        
        Args:
            level: Nível desejado (1, 2, 3, etc.)
            
        Returns:
            Lista de títulos do nível especificado
        """
        return [item for item in self.estrutura if item['level'] == level]
    
    def validar_estrutura(self) -> List[str]:
        """
        Valida a estrutura hierárquica
        
        Returns:
            Lista de avisos/problemas encontrados (vazia se OK)
        """
        problemas = []
        
        if not self.estrutura:
            problemas.append("Nenhum título encontrado no sumário")
            return problemas
        
        # Verificar se primeiro título é nível 1
        if self.estrutura[0]['level'] != 1:
            problemas.append(f"Primeiro título deveria ser nível 1, mas é nível {self.estrutura[0]['level']}")
        
        # Verificar saltos de nível
        for i in range(1, len(self.estrutura)):
            nivel_anterior = self.estrutura[i-1]['level']
            nivel_atual = self.estrutura[i]['level']
            
            # Não pode pular mais de 1 nível de uma vez
            if nivel_atual > nivel_anterior + 1:
                problemas.append(
                    f"Salto de nível inválido: de {nivel_anterior} para {nivel_atual} "
                    f"no título '{self.estrutura[i]['chave']}'"
                )
        
        return problemas
