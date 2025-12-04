"""
Módulo principal para leitura e integração de templates
Orquestra o StructureParser e ContentMapper
"""

from .structure_parser import StructureParser
from .content_mapper import ContentMapper
from typing import Dict, List, Any


class TemplateReader:
    """Leitor e integrador de templates"""
    
    def __init__(self, sumario_path: str, conteudo_path: str):
        """
        Inicializa o leitor de templates
        
        Args:
            sumario_path: Caminho para Sumario_Modelo.docx
            conteudo_path: Caminho para Conteudo_Fonte.docx
        """
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
        stats_estrutura = self.parser.obter_estatisticas()
        print(f"   ✅ {stats_estrutura['total_titulos']} títulos identificados")
        print(f"   📊 Distribuição por nível: {stats_estrutura['por_nivel']}")
        
        # Validar estrutura
        problemas = self.parser.validar_estrutura()
        if problemas:
            print(f"   ⚠️  {len(problemas)} problema(s) na estrutura:")
            for problema in problemas:
                print(f"      • {problema}")
        
        print("\n📝 Processando template de conteúdo...")
        self.conteudo = self.mapper.mapear_conteudo()
        stats_conteudo = self.mapper.obter_estatisticas()
        print(f"   ✅ {stats_conteudo['total_secoes']} seções mapeadas")
        print(f"   📦 {stats_conteudo['total_blocos']} blocos de conteúdo")
        print(f"   📋 Tipos de blocos: {stats_conteudo['tipos_blocos']}")
        
        # Validar correspondência
        print("\n🔍 Validando correspondência...")
        avisos = self._validar_correspondencia()
        
        if not avisos:
            print("   ✅ Todos os títulos têm conteúdo correspondente")
        
        return {
            'estrutura': self.estrutura,
            'conteudo': self.conteudo,
            'metadados': {
                'total_titulos': stats_estrutura['total_titulos'],
                'total_secoes': stats_conteudo['total_secoes'],
                'total_blocos': stats_conteudo['total_blocos'],
                'distribuicao_niveis': stats_estrutura['por_nivel'],
                'tipos_blocos': stats_conteudo['tipos_blocos'],
                'sem_conteudo': self._identificar_secoes_vazias(),
                'avisos': avisos
            }
        }
    
    def _validar_correspondencia(self) -> List[str]:
        """
        Valida se todos os títulos do sumário têm conteúdo
        
        Returns:
            Lista de avisos (vazia se tudo OK)
        """
        avisos = []
        
        titulos_sumario = {item['chave'] for item in self.estrutura}
        titulos_conteudo = set(self.conteudo.keys())
        
        # Títulos sem conteúdo
        sem_conteudo = titulos_sumario - titulos_conteudo
        if sem_conteudo:
            avisos.append(f"{len(sem_conteudo)} título(s) sem conteúdo")
            print(f"   ⚠️  {len(sem_conteudo)} título(s) sem conteúdo:")
            for titulo in sorted(sem_conteudo):
                print(f"      • {titulo}")
        
        # Conteúdo sem título no sumário
        extras = titulos_conteudo - titulos_sumario
        if extras:
            avisos.append(f"{len(extras)} seção(ões) sem título no sumário")
            print(f"   ⚠️  {len(extras)} seção(ões) sem título no sumário:")
            for titulo in sorted(extras):
                print(f"      • {titulo}")
        
        return avisos
    
    def _identificar_secoes_vazias(self) -> List[str]:
        """
        Retorna lista de títulos sem conteúdo
        
        Returns:
            Lista de chaves de títulos sem conteúdo
        """
        titulos_sumario = {item['chave'] for item in self.estrutura}
        titulos_conteudo = set(self.conteudo.keys())
        return sorted(list(titulos_sumario - titulos_conteudo))
    
    def obter_conteudo_por_chave(self, chave: str) -> List[Dict]:
        """
        Retorna o conteúdo associado a uma chave específica
        
        Args:
            chave: Chave do título (ex: "1 INTRODUÇÃO")
            
        Returns:
            Lista de blocos de conteúdo ou lista vazia
        """
        return self.conteudo.get(chave, [])
    
    def tem_conteudo(self, chave: str) -> bool:
        """
        Verifica se uma chave tem conteúdo associado
        
        Args:
            chave: Chave do título
            
        Returns:
            True se tem conteúdo
        """
        return chave in self.conteudo and len(self.conteudo[chave]) > 0
