"""
Mapeamento de Metas para Superintendências - TJMG
Versão: 4.0 - Totalmente Modularizado

Este arquivo mantém retrocompatibilidade importando do package base_dados/
Para novos desenvolvimentos, importe diretamente de base_dados
"""

# Importar tudo do package modularizado
from base_dados import *

# Manter compatibilidade com código antigo
from base_dados.mapeamentos import META_SUPERINTENDENCIA, ORDEM_SUPERINTENDENCIAS, estatisticas
from base_dados.metas_cnj import METAS_CNJ_ALVOS, METAS_CNJ_PARA_MACRODESAFIO
from base_dados.historicos import HISTORICO_METAS_APROVADAS, HISTORICO_METAS_POR_MACRODESAFIO, VARIACAO_2021_2024
from base_dados.calculos import (
    calcular_metas_atuais,
    atualizar_historico_com_ano_atual,
    calcular_metas_por_superintendencia,
    calcular_metas_atuais_por_macrodesafio,
    atualizar_historico_macrodesafio_com_ano_atual,
    calcular_cumprimento_metas_cnj,
    gerar_relatorio_historico
)


if __name__ == "__main__":
    # Executar este arquivo diretamente mostra as estatísticas
    print("📊 ANÁLISE COMPLETA DO PLANEJAMENTO ESTRATÉGICO\n")
    
    # Estatísticas básicas
    stats = estatisticas()
    print("=" * 70)
    print("ESTATÍSTICAS DO MAPEAMENTO")
    print("=" * 70)
    print(f"Total de metas mapeadas: {stats['total_metas']}")
    print(f"  - Metas CNJ: {stats['metas_cnj']}")
    print(f"  - Metas TJMG: {stats['metas_tjmg']}")
    print("\nDistribuição por Superintendência:")
    for super_nome, count in sorted(stats['por_superintendencia'].items()):
        print(f"  {super_nome}: {count} metas")
    
    print("\n")
    
    # Metas atuais das planilhas
    print("=" * 70)
    print("METAS ATUAIS (DAS PLANILHAS)")
    print("=" * 70)
    metas_atuais = calcular_metas_atuais()
    print(f"Ano: {metas_atuais['ano']}")
    print(f"Metas Nacionais: {metas_atuais['nacionais']}")
    print(f"Metas Institucionais: {metas_atuais['institucionais']}")
    print(f"Total: {metas_atuais['total']}")
    
    print("\n")
    
    # Relatório histórico
    print(gerar_relatorio_historico())
    
    print("\n")
    
    # Distribuição por superintendência (atual)
    print("=" * 70)
    print("DISTRIBUIÇÃO ATUAL POR SUPERINTENDÊNCIA")
    print("=" * 70)
    dist = calcular_metas_por_superintendencia()
    for super_nome in ORDEM_SUPERINTENDENCIAS:
        if super_nome in dist:
            info = dist[super_nome]
            print(f"\n{super_nome}:")
            print(f"  Nacionais: {info['nacionais']}")
            print(f"  Institucionais: {info['institucionais']}")
            print(f"  Total: {info['total']}")
            print(f"  Metas: {', '.join(info['metas'][:5])}{'...' if len(info['metas']) > 5 else ''}")
