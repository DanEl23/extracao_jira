"""
Script de exemplo para atualizar permanentemente o histórico de metas
no arquivo meta_por_superintendencia.py

IMPORTANTE: Este script modifica o código-fonte do módulo!
Use apenas quando quiser salvar permanentemente os dados do ano atual.
"""

from meta_por_superintendencia import (
    atualizar_historico_com_ano_atual,
    salvar_historico_atualizado,
    gerar_relatorio_historico
)

def main():
    print("=" * 80)
    print("ATUALIZAÇÃO AUTOMÁTICA DO HISTÓRICO DE METAS")
    print("=" * 80)
    print()
    
    # 1. Visualizar dados que serão salvos
    print("📊 Dados que serão adicionados ao histórico:")
    print("-" * 80)
    dados = atualizar_historico_com_ano_atual()
    
    ano_atual = max(dados['historico'].keys())
    dados_ano = dados['historico'][ano_atual]
    
    print(f"\nAno: {ano_atual}")
    print(f"  - Metas Nacionais: {dados_ano['nacionais']}")
    print(f"  - Metas Institucionais: {dados_ano['institucionais']}")
    print(f"  - Total: {dados_ano['total']}")
    print()
    print(f"Período de análise: {dados['periodo']}")
    print(f"Variação no período:")
    print(f"  - Metas Nacionais: {dados['variacao']['nacionais']:+}%")
    print(f"  - Metas Institucionais: {dados['variacao']['institucionais']:+}%")
    print(f"  - Total: {dados['variacao']['total']:+}%")
    print()
    
    # 2. Perguntar confirmação
    resposta = input("⚠️  Deseja salvar permanentemente no arquivo? (s/N): ").strip().lower()
    
    if resposta in ['s', 'sim', 'yes', 'y']:
        print("\n🔄 Salvando...")
        sucesso = salvar_historico_atualizado()
        
        if sucesso:
            print("\n✅ Histórico atualizado com sucesso!")
            print("\n📋 Relatório atualizado:")
            print("-" * 80)
            print(gerar_relatorio_historico())
        else:
            print("\n❌ Falha ao atualizar o histórico.")
    else:
        print("\n❌ Operação cancelada. Nenhuma alteração foi feita.")
    
    print("\n" + "=" * 80)
    print("Fim da execução")
    print("=" * 80)

if __name__ == "__main__":
    main()
