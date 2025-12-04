"""
Script de Validação - Compara extração original vs unificada
"""

import pandas as pd
from pathlib import Path
import sys

def comparar_arquivos(arquivo_original, arquivo_novo, nome_descritivo):
    """Compara dois arquivos Excel e reporta diferenças"""
    print("\n" + "="*70)
    print(f"📊 COMPARANDO: {nome_descritivo}")
    print("="*70)
    
    try:
        df_orig = pd.read_excel(arquivo_original)
        df_novo = pd.read_excel(arquivo_novo)
        
        print(f"\n✓ Original: {arquivo_original}")
        print(f"  - Linhas: {len(df_orig)}")
        print(f"  - Colunas: {len(df_orig.columns)}")
        
        print(f"\n✓ Novo: {arquivo_novo}")
        print(f"  - Linhas: {len(df_novo)}")
        print(f"  - Colunas: {len(df_novo.columns)}")
        
        # Comparar número de linhas
        if len(df_orig) != len(df_novo):
            print(f"\n⚠️  DIFERENÇA no número de linhas:")
            print(f"   Original: {len(df_orig)} | Novo: {len(df_novo)}")
            print(f"   Diferença: {abs(len(df_orig) - len(df_novo))} linhas")
        else:
            print(f"\n✅ Número de linhas: IGUAL ({len(df_orig)})")
        
        # Comparar colunas
        colunas_orig = set(df_orig.columns)
        colunas_novo = set(df_novo.columns)
        
        if colunas_orig == colunas_novo:
            print(f"✅ Colunas: IDÊNTICAS ({len(colunas_orig)} colunas)")
        else:
            print(f"\n⚠️  DIFERENÇA nas colunas:")
            
            apenas_orig = colunas_orig - colunas_novo
            if apenas_orig:
                print(f"   Apenas no original: {sorted(apenas_orig)}")
            
            apenas_novo = colunas_novo - colunas_orig
            if apenas_novo:
                print(f"   Apenas no novo: {sorted(apenas_novo)}")
        
        # Comparar ordem das colunas
        if list(df_orig.columns) == list(df_novo.columns):
            print("✅ Ordem das colunas: IDÊNTICA")
        else:
            print("\n⚠️  DIFERENÇA na ordem das colunas:")
            print(f"   Original (5 primeiras): {list(df_orig.columns[:5])}")
            print(f"   Novo (5 primeiras):     {list(df_novo.columns[:5])}")
        
        # Comparar primeiras linhas (se houver linhas em comum)
        if len(df_orig) > 0 and len(df_novo) > 0:
            # Comparar apenas colunas comuns
            colunas_comuns = sorted(list(colunas_orig & colunas_novo))
            
            if len(colunas_comuns) > 0:
                min_linhas = min(3, len(df_orig), len(df_novo))
                
                df_orig_comum = df_orig[colunas_comuns].head(min_linhas)
                df_novo_comum = df_novo[colunas_comuns].head(min_linhas)
                
                # Comparar valores nas colunas-chave
                colunas_chave = ['META_ID', 'Chave', 'Resumo']
                colunas_chave_disponiveis = [c for c in colunas_chave if c in colunas_comuns]
                
                if colunas_chave_disponiveis:
                    print(f"\n📋 Comparação de chaves (primeiras {min_linhas} linhas):")
                    for col in colunas_chave_disponiveis:
                        vals_orig = df_orig[col].head(min_linhas).tolist()
                        vals_novo = df_novo[col].head(min_linhas).tolist()
                        
                        if vals_orig == vals_novo:
                            print(f"   ✅ {col}: IGUAL")
                        else:
                            print(f"   ⚠️  {col}: DIFERENTE")
                            print(f"      Original: {vals_orig}")
                            print(f"      Novo:     {vals_novo}")
        
        return True
        
    except FileNotFoundError as e:
        print(f"\n❌ Arquivo não encontrado: {e}")
        return False
    except Exception as e:
        print(f"\n❌ Erro ao comparar: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Função principal"""
    print("\n" + "="*70)
    print("🔍 VALIDAÇÃO DE EXTRAÇÃO - Original vs Unificado")
    print("="*70)
    
    pasta = Path("exports")
    
    # Teste 1: Extração Simples
    sucesso1 = comparar_arquivos(
        pasta / "dados_exportados_jira.xlsx",
        pasta / "dados_exportados_jira.xlsx",  # Mesmo arquivo (teste de baseline)
        "Extração Simples (Baseline)"
    )
    
    # Teste 2: Extração Anual
    sucesso2 = comparar_arquivos(
        pasta / "dados_exportados_jira_por_ano.xlsx",
        pasta / "dados_jira_anual.xlsx",  # Arquivo gerado pelo script novo (antes da correção)
        "Extração Anual (Original vs Novo)"
    )
    
    # Teste 3: CNJ
    print("\n" + "="*70)
    print("📊 Arquivo CNJ")
    print("="*70)
    arquivo_cnj = pasta / "resultados_cnj.xlsx"
    if arquivo_cnj.exists():
        df_cnj = pd.read_excel(arquivo_cnj)
        print(f"\n✓ {arquivo_cnj}")
        print(f"  - Linhas: {len(df_cnj)}")
        print(f"  - Colunas: {df_cnj.columns.tolist()}")
        print(f"\n📋 Metas únicas coletadas:")
        if 'Meta' in df_cnj.columns:
            metas_unicas = df_cnj['Meta'].unique()
            for i, meta in enumerate(metas_unicas, 1):
                print(f"   {i}. {meta}")
            print(f"\n✅ Total de metas: {len(metas_unicas)}")
        else:
            print("   ⚠️  Coluna 'Meta' não encontrada")
    else:
        print(f"\n❌ Arquivo não encontrado: {arquivo_cnj}")
    
    # Resumo Final
    print("\n" + "="*70)
    print("📝 RESUMO")
    print("="*70)
    print("\n💡 Para testar o extrator unificado:")
    print("   1. Renomeie os arquivos atuais (backup)")
    print("   2. Execute: python extrator_unificado.py 2023 2024 2025")
    print("   3. Execute: python validar_extracao.py")
    print("   4. Compare os resultados")
    print("\n")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Validação cancelada pelo usuário")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ ERRO INESPERADO: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
