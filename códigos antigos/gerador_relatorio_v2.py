"""
GERADOR DE RELATÓRIO DE METAS ESTRATÉGICAS - TJMG
Versão: 3.0 (Totalmente Modularizado)

Este arquivo agora importa todas as funções dos módulos organizados.
"""

import os
from datetime import datetime

# Importar do package modularizado
from gerador_relatorio import (
    Config,
    criar_pasta_saida,
    carregar_dados,
    carregar_mapeamento_superintendencias,
    adicionar_coluna_superintendencia,
    agrupar_por_superintendencia_e_macro,
    criar_documento,
    criar_secao_paisagem_inicial,
    adicionar_tabela_historica,
    adicionar_tabela_macrodesafio,
    adicionar_tabela_metas_nacionais,
    adicionar_tabela_resultado_monitoramento,
    adicionar_nova_secao_superintendencia,
    adicionar_secao_macrodesafio
)


def adicionar_cabecalho_relatorio(doc):
    """Adiciona o cabeçalho completo do relatório na primeira página (retrato)"""
    # Adicionar tabelas históricas
    adicionar_tabela_historica(doc)
    adicionar_tabela_macrodesafio(doc)
    
    # Adicionar tabela de metas nacionais do CNJ
    adicionar_tabela_metas_nacionais(doc)
    
    # Adicionar tabela de resultado do monitoramento
    adicionar_tabela_resultado_monitoramento(doc)


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
    
    # 3. Carregar mapeamento de superintendências
    mapeamento = carregar_mapeamento_superintendencias()
    if not mapeamento:
        print("⚠️  Aviso: Mapeamento não carregado. Continuando sem classificação por superintendência.")
        return
    
    # 4. Adicionar coluna de superintendência ao DataFrame
    df = adicionar_coluna_superintendencia(df, mapeamento)
    
    # 5. Agrupar dados por superintendência e macrodesafio
    grupos_super = agrupar_por_superintendencia_e_macro(df)
    
    # 6. Criar documento com primeira página em retrato
    print("📝 Criando documento Word...")
    primeira_superintendencia = list(grupos_super.keys())[0] if grupos_super else 'Presidência'
    doc = criar_documento(primeira_superintendencia)
    
    # 7. Adicionar tabelas na primeira página (retrato)
    print("📊 Adicionando tabelas históricas e de monitoramento...")
    adicionar_cabecalho_relatorio(doc)
    
    # 8. Criar segunda seção em paisagem para as tabelas de metas
    print("📄 Criando seção em paisagem...")
    criar_secao_paisagem_inicial(doc, primeira_superintendencia)
    
    # 9. Gerar seções por superintendência no mesmo documento
    print("✍️  Gerando seções do relatório por Superintendência...")
    
    primeira_super = True
    for superintendencia, grupos_macro in grupos_super.items():
        print(f"\n📋 Superintendência: {superintendencia}")
        
        # Adicionar nova seção com cabeçalho específico (exceto para a primeira)
        if not primeira_super:
            adicionar_nova_secao_superintendencia(doc, superintendencia, primeira=False)
        
        # Adicionar cada Macrodesafio desta superintendência
        primeira_secao_super = True
        for idx, (macrodesafio, df_grupo) in enumerate(grupos_macro):
            print(f"   → {macrodesafio} ({len(df_grupo)} registros)")
            adicionar_secao_macrodesafio(doc, macrodesafio, df_grupo, primeira_secao=primeira_secao_super)
            primeira_secao_super = False
        
        primeira_super = False
    
    # 10. Salvar documento único
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome_arquivo = f"{Config.NOME_RELATORIO}_{timestamp}.docx"
    caminho_completo = os.path.join(Config.PASTA_SAIDA, nome_arquivo)
    
    doc.save(caminho_completo)
    
    print(f"\n✅ RELATÓRIO GERADO COM SUCESSO!")
    print(f"📁 Localização: {os.path.abspath(caminho_completo)}")
    print(f"📄 Total de Superintendências: {len(grupos_super)}")
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

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome_arquivo = f"{Config.NOME_RELATORIO}_{timestamp}.docx"
    caminho_completo = os.path.join(Config.PASTA_SAIDA, nome_arquivo)
    
    doc.save(caminho_completo)
    
    print("\n" + "="*60)
    print("✅ RELATÓRIO GERADO COM SUCESSO!")
    print("="*60)
    print(f"\n📁 Arquivo: {caminho_completo}")
    print(f"📊 Total de registros processados: {len(df)}")
    print(f"🏢 Total de superintendências: {len(grupos)}")
    
    input("\nPressione ENTER para sair...")
