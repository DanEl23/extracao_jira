"""
GERADOR DE RELATÓRIO DE METAS ESTRATÉGICAS - TJMG
Versão: 4.0 (Com Sistema de Templates)

Este gerador utiliza templates Word para definir estrutura e conteúdo,
integrando com dados extraídos do CNJ e Jira.
"""

import os
from datetime import datetime
from pathlib import Path

# Importações do sistema existente
from gerador_relatorio import (
    Config,
    criar_pasta_saida,
    carregar_dados,
    carregar_mapeamento_superintendencias,
    adicionar_coluna_superintendencia,
    agrupar_por_superintendencia_e_macro,
)

# 🆕 Importações dos novos módulos de templates
from gerador_relatorio import (
    TemplateReader,
    DocumentGenerator
)

# Caminhos dos templates
TEMPLATE_DIR = Path(__file__).parent / 'templates'
SUMARIO_PATH = TEMPLATE_DIR / 'Sumario_Modelo.docx'
CONTEUDO_PATH = TEMPLATE_DIR / 'Conteudo_Fonte.docx'


def calcular_variaveis(df, grupos_super) -> dict:
    """
    Calcula variáveis dinâmicas a partir dos dados
    
    Args:
        df: DataFrame com dados das metas
        grupos_super: Dicionário com grupos por superintendência
        
    Returns:
        Dicionário com variáveis calculadas
    """
    # Total de metas
    total_metas = len(df)
    
    # Separar CNJ e TJMG
    # Verificar qual coluna tem o código da meta
    coluna_meta = 'Código Meta' if 'Código Meta' in df.columns else 'Meta'
    
    if coluna_meta in df.columns:
        total_cnj = len(df[df[coluna_meta].str.contains('CNJ', na=False)])
        total_tjmg = len(df[df[coluna_meta].str.contains('TJMG', na=False)])
    else:
        # Fallback: tentar pela coluna MACRODESAFIO ou outra
        total_cnj = 0
        total_tjmg = total_metas
    
    # Total de macrodesafios
    total_macros = df['MACRODESAFIO'].nunique() if 'MACRODESAFIO' in df.columns else 0
    
    # Calcular percentuais por faixa (simulado - ajustar conforme lógica real)
    # Aqui você deve implementar a lógica real de classificação
    percentual_verde = 75.5  # Placeholder
    percentual_amarelo = 15.8  # Placeholder
    percentual_vermelho = 8.7  # Placeholder
    
    # Total de superintendências
    total_superintendencias = len(grupos_super)
    
    # Ano atual
    ano_atual = datetime.now().year
    
    return {
        'var_total_metas': total_metas,
        'var_total_metas_cnj': total_cnj,
        'var_total_metas_tjmg': total_tjmg,
        'var_total_macros': total_macros,
        'var_percentual_verde': percentual_verde,
        'var_percentual_amarelo': percentual_amarelo,
        'var_percentual_vermelho': percentual_vermelho,
        'var_total_superintendencias': total_superintendencias,
        'var_ano_atual': ano_atual
    }


def gerar_relatorio_com_templates():
    """Função principal - geração de relatório com templates"""
    
    print("\n" + "="*70)
    print("📊 GERADOR DE RELATÓRIO - VERSÃO 4.0 (COM TEMPLATES)")
    print("="*70 + "\n")
    
    # ============================================
    # FASE 1: VALIDAR E PROCESSAR TEMPLATES
    # ============================================
    print("="*70)
    print("FASE 1: PROCESSAMENTO DOS TEMPLATES")
    print("="*70 + "\n")
    
    if not SUMARIO_PATH.exists():
        print(f"❌ Template de sumário não encontrado: {SUMARIO_PATH}")
        print("   Execute: python templates/criar_templates.py")
        return
    
    if not CONTEUDO_PATH.exists():
        print(f"❌ Template de conteúdo não encontrado: {CONTEUDO_PATH}")
        print("   Execute: python templates/criar_templates.py")
        return
    
    try:
        reader = TemplateReader(str(SUMARIO_PATH), str(CONTEUDO_PATH))
        template_data = reader.processar_templates()
    except Exception as e:
        print(f"❌ Erro ao processar templates: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # ============================================
    # FASE 2: CARREGAR DADOS
    # ============================================
    print("\n" + "="*70)
    print("FASE 2: CARREGAMENTO DOS DADOS")
    print("="*70 + "\n")
    
    criar_pasta_saida()
    
    df = carregar_dados()
    if df is None:
        return
    
    mapeamento = carregar_mapeamento_superintendencias()
    if not mapeamento:
        print("⚠️  Aviso: Mapeamento não carregado.")
        return
    
    df = adicionar_coluna_superintendencia(df, mapeamento)
    grupos_super = agrupar_por_superintendencia_e_macro(df)
    
    print(f"✅ Dados carregados: {len(df)} registros")
    print(f"📋 Superintendências: {len(grupos_super)}")
    
    # ============================================
    # FASE 3: CALCULAR VARIÁVEIS DINÂMICAS
    # ============================================
    print("\n" + "="*70)
    print("FASE 3: CÁLCULO DE VARIÁVEIS")
    print("="*70 + "\n")
    
    variaveis = calcular_variaveis(df, grupos_super)
    
    print(f"   📊 Total de Metas: {variaveis['var_total_metas']}")
    print(f"   🔵 Metas CNJ: {variaveis['var_total_metas_cnj']}")
    print(f"   🟡 Metas TJMG: {variaveis['var_total_metas_tjmg']}")
    print(f"   📁 Macrodesafios: {variaveis['var_total_macros']}")
    print(f"   🟢 % Faixa Verde: {variaveis['var_percentual_verde']}%")
    print(f"   🟡 % Faixa Amarela: {variaveis['var_percentual_amarelo']}%")
    print(f"   🔴 % Faixa Vermelha: {variaveis['var_percentual_vermelho']}%")
    print(f"   🏢 Superintendências: {variaveis['var_total_superintendencias']}")
    print(f"   📅 Ano: {variaveis['var_ano_atual']}")
    
    # ============================================
    # FASE 4: GERAR DOCUMENTO
    # ============================================
    print("\n" + "="*70)
    print("FASE 4: GERAÇÃO DO DOCUMENTO")
    print("="*70 + "\n")
    
    try:
        generator = DocumentGenerator(template_data, df, grupos_super, variaveis)
        doc = generator.gerar_documento()
    except Exception as e:
        print(f"❌ Erro ao gerar documento: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # ============================================
    # FASE 5: SALVAR
    # ============================================
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome_arquivo = f"Relatorio_Metas_Estrategicas_{timestamp}.docx"
    caminho_completo = os.path.join(Config.PASTA_SAIDA, nome_arquivo)
    
    doc.save(caminho_completo)
    
    print("\n" + "="*70)
    print("✅ RELATÓRIO GERADO COM SUCESSO!")
    print("="*70)
    print(f"\n📁 Localização: {os.path.abspath(caminho_completo)}")
    print(f"📊 Total de registros: {len(df)}")
    print(f"📄 Estrutura: {template_data['metadados']['total_titulos']} títulos")
    print(f"📝 Conteúdo: {template_data['metadados']['total_blocos']} blocos")
    
    if template_data['metadados']['avisos']:
        print(f"\n⚠️  Avisos: {len(template_data['metadados']['avisos'])}")
        for aviso in template_data['metadados']['avisos']:
            print(f"   • {aviso}")
    
    print("\n" + "="*70 + "\n")


# ============================================
# EXECUÇÃO
# ============================================

if __name__ == "__main__":
    try:
        gerar_relatorio_com_templates()
    except KeyboardInterrupt:
        print("\n\n⚠️  Geração cancelada pelo usuário.")
    except Exception as e:
        print(f"\n❌ ERRO INESPERADO: {e}")
        import traceback
        traceback.print_exc()
