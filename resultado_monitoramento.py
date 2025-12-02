"""
Módulo para calcular estatísticas de resultado do monitoramento de metas
"""

import pandas as pd


def calcular_resultado_monitoramento():
    """
    Calcula a distribuição de metas por faixa de cumprimento.
    Inclui metas nacionais (resultados_cnj.xlsx) e institucionais (teste_integração.xlsx).
    
    Returns:
        dict: {
            'maior_100': int,
            'entre_70_100': int,
            'abaixo_70': int,
            'sem_apuracao': int,
            'total': int
        }
    """
    # Inicializar contadores
    maior_100 = 0
    entre_70_100 = 0
    abaixo_70 = 0
    sem_apuracao = 0
    
    # ===== PROCESSAR METAS NACIONAIS (CNJ) =====
    try:
        df_cnj = pd.read_excel('exports/resultados_cnj.xlsx')
        
        for _, row in df_cnj.iterrows():
            resultado = row.get('Resultado', None)
            
            # Verificar se tem resultado
            if pd.isna(resultado):
                sem_apuracao += 1
                continue
            
            try:
                # Converter resultado (formato: "108,62%")
                if isinstance(resultado, str):
                    # Remover % e converter vírgula para ponto
                    percentual_str = resultado.replace('%', '').replace(',', '.')
                    percentual = float(percentual_str)
                else:
                    percentual = float(resultado)
                
                # Classificar
                if percentual >= 100:
                    maior_100 += 1
                elif percentual >= 70:
                    entre_70_100 += 1
                else:
                    abaixo_70 += 1
                    
            except (ValueError, TypeError):
                sem_apuracao += 1
                continue
                
    except Exception as e:
        print(f"⚠️  Erro ao processar metas nacionais (CNJ): {e}")
    
    # ===== PROCESSAR METAS INSTITUCIONAIS (TJMG) =====
    try:
        df_tjmg = pd.read_excel('exports/teste_integração.xlsx')
        
        for _, row in df_tjmg.iterrows():
            valor_apurado = row.get('Valor Apurado', None)
            valor_meta = row.get('Valor da Meta', None)
            
            # Verificar se tem apuração
            if pd.isna(valor_apurado) or pd.isna(valor_meta):
                sem_apuracao += 1
                continue
            
            try:
                # Converter valor apurado (pode estar como string com vírgula)
                if isinstance(valor_apurado, str):
                    val_apurado_float = float(valor_apurado.replace('.', '').replace(',', '.'))
                else:
                    val_apurado_float = float(valor_apurado)
                
                # Converter valor da meta
                val_meta_float = float(valor_meta)
                
                # Calcular percentual
                if val_meta_float > 0:
                    percentual = (val_apurado_float / val_meta_float) * 100
                    
                    # Classificar
                    if percentual >= 100:
                        maior_100 += 1
                    elif percentual >= 70:
                        entre_70_100 += 1
                    else:
                        abaixo_70 += 1
                else:
                    sem_apuracao += 1
                    
            except (ValueError, TypeError):
                sem_apuracao += 1
                continue
                
    except Exception as e:
        print(f"⚠️  Erro ao processar metas institucionais (TJMG): {e}")
    
    total = maior_100 + entre_70_100 + abaixo_70 + sem_apuracao
    
    return {
        'maior_100': maior_100,
        'entre_70_100': entre_70_100,
        'abaixo_70': abaixo_70,
        'sem_apuracao': sem_apuracao,
        'total': total
    }


if __name__ == "__main__":
    # Teste
    resultado = calcular_resultado_monitoramento()
    print("\n📊 RESULTADO DO MONITORAMENTO DE METAS")
    print("=" * 50)
    print(f"Metas com resultado maior ou igual 100%: {resultado['maior_100']}")
    print(f"Metas com resultado entre 70% e 100%: {resultado['entre_70_100']}")
    print(f"Metas com resultado abaixo de 70%: {resultado['abaixo_70']}")
    print(f"Metas sem apuração até outubro/2024: {resultado['sem_apuracao']}")
    print(f"Total: {resultado['total']}")
