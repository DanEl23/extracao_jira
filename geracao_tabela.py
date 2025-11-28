import json
import pandas as pd
import os
import sys

def exportar_dicionario_para_excel(nome_arquivo_json, nome_arquivo_excel):
    """
    Carrega o dicionário hierárquico, achata (flatten) a estrutura
    (Meta Pai -> Filhos) e exporta o resultado para um arquivo Excel.
    """
    # 1. Carrega o dicionário
    if not os.path.exists(nome_arquivo_json):
        print(f"❌ Erro: Arquivo JSON '{nome_arquivo_json}' não encontrado no diretório.")
        print("Certifique-se de ter executado o script de extração anteriormente.")
        return
    
    with open(nome_arquivo_json, 'r', encoding='utf-8') as f:
        dicionario_metas = json.load(f)

    dados_planos = []
    
    print(f"🔄 Iniciando o achatamento de {len(dicionario_metas)} Meta(s) Raiz...")
    
    # 2. Itera sobre o dicionário e achata os dados
    for chave_pai, meta_pai in dicionario_metas.items():
        dados_pai = meta_pai['Dados']
        filhos = meta_pai['Filhos']
        
        # Prefixa os campos do Pai (Chave, Resumo, etc.) para evitar conflito com os do Filho
        # Ex: 'Chave' do Pai se torna 'PAI_Chave'
        dados_pai_prefixados = {f"PAI_{k}": v for k, v in dados_pai.items()}
        
        if not filhos:
            # Se o Pai não tem filhos, adicionamos apenas os dados do Pai à lista
            dados_planos.append(dados_pai_prefixados)
        else:
            # Para cada Filho (Apuração), criamos uma linha no Excel
            for filho in filhos:
                dados_filho = filho['Dados']
                
                # Combina os dados do Filho com os dados prefixados do Pai
                registro_completo = {**dados_pai_prefixados, **dados_filho}
                
                dados_planos.append(registro_completo)

    # 3. Cria e salva o DataFrame
    df = pd.DataFrame(dados_planos)

    # Define a ordem das colunas para melhor visualização (Pai primeiro, depois Filho)
    colunas_principais = [
        # Campos do Pai
        'PAI_Chave', 'PAI_Resumo', 'PAI_META_ID', 'PAI_Nº_Meta', 
        # Campos do Filho (Apuração)
        'Chave', 'Resumo', 'Meta_apuração', 'META_ID',
    ]
    
    # Colunas que realmente existem no DataFrame
    colunas_existentes = [col for col in colunas_principais if col in df.columns]
    
    # Adiciona todas as outras colunas (campos customizados, datas, etc.)
    outras_colunas = [col for col in df.columns if col not in colunas_existentes]
    
    # Aplica a ordem das colunas
    df = df[colunas_existentes + outras_colunas]
    
    df.to_excel(nome_arquivo_excel, index=False)
    
    print(f"\n💾 Tabela Excel criada com sucesso em: {nome_arquivo_excel}")
    print(f"📊 Total de linhas (registros de Apuração/Itens): {len(df)}")


# ==================== EXECUTOR ====================

if __name__ == '__main__':
    # NOME DO ARQUIVO JSON GERADO PELO SCRIPT DE EXTRAÇÃO ANTERIOR
    NOME_JSON = 'dicionario_metas_hierarquico_lote_unico.json' 
    # NOME DO NOVO ARQUIVO EXCEL
    NOME_EXCEL = 'tabela_jira_hierarquica_final.xlsx'
    
    print("🚀 INICIANDO CONVERSÃO DO DICIONÁRIO JSON PARA EXCEL")
    exportar_dicionario_para_excel(NOME_JSON, NOME_EXCEL)