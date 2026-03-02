import pandas as pd
from bs4 import BeautifulSoup
import re
import os

def limpar_html(texto):
    """Remove tags HTML para limpeza do relatório Word."""
    if pd.isna(texto) or texto == "":
        return ""
    soup = BeautifulSoup(str(texto), 'html.parser')
    return soup.get_text(separator=' ', strip=True)

def extrair_numero(texto):
    """Extrai o primeiro número de uma string para ordenação correta."""
    if pd.isna(texto): return 999
    match = re.search(r'(\d+)', str(texto))
    return int(match.group(1)) if match else 999

def converter_valor_brasileiro(valor):
    """Converte strings brasileiras (ex: 55.266,00) para float puro."""
    if pd.isna(valor) or str(valor).strip() == "":
        return None
    v = str(valor).strip()
    # Se tem ponto e vírgula: 55.266,00 -> 55266.00
    if '.' in v and ',' in v:
        v = v.replace('.', '').replace(',', '.')
    # Se tem apenas vírgula: 87,5 -> 87.5
    elif ',' in v:
        v = v.replace(',', '.')
    # Se tem apenas ponto: 55.266 -> 55266 (milhar)
    elif '.' in v:
        partes = v.split('.')
        if len(partes[-1]) == 3:
            v = v.replace('.', '')
    try:
        return float(v)
    except:
        return None

def formatar_valor_final(valor, unidade=""):
    """
    Formatação Rigorosa para Relatório:
    1. Pontos de milhar sempre presentes (55.266).
    2. Vírgula decimal apenas se houver casas decimais (87,5).
    3. Símbolo de % apenas se a unidade for Percentual.
    """
    if pd.isna(valor) or str(valor).strip() == "" or valor is None:
        return ""
    
    try:
        val_f = float(valor)
        # Formato 'g' remove o .0 de inteiros automaticamente
        string_limpa = format(val_f, 'g')
        partes = string_limpa.split('.')
        
        # Formata parte inteira com ponto de milhar
        inteiro_com_ponto = "{:,}".format(int(partes[0])).replace(",", ".")
        
        # Monta o número final
        resultado = inteiro_com_ponto
        if len(partes) > 1:
            resultado += "," + partes[1]
            
        # Adiciona % se for meta percentual
        if str(unidade).lower() == "percentual":
            resultado += "%"
            
        return resultado
    except:
        return str(valor)

def unificar_grupo(group):
    """
    Regra de Unificação Estrita:
    - Valor Apurado: Coletado EXCLUSIVAMENTE da linha 'Apuração'.
    - Texto: Coletado da linha 'Dezembro' (se for mais extenso).
    """
    # Filtros por palavra-chave
    row_valor = group[group['Resumo'].str.contains('Apur', na=False, case=False)]
    row_texto = group[group['Resumo'].str.strip().str.lower() == 'dezembro']
    
    # 1. Definir linha base (âncora)
    if not row_valor.empty:
        # Pega a linha de Apuração. Valor Apurado fica o que estiver aqui (pode ser vazio).
        final_row = row_valor.iloc[0].copy()
    else:
        # Se não existe linha de Apuração, usamos Dezembro/Última, mas ZERAMOS o valor apurado.
        if not row_texto.empty:
            final_row = row_texto.iloc[0].copy()
        else:
            final_row = group.iloc[-1].copy()
        final_row['Valor Apurado'] = None

    # 2. Herança de Texto qualitativo (Sempre de Dezembro se houver mais conteúdo)
    if not row_texto.empty:
        txt_dez = row_texto.iloc[0].get('Informação complementar texto', '')
        txt_base = final_row.get('Informação complementar texto', '')
        if pd.notna(txt_dez) and len(str(txt_dez)) > len(str(txt_base)):
            final_row['Informação complementar texto'] = txt_dez

    return final_row

def processar_mala_direta():
    print("🚀 Iniciando processamento unificado de metas...")

    # 1. CARREGAMENTO
    try:
        df_ano = pd.read_excel('exports/dados_exportados_jira_por_ano.xlsx')
        df_apuracao = pd.read_excel('exports/dados_exportados_jira.xlsx')
    except Exception as e:
        print(f"❌ Erro ao ler arquivos: {e}")
        return

    # 2. LIMPEZA NUMÉRICA PARA PROCESSAMENTO
    df_apuracao['Valor Apurado'] = df_apuracao['Valor Apurado'].apply(converter_valor_brasileiro)
    df_apuracao['Valor da Meta'] = df_apuracao['Valor da Meta'].apply(converter_valor_brasileiro)

    # 3. MAPEAMENTO DE HERANÇA (Dados da Meta Pai)
    # Criamos dicionários para transferir atributos das metas principais para os meses
    map_base = df_ano.set_index('Chave')
    cols_mapping = {
        'Ano da Meta': 'Ano da Meta',
        'Macrodesafio': 'Macrodesafio',
        'Indicador Estratégico': 'Indicador',
        'Unidade Gestora': 'Unidade Gestora',
        'Polaridade': 'Polaridade',
        'Resumo': 'MetaKey',
        'Valor da Meta': 'Valor da Meta_Pai',
        'Iniciativas Estratégicas 2025': 'IE_Raw',
        'Unidade de Medida': 'Unidade_Medida'
    }

    for col_origem, col_destino in cols_mapping.items():
        df_apuracao[col_destino] = df_apuracao['META_ID'].map(map_base[col_origem].to_dict())

    # 4. TRATAMENTOS DE CAMPO (Iniciativas IE e Limpeza de Indicador)
    df_apuracao['Iniciativa'] = df_apuracao['IE_Raw'].apply(
        lambda x: f"IE {re.search(r'\d+', str(x)).group(0)}" if pd.notna(x) and re.search(r'\d+', str(x)) else ""
    )
    df_apuracao['Informação complementar texto'] = df_apuracao['Informação complementar'].apply(limpar_html)
    df_apuracao['Indicador'] = df_apuracao['Indicador'].astype(str).str.replace(r'^\d{4}\s*-\s*', '', regex=True)

    # 5. ARQUIVO DE CONFERÊNCIA (dados_exportados_jira_adaptado.xlsx)
    df_apuracao.to_excel('dados_exportados_jira_adaptado.xlsx', index=False)

    # 6. FILTRAGEM POR ANO E UNIFICAÇÃO
    df_2025 = df_apuracao[df_apuracao['Ano da Meta'] == 2025].copy()
    print(f"🔗 Unificando {len(df_2025['MetaKey'].unique())} metas...")
    
    df_final = df_2025.groupby('MetaKey', group_keys=False).apply(unificar_grupo).reset_index(drop=True)

    # 7. ORDENAÇÃO DUPLA (Macrodesafio + Número da Meta)
    df_final['macro_num'] = df_final['Macrodesafio'].apply(extrair_numero)
    df_final['meta_num'] = df_final['MetaKey'].apply(extrair_numero)
    df_final = df_final.sort_values(by=['macro_num', 'meta_num']).drop(columns=['macro_num', 'meta_num'])

    # 8. FORMATAÇÃO VISUAL (Ponto Milhar, Vírgula Decimal e %)
    # Aplicado apenas na saída final para o Excel de integração
    df_final['Valor Apurado'] = df_final.apply(
        lambda x: formatar_valor_final(x['Valor Apurado'], x['Unidade_Medida']), axis=1
    )
    df_final['Valor da Meta'] = df_final.apply(
        lambda x: formatar_valor_final(x['Valor da Meta_Pai'], x['Unidade_Medida']), axis=1
    )

    # 9. EXPORTAÇÃO FINAL
    colunas_finais = [
        'Macrodesafio', 'MetaKey', 'Ano da Meta', 'Indicador', 
        'Unidade Gestora', 'Polaridade', 'Valor da Meta', 'Valor Apurado', 
        'Iniciativa', 'Informação complementar texto'
    ]
    
    df_final[colunas_finais].to_excel('teste_integração.xlsx', index=False)
    print("✅ Sucesso! O arquivo 'teste_integração.xlsx' está pronto para o relatório.")

if __name__ == "__main__":
    processar_mala_direta()