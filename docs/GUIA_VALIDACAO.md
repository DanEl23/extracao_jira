# 🔍 GUIA DE VALIDAÇÃO - Extrator Unificado

## 📋 Objetivo
Garantir que o `extrator_unificado.py` gera **exatamente os mesmos dados** que os scripts originais.

---

## 🚀 PASSO 1: Backup dos Arquivos Atuais

```powershell
# Criar pasta de backup
New-Item -ItemType Directory -Path "exports\backup_original" -Force

# Fazer backup dos arquivos originais
Copy-Item "exports\dados_exportados_jira.xlsx" "exports\backup_original\" -Force
Copy-Item "exports\dados_exportados_jira_por_ano.xlsx" "exports\backup_original\" -Force
Copy-Item "exports\dicionario_metas_hierarquico.json" "exports\backup_original\" -Force
Copy-Item "exports\resultados_cnj.xlsx" "exports\backup_original\" -Force
```

---

## 🧪 PASSO 2: Testar Extrator Unificado

### Opção A: Apenas Jira (mais rápido para teste)
```powershell
.\venv\Scripts\python.exe extrator_unificado.py 2024 --sem-cnj
```

### Opção B: Completo (Jira + CNJ)
```powershell
.\venv\Scripts\python.exe extrator_unificado.py 2023 2024 2025
```

**Arquivos gerados:**
- `exports/dados_exportados_jira_por_ano.xlsx` (substitui o original)
- `exports/dicionario_metas_hierarquico.json` (substitui o original)
- `exports/resultados_cnj.xlsx` (se incluir CNJ)

---

## 🔍 PASSO 3: Validar Resultados

### 3.1 Validação Automática
```powershell
.\venv\Scripts\python.exe validar_extracao.py
```

### 3.2 Validação Manual no Excel

**Abra ambos os arquivos:**
- `exports/backup_original/dados_exportados_jira_por_ano.xlsx`
- `exports/dados_exportados_jira_por_ano.xlsx`

**Verifique:**
1. ✅ **Número de linhas** é igual?
2. ✅ **Colunas** são as mesmas?
3. ✅ **Primeiras 10 linhas** têm os mesmos valores?
4. ✅ **Campo META_ID** está preenchido corretamente?
5. ✅ **Campo Meta_apuração** está no formato correto?

---

## 📊 PASSO 4: Comparação Detalhada (PowerShell)

```powershell
# Comparar número de linhas
python -c "import pandas as pd; df1 = pd.read_excel('exports/backup_original/dados_exportados_jira_por_ano.xlsx'); df2 = pd.read_excel('exports/dados_exportados_jira_por_ano.xlsx'); print(f'Original: {len(df1)} linhas'); print(f'Novo: {len(df2)} linhas'); print(f'Diferença: {abs(len(df1) - len(df2))} linhas')"

# Comparar colunas
python -c "import pandas as pd; df1 = pd.read_excel('exports/backup_original/dados_exportados_jira_por_ano.xlsx'); df2 = pd.read_excel('exports/dados_exportados_jira_por_ano.xlsx'); print('Colunas originais:', len(df1.columns)); print('Colunas novas:', len(df2.columns)); print('Apenas no original:', set(df1.columns) - set(df2.columns)); print('Apenas no novo:', set(df2.columns) - set(df1.columns))"

# Comparar primeiras linhas
python -c "import pandas as pd; df1 = pd.read_excel('exports/backup_original/dados_exportados_jira_por_ano.xlsx'); df2 = pd.read_excel('exports/dados_exportados_jira_por_ano.xlsx'); print('=== CHAVES ORIGINAIS ==='); print(df1[['META_ID', 'Chave', 'Resumo']].head(3)); print('\n=== CHAVES NOVAS ==='); print(df2[['META_ID', 'Chave', 'Resumo']].head(3))"
```

---

## ✅ CRITÉRIOS DE SUCESSO

### Jira (dados_exportados_jira_por_ano.xlsx)
- [ ] Número de linhas: **IGUAL** (±5 linhas aceitável se filtros de ano mudaram)
- [ ] Colunas: **IDÊNTICAS** (mesmos nomes e ordem)
- [ ] META_ID: **PREENCHIDO** corretamente
- [ ] Nº_Meta: **PREENCHIDO** quando disponível
- [ ] Meta_apuração: **FORMATO CORRETO** `[CHAVE] Resumo`

### JSON (dicionario_metas_hierarquico.json)
- [ ] Estrutura hierárquica: **Pai → Filhos**
- [ ] Número de metas raiz: **SIMILAR** ao original

### CNJ (resultados_cnj.xlsx)
- [ ] Número de metas: **10 metas** coletadas
- [ ] Colunas: Meta, Subtítulo, Descrição, Categoria, Resultado, Data
- [ ] Todas as metas de 1 a 10 presentes

---

## 🐛 PROBLEMAS COMUNS E SOLUÇÕES

### Problema 1: Mais linhas no novo arquivo
**Causa:** Extração de mais anos ou mudanças no filtro JQL  
**Solução:** Verificar se os anos configurados são os mesmos

### Problema 2: Coluna extra no novo arquivo
**Causa:** Campo novo detectado no HTML do Jira  
**Solução:** Verificar se é campo relevante ou remover

### Problema 3: META_ID vazio em algumas linhas
**Causa:** Tickets que são "pais" (não têm parent_issue_key)  
**Solução:** Esperado - metas raiz não têm META_ID

---

## 📝 REPORT DE VALIDAÇÃO

Após os testes, preencha:

```
Data do teste: ___________
Versão testada: extrator_unificado.py

RESULTADOS:
[ ] Jira Simples: ✅ OK / ❌ FALHOU
[ ] Jira Anual: ✅ OK / ❌ FALHOU  
[ ] CNJ: ✅ OK / ❌ FALHOU
[ ] JSON Hierárquico: ✅ OK / ❌ FALHOU

OBSERVAÇÕES:
_________________________________
_________________________________
_________________________________

APROVADO PARA PRODUÇÃO: [ ] SIM / [ ] NÃO
```

---

## 🔄 ROLLBACK (Se necessário)

```powershell
# Restaurar arquivos originais
Copy-Item "exports\backup_original\*" "exports\" -Force
```

---

## 📞 PRÓXIMOS PASSOS

1. ✅ Validação completa aprovada?
2. 🗑️ Deprecar scripts antigos:
   - `extracao_jira.py`
   - `extracao_anual_jira.py`
   - `teste_extracao_cnj.py`
3. 📚 Atualizar documentação
4. 🎉 Usar apenas `extrator_unificado.py`
