# 📊 LISTA COMPLETA DE MARCADORES E VARIÁVEIS

## 🔷 MARCADORES DE TABELAS DINÂMICAS

### Tabelas Geradas Automaticamente

| Marcador | Descrição | Localização Sugerida |
|----------|-----------|---------------------|
| `[INSERIR_TABELA_HISTORICA]` | Tabela de metas aprovadas (2021-2024 + ano atual) | Seção "2.1 Total de Metas Estratégicas" |
| `[INSERIR_TABELA_MACRODESAFIO]` | Distribuição de metas por macrodesafio (4 anos) | Seção "2.2 Distribuição por Macrodesafio" |
| `[INSERIR_TABELA_MONITORAMENTO]` | Resultado do monitoramento por faixa de desempenho | Seção "3.1 Indicadores de Cumprimento" |
| `[INSERIR_TABELA_CNJ]` | Metas nacionais do CNJ com alvos e resultados | Seção "4.1 Visão Geral" |
| `[INICIAR_SECAO_SUPERINTENDENCIAS]` | Seções detalhadas com tabelas em paisagem | Seção "5 SUPERINTENDÊNCIAS" |

**Observação:** Cada marcador deve estar sozinho em uma linha separada.

---

## 🔢 VARIÁVEIS DINÂMICAS

### Variáveis Calculadas Automaticamente

| Marcador | Tipo | Descrição | Exemplo de Valor |
|----------|------|-----------|------------------|
| `[VAR_ANO_ATUAL]` | Numérico | Ano atual do relatório | 2025 |
| `[NUMERO_METAS_2025]` | Numérico | Total de metas aprovadas no ano | 57 |
| `[NUMERO_METAS_CNJ]` | Numérico | Total de metas nacionais do CNJ | 9 |
| `[NUMERO_METAS_TJMG]` | Numérico | Total de metas institucionais TJMG | 48 |
| `[NUMERO_MACRODESAFIOS]` | Numérico | Total de macrodesafios distintos | 6 |
| `[PERCENTUAL_VERDE]` | Decimal | % de metas na faixa verde (≥90%) | 75.5 |
| `[PERCENTUAL_AMARELO]` | Decimal | % de metas na faixa amarela (60-89%) | 15.8 |
| `[PERCENTUAL_VERMELHO]` | Decimal | % de metas na faixa vermelha (<60%) | 8.7 |
| `[TOTAL_SUPERINTENDENCIAS]` | Numérico | Total de superintendências com metas | 6 |

**Uso em texto:**

```
Foram aprovadas [NUMERO_METAS_2025] metas para o ano de [VAR_ANO_ATUAL], 
sendo [NUMERO_METAS_CNJ] metas nacionais do CNJ e [NUMERO_METAS_TJMG] 
metas institucionais do TJMG.
```

**Resultado após processamento:**

```
Foram aprovadas 57 metas para o ano de 2025, sendo 9 metas nacionais 
do CNJ e 48 metas institucionais do TJMG.
```

---

## 🎨 MARCADORES DE FORMATAÇÃO

### Controle de Estilo e Layout

| Marcador | Função | Exemplo |
|----------|--------|---------|
| `#Texto` | Aplica negrito + cor vermelha | `#Destaque: 75% das metas foram cumpridas!` |
| `[INICIAR_LISTA_NUMERADA]` | Inicia lista numerada (1. 2. 3...) | Ver seção abaixo |
| `[FINALIZAR_LISTA_NUMERADA]` | Finaliza lista numerada | Ver seção abaixo |
| `[INICIAR_LISTA_MARCADORES]` | Inicia lista com bullets (• • •) | Ver seção abaixo |
| `[FINALIZAR_LISTA_MARCADORES]` | Finaliza lista com bullets | Ver seção abaixo |
| `[QUEBRA_PAGINA]` | Insere quebra de página | (marcador único em linha) |

---

## 📝 EXEMPLOS DE USO

### Exemplo 1: Texto com Variáveis

**No template:**

```
O TJMG monitora [NUMERO_METAS_2025] metas estratégicas no ano de [VAR_ANO_ATUAL].

Destas, [PERCENTUAL_VERDE]% alcançaram resultado satisfatório (faixa verde), 
[PERCENTUAL_AMARELO]% apresentaram resultado intermediário (faixa amarela) e 
[PERCENTUAL_VERMELHO]% necessitam atenção especial (faixa vermelha).
```

**Resultado gerado:**

```
O TJMG monitora 57 metas estratégicas no ano de 2025.

Destas, 75.5% alcançaram resultado satisfatório (faixa verde), 15.8% 
apresentaram resultado intermediário (faixa amarela) e 8.7% necessitam 
atenção especial (faixa vermelha).
```

---

### Exemplo 2: Lista Numerada

**No template:**

```
Os objetivos deste relatório são:

[INICIAR_LISTA_NUMERADA]
Apresentar o panorama geral das metas estratégicas
Analisar os resultados apurados no período
Identificar metas que demandam atenção
Subsidiar decisões estratégicas dos gestores
[FINALIZAR_LISTA_NUMERADA]

Estes objetivos orientam a estrutura do documento.
```

**Resultado gerado:**

```
Os objetivos deste relatório são:

1. Apresentar o panorama geral das metas estratégicas
2. Analisar os resultados apurados no período
3. Identificar metas que demandam atenção
4. Subsidiar decisões estratégicas dos gestores

Estes objetivos orientam a estrutura do documento.
```

---

### Exemplo 3: Lista com Marcadores

**No template:**

```
As principais ações incluem:

[INICIAR_LISTA_MARCADORES]
Implementação de novas ferramentas
Capacitação de equipes
Aprimoramento de processos
Integração entre unidades
[FINALIZAR_LISTA_MARCADORES]
```

**Resultado gerado:**

```
As principais ações incluem:

• Implementação de novas ferramentas
• Capacitação de equipes
• Aprimoramento de processos
• Integração entre unidades
```

---

### Exemplo 4: Texto com Destaque

**No template:**

```
A análise dos dados revela importantes insights:

#Destaque Positivo: [PERCENTUAL_VERDE]% das metas atingiram resultado satisfatório!

Este resultado demonstra o comprometimento institucional com o planejamento estratégico.
```

**Resultado gerado:**

```
A análise dos dados revela importantes insights:

**Destaque Positivo: 75.5% das metas atingiram resultado satisfatório!** (em vermelho)

Este resultado demonstra o comprometimento institucional com o planejamento estratégico.
```

---

### Exemplo 5: Inserção de Tabela

**No template:**

```
2.1 Total de Metas Estratégicas

A evolução histórica das metas aprovadas é apresentada a seguir:

[INSERIR_TABELA_HISTORICA]

O quadro acima demonstra o crescimento do número de metas monitoradas ao longo 
dos últimos 4 anos, evidenciando o amadurecimento do processo de planejamento 
estratégico do TJMG.
```

**Resultado gerado:**

```
2.1 Total de Metas Estratégicas

A evolução histórica das metas aprovadas é apresentada a seguir:

[AQUI APARECE A TABELA FORMATADA COM OS DADOS REAIS]

┌──────┬───────────┬──────────┐
│ Ano  │ Metas CNJ │ Metas TJ │
├──────┼───────────┼──────────┤
│ 2021 │     10    │    40    │
│ 2022 │      9    │    42    │
│ 2023 │      9    │    45    │
│ 2024 │      9    │    46    │
│ 2025 │      9    │    48    │
└──────┴───────────┴──────────┘

O quadro acima demonstra o crescimento do número de metas monitoradas...
```

---

## 🔒 REGRAS IMPORTANTES

### ✅ FAZER:

1. **Manter marcadores exatos**: Copie e cole os marcadores, não digite manualmente
2. **Linha separada para marcadores de tabela**: `[INSERIR_TABELA_X]` deve estar sozinho
3. **Fechar listas**: Sempre use `[FINALIZAR_...]` após `[INICIAR_...]`
4. **Testar após edição**: Execute o gerador para validar alterações

### ❌ NÃO FAZER:

1. ❌ Adicionar espaços: `[ VARIAVEL ]` ou `[VARIAVEL ]`
2. ❌ Alterar nomes: `[NUMERO_META]` ao invés de `[NUMERO_METAS_2025]`
3. ❌ Misturar formatação: Não use `#` em listas ou tabelas
4. ❌ Esquecer de fechar listas: Sempre finalize com `[FINALIZAR_...]`

---

## 🎯 SEQUÊNCIA RECOMENDADA DE SEÇÕES

Para um relatório completo, sugere-se a seguinte estrutura:

```
1 INTRODUÇÃO
  Texto introdutório
  
2 METAS APROVADAS
  2.1 Total → [INSERIR_TABELA_HISTORICA]
  2.2 Distribuição → [INSERIR_TABELA_MACRODESAFIO]
  
3 RESULTADO DO MONITORAMENTO
  3.1 Indicadores → [INSERIR_TABELA_MONITORAMENTO]
  
4 METAS NACIONAIS DO CNJ
  4.1 Visão Geral → [INSERIR_TABELA_CNJ]
  
5 SUPERINTENDÊNCIAS
  → [INICIAR_SECAO_SUPERINTENDENCIAS]
  
6 CONSIDERAÇÕES FINAIS
  Texto conclusivo
```

---

## 📞 SUPORTE E DÚVIDAS

Se encontrar problemas:

1. **Verifique a sintaxe** dos marcadores
2. **Consulte o README_TEMPLATES.md** para instruções detalhadas
3. **Revise os exemplos** neste arquivo
4. **Teste incrementalmente**: Adicione marcadores um por vez

---

**Versão:** 1.0  
**Última Atualização:** Dezembro/2025  
**Compatível com:** Sistema de Relatórios TJMG v4.0+
