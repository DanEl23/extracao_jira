# 📘 GUIA DE USO DOS TEMPLATES

## 🎯 Visão Geral

Os templates permitem controlar a estrutura e conteúdo do relatório **sem modificar código**.

### **Arquivos:**

1. **Sumario_Modelo.docx** - Define estrutura hierárquica do relatório
2. **Conteudo_Fonte.docx** - Contém textos, marcadores e variáveis

---

## 📝 SUMARIO_MODELO.DOCX

### **Formato:**

```
1 TÍTULO NÍVEL 1
1.1 Título Nível 2
1.2 Outro Título Nível 2
2 OUTRO TÍTULO NÍVEL 1
2.1 Subtítulo
2.1.1 Nível 3
```

### **Regras:**

✅ Use numeração hierárquica (1, 1.1, 1.1.1)  
✅ Ponto final no prefixo é opcional (1. ou 1)  
✅ Títulos devem ser EXATOS - serão mapeados no conteúdo  
❌ Não adicione números de página (gerados automaticamente)  
❌ Não use tabulações ou pontos de preenchimento  

### **Exemplo:**

```
1 INTRODUÇÃO
1.1 Contextualização
1.2 Objetivos do Relatório

2 METAS APROVADAS
2.1 Total de Metas Estratégicas
```

---

## 📄 CONTEUDO_FONTE.DOCX

### **Estrutura:**

O conteúdo deve seguir a mesma hierarquia do sumário. Cada título do sumário deve ter uma seção correspondente no conteúdo.

### **Marcadores de Tabelas Dinâmicas:**

Estes marcadores são substituídos por tabelas geradas automaticamente a partir dos dados:

| Marcador | Descrição |
|----------|-----------|
| `[INSERIR_TABELA_HISTORICA]` | Tabela de metas aprovadas (4 anos) |
| `[INSERIR_TABELA_MACRODESAFIO]` | Distribuição por macrodesafio |
| `[INSERIR_TABELA_MONITORAMENTO]` | Resultado do monitoramento (faixas) |
| `[INSERIR_TABELA_CNJ]` | Metas nacionais do CNJ |
| `[INICIAR_SECAO_SUPERINTENDENCIAS]` | Seções detalhadas por superintendência |

**Uso:**

```
2.1 Total de Metas Estratégicas

[INSERIR_TABELA_HISTORICA]

O quadro acima demonstra a evolução...
```

### **Marcadores de Variáveis:**

Estes marcadores são substituídos por valores calculados automaticamente:

| Marcador | Descrição | Exemplo |
|----------|-----------|---------|
| `[NUMERO_METAS_2025]` | Total de metas do ano atual | 57 |
| `[NUMERO_METAS_CNJ]` | Total de metas nacionais CNJ | 9 |
| `[NUMERO_METAS_TJMG]` | Total de metas institucionais TJMG | 48 |
| `[NUMERO_MACRODESAFIOS]` | Total de macrodesafios | 6 |
| `[PERCENTUAL_VERDE]` | % de metas na faixa verde | 75.5 |
| `[VAR_ANO_ATUAL]` | Ano atual do relatório | 2025 |

**Uso:**

```
Foram aprovadas [NUMERO_METAS_2025] metas para o ano de [VAR_ANO_ATUAL].
```

### **Formatação Especial:**

#### **1. Texto com Destaque (Negrito + Vermelho)**

Inicie a linha com `#` para aplicar formatação de destaque:

```
#Este texto será exibido em negrito e cor vermelha
```

#### **2. Listas Numeradas**

Use marcadores de controle:

```
Os principais objetivos são:

[INICIAR_LISTA_NUMERADA]
Apresentar panorama geral das metas
Analisar resultados apurados
Identificar oportunidades de melhoria
[FINALIZAR_LISTA_NUMERADA]
```

Resultado:
```
1. Apresentar panorama geral das metas
2. Analisar resultados apurados
3. Identificar oportunidades de melhoria
```

#### **3. Listas com Marcadores (Bullets)**

```
[INICIAR_LISTA_MARCADORES]
Item A
Item B
Item C
[FINALIZAR_LISTA_MARCADORES]
```

Resultado:
```
• Item A
• Item B
• Item C
```

---

## 🔄 WORKFLOW DE USO

### **Passo 1: Ajustar Estrutura**

1. Abra `Sumario_Modelo.docx`
2. Adicione, remova ou reordene títulos conforme necessário
3. Mantenha numeração hierárquica consistente
4. Salve o arquivo

### **Passo 2: Ajustar Conteúdo**

1. Abra `Conteudo_Fonte.docx`
2. Edite os textos conforme necessário
3. Use marcadores `[INSERIR_TABELA_*]` onde tabelas devem aparecer
4. Use marcadores `[VARIAVEL]` onde valores dinâmicos devem aparecer
5. Salve o arquivo

### **Passo 3: Gerar Relatório**

```bash
python gerador_relatorio_com_templates.py
```

O sistema irá:
- ✅ Ler a estrutura do sumário
- ✅ Ler o conteúdo mapeado
- ✅ Carregar dados do Excel
- ✅ Calcular variáveis dinâmicas
- ✅ Gerar documento final completo

**Nota:** Para gerar relatório SEM templates (modo antigo), use:
```bash
python gerador_relatorio.py
```

---

## ⚠️ CUIDADOS IMPORTANTES

### **❌ NÃO FAÇA:**

- ❌ Não altere os marcadores `[MARCADOR]` (copie exatamente como está)
- ❌ Não adicione espaços dentro dos marcadores: `[ MARCADOR ]` ❌
- ❌ Não use marcadores inexistentes
- ❌ Não remova títulos do sumário se eles têm conteúdo

### **✅ FAÇA:**

- ✅ Mantenha títulos idênticos entre Sumario e Conteudo
- ✅ Use marcadores exatamente como documentado
- ✅ Teste após alterações estruturais
- ✅ Valide correspondência sumário ↔ conteúdo

---

## 🐛 TROUBLESHOOTING

### **Erro: "Título sem conteúdo"**

**Causa:** Título existe no sumário mas não no conteúdo.

**Solução:** Adicione seção correspondente no `Conteudo_Fonte.docx` ou remova do sumário.

### **Erro: "Conteúdo sem título no sumário"**

**Causa:** Seção existe no conteúdo mas não no sumário.

**Solução:** Adicione título correspondente no `Sumario_Modelo.docx` ou remova seção do conteúdo.

### **Variável não substituída (aparece [VARIAVEL] no relatório)**

**Causa:** Marcador de variável incorreto ou não implementado.

**Solução:** Verifique lista de marcadores válidos e corrija a digitação.

### **Tabela não aparece**

**Causa:** Marcador de tabela incorreto ou mal posicionado.

**Solução:** Verifique se o marcador está em uma linha separada, sem espaços extras.

---

## 📊 EXEMPLO COMPLETO

### **Sumario_Modelo.docx:**

```
1 INTRODUÇÃO
1.1 Contextualização

2 METAS APROVADAS
2.1 Total de Metas
```

### **Conteudo_Fonte.docx:**

```
1 INTRODUÇÃO

Este relatório apresenta os resultados de [NUMERO_METAS_2025] metas.

1.1 Contextualização

O processo de monitoramento tem como objetivo...

2 METAS APROVADAS

[INSERIR_TABELA_HISTORICA]

A tabela acima demonstra...

2.1 Total de Metas

#Destaque: [PERCENTUAL_VERDE]% das metas foram cumpridas!
```

### **Resultado Final:**

O sistema gera um documento Word com:
- ✅ Estrutura do sumário
- ✅ Textos formatados
- ✅ Tabelas dinâmicas inseridas
- ✅ Variáveis substituídas por valores reais
- ✅ Formatação especial aplicada

---

## 📞 SUPORTE

Para dúvidas ou problemas:

1. Consulte a documentação completa em `docs/PROPOSTA_IMPLEMENTACAO_TEMPLATES.md`
2. Verifique exemplos neste diretório
3. Valide templates antes de gerar relatório
4. Revise logs de processamento no console

---

**Versão:** 1.0  
**Última Atualização:** Dezembro/2025  
**Autor:** Sistema de Geração de Relatórios - TJMG
