# 📊 Sistema de Extração e Geração de Relatórios - TJMG

Sistema automatizado para extração de dados de metas (Jira e CNJ) e geração de relatórios estratégicos em Word.

---

## 🚀 Fluxo de Produção Principal

```
1. EXTRAÇÃO CNJ
   extracao_cnj.py
   └─> exports/resultados_cnj.xlsx

2. EXTRAÇÃO JIRA
   extrator_unificado.py
   └─> exports/teste_integração.xlsx

3. GERAÇÃO DE RELATÓRIO (MODULARIZADO)
   gerador_relatorio_novo.py
   ├─> Usa módulos de gerador_relatorio/
   └─> relatorios_gerados/Relatorio_Metas_Estrategicas_[timestamp].docx
```

---

## 📁 Estrutura de Arquivos

### **Arquivos Principais (Raiz)**

#### Extração de Dados
- `extracao_cnj.py` - Extrator PowerBI para metas nacionais CNJ
- `extrator_unificado.py` - Extrator unificado Jira + CNJ (em desenvolvimento)

#### Geração de Relatórios
- **`gerador_relatorio.py`** - 🆕 Gerador principal MODULARIZADO (versão 3.0) ⭐
- `gerador_relatorio_gui.py` - Interface gráfica alternativa (Tkinter)

#### Dados e Cálculos
- **`base_dados_fixos.py`** - 🆕 Wrapper compatibilidade (importa do package base_dados/) ⭐
- `resultado_monitoramento.py` - Classificação de metas por faixas de desempenho

#### Configuração
- `requirements.txt` - Dependências do projeto

---

### **📦 Package `gerador_relatorio/` (MODULARIZADO)**

Novo package estruturado em módulos especializados:

```
gerador_relatorio/
├── __init__.py              # Exporta todas as funções
├── config.py                # Configurações, constantes, cores
├── data_loader.py           # Carregamento e agrupamento de dados
├── formatters.py            # Formatação de valores e cores
├── styles.py                # Estilos de células, bordas, backgrounds
├── document_builder.py      # Criação de documento e seções
├── table_historico.py       # Tabelas de histórico de metas
├── table_cnj.py             # Tabela de metas nacionais CNJ
├── table_monitoramento.py   # Tabela de resultado de monitoramento
└── table_superintendencia.py # Tabelas e seções de superintendências
```

**Benefícios da Modularização:**
- ✅ Código organizado por responsabilidade
- ✅ Fácil manutenção e debugging
- ✅ Reutilização de componentes
- ✅ Testes unitários facilitados

---

### **Pastas Organizacionais**

#### 📂 `legacy/`
Arquivos legados mantidos para referência (substituídos por `extrator_unificado.py`):
- `extracao_jira.py` - Extrator Jira original (single-year)
- `extracao_anual_jira.py` - Extrator Jira multi-year

#### 📂 `utils/`
Scripts utilitários e ferramentas auxiliares:
- `geracao_tabela.py` - Conversor JSON hierárquico → Excel plano
- `atualizar_historico_exemplo.py` - Atualizador de dados históricos
- `validar_extracao.py` - Validador de extrações (QA)

#### 📂 `notebooks/`
Notebooks Jupyter para análise e processamento de dados:
- `tratamento_2.ipynb` - Mapeamento de anos entre bases
- `Tratamento_mala_direta.ipynb` - Preparação de dados para mail merge

#### 📂 `docs/`
Documentação do projeto:
- `GUIA_VALIDACAO.md` - Guia de validação do extrator unificado
- `README.md` - Este arquivo

#### 📂 `exports/`
Arquivos Excel gerados pelas extrações:
- `resultados_cnj.xlsx` - Dados CNJ
- `teste_integração.xlsx` - Dados Jira processados

#### 📂 `relatorios_gerados/`
Documentos Word gerados pelo sistema

#### 📂 `templates/` 🆕
Templates Word para geração de relatórios:
- `Sumario_Modelo.docx` - Estrutura hierárquica do relatório
- `Conteudo_Fonte.docx` - Textos e marcadores de conteúdo
- `README_TEMPLATES.md` - Guia completo de uso dos templates
- `LISTA_MARCADORES.md` - Referência de marcadores e variáveis
- `criar_templates.py` - Script para recriar templates

#### 📂 `TESTE/`
Arquivos de teste e templates XML

---

## 🔧 Como Usar

### 1. Instalar Dependências
```powershell
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### 2. Extrair Dados CNJ
```powershell
python extracao_cnj.py
```
- Aguarde o navegador abrir
- Os dados serão salvos em `exports/resultados_cnj.xlsx`

### 3. Extrair Dados Jira
```powershell
# Opção A: Extrator unificado (em desenvolvimento)
python extrator_unificado.py 2024 2025

# Opção B: Extrator legado
python legacy/extracao_anual_jira.py
```

### 4. Gerar Relatório
```powershell
# Via linha de comando
python gerador_relatorio.py

# Via interface gráfica
python gerador_relatorio_gui.py
```

---

## 📊 Arquivos de Dados Principais

### Entrada (Excel)
- `exports/resultados_cnj.xlsx` - 9 metas nacionais CNJ
- `exports/teste_integração.xlsx` - 48 metas institucionais TJMG

### Saída (Word)
- `relatorios_gerados/Relatorio_Metas_Estrategicas_YYYYMMDD_HHMMSS.docx`

---

## 🗃️ Dados Centralizados

### `base_dados_fixos.py`

**Dicionários:**
- `META_SUPERINTENDENCIA` - Mapeia 64 metas para 6 superintendências
- `METAS_CNJ_ALVOS` - Targets para 21 categorias CNJ
- `METAS_CNJ_PARA_MACRODESAFIO` - Mapeia 10 metas CNJ para macrodesafios
- `HISTORICO_METAS_APROVADAS` - Histórico 2021-2024
- `HISTORICO_METAS_POR_MACRODESAFIO` - Histórico por macrodesafio 2022-2024

**Funções:**
- `calcular_cumprimento_metas_cnj()` - Aplica fórmula de ajuste CNJ
- `calcular_metas_atuais_por_macrodesafio()` - Conta metas CNJ + TJMG
- `atualizar_historico_com_ano_atual()` - Atualiza histórico com dados atuais

---

## 🎨 Formatação do Relatório

### Seção Retrato (Portrait)
- Tabela de Metas Nacionais CNJ
- Tabela de Resultado do Monitoramento (4 faixas)
- Tabela Histórica por Macrodesafio (2022-2025)
- Tabela de Total de Metas Aprovadas (2021-2025)

### Seção Paisagem (Landscape)
- Tabelas por Superintendência (6 tabelas)
- Detalhamento de cada meta com indicador, polaridade, valores

---

## 📈 Estatísticas do Projeto

- **Total de Metas:** 64 (9 CNJ + 55 TJMG, sendo 48 institucionais ativas)
- **Superintendências:** 6
- **Macrodesafios:** 11
- **Faixas de Desempenho:** 4 (≥100%, 70-100%, <70%, sem apuração)

---

## 🐍 Requisitos Python

- Python 3.8+
- pandas
- python-docx
- selenium
- webdriver-manager
- beautifulsoup4

---

## 👥 Suporte

Para dúvidas sobre o sistema, consulte a documentação em `docs/` ou os comentários no código-fonte.

---

**Última atualização:** Dezembro 2024
