"""
EXTRATOR UNIFICADO DE DADOS - TJMG
Versão: 2.1 (Refatorado)
Integra extração de Jira e CNJ (Lógica atualizada do extracao_cnj.py)

Modos de Operação:
1. Jira Simples: Extração única do estado atual
2. Jira Anual: Extração filtrada por múltiplos anos
3. CNJ: Extração do painel de metas nacionais
4. Completo: Executa Jira + CNJ
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
import pandas as pd
import time
import traceback
import re
import json
import argparse
import sys
from urllib.parse import quote_plus
from datetime import datetime
from pathlib import Path


# ============================================
# CONFIGURAÇÕES GLOBAIS
# ============================================

class Config:
    """Configurações centralizadas do sistema"""
    
    # Jira
    URL_JIRA = "https://tjmg.atlassian.net/"
    JQL_BASE = "project = ASPLAGMETA ORDER BY created DESC"
    ANOS_EXTRACAO = ["2022", "2023", "2024", "2025", "2026"]
    
    # CNJ
    URL_CNJ = "https://justica-em-numeros.cnj.jus.br/painel-metas/"
    TRIBUNAL_CNJ = "TJMG"
    
    # Arquivos de Saída
    PASTA_SAIDA = "exports"
    ARQUIVO_JIRA_SIMPLES = "dados_exportados_jira.xlsx"
    ARQUIVO_JIRA_ANUAL = "dados_exportados_jira_por_ano.xlsx"
    ARQUIVO_JIRA_JSON = "dicionario_metas_hierarquico.json"
    ARQUIVO_CNJ = "resultados_cnj.xlsx"
    
    # Navegador
    NAVEGADOR = "edge"  # "edge" ou "chrome"
    TIMEOUT = 20


# ============================================
# CLASSE BASE - EXTRATOR
# ============================================

class ExtratorBase:
    """Classe base com funcionalidades comuns de navegação"""
    
    def __init__(self, config=None):
        self.config = config or Config()
        self.driver = None
        self.wait = None
        self.dados_extraidos = []
        
    def iniciar_navegador(self):
        """Inicializa o navegador (Edge ou Chrome)"""
        print(f"🌐 Iniciando {self.config.NAVEGADOR.upper()}...")
        
        if self.config.NAVEGADOR.lower() == "edge":
            options = webdriver.EdgeOptions()
            options.add_argument('--start-maximized')
            options.add_argument('--disable-blink-features=AutomationControlled')
            self.driver = webdriver.Edge(options=options)
        else:
            options = webdriver.ChromeOptions()
            options.add_argument('--start-maximized')
            options.add_argument('--disable-notifications')
            self.driver = webdriver.Chrome(options=options)
        
        self.wait = WebDriverWait(self.driver, self.config.TIMEOUT)
        print("✅ Navegador iniciado com sucesso!")
        
    def fechar(self):
        """Fecha o navegador"""
        if self.driver:
            self.driver.quit()
            print("🔴 Navegador fechado")
    
    def criar_pasta_saida(self):
        """Cria pasta de saída se não existir"""
        Path(self.config.PASTA_SAIDA).mkdir(exist_ok=True)


# ============================================
# EXTRATOR JIRA
# ============================================

class ExtratorJira(ExtratorBase):
    """Extrator especializado para Jira"""
    
    def __init__(self, config=None):
        super().__init__(config)
        self.janela_principal = None
        
    def login_manual_e_aguardar(self, jql_base=None):
        """Abre o Jira e aguarda login manual"""
        jql = jql_base or self.config.JQL_BASE
        
        self.driver.get(self.config.URL_JIRA)
        
        print("\n" + "="*60)
        print("🔐 MODO MANUAL DE LOGIN")
        print("="*60)
        print("1. Faça o LOGIN MANUALMENTE no navegador")
        print(f"2. Após o login, navegue para: {jql}")
        print("\n⏸️  Pressione ENTER após fazer login e carregar a lista de issues...")
        
        input()
        
        print("\n✅ Continuando a extração...")
        time.sleep(2)
        
        # Navega para a JQL
        jql_encoded = quote_plus(jql)
        url_busca = f"{self.config.URL_JIRA}issues/?jql={jql_encoded}"
        print(f"➡️  Navegando para filtro JQL...")
        self.driver.get(url_busca)
        
        try:
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "[data-testid='issue-navigator-list']"))
            )
            print("   ✅ Lista de issues carregada")
        except:
            print("   ⚠️  Lista demorou para carregar, tentando continuar...")
        
        self.janela_principal = self.driver.current_window_handle
        time.sleep(3)
    
    def navegar_para_jql(self, jql_encoded):
        """Navega para uma JQL específica (usado para resetar filtros)"""
        url_busca = f"{self.config.URL_JIRA}issues/?jql={jql_encoded}"
        self.driver.get(url_busca)
        try:
            self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "[data-testid='issue-navigator-list']"))
            )
        except:
            pass
        time.sleep(3)
    
    def aplicar_filtro_por_ano(self, nome_campo="Ano da Meta", valor_ano="2024"):
        """Aplica filtro customizado por ano na interface JQL"""
        print(f"\n⚙️  Aplicando filtro: {nome_campo} = {valor_ano}")
        time.sleep(3)
        try:
            more_filters_button_selector = "button[data-testid='jql-builder-basic.ui.jql-editor.add-filter']"
            more_filters_button = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, more_filters_button_selector))
            )
            more_filters_button.click()
            time.sleep(1)

            search_input_xpath = "//input[@aria-label='Search more filters']"
            search_input = self.wait.until(
                EC.presence_of_element_located((By.XPATH, search_input_xpath))
            )
            search_input.send_keys(nome_campo)
            time.sleep(1)

            field_option_xpath = f"//div[@role='option']//div[text()='{nome_campo}']"
            field_option = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, field_option_xpath))
            )
            field_option.click()
            time.sleep(2)

            value_input_xpath = f"//input[@aria-label='Search {nome_campo}']"
            value_input = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, value_input_xpath))
            )
            value_input.send_keys(valor_ano)
            time.sleep(1)

            value_option_xpath = f"//div[@role='option' and .//div[text()='{valor_ano}']]"
            value_option = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, value_option_xpath))
            )
            value_option.click()

            self.wait.until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, "[data-testid='issue-navigator.issue-list.content-loading-spinner']"))
            )
            print("   ✅ Filtro aplicado com sucesso!")
            time.sleep(2)
        except Exception as e:
            print(f"❌ Erro ao aplicar filtro: {e}")
            raise
    
    def exportar_detalhes_impressao(self):
        """Abre o menu de exportação e seleciona 'Detalhes de impressão'"""
        print("\n⚙️  Iniciando exportação...")
        
        try:
            export_button_selector = "button[data-testid='issue-navigator-action-export-issues.ui.filter-button--trigger']"
            export_button = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, export_button_selector))
            )
            export_button.click()
            time.sleep(1)
            
            print_details_link_xpath = "//a[@data-vc='link-item' and .//span[text()='Detalhes de impressão']]"
            print_details_link = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, print_details_link_xpath))
            )
            print_details_link.click()
            print("   ✅ 'Detalhes de impressão' clicado. Aguardando nova aba...")
            time.sleep(3)
            
        except Exception as e:
            print(f"❌ Erro durante a exportação: {e}")
            raise
    

    def processar_aba_exportacao(self):
        """
        Extração Universal com Lógica de Herança.
        Captura campos em tabelas 'tableBorder' e 'grid', e vincula nomes de metas aos tickets de apuração.
        """
        janelas = self.driver.window_handles
        if len(janelas) < 2:
            print("❌ Nova aba de exportação não detectada")
            return 0
        
        janela_exportacao = [w for w in janelas if w != self.janela_principal][0]
        self.driver.switch_to.window(janela_exportacao)
        print("\n🔄 Foco mudado para aba de exportação. Processando conteúdo...")
        
        html_content = self.driver.page_source
        start_time = time.time()
        
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            # Localiza os blocos principais de cada ticket
            blocos_tickets = soup.find_all('table', class_='tableBorder')
            num_tickets = len(blocos_tickets)
            print(f"📊 Encontrados {num_tickets} tickets para processamento.")
            
            for idx, tabela_inicio in enumerate(blocos_tickets):
                registro = {}
                
                # 1. CAPTURA DE IDENTIFICADORES E HERANÇA DO PAI
                # Extrai a chave do pai (META_ID)
                parent_key_tag = tabela_inicio.find('a', id='parent_issue_key')
                registro['META_ID'] = parent_key_tag.get_text(strip=True) if parent_key_tag else ""
                
                # Extrai o resumo do pai (Nome da Meta Principal)
                parent_summary_tag = tabela_inicio.find('a', id='parent_issue_summary')
                nome_meta_pai = parent_summary_tag.get_text(strip=True) if parent_summary_tag else ""
                registro['Nº_Meta'] = nome_meta_pai
                
                # 2. TRATAMENTO DO TÍTULO E RESUMO (H3)
                h3_element = tabela_inicio.find('h3', class_='formtitle')
                if h3_element:
                    titulo_texto = h3_element.get_text(separator=' ', strip=True)
                    # Extrai a Chave do ticket atual: ex [ASPLAGMETA-2814]
                    chave_match = re.search(r'\[([A-Z]+-\d+)\]', titulo_texto)
                    current_chave = chave_match.group(1) if chave_match else f"TICKET-{idx+1}"
                    registro['Chave'] = current_chave
                    
                    # Extrai o link do resumo
                    resumo_link = h3_element.find('a')
                    resumo_original = resumo_link.get_text(strip=True) if resumo_link else ""
                    
                    # Se o resumo for genérico ("Apurado no período"), herda o nome da meta do pai
                    if "Apurado no período" in resumo_original:
                        registro['Resumo'] = f"Apuração: {nome_meta_pai}"
                        registro['Apurado no período'] = resumo_original # Mantém o marcador na coluna específica
                    else:
                        registro['Resumo'] = resumo_original
                    
                    registro['Meta_apuração'] = f"[{current_chave}] {registro['Resumo']}"

                # 3. VARREDURA UNIVERSAL DE TABELAS (tableBorder e grid)
                # O Valor da Meta (ex: 100) costuma estar em tabelas de classe 'grid'
                current_element = tabela_inicio
                while current_element:
                    if current_element.name == 'table':
                        for linha in current_element.find_all('tr'):
                            # Captura tanto células normais (td) quanto cabeçalhos (th)
                            celulas = linha.find_all(['td', 'th'])
                            
                            for c_idx, celula in enumerate(celulas):
                                b_tag = celula.find('b')
                                # Se houver um rótulo em negrito e uma célula de valor à direita
                                if b_tag and (c_idx + 1) < len(celulas):
                                    rotulo = re.sub(r'\s+', ' ', b_tag.get_text(strip=True).rstrip(':')).strip()
                                    valor_td = celulas[c_idx + 1]
                                    
                                    # Tratamento de datas/tempo via tag <time>
                                    time_tag = valor_td.find('time')
                                    if time_tag and time_tag.get('datetime'):
                                        valor = time_tag['datetime']
                                    elif rotulo.lower() == 'informação complementar':
                                        valor = valor_td.decode_contents().strip()
                                    else:
                                        valor = valor_td.get_text(separator=' ', strip=True)
                                    
                                    if valor and valor.lower() != "desconhecido":
                                        registro[rotulo] = valor
                    
                    # Navega para o próximo elemento, parando antes do próximo ticket (hr class="fullcontent")
                    proximo = current_element.find_next_sibling()
                    if not proximo or (proximo.name == 'hr' and 'fullcontent' in proximo.get('class', [])):
                        break
                    current_element = proximo
                
                # Adiciona o ticket ao dataset global se tiver uma chave válida
                if registro.get('Chave'):
                    self.dados_extraidos.append(registro)
            
            elapsed = time.time() - start_time
            print(f"✅ Extração finalizada em {elapsed:.2f}s.")
            
        except Exception as e:
            print(f"❌ Erro crítico na extração: {e}")
            traceback.print_exc()
        finally:
            self.driver.close()
            self.driver.switch_to.window(self.janela_principal)
            print("↩️  Retornando à aba de controle do Jira.")
            return num_tickets
        

    def montar_dicionario_hierarquico(self):
        """Gera estrutura hierárquica Pai → Filhos usando META_ID"""
        print("\n🧱 Montando dicionário hierárquico...")
        mapa_tickets = {
            registro.get('Chave'): {'Dados': registro, 'Filhos': []}
            for registro in self.dados_extraidos if 'Chave' in registro
        }
        dicionario_final = {}
        for chave, item in mapa_tickets.items():
            registro = item['Dados']
            chave_pai = registro.get('META_ID')
            if chave_pai and chave_pai in mapa_tickets:
                mapa_tickets[chave_pai]['Filhos'].append(item)
            else:
                dicionario_final[chave] = item
        return dicionario_final
    
    def salvar_excel(self, nome_arquivo):
        """Salva dados extraídos em Excel"""
        if not self.dados_extraidos:
            print("\n❌ Nenhum dado para salvar!")
            return
        
        df = pd.DataFrame(self.dados_extraidos)
        colunas_prioritarias = ['META_ID', 'Chave', 'Resumo']
        if 'Nº_Meta' in df.columns: colunas_prioritarias.append('Nº_Meta')
        if 'Meta_apuração' in df.columns: colunas_prioritarias.append('Meta_apuração')
        
        outras_colunas = sorted([col for col in df.columns if col not in colunas_prioritarias])
        colunas_finais = colunas_prioritarias + outras_colunas
        colunas_existentes = [c for c in colunas_finais if c in df.columns]
        df = df[colunas_existentes]
        
        caminho = Path(self.config.PASTA_SAIDA) / nome_arquivo
        df.to_excel(caminho, index=False)
        print(f"\n💾 Excel salvo: {caminho}")
    
    def salvar_json(self, dicionario, nome_arquivo):
        """Salva dicionário hierárquico em JSON"""
        caminho = Path(self.config.PASTA_SAIDA) / nome_arquivo
        with open(caminho, 'w', encoding='utf-8') as f:
            json.dump(dicionario, f, indent=4, ensure_ascii=False)
        print(f"💾 JSON salvo: {caminho}")
    
    def extrair_simples(self):
        """Modo 1: Extração simples do estado atual"""
        print("\n" + "="*60)
        print("🚀 MODO: EXTRAÇÃO JIRA SIMPLES")
        print("="*60 + "\n")
        self.criar_pasta_saida()
        self.iniciar_navegador()
        try:
            self.login_manual_e_aguardar()
            self.exportar_detalhes_impressao()
            self.processar_aba_exportacao()
            self.salvar_excel(self.config.ARQUIVO_JIRA_SIMPLES)
        except Exception as e:
            print(f"\n❌ Erro durante execução: {e}")
            traceback.print_exc()
        finally:
            input("\n⏸️  Pressione ENTER para fechar o navegador...")
            self.fechar()
    
    def extrair_por_anos(self):
        """Modo 2: Extração filtrada por múltiplos anos"""
        print("\n" + "="*60)
        print("🚀 MODO: EXTRAÇÃO JIRA POR ANOS")
        print("="*60 + "\n")
        self.criar_pasta_saida()
        self.iniciar_navegador()
        try:
            self.login_manual_e_aguardar()
            for ano in self.config.ANOS_EXTRACAO:
                print(f"\n" + "="*60)
                print(f"📅 EXTRAINDO ANO: {ano}")
                print("="*60)
                self.aplicar_filtro_por_ano("Ano da Meta", ano)
                self.exportar_detalhes_impressao()
                self.processar_aba_exportacao()
                jql_encoded_base = quote_plus(self.config.JQL_BASE)
                self.navegar_para_jql(jql_encoded_base)
            
            self.salvar_excel(self.config.ARQUIVO_JIRA_ANUAL)
            dicionario = self.montar_dicionario_hierarquico()
            self.salvar_json(dicionario, self.config.ARQUIVO_JIRA_JSON)
        except Exception as e:
            print(f"\n❌ Erro durante execução: {e}")
            traceback.print_exc()
        finally:
            input("\n⏸️  Pressione ENTER para fechar o navegador...")
            self.fechar()


# ============================================
# EXTRATOR CNJ (ATUALIZADO DO extracao_cnj.py)
# ============================================

class ExtratorCNJ(ExtratorBase):
    """Extrator especializado para o painel do CNJ (Lógica atualizada do extracao_cnj.py)"""
    
    def acessar_painel(self):
        """Acessa o painel de metas do CNJ"""
        print(f"🌐 Acessando {self.config.URL_CNJ}...")
        self.driver.get(self.config.URL_CNJ)
        time.sleep(10)
    
    def entrar_no_iframe(self):
        """Entra no contexto do iframe PowerBI"""
        try:
            iframe = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "iframe")))
            self.driver.switch_to.frame(iframe)
            print("✅ Entrou no contexto do Iframe PowerBI")
        except:
            print("⚠️ Iframe não encontrado (talvez já esteja nele).")

    def clicar_elemento_por_texto(self, texto_parcial):
        print(f"Procurando elemento com texto: '{texto_parcial}'...")
        try:
            xpath = f"//*[contains(text(), '{texto_parcial}')]"
            elementos = self.wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
            elemento_alvo = elementos[-1] 
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", elemento_alvo)
            time.sleep(0.5)
            self.driver.execute_script("arguments[0].dispatchEvent(new MouseEvent('click', {view: window, bubbles:true, cancelable: true}))", elemento_alvo)
            
            print(f"✅ Clique em '{texto_parcial}' realizado com sucesso!")
            time.sleep(4)
            return True
        except Exception as e:
            print(f"❌ Não foi possível clicar em '{texto_parcial}'.")
            return False

    def clicar_botao_laranja_estadual(self, indice_alvo=0):
        print(f"Procurando botões laranjas (Alvo: índice {indice_alvo})...")
        try:
            xpath_cor = "//*[local-name()='path' and contains(@fill, 'e1874d')]"
            elementos = self.wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath_cor)))
            
            qtd = len(elementos)
            print(f"🔎 Encontrados {qtd} elementos laranjas.")

            if qtd > indice_alvo:
                botao_alvo = elementos[indice_alvo]
                self.driver.execute_script("arguments[0].scrollIntoView(true);", botao_alvo)
                self.driver.execute_script("arguments[0].dispatchEvent(new MouseEvent('click', {view: window, bubbles:true, cancelable: true}))", botao_alvo)
                print(f"✅ Clique no botão Laranja (índice {indice_alvo}) realizado!")
                time.sleep(4)
                return True
            else:
                print(f"⚠️ ERRO: Índice {indice_alvo} fora do alcance (só encontrei {qtd} elementos).")
                return False
        except Exception as e:
            print(f"❌ Falha ao processar botões laranjas: {e}")
            return False

    def aplicar_filtro_powerbi(self, nome_interno_filtro, valor_desejado):
        print(f"--- Filtrando '{nome_interno_filtro}' para '{valor_desejado}' ---")
        try:
            dropdown_xpath = f"//div[@class='slicer-dropdown-menu' and @aria-label='{nome_interno_filtro}']"
            # Espera o dropdown estar presente
            dropdown = self.wait.until(EC.presence_of_element_located((By.XPATH, dropdown_xpath)))
            
            # Scroll to center and wait for clickability
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
            time.sleep(1)
            
            # Clica no dropdown para abrir
            self.driver.execute_script("arguments[0].dispatchEvent(new MouseEvent('click', {view: window, bubbles:true, cancelable: true}))", dropdown)
            time.sleep(1.5)

            # Seleciona a opção
            opcao_xpath = f"//div[@class='slicerItemContainer']//span[@title='{valor_desejado}' or text()='{valor_desejado}']"
            opcao = self.wait.until(EC.element_to_be_clickable((By.XPATH, opcao_xpath)))
            self.driver.execute_script("arguments[0].click();", opcao)
            print(f"✅ Opção '{valor_desejado}' selecionada!")
            
            # Fecha o dropdown
            self.driver.execute_script("arguments[0].dispatchEvent(new MouseEvent('click', {view: window, bubbles:true, cancelable: true}))", dropdown)
            time.sleep(2)
            
        except Exception as e:
            print(f"❌ Erro ao filtrar {nome_interno_filtro}: {e}")

    def _identificar_titulo_descricao(self, numero_meta_esperada):
        meta_titulo = f"Meta {numero_meta_esperada}"
        meta_subtitulo = "N/D"
        meta_descricao = "Descrição não encontrada"
        
        try:
            textboxes = self.driver.find_elements(By.CSS_SELECTOR, "div.textbox")
            
            for box in textboxes:
                texto_completo = box.text
                
                if f"Meta {numero_meta_esperada}" in texto_completo or f"Meta{numero_meta_esperada}" in texto_completo:
                    
                    # 1. Extração do Título
                    try:
                        xpath_titulo = f".//span[contains(text(), 'Meta {numero_meta_esperada}')]"
                        elem_titulo = box.find_element(By.XPATH, xpath_titulo)
                        meta_titulo = elem_titulo.text.strip()
                    except:
                        meta_titulo = texto_completo.split('\n')[0].strip()

                    if str(numero_meta_esperada) == "9":
                        meta_titulo = meta_titulo.replace(" de 2025", "")
                        
                    # 2. Extração do Subtítulo (Apenas Metas 2, 6, 7, 8)
                    if str(numero_meta_esperada) in ["2", "6", "7", "8"]:
                        try:
                            # ESTRATÉGIA A: Busca por TEXTO (Mais seguro contra mudança de cores)
                            xpath_sub_texto = ".//span[contains(text(), 'Identificar e julgar')]"
                            elem_sub = box.find_element(By.XPATH, xpath_sub_texto)
                            meta_subtitulo = elem_sub.text.replace(":", "").strip()
                        except:
                            try:
                                # ESTRATÉGIA B: Busca por COR (Fallback)
                                xpath_sub_cor = ".//span[contains(@style, 'rgb(204, 204, 204)') or contains(@style, 'rgb(179, 179, 179)')]"
                                elem_sub = box.find_element(By.XPATH, xpath_sub_cor)
                                meta_subtitulo = elem_sub.text.replace(":", "").strip()
                            except:
                                pass # Se falhar tudo, mantém N/D

                    # 3. Extração da Descrição (Prioridade: Justiça Estadual)
                    try:
                        xpath_estadual = ".//li[contains(., 'Justiça Estadual')]"
                        elem_estadual = box.find_element(By.XPATH, xpath_estadual)
                        texto_bruto = elem_estadual.text
                        meta_descricao = texto_bruto.replace("Justiça Estadual:", "").replace("Justiça Estadual", "").strip()
                        return meta_titulo, meta_subtitulo, meta_descricao
                        
                    except:
                        # Fallback para metas sem lista de justiça estadual
                        linhas = [l for l in texto_completo.split('\n') if len(l) > 10 and "Meta" not in l and l != meta_subtitulo]
                        if linhas:
                            meta_descricao = linhas[0]
                    
                    return meta_titulo, meta_subtitulo, meta_descricao
                    
        except Exception as e:
            print(f"⚠️ Erro ao ler textos da Meta {numero_meta_esperada}: {e}")

        return meta_titulo, meta_subtitulo, meta_descricao

    def adicionar_linha(self, titulo, subtitulo, desc, cat, val):
        print(f"   > Capturado: {cat} -> {val} (Sub: {subtitulo})")
        self.dados_extraidos.append({
            "Meta": titulo,
            "Subtítulo": subtitulo,
            "Descrição": desc,
            "Categoria": cat,
            "Resultado": val,
            "Data": datetime.now().strftime("%Y-%m-%d %H:%M")
        })

    # --- FUNÇÕES DE EXTRAÇÃO ESPECÍFICAS ---

    def extrair_dados_grafico(self, numero_meta_esperada=None):
        """Extrai dados de gráficos (Metas 1 e 2)"""
        print(f"\n--- Iniciando Extração Gráfico (Alvo: Meta {numero_meta_esperada if numero_meta_esperada else 'Qualquer'}) ---")
        
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao(numero_meta_esperada)
        print(f"📌 Meta Identificada: {meta_titulo}")

        print("🔍 Localizando gráfico de barras...")
        container = None
        seletores = ["div[aria-label*='por Ramo']", "div[aria-label*='por Tribunal']", "div[aria-label*='Meta']"]
        
        for seletor in seletores:
            try:
                elems = self.driver.find_elements(By.CSS_SELECTOR, seletor)
                for el in elems:
                    if el.find_elements(By.CSS_SELECTOR, "g.axis"):
                        container = el
                        break
                if container: break
            except: continue
            
        if container:
            try:
                categorias = container.find_elements(By.CSS_SELECTOR, "g.axis.y g.tick text")
                valores = container.find_elements(By.CSS_SELECTOR, "g.label-container tspan.label-tspan")
                
                l_cats = [c.get_attribute('textContent').strip() for c in categorias if c.get_attribute('textContent').strip()]
                l_vals = [v.get_attribute('textContent').strip() for v in valores if v.get_attribute('textContent').strip()]
                
                # Remover duplicação de categorias (ex: "1º Grau1º Grau" -> "1º Grau")
                l_cats_limpo = []
                for cat in l_cats:
                    if len(cat) > 0 and len(cat) % 2 == 0:
                        metade = len(cat) // 2
                        primeira_metade = cat[:metade]
                        segunda_metade = cat[metade:]
                        if primeira_metade == segunda_metade:
                            l_cats_limpo.append(primeira_metade)
                        else:
                            l_cats_limpo.append(cat)
                    else:
                        l_cats_limpo.append(cat)

                if l_cats_limpo and l_vals:
                    limite = min(len(l_cats_limpo), len(l_vals))
                    for i in range(limite):
                        self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, l_cats_limpo[i], l_vals[i])
                    print(f"✅ Extraídos {limite} registros.")
                else:
                    print("⚠️ Container achado, mas dados vazios.")
            except:
                print("⚠️ Erro na leitura interna do gráfico.")
        else:
            print("⚠️ Nenhum gráfico encontrado.")

    def extrair_kpi_meta_1_total(self):
        print(f"\n--- Iniciando Extração KPI Total (Meta 1) ---")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao(numero_meta_esperada="1")
        print(f"📌 Meta Identificada: {meta_titulo}")
        
        print(f"🔍 Buscando cartão 'Julgar mais processos que os distribuídos'...")
        try:
            xpath_card = "//div[@title='Julgar mais processos que os distribuídos']/ancestor::div[contains(@class, 'visualWrapper')]"
            card = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_card)))
            valor = card.find_element(By.CSS_SELECTOR, "text.value tspan").text.strip()
            print(f"💎 Valor encontrado: {valor}")
            self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, "Total", valor)
        except Exception as e:
            print(f"❌ Erro ao extrair KPI Total da Meta 1: {e}")
    
    def extrair_kpis_meta_2(self):
        print(f"\n--- Iniciando Extração KPIs Meta 2 ---")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao(numero_meta_esperada="2")
        print(f"📌 Meta Identificada: {meta_titulo}")
        
        instancias_titles = {
            "1º Grau": "1º Grau",
            "2º Grau": "2º Grau",
            "Juizados e Turmas": "Juizados e Turmas",
            "Processos mais Antigos": "Processos mais Antigos"
        }
        
        for titulo_card, instancia in instancias_titles.items():
            print(f"\n🔍 Buscando card 'Cumprimento' da instância '{titulo_card}'...")
            try:
                xpath_wrapper = f"//div[@title='{titulo_card}']/ancestor::div[contains(@class, 'visualWrapper')]"
                wrapper = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_wrapper)))
                try:
                    cumprimento_xpath = ".//h4[contains(text(), 'Cumprimento')]/following-sibling::p"
                    valor_elem = wrapper.find_element(By.XPATH, cumprimento_xpath)
                    valor = valor_elem.text.strip()
                    print(f"   💎 {instancia} - Cumprimento: {valor}")
                    self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, instancia, valor)
                except Exception as e:
                    print(f"   ⚠️ Erro ao extrair Cumprimento: {e}")
            except Exception as e:
                print(f"❌ Erro ao processar instância '{titulo_card}': {e}")
    
    def extrair_kpi_cumprimento(self, numero_meta, kpi_title):
        print(f"\n--- Iniciando Extração KPI (Alvo: Meta {numero_meta}) ---")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao(numero_meta_esperada=str(numero_meta))
        print(f"📌 Meta Identificada: {meta_titulo}")
        
        print(f"🔍 Buscando cartão '{kpi_title}'...")
        try:
            xpath_card = f"//div[@title='{kpi_title}']/ancestor::div[contains(@class, 'visualWrapper')]"
            card = self.driver.find_element(By.XPATH, xpath_card)
            valor = card.find_element(By.CSS_SELECTOR, "text.value").text.strip()
            print(f"💎 Valor encontrado: {valor}")
            self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, "Total", valor)
        except Exception as e:
            print(f"❌ Erro ao extrair KPI da Meta {numero_meta} (Título: {kpi_title}): {e}")

    def extrair_kpis_meta_4(self):
        print(f"\n--- Iniciando Extração KPI (Alvo: Meta 4) ---")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao(numero_meta_esperada="4")
        print(f"📌 Meta Identificada: {meta_titulo}")
        
        def extrair_valor_do_cartao(container_title, label_nome, campo_nome_saida):
            try:
                xpath_container = f"//div[@title='{container_title}']/ancestor::transform[1]"
                container = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_container)))
                xpath_value = f".//h4[contains(text(), '{label_nome}')]/following-sibling::p[contains(@class, 'bottom')]"
                valor = container.find_element(By.XPATH, xpath_value).text.strip()
                if valor:
                    self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, campo_nome_saida, valor)
                    return True
            except: return False

        extrair_valor_do_cartao("Meta 4", "Cumprimento", "Crimes contra a administração pública")
        extrair_valor_do_cartao("Meta 4 Improb. Administrativa", "Cumprimento", "Improbidade administrativa")

    def extrair_kpi_meta_6(self):
        print(f"\n--- Iniciando Extração KPI (Alvo: Meta 6) ---")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao(numero_meta_esperada="6")
        print(f"📌 Meta Identificada: {meta_titulo}")
        
        print("🔍 Buscando cartão pelo background SVG...")
        try:
            xpath_bg = "//*[local-name()='path' and @data-sub-selection-display-name='Card_Background_Color']"
            backgrounds = self.wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath_bg)))
            valor_encontrado = None
            for bg in backgrounds:
                try:
                    container = bg.find_element(By.XPATH, "./ancestor::visual-modern[1]")
                    texto_container = container.text
                    if "Cumprimento" in texto_container:
                        try: valor_encontrado = container.find_element(By.CSS_SELECTOR, "p.content").text.strip()
                        except: valor_encontrado = container.find_element(By.CSS_SELECTOR, "text.value").text.strip()
                        if valor_encontrado: break
                except: continue
            
            if valor_encontrado:
                print(f"💎 Valor encontrado: {valor_encontrado}")
                self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, "Total", valor_encontrado)
            else:
                print("❌ Erro: Cartão 'Cumprimento' não encontrado ou valor vazio.")
        except Exception as e:
            print(f"❌ Erro grave ao extrair Meta 6: {e}")

    def extrair_kpis_meta_7(self):
        print(f"\n--- Iniciando Extração KPI (Alvo: Meta 7) ---")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao(numero_meta_esperada="7")
        print(f"📌 Meta Identificada: {meta_titulo}")
        
        def extrair_valor_do_cartao(container_title, label_nome, campo_nome_saida):
            try:
                xpath_container = f"//div[@title='{container_title}']/ancestor::transform[1]"
                container = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_container)))
                xpath_value = f".//h4[contains(text(), '{label_nome}')]/following-sibling::p[contains(@class, 'bottom')]"
                valor = container.find_element(By.XPATH, xpath_value).text.strip()
                if valor:
                    self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, campo_nome_saida, valor)
                    return True
            except: return False

        extrair_valor_do_cartao("Meta 7 Indígenas", "Cumprimento", "Total Indígenas")
        extrair_valor_do_cartao("Meta 7 Quilombola", "Cumprimento", "Total Quilombola")

    def extrair_kpis_meta_8(self):
        print(f"\n--- Iniciando Extração KPI (Alvo: Meta 8) ---")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao(numero_meta_esperada="8")
        print(f"📌 Meta Identificada: {meta_titulo}")
        
        def extrair_valor_do_cartao(container_title, label_nome, campo_nome_saida):
            try:
                xpath_container = f"//div[@title='{container_title}']/ancestor::transform[1]"
                container = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_container)))
                xpath_value = f".//h4[contains(text(), '{label_nome}')]/following-sibling::p[contains(@class, 'bottom')]"
                valor = container.find_element(By.XPATH, xpath_value).text.strip()
                if valor:
                    self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, campo_nome_saida, valor)
                    return True
            except: return False

        extrair_valor_do_cartao("Violência Doméstica", "Cumprimento", "Total Violência Doméstica")
        extrair_valor_do_cartao("Feminicídio", "Cumprimento", "Total Feminicídio")

    def extrair_kpis_meta_9(self):
        print(f"\n--- Iniciando Extração KPI (Alvo: Meta 9) ---")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao(numero_meta_esperada="9")
        print(f"📌 Meta Identificada: {meta_titulo}")
        
        def extrair_valor_do_cartao(label_busca, campo_nome_saida):
            try:
                xpath_valor = f"//*[contains(text(), '{label_busca}')]/following::*[string-length(text()) > 0 and contains(@class, 'value') or contains(@class, 'content')][1]"
                elementos = self.driver.find_elements(By.XPATH, xpath_valor)
                if elementos:
                    valor = elementos[0].text.strip()
                    print(f"💎 Valor encontrado para '{campo_nome_saida}': {valor}")
                    self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, campo_nome_saida, valor)
                    return True
                else:
                    print(f"⚠️ Valor não encontrado para o rótulo '{label_busca}'")
                    return False
            except Exception as e: 
                print(f"❌ Erro ao ler cartão Meta 9: {e}")
                return False

        if not extrair_valor_do_cartao("Cumprimento", "Total Meta 9"):
            try:
                xpath_percent = "//*[contains(text(), '%') and (@class='value' or contains(@class, 'label'))]"
                elem = self.driver.find_element(By.XPATH, xpath_percent)
                self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, "Total Estimado", elem.text)
            except:
                print("❌ Não foi possível extrair dados da Meta 9 pelos métodos padrão.")

    def extrair_kpis_meta_10(self):
        print(f"\n--- Iniciando Extração KPI (Alvo: Meta 10) ---")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao(numero_meta_esperada="10")
        print(f"📌 Meta Identificada: {meta_titulo}")
        
        def extrair_valor_do_cartao(container_title, label_nome, campo_nome_saida):
            try:
                xpath_container = f"//div[@title='{container_title}']/ancestor::transform[1]"
                container = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_container)))
                xpath_value = f".//h4[contains(text(), '{label_nome}')]/following-sibling::p[contains(@class, 'bottom')]"
                valor = container.find_element(By.XPATH, xpath_value).text.strip()
                if valor:
                    self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, campo_nome_saida, valor)
                    return True
            except: return False

        extrair_valor_do_cartao("1º Grau", "Cumprimento", "1º Grau")
        extrair_valor_do_cartao("2º Grau", "Cumprimento", "2º Grau")

    def salvar_excel(self):
        """Salva dados extraídos em Excel"""
        if self.dados_extraidos:
            df = pd.DataFrame(self.dados_extraidos)
            caminho = Path(self.config.PASTA_SAIDA) / self.config.ARQUIVO_CNJ
            df.to_excel(caminho, index=False)
            print(f"\n💾 Arquivo salvo: {caminho}")
            print(f"📊 Total de registros: {len(df)}")
        else:
            print("\n⚠️  Nenhum dado para salvar")
    
    def extrair_completo(self):
        """Modo 3: Extração completa das metas do CNJ"""
        print("\n" + "="*60)
        print("🚀 MODO: EXTRAÇÃO CNJ")
        print("="*60 + "\n")

        self.criar_pasta_saida()
        self.iniciar_navegador()

        try:
            self.acessar_painel()
            self.entrar_no_iframe()

            # META 1
            print("\n=== 🏁 INICIANDO META 1 ===")
            self.aplicar_filtro_powerbi("ramo_justica", "Justiça Estadual")
            time.sleep(2)
            self.aplicar_filtro_powerbi("sigla_tribunal", "TJMG")
            self.extrair_dados_grafico("1")
            self.extrair_kpi_meta_1_total()

            # META 2
            print("\n=== 🏁 INICIANDO META 2 ===")
            if self.clicar_elemento_por_texto("Meta 2"):
                self.clicar_botao_laranja_estadual(indice_alvo=1)
                self.aplicar_filtro_powerbi("sigla_tribunal", "TJMG")
                self.extrair_kpis_meta_2()

            # META 3
            print("\n=== 🏁 INICIANDO META 3 ===")
            if self.clicar_elemento_por_texto("Meta 3"):
                self.aplicar_filtro_powerbi("sigla_tribunal", "TJMG")
                self.extrair_kpi_cumprimento(numero_meta="3", kpi_title="Percentual de Cumprimento")

            # META 4
            print("\n=== 🏁 INICIANDO META 4 ===")
            if self.clicar_elemento_por_texto("Meta 4"):
                self.clicar_botao_laranja_estadual(indice_alvo=1)
                self.aplicar_filtro_powerbi("sigla_tribunal", "TJMG")
                self.extrair_kpis_meta_4()

            # META 5
            print("\n=== 🏁 INICIANDO META 5 ===")
            if self.clicar_elemento_por_texto("Meta 5"):
                self.clicar_botao_laranja_estadual(indice_alvo=2)
                self.aplicar_filtro_powerbi("Tribunal", "TJMG")
                self.extrair_kpi_cumprimento(numero_meta="5", kpi_title="Cumprimento Meta 5")

            # META 6
            print("\n=== 🏁 INICIANDO META 6 ===")
            if self.clicar_elemento_por_texto("Meta 6"):
                self.clicar_botao_laranja_estadual(indice_alvo=1)
                try: self.aplicar_filtro_powerbi("sigla_tribunal", "TJMG")
                except: self.aplicar_filtro_powerbi("Tribunal", "TJMG")
                self.extrair_kpi_meta_6()

            # META 7
            print("\n=== 🏁 INICIANDO META 7 ===")
            if self.clicar_elemento_por_texto("Meta 7"):
                self.clicar_botao_laranja_estadual(indice_alvo=1)
                try: self.aplicar_filtro_powerbi("sigla_tribunal", "TJMG")
                except: self.aplicar_filtro_powerbi("Tribunal", "TJMG")
                self.extrair_kpis_meta_7()

            # META 8
            print("\n=== 🏁 INICIANDO META 8 ===")
            if self.clicar_elemento_por_texto("Meta 8"):
                self.clicar_botao_laranja_estadual(indice_alvo=1)
                try: self.aplicar_filtro_powerbi("sigla_tribunal", "TJMG")
                except: self.aplicar_filtro_powerbi("Tribunal", "TJMG")
                self.extrair_kpis_meta_8()

            # META 9
            print("\n=== 🏁 INICIANDO META 9 ===")
            if self.clicar_elemento_por_texto("Meta 9"):
                try: self.aplicar_filtro_powerbi("sigla_tribunal", "TJMG")
                except: self.aplicar_filtro_powerbi("Tribunal", "TJMG")
                self.extrair_kpis_meta_9()

            # META 10
            print("\n=== 🏁 INICIANDO META 10 ===")
            if self.clicar_elemento_por_texto("Meta 10"):
                self.clicar_botao_laranja_estadual(indice_alvo=1)
                try: self.aplicar_filtro_powerbi("sigla_tribunal", "TJMG")
                except: self.aplicar_filtro_powerbi("Tribunal", "TJMG")
                self.extrair_kpis_meta_10()

            print("\n⏸️ Processo finalizado.")
            self.salvar_excel()
            input("\nPressione ENTER para fechar o navegador e encerrar o robô...")
            self.fechar()

        except Exception as e:
            print(f"\n❌ Erro durante execução: {e}")
            traceback.print_exc()
            input("\nPressione ENTER para fechar o navegador e encerrar o robô...")
            self.fechar()


# ============================================
# INTERFACE DE LINHA DE COMANDO
# ============================================

def criar_parser():
    """Cria parser de argumentos de linha de comando"""
    parser = argparse.ArgumentParser(
        description='🎯 Extrator Unificado de Dados - TJMG',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos de uso:

  # Menu interativo (sem argumentos)
  python extrator_unificado.py

  # Extração automática de um ano específico (Jira + CNJ)
  python extrator_unificado.py 2024

  # Extração de múltiplos anos
  python extrator_unificado.py 2023 2024 2025

  # Apenas Jira (sem CNJ)
  python extrator_unificado.py 2024 --sem-cnj

  # Apenas CNJ (sem Jira)
  python extrator_unificado.py --apenas-cnj

  # Alterar navegador
  python extrator_unificado.py 2024 --navegador chrome
        """
    )
    
    parser.add_argument(
        'anos',
        nargs='*',
        help='Anos para extração (ex: 2024 ou 2023 2024 2025). Se não informado, abre menu interativo.'
    )
    
    parser.add_argument(
        '--sem-cnj',
        action='store_true',
        help='Extrai apenas Jira (não extrai CNJ)'
    )
    
    parser.add_argument(
        '--apenas-cnj',
        action='store_true',
        help='Extrai apenas CNJ (não extrai Jira)'
    )
    
    parser.add_argument(
        '--navegador', '-n',
        choices=['edge', 'chrome'],
        help='Navegador a ser usado (padrão: edge)'
    )
    
    parser.add_argument(
        '--output', '-o',
        help='Pasta de saída para arquivos (padrão: exports)'
    )
    
    return parser


def aplicar_argumentos_config(args, config):
    """Aplica argumentos de linha de comando à configuração"""
    if args.anos:
        config.ANOS_EXTRACAO = args.anos
        print(f"📅 Anos configurados: {', '.join(args.anos)}")
    
    if args.output:
        config.PASTA_SAIDA = args.output
        print(f"📂 Pasta de saída: {args.output}")
    
    if args.navegador:
        config.NAVEGADOR = args.navegador
        print(f"🌐 Navegador: {args.navegador}")


def menu_principal():
    """Exibe menu de opções"""
    print("\n" + "="*60)
    print("🎯 EXTRATOR UNIFICADO DE DADOS - TJMG")
    print("="*60)
    print("\nEscolha o modo de operação:")
    print()
    print("1️⃣  Jira Simples     - Extração única do estado atual")
    print("2️⃣  Jira Anual       - Extração filtrada por anos")
    print("3️⃣  CNJ              - Extração do painel de metas nacionais")
    print("4️⃣  Completo         - Executa Jira Anual + CNJ")
    print("0️⃣  Sair")
    print()
    
    return input("Digite sua escolha: ").strip()


def executar_modo(modo, config):
    """Executa modo específico de extração"""
    if modo == "jira-simples" or modo == "1":
        extrator = ExtratorJira(config)
        extrator.extrair_simples()
    
    elif modo == "jira-anual" or modo == "2":
        extrator = ExtratorJira(config)
        extrator.extrair_por_anos()
    
    elif modo == "cnj" or modo == "3":
        extrator = ExtratorCNJ(config)
        extrator.extrair_completo()
    
    elif modo == "completo" or modo == "4":
        print("\n🔄 Executando extração completa...")
        
        # Jira
        print("\n📌 ETAPA 1/2: Extração Jira")
        extrator_jira = ExtratorJira(config)
        extrator_jira.extrair_por_anos()
        
        # CNJ
        print("\n📌 ETAPA 2/2: Extração CNJ")
        extrator_cnj = ExtratorCNJ(config)
        extrator_cnj.extrair_completo()
        
        print("\n" + "="*60)
        print("✅ EXTRAÇÃO COMPLETA FINALIZADA!")
        print("="*60)


def executar_extracao_automatica(config, sem_cnj=False, apenas_cnj=False):
    """Executa extração automática baseada nos anos configurados"""
    print("\n" + "="*60)
    print("🚀 EXTRAÇÃO AUTOMÁTICA")
    print("="*60)
    print(f"📅 Anos: {', '.join(config.ANOS_EXTRACAO)}")
    print(f"📂 Saída: {config.PASTA_SAIDA}")
    print("="*60 + "\n")
    
    if apenas_cnj:
        # Apenas CNJ
        print("📌 Modo: CNJ apenas")
        extrator_cnj = ExtratorCNJ(config)
        extrator_cnj.extrair_completo()
    
    elif sem_cnj:
        # Apenas Jira
        print("📌 Modo: Jira apenas")
        extrator_jira = ExtratorJira(config)
        extrator_jira.extrair_por_anos()
    
    else:
        # Completo (Jira + CNJ)
        print("📌 Modo: Completo (Jira + CNJ)")
        
        # Jira
        print("\n📌 ETAPA 1/2: Extração Jira")
        extrator_jira = ExtratorJira(config)
        extrator_jira.extrair_por_anos()
        
        # CNJ
        print("\n📌 ETAPA 2/2: Extração CNJ")
        extrator_cnj = ExtratorCNJ(config)
        extrator_cnj.extrair_completo()
    
    print("\n" + "="*60)
    print("✅ EXTRAÇÃO FINALIZADA COM SUCESSO!")
    print("="*60)


def executar_menu_interativo():
    """Executa modo interativo com menu"""
    config = Config()
    
    while True:
        escolha = menu_principal()
        
        if escolha == "0":
            print("\n👋 Até logo!")
            break
        elif escolha in ["1", "2", "3", "4"]:
            executar_modo(escolha, config)
        else:
            print("\n❌ Opção inválida! Tente novamente.")


def executar():
    """Função principal de execução"""
    parser = criar_parser()
    args = parser.parse_args()
    
    # Se nenhum argumento foi passado, usa menu interativo
    if not args.anos and not args.apenas_cnj:
        executar_menu_interativo()
    else:
        # Modo automático com anos
        config = Config()
        aplicar_argumentos_config(args, config)
        
        # Valida conflito de flags
        if args.sem_cnj and args.apenas_cnj:
            print("❌ Erro: Não é possível usar --sem-cnj e --apenas-cnj simultaneamente")
            sys.exit(1)
        
        executar_extracao_automatica(config, sem_cnj=args.sem_cnj, apenas_cnj=args.apenas_cnj)


# ============================================
# EXECUÇÃO
# ============================================

if __name__ == "__main__":
    try:
        executar()
    except KeyboardInterrupt:
        print("\n\n⚠️  Execução cancelada pelo usuário")
    except Exception as e:
        print(f"\n❌ ERRO INESPERADO: {e}")
        traceback.print_exc()