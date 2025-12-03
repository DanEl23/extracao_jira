"""
EXTRATOR UNIFICADO DE DADOS - TJMG
Versão: 2.0
Integra extração de Jira e CNJ em um único módulo

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
    JQL_BASE = "project = ASPLAGMETA OR4DER BY created DESC"
    ANOS_EXTRACAO = ["2024", "2025", "2026"]
    
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
        
        try:
            # 1. Clicar no botão 'Mais filtros'
            more_filters_button_selector = "button[data-testid='jql-builder-basic.ui.jql-editor.add-filter.more-button']"
            more_filters_button = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, more_filters_button_selector))
            )
            more_filters_button.click()
            print("   1. Clicou em 'Mais filtros'")
            
            # 2. Digitar o nome do campo
            search_input_xpath = "//input[@aria-label='Pesquisar mais filtros']"
            search_input = self.wait.until(
                EC.presence_of_element_located((By.XPATH, search_input_xpath))
            )
            search_input.send_keys(nome_campo)
            time.sleep(1)
            
            # 3. Selecionar o campo
            field_option_xpath = f"//div[@role='option']//div[text()='{nome_campo}']"
            field_option = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, field_option_xpath))
            )
            field_option.click()
            print(f"   2. Selecionou o campo '{nome_campo}'")
            time.sleep(2)
            
            # 4. Digitar valor do ano
            value_input_xpath = f"//input[@aria-label='Pesquisar {nome_campo}']"
            value_input = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, value_input_xpath))
            )
            value_input.send_keys(valor_ano)
            print(f"   3. Digitou '{valor_ano}'")
            time.sleep(1)
            
            # 5. Selecionar opção
            value_option_xpath = f"//div[@role='listbox']//div[text()='{valor_ano}']"
            value_option = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, value_option_xpath))
            )
            value_option.click()
            print(f"   4. Selecionou '{valor_ano}'")
            
            # 6. Aguardar carregamento
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
        """Extrai dados da aba de exportação usando BeautifulSoup"""
        janelas = self.driver.window_handles
        if len(janelas) < 2:
            print("❌ Nova aba de exportação não detectada")
            return 0
        
        janela_exportacao = [w for w in janelas if w != self.janela_principal][0]
        self.driver.switch_to.window(janela_exportacao)
        print("\n🔄 Foco mudado para aba de exportação")
        
        html_content = self.driver.page_source
        start_time = time.time()
        print("   Iniciando extração com BeautifulSoup...")
        
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            blocos_tickets = soup.find_all('table', class_='tableBorder')
            num_tickets = len(blocos_tickets)
            print(f"📊 Encontrados {num_tickets} tickets neste lote")
            
            for idx, tabela_inicio in enumerate(blocos_tickets):
                registro = {}
                
                # 1. META_ID (Parent Key)
                parent_key_tag = tabela_inicio.find('a', id='parent_issue_key')
                registro['META_ID'] = parent_key_tag.get_text(strip=True) if parent_key_tag else ""
                
                # 2. Nº_Meta (Parent Summary)
                parent_summary_tag = tabela_inicio.find('a', id='parent_issue_summary')
                if parent_summary_tag:
                    registro['Nº_Meta'] = parent_summary_tag.get_text(strip=True)
                
                # 3. Chave e Resumo do ticket atual
                h3_element = tabela_inicio.find('h3', class_='formtitle')
                if h3_element:
                    titulo_completo = h3_element.get_text(strip=True)
                    resumo_link = h3_element.find('a')
                    
                    if resumo_link:
                        current_summary = resumo_link.get_text(strip=True)
                        chave_match = re.search(r'\[([A-Z]+-\d+)\]', titulo_completo)
                        current_chave = chave_match.group(1) if chave_match else f"TICKET-{len(self.dados_extraidos) + idx + 1}"
                        
                        registro['Meta_apuração'] = f"[{current_chave}] {current_summary}"
                        registro['Chave'] = current_chave
                        registro['Resumo'] = current_summary
                    else:
                        registro['Chave'] = f"TICKET-{len(self.dados_extraidos) + idx + 1}"
                        registro['Resumo'] = "Resumo não encontrado"
                else:
                    registro['Chave'] = f"TICKET-{len(self.dados_extraidos) + idx + 1}"
                    registro['Resumo'] = "Resumo não encontrado"
                
                # 4. Extração de campos customizados (tabelas dinâmicas)
                tabelas_ticket = []
                current_element = tabela_inicio
                while current_element:
                    if current_element.name == 'table':
                        tabelas_ticket.append(current_element)
                    proximo = current_element.find_next_sibling()
                    if not proximo:
                        break
                    if proximo.name == 'hr' and 'fullcontent' in proximo.get('class', []):
                        break
                    current_element = proximo
                
                for tabela in tabelas_ticket:
                    for linha in tabela.find_all('tr'):
                        colunas = linha.find_all('td')
                        if not colunas:
                            continue
                        
                        i = 0
                        while i < len(colunas):
                            label_cell = colunas[i]
                            b_tag = label_cell.find('b')
                            if not b_tag:
                                i += 1
                                continue
                            
                            rotulo_bruto = b_tag.get_text(separator=' ', strip=True).rstrip(':')
                            rotulo = re.sub(r'\s+', ' ', rotulo_bruto).strip()
                            if not rotulo or i + 1 >= len(colunas):
                                i += 1
                                continue
                            
                            valor_td = colunas[i + 1]
                            valor = valor_td.get_text(separator=' ', strip=True)
                            
                            # Tratamento especial para datas
                            if rotulo.lower() in ['data de apuração', 'data de criação', 'atualizado']:
                                time_tag = valor_td.find('time')
                                if time_tag and time_tag.get('datetime'):
                                    valor = time_tag['datetime']
                            
                            # Tratamento para campos HTML
                            if rotulo.lower() == 'informação complementar' or valor_td.find_all('p'):
                                valor = valor_td.decode_contents().strip()
                            
                            if valor:
                                registro[rotulo] = valor
                            
                            i += 2
                
                self.dados_extraidos.append(registro)
                
                if (idx + 1) % 100 == 0:
                    print(f"   Processados {idx+1} tickets...")
            
            elapsed = time.time() - start_time
            print(f"\n⏱️  Tempo: {elapsed:.2f}s ({num_tickets} tickets)")
            print(f"   Total acumulado: {len(self.dados_extraidos)}")
            
        except Exception as e:
            print(f"❌ Erro na extração: {e}")
            traceback.print_exc()
        finally:
            self.driver.close()
            self.driver.switch_to.window(self.janela_principal)
            self.wait = WebDriverWait(self.driver, self.config.TIMEOUT)
            print("↩️  Retornando à aba principal")
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
        
        print(f"   ✅ Dicionário montado. Metas Raiz: {len(dicionario_final)}")
        return dicionario_final
    
    def salvar_excel(self, nome_arquivo):
        """Salva dados extraídos em Excel"""
        if not self.dados_extraidos:
            print("\n❌ Nenhum dado para salvar!")
            return
        
        df = pd.DataFrame(self.dados_extraidos)
        
        # Ordenação de colunas - mantém compatibilidade com scripts originais
        colunas_prioritarias = ['META_ID', 'Chave', 'Resumo']
        
        # Adiciona Nº_Meta e Meta_apuração se existirem
        if 'Nº_Meta' in df.columns:
            colunas_prioritarias.append('Nº_Meta')
        if 'Meta_apuração' in df.columns:
            colunas_prioritarias.append('Meta_apuração')
        
        # Restante das colunas em ordem alfabética (como nos scripts originais)
        outras_colunas = sorted([col for col in df.columns if col not in colunas_prioritarias])
        colunas_finais = colunas_prioritarias + outras_colunas
        
        # Filtra apenas colunas existentes
        colunas_existentes = [c for c in colunas_finais if c in df.columns]
        df = df[colunas_existentes]
        
        caminho = Path(self.config.PASTA_SAIDA) / nome_arquivo
        df.to_excel(caminho, index=False)
        print(f"\n💾 Excel salvo: {caminho}")
        print(f"📊 Total de registros: {len(df)}")
        print(f"📋 Primeiras colunas: {', '.join(df.columns[:5].tolist())}...")
    
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
            
            print("\n" + "="*60)
            print("INICIANDO EXTRAÇÃO")
            print("="*60)
            
            self.exportar_detalhes_impressao()
            self.processar_aba_exportacao()
            self.salvar_excel(self.config.ARQUIVO_JIRA_SIMPLES)
            
            print("\n✅ EXTRAÇÃO SIMPLES CONCLUÍDA COM SUCESSO!")
            
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
                
                # Resetar filtro para próximo ano
                print("🔁 Resetando filtro...")
                jql_encoded_base = quote_plus(self.config.JQL_BASE)
                self.navegar_para_jql(jql_encoded_base)
            
            # Salvar Excel plano
            self.salvar_excel(self.config.ARQUIVO_JIRA_ANUAL)
            
            # Salvar JSON hierárquico
            dicionario = self.montar_dicionario_hierarquico()
            self.salvar_json(dicionario, self.config.ARQUIVO_JIRA_JSON)
            
            print("\n✅ EXTRAÇÃO POR ANOS CONCLUÍDA COM SUCESSO!")
            
        except Exception as e:
            print(f"\n❌ Erro durante execução: {e}")
            traceback.print_exc()
        finally:
            input("\n⏸️  Pressione ENTER para fechar o navegador...")
            self.fechar()


# ============================================
# EXTRATOR CNJ
# ============================================

class ExtratorCNJ(ExtratorBase):
    """Extrator especializado para o painel do CNJ"""
    
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
            print("✅ Entrou no iframe PowerBI")
        except:
            print("⚠️  Iframe não encontrado (talvez já esteja no contexto)")
    
    def clicar_elemento_por_texto(self, texto_parcial):
        """Clica em elemento que contém o texto especificado"""
        print(f"🔍 Procurando elemento com texto: '{texto_parcial}'...")
        try:
            xpath = f"//*[contains(text(), '{texto_parcial}')]"
            elementos = self.wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
            elemento_alvo = elementos[-1]
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", elemento_alvo)
            time.sleep(0.5)
            self.driver.execute_script(
                "arguments[0].dispatchEvent(new MouseEvent('click', {view: window, bubbles:true, cancelable: true}))",
                elemento_alvo
            )
            
            print(f"   ✅ Clique em '{texto_parcial}' realizado")
            time.sleep(4)
            return True
        except Exception as e:
            print(f"   ❌ Não foi possível clicar em '{texto_parcial}'")
            return False
    
    def clicar_botao_laranja_estadual(self, indice_alvo=0):
        """Clica em botões laranjas (geralmente para selecionar Justiça Estadual)"""
        print(f"🔍 Procurando botões laranjas (índice {indice_alvo})...")
        try:
            xpath_cor = "//*[local-name()='path' and contains(@fill, 'e1874d')]"
            elementos = self.wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath_cor)))
            
            qtd = len(elementos)
            print(f"   Encontrados {qtd} elementos laranjas")
            
            if qtd > indice_alvo:
                botao_alvo = elementos[indice_alvo]
                self.driver.execute_script("arguments[0].scrollIntoView(true);", botao_alvo)
                self.driver.execute_script(
                    "arguments[0].dispatchEvent(new MouseEvent('click', {view: window, bubbles:true, cancelable: true}))",
                    botao_alvo
                )
                print(f"   ✅ Clique no botão laranja (índice {indice_alvo}) realizado")
                time.sleep(4)
                return True
            else:
                print(f"   ⚠️  Índice {indice_alvo} fora do alcance (só encontrei {qtd})")
                return False
        except Exception as e:
            print(f"   ❌ Falha ao processar botões laranjas: {e}")
            return False
    
    def aplicar_filtro_powerbi(self, nome_interno_filtro, valor_desejado):
        """Aplica filtro no PowerBI"""
        print(f"⚙️  Filtrando '{nome_interno_filtro}' = '{valor_desejado}'")
        try:
            dropdown_xpath = f"//div[@class='slicer-dropdown-menu' and @aria-label='{nome_interno_filtro}']"
            dropdown = self.wait.until(EC.presence_of_element_located((By.XPATH, dropdown_xpath)))
            
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
            time.sleep(1)
            
            self.driver.execute_script(
                "arguments[0].dispatchEvent(new MouseEvent('click', {view: window, bubbles:true, cancelable: true}))",
                dropdown
            )
            time.sleep(1.5)
            
            opcao_xpath = f"//div[@class='slicerItemContainer']//span[@title='{valor_desejado}' or text()='{valor_desejado}']"
            opcao = self.wait.until(EC.element_to_be_clickable((By.XPATH, opcao_xpath)))
            self.driver.execute_script("arguments[0].click();", opcao)
            print(f"   ✅ Opção '{valor_desejado}' selecionada")
            
            self.driver.execute_script(
                "arguments[0].dispatchEvent(new MouseEvent('click', {view: window, bubbles:true, cancelable: true}))",
                dropdown
            )
            time.sleep(2)
            
        except Exception as e:
            print(f"   ❌ Erro ao filtrar {nome_interno_filtro}: {e}")
    
    def _identificar_titulo_descricao(self, numero_meta_esperada):
        """Identifica título, subtítulo e descrição da meta"""
        meta_titulo = f"Meta {numero_meta_esperada}"
        meta_subtitulo = "N/D"
        meta_descricao = "Descrição não encontrada"
        
        try:
            textboxes = self.driver.find_elements(By.CSS_SELECTOR, "div.textbox")
            
            for box in textboxes:
                texto_completo = box.text
                
                if f"Meta {numero_meta_esperada}" in texto_completo or f"Meta{numero_meta_esperada}" in texto_completo:
                    # Extração do título
                    try:
                        xpath_titulo = f".//span[contains(text(), 'Meta {numero_meta_esperada}')]"
                        elem_titulo = box.find_element(By.XPATH, xpath_titulo)
                        meta_titulo = elem_titulo.text.strip()
                    except:
                        meta_titulo = texto_completo.split('\n')[0].strip()
                    
                    # Extração do subtítulo (apenas Metas 2, 6, 7, 8)
                    if str(numero_meta_esperada) in ["2", "6", "7", "8"]:
                        try:
                            xpath_sub_texto = ".//span[contains(text(), 'Identificar e julgar')]"
                            elem_sub = box.find_element(By.XPATH, xpath_sub_texto)
                            meta_subtitulo = elem_sub.text.replace(":", "").strip()
                        except:
                            try:
                                xpath_sub_cor = ".//span[contains(@style, 'rgb(204, 204, 204)') or contains(@style, 'rgb(179, 179, 179)')]"
                                elem_sub = box.find_element(By.XPATH, xpath_sub_cor)
                                meta_subtitulo = elem_sub.text.replace(":", "").strip()
                            except:
                                pass
                    
                    # Extração da descrição
                    try:
                        xpath_estadual = ".//li[contains(., 'Justiça Estadual')]"
                        elem_estadual = box.find_element(By.XPATH, xpath_estadual)
                        texto_bruto = elem_estadual.text
                        meta_descricao = texto_bruto.replace("Justiça Estadual:", "").replace("Justiça Estadual", "").strip()
                        return meta_titulo, meta_subtitulo, meta_descricao
                    except:
                        linhas = [l for l in texto_completo.split('\n') if len(l) > 10 and "Meta" not in l and l != meta_subtitulo]
                        if linhas:
                            meta_descricao = linhas[0]
                    
                    return meta_titulo, meta_subtitulo, meta_descricao
        
        except Exception as e:
            print(f"   ⚠️  Erro ao ler textos da Meta {numero_meta_esperada}: {e}")
        
        return meta_titulo, meta_subtitulo, meta_descricao
    
    def adicionar_linha(self, titulo, subtitulo, desc, cat, val):
        """Adiciona registro aos dados extraídos"""
        print(f"   > Capturado: {cat} → {val}")
        self.dados_extraidos.append({
            "Meta": titulo,
            "Subtítulo": subtitulo,
            "Descrição": desc,
            "Categoria": cat,
            "Resultado": val,
            "Data": datetime.now().strftime("%Y-%m-%d %H:%M")
        })
    
    def extrair_dados_grafico(self, numero_meta_esperada):
        """Extrai dados de gráficos (Metas 1 e 2)"""
        print(f"\n📊 Extraindo gráfico (Meta {numero_meta_esperada})...")
        
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao(numero_meta_esperada)
        print(f"   Meta: {meta_titulo}")
        
        container = None
        seletores = ["div[aria-label*='por Ramo']", "div[aria-label*='por Tribunal']", "div[aria-label*='Meta']"]
        
        for seletor in seletores:
            try:
                elems = self.driver.find_elements(By.CSS_SELECTOR, seletor)
                for el in elems:
                    if el.find_elements(By.CSS_SELECTOR, "g.axis"):
                        container = el
                        break
                if container:
                    break
            except:
                continue
        
        if container:
            try:
                categorias = container.find_elements(By.CSS_SELECTOR, "g.axis.y g.tick text")
                valores = container.find_elements(By.CSS_SELECTOR, "g.label-container tspan.label-tspan")
                
                l_cats = [c.get_attribute('textContent').strip() for c in categorias if c.get_attribute('textContent').strip()]
                l_vals = [v.get_attribute('textContent').strip() for v in valores if v.get_attribute('textContent').strip()]
                
                if l_cats and l_vals:
                    limite = min(len(l_cats), len(l_vals))
                    for i in range(limite):
                        self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, l_cats[i], l_vals[i])
                    print(f"   ✅ Extraídos {limite} registros")
                else:
                    print("   ⚠️  Dados vazios")
            except:
                print("   ⚠️  Erro ao ler gráfico")
        else:
            print("   ⚠️  Gráfico não encontrado")
    
    def extrair_kpi(self, numero_meta, kpi_title):
        """Extrai KPI individual (Metas 3, 5)"""
        print(f"\n💎 Extraindo KPI (Meta {numero_meta})...")
        
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao(str(numero_meta))
        print(f"   Meta: {meta_titulo}")
        
        try:
            xpath_card = f"//div[@title='{kpi_title}']/ancestor::div[contains(@class, 'visualWrapper')]"
            card = self.driver.find_element(By.XPATH, xpath_card)
            valor = card.find_element(By.CSS_SELECTOR, "text.value").text.strip()
            print(f"   Valor: {valor}")
            self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, "Total", valor)
        except Exception as e:
            print(f"   ❌ Erro ao extrair KPI: {e}")
    
    def extrair_kpis_meta_4(self):
        """Extrai KPIs da Meta 4 (múltiplos cartões)"""
        print(f"\n💎 Extraindo KPI (Meta 4)...")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao("4")
        print(f"   Meta: {meta_titulo}")
        
        def extrair_valor_do_cartao(container_title, label_nome, campo_nome_saida):
            try:
                xpath_container = f"//div[@title='{container_title}']/ancestor::transform[1]"
                container = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_container)))
                xpath_value = f".//h4[contains(text(), '{label_nome}')]/following-sibling::p[contains(@class, 'bottom')]"
                valor = container.find_element(By.XPATH, xpath_value).text.strip()
                if valor:
                    self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, campo_nome_saida, valor)
                    return True
            except:
                return False
        
        extrair_valor_do_cartao("Meta 4", "Cumprimento", "Total Real")
        extrair_valor_do_cartao("Meta 4 Improb. Administrativa", "Cumprimento", "Total Submeta")
    
    def extrair_kpi_meta_6(self):
        """Extrai KPI da Meta 6 (background SVG)"""
        print(f"\n💎 Extraindo KPI (Meta 6)...")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao("6")
        print(f"   Meta: {meta_titulo}")
        
        try:
            xpath_bg = "//*[local-name()='path' and @data-sub-selection-display-name='Card_Background_Color']"
            backgrounds = self.wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath_bg)))
            valor_encontrado = None
            for bg in backgrounds:
                try:
                    container = bg.find_element(By.XPATH, "./ancestor::visual-modern[1]")
                    texto_container = container.text
                    if "Cumprimento" in texto_container:
                        try:
                            valor_encontrado = container.find_element(By.CSS_SELECTOR, "p.content").text.strip()
                        except:
                            valor_encontrado = container.find_element(By.CSS_SELECTOR, "text.value").text.strip()
                        if valor_encontrado:
                            break
                except:
                    continue
            
            if valor_encontrado:
                print(f"   Valor: {valor_encontrado}")
                self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, "Total", valor_encontrado)
            else:
                print("   ❌ Valor não encontrado")
        except Exception as e:
            print(f"   ❌ Erro ao extrair Meta 6: {e}")
    
    def extrair_kpis_meta_7(self):
        """Extrai KPIs da Meta 7 (Indígenas e Quilombola)"""
        print(f"\n💎 Extraindo KPI (Meta 7)...")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao("7")
        print(f"   Meta: {meta_titulo}")
        
        def extrair_valor_do_cartao(container_title, label_nome, campo_nome_saida):
            try:
                xpath_container = f"//div[@title='{container_title}']/ancestor::transform[1]"
                container = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_container)))
                xpath_value = f".//h4[contains(text(), '{label_nome}')]/following-sibling::p[contains(@class, 'bottom')]"
                valor = container.find_element(By.XPATH, xpath_value).text.strip()
                if valor:
                    self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, campo_nome_saida, valor)
                    return True
            except:
                return False
        
        extrair_valor_do_cartao("Meta 7 Indígenas", "Cumprimento", "Total Indígenas")
        extrair_valor_do_cartao("Meta 7 Quilombola", "Cumprimento", "Total Quilombola")
    
    def extrair_kpis_meta_8(self):
        """Extrai KPIs da Meta 8 (Violência Doméstica e Feminicídio)"""
        print(f"\n💎 Extraindo KPI (Meta 8)...")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao("8")
        print(f"   Meta: {meta_titulo}")
        
        def extrair_valor_do_cartao(container_title, label_nome, campo_nome_saida):
            try:
                xpath_container = f"//div[@title='{container_title}']/ancestor::transform[1]"
                container = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_container)))
                xpath_value = f".//h4[contains(text(), '{label_nome}')]/following-sibling::p[contains(@class, 'bottom')]"
                valor = container.find_element(By.XPATH, xpath_value).text.strip()
                if valor:
                    self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, campo_nome_saida, valor)
                    return True
            except:
                return False
        
        extrair_valor_do_cartao("Violência Doméstica", "Cumprimento", "Total Violência Doméstica")
        extrair_valor_do_cartao("Feminicídio", "Cumprimento", "Total Feminicídio")
    
    def extrair_kpis_meta_10(self):
        """Extrai KPIs da Meta 10 (1º e 2º Grau)"""
        print(f"\n💎 Extraindo KPI (Meta 10)...")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao("10")
        print(f"   Meta: {meta_titulo}")
        
        def extrair_valor_do_cartao(container_title, label_nome, campo_nome_saida):
            try:
                xpath_container = f"//div[@title='{container_title}']/ancestor::transform[1]"
                container = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_container)))
                xpath_value = f".//h4[contains(text(), '{label_nome}')]/following-sibling::p[contains(@class, 'bottom')]"
                valor = container.find_element(By.XPATH, xpath_value).text.strip()
                if valor:
                    self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, campo_nome_saida, valor)
                    return True
            except:
                return False
        
        extrair_valor_do_cartao("1º Grau", "Cumprimento", "Total 1º Grau")
        extrair_valor_do_cartao("2º Grau", "Cumprimento", "Total 2º Grau")
    
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
            print("\n" + "="*60)
            print("📋 META 1")
            print("="*60)
            self.aplicar_filtro_powerbi("ramo_justica", "Justiça Estadual")
            time.sleep(2)
            self.aplicar_filtro_powerbi("sigla_tribunal", self.config.TRIBUNAL_CNJ)
            self.extrair_dados_grafico("1")
            
            # META 2
            print("\n" + "="*60)
            print("📋 META 2")
            print("="*60)
            if self.clicar_elemento_por_texto("Meta 2"):
                self.clicar_botao_laranja_estadual(indice_alvo=1)
                self.aplicar_filtro_powerbi("sigla_tribunal", self.config.TRIBUNAL_CNJ)
                self.extrair_dados_grafico("2")
            
            # META 3
            print("\n" + "="*60)
            print("📋 META 3")
            print("="*60)
            if self.clicar_elemento_por_texto("Meta 3"):
                self.aplicar_filtro_powerbi("sigla_tribunal", self.config.TRIBUNAL_CNJ)
                self.extrair_kpi(3, "Percentual de Cumprimento")
            
            # META 4
            print("\n" + "="*60)
            print("📋 META 4")
            print("="*60)
            if self.clicar_elemento_por_texto("Meta 4"):
                self.clicar_botao_laranja_estadual(indice_alvo=1)
                self.aplicar_filtro_powerbi("sigla_tribunal", self.config.TRIBUNAL_CNJ)
                self.extrair_kpis_meta_4()
            
            # META 5
            print("\n" + "="*60)
            print("📋 META 5")
            print("="*60)
            if self.clicar_elemento_por_texto("Meta 5"):
                self.clicar_botao_laranja_estadual(indice_alvo=2)
                self.aplicar_filtro_powerbi("Tribunal", self.config.TRIBUNAL_CNJ)
                self.extrair_kpi(5, "Cumprimento Meta 5")
            
            # META 6
            print("\n" + "="*60)
            print("📋 META 6")
            print("="*60)
            if self.clicar_elemento_por_texto("Meta 6"):
                self.clicar_botao_laranja_estadual(indice_alvo=1)
                try:
                    self.aplicar_filtro_powerbi("sigla_tribunal", self.config.TRIBUNAL_CNJ)
                except:
                    self.aplicar_filtro_powerbi("Tribunal", self.config.TRIBUNAL_CNJ)
                self.extrair_kpi_meta_6()
            
            # META 7
            print("\n" + "="*60)
            print("📋 META 7")
            print("="*60)
            if self.clicar_elemento_por_texto("Meta 7"):
                self.clicar_botao_laranja_estadual(indice_alvo=1)
                try:
                    self.aplicar_filtro_powerbi("sigla_tribunal", self.config.TRIBUNAL_CNJ)
                except:
                    self.aplicar_filtro_powerbi("Tribunal", self.config.TRIBUNAL_CNJ)
                self.extrair_kpis_meta_7()
            
            # META 8
            print("\n" + "="*60)
            print("📋 META 8")
            print("="*60)
            if self.clicar_elemento_por_texto("Meta 8"):
                self.clicar_botao_laranja_estadual(indice_alvo=1)
                try:
                    self.aplicar_filtro_powerbi("sigla_tribunal", self.config.TRIBUNAL_CNJ)
                except:
                    self.aplicar_filtro_powerbi("Tribunal", self.config.TRIBUNAL_CNJ)
                self.extrair_kpis_meta_8()
            
            # META 9 (nota: pode não existir em todos os anos)
            print("\n" + "="*60)
            print("📋 META 9")
            print("="*60)
            if self.clicar_elemento_por_texto("Meta 9"):
                try:
                    self.clicar_botao_laranja_estadual(indice_alvo=1)
                    self.aplicar_filtro_powerbi("sigla_tribunal", self.config.TRIBUNAL_CNJ)
                    self.extrair_kpi(9, "Percentual de Cumprimento")
                except:
                    print("   ⚠️  Meta 9 não disponível ou erro ao extrair")
            
            # META 10
            print("\n" + "="*60)
            print("📋 META 10")
            print("="*60)
            if self.clicar_elemento_por_texto("Meta 10"):
                self.clicar_botao_laranja_estadual(indice_alvo=1)
                try:
                    self.aplicar_filtro_powerbi("sigla_tribunal", self.config.TRIBUNAL_CNJ)
                except:
                    self.aplicar_filtro_powerbi("Tribunal", self.config.TRIBUNAL_CNJ)
                self.extrair_kpis_meta_10()
            
            self.salvar_excel()
            
            print("\n✅ EXTRAÇÃO CNJ CONCLUÍDA COM SUCESSO!")
            print(f"📊 Total de registros coletados: {len(self.dados_extraidos)}")
            
        except Exception as e:
            print(f"\n❌ Erro durante execução: {e}")
            traceback.print_exc()
        finally:
            input("\n⏸️  Pressione ENTER para fechar o navegador...")
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
    
    # Argumento posicional para anos (mais simples)
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
