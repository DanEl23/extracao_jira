from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup 
import pandas as pd
import time
import traceback
import sys 
import re 
from urllib.parse import quote_plus
import json 

class JiraExtractor:
    def __init__(self, url_jira):
        self.url_jira = url_jira
        self.driver = None
        # Lista plana de todos os tickets extraídos (acumula todos os anos)
        self.dados_extraidos = []
        self.janela_principal = None 
        
    def iniciar_navegador(self):
        """Inicializa o navegador Edge"""
        options = webdriver.EdgeOptions()
        options.add_argument('--start-maximized')
        options.add_argument('--disable-blink-features=AutomationControlled')
        
        print("Iniciando Microsoft Edge...")
        self.driver = webdriver.Edge(options=options)
        self.wait = WebDriverWait(self.driver, 20)
        self.janela_principal = self.driver.current_window_handle
        print("Edge iniciado com sucesso!")
        
    def login_manual_e_aguardar(self, url_jira, jql_base):
        """Abre o Jira e aguarda login, e navega para a JQL inicial."""
        
        self.driver.get(url_jira)
        
        print("\n" + "="*60)
        print("MODO MANUAL DE LOGIN")
        print("="*60)
        print("1. Faça o LOGIN MANUALMENTE no navegador que abriu")
        print(f"2. Após o login, a busca JQL '{jql_base}' será carregada.")
        print("\n⏸️  Pressione ENTER depois de fazer login...")
        
        input() 
        
        print("\n✅ Continuando a extração...")
        time.sleep(2)
        
        # Navega para a JQL BASE para garantir o ponto de partida
        jql_encoded = quote_plus(jql_base)
        url_busca = f"{self.url_jira}issues/?jql={jql_encoded}"
        print(f"➡️ Navegando para o filtro JQL base: {jql_base}")
        self.driver.get(url_busca)
        
        try:
             self.wait.until(
                 EC.presence_of_element_located((By.CSS_SELECTOR, "[data-testid='issue-navigator-list']"))
             )
             print("   ✅ Lista de issues carregada.")
        except:
             print("❌ Aviso: A lista de issues demorou para carregar. Tentando continuar...")
             
        time.sleep(3) 

    def navegar_para_jql(self, jql_encoded):
        """Recarrega a JQL base, usada para resetar o filtro no loop entre anos."""
        url_busca = f"{self.url_jira}issues/?jql={jql_encoded}"
        self.driver.get(url_busca)
        try:
             self.wait.until(
                 EC.presence_of_element_located((By.CSS_SELECTOR, "[data-testid='issue-navigator-list']"))
             )
        except:
             pass 
        time.sleep(3)


    def aplicar_filtro_por_ano(self, nome_campo="Ano da Meta", valor_ano="2024"):
        """
        Simula a adição de um filtro customizado (Ano da Meta) na interface JQL.
        """
        print(f"\n⚙️  Aplicando filtro: {nome_campo} = {valor_ano}")

        try:
            # 1. Clicar no botão 'Mais filtros'
            more_filters_button_selector = "button[data-testid='jql-builder-basic.ui.jql-editor.add-filter.more-button']"
            more_filters_button = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, more_filters_button_selector))
            )
            more_filters_button.click()
            print("   1. Clicou em 'Mais filtros'.")
            
            # 2. Digitar o nome do campo ("Ano da Meta") na busca de filtros
            search_input_xpath = "//input[@aria-label='Pesquisar mais filtros']"
            search_input = self.wait.until(
                EC.presence_of_element_located((By.XPATH, search_input_xpath))
            )
            search_input.send_keys(nome_campo)
            time.sleep(1) 

            # 3. Clicar na opção "Ano da Meta" na lista de filtros
            field_option_xpath = f"//div[@role='option']//div[text()='{nome_campo}']"
            field_option = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, field_option_xpath))
            )
            field_option.click()
            print(f"   2. Selecionou o campo '{nome_campo}'.")
            time.sleep(2) # Pausa crucial para o campo habilitar

            # 4. Localizar o input de busca/seleção para o valor do ano
            value_input_xpath = "//input[@aria-label='Pesquisar Ano da Meta']"
            
            value_input = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, value_input_xpath))
            )
            
            value_input.send_keys(valor_ano)
            print(f"   3a. Digitou o valor '{valor_ano}'.")
            time.sleep(1) 

            # 5. Clicar diretamente na opção de ano filtrada
            value_option_xpath = f"//div[@role='listbox']//div[text()='{valor_ano}']"
            
            value_option = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, value_option_xpath))
            )
            value_option.click() 
            
            print(f"   3b. Clicou na opção '{valor_ano}'.")

            # 6. Aguarda o carregamento da lista filtrada
            self.wait.until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, "[data-testid='issue-navigator.issue-list.content-loading-spinner']"))
            )
            
            print("   ✅ Filtro de ano aplicado com sucesso!")
            time.sleep(2)

        except Exception as e:
            print(f"❌ Erro ao aplicar filtro '{nome_campo}'. Verifique seletores. Erro: {e}")
            raise 


    def exportar_detalhes_impressao(self):
        """Abre o menu de exportação e seleciona 'Detalhes de impressão'"""
        print("\n⚙️  Iniciando exportação por clique...")
        
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
            print("   ✅ Opção 'Detalhes de impressão' clicada. Aguardando nova aba...")
            time.sleep(3) 
            
        except Exception as e:
            print(f"❌ Erro durante a exportação. {e}")
            raise 
    
    
    def processar_aba_exportacao(self):
        """
        Muda para a aba de exportação e extrai os dados, incluindo META_ID e Nº_Meta.
        """
        janelas = self.driver.window_handles
        if len(janelas) < 2:
            print("❌ Nova aba de exportação não foi detectada.")
            return 0

        janela_exportacao = [w for w in janelas if w != self.janela_principal][0]
        self.driver.switch_to.window(janela_exportacao)
        print("\n🔄 Foco mudado para a aba de exportação.")
        
        html_content = self.driver.page_source
        start_time = time.time()
        print("   Iniciando extração de dados com BeautifulSoup...")

        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            blocos_tickets = soup.find_all('table', class_='tableBorder')
            num_tickets_bloco = len(blocos_tickets)
            print(f"Encontrados {num_tickets_bloco} tickets neste lote.")

            for idx, tabela_inicio in enumerate(blocos_tickets):
                registro = {}
                
                # --- 1. META_ID (Parent Key) ---
                parent_key_tag = tabela_inicio.find('a', id='parent_issue_key')
                if parent_key_tag:
                    registro['META_ID'] = parent_key_tag.get_text(strip=True)
                else:
                    registro['META_ID'] = ""
                
                # --- 2. Nº_Meta (Parent Summary) ---
                parent_summary_tag = tabela_inicio.find('a', id='parent_issue_summary')
                if parent_summary_tag:
                    registro['Nº_Meta'] = parent_summary_tag.get_text(strip=True)
                
                
                # --- 3. Chave/Resumo e Meta_apuração ---
                h3_element = tabela_inicio.find('h3', class_='formtitle')
                if h3_element:
                    titulo_completo = h3_element.get_text(strip=True)
                    resumo_link = h3_element.find('a')

                    if resumo_link:
                        current_summary = resumo_link.get_text(strip=True)
                        chave_match = re.search(r'\[([A-Z]+-\d+)\]', titulo_completo)
                        current_chave = chave_match.group(1) if chave_match else f"TICKET-{len(self.dados_extraidos) + idx + 1}"
                        
                        registro['Meta_apuração'] = f"[{current_chave}] {current_summary}"
                        
                        # Mantém as chaves originais
                        registro['Chave'] = current_chave
                        registro['Resumo'] = current_summary 
                    else:
                        registro['Chave'] = f"TICKET-{len(self.dados_extraidos) + idx + 1}"
                        registro['Resumo'] = "Resumo não encontrado"
                else:
                    registro['Chave'] = f"TICKET-{len(self.dados_extraidos) + idx + 1}"
                    registro['Resumo'] = "Resumo não encontrado"


                # --- 4. Extração Dinâmica de Campos ---
                tabelas_ticket = []
                current_element = tabela_inicio
                while current_element:
                    if current_element.name == 'table': tabelas_ticket.append(current_element)
                    proximo = current_element.find_next_sibling()
                    if not proximo: break
                    if proximo.name == 'hr' and 'fullcontent' in proximo.get('class', []): break
                    current_element = proximo

                for tabela in tabelas_ticket:
                    for linha in tabela.find_all('tr'):
                        colunas = linha.find_all('td')
                        if not colunas: continue
                        i = 0
                        while i < len(colunas):
                            label_cell = colunas[i]
                            b_tag = label_cell.find('b')
                            if not b_tag: i += 1; continue
                            
                            rotulo_bruto = b_tag.get_text(separator=' ', strip=True).rstrip(':')
                            rotulo = re.sub(r'\s+', ' ', rotulo_bruto).strip()
                            if not rotulo: i += 1; continue
                            if i + 1 >= len(colunas): i += 1; continue

                            valor_td = colunas[i + 1]
                            valor = valor_td.get_text(separator=' ', strip=True)

                            # Tratamento para datas e HTML
                            if rotulo.lower() in ['data de apuração', 'data de criação', 'atualizado']:
                                time_tag = valor_td.find('time')
                                if time_tag and time_tag.get('datetime'):
                                    valor = time_tag['datetime']

                            if rotulo.lower() == 'informação complementar' or valor_td.find_all('p'):
                                valor = valor_td.decode_contents().strip()

                            if valor:
                                registro[rotulo] = valor

                            i += 2

                self.dados_extraidos.append(registro)

                if (idx + 1) % 100 == 0:
                    print(f"   Processados {idx+1} tickets deste bloco...")

            end_time = time.time()
            elapsed_time = end_time - start_time
            print(f"\n⏱️  Tempo deste bloco ({num_tickets_bloco} tickets): {elapsed_time:.2f} s.")
            print(f"   Total acumulado: {len(self.dados_extraidos)}")
            
        except Exception as e:
            print(f"❌ Erro fatal na extração da aba: {e}")
            traceback.print_exc()
        finally:
            self.driver.close()
            self.driver.switch_to.window(self.janela_principal)
            self.wait = WebDriverWait(self.driver, 20)
            print("↩️  Retornando à aba principal.")
            return num_tickets_bloco


    def montar_dicionario_hierarquico(self):
        """
        Gera o dicionário Pai -> Filhos usando META_ID.
        """
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


    def salvar_em_excel(self, nome_arquivo="dados_jira_exportados_anos.xlsx"):
        """Salva os dados com META_ID na primeira coluna"""
        if not self.dados_extraidos:
            print("\n❌ Nenhum dado para salvar!")
            return
        
        df = pd.DataFrame(self.dados_extraidos)
        
        # ORDENAÇÃO: META_ID PRIMEIRO
        colunas_prioritarias = ['META_ID', 'Chave', 'Resumo', 'Nº_Meta', 'Meta_apuração']
        outras_colunas = [col for col in df.columns if col not in colunas_prioritarias]
        colunas_finais = colunas_prioritarias + outras_colunas
        
        # Filtra colunas existentes para evitar erros
        colunas_existentes = [c for c in colunas_finais if c in df.columns]
        df = df[colunas_existentes]
        
        df.to_excel(nome_arquivo, index=False)
        print(f"\n💾 Excel salvo em: {nome_arquivo}")
        print(f"📋 Colunas iniciais: {', '.join(df.columns[:5].tolist())}...")

    def fechar(self):
        if self.driver:
            self.driver.quit()
            print("\n🔴 Navegador fechado")


# ==================== EXECUÇÃO COM FILTRO DE ANOS ====================

if __name__ == "__main__":
    # --- CONFIGURAÇÕES ---
    URL_JIRA = "https://tjmg.atlassian.net/"
    JQL_BASE = "project = ASPLAGMETA ORDER BY created DESC"
    ANOS_PARA_EXTRAIR = ["2024", "2025", "2026"] 
    # ---------------------

    extrator = JiraExtractor(URL_JIRA)
    
    try:
        print("\n" + "🚀 INICIANDO EXTRAÇÃO POR ANOS " + "\n")
        
        extrator.iniciar_navegador()
        
        # 1. Login Manual
        extrator.login_manual_e_aguardar(URL_JIRA, JQL_BASE)
        
        # 2. Loop pelos Anos
        for ano in ANOS_PARA_EXTRAIR:
            print(f"\n" + "="*60)
            print(f"EXTRAINDO ANO: {ano}")
            print("="*60)
            
            # Aplica filtro
            extrator.aplicar_filtro_por_ano("Ano da Meta", ano)
            
            # Exporta e Processa
            extrator.exportar_detalhes_impressao()
            extrator.processar_aba_exportacao()
            
            # Reseta filtro para o próximo loop
            print("🔁 Resetando filtro...")
            jql_encoded_base = quote_plus(JQL_BASE)
            extrator.navegar_para_jql(jql_encoded_base)
        
        
        # 3. Salva Excel Plano
        extrator.salvar_em_excel("dados_exportados_jira_por_ano.xlsx")
        
        # 4. Salva JSON Hierárquico
        dicionario_metas = extrator.montar_dicionario_hierarquico()
        with open('dicionario_metas_hierarquico_anos.json', 'w', encoding='utf-8') as f:
             json.dump(dicionario_metas, f, indent=4, ensure_ascii=False)
        
        print("\n✅ PROCESSO CONCLUÍDO COM SUCESSO!")
        
    except Exception as e:
        print(f"\n❌ Erro durante a execução: {e}")
        traceback.print_exc()
        
    finally:
        input("\n⏸️  Pressione ENTER para fechar o navegador...")
        extrator.fechar()