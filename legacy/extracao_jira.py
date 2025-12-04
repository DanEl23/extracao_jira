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

class JiraExtractor:
    def __init__(self, url_jira):
        self.url_jira = url_jira
        self.driver = None
        # Lista de dados principal do extrator
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
        """Abre o Jira, aguarda login e navega para a JQL inicial."""
        
        self.driver.get(url_jira)
        
        print("\n" + "="*60)
        print("MODO MANUAL DE LOGIN")
        print("="*60)
        print("1. Faça o LOGIN MANUALMENTE no navegador que abriu")
        print(f"2. NAVEGUE ATÉ O FILTRO JIRA (Ex: {jql_base})")
        print("\n⏸️  Pressione ENTER depois de fazer login e carregar a lista de tickets...")
        
        input() 
        
        print("\n✅ Continuando a extração...")
        time.sleep(3)
        
        try:
             self.wait.until(
                 EC.presence_of_element_located((By.CSS_SELECTOR, "[data-testid='issue-navigator-list']"))
             )
             print("   ✅ Lista de issues carregada.")
        except:
             print("❌ Aviso: A lista de issues demorou para carregar. Tentando continuar...")
             
        time.sleep(2) 


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
        Muda para a aba de exportação e extrai os dados, incluindo META_ID.
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
            print(f"Encontrados {num_tickets_bloco} blocos de tickets para processamento.")

            for idx, tabela_inicio in enumerate(blocos_tickets):
                registro = {}
                
                # --- [NOVO] EXTRAÇÃO DO META_ID (PAI) ---
                # Procura a tag específica que contém a chave do pai
                parent_key_tag = tabela_inicio.find('a', id='parent_issue_key')
                if parent_key_tag:
                    registro['META_ID'] = parent_key_tag.get_text(strip=True)
                else:
                    # Se não tiver pai, deixa vazio ou coloque "Raiz" se preferir
                    registro['META_ID'] = "" 

                # --- EXTRAÇÃO DE CHAVE E RESUMO DO TICKET ATUAL ---
                h3_element = tabela_inicio.find('h3', class_='formtitle')
                if h3_element:
                    titulo_completo = h3_element.get_text(strip=True)
                    resumo_link = h3_element.find('a')

                    if resumo_link:
                        registro['Resumo'] = resumo_link.get_text(strip=True)
                        chave_match = re.search(r'\[([A-Z]+-\d+)\]', titulo_completo)
                        registro['Chave'] = chave_match.group(1) if chave_match else f"TICKET-{idx+1}"
                    else:
                        registro['Resumo'] = titulo_completo
                        registro['Chave'] = f"TICKET-{idx+1}"
                else:
                    registro['Chave'] = f"TICKET-{idx+1}"
                    registro['Resumo'] = "Resumo não encontrado"


                # --- EXTRAÇÃO DE COLUNAS DINÂMICAS (Tabelas) ---
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

                            # Tratamento específico para campos de data e texto longo
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
            print(f"\n⏱️  Tempo total de extração deste bloco ({num_tickets_bloco} tickets): {elapsed_time:.2f} segundos.")
            print(f"   Total de registros acumulados: {len(self.dados_extraidos)}")
            
        except Exception as e:
            print(f"❌ Erro fatal na extração da aba: {e}")
            traceback.print_exc()
        finally:
            self.driver.close()
            self.driver.switch_to.window(self.janela_principal)
            self.wait = WebDriverWait(self.driver, 20)
            print("↩️  Retornando à aba principal.")
            return num_tickets_bloco


    def salvar_em_excel(self, nome_arquivo="dados_jira_exportados.xlsx"):
        """Salva os dados extraídos em Excel com META_ID na primeira coluna"""
        if not self.dados_extraidos:
            print("\n❌ Nenhum dado para salvar!")
            return
        
        df = pd.DataFrame(self.dados_extraidos)
        
        # --- [NOVO] ORDENAÇÃO DE COLUNAS ---
        # Garante que META_ID seja a primeira, seguida de Chave e Resumo
        colunas_prioritarias = ['META_ID', 'Chave', 'Resumo']
        
        # Pega o restante das colunas dinâmicas que não estão nas prioritárias
        outras_colunas = [col for col in df.columns if col not in colunas_prioritarias]
        
        # Junta tudo
        colunas_finais = colunas_prioritarias + outras_colunas
        
        # Filtra o DataFrame (apenas colunas que realmente existem)
        # (Isso evita erro se por acaso 'META_ID' não for encontrada em nenhum ticket, embora tenhamos forçado a criação)
        colunas_existentes = [c for c in colunas_finais if c in df.columns]
        
        df = df[colunas_existentes]
        
        df.to_excel(nome_arquivo, index=False)
        print(f"\n💾 Dados salvos em: {nome_arquivo}")
        print(f"📊 Total de registros: {len(df)}")
        print(f"📋 Colunas (primeiras): {', '.join(df.columns[:5].tolist())}...")

    def fechar(self):
        """Fecha o navegador"""
        if self.driver:
            self.driver.quit()
            print("\n🔴 Navegador fechado")


# ==================== EXEMPLO DE USO FINAL ====================

if __name__ == "__main__":
    # CONFIGURE ESTAS VARIÁVEIS
    URL_JIRA = "https://tjmg.atlassian.net/"
    JQL_BASE = "project = ASPLAGMETA ORDER BY created DESC" 
    #---------------------------

    extrator = JiraExtractor(URL_JIRA)
    
    try:
        print("\n" + "🚀 INICIANDO EXTRAÇÃO DO JIRA (LOTE ÚNICO) " + "\n")
        
        extrator.iniciar_navegador()
        
        # 1. Login Manual e Navegação para a JQL
        extrator.login_manual_e_aguardar(URL_JIRA, JQL_BASE)
        
        print(f"\n" + "="*60)
        print("INICIANDO EXTRAÇÃO")
        print("="*60)
            
        # 2. Inicia a exportação (simula clique e abre nova aba)
        extrator.exportar_detalhes_impressao()
            
        # 3. Processa a nova aba e extrai os dados
        extrator.processar_aba_exportacao()
        
        # 4. Salva no arquivo solicitado com a nova coluna
        extrator.salvar_em_excel("dados_exportados_jira.xlsx")
        
        print("\n✅ PROCESSO CONCLUÍDO COM SUCESSO!")
        
    except Exception as e:
        print(f"\n❌ Erro durante a execução: {e}")
        traceback.print_exc()
        
    finally:
        input("\n⏸️  Pressione ENTER para fechar o navegador...")
        extrator.fechar()