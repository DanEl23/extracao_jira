from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
from datetime import datetime
import pandas as pd
from selenium.common.exceptions import TimeoutException

class AutomacaoPainelCNJ:
    def __init__(self):
        print("Inicializando robô...")
        opcoes = webdriver.ChromeOptions()
        opcoes.add_argument('--start-maximized')
        opcoes.add_argument('--disable-notifications')
        
        self.driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=opcoes
        )
        self.wait = WebDriverWait(self.driver, 20)
        self.dados_extraidos = []

    def acessar_painel(self):
        print("Acessando site do CNJ...")
        self.driver.get("https://justica-em-numeros.cnj.jus.br/painel-metas/")
        time.sleep(10)

    def entrar_no_iframe(self):
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

    # --- MÉTODO CORRIGIDO: Suporte a múltiplas cores de cinza e busca por texto ---
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

                    # 2. Extração do Subtítulo (Apenas Metas 2, 6, 7, 8)
                    if str(numero_meta_esperada) in ["2", "6", "7", "8"]:
                        try:
                            # ESTRATÉGIA A: Busca por TEXTO (Mais seguro contra mudança de cores)
                            # Todas essas metas começam com "Identificar e julgar"
                            xpath_sub_texto = ".//span[contains(text(), 'Identificar e julgar')]"
                            elem_sub = box.find_element(By.XPATH, xpath_sub_texto)
                            meta_subtitulo = elem_sub.text.replace(":", "").strip()
                        except:
                            try:
                                # ESTRATÉGIA B: Busca por COR (Fallback)
                                # rgb(204, 204, 204) = Meta 2
                                # rgb(179, 179, 179) = Meta 6, 7, 8
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

    # --- FUNÇÃO GRÁFICOS (METAS 1 e 2) ---
    def extrair_dados_da_aba(self, numero_meta_esperada=None):
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
                    # Verifica se a string é duplicada (primeira metade = segunda metade)
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

    # --- FUNÇÃO KPI META 1 (Julgar mais processos que os distribuídos) ---
    def extrair_kpi_meta_1_total(self):
        print(f"\n--- Iniciando Extração KPI Total (Meta 1) ---")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao(numero_meta_esperada="1")
        print(f"📌 Meta Identificada: {meta_titulo}")
        
        print(f"🔍 Buscando cartão 'Julgar mais processos que os distribuídos'...")
        try:
            # Busca específica pelo título do card
            xpath_card = "//div[@title='Julgar mais processos que os distribuídos']/ancestor::div[contains(@class, 'visualWrapper')]"
            card = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_card)))
            
            # Extrai o valor do SVG text com classe "value"
            valor = card.find_element(By.CSS_SELECTOR, "text.value tspan").text.strip()
            print(f"💎 Valor encontrado: {valor}")
            
            # Adiciona com categoria "Total" como solicitado
            self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, "Total", valor)
        except Exception as e:
            print(f"❌ Erro ao extrair KPI Total da Meta 1: {e}")
    
    # --- FUNÇÃO KPI META 2 (Extrair múltiplos KPIs por instância) ---
    def extrair_kpis_meta_2(self):
        print(f"\n--- Iniciando Extração KPIs Meta 2 ---")
        meta_titulo, meta_subtitulo, meta_descricao = self._identificar_titulo_descricao(numero_meta_esperada="2")
        print(f"📌 Meta Identificada: {meta_titulo}")
        
        # Mapeamento dos títulos dos cards e suas instâncias
        instancias_titles = {
            "1º Grau": "1º Grau",
            "2º Grau": "2º Grau",
            "Juizados e Turmas": "Juizados e Turmas",
            "Processos mais Antigos": "Processos mais Antigos"
        }
        
        for titulo_card, instancia in instancias_titles.items():
            print(f"\n🔍 Buscando card 'Cumprimento' da instância '{titulo_card}'...")
            try:
                # Localizar o card principal pelo título
                xpath_wrapper = f"//div[@title='{titulo_card}']/ancestor::div[contains(@class, 'visualWrapper')]"
                wrapper = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_wrapper)))
                
                # Extrair apenas o KPI "Cumprimento" (primeiro card)
                try:
                    # Busca específica pelo card com label "Cumprimento"
                    cumprimento_xpath = ".//h4[contains(text(), 'Cumprimento')]/following-sibling::p"
                    valor_elem = wrapper.find_element(By.XPATH, cumprimento_xpath)
                    valor = valor_elem.text.strip()
                    
                    print(f"   💎 {instancia} - Cumprimento: {valor}")
                    
                    # Adiciona linha com categoria sendo a instância
                    self.adicionar_linha(meta_titulo, meta_subtitulo, meta_descricao, instancia, valor)
                    
                except Exception as e:
                    print(f"   ⚠️ Erro ao extrair Cumprimento: {e}")
                        
            except Exception as e:
                print(f"❌ Erro ao processar instância '{titulo_card}': {e}")
    
    # --- FUNÇÃO KPI GENERALIZADA (META 3, 5, 6) ---
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

    # --- FUNÇÃO META 4 (Coleta Múltipla) ---
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

    # --- FUNÇÃO META 6 (Extração por Background SVG) ---
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

    # --- FUNÇÃO META 7 ---
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

    # --- FUNÇÃO META 8 ---
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

    # --- FUNÇÃO META 10 ---
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
        if self.dados_extraidos:
            df = pd.DataFrame(self.dados_extraidos)
            arquivo = "exports/resultados_cnj.xlsx"
            df.to_excel(arquivo, index=False)
            print(f"\n✅ Arquivo salvo: {arquivo}")
        else:
            print("\n⚠️ Nenhum dado para salvar.")

    def executar(self):
        """Fluxo Principal"""
        self.acessar_painel()
        self.entrar_no_iframe()
        
        # META 1
        print("\n=== 🏁 INICIANDO META 1 ===")
        self.aplicar_filtro_powerbi("ramo_justica", "Justiça Estadual")
        time.sleep(2)
        self.aplicar_filtro_powerbi("sigla_tribunal", "TJMG")
        self.extrair_dados_da_aba(numero_meta_esperada="1")
        self.extrair_kpi_meta_1_total()  # Extrai o KPI "Julgar mais processos que os distribuídos"
        
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
        self.driver.quit()

if __name__ == "__main__":
    robo = AutomacaoPainelCNJ()
    robo.executar()