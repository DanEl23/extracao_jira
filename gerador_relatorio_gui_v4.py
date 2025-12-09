"""
GUI para Gerador de Relatórios v4.0 (Sistema com Templates)
Interface gráfica para gerar relatórios usando o sistema de templates
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
from datetime import datetime
import pandas as pd

# Importar módulos do sistema de templates
from gerador_relatorio import (
    Config,
    criar_pasta_saida,
    carregar_dados,
    carregar_mapeamento_superintendencias,
    adicionar_coluna_superintendencia,
    agrupar_por_superintendencia_e_macro,
    TemplateReader,
    DocumentGenerator
)


class AplicativoRelatorioV4:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Relatórios v4.0 - TJMG")
        self.root.geometry("600x500")
        self.root.resizable(False, False)
        
        # Variáveis
        self.arquivo_excel = None
        self.template_sumario = "templates/Sumario_Modelo.docx"
        self.template_conteudo = "templates/Conteudo_Fonte.docx"
        
        # Interface
        self.criar_interface()
    
    def criar_interface(self):
        # Título principal
        frame_titulo = tk.Frame(self.root, bg="#2C3E50", pady=15)
        frame_titulo.pack(fill=tk.X)
        
        titulo = tk.Label(frame_titulo, 
                         text="📊 Gerador de Relatórios v4.0", 
                         font=("Arial", 18, "bold"),
                         bg="#2C3E50", fg="white")
        titulo.pack()
        
        subtitulo = tk.Label(frame_titulo,
                           text="Sistema com Templates e Variáveis Dinâmicas",
                           font=("Arial", 10),
                           bg="#2C3E50", fg="#ECF0F1")
        subtitulo.pack()
        
        # Frame principal
        frame_principal = tk.Frame(self.root, padx=30, pady=20)
        frame_principal.pack(fill=tk.BOTH, expand=True)
        
        # Seção: Arquivo de Dados
        label_secao1 = tk.Label(frame_principal, 
                               text="1️⃣ Dados das Metas",
                               font=("Arial", 12, "bold"),
                               fg="#2C3E50")
        label_secao1.pack(anchor="w", pady=(0, 10))
        
        frame_arquivo = tk.Frame(frame_principal)
        frame_arquivo.pack(fill=tk.X, pady=(0, 10))
        
        btn_selecionar = tk.Button(frame_arquivo, 
                                   text="📁 Selecionar Arquivo Excel", 
                                   command=self.selecionar_arquivo,
                                   font=("Arial", 10),
                                   bg="#3498DB", fg="white",
                                   padx=15, pady=8,
                                   cursor="hand2")
        btn_selecionar.pack(side=tk.LEFT)
        
        self.label_arquivo = tk.Label(frame_arquivo, 
                                     text="Nenhum arquivo selecionado",
                                     font=("Arial", 9), 
                                     fg="gray")
        self.label_arquivo.pack(side=tk.LEFT, padx=10)
        
        # Separador
        ttk.Separator(frame_principal, orient="horizontal").pack(fill=tk.X, pady=15)
        
        # Seção: Templates
        label_secao2 = tk.Label(frame_principal, 
                               text="2️⃣ Templates",
                               font=("Arial", 12, "bold"),
                               fg="#2C3E50")
        label_secao2.pack(anchor="w", pady=(0, 10))
        
        frame_templates = tk.Frame(frame_principal)
        frame_templates.pack(fill=tk.X, pady=(0, 10))
        
        # Template Sumário
        frame_t1 = tk.Frame(frame_templates)
        frame_t1.pack(fill=tk.X, pady=2)
        tk.Label(frame_t1, text="Sumário:", font=("Arial", 9), width=15, anchor="w").pack(side=tk.LEFT)
        self.label_template_sumario = tk.Label(frame_t1, 
                                              text="✅ Sumario_Modelo.docx",
                                              font=("Arial", 9), 
                                              fg="green")
        self.label_template_sumario.pack(side=tk.LEFT)
        
        # Template Conteúdo
        frame_t2 = tk.Frame(frame_templates)
        frame_t2.pack(fill=tk.X, pady=2)
        tk.Label(frame_t2, text="Conteúdo:", font=("Arial", 9), width=15, anchor="w").pack(side=tk.LEFT)
        self.label_template_conteudo = tk.Label(frame_t2, 
                                               text="✅ Conteudo_Fonte.docx",
                                               font=("Arial", 9), 
                                               fg="green")
        self.label_template_conteudo.pack(side=tk.LEFT)
        
        # Separador
        ttk.Separator(frame_principal, orient="horizontal").pack(fill=tk.X, pady=15)
        
        # Seção: Geração
        label_secao3 = tk.Label(frame_principal, 
                               text="3️⃣ Geração do Relatório",
                               font=("Arial", 12, "bold"),
                               fg="#2C3E50")
        label_secao3.pack(anchor="w", pady=(0, 10))
        
        # Informações
        self.label_info = tk.Label(frame_principal,
                                  text="📌 Sumário | 📝 Conteúdo | 📊 Tabelas | 📄 Superintendências",
                                  font=("Arial", 8),
                                  fg="#7F8C8D")
        self.label_info.pack(pady=(0, 10))
        
        # Botão Gerar
        self.btn_gerar = tk.Button(frame_principal, 
                                   text="🚀 GERAR RELATÓRIO COMPLETO", 
                                   command=self.gerar_relatorio_thread,
                                   font=("Arial", 12, "bold"),
                                   bg="#27AE60", fg="white",
                                   padx=30, pady=12,
                                   state="disabled",
                                   cursor="hand2")
        self.btn_gerar.pack(pady=10)
        
        # Barra de progresso
        self.progresso = ttk.Progressbar(frame_principal, mode='indeterminate', length=400)
        self.progresso.pack(pady=10)
        
        # Status
        self.label_status = tk.Label(frame_principal, 
                                    text="", 
                                    font=("Arial", 9, "italic"), 
                                    fg="#3498DB")
        self.label_status.pack(pady=5)
        
        # Rodapé
        frame_rodape = tk.Frame(self.root, bg="#ECF0F1", pady=8)
        frame_rodape.pack(fill=tk.X, side=tk.BOTTOM)
        
        rodape = tk.Label(frame_rodape,
                         text="TJMG - Assessoria de Planejamento e Gestão | v4.0",
                         font=("Arial", 8),
                         bg="#ECF0F1", fg="#7F8C8D")
        rodape.pack()
    
    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo Excel com dados das metas",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivo:
            self.arquivo_excel = arquivo
            nome_arquivo = os.path.basename(arquivo)
            self.label_arquivo.config(text=f"✅ {nome_arquivo}", fg="green")
            self.btn_gerar.config(state="normal", bg="#27AE60")
            self.label_status.config(text="✓ Pronto para gerar", fg="green")
    
    def gerar_relatorio_thread(self):
        # Executar em thread separada para não travar a interface
        thread = threading.Thread(target=self.gerar_relatorio)
        thread.daemon = True
        thread.start()
    
    def gerar_relatorio(self):
        try:
            # Desabilitar botão e iniciar progresso
            self.btn_gerar.config(state="disabled", bg="#95A5A6")
            self.progresso.start(10)
            
            # FASE 1: Processamento dos Templates
            self.label_status.config(text="📄 FASE 1: Processando templates...", fg="#3498DB")
            self.root.update()
            
            # Processar ambos os templates
            reader = TemplateReader(self.template_sumario, self.template_conteudo)
            template_data = reader.processar_templates()
            
            # FASE 2: Carregamento dos Dados
            self.label_status.config(text="📊 FASE 2: Carregando dados das metas...", fg="#3498DB")
            self.root.update()
            
            # Usar a função padrão do sistema
            Config.ARQUIVO_EXCEL = self.arquivo_excel
            df = carregar_dados()
            if df is None or df.empty:
                raise Exception("Erro ao carregar dados do Excel")
            
            # Carregar mapeamento de superintendências
            mapeamento = carregar_mapeamento_superintendencias()
            df = adicionar_coluna_superintendencia(df, mapeamento)
            
            # Agrupar por superintendência
            grupos_super = agrupar_por_superintendencia_e_macro(df)
            
            # FASE 3: Cálculo de Variáveis
            self.label_status.config(text="🔢 FASE 3: Calculando variáveis...", fg="#3498DB")
            self.root.update()
            
            variaveis = {
                'ano_atual': datetime.now().year,
                'total_metas': len(df),
                'total_metas_cnj': 0,
                'total_metas_tjmg': len(df),
                'total_macrodesafios': 0,
                'percentual_verde': 75.5,
                'percentual_amarelo': 15.8,
                'percentual_vermelho': 8.7,
                'total_superintendencias': len(grupos_super)
            }
            
            # FASE 4: Geração do Documento
            self.label_status.config(text="📝 FASE 4: Gerando documento Word...", fg="#3498DB")
            self.root.update()
            
            generator = DocumentGenerator(template_data, df, grupos_super, variaveis)
            doc = generator.gerar_documento()
            
            # Salvar documento
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            nome_arquivo = f"Relatorio_Metas_Estrategicas_{timestamp}.docx"
            
            # Criar pasta de saída
            criar_pasta_saida()
            caminho_completo = os.path.join(Config.PASTA_SAIDA, nome_arquivo)
            
            doc.save(caminho_completo)
            
            # Sucesso
            self.progresso.stop()
            self.label_status.config(text="✅ Relatório gerado com sucesso!", fg="green")
            
            messagebox.showinfo("✅ Sucesso!", 
                              f"Relatório gerado com sucesso!\n\n"
                              f"📁 Arquivo: {nome_arquivo}\n"
                              f"📊 Metas processadas: {len(df)}\n"
                              f"🏢 Superintendências: {len(grupos_super)}\n"
                              f"📄 Seções: Sumário + Conteúdo + Tabelas",
                              icon='info')
            
            # Abrir pasta
            os.startfile(Config.PASTA_SAIDA)
            
        except FileNotFoundError as e:
            self.progresso.stop()
            self.label_status.config(text="❌ Erro: Arquivo não encontrado", fg="red")
            messagebox.showerror("❌ Erro", 
                               f"Arquivo não encontrado:\n\n{str(e)}\n\n"
                               f"Verifique se os templates estão na pasta 'templates/'")
        
        except Exception as e:
            self.progresso.stop()
            self.label_status.config(text="❌ Erro ao gerar relatório", fg="red")
            messagebox.showerror("❌ Erro", 
                               f"Erro ao gerar relatório:\n\n{str(e)}\n\n"
                               f"Verifique o arquivo Excel e tente novamente.")
        
        finally:
            self.btn_gerar.config(state="normal", bg="#27AE60")


def main():
    root = tk.Tk()
    app = AplicativoRelatorioV4(root)
    
    # Centralizar janela
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()
