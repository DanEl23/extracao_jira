import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from gerador_relatorio import *

class AplicativoRelatorio:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Relatórios - TJMG")
        self.root.geometry("500x300")
        
        # Arquivo selecionado
        self.arquivo_excel = None
        
        # Interface
        self.criar_interface()
    
    def criar_interface(self):
        # Título
        titulo = tk.Label(self.root, text="📊 Gerador de Relatórios de Metas", 
                         font=("Arial", 16, "bold"))
        titulo.pack(pady=20)
        
        # Botão selecionar arquivo
        btn_selecionar = tk.Button(self.root, text="📁 Selecionar Arquivo Excel", 
                                   command=self.selecionar_arquivo,
                                   font=("Arial", 12),
                                   bg="#4ECDC4", fg="white",
                                   padx=20, pady=10)
        btn_selecionar.pack(pady=10)
        
        # Label arquivo selecionado
        self.label_arquivo = tk.Label(self.root, text="Nenhum arquivo selecionado",
                                     font=("Arial", 10), fg="gray")
        self.label_arquivo.pack()
        
        # Botão gerar
        self.btn_gerar = tk.Button(self.root, text="🚀 Gerar Relatório", 
                                   command=self.gerar_relatorio_thread,
                                   font=("Arial", 12, "bold"),
                                   bg="#FF6B35", fg="white",
                                   padx=20, pady=10,
                                   state="disabled")
        self.btn_gerar.pack(pady=20)
        
        # Barra de progresso
        self.progresso = ttk.Progressbar(self.root, mode='indeterminate')
        self.progresso.pack(fill=tk.X, padx=50)
        
        # Status
        self.label_status = tk.Label(self.root, text="", font=("Arial", 9), fg="blue")
        self.label_status.pack(pady=10)
    
    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivo:
            self.arquivo_excel = arquivo
            nome_arquivo = os.path.basename(arquivo)
            self.label_arquivo.config(text=f"✅ {nome_arquivo}", fg="green")
            self.btn_gerar.config(state="normal")
    
    def gerar_relatorio_thread(self):
        # Executar em thread separada para não travar a interface
        thread = threading.Thread(target=self.gerar_relatorio)
        thread.start()
    
    def gerar_relatorio(self):
        try:
            # Desabilitar botão
            self.btn_gerar.config(state="disabled")
            self.progresso.start()
            self.label_status.config(text="⏳ Gerando relatório...", fg="blue")
            
            # Configurar arquivo
            Config.ARQUIVO_EXCEL = self.arquivo_excel
            
            # Gerar relatório
            criar_pasta_saida()
            df = carregar_dados()
            if df is None:
                raise Exception("Erro ao carregar dados")
            
            grupos = agrupar_por_macrodesafio(df)
            doc = criar_documento()
            adicionar_cabecalho_relatorio(doc)
            
            for idx, (macrodesafio, df_grupo) in enumerate(grupos):
                adicionar_secao_macrodesafio(doc, macrodesafio, df_grupo, primeira_secao=(idx==0))
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            nome_arquivo = f"{Config.NOME_RELATORIO}_{timestamp}.docx"
            caminho_completo = os.path.join(Config.PASTA_SAIDA, nome_arquivo)
            
            doc.save(caminho_completo)
            
            # Sucesso
            self.progresso.stop()
            self.label_status.config(text="✅ Relatório gerado com sucesso!", fg="green")
            
            messagebox.showinfo("Sucesso", 
                              f"Relatório gerado!\n\n📁 {caminho_completo}\n\n"
                              f"📊 {len(df)} registros processados")
            
            # Abrir pasta
            os.startfile(Config.PASTA_SAIDA)
            
        except Exception as e:
            self.progresso.stop()
            self.label_status.config(text="❌ Erro ao gerar relatório", fg="red")
            messagebox.showerror("Erro", f"Erro ao gerar relatório:\n\n{str(e)}")
        
        finally:
            self.btn_gerar.config(state="normal")

if __name__ == "__main__":
    root = tk.Tk()  # Alterado de TkRoot() para Tk()
    app = AplicativoRelatorio(root)
    root.mainloop()