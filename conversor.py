import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
import threading
import time
import os

class ConversorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor CSV para Site")
        self.root.geometry("500x250")
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f0f0")

        # Estilo
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TProgressbar", thickness=25)

        # Título
        self.lbl_titulo = tk.Label(root, text="Conversor de Acervo", font=("Segoe UI", 16, "bold"), bg="#f0f0f0", fg="#002b5c")
        self.lbl_titulo.pack(pady=20)

        # Instrução
        self.lbl_info = tk.Label(root, text="Selecione seu arquivo .csv (separado por ponto e vírgula)", font=("Segoe UI", 10), bg="#f0f0f0")
        self.lbl_info.pack(pady=5)

        # Botão
        self.btn_selecionar = tk.Button(root, text="📂 Selecionar CSV e Converter", command=self.iniciar_thread, 
                                      font=("Segoe UI", 11), bg="#29a771", fg="white", padx=20, pady=5, relief="flat")
        self.btn_selecionar.pack(pady=15)

        # Barra de Progresso
        self.progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate", style="TProgressbar")
        self.progress.pack(pady=10)

        # Status
        self.lbl_status = tk.Label(root, text="Aguardando...", font=("Segoe UI", 9, "italic"), bg="#f0f0f0", fg="#666")
        self.lbl_status.pack(pady=5)

    def iniciar_thread(self):
        # Roda o processo em paralelo para não travar a janela
        threading.Thread(target=self.processar_arquivo).start()

    def processar_arquivo(self):
        # 1. Selecionar Arquivo de Entrada
        csv_path = filedialog.askopenfilename(title="Selecione o arquivo CSV", filetypes=[("Arquivo CSV", "*.csv")])
        if not csv_path:
            return

        self.btn_selecionar.config(state="disabled")
        self.lbl_status.config(text="Lendo arquivo...")
        self.progress['value'] = 0

        try:
            # Conta linhas para a barra de progresso
            with open(csv_path, 'r', encoding='utf-8', errors='ignore') as f:
                total_linhas = sum(1 for row in f)

            dados_convertidos = []
            
            # 2. Processamento
            with open(csv_path, 'r', encoding='utf-8', errors='ignore') as f:
                # Assume que o separador é ponto e vírgula (padrão do seu arquivo)
                leitor = csv.reader(f, delimiter=';')
                
                contador = 0
                for linha in leitor:
                    contador += 1
                    # Atualiza barra de progresso
                    progresso = (contador / total_linhas) * 100
                    self.progress['value'] = progresso
                    self.lbl_status.config(text=f"Processando linha {contador} de {total_linhas}...")
                    self.root.update_idletasks()

                    # Pula linhas vazias ou incompletas
                    if len(linha) < 4:
                        continue

                    # Extrai os dados (Ano; Edicao; Data; Link)
                    ano = linha[0].strip()
                    edicao = linha[1].strip()
                    data = linha[2].strip()
                    link = linha[3].strip()

                    # Formata a linha no padrão JavaScript Object
                    obj_str = f"    {{ano: {ano}, edicao: '{edicao}', data: '{data}', link: '{link}'}}"
                    dados_convertidos.append(obj_str)
                    
                    # Pequena pausa só pra dar pra ver a barra enchendo (opcional)
                    # time.sleep(0.001) 

            # 3. Montagem do Arquivo Final
            conteudo_final = "const baseDeDados = [\n" + ",\n".join(dados_convertidos) + "\n];"

            # 4. Salvar
            save_path = filedialog.asksaveasfilename(defaultextension=".js", initialfile="dados.js", title="Salvar como")
            
            if save_path:
                with open(save_path, 'w', encoding='utf-8') as f:
                    f.write(conteudo_final)
                
                self.lbl_status.config(text="Concluído com sucesso!", fg="green")
                messagebox.showinfo("Sucesso", "Arquivo dados.js gerado com sucesso!\nAgora é só colocar na pasta 'js' do site.")
            else:
                self.lbl_status.config(text="Cancelado pelo usuário.")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
            self.lbl_status.config(text="Erro na conversão.")
        
        finally:
            self.btn_selecionar.config(state="normal")
            self.progress['value'] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = ConversorApp(root)
    root.mainloop()