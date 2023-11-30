import tkinter as tk
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns

def abrir_excel():
    root.filename = filedialog.askopenfilename(initialdir="/", title="Selecione o arquivo Excel", filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    entry_arquivo.delete(0, tk.END)
    entry_arquivo.insert(0, root.filename)

def carregar_dados():
    arquivo_excel = entry_arquivo.get()
    nome_planilha = entry_planilha.get()
    
    if arquivo_excel and nome_planilha:
        try:
            df = pd.read_excel(arquivo_excel, sheet_name=nome_planilha)
            
            # Criar nova janela para exibir os gráficos
            nova_janela = tk.Toplevel(root)
            nova_janela.title("Visualização de Dados")

            # Gráfico de linhas para visualizar as vendas ao longo dos anos
            fig1 = plt.figure(figsize=(8, 5))
            plt.plot(df['Ano'], df['Vendas'], marker='o', linestyle='-', color='b')
            plt.title('Vendas ao longo dos anos')
            plt.xlabel('Ano')
            plt.ylabel('Vendas')
            plt.grid(True)
            canvas_vendas = FigureCanvasTkAgg(fig1, master=nova_janela)
            canvas_vendas.draw()
            canvas_vendas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

            # Gráfico de barras para comparar vendas e lucro por ano usando Seaborn
            fig2 = plt.figure(figsize=(8, 5))
            sns.barplot(x='Ano', y='Vendas', data=df, color='skyblue', label='Vendas')
            sns.barplot(x='Ano', y='Lucro', data=df, color='orange', label='Lucro')
            plt.title('Comparação de Vendas e Lucro por Ano')
            plt.xlabel('Ano')
            plt.ylabel('Valor')
            plt.legend()
            canvas_lucro = FigureCanvasTkAgg(fig2, master=nova_janela)
            canvas_lucro.draw()
            canvas_lucro.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=1)

            # Configura a função para fechar a janela e parar a execução do programa
            def fechar_janela():
                nova_janela.destroy()
                root.quit()

            # Adiciona um evento para quando a janela secundária for fechada
            nova_janela.protocol("WM_DELETE_WINDOW", fechar_janela)

        except Exception as e:
            print("Erro ao carregar os dados:", e)
    else:
        print("Por favor, selecione o arquivo do Excel e insira o nome da planilha.")

root = tk.Tk()
root.title("Seleção de Arquivo Excel e Planilha")

label_arquivo = tk.Label(root, text="Arquivo Excel:")
label_arquivo.grid(row=0, column=0)

entry_arquivo = tk.Entry(root, width=50)
entry_arquivo.grid(row=0, column=1)

btn_selecionar_arquivo = tk.Button(root, text="Selecionar Arquivo", command=abrir_excel)
btn_selecionar_arquivo.grid(row=0, column=2)

label_planilha = tk.Label(root, text="Nome da Planilha:")
label_planilha.grid(row=1, column=0)

entry_planilha = tk.Entry(root)
entry_planilha.grid(row=1, column=1)

btn_carregar = tk.Button(root, text="Carregar Dados", command=carregar_dados)
btn_carregar.grid(row=2, columnspan=3)

root.mainloop()
