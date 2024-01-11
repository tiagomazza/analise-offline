import tkinter as tk
from tkinter import filedialog
import pandas as pd

# Listas globais para armazenar os valores de cada coluna
nomeLista = []
emailVendedorLista = []
codigoVendedorLista = []
meta1Lista = []
meta2Lista = []
meta3Lista = []
meta4Lista = []

class Aplicacao:
    def __init__(self, janela):
        self.janela = janela
        self.janela.title("Carregar XLSX")
        
        self.dataframe = None  # Variável para armazenar o DataFrame

        # Criar um botão para carregar o arquivo XLSX
        self.botao_carregar_xlsx = tk.Button(self.janela, text="Carregar XLSX", command=self.carregar_xlsx)
        self.botao_carregar_xlsx.pack(pady=20)

    def carregar_xlsx(self):
        # Abrir a caixa de diálogo para seleção do arquivo XLSX
        caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])

        if caminho_arquivo:
            try:
                # Ler o arquivo XLSX e criar um DataFrame
                self.dataframe = pd.read_excel(caminho_arquivo)

                # Preencher as listas globais com os valores de cada coluna
                global nomeLista, emailVendedorLista, codigoVendedorLista, meta1Lista, meta2Lista, meta3Lista, meta4Lista
                nomeLista = self.dataframe['Nome'].astype(str).tolist()
                emailVendedorLista = self.dataframe['Email'].astype(str).tolist()
                codigoVendedorLista = self.dataframe['Código'].astype(str).tolist()
                meta1Lista = self.dataframe['Meta 1'].astype(str).tolist()
                meta2Lista = self.dataframe['Meta 2'].astype(str).tolist()
                meta3Lista = self.dataframe['Meta 3'].astype(str).tolist()
                meta4Lista = self.dataframe['Meta 4'].astype(str).tolist()

                # Exibir mensagem de sucesso
                mensagem_sucesso = "XLSX carregado com sucesso!"
                tk.messagebox.showinfo("Sucesso", mensagem_sucesso)
            except Exception as e:
                # Exibir mensagem de erro
                tk.messagebox.showerror("Erro", f"Erro ao carregar o arquivo XLSX: {str(e)}")

# Criar a janela principal
janela_principal = tk.Tk()

# Criar uma instância da classe Aplicacao
app = Aplicacao(janela_principal)

# Iniciar o loop principal
janela_principal.mainloop()

# Imprimir o conteúdo das listas no final do código
print("Conteúdo das Listas:")
print("nomeLista:", nomeLista)
print("emailVendedorLista:", emailVendedorLista)
print("codigoVendedorLista:", codigoVendedorLista)
print("meta1Lista:", meta1Lista)
print("meta2Lista:", meta2Lista)
print("meta3Lista:", meta3Lista)
print("meta4Lista:", meta4Lista)
