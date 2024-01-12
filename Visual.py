import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import tempfile
import os
from datetime import datetime
import win32com.client as win32

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
        self.janela.title("Envio de metas por email")

        self.dataframe = None  # Variável para armazenar o DataFrame

        # Criar um botão para carregar o arquivo XLSX
        self.botao_carregar_xlsx = tk.Button(self.janela, text="Carregue os dados", command=self.carregar_xlsx)
        self.botao_carregar_xlsx.pack(pady=20)

        # Criar um segundo botão para carregar o arquivo CSV
        self.botao_carregar_csv = tk.Button(self.janela, text="Carregar CSV", command=self.carregar_csv)
        self.botao_carregar_csv.pack(pady=20)
        self.botao_carregar_csv.pack_forget()  # Ocultar o segundo botão inicialmente

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
                codigoVendedorLista = self.dataframe['Código'].apply(lambda x: f'{int(x):02}').tolist()
                meta1Lista = self.dataframe['Meta 1'].astype(int).tolist()
                meta2Lista = self.dataframe['Meta 2'].astype(int).tolist()
                meta3Lista = self.dataframe['Meta 3'].astype(int).tolist()
                meta4Lista = self.dataframe['Meta 4'].astype(int).tolist()

                # Exibir mensagem de sucesso
                mensagem_sucesso = "XLSX carregado com sucesso! Feche o programa que enviaremos os emails"
                messagebox.showinfo("Sucesso", mensagem_sucesso)

                # Mostrar o segundo botão após o carregamento do XLSX
                self.botao_carregar_csv.pack()

            except Exception as e:
                # Exibir mensagem de erro
                messagebox.showerror("Erro", f"Erro ao carregar o arquivo XLSX: {str(e)}")

    def carregar_csv(self):
        # Abrir a caixa de diálogo para seleção do arquivo CSV
        caminho_csv = filedialog.askopenfilename(filetypes=[("Arquivos CSV", "*.csv")])

        if caminho_csv:
            try:
                # Ler o arquivo CSV
                dataframe_csv = pd.read_csv(caminho_csv, encoding='latin-1', decimal=',', header=0, skiprows=1)

                # Realizar operações desejadas com o dataframe_csv, se necessário
                # ...

                # Exibir mensagem de sucesso
                mensagem_sucesso = "CSV carregado com sucesso!"
                messagebox.showinfo("Sucesso", mensagem_sucesso)

            except Exception as e:
                # Exibir mensagem de erro
                messagebox.showerror("Erro", f"Erro ao carregar o arquivo CSV: {str(e)}")
janela_principal = tk.Tk()

# Criar uma instância da classe Aplicacao
app = Aplicacao(janela_principal)

# Iniciar o loop principal
janela_principal.mainloop()