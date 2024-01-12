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
        self.janela.title("METAS")

        self.dataframe = None  # Variável para armazenar o DataFrame
        self.caminho_salvar_resultados = None  # Variável para armazenar o caminho escolhido pelo usuário
        self.nome_arquivo = "teste.txt"  # Nome do arquivo a ser adicionado ao caminho

        # Criar um botão para carregar o arquivo XLSX
        self.botao_carregar_xlsx = tk.Button(self.janela, text="Carregue os dados básicos da análise", command=self.carregar_xlsx)
        self.botao_carregar_xlsx.pack(pady=20)

        # Criar um segundo botão para carregar o arquivo CSV
        self.botao_carregar_csv = tk.Button(self.janela, text="Carregar a planilha com os dados a serem analisados", command=self.carregar_csv)
        self.botao_carregar_csv.pack(pady=20)
        self.botao_carregar_csv.pack_forget()  # Ocultar o segundo botão inicialmente

        # Criar um terceiro botão para selecionar a pasta de destino
        self.botao_selecionar_pasta = tk.Button(self.janela, text="Selecionar Pasta de Destino", command=self.selecionar_pasta_destino)
        self.botao_selecionar_pasta.pack(pady=20)
        self.botao_selecionar_pasta.pack_forget()  # Ocultar o terceiro botão inicialmente

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
        global caminho_csv
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

                # Mostrar o terceiro botão após o carregamento do CSV
                self.botao_selecionar_pasta.pack()

            except Exception as e:
                # Exibir mensagem de erro
                messagebox.showerror("Erro", f"Erro ao carregar o arquivo CSV: {str(e)}")

    def selecionar_pasta_destino(self):
        # Abrir a caixa de diálogo para seleção da pasta de destino
        pasta_destino = filedialog.askdirectory()

        if pasta_destino:
            try:
                # Concatenar o caminho escolhido pelo usuário com o nome do arquivo
                self.caminho_salvar_resultados = os.path.join(pasta_destino, self.nome_arquivo)

                # Exibir mensagem de sucesso
                mensagem_sucesso = "Pasta de destino selecionada com sucesso!"
                messagebox.showinfo("Sucesso", mensagem_sucesso)

            except Exception as e:
                # Exibir mensagem de erro
                messagebox.showerror("Erro", f"Erro ao selecionar a pasta de destino: {str(e)}")

janela_principal = tk.Tk()

# Criar uma instância da classe Aplicacao
app = Aplicacao(janela_principal)

# Iniciar o loop principal
janela_principal.mainloop()

print(app.caminho_salvar_resultados)  # Exibir o caminho escolhido pelo usuário após o encerramento da interface gráfica
