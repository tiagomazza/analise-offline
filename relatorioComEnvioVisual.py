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

        # Criar um botão para carregar o arquivo XLSX
        self.botao_carregar_xlsx = tk.Button(self.janela, text="Carregue os dados basicos da análise", command=self.carregar_xlsx)
        self.botao_carregar_xlsx.pack(pady=20)

        # Criar um segundo botão para carregar o arquivo CSV
        self.botao_carregar_csv = tk.Button(self.janela, text="Carregar a planilha com os dados a serem analisados", command=self.carregar_csv)
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

            except Exception as e:
                # Exibir mensagem de erro
                messagebox.showerror("Erro", f"Erro ao carregar o arquivo CSV: {str(e)}")
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




data_e_horario_atual = datetime.now()
data_e_horario_formatados = data_e_horario_atual.strftime("%d-%m-%Y %H:%M")


for salesman, emailVendedor, codigoVendedor, meta1, meta2, meta3, meta4 in zip (nomeLista, emailVendedorLista, codigoVendedorLista, meta1Lista, meta2Lista, meta3Lista, meta4Lista):
    #caminho_arquivo_csv = 'analise.csv'
    dataframe = pd.read_csv(caminho_csv, encoding='latin-1', decimal=',', header=0, skiprows=1)
    dataframe['Cliente'] = dataframe['Cliente'].str.replace(r'\D', '', regex=True)
    dataframe['Vd'] = dataframe['Vd'].str.replace(r'\D', '', regex=True)
    dataframe['Ano'] = pd.to_datetime(dataframe['Ano'], format='%Y')
    dataframe_vd = dataframe.loc[dataframe['Vd'] == codigoVendedor]
    dataframe_2023 = dataframe_vd.loc[dataframe['Ano'].dt.year == 2023]
    dataframe_2024 = dataframe_vd.loc[dataframe['Ano'].dt.year == 2024]
    dataframe_2023['Janeiro'].fillna(0, inplace=True)
    dataframe_2024['Janeiro'].fillna(0, inplace=True)
    # Filtrar clientes com valor maior que zero em 2023 ou 2024
    dataframe_2023 = dataframe_2023[dataframe_2023['Janeiro'] > 0]
    dataframe_2024 = dataframe_2024[dataframe_2024['Janeiro'] > 0]
    # Definir o índice como 'Nome'
    dataframe_2023.set_index('Nome', inplace=True)
    dataframe_2024.set_index('Nome', inplace=True)

    # Ordenar o DataFrame merge por 'Janeiro_2023'
    dataframe_merge = pd.merge(dataframe_2023, dataframe_2024, left_index=True, right_index=True, how='outer', suffixes=('_2023', '_2024'))
    dataframe_merge = dataframe_merge.sort_values(by='Janeiro_2023', ascending=True)  # Altere para 'ascending=False' se desejar ordem decrescente

    # Ajuste do tamanho do gráfico e espaçamento entre barras
    fig, ax = plt.subplots(figsize=(10, len(dataframe_merge) * 0.2 ))  # Ajuste da altura e da margem inferior
    largura_barra = 0.5  # Ajuste da largura

    # Posições das barras
    posicoes = range(len(dataframe_merge))

    ax.barh(posicoes, dataframe_merge['Janeiro_2023'], height=largura_barra, label='2023', color='red', edgecolor='none')
    ax.barh(posicoes, dataframe_merge['Janeiro_2024'], height=largura_barra, label='2024', color='blue', edgecolor='none', left=0)

    # Configurar o eixo y
    ax.set_yticks([pos + largura_barra / 2 for pos in posicoes])
    ax.set_yticklabels(dataframe_merge.index, fontsize=9, ha='right')  # Ajuste o tamanho da fonte do índice

    # Adicionar rótulos e título
    ax.set_xlabel('Valores de Janeiro', fontsize=8)
    ax.set_ylabel('Clientes', fontsize=12)
    ax.set_title('Comparação entre 2023 e 2024 - Janeiro', fontsize=14)


    for spine in ax.spines.values():
        spine.set_visible(False)

    # Salvar o primeiro gráfico como uma imagem temporária
    bar_chart_path = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
    plt.savefig(bar_chart_path, format='png', bbox_inches='tight')
    plt.close()

    #######GRAFICO CIRCULAR###########################

    vendasMes = dataframe_2024['Janeiro'].sum()

    metas = [meta1, meta2, meta3, meta4]

    fig, ax = plt.subplots(1, 4, figsize=(15, 4))

    for i, meta in enumerate(metas):
        porcentagem_vendas_mes = (vendasMes / meta) * 100
        porcentagem_meta = 100 - porcentagem_vendas_mes

        # Limitar a porcentagem a 100%
        porcentagem_vendas_mes = min(porcentagem_vendas_mes, 100)
        porcentagem_meta = 100 - porcentagem_vendas_mes

        # Cores para os gráficos
        cores = ['white', 'red']

        # Criar o gráfico de donut
        donut = ax[i].pie([porcentagem_meta, porcentagem_vendas_mes], startangle=90, colors=cores, wedgeprops=dict(width=0.3))

        # Adicionar um círculo branco no meio para criar o efeito de donut
        centro_do_circulo = plt.Circle((0, 0), 0.7, color='white')
        ax[i].add_patch(centro_do_circulo)

        # Adicionar o número no meio representando a porcentagem
        ax[i].text(0, 0, f'{porcentagem_vendas_mes:.1f}%', ha='center', va='center', fontsize=12, color='black')

        # Adicionar título com a variável e valor correspondentes
        ax[i].set_title(f'Meta: {meta}')

        # Se a porcentagem ultrapassar 100%, atualizar a variável meta e atualizar o gráfico
        if porcentagem_vendas_mes > 100:
            meta = metas[i+1]
            porcentagem_vendas_mes = (vendasMes / meta) * 100

            # Atualizar as cores para indicar a mudança
            cores = ['white', 'orange']

            # Limitar a porcentagem novamente a 100%
            porcentagem_vendas_mes = min(porcentagem_vendas_mes, 100)
            # Atualizar o gráfico com a nova meta
            donut[0][1].set_color(cores[1])
            # Adicionar título para a nova meta
            ax[i].set_title(f'Meta: {meta}')

    plt.suptitle(f'Porcentagem de Vendas em Relação à Meta. Já tem {vendasMes}€ vendidos ')
    plt.tight_layout(rect=[0, 0, 1, 0.95])

    # Salvar o segundo gráfico como uma imagem temporária
    donut_chart_path = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
    plt.savefig(donut_chart_path, format='png', bbox_inches='tight')
    plt.close()

    localDoArquivo = f'C:/Windows/Users/tiagomazza/Desktop/analise-offline/relatorio de vendas {salesman}.pdf'
    # Criar um arquivo PDF e inserir as imagens
    with open(localDoArquivo, 'wb') as pdf_file:
        pdf = canvas.Canvas(pdf_file, pagesize=A4)

        # Inserir imagem do primeiro gráfico
        pdf.drawInlineImage(donut_chart_path, 35, 600, width=550, height=150)
        #pdf.drawInlineImage(bar_chart_path, 20, 60, width=500, height=550)

        pdf.drawString(72, 780, f"Relatório de vendas do vendedor {salesman} - realizado as {data_e_horario_formatados}")
        pdf.drawImage(bar_chart_path, 20, 20, width=500, height=550)
        # Adicionar uma nova página para o segundo gráfico
        pdf.showPage()

        pdf.save()

    # Remover os arquivos temporários
    os.remove(bar_chart_path)
    os.remove(donut_chart_path)

    try:
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)

        email.to = emailVendedor
        email.Subject = 'Relatório de vendas'
        email.HtmlBody = f'Olá {salesman} segue anexo o relatório de vendas do mês até o presente momento.'
        
        attachment_path = localDoArquivo
        email.Attachments.Add(attachment_path)
        email.Send()
        print('E-mail enviado com sucesso!')
    except Exception as e:
        print(f'Erro: {e}')

