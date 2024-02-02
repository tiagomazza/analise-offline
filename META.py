import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import matplotlib.pyplot as plt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
import tempfile
import os
from datetime import datetime
import win32com.client as win32

nomeLista = []
emailVendedorLista = []
codigoVendedorLista = []
meta1Lista = []
meta2Lista = []
meta3Lista = []
meta4Lista = []
bonus1Lista = []
bonus2Lista = []
bonus3Lista = []
bonus4Lista = []
coluna_selecionada = None
data_e_horario_atual = datetime.now()
data_e_horario_formatados = data_e_horario_atual.strftime("%d-%m-%Y %H:%M")

class Aplicacao:
    def __init__(self, janela):
        self.janela = janela
        self.janela.title("METAS")
        self.janela.geometry("270x250")

        self.dataframe = None  # Variável para armazenar o DataFrame
        self.meses_variaveis = {}

        self.botao_carregar_xlsx = tk.Button(self.janela, text="Carregar Parâmetros", command=self.carregar_xlsx)
        self.botao_carregar_xlsx.pack(pady=20)

        self.botao_carregar_csv = tk.Button(self.janela, text="Carregar dados a serem analisados", command=self.carregar_csv)
        self.botao_carregar_csv.pack(pady=20)
        #self.botao_carregar_csv.pack_forget()  # Ocultar o segundo botão inicialmente

        self.botao_on_select = tk.Button(janela_principal, text="Escolher local onde salvar os ficheiros", command=self.save_local)
        self.botao_on_select.pack(pady=10)

        #self.gerar_relatorios = tk.Button(janela_principal, text="Gerar relatórios", command=self.save_local)
        self.botao_on_select.pack(pady=30)

        self.column_combobox = None  # Adicionado atributo column_combobox

        self.rotulo = tk.Label(self.janela, text="Escolha o mês de análise")
        



    def carregar_xlsx(self):
        # Abrir a caixa de diálogo para seleção do arquivo XLSX
        caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
        
        if caminho_arquivo:
            try:
                self.dataframe = pd.read_excel(caminho_arquivo)

                global nomeLista, emailVendedorLista, codigoVendedorLista, meta1Lista, meta2Lista, meta3Lista, meta4Lista, bonus1Lista, bonus2Lista, bonus3Lista, bonus4Lista
                nomeLista = self.dataframe['Nome'].astype(str).tolist()
                emailVendedorLista = self.dataframe['Email'].astype(str).tolist()
                codigoVendedorLista = self.dataframe['Código'].astype(str).tolist()
                codigoVendedorLista = self.dataframe['Código'].apply(lambda x: f'{int(x):02}').tolist()
                meta1Lista = self.dataframe['Meta 1'].astype(int).tolist()
                meta2Lista = self.dataframe['Meta 2'].astype(int).tolist()
                meta3Lista = self.dataframe['Meta 3'].astype(int).tolist()
                meta4Lista = self.dataframe['Meta 4'].astype(int).tolist()
                bonus1Lista = self.dataframe['Bonus da Meta 1'].astype(int).tolist()
                bonus2Lista = self.dataframe['Bonus da Meta 2'].astype(int).tolist()
                bonus3Lista = self.dataframe['Bonus da Meta 3'].astype(int).tolist()
                bonus4Lista = self.dataframe['Bonus da Meta 4'].astype(int).tolist()

                mensagem_sucesso = "Paramentros XLSX carregado com sucesso! Hora de carregar os dados a serem analisádos"
                messagebox.showinfo("Sucesso", mensagem_sucesso)
                self.botao_carregar_xlsx.config(fg="green")
                self.botao_carregar_csv.pack()

            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar o arquivo XLSX: {str(e)}")

    def carregar_csv(self):
        global caminho_csv
        caminho_csv = filedialog.askopenfilename(filetypes=[("Arquivos CSV", "*.csv")])

        if caminho_csv:
            try:
                dataframe_csv = pd.read_csv(caminho_csv, encoding='latin-1', decimal=',', header=0, skiprows=1)
                df_menu = pd.read_csv(caminho_csv, encoding='latin-1', decimal=',', header=0, skiprows=1)
                df_menu = df_menu.iloc[:, 5:]
                columns = df_menu.columns.tolist()

                self.column_combobox = ttk.Combobox(janela_principal, values=columns)
                self.column_combobox.set("Selecione um mês")  
                self.column_combobox.bind("<<ComboboxSelected>>", self.on_select)
                self.column_combobox.pack(pady=10)

                mensagem_sucesso = "dados em CSV carregados! Feche o programa para gerar os relatórios"
                self.botao_carregar_csv.config(fg="green")
                messagebox.showinfo("Sucesso", mensagem_sucesso)
                
   

            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar o arquivo CSV: {str(e)}")
    
    def on_select(self, event):
            global coluna_selecionada
            selected_column = self.column_combobox.get()
            coluna_selecionada = selected_column
            print("Coluna selecionada:", coluna_selecionada)    

    def save_local(self):
        global local
        local = filedialog.askdirectory()
        self.botao_on_select.config(fg="green")

janela_principal = tk.Tk()
janela_principal.iconbitmap('icone.ico')
app = Aplicacao(janela_principal)





janela_principal.mainloop()


for salesman, emailVendedor, codigoVendedor, meta1, meta2, meta3, meta4, valorBonus1, valorBonus2, valorBonus3, valorBonus4 in zip (nomeLista, emailVendedorLista, codigoVendedorLista, meta1Lista, meta2Lista, meta3Lista, meta4Lista, bonus1Lista, bonus2Lista, bonus3Lista, bonus4Lista):
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

    
    fig, ax = plt.subplots(figsize=(10, len(dataframe_merge) * 0.3))  # Ajuste da altura e da margem inferior
    largura_barra = 0.5  # Ajuste da largura

    # Posições das barras
    posicoes = range(len(dataframe_merge))

    ax.barh(posicoes, dataframe_merge['Janeiro_2023'], height=largura_barra, label='2023', color='red', edgecolor='none')
    ax.barh(posicoes, dataframe_merge['Janeiro_2024'], height=largura_barra, label='2024', color='blue', edgecolor='none')

    # Adicionar valores após as barras no gráfico de barras
    for pos, valor_2023, valor_2024 in zip(posicoes, dataframe_merge['Janeiro_2023'], dataframe_merge['Janeiro_2024']):
        ax.text(max(valor_2023, 0.5), pos + largura_barra/2, f'{valor_2023}€', ha='left', va='center', fontsize=8, color='red')

    # Configurar o eixo y
    ax.set_yticks([pos + largura_barra / 2 for pos in posicoes])
    ax.set_yticklabels(dataframe_merge.index, fontsize=9, ha='right')  # Ajuste o tamanho da fonte do índice

    # Adicionar rótulos e título
    #ax.set_xlabel('Valores de Janeiro', fontsize=8)
    #ax.set_ylabel('Clientes', fontsize=12)
    #ax.set_title('Comparação entre 2023 e 2024 - Janeiro', fontsize=14)

    for pos, valor_2023, valor_2024 in zip(posicoes, dataframe_merge['Janeiro_2023'], dataframe_merge['Janeiro_2024']):
        ax.text(valor_2023, pos + largura_barra/2, f'{valor_2023}€', ha='left', va='center', fontsize=8, color='red')
        
    for spine in ax.spines.values():
        spine.set_visible(False)

    # Salvar o primeiro gráfico como uma imagem temporária
    bar_chart_path = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
    plt.savefig(bar_chart_path, format='png', bbox_inches='tight')
    plt.close()

    #######GRAFICO CIRCULAR###########################

    vendasMes = dataframe_2024['Janeiro'].sum()
    metas = [meta1, meta2, meta3]
    metas_formatadas = list(map(lambda x: '{:,.2f}'.format(x).replace(',', '.'), metas))
    bonus = [valorBonus1, valorBonus2, valorBonus3]
    bonus_formatados = list(map(lambda x: '{:,.2f}'.format(x).replace(',', '.'), bonus))
    bonus = bonus_formatados

    fig, ax = plt.subplots(1, 3, figsize=(15, 4))

    for i, (meta, bonu, meta_format) in enumerate(zip(metas,bonus,metas_formatadas)):
        porcentagem_vendas_mes = (vendasMes / meta) * 100
        porcentagem_meta = 100 - porcentagem_vendas_mes

        # Limitar a porcentagem a 100%
        porcentagem_vendas_mes = min(porcentagem_vendas_mes, 100)
        porcentagem_meta = 100 - porcentagem_vendas_mes

        # Cores para os gráficos
        cores = ['lightgrey', 'blue']

        # Criar o gráfico de donut
        donut = ax[i].pie([porcentagem_meta, porcentagem_vendas_mes], startangle=90, colors=cores, wedgeprops=dict(width=0.3))

        # Adicionar um círculo branco no meio para criar o efeito de donut
        centro_do_circulo = plt.Circle((0, 0), 0.7, color='white')
        ax[i].add_patch(centro_do_circulo)

        # Adicionar o número no meio representando a porcentagem
        ax[i].text(0, 0, f'{porcentagem_vendas_mes:.1f}%', ha='center', va='center', fontsize=30, color='black', fontstyle ='normal')

        # Adicionar título com a variável e valor correspondentes
        ax[i].set_title(f'Meta: {meta_format}€ \n Bonus:{bonu}€', color='dimgrey')

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

    vendasMes = round(vendasMes,2)
    vendasMes_format = '{:,}'.format(vendasMes)

    plt.suptitle(f'Porcentagem de Vendas em Relação à Meta. Já tem {vendasMes_format}€ vendidos ')
    plt.tight_layout(rect=[0, 0, 1, 0.95])

    # Salvar o segundo gráfico como uma imagem temporária
    donut_chart_path = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
    plt.savefig(donut_chart_path, format='png', bbox_inches='tight')
    plt.close()

    #C:/Users/tiagomazza/Desktop/analise-offline
    localDoArquivo = f'{local}/relatorio de vendas {salesman}.pdf'
    #localDoArquivo = f'C:/Users/tiagomazza/Desktop/analise-offline/relatorio de vendas {salesman}.pdf'
    # Criar um arquivo PDF e inserir as imagens
    with open(localDoArquivo, 'wb') as pdf_file:
        pdf = canvas.Canvas(pdf_file, pagesize=A4)
       
        # Inserir imagem do primeiro gráfico
        pdf.drawInlineImage(donut_chart_path, 35, 575, width=500, height=150)
        #pdf.drawInlineImage(bar_chart_path, 20, 60, width=500, height=550)

        pdf.setFont("Helvetica-Bold", 30)
        pdf.drawString(40, 780, f"Relatório de vendas")
        pdf.setFont("Helvetica-Oblique", 12)
        pdf.drawString(40, 760, f"vendedor {salesman} - realizado as {data_e_horario_formatados}")
        pdf.drawImage(bar_chart_path, 10, 30, width=570, height=500)
        # Adicionar uma nova página para o segundo gráfico
        img = ImageReader("logo.jpg")
        pdf.drawImage(img, 460, 750, width=75, height=60)
        pdf.setFont("Helvetica", 20)
        pdf.drawString(40, 530, f"Comparativo de vendas")
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
        #email.Send()
        print('E-mail enviado com sucesso!')
    except Exception as e:
        print(f'Erro: {e}')

