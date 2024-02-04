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
import re

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
nome_ano_base = None
nome_ano_atual = None


data_e_horario_atual = datetime.now()
data_e_horario_formatados = data_e_horario_atual.strftime("%d-%m-%Y %H:%M")

class Aplicacao:
    def __init__(self, janela):
        self.janela = janela
        self.janela.title("METAS")
        self.janela.geometry("270x350")

        self.dataframe = None 
        self.meses_variaveis = {}

        self.botao_carregar_xlsx = tk.Button(self.janela, text="Carregar Parâmetros", command=self.carregar_xlsx)
        self.botao_carregar_xlsx.pack(pady=10)

        self.botao_carregar_csv = tk.Button(self.janela, text="Carregar dados a serem analisados", command=self.carregar_csv)
        self.botao_carregar_csv.pack(pady=10)


        self.botao_on_select = tk.Button(janela_principal, text="Escolher local onde salvar os ficheiros", command=self.save_local)
        self.botao_on_select.pack(pady=10)

        self.column_combobox = None 
        self.column_combobox_ano_base = None 
        self.column_combobox_ano_atual = None 
    
        self.botao_fechar = None    
        self.check_var = None
                

    def carregar_xlsx(self):
        caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xls")])
        
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

                self.botao_carregar_xlsx.config(fg="green")
            
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
                df_menu_ano = pd.read_csv(caminho_csv, encoding='latin-1', decimal=',', header=0, skiprows=1)
                columns = df_menu.columns.tolist()

                self.column_combobox_ano_base = ttk.Combobox(janela_principal, values=df_menu_ano['Ano'].unique())
                self.column_combobox_ano_base.set("Selecione o ano base")  
                self.column_combobox_ano_base.bind("<<ComboboxSelected>>", self.on_ano_base)
                self.column_combobox_ano_base.pack(pady=2)

                self.column_combobox_ano_atual = ttk.Combobox(janela_principal, values=df_menu_ano['Ano'].unique())
                self.column_combobox_ano_atual.set("Selecione o ano atual")  
                self.column_combobox_ano_atual.bind("<<ComboboxSelected>>", self.on_ano_atual)
                self.column_combobox_ano_atual.pack(pady=2)


                self.column_combobox = ttk.Combobox(janela_principal, values=columns)
                self.column_combobox.set("Selecione um mês")  
                self.column_combobox.bind("<<ComboboxSelected>>", self.on_select)
                self.column_combobox.pack(pady=10)

                mensagem_sucesso = "dados em CSV carregados!"
                self.botao_carregar_csv.config(fg="green")  

                self.botao_fechar = tk.Button(self.janela, text="Gerar relatórios", command=self.janela.destroy)

                self.botao_fechar.config(fg="green", font=("Helvetica", 15, "bold"), bd=4)
                self.botao_fechar.pack(pady=10)
   
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar o arquivo CSV: {str(e)}")
    
    def on_select(self, event):
            global coluna_selecionada
            selected_column = self.column_combobox.get()
            coluna_selecionada = selected_column

    def on_ano_base (self, event):
        global nome_ano_base
        ano_base_escolhido = self.column_combobox_ano_base.get()
        nome_ano_base = ano_base_escolhido
        nome_ano_base = re.sub(r'[^0-9]', '', nome_ano_base)
        nome_ano_base = int(nome_ano_base)

    def on_ano_atual (self, event):
        global nome_ano_atual
        ano_atual_escolhido = self.column_combobox_ano_atual.get()
        nome_ano_atual = ano_atual_escolhido
        nome_ano_atual = re.sub(r'[^0-9]', '', nome_ano_atual)
        nome_ano_atual = int(nome_ano_atual)

    def save_local(self):
        global local
        local = filedialog.askdirectory()
        self.botao_on_select.config(fg="green")

    
    def enviar_email():
    
        try:
            outlook = win32.client.Dispatch('outlook.application')
            email = outlook.CreateItem(0)

            email.to = emailVendedor
            email.Subject = 'Relatório de vendas'
            email.HtmlBody = f'Olá {salesman} segue anexo o relatório de vendas do mês até o presente momento.'
            
            attachment_path = localDoArquivo
            email.Attachments.Add(attachment_path)
            email.Send()
            messagebox.showinfo("Sucesso", "E-mail enviado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro: {e}")


janela_principal = tk.Tk()
icone = 'icone.ico'
janela_principal.iconbitmap(icone)
janela_principal.configure(bg="#dcdcdc")
app = Aplicacao(janela_principal)

janela_principal.mainloop()


for salesman, emailVendedor, codigoVendedor, meta1, meta2, meta3, meta4, valorBonus1, valorBonus2, valorBonus3, valorBonus4 in zip (nomeLista, emailVendedorLista, codigoVendedorLista, meta1Lista, meta2Lista, meta3Lista, meta4Lista, bonus1Lista, bonus2Lista, bonus3Lista, bonus4Lista):

    dataframe = pd.read_csv(caminho_csv, encoding='latin-1', decimal=',', header=0, skiprows=1)
    dataframe['Cliente'] = dataframe['Cliente'].str.replace(r'\D', '', regex=True)
    dataframe['Vd'] = dataframe['Vd'].str.replace(r'\D', '', regex=True)
    dataframe['Ano'] = pd.to_datetime(dataframe['Ano'], format='%Y')
    dataframe_vd = dataframe.loc[dataframe['Vd'] == codigoVendedor]
    dataframe_ano_anterior = dataframe_vd.loc[dataframe['Ano'].dt.year == nome_ano_base]
    dataframe_ano_atual = dataframe_vd.loc[dataframe['Ano'].dt.year == nome_ano_atual]
    dataframe_ano_anterior[coluna_selecionada].fillna(0, inplace=True)
    dataframe_ano_atual[coluna_selecionada].fillna(0, inplace=True)
    dataframe_ano_anterior = dataframe_ano_anterior[dataframe_ano_anterior[coluna_selecionada] != 0]
    dataframe_ano_atual = dataframe_ano_atual[dataframe_ano_atual[coluna_selecionada] > 0]
    dataframe_ano_anterior.set_index('Nome', inplace=True)
    dataframe_ano_atual.set_index('Nome', inplace=True)
    dataframe_merge = pd.merge(dataframe_ano_anterior, dataframe_ano_atual, left_index=True, right_index=True, how='outer', suffixes=('_ano_anterior', '_ano_atual'))
    dataframe_merge = dataframe_merge.sort_values(by=f'{coluna_selecionada}_ano_anterior', ascending=True)  

    fig, ax = plt.subplots(figsize=(10, len(dataframe_merge) * 0.3)) 
    largura_barra = 0.5 

    posicoes = range(len(dataframe_merge))

    ax.barh(posicoes, dataframe_merge[f'{coluna_selecionada}_ano_anterior'], height=largura_barra, label=nome_ano_base, color='red', edgecolor='none')
    ax.barh(posicoes, dataframe_merge[f'{coluna_selecionada}_ano_atual'], height=largura_barra, label=nome_ano_atual, color='blue', edgecolor='none')

    for pos, valor_ano_anterior, valor_ano_atual in zip(posicoes, dataframe_merge[f'{coluna_selecionada}_ano_anterior'], dataframe_merge[f'{coluna_selecionada}_ano_atual']):
        ax.text(max(valor_ano_anterior, 0.5), pos + largura_barra/2, f'{valor_ano_anterior}€', ha='left', va='center', fontsize=8, color='red')

    ax.set_yticks([pos + largura_barra / 2 for pos in posicoes])
    ax.set_yticklabels(dataframe_merge.index, fontsize=9, ha='right')  

    for pos, valor_ano_anterior, valor_ano_atual in zip(posicoes, dataframe_merge[f'{coluna_selecionada}_ano_anterior'], dataframe_merge[f'{coluna_selecionada}_ano_atual']):
        ax.text(valor_ano_anterior, pos + largura_barra/2, f'{valor_ano_anterior}€', ha='left', va='center', fontsize=8, color='red')
        
    for spine in ax.spines.values():
        spine.set_visible(False)

    bar_chart_path = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
    plt.savefig(bar_chart_path, format='png', bbox_inches='tight')
    plt.close()

    vendasMes = dataframe_ano_atual[coluna_selecionada].sum()
    metas = [meta1, meta2, meta3]
    metas_formatadas = list(map(lambda x: '{:,.2f}'.format(x).replace(',', '.'), metas))
    bonus = [valorBonus1, valorBonus2, valorBonus3]
    bonus_formatados = list(map(lambda x: '{:,.2f}'.format(x).replace(',', '.'), bonus))
    bonus = bonus_formatados
    fig, ax = plt.subplots(1, 3, figsize=(15, 4))

    for i, (meta, bonu, meta_format) in enumerate(zip(metas,bonus,metas_formatadas)):
        porcentagem_vendas_mes = (vendasMes / meta) * 100
        porcentagem_meta = 100 - porcentagem_vendas_mes

        porcentagem_vendas_mes = min(porcentagem_vendas_mes, 100)
        porcentagem_meta = 100 - porcentagem_vendas_mes

        cores = ['lightgrey', 'blue']

        donut = ax[i].pie([porcentagem_meta, porcentagem_vendas_mes], startangle=90, colors=cores, wedgeprops=dict(width=0.3))
        centro_do_circulo = plt.Circle((0, 0), 0.7, color='white')
        ax[i].add_patch(centro_do_circulo)
        ax[i].text(0, 0, f'{porcentagem_vendas_mes:.1f}%', ha='center', va='center', fontsize=30, color='black', fontstyle ='normal')
        ax[i].set_title(f'Meta: {meta_format}€ \n Bonus:{bonu}€', color='dimgrey')

        if porcentagem_vendas_mes > 100:
            meta = metas[i+1]
            porcentagem_vendas_mes = (vendasMes / meta) * 100
            cores = ['white', 'orange']
            porcentagem_vendas_mes = min(porcentagem_vendas_mes, 100)
            donut[0][1].set_color(cores[1])
            ax[i].set_title(f'Meta: {meta}')

    vendasMes = round(vendasMes,2)
    vendasMes_format = '{:,}'.format(vendasMes)

    plt.suptitle(f'Porcentagem de Vendas em Relação à Meta. Já tem {vendasMes_format}€ vendidos ')
    plt.tight_layout(rect=[0, 0, 1, 0.95])

    donut_chart_path = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
    plt.savefig(donut_chart_path, format='png', bbox_inches='tight')
    plt.close()

    localDoArquivo = f'{local}/relatorio de vendas {salesman}.pdf'
    with open(localDoArquivo, 'wb') as pdf_file:
        pdf = canvas.Canvas(pdf_file, pagesize=A4)
        pdf.drawInlineImage(donut_chart_path, 35, 575, width=500, height=150)
        pdf.setFont("Helvetica-Bold", 30)
        pdf.drawString(40, 780, f"Relatório de vendas")
        pdf.setFont("Helvetica-Oblique", 12)
        pdf.drawString(40, 760, f"vendedor {salesman} - realizado as {data_e_horario_formatados}")
        pdf.drawImage(bar_chart_path, 10, 30, width=570, height=500)
        img = ImageReader("logo.jpg")
        pdf.drawImage(img, 460, 750, width=75, height=60)
        pdf.setFont("Helvetica", 20)
        pdf.drawString(40, 530, f"Comparativo de vendas")
        pdf.showPage()
        pdf.save()

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



