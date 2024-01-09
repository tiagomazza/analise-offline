import pandas as pd
import matplotlib.pyplot as plt
import tempfile
from datetime import datetime
import win32com.client as win32
from matplotlib.backends.backend_pdf import PdfPages

data_e_horario_atual = datetime.now()
data_e_horario_formatados = data_e_horario_atual.strftime("%d-%m-%Y %H:%M")


salesman = "06"
# Carregar o DataFrame a partir do arquivo CSV
caminho_arquivo_csv = 'analise.csv'
dataframe = pd.read_csv(caminho_arquivo_csv, encoding='latin-1', decimal=',', header=0, skiprows=1)
dataframe['Cliente'] = dataframe['Cliente'].str.replace(r'\D', '', regex=True)
dataframe['Vd'] = dataframe['Vd'].str.replace(r'\D', '', regex=True)

dataframe['Ano'] = pd.to_datetime(dataframe['Ano'], format='%Y')
dataframe_vd = dataframe.loc[dataframe['Vd'] == salesman]

dataframe_2023 = dataframe_vd.loc[dataframe['Ano'].dt.year == 2023]
dataframe_2024 = dataframe_vd.loc[dataframe['Ano'].dt.year == 2024]

# Preencher valores nulos com 0
dataframe_2023['Janeiro'].fillna(0, inplace=True)
dataframe_2024['Janeiro'].fillna(0, inplace=True)

# Filtrar clientes com valor maior que zero em 2023 ou 2024
dataframe_2023 = dataframe_2023[dataframe_2023['Janeiro'] > 0]
dataframe_2024 = dataframe_2024[dataframe_2024['Janeiro'] > 0]

# Ordenar o DataFrame por 'Janeiro' de 2023 em ordem decrescente
dataframe_2023 = dataframe_2023.sort_values(by='Janeiro', ascending=False)

# Definir o índice como 'Nome'
dataframe_2023.set_index('Nome', inplace=True)
dataframe_2024.set_index('Nome', inplace=True)

# Ordenar o DataFrame merge por 'Janeiro_2023'
dataframe_merge = pd.merge(dataframe_2023, dataframe_2024, left_index=True, right_index=True, how='outer', suffixes=('_2023', '_2024'))
dataframe_merge = dataframe_merge.sort_values(by='Janeiro_2023', ascending=True)  # Altere para 'ascending=False' se desejar ordem decrescente

# Ajuste do tamanho do gráfico e espaçamento entre barras
fig, ax = plt.subplots(figsize=(8, len(dataframe_merge) * 0.2 + 1))  # Ajuste da altura e da margem inferior
largura_barra = 0.3  # Ajuste da largura

# Posições das barras
posicoes = range(len(dataframe_merge))

# Plotar as barras de 2023 transparentes com borda preta
ax.barh(posicoes, dataframe_merge['Janeiro_2023'], height=largura_barra, label='2023', color='red', edgecolor='none')

# Plotar as barras de 2024 cobrindo totalmente as de 2023
ax.barh(posicoes, dataframe_merge['Janeiro_2024'], height=largura_barra, label='2024', color='blue', edgecolor='none', left=0)

# Configurar o eixo y
ax.set_yticks([pos + largura_barra / 2 for pos in posicoes])
ax.set_yticklabels(dataframe_merge.index, fontsize=6, ha='right')  # Ajuste o tamanho da fonte do índice

# Ajustar o espaço entre os rótulos das legendas
ax.legend(loc='upper right', bbox_to_anchor=(1.2, 1), ncol=1, fontsize=8)

# Adicionar rótulos e título
ax.set_xlabel('Valores de Janeiro', fontsize=8)
ax.set_ylabel('Clientes', fontsize=8)
ax.set_title('Comparação entre 2023 e 2024 - Janeiro', fontsize=10)

for spine in ax.spines.values():
    spine.set_visible(False)

# Salvar o primeiro gráfico como uma imagem temporária
bar_chart_path = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
plt.savefig(bar_chart_path, format='png', bbox_inches='tight')
plt.close()

#######GRAFICO CIRCULAR###########################

vendasMes = dataframe_2024['Janeiro'].sum()

metas = [2000, 3000, 4000, 5000]



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

plt.suptitle('Porcentagem de Vendas em Relação à Meta')
plt.tight_layout(rect=[0, 0, 1, 0.95])

# Salvar o segundo gráfico como uma imagem temporária
donut_chart_path = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
plt.savefig(donut_chart_path, format='png', bbox_inches='tight')
plt.close()

# Criar um arquivo PDF temporário para os gráficos
pdf_path = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name

with PdfPages(pdf_path) as pdf:
    # Gráfico 1
    fig, ax = plt.subplots(figsize=(8, len(dataframe_merge) * 0.2 + 1))
    # Adicionar código para plotar o gráfico 1
    pdf.savefig()
    plt.close()

    # Gráfico 2
    fig, ax = plt.subplots(1, 4, figsize=(15, 4))
    # Adicionar código para plotar o gráfico 2
    pdf.savefig()
    plt.close()

# Criar um arquivo PDF para o relatório de vendas
report_pdf_path = 'relatorio_de_vendas.pdf'
with PdfPages(report_pdf_path) as pdf:
    # Gráfico 1
    fig, ax = plt.subplots(figsize=(8, len(dataframe_merge) * 0.2 + 1))
    # Adicionar código para plotar o gráfico 1
    pdf.savefig()
    plt.close()

    # Gráfico 2
    fig, ax = plt.subplots(1, 4, figsize=(15, 4))
    # Adicionar código para plotar o gráfico 2
    pdf.savefig()
    plt.close()

print("Relatório de vendas salvo como 'relatorio_de_vendas.pdf'")

# Enviar e-mail com o relatório anexado
outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)

email.to = 'tiagomazza@gmail.com'
email.Subject = 'Relatório de Vendas'
email.Body = 'Por favor, encontre anexado o relatório de vendas.'
email.Attachments.Add(report_pdf_path)

# Enviar e-mail
email.Send()
