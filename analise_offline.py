import pandas as pd
import matplotlib.pyplot as plt

# Carregar o DataFrame a partir do arquivo CSV
caminho_arquivo_csv = 'analise.csv'
dataframe = pd.read_csv(caminho_arquivo_csv, encoding='latin-1', header=0, skiprows=1)
dataframe['Cliente'] = dataframe['Cliente'].str.replace(r'\D', '', regex=True)
dataframe['Vd'] = dataframe['Vd'].str.replace(r'\D', '', regex=True)

dataframe['Ano'] = pd.to_datetime(dataframe['Ano'], format='%Y')
dataframe_vd_5 = dataframe.loc[dataframe['Vd'] == "05"]

dataframe_2023 = dataframe.loc[dataframe['Ano'].dt.year == 2023]
dataframe_2024 = dataframe.loc[dataframe['Ano'].dt.year == 2024]

dataframe_2023['Janeiro'] = pd.to_numeric(dataframe_2023['Janeiro'], errors='coerce')
dataframe_2024['Janeiro'] = pd.to_numeric(dataframe_2024['Janeiro'], errors='coerce')

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
fig, ax = plt.subplots(figsize=(8, len(dataframe_merge) * 0.6 + 1))  # Ajuste da altura e da margem inferior
largura_barra = 0.3  # Ajuste da largura

# Posições das barras
posicoes = range(len(dataframe_merge))

# Plotar as barras
ax.barh(posicoes, dataframe_merge['Janeiro_2023'], height=largura_barra, label='2023', color='orange', edgecolor='black')
ax.barh(posicoes, dataframe_merge['Janeiro_2024'], height=largura_barra, label='2024', left=dataframe_merge['Janeiro_2023'], color='lightblue', edgecolor='black')

# Configurar o eixo y
ax.set_yticks(posicoes)
ax.set_yticklabels(dataframe_merge.index, fontsize=6)  # Ajuste o tamanho da fonte do índice

# Ajustar o espaço entre os rótulos das legendas
ax.legend(loc='upper right', bbox_to_anchor=(1.2, 1), ncol=1, fontsize=8)

# Rotacionar os rótulos do eixo y para melhor legibilidade
plt.yticks(rotation=0, ha='right')

# Adicionar rótulos e título
ax.set_xlabel('Valores de Janeiro', fontsize=8)
ax.set_ylabel('Clientes', fontsize=8)
ax.set_title('Comparação entre 2023 e 2024 - Janeiro', fontsize=10)

# Exibir o gráfico
plt.tight_layout()  # Ajuste automático de layout para evitar sobreposições
plt.show()

#