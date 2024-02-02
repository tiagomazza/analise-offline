import tkinter as tk
from tkinter import ttk
import pandas as pd

# Suponha que você já tenha um DataFrame chamado 'df'
# Aqui, estamos apenas criando um DataFrame de exemplo
data = {'Column1': [1, 2, 3], 'Column2': [4, 5, 6], 'Column3': [7, 8, 9]}
df = pd.DataFrame(data)

# Variável para armazenar a coluna selecionada
coluna_selecionada = None

def on_select(event):
    global coluna_selecionada
    selected_column = column_combobox.get()
    coluna_selecionada = selected_column
    print("Coluna selecionada:", coluna_selecionada)

# Criar a janela principal
root = tk.Tk()
root.title("Caixa Suspensa de Colunas")

# Obter as colunas do DataFrame
columns = df.columns.tolist()

# Criar a caixa suspensa (Combobox)
column_combobox = ttk.Combobox(root, values=columns)
column_combobox.set("Selecione uma coluna")  # Texto padrão exibido na caixa suspensa
column_combobox.bind("<<ComboboxSelected>>", on_select)

# Exibir a caixa suspensa
column_combobox.pack(pady=10)

# Iniciar o loop principal do Tkinter
root.mainloop()

# Agora você pode usar a variável 'coluna_selecionada' para acessar a coluna escolhida em outros lugares do seu código.
print (coluna_selecionada)
