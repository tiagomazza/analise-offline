import tkinter as tk
from tkinter import ttk
import pandas as pd


data = {'Column1': [1, 2, 3], 'Column2': [4, 5, 6], 'Column3': [7, 8, 9]}
df = pd.DataFrame(data)

coluna_selecionada = None

def on_select(event):
    global coluna_selecionada
    selected_column = column_combobox.get()
    coluna_selecionada = selected_column
    print("Coluna selecionada:", coluna_selecionada)

janela = tk.Tk()
janela.title("Titulo da janela")
janela.iconbitmap('icone.ico')
janela.configure(bg="#add8e6")
janela.geometry("200x400")


columns = df.columns.tolist()

# Criar a caixa suspensa (Combobox)
column_combobox = ttk.Combobox(janela, values=columns)
column_combobox.set("Selecione uma coluna")  
column_combobox.bind("<<ComboboxSelected>>", on_select)
column_combobox.pack(pady=10)

# Criar a checkbox
check_var = tk.BooleanVar()
checkbox = tk.Checkbutton(janela, text="Texto do checkbox", variable=check_var)
checkbox.pack(pady=10)

# Criar botão
def exibir_mensagem():
    print("Pressionado")

botao = tk.Button(janela, text="Pressione-me", command=exibir_mensagem)
botao.pack(pady=10)

spin_valor = tk.Spinbox(janela, from_=0, to=100, increment=1)
spin_valor.pack(pady=10)

        


janela.mainloop()

print (coluna_selecionada)
print (check_var)
print(f"Valor do Spinbox: {valor}")
