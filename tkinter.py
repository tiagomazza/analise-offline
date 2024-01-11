import tkinter as tk
from tkinter import messagebox

def mostrar_mensagem():
    messagebox.showinfo("Saudação", "Olá, Mundo!")

# Criar a janela principal
janela_principal = tk.Tk()
janela_principal.title("Exemplo com Tkinter")

# Criar um botão na janela
botao_saudacao = tk.Button(janela_principal, text="Clique para Saudação", command=mostrar_mensagem)
botao_saudacao.pack(pady=20)

# Iniciar o loop principal
janela_principal.mainloop()
