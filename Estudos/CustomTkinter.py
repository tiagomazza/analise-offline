import customtkinter

janela = customtkinter.CTk()
janela.title('Titulo da janela')
janela.geometry ("200x300")

def clique():
    print('fazer login')

texto = customtkinter.CTkLabel(janela, text='Fazer login')
texto.pack(padx =10, pady=10)

email = customtkinter.CTkEntry(janela, placeholder_text='Coloque seu email')
email.pack(padx =10, pady=10)

senha = customtkinter.CTkEntry(janela,placeholder_text='Sua senha', show= "*" )
senha.pack(padx =10, pady=10)

combobox = customtkinter.CTkComboBox(janela, values=["option 1", "option 2"])
combobox.pack(padx =10, pady=10)

spinbox = customtkinter.CTkSwitch(janela, text='enviar email')
spinbox.pack(padx =10, pady=10)

botao = customtkinter.CTkButton(janela, text='Aceder')
botao.pack(padx =10, pady=10)

imagem =customtkinter.CTkCanvas (janela)
imagem.pack(padx =10, pady=10)

janela.mainloop()
