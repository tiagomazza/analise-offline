import customtkinter

janela = customtkinter.CTk()
janela.geometry ("200x300")

def clique():
    print('fazer login')

texto = customtkinter.CTkLabel(janela, text='Fazer login')
texto.pack(padx =10, pady=10)

email = customtkinter.CTkEntry(janela, placeholder_text='Coloque seu email')
email.pack(padx =10, pady=10)

senha = customtkinter.CTkEntry(janela,placeholder_text='Sua senha', show= "*" )
senha.pack(padx =10, pady=10)

botao = customtkinter.CTkButton(janela, text='Aceder')
botao.pack(padx =10, pady=10)

janela.mainloop()
