import win32com.client as win32

try:
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    email.to = 'tiagomazza@gmail.com'
    email.Subject = 'Teste envio python'
    email.HtmlBody = '''
    teste de envio novo codigo
    '''
    attachment_path = 'C:/caminho/do/seu/anexo.txt'  
    email.Attachments.Add(attachment_path)
    email.Send()
    print('E-mail enviado com sucesso!')
except Exception as e:
    print(f'Erro: {e}')


    