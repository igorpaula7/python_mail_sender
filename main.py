import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

email = outlook.CreateItem(0)

name = input("Qual o nome da pessoa?")
mail = input("Qual o e-mail da pessoa?")
subject = input("Qual o assunto do e-mail?")
message = input("Qual o texto do e-mail?")

# configurar as informações do seu e-mail
email.To = mail
email.Subject = subject
email.HTMLBody = f"""
Oi, {name}.

{message}
"""

anexo = "C://Users/user/Documentos/arquivo.txt"
email.Attachments.Add(anexo)

email.Send()
print("Email Enviado")
