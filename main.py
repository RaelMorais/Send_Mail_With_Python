import win32com.client as win32#importando a biblioteca de python
import pandas as pd
import datetime 

outlook = win32.Dispatch('outlook.application')#criando integração com outlook
excel_arquivo = r'#'#r tira necessidade de barras duplas
dataframe = pd.read_excel(excel_arquivo)
x = datetime.datetime.today()
data_atual = x .strftime('%d/%m/%Y')
print(data_atual)


email = outlook.CreateItem(0)#criando item "email"
soma1 = 4
soma2 = soma1/5
#corpo do email

#variavel_email = input("entre com o email: ")
email.To = "#"
#variavel_assunto = input("Entre com o assunto: ")
email.Subject = "Isso é um email com python"
email.HTMLBody = f"""
<p>isso é um teste<p>
<p>a soma é {soma2} <p>
Atividade em python, como enviar um email automatizado------------- e anexo e hoje é {x} 
"""
anexo =r"#" #entrar com caminho do anexo 
email.Attachments.Add(anexo)
email.Attachments.Add(excel_arquivo)
email.Send()
print(dataframe)
print("Email enviado")