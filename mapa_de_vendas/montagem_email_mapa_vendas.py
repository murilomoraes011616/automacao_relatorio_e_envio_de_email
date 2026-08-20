import win32com.client
from datetime import date
import time

#---------------------
data_de_hoje = date.today()
outlook = win32com.client.Dispatch("Outlook.Application") # apenas liga o python ao outlook 
print("1 - conectado ao outlook") 
#---------------------


print("--------")
mail = outlook.CreateItem(0) #aqui ele segue a arvore, sendo outlook, ou outloo. aberto e crate item, que cria um item, o 0 significa o tipo de itrem, e 0 nesse caso significa email, entao a variavel email tem como resultado a criação de um email dentro do outlook 
print("2 - mail criado") 



print("--------")
inspetor = mail.GetInspector
assinatura = mail.HTMLBody
meu_texto = (
    f"Bom dia,<br><br>"
    f"Seguem anexos o mapa de venda, atualizado até {data_de_hoje}. "
    f"Atenciosamente,<br><br>"
)
mail.HTMLBody = meu_texto + assinatura
print("--------")



print("--------")
mail.Subject = f"Mapa de vendas do dia {data_de_hoje}." # feito para definir o assunto do email
assunto_do_email = mail.Subject
print(f"3 - o assunto do email é: {assunto_do_email}")
print("--------")


print("--------")
lista_emails = [
    "rafael.gomes@greentech.log.br",
    "paulo.chequetti@greentech.log.br",
    "ariel.rosenblatt@greentech.log.br",
    "felipe.andriolo@greentech.log.br",
    "marcelo.mota@greentech.log.br",
]
mail.To = ";".join(lista_emails) #lista de destinatarios do email, o join formaata cada valor entre ;, pois e o formato que o COM do outlook aceita 
destinatarios = mail.to 
print(f"5 - os destinatarios dos email são: {destinatarios}")
print("--------")


print("--------")
lista_copias = [
    "andre.santos@greentech.log.br", 
    "rodrigo.ferrarezzo@greentech.log.br", 
    "miguel.savtchen@greentech.log.br", 
    "karina.palmieri@greentech.log.br", 
    "marcelo.valerio@greentech.log.br", 
    "patricia.pinheiro@greentech.log.br", 
    "sarah@greentech.log.br",  
    "elisabete.ferreira@greentech.log.br",  
]
mail.CC = ";".join(lista_copias) #lista de copias do email
copias = mail.CC
print(f"6 - as copias dos email são: {copias}")
time.sleep(10)
print("--------")



print("--------")
arquivo_MAPA_VENDAS = r'C:\Users\murilo.oliveira\OneDrive - Greentech\Perfil\Desktop\pastas para coisas da  automações\automação de tabela toda segunda\MAPA DE VENDAS.pdf'
mail.Attachments.Add(arquivo_MAPA_VENDAS)
print(arquivo_MAPA_VENDAS)

print("esperando 10 segundos para poder abrir o display")
time.sleep(10) #trocar para 30 segundos quando entrar em produção 
mail.Send()    #decidi colocar display pra poder dar o aval e conferir o email, mas futuramente vou mandar automaticamente 
print("--------")
print(data_de_hoje)