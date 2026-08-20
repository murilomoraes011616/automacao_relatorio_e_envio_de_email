import xlwings as xw # importa a biblioteca para manipular o excel .
from datetime import date   # importa sobemente a função date da biblioteca datetime de pega a data, biblioteca do python ja.
import time # importa biblioteca para poder dar o comando de esperar 10 segundos 


app = xw.App(visible=True)   # cria a instância do Excel; visible=False roda em segundo plano
app.display_alerts = False   # suprime qualquer alerta/pop-up do Excel, incluindo esse
wb = app.books.open(
    r'U:\AREA_DE_DADOS\Indicadores\Gestao de Contratos\FILIAL SP\KPI - Mapa de Vendas\Mapa de vendas v0 - Jul26.xlsx',
    update_links=0   # 0 = não atualiza vínculos automaticamente ao abrir, e não pergunta nada
)


abrir_planilha = wb.sheets('Pedido de venda') #nessa linha tranformamos o wb.sheets(Pedido de Venda) em uma varaivel, basicamente esse sheet é uma função do wlwin
abrir_planilha.activate() #essa linha faz com que mostre pra mim que aba pedido de vendas foi aberta, pois na linha de cima ela so entrou na aba, mas não significa que mostrou pra mim, o usúario
wb.api.RefreshAll()
time.sleep(15) #como a linha de cima so atualiza e nas proximas linhas vou precisar trabalhar com os dados atualizados, essa linha garante de uma forma bem ruim que ls dados estejam atualizados antes de eu dar o proximo comando para o excel, é uma forma que da pra melhorar, mas por enquanto, vai servir.


ultima_linha = abrir_planilha.range('P2').end('down').row #aqui range sgnifica um pedaçõ do codigo(P2) é o ponrto que ele usa como referencia, o .end(down) siginifia a mesma coisa que aperta ctrl seta ora baixo, entao vai pra ultima linha e .row te fala o nuemro dessa linha, ou seja ele usa a celula p2 como referencia, vai pra ultima linha e ega esse n8mero, oque sinigiffica a ultima linha da planilha 
ultima_coluna = abrir_planilha.range('P2').end('right').column # mesma logica da linha de cima, porem ele quer saber aultima coluna, afim de fechar o quadrado da tabela total que vai ser selecionado para ser copiado no futuro 
tabela = abrir_planilha.range((2, 1), (ultima_linha, ultima_coluna)) # abrir_planilha.tange(2,) significa o ponto que usaremos como referencia para começar a range, que seria a linha 2 e a celula 1, igual A2 (pra não pegar o cabeçalho) e (ultima_linha, ultima_coluna) gnifica a ultima celula que ele vai pegar, que e a ultima celula x ultima linha, assim pegando o quadrado todo para que possamos palicar o filtro, no caso retornara A2:P3427.
print(tabela) #print o valor da range acima 

time.sleep(5)

print("Última linha:", ultima_linha)

tabela_filtro = abrir_planilha.range((2, 16), (ultima_linha, 16))
print("Endereço:", tabela_filtro.address)
#-----------------------------------------------------------------------------------
range_para_filtro = abrir_planilha.api.ListObjects(1).Range  # já é o range da tabela, direto do COM
range_para_filtro.AutoFilter(          # sem .api aqui — já é objeto COM cru
    Field=16,
    Criteria1=["Avaria", "avaria"],
    Operator=7  # xlFilterValues — diz "filtra por essa lista de valores"
) 

range_para_filtro = abrir_planilha.api.ListObjects(1).Range  # já é o range da tabela, direto do COM
range_para_filtro.AutoFilter(          # sem .api aqui — já é objeto COM cru
    Field=17,
    Criteria1=["", "0", " "],
    Operator=7  # xlFilterValues — diz "filtra por essa lista de valores"
)

# ListColumns("DocNum").DataBodyRange pega só os dados da coluna DocNum,
# já excluindo a linha de cabeçalho automaticamente
coluna_docnum = abrir_planilha.api.ListObjects(1).ListColumns("DocNum").DataBodyRange

# agora sim: só as células visíveis DESSA coluna, depois dos dois filtros
linhas_visiveis = coluna_docnum.SpecialCells(12)

valores = set()
for celula in linhas_visiveis:
    valores.add(celula.Value)

for valor in valores:
    print(valor)


#---------------------------------------



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
    f"seguem os PVS de avaria que estão sem OS: {valores} <br><br>"
    f"Atenciosamente,<br><br>"
)
mail.HTMLBody = meu_texto + assinatura
print("--------")
print("--------")
lista_copias = [
    "", 
]
mail.CC = ";".join(lista_copias) #lista de copias do email
copias = mail.CC
print(f"6 - as copias dos email são: {copias}")
time.sleep(10)
print("--------")


print("--------")
mail.Subject = f"PVS sem OS avaria do Mapa de vendas, dia {data_de_hoje}." # feito para definir o assunto do email
assunto_do_email = mail.Subject
print(f"3 - o assunto do email é: {assunto_do_email}")
print("--------")


print("--------")
lista_emails = [
    "nathiele.belo@greentech.log.br",
]
mail.To = ";".join(lista_emails) #lista de destinatarios do email, o join formaata cada valor entre ;, pois e o formato que o COM do outlook aceita 
destinatarios = mail.to 
print(f"5 - os destinatarios dos email são: {destinatarios}")
print("--------")



mail.CC = ";".join(lista_copias) #lista de copias do email
copias = mail.CC
print(f"6 - as copias dos email são: {copias}")
time.sleep(10)
print("--------")



print("esperando 10 segundos para poder abrir o display")
time.sleep(10) #trocar para 30 segundos quando entrar em produção 
mail.Display()    #decidi colocar display pra poder dar o aval e conferir o email, mas futuramente vou mandar automaticamente 
print("--------")
print(data_de_hoje)