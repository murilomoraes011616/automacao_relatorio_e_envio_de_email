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


time.sleep(3)
range_para_filtro = abrir_planilha.api.ListObjects(1).Range  # já é o range da tabela, direto do COM
range_para_filtro.AutoFilter(          # sem .api aqui — já é objeto COM cru
    Field=16,
    Criteria1=["-", " ", ""],
    Operator=7  # xlFilterValues — diz "filtra por essa lista de valores"
)
time.sleep(10)

try:
    print("---------- listagem dos pvs a serem excluidos ----------")
    ultima_linha_filtro =abrir_planilha.range('A2').end('down').row
    range_colunaA = abrir_planilha.range((2, 1), (ultima_linha_filtro, 1)) #aqui nessa linha ele pega a range que passou pelo filtro 
    coluna_a_visivel = range_colunaA.api.SpecialCells(12)  # ja aqui tonra essa rnage filtrada, visivel, para que possamos manuipular os valores dela 
except Exception:
    print("⚠️ Nenhum PV para excluir dessa vez: o filtro não encontrou '-', ' ' ou '' na coluna P.")
    valores = set()


valores = set()      #set tem a função dde receber valores e excluir aqueles repetidos                                   
for celula in coluna_a_visivel:    #os valores da range que a gente pegou, pega valor por valor                      
    valores.add(celula.Value)                 #adiciona o campo value desses valores dentro do nosso set                                            
for valor in valores:
    print(valor)    #esse fot foi para pegar os valores que estao dentro do set valores, que sao aqueles que nao estão repetidos, pois se nao pegasse os valores dentro dele, pegaria os que nao tao dentro do set, logo os repetidos.



time.sleep(5)
abrir_planilha_PVS_deletados = wb.sheets("pv_excluidos") #abre a panilha de pvs pra serem deletados 
proxima_linha_vazia_pv_excluidos = abrir_planilha_PVS_deletados.range('A1').end('down').row + 1 #faz a mesma coisa de antes, porem agora usa como referencia a primeira celula da minha range,.end('down') vai até a ultima linha preenchida + 1, oque da na primeira linha vazia da primeira coluna, que e onde a gente vai colar nossas infromações 

print(f"a proxima linha vazia é: {proxima_linha_vazia_pv_excluidos}")


time.sleep(5)

print("---------- jogar os pvs na outra planilha ----------")
linha_pv_excluido = proxima_linha_vazia_pv_excluidos # tranforma o nome da varaivel que busca a ultima linha 
for docnum in valores:         #para cada valor dentro do set valores 
    abrir_planilha_PVS_deletados.range((linha_pv_excluido,1)).value = docnum       #pega a proxima linha vazia na coluna a
    linha_pv_excluido += 1           #faz virar a proima linha 
    print(f" a linha 'A{linha_pv_excluido}' recebe o valor {docnum}") #print apenas para vizualização 

#para diminuir os erros, criei uma coluna igual a P, porem como ela usa os valores fixos como referencia, na hra que eu atualizar  o power query os valores nao serao formatados, assim posso fazer eventuais correcoes no tipo de canal caso exista algum erro, e passo a usar essa nova coluna para fazer asreferencias da aba de mapa diario 
ultima_linha = abrir_planilha.range(
    "P" + str(abrir_planilha.cells.last_cell.row)
).end("up").row

# Limpa a coluna R
abrir_planilha.range(f"R2:R{ultima_linha}").clear_contents()

# Copia P para R, linha por linha, como valor
for linha in range(2, ultima_linha + 1):
    valor = abrir_planilha.range(f"P{linha}").value
    abrir_planilha.range(f"R{linha}").value = valor
#----------------------------------------



wb.api.RefreshAll()
time.sleep(20)
time.sleep(10)
#wb.save()






aba = wb.sheets('Mapa Diário')   # 1. ele pega o wb.sheets na aba da tabela dinamica e trasnforma na variavel aba

aba.api.PageSetup.PrintArea = 'A2:S44'             # chegamos em uma parte que a biblioteca nao traduziu, então criou uma especie de porta dos fundos, a api., que usando a aba que queremos, e ela, depois podemos dar comandos que o excel usa porem nao traduzidos, normalmente em VBA, fazendo que possamos continuar a  programar em python, o .PageSetup

aba.api.PageSetup.Orientation = 2               # O que é PageSetup? É um objeto nativo do Excel que reúne todas as configurações relacionadas a impressão/exportação de página: margens, orientação, cabeçalho, rodapé, área de impressão, escala, etc. É exatamente o que você configura manualmente indo em Layout da Página no Excel, O que é Orientation? Define se a exportação será em retrato (vertical, como uma folha de carta em pé) ou paisagem (horizontal, deitada) — útil pra tabelas largas, como a sua.
                                                #Por que o número 2? Aqui é importante entender: como estamos usando o Excel/VBA "cru" através do .api, e não a versão traduzida do Python, não temos nomes bonitos disponíveis (tipo "paisagem"). O VBA original usa constantes numéricas pra isso:retrato = 1 e paisagem = 2

aba.api.PageSetup.Zoom = False                    # 4. Por que isso é necessário? O Excel tem duas formas de controlar o tamanho da exportação, que não podem ser usadas ao mesmo tempo:
                                                  #Zoom fixo (ex: "exportar em 100% do tamanho original")
                                                  #Ajuste automático pra caber em X páginas (que é o que vamos configurar nas próximas duas linhas)
                                                  #Por padrão, o Excel geralmente já vem com um Zoom fixo ativo (tipo 100), o que bloquearia o ajuste automático. Zoom = False desativa esse zoom fixo, "abrindo espaço" pra usar o ajuste automático nas linhas seguintes.
                                                  #Por que False e não um número? Porque, diferente de Orientation (que sempre é numérico), essa propriedade específica aceita tanto um número (se você quisesse um zoom fixo) quanto False (pra dizer "não use zoom fixo, vou usar ajuste automático"


aba.api.PageSetup.FitToPagesWide = 1              # 5. FitToPagesWide = 1 → a tabela inteira, não importa quantas colunas tenha, deve caber na largura de uma única página

aba.api.PageSetup.FitToPagesTall = 1              # 6. FitToPagesTall = 1 → a tabela inteira, não importa quantas linhas tenha, deve caber na altura de uma única página

aba.api.ExportAsFixedFormat(0, fr'C:\Users\murilo.oliveira\OneDrive - Greentech\Perfil\Desktop\pastas para coisas da  automações\automação de tabela toda segunda\MAPA DE VENDAS.pdf')  # 7. Claro! Vamos ler essa linha inteira em texto corrido, explicando o papel de cada parte conforme ela aparece.
print("---------- Programa finalizado ----------")

