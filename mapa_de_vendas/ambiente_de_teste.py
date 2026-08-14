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


print("chegou aqui")
range_para_filtro = abrir_planilha.api.ListObjects(1).Range  # já é o range da tabela, direto do COM
range_para_filtro.AutoFilter(          # sem .api aqui — já é objeto COM cru
    Field=13,
    Criteria1=["Pecas", "EQPUsado", "EQPNovo", "Avaria", "Servicos", "Avaria", "PLPrev", "Comissao", "Comissão"],
    Operator=7  # xlFilterValues — diz "filtra por essa lista de valores"
)

print("Última linha:", ultima_linha)
tabela_filtro= abrir_planilha.range((2, 16), (ultima_linha, 16))
print("Endereço:", tabela_filtro.address)
coluna_P_visivel = tabela_filtro.api.SpecialCells(12)
coluna_P_visivel.FormulaR1C1 = "=RC[-3]"

time.sleep(3)
abrir_planilha.api.ShowAllData() # tira os filtros, tudo visível de novo
time.sleep(5)

print("Última linha:", ultima_linha)

tabela_filtro = abrir_planilha.range((2, 17), (ultima_linha, 17))
print("Endereço:", tabela_filtro.address)

coluna_q_visivel = tabela_filtro.api.SpecialCells(12)

for area in coluna_q_visivel.Areas:
    for celula in area:
        linha = celula.Row

        valor_m = abrir_planilha.range(f"M{linha}").value

        celula.Value = valor_m

time.sleep(3)

coluna_p_correcao = abrir_planilha.range((2, 16), (ultima_linha, 16))
coluna_p_correcao.api.Replace(What="Venda servicos", Replacement="Servicos", LookAt=1)
coluna_p_correcao.api.Replace(What="Novonegocio", Replacement="Novo contrato", LookAt=1) 
tabela_filtro.api.AutoFilter(Field=16, Criteria1 = "-".upper().strip() )
tabela_filtro2= abrir_planilha.range((2, 16), (ultima_linha, 16))

time.sleep(5)

print("---------- listagem dos pvs a serem excluidos ----------")
range_colunaA = abrir_planilha.range((2, 1), (ultima_linha, 1)) #aqui nessa linha ele pega a range que passou pelo filtro 
coluna_a_visivel = range_colunaA.api.SpecialCells(12)  # ja aqui tonra essa rnage filtrada, visivel, para que possamos manuipular os valores dela 

valores = set()      #set tem a função dde receber valores e excluir aqueles repetidos                                   
for celula in coluna_a_visivel:    #os valores da range que a gente pegou, pega valor por valor                      
    valores.add(celula.Value)                 #adiciona o campo value desses valores dentro do nosso set                                            
for valor in valores:
    print(valor)    #esse fot foi para pegar os valores que estao dentro do set valores, que sao aqueles que nao estão repetidos, pois se nao pegasse os valores dentro dele, pegaria os que nao tao dentro do set, logo os repetidos.



print("--------------- processo de listar PVS a serem excluidos ---------------")
time.sleep(5)
abrir_planilha_PVS_deletados = wb.sheets("pv_excluidos") #abre a panilha de pvs pra serem deletados 
proxima_linha_vazia_pv_excluidos = abrir_planilha_PVS_deletados.range('A1').end('down').row + 1 #faz a mesma coisa de antes, porem agora usa como referencia a primeira celula da minha range,.end('down') vai até a ultima linha preenchida + 1, oque da na primeira linha vazia da primeira coluna, que e onde a gente vai colar nossas infromações 

print(proxima_linha_vazia_pv_excluidos)


time.sleep(5)

print("---------- jogar os pvs na outra planilha ----------")
linha_pv_excluido = proxima_linha_vazia_pv_excluidos # tranforma o nome da varaivel que busca a ultima linha 
for docnum in valores:         #para cada valor dentro do set valores 
    abrir_planilha_PVS_deletados.range((linha_pv_excluido,1)).value = docnum       #pega a proxima linha vazia na coluna a
    linha_pv_excluido += 1           #faz virar a proima linha 
    print(f" a linha 'A{linha_pv_excluido}' recebe o valor {docnum}") #print apenas para vizualização 

