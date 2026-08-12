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


#tabela.api.AutoFilter(Field=16, Criteria1 = "#N/D") #tabela.api.autofilter tem o campo FIELD=16 que significa a coluna, que no caso de A até P, a coluna P é a 16, e o campo criterial é o filtro que vao colocar naquela coluna, e o filtro e so na quela coluna pois uarem os dados dela como parametro ams usa a tabela range inteira por que queremos os dados de todas as linhas mas o filtroé só em uma coluna 
#tabela.select() #apenas mostra a tange selecionada visualmente pro programador 
#tabela_filtrada = tabela.api.SpecialCells(12) #nessa linha tranforma a range definida acima em obejto para se tornar manipulavel  

#---------------------------------------------------------------------------------------------------------------------------
#abrir planilho filtro -Ajustado:
#proxima_linha_vazia = ultima_linha + 1
#----------------
#abrir_planilha.range(f'A{proxima_linha_vazia}').paste(paste='values') #Isso resolveria os dois sintomas de uma vez: o #VALOR! sumiria (porque você colaria o
#print(proxima_linha_vazia)
#------------------
#abrir_planilha_filtro_ajustado.api.ShowAllData()
# ultima_linha_filtro_ajustado = abrir_planilha_filtro_ajustado.range('A2').end('down').row #aqui range sgnifica um pedaçõ do codigo(P2) é o ponrto que ele usa como referencia, o .end(down) siginifia a mesma coisa que aperta ctrl seta ora baixo, entao vai pra ultima linha e .row te fala o nuemro dessa linha, ou seja ele usa a celula p2 como referencia, vai pra ultima linha e ega esse n8mero, oque sinigiffica a ultima linha da planilha 
# ultima_coluna_filtro_ajustado = abrir_planilha_filtro_ajustado.range('A2').end('right').column # mesma logica da linha de cima, porem ele quer saber aultima coluna, afim de fechar o quadrado da tabela total que vai ser selecionado para ser copiado no futuro 
#---------------------------------------------------------------------------------------------------------------------------


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
coluna_n_visivel = tabela_filtro.api.SpecialCells(12)
coluna_n_visivel.FormulaR1C1 = "=RC[-3]"


time.sleep(3)
abrir_planilha.api.ShowAllData() # tira os filtros, tudo visível de novo

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

#---------------------ATÉ AQUI TA CERTO ------------------------


