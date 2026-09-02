################
import xlwings as xw # importa a biblioteca para manipular o excel .
from datetime import date,  timedelta
import time

app = xw.App(visible=False)   # cria a instância do Excel; visible=False roda em segundo plano
app.display_alerts = False   # suprime qualquer alerta/pop-up do Excel, incluindo esse
wb = app.books.open(
    r'U:\AREA_DE_DADOS\Indicadores\Gestao de Contratos\FILIAL SP\KPI - Faturamento\Mapa de Faturamento\Mapa_Faturamento_SAPHANA_v2.08.xlsm',
    update_links=0   # 0 = não atualiza vínculos automaticamente ao abrir, e não pergunta nada
)

abrir_planilha = wb.sheets('tabelas-auxiliares') #cria uma variavel, que dentro do objeto tem o metodo sheets, que abre a planilha que estiver em seu paramtro ().
qual_qual_valor_de_uma_celula = abrir_planilha.range('X41').value #encontra o valor da celullor_de_uma_celula) #printa esse valor.
print(f"data antiga: {qual_qual_valor_de_uma_celula}")#mostra o valor da celula X41


data_de_hoje = date.today()
primeiro_dia_mes_atual = data_de_hoje.replace(day=1)
ultimo_dia_mes_passado = primeiro_dia_mes_atual - timedelta(days=1)
hoje_formatado = data_de_hoje.strftime('%d-%m-%Y')   # pega a função de hoje da biblioteca datetime e guarda esse valor me um variavel.
abrir_planilha.range('X41').value = ultimo_dia_mes_passado #encontra esse valor da celula A1 e modifica ele.
qual_valor_ATUALIZADO_de_uma_celula = abrir_planilha.range('X41').value #variavel que equivale ao novo valorr.
print(f"data atualizada: {qual_valor_ATUALIZADO_de_uma_celula}")
wb.api.RefreshAll()  ##atualiza o arquivo todo pra puxar com as novas datas 
time.sleep(40) #espera 10 segundos no codigo apenas para poder para garantir a atualização dos dados

##ETAPA DE TIRAR O PRINT DA IMAGEM DA TABELA 

aba = wb.sheets('mapa_diario')   # 1. ele pega o wb.sheets na aba da tabela dinamica e trasnforma na variavel aba

aba.api.PageSetup.PrintArea = 'A2:O50'             # chegamos em uma parte que a biblioteca nao traduziu, então criou uma especie de porta dos fundos, a api., que usando a aba que queremos, e ela, depois podemos dar comandos que o excel usa porem nao traduzidos, normalmente em VBA, fazendo que possamos continuar a  programar em python, o .PageSetup

aba.api.PageSetup.Orientation = 2               # O que é PageSetup? É um objeto nativo do Excel que reúne todas as configurações relacionadas a impressão/exportação de página: margens, orientação, cabeçalho, rodapé, área de impressão, escala, etc. É exatamente o que você configura manualmente indo em Layout da Página no Excel, O que é Orientation? Define se a exportação será em retrato (vertical, como uma folha de carta em pé) ou paisagem (horizontal, deitada) — útil pra tabelas largas, como a sua.
                                                #Por que o número 2? Aqui é importante entender: como estamos usando o Excel/VBA "cru" através do .api, e não a versão traduzida do Python, não temos nomes bonitos disponíveis (tipo "paisagem"). O VBA original usa constantes numéricas pra isso:retrato = 1 e paisagem = 2

aba.api.PageSetup.Zoom = False                    # 4. Por que isso é necessário? O Excel tem duas formas de controlar o tamanho da exportação, que não podem ser usadas ao mesmo tempo:
                                                  #Zoom fixo (ex: "exportar em 100% do tamanho original")
                                                  #Ajuste automático pra caber em X páginas (que é o que vamos configurar nas próximas duas linhas)
                                                  #Por padrão, o Excel geralmente já vem com um Zoom fixo ativo (tipo 100), o que bloquearia o ajuste automático. Zoom = False desativa esse zoom fixo, "abrindo espaço" pra usar o ajuste automático nas linhas seguintes.
                                                  #Por que False e não um número? Porque, diferente de Orientation (que sempre é numérico), essa propriedade específica aceita tanto um número (se você quisesse um zoom fixo) quanto False (pra dizer "não use zoom fixo, vou usar ajuste automático"


aba.api.PageSetup.FitToPagesWide = 1              # 5. FitToPagesWide = 1 → a tabela inteira, não importa quantas colunas tenha, deve caber na largura de uma única página

aba.api.PageSetup.FitToPagesTall = 1              # 6. FitToPagesTall = 1 → a tabela inteira, não importa quantas linhas tenha, deve caber na altura de uma única página

aba.api.ExportAsFixedFormat(0, fr'C:\Users\murilo.oliveira\OneDrive - Greentech\Perfil\Desktop\pastas para coisas da  automações\automação de tabela toda segunda\Mapa de faturamento Resumo.pdf')  # 7. Claro! Vamos ler essa linha inteira em texto corrido, explicando o papel de cada parte conforme ela aparece.
                                                                                                                                          #A linha começa com aba, que é a variávelde você guardou a aba específica do Excel que contém a tabela (aquela que veio de wb.sheets('Indicador Faturamento-julho'), por exemplo). Em seguida vem .api, que é a "porta de entrada" para o objeto original do Excel — ou seja, a partir daqui você não está mais usando os comandos traduzidos pelo xlwings, mas sim os comandos nativos que o próprio Excel (através do VBA) sempre teve disponíveis. Depois vem .ExportAsFixedFormat, que é o método nativo do Excel responsável por gerar um arquivo em "formato fixo" — isto é, um formato que não muda de layout depois de criado, como PDF ou XPS. Esse método é o que efetivamente executa a ação de criar o arquivo; tudo que veio antes dele nas linhas anteriores (como configurar o PrintArea, a orientação, o ajuste de página) só preparou as condições, mas foi essa linha que de fato gerou o resultado.
                                                                                                                                          #Dentro dos parênteses desse método, você passa duas informações que ele precisa para funcionar. A primeira é o número 0, que representa o formato do arquivo a ser gerado — no sistema de constantes numéricas do VBA, 0 significa PDF e 1 significaria XPS, então usar 0 garante que o arquivo final seja um PDF. A segunda informação, separada por vírgula, é uma string que representa o caminho completo de onde esse arquivo será salvo no seu computador: r'C:\Users\murilo.oliveira\OneDrive - Greentech\Perfil\Desktop\lugar dos pdf para autoamação\tabela.pdf'. O r logo antes das aspas indica que essa é uma "raw string", ou seja, o Python deve interpretar as barras invertidas do caminho de forma literal, sem tratá-las como caracteres especiais — algo necessário porque caminhos do Windows sempre usam barras invertidas para separar pastas. Esse caminho, lido da esquerda para a direita, representa a navegação por pastas até o destino final: começa no disco C:, passa pela pasta do seu usuário Users\murilo.oliveira, entra na pasta sincronizada do OneDrive da empresa OneDrive - Greentech, segue para Perfil\Desktop, depois para uma pasta personalizada que você criou chamada lugar dos pdf para automação, e finalmente termina no nome do arquivo que será criado ali dentro, tabela.pdf.
                                                                                                                                          #Ou seja, lendo a linha inteira de forma corrida: você está pegando a aba já configurada, acessando o comando nativo do Excel para exportação, dizendo que o formato desejado é PDF, e informando exatamente em qual pasta e com qual nome esse arquivo deve ser salvo assim que for gerad

#ESQUEMA DESSE ALGORITMO POR HIERARQUIA:
#wb (Book, xlwings)
#  └── aba = wb.sheets(...) (Sheet, xlwings)
#        └── aba.api (Sheet original do Excel/VBA)
#              ├── .PageSetup (configurações de página)
#              │     ├── .PrintArea = 'C4:R20'
#              │     ├── .Orientation = 2
#              │     ├── .Zoom = False
#              │     ├── .FitToPagesWide = 1
#              │     └── .FitToPagesTall = 1
#              └── .ExportAsFixedFormat(0, 'caminho.pdf')  (método, executa a exportação)


###### ETAPA DE MANDAR O EMAIL ANEXADO PARA OS EMAIL ESCOLHIDOS:

import win32com.client
wb.save()
wb.close()