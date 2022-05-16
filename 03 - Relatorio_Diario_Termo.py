'''
#Author: Ricardo Vila Longarço
#Supervisor: Lucas Massaro
#Date: 21/01/2022
'''
# Importação de bibliotecas necessárias para a criação do relatório
import pandas as pd
import datetime as dt
import xlwings as xw
#Importação de módulo com funções criadas para a elaboração do relatório
from funcoes_relatorio_termo import *

#Quantidades de dados(diários) que deseja-se ter no gráfico e dia de hoje no formato corrido e no formato para exportação
qnt_dados_grafico = 10
#hoje_completo = sys.argv[1]
#hoje_completo = '14/01/2022'
hoje_completo = '2022_05_13'

dia = hoje_completo[8:]
mes = hoje_completo[5:7]
ano = hoje_completo[:4]
hoje_corrido = int(ano+mes+dia)

#Função data explicada em funcoes_relaorio_termo.py
hoje_final_corrido = obter_lista_du(hoje_corrido, qnt_dados_grafico, 'Dia_Numero') 
hoje_final_dados = obter_lista_du(hoje_corrido, qnt_dados_grafico, 'Dia_Barra_Mes_Barra_Ano') 

#Obtendo dados dos setores
df_setores = pd.read_excel(r'R:/Aluguel/Desenvolvimento/Rotina Relatorio/setores_portugues.xlsx')

#Obtenção de dados de D-1 e D-2
url_1d = 'http://www.b3.com.br/pt_br/market-data-e-indices/servicos-de-dados/market-data/consultas/mercado-a-vista/termo/posicoes-em-aberto/posicoes-em-aberto-8AA8D0CC77D179750177DF167F150965.htm?data=' + str(hoje_final_dados[0]) +'&f=0'
url_2d = 'http://www.b3.com.br/pt_br/market-data-e-indices/servicos-de-dados/market-data/consultas/mercado-a-vista/termo/posicoes-em-aberto/posicoes-em-aberto-8AA8D0CC77D179750177DF167F150965.htm?data=' + str(hoje_final_dados[1]) +'&f=0'
df_1d = pd.read_html(url_1d)    
df_2d = pd.read_html(url_2d)
df_1d_valores = df_1d[0]
df_1d_empresas = df_1d[1]
df_2d_valores = df_2d[0]
df_2d_empresas = df_2d[1]

df_5d_empresas = pd.read_excel(r'R:/Aluguel/Desenvolvimento/Rotina Relatorio/Relatorio_Diario_Termo/PosicaoTermo/posicao_aluguel_' + str(hoje_final_corrido[4]) + '.xlsx')

#Obtenção de dados de D-3 até D-10
valores_grafico = {}
for i in range(2, qnt_dados_grafico):
    valores_grafico[i] = pd.read_excel((r'R:/Aluguel/Desenvolvimento/Rotina Relatorio/Relatorio_Diario_Termo/PosicaoTermo/posicao_termo_' +str(hoje_final_corrido[i]) + '.xlsx' ))   

#Preparando DataFrame para receber o valor total em contrato de termo
df_total_final = pd.DataFrame(columns= ['Dia', 'Total'])
hoje_final_dados.reverse()
df_total_final['Dia'] = hoje_final_dados

df_1d_empresas = Filtro_Ibov(df_1d_empresas, df_setores)
df_2d_empresas = Filtro_Ibov(df_2d_empresas, df_setores)
df_5d_empresas = Filtro_Ibov(df_5d_empresas, df_setores)

#Abastecendo o DataFrame com os dados dos ultimos 10 dias
#Função data explicada em funcoes_relaorio_termo.py
df_total_final = Valores_Totais(df_1d_valores, df_2d_valores, valores_grafico, df_total_final, qnt_dados_grafico)    
lista_temp = df_total_final['Total'].tolist()
lista_temp.reverse()
df_total_final['Total'] = lista_temp

#Calculando variação diária e semanal do volume de contratos de termo
var_dia = (df_total_final.iloc[0][1] / df_total_final.iloc[1][1] - 1) * 100
var_semana = (df_total_final.iloc[0][1] / df_total_final.iloc[4][1] - 1) * 100

#Define qual o menor DataFrame para nao estourar o contador da iteração abaixo
if len(df_1d_empresas) >= len(df_2d_empresas):
    menor = df_2d_empresas
else:
    menor = df_1d_empresas     

#Criando DataFrame com variacoes diarias e semanais da quantidade de cada ativo em contrato de termo
#Função data explicada em funcoes_relaorio_termo.py
df_var_final = Var_Termo(menor, df_1d_empresas, df_2d_empresas, df_5d_empresas) 
df_var_final = df_var_final.drop_duplicates(subset = ['Código'])

#Obtendo DataFrames com as maiores e menores variações diárias da quantidade de cada ativo em contrato de termo
#Função data explicada em funcoes_relaorio_termo.py
df_maiores_var = Top5_Var('maiores', df_var_final)
df_menores_var = Top5_Var('menores', df_var_final)

#Criando o dataframe com variações diárias, semanal e poisção atual
df_posicao = pd.DataFrame({'Posição': df_total_final.iloc[0][1] / 100, 'Variação diária': [str(var_dia) + '%'], 'Variação semanal': [str(round((var_semana),2)) + '%']}, index= [0])

#Adicionando os setores aos ativos com maiores e menores variações diárias
df_final_maiores_var = Setores_Termo(df_maiores_var, df_menores_var, df_setores, 'maiores')
df_final_menores_var = Setores_Termo(df_maiores_var, df_menores_var, df_setores, 'menores')

#Formatando valores para importação ao excel
df_total_final['Total'] = df_total_final['Total'].div(100)
lista_temp = list(df_total_final['Total'])
lista_temp.reverse()
df_total_final['Total'] = lista_temp

#Enviando tudo para excel com algumas formatações
excel_app = xw.App(visible=False) #criando um app com a planilha (melhor controle dentro da planilha)
excel_book = excel_app.books.add() #criando um excel novo
ws = excel_book.sheets["Sheet1"] #adicionando a aba a uma variável ##Aqui da erro pq meu excel ta em ptbr colocar Planilha1 um ao inves de Sheet1
ws.name = 'Maiores variações' #mudando o nome da aba da planilha
ws.range("A1").options(index = False).value = df_final_maiores_var #a aba recebe o dataframe
ws.range("A2:A6").api.Font.Bold = True #colocando em negrito
ws.range("D2:D6").number_format = "#.##0" #colocando separador de milhar nos valores
ws.range("D2:D6").api.HorizontalAlignment = -4108 #alinhando ao centro
ws.range("B2:B6").api.HorizontalAlignment = -4108 #alinhando ao centro
ws.range("C2:C6").api.HorizontalAlignment = -4108 #alinhando ao centro
ws.range("B2:B6").api.Font.Color = 0x008000 #colocado cor
ws.range("E2:E6").api.HorizontalAlignment = -4108
ws2 = excel_book.sheets.add("Valores gráfico") #adicionando nova aba
ws2.range("A1").options(dates= dt.date, index=False).value = df_total_final #nova aba recebe o dataframe
ws2.range("A1:A15").number_format = "mm/dd" #formatando a data
ws3 = excel_book.sheets.add("Posição + variações") #adicionando aba
ws3.range("A1").options(index=False).value = df_posicao #aba recebe o dataframe
ws3.range("A1:A2").number_format = "#.##0" #colocando separador de milhat nos valores
ws3.range("B1:C1").number_format = "0%" #formatando os valores para porcentagem
ws4 = excel_book.sheets.add("Menores variações") #adicionando nova aba
ws4.range("A1").options(index=False).value = df_final_menores_var #aba recebe dataframe
ws4.range("A1:A10").api.Font.Bold = True #colcando em negrito
ws4.range("D2:D6").api.HorizontalAlignment = -4108 #alinhando ao centro
ws4.range("B2:B6").api.HorizontalAlignment = -4108 #alinhando ao centro
ws4.range("C2:C6").api.HorizontalAlignment = -4108 #alinhando ao centro
ws4.range("B2:B6").api.Font.Color = 0x000FF #colocando cor
ws4.range("E2:E6").api.HorizontalAlignment = -4108
ws4.range("D2:D6").number_format = "#.##0" #colocando separador de milhar nos valores
excel_book.save(r'R:/Aluguel/Desenvolvimento/Rotina Relatorio/Relatorio_Diario_Termo/Dados_Calculados/' + str(ano) + str(mes) + str(dia) + '-RDT.xlsx') #salvando em excel
excel_book.close()
excel_app.quit()
