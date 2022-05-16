'''
#Author: Ricardo Vila Longarço
#Supervisor: Lucas Massaro
#Date: 19/01/2022
'''

# Importação de bibliotecas necessárias para a criação do relatório
import pandas as pd
import datetime as dt
import xlwings as xw
#import sys
#Importação de módulo com funções criadas para a elaboração do relatório
from funcoes_relatorio_aluguel import *

#Quantidades de dados(diários) que deseja-se ter no gráfico e dia de hoje no formato corrido e no formato para exportação
qnt_dados_grafico = 10
#hoje_completo = sys.argv[1]
#hoje_completo = '14/01/2022'
hoje_completo = '2022_05_13'

#dia = hoje_completo[:2]
#mes = hoje_completo[3:5]
#ano = hoje_completo[6:]
dia = hoje_completo[8:]
mes = hoje_completo[5:7]
ano = hoje_completo[:4]

hoje_corrido = int(ano+mes+dia)

#Função data explicada em funcoes_relaorio_aluguel.py
hoje_final_dados = obter_lista_du(hoje_corrido, qnt_dados_grafico, 'Dia_Numero') 

# Download de dados  
df_free_float = pd.read_excel(r'R:/Aluguel/Desenvolvimento/Rotina Relatorio/free_float_b3.xlsx')
taxas_1d = pd.read_csv(r'R:/Aluguel/Desenvolvimento/Documentos Diarios/B3/Loan Balance/LoanBalanceFile_' + str(hoje_final_dados[0]) + '_1.csv', sep = ';')
taxas_2d = pd.read_csv(r'R:/Aluguel/Desenvolvimento/Documentos Diarios/B3/Loan Balance/LoanBalanceFile_' + str(hoje_final_dados[1]) + '_1.csv', sep = ';')
taxas_5d = pd.read_csv(r'R:/Aluguel/Desenvolvimento/Documentos Diarios/B3/Loan Balance/LoanBalanceFile_' + str(hoje_final_dados[4]) + '_1.csv', sep = ';')
df_setores = pd.read_excel(r'R:/Aluguel/Desenvolvimento/Rotina Relatorio/setores_portugues.xlsx')               
# Dicionario com DataFrames diários e cada umade suas chaves
valores_grafico = {}
for i in range(0, qnt_dados_grafico):  
    valores_grafico[i] = pd.read_csv(r'R:/Aluguel/Desenvolvimento/Documentos Diarios/B3/Lending Open Position/LendingOpenPositionFile_' + str(hoje_final_dados[i]) + '_1.csv', sep = ';', usecols = [1, 4, 7])   

# Criação de dicionario com quantidade de ações alugadas
df_aluguel_1d = valores_grafico[0][['TckrSymb', 'BalQty']]
df_aluguel_2d = valores_grafico[1][['TckrSymb', 'BalQty']]
df_aluguel_5d = valores_grafico[4][['TckrSymb', 'BalQty']]

# Função valor_total explicada em funcoes_relaorio_aluguel.py
valores_total_grafico = valor_total(valores_grafico) 

# Calculando variações diária e semanal do volume total de ativos alugados
var_diaria = (valores_total_grafico[0] - valores_total_grafico[1]) / valores_total_grafico[1]  
var_semanal = (valores_total_grafico[0] - valores_total_grafico[5]) / valores_total_grafico[5]

# Função filtro_LoanBalance explicada em funcoes_relaorio_aluguel.py
df_final = filtro_LoanBalance(df_free_float, taxas_1d, df_aluguel_1d)
df_final_ontem = filtro_LoanBalance(df_free_float, taxas_2d, df_aluguel_2d)
df_final_semana = filtro_LoanBalance(df_free_float, taxas_5d, df_aluguel_5d)

# Função short_interest explicada em funcoes_relaorio_aluguel.py
df_si = short_interest(df_final)
df_si_ontem = short_interest(df_final_ontem)
df_si_semana = short_interest(df_final_semana)

# Definição do tamanho da iteração realizada na função var_si
if len(df_si) > len(df_si_ontem):
    df_menor = df_si_ontem
else:
    df_menor = df_si

#Função var_si explicada em funcoes_relaorio_aluguel.py
df_var_final = var_si(df_aluguel_1d, df_aluguel_2d, df_aluguel_5d,
                      df_si, df_si_ontem, df_si_semana, df_menor, df_setores)


# Obtendo as maiores e menores variações do SI, sendo a métrica a variação diária
df_var_final = df_var_final.sort_values(by= ['Var % SI (1D)'], ascending= False)
df_maiores = df_var_final.head(5) 
df_menores = df_var_final.tail(5) 

# Criando um DaraFrame com a posição do dia e a variação diária e semanal do volume de ativos alugados
df_posicao = pd.DataFrame({'Posição': valores_total_grafico[0], 
                           'Variação diária': [str(round((var_diaria * 100),2)) + '%'],
                           'Variação semanal': [str(round((var_semanal * 100),2)) + '%']}, index= [0])

# DataFrame com dados para gráfico de ativos mais alugados (em valores relativos ao mercado de aluguel)
maior_volume = pd.DataFrame(valores_grafico[0]) 
maior_volume['BalVal'] = maior_volume['BalVal'].astype(float)
maior_volume['BalVal'] = maior_volume['BalVal'].div(valores_total_grafico[0])
maior_volume = maior_volume.sort_values(by= ['BalVal'], ascending = False).head(10)

# DataFrame com dados para gráfico de ativos com maior SI no contrato de aluguel
maior_si = df_si[{'Código': df_si['Código'], 'Short interest': df_si['Short interest']}]
maior_si = maior_si.sort_values(by= ['Short interest'], ascending = False).head(10)

# Formatação dos dados para exportação ao excel
datas = obter_lista_du(hoje_corrido,qnt_dados_grafico,'Dia_Barra_Mes')
datas.reverse()
dados_grafico = list(valores_total_grafico.values())
dados_grafico.reverse()
valores_total_grafico = pd.DataFrame(list(zip(datas, dados_grafico)),columns =['Dia', 'Total'])

#enviando tudo para excel com algumas formatações
excel_app = xw.App(visible=False) #criando um app com a planilha (melhor controle dentro da planilha)
excel_book = excel_app.books.add() #criando um excel novo
ws = excel_book.sheets["Sheet1"] #adicionando a aba a uma variável
ws.name = 'Maiores' #mudando o nome da aba da planilha
ws.range("A1").options(index = False).value = df_maiores #a aba recebe o dataframe 
ws.range("A2:A6").api.Font.Bold = True #colocando em negrito
ws.range("B2:B6").number_format = "0,00%" #formatando para porcentagem
ws.range("F2:F6").api.HorizontalAlignment = -4108 #alinhando ao centro
ws.range("B2:B6").api.Font.Color = 0x008000 #colocado cor
ws2 = excel_book.sheets.add("Grafico_Linha") #adicionando nova aba
ws2.range("A1").options(dates= dt.date, index=False).value = valores_total_grafico #nova aba recebe o dataframe 
ws2.range("B1:B15").number_format = "#.##0" #colocando separador de milhar nos valores
ws2.range("A1:A15").number_format = "mm/dd" #formatando a data
ws3 = excel_book.sheets.add("Posicao") #adicionando aba
ws3.range("A1").options(index=False).value = df_posicao #aba recebe o dataframe 
ws3.range("A1:A2").number_format = "#.##0" #colocando separador de milhat nos valores
ws4 = excel_book.sheets.add("Menores") #adicionando nova aba
ws4.range("A1").options(index=False).value = df_menores #aba recebe dataframe 
ws4.range("A1:A10").api.Font.Bold = True #colcando em negrito
ws4.range("B2:B6").number_format = "0,00%" #fromatando para porcentagem
ws4.range("F2:F6").api.HorizontalAlignment = -4108 #alinhando ao centro
ws4.range("B2:B6").api.Font.Color = 0x000FF #colocando cor
ws5 = excel_book.sheets.add("Grafico_Volume") #adicionando nova aba
ws5.range("A1").options(index = False).value = maior_volume #aba recebe o dataframe
ws6 = excel_book.sheets.add("Grafico_SI") #adicionando nova aba
ws6.range("A1").options(index = False).value = maior_si #aba recebe o dataframe
excel_book.save(r'R:/Aluguel/Desenvolvimento/Rotina Relatorio/Relatorio_Diario_Aluguel/Dados_Calculados/' + ano+mes+dia + '-RDA.xlsx') #salvando em excel
excel_book.close()
excel_app.quit()

df_setores = pd.read_excel(r'R:/Aluguel/Desenvolvimento/Rotina Relatorio/setores_ibovmais.xlsx') 
df_var_final = var_si(df_aluguel_1d, df_aluguel_2d, df_aluguel_5d, df_si, df_si_ontem, df_si_semana, df_menor, df_setores)
#Importações de taxas e SI para o excel de Calculo do Aluguel
df_calc_aluguel = df_var_final[['Código','Short interest','Var % SI (1S)', 'Var % SI (1D)', 'Taxa']]
excel_app = xw.App(visible=False) #criando um app com a planilha (melhor controle dentro da planilha)
excel_book = excel_app.books.add() #criando um excel novo
ws = excel_book.sheets["Sheet1"] #adicionando a aba a uma variável
ws.name = 'Taxas e SI' #mudando o nome da aba da planilha
ws.range("A1").options(index = False).value = df_calc_aluguel
excel_book.save(r'R:\Aluguel\Desenvolvimento\Rotina Relatorio\Calculo Aluguel\Historico Taxas/Taxas_SI_' + hoje_completo + '.xlsx') #salvando em excel
excel_book.close()
excel_app.quit()

                                                           