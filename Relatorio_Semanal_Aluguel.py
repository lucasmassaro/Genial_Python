'''
#Author: Ricardo Vila Longarço
#Supervisor: Lucas Massaro
#Date: 31/01/2022
'''
#Importação de módulo com funções criadas e valores de outros códigos para a elaboração do relatório
from Relatorio_Diario_Aluguel import valores_total_grafico, df_posicao, df_var_final
from funcoes_relatorio_semanal_aluguel import *
# Importação de bibliotecas necessárias para a criação do relatório
import pandas as pd
import xlwings as xw

hoje_corrido = '20220304'
df_var_final['Short interest'] = df_var_final['Short interest'].str.replace('%', '')
df_var_final['Short interest'] = df_var_final['Short interest'].astype(float)
df_var_final = df_var_final.sort_values(by= ['Short interest'], ascending= False)

df_top_si = df_var_final[['Código', 'Short interest', 'Taxa']].head(5)

colunas_utlizadas = [7, 15, 24, 27, 33]                                           
df_aberto = pd.read_excel(r'R:/Aluguel/Desenvolvimento/Documentos Diarios/Email/Posicao Atual/POSICAO ATUAL - ' + hoje_corrido + '.xls', usecols = colunas_utlizadas) 
df_aberto.rename(columns={'Cód. de Neg. do Ativo Obj.': 'Ticker', 'Preço de referência do Ativo-Objeto': 'Preço'}, inplace = True)

df_aberto['Volume Financeiro'] = df_aberto['Quantidade Atual'] * df_aberto['Preço']

df_nomes = pd.read_excel(r'R:/Aluguel/Desenvolvimento/Rotina Relatorio/Relatorio_Semanal_Aluguel/Base_Nome_Classificacao.xlsx')
df_aberto = pd.merge(df_aberto, df_nomes, on = 'Conta do Executor', how = 'inner')

dados_doador = filtro_geo(df_aberto, 'Doador')
dados_tomador = filtro_geo(df_aberto, 'Tomador')

df_dados_doador = pd.DataFrame(list(dados_doador.items()), columns = ['Geo Category', 'Total Lender Volume'])
df_dados_tomador = pd.DataFrame(list(dados_tomador.items()), columns = ['Geo Category', 'Total Borrower Volume'])

#enviando tudo para excel com algumas formatações
excel_app = xw.App(visible=False) #criando um app com a planilha (melhor controle dentro da planilha)
excel_book = excel_app.books.add() #criando um excel novo

ws = excel_book.sheets["Sheet1"] #adicionando a aba a uma variável
ws.name = 'Dados' #mudando o nome da aba da planilha
ws.range("A1").options(index = False).value = valores_total_grafico #a aba recebe o dataframe 
ws.range("D1").options(index = False).value = df_posicao #a aba recebe o dataframe 
ws.range("I1").options(index = False).value = df_top_si #a aba recebe o dataframe 
ws.range("A13").options(index = False).value = df_dados_doador #a aba recebe o dataframe 
ws.range("D13").options(index = False).value = df_dados_tomador #a aba recebe o dataframe 
excel_book.save(r'R:/Aluguel/Desenvolvimento/Rotina Relatorio/Relatorio_Semanal_Aluguel/Dados_Calculados/' + hoje_corrido + '-RSA.xlsx') #salvando em excel
excel_book.close()
excel_app.quit()