'''
#Author: Ricardo Vila Longarço
#Supervisor: Lucas Massaro
#Date: 19/01/2022
'''
# Função utilizadas na formulação do relatório de aluguel
import pandas as pd

def obter_lista_du(hoje_corrido, qnt_dias, tipo):  # -> List[int] or List[str]: 
    #Fornece uma lista com dias uteis em um determiandi periodo 
    #Inputs: hoje_corrido: dia anterior(ontem) no formato 'anomesdia' ex: 20220112
    # qnt_dias: quantidade de dias desejada 
    #tipo: formato do output desejado 'Data_Traco' = 2022-03-14 (str) 'Dia_Barra_Mes' = 14/01 (str)
    # 'Dia_Numero' = 20220314 (int), 'Dia_Barra_Mes_Barra_Ano' = 14/03/2022 (str)

    hoje_final_dados_bruto = pd.read_excel('R:/Aluguel/Desenvolvimento/Documentos Compartilhados/CalendarioDU2.xlsx', sheet_name='Sheet1')
    index = int(hoje_final_dados_bruto[hoje_final_dados_bruto['Dia_Numero'] == hoje_corrido].index.values) - 1
    
    obter_lista_du = []
    for i in range(0,qnt_dias):    
            obter_lista_du.append(hoje_final_dados_bruto[tipo][index-i])
    return obter_lista_du

def valor_total(dicionario): 
    # Calcula a soma das posições abertas em contratos de aluguel em casa dia
    # Input: Dicionário contendo seus Keys (que são DataFrames) já filtradas pela função filtro
    valores_totais = {}    
    
    for i in range(0, len(dicionario)):
        df_temporario = dicionario[i]
        df_temporario['BalVal'] = df_temporario['BalVal'].replace(',','.', regex = True)
        df_temporario['BalVal'] = df_temporario['BalVal'].astype(float)
        valores_totais[i] = round(df_temporario['BalVal'].sum(), 0)
    
    # Output: Dicionário com a soma das posições abertas de cada dia em cada chave
    return valores_totais
     
def filtro_LoanBalance(free_float, taxas, df_aluguel):
    # Função que compila em um DataFrame o Ticker, quantiade de ações shorteadas, free float e taxa média de aluguel
    # Inputs: Tres DataFrames. O primeiro com o free float, o segundo com os dados de contratos de aluguel
    # e terceiro com a quantidade de ações alugadas de cada ativo
    df_final = pd.DataFrame(columns = ['Código', 'Quantidade ações', 'Free float'])
    for i in range(len(free_float)):
        cod = free_float['Código'].iloc[i]
        df_Fil = taxas[taxas.TckrSymb == cod]
        df_loan = df_aluguel[df_aluguel.TckrSymb == cod]
        if df_Fil.empty:
            continue
        if df_Fil.iloc[0][13] == 'Balcao':
            df_prov = pd.DataFrame({'Código': df_Fil.iloc[0][1],
                                 'Quantidade ações': df_loan.iloc[0][1],
                                 'Free float': free_float['Free Float'].iloc[i],
                                    'Taxa': df_Fil.iloc[0][11]}, index=[0])
    
            df_final = df_final.append(df_prov)
    # Outouts:  DataFrame com o Ticker, quantiade de ações shorteadas(alugadas), free float e taxa média de aluguel        
    return df_final

def short_interest(df_final_periodo):
    # Calcula o Short Interest de cada ativo presente nos df_final_periodo
    # Input: df_final_periodo calculado pela função filtro_LoanBalance
    df_si = pd.DataFrame(columns = ['Código', 'Short interest', 'Taxa'])
    for i in range(len(df_final_periodo)):
        if df_final_periodo['Free float'].iloc[i] == 0:
            continue
        else:
            si = (df_final_periodo['Quantidade ações'].iloc[i] / df_final_periodo['Free float'].iloc[i]) * 100
            if si >= 0.05:
                df_prov2 = pd.DataFrame({'Código': df_final_periodo['Código'].iloc[i],
                                     'Short interest' : si,
                                     'Taxa': df_final_periodo['Taxa'].iloc[i]}, index = [0])
                
                df_si = df_si.append(df_prov2)
    
    #Output: DataFrame com Ticker, Short Interest e Taxa 
    return df_si    

def var_si(df_aluguel_1d, df_aluguel_2d, df_aluguel_5d,
           df_si, df_si_ontem, df_si_semana, menor, df_setores):
    
    # Função que compila os dados finais que serão exportados para o excel
    # Inputs:DataFrames resultantes da função short_interest, o DataFrame de menor tamanho e DataFrame com o setor de cada ativo
    df_var_final = pd.DataFrame(columns = ['Código', 'Var % SI (1D)', 'Var % SI (1S)', 'Short interest', 'Taxa'])
    for i in range(0, len(menor)):
        cod = menor['Código'].iloc[i]
        df_Fil_c = df_si[df_si.Código == cod]  
        df_Fil1 = df_aluguel_1d[df_aluguel_1d.TckrSymb == cod]
        df_Fil2 = df_aluguel_2d[df_aluguel_2d.TckrSymb == cod]
        df_Fil3 = df_aluguel_5d[df_aluguel_5d.TckrSymb == cod]
        df_Fil4 = df_setores[df_setores.Código == cod]
        if df_Fil1.empty or df_Fil2.empty or df_Fil3.empty or df_Fil_c.empty:
            continue
        if df_Fil4.empty:
            #print("Ativo não encontrado")
            #print("Codigo: ",cod)
            continue                                                                                                 ##variação percentual de dois ratios
        si_var_dia = (df_Fil1['BalQty'].iloc[0] / df_Fil2['BalQty'].iloc[0]) - 1
        si_var_semana = (df_Fil1['BalQty'].iloc[0] / df_Fil3['BalQty'].iloc[0]) - 1
        df_prov3 = pd.DataFrame({'Código': df_Fil_c.iloc[0][0],
                                 'Var % SI (1D)' : round(si_var_dia,3),  
                                 'Var % SI (1S)' : str(round(si_var_semana,3) * 100) + '%',
                                 'Short interest': str(round(df_Fil_c['Short interest'].iloc[0], 4)) + '%',
                                 'Taxa': (df_Fil_c['Taxa'].iloc[0].replace(',','.')),
                                 'Setores': df_Fil4.iloc[0].iloc
                                  [5]}, index=[0])
        df_var_final = df_var_final.append(df_prov3)
    
        # Output: DataFrame df_var_final
    return df_var_final

