'''
#Author: Ricardo Vila Longarço
#Supervisor: Lucas Massaro
#Date: 
'''
# Funções utilizadas na formulação do relatório semanal de aluguel

def filtro_geo(df_aberto, doador_tomador):
    dicionario = {}
    
    if doador_tomador == 'Doador':
        df_temp = df_aberto.loc[df_aberto['Natureza'] == 'Doador']
        df_temp = df_temp[['Volume Financeiro', 'Geo Classification']]
        
        df_local_fund = df_temp.loc[df_temp['Geo Classification'] == 'Local Fund']
        soma = df_local_fund['Volume Financeiro'].sum()
        dicionario['Local Fund'] = soma 
        
        df_foreign_fund = df_temp.loc[df_temp['Geo Classification'] == 'Foreign Fund']
        soma = df_foreign_fund['Volume Financeiro'].sum()
        dicionario['Foreign Fund'] = soma
        
        #df_pension_fund = df_temp.loc[df_temp['Geo Classification'] == 'Pension Fund']
        #soma = df_pension_fund['Volume Financeiro'].sum()
        #dicionario['Pension Fund'] = soma 
        
        df_investment_club = df_temp.loc[df_temp['Geo Classification'] == 'Investment Club']
        soma = df_investment_club['Volume Financeiro'].sum()
        dicionario['Investment Club'] = soma
        
        df_retail = df_temp.loc[df_temp['Geo Classification'] == 'Retail']
        soma = df_retail['Volume Financeiro'].sum()
        dicionario['Retail'] = soma
        
        return dicionario
    
    elif doador_tomador == 'Tomador':
        df_temp = df_aberto.loc[df_aberto['Natureza'] == 'Tomador']
        df_temp = df_temp[['Volume Financeiro', 'Geo Classification']]
        
        df_local_fund = df_temp.loc[df_temp['Geo Classification'] == 'Local Fund']
        soma = df_local_fund['Volume Financeiro'].sum()
        dicionario['Local Fund'] = soma 
        
        df_foreign_fund = df_temp.loc[df_temp['Geo Classification'] == 'Foreign Fund']
        soma = df_foreign_fund['Volume Financeiro'].sum()
        dicionario['Foreign Fund'] = soma
        
        #df_pension_fund = df_temp.loc[df_temp['Geo Classification'] == 'Pension Fund']
        #soma = df_pension_fund['Volume Financeiro'].sum()
        #dicionario['Pension Fund'] = soma 
        
        #df_investment_club = df_temp.loc[df_temp['Geo Classification'] == 'Investment Club']
        #soma = df_investment_club['Volume Financeiro'].sum()
        #dicionario['Investment Club'] = soma
        
        df_retail = df_temp.loc[df_temp['Geo Classification'] == 'Retail']
        soma = df_retail['Volume Financeiro'].sum()
        dicionario['Retail'] = soma
        
        return dicionario
