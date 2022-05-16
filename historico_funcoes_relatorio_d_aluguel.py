def Datas(hoje_corrido, qnt_dias, tipo): 
    #Fornece uma lista com dias uteis em um determiandi periodo 
    #Inputs: hoje_corrido: dia anterior(ontem) no formato 'anomesdia' ex: 20220112
    # qnt_dias: quantidade de dias desejada 
    #tipo: formato do output desejado 'barra' = 14/03 (str) 'corrido' = 20220314 (int)

    hoje_final_dados_bruto = pd.read_excel('R:/Aluguel/Desenvolvimento/Summer/Ricardo Longarco/Relatorios/CalendarioDU2.xlsx',sheet_name='Sheet1')
    index = int(hoje_final_dados_bruto[hoje_final_dados_bruto['Dia_Numero'] == hoje_corrido].index.values) - 1
    
    dias = []
    
    for i in range(0,qnt_dias):
        if tipo == 'barra':       
            dias.append(hoje_final_dados_bruto['Dia_Barra_Mes'][index-i])
            
        elif tipo == 'corrido':
            dias.append(hoje_final_dados_bruto['Dia_Numero'][index-i])
    # Output: lista com dias uteis no periodo e formato selecionados
    return dias


def filtro(dicionario): 
    # Redefine cada Key do dicionario filtrando o DataFrame de cada dia e mantém colunas de TckrSymb (Ticker) e BalVal (Valor Transacionado)
    
    for i in range(0, len(dicionario)):
        df_temporario = pd.DataFrame()
        df_temporario = dicionario[i]
        df_temporario = df_temporario[['TckrSymb', 'BalVal']]
        dicionario[i] = df_temporario
    
    # Output: Keys do dicionario redefinidas para DataFrames com apenas as colunas TckrSymb e BalVal
    return dicionario

def filtro_aluguel(dicionario, periodo): 
    # Redefine cada Key do dicionario filtrando o DataFrame de cada dia e mantém colunas de TckrSymb (Ticker) e BalQty (Número de ações alugadas)       
        df_temporario = pd.DataFrame()
        if periodo == 'hoje':
            df_temporario = dicionario[0]
            df_temporario = df_temporario[['TckrSymb', 'BalQty']]
            return df_temporario 
        
        elif periodo == 'ontem':
            df_temporario = dicionario[1]   
            df_temporario = df_temporario[['TckrSymb', 'BalQty']]
            return df_temporario 
        
        elif periodo == 'semana':
            df_temporario = dicionario[4]   
            df_temporario = df_temporario[['TckrSymb', 'BalQty']]
    # Output: Keys do dicionario redefinidas para DataFrames com apenas as colunas TckrSymb e BalQty
        return df_temporario 
    