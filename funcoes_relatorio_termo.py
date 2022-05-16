'''
#Author: Ricardo Vila Longarço
#Supervisor: Lucas Massaro
#Date: 21/01/2022
'''
# Função utilizadas na formulação do relatório de termo
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

def Filtro_Ibov(df_d_empresas, df_setores):
    df_f = pd.DataFrame( columns = ['Código', 'Empresa', 'Tipo', 'Qtde de Contratos', 'Qtde de Ativos', 'Valor de Contratos (R$)'])

    for i in range(len(df_setores)):
        cod = df_setores.iloc[i][0]
        df_p = df_d_empresas[df_d_empresas['Código'] == cod]
        
        df_f = df_f.append(df_p)       
    return df_f
       
def Valores_Totais(df_1d_valores, df_2d_valores, valores_grafico, df_total_final, qnt_dados_grafico):
    df_total_final['Total'].iloc[0] = float(df_1d_valores.iloc[0][2].translate({ord(c): None for c in '.'}).replace(',', ''))
    df_total_final['Total'].iloc[1] = float(df_2d_valores.iloc[0][2].translate({ord(c): None for c in '.'}).replace(',', ''))
    for i in range(2, qnt_dados_grafico):
        df_temp = valores_grafico[i]
        df_total_final['Total'].iloc[i] = float(df_temp.iloc[0][2].translate({ord(c): None for c in '.'}).replace(',', ''))

    return df_total_final

def Var_Termo(menor, df_1d_empresas, df_2d_empresas, df_5d_empresas):


    df_var_final = pd.DataFrame(columns= ['Código', 'Quantidade', 'Variação diária', 'Variação semanal'])
    
    for i in range(len(menor)):        
        cod = df_1d_empresas.iloc[i][0]
        df_Fil1 = df_1d_empresas[df_1d_empresas.Código == cod]
        df_Fil2 = df_2d_empresas[df_2d_empresas.Código == cod]
        df_Fil3 = df_5d_empresas[df_5d_empresas.Código == cod]
        if df_Fil1.empty or df_Fil2.empty or df_Fil3.empty:
            continue
        var_dia = (float(df_Fil1.iloc[0][4].translate({ord(c): None for c in "."}).replace(',', '.')) / float(df_Fil2.iloc[0][4].translate({ord(c): None for c in "."}).replace(',', '.'))) - 1
        var_sem = (float(df_Fil1.iloc[0][4].translate({ord(c): None for c in "."}).replace(',', '.')) / float(df_Fil3.iloc[0][4].translate({ord(c): None for c in "."}).replace(',', '.'))) - 1
        #contruindo dataframe provisório
        if float(df_Fil2.iloc[0][4].translate({ord(c): None for c in "."})) > 100000:
            df_var = pd.DataFrame({'Código' : df_Fil2.iloc[0][0],
                               'Quantidade' : df_Fil2.iloc[0][4],
                               'Variação diária' : [(var_dia) * 100],
                                   'Variação semanal' : [(var_sem) * 100]})
            df_var_final = df_var_final.append(df_var)
            
    return df_var_final

def Top5_Var(tipo, df_var_final):

    df_rank_menor = pd.DataFrame( columns = ['Código', 'Quantidade', 'Menor variação diária', 'Variação semanal'])
    df_rank_maior = pd.DataFrame( columns = ['Código', 'Quantidade', 'Maior variação diária', 'Variação semanal'])
    
    if tipo == 'maiores':
        df_menor = df_var_final.sort_values(by= ['Variação diária'], ascending = False)
        
        for i in range(5):
            df_prov_menor = pd.DataFrame({'Código' : df_menor.iloc[i][0],
                                          'Quantidade' : df_menor.iloc[i][1],
                                          'Menor variação diária' : df_menor.iloc[i][2],
                                          'Variação semanal' : df_menor.iloc[i][3]
                                          }, index = [0])
            df_rank_menor = df_rank_menor.append(df_prov_menor)
        return df_rank_menor
    
    elif tipo == 'menores':
        df_maior = df_var_final.sort_values(by= ['Variação diária'], ascending = True)
        
        for i in range(5):
            df_prov_maior = pd.DataFrame({'Código' : df_maior.iloc[i][0],
                                          'Quantidade' : df_maior.iloc[i][1],
                                          'Maior variação diária' : df_maior.iloc[i][2],
                                          'Variação semanal': df_maior.iloc[i][3]
                                          }, index = [0])
            df_rank_maior = df_rank_maior.append(df_prov_maior)  
        return  df_rank_maior

def Setores_Termo(df_maiores_var, df_menores_var, df_setores, tipo):
    df_final_pdf = pd.DataFrame(columns = ['Código', 'Variação diária', 'Variação semanal','Quantidade', 'Setor'])
    for j in range(len(df_maiores_var)):
        maiores = df_maiores_var.iloc[j][0]      
        menores = df_menores_var.iloc[j][0]
        df_setores_maior = df_setores[df_setores.Códig_T == maiores] 
        df_setores_menor = df_setores[df_setores.Códig_T == menores]
        if df_setores_maior.empty or df_setores_menor.empty:
            print('Ativo não está na planilha de setores, adicionar')
            print(maiores)
            print(menores)
        if tipo == 'maiores':
            df_setores_final = pd.DataFrame({'Código': df_maiores_var.iloc[j][0],
                                                    'Variação diária': [str(df_maiores_var.iloc[j][2]) + '%'],
                                                   'Variação semanal' :  [str(df_maiores_var.iloc[j][3]) + '%'],
                                                   'Quantidade': df_maiores_var.iloc[j][1].translate({ord(c): None for c in "."}),
                                                     'Setor': df_setores_maior.iloc[0][5]}, index= [0])
            df_final_pdf = df_final_pdf.append(df_setores_final)

        elif tipo == 'menores':
            df_setores_final = pd.DataFrame({'Código': df_menores_var.iloc[j][0],
                                                   'Variação diária': [str(df_menores_var.iloc[j][2]) + '%'],
                                                   'Variação semanal': [str(df_menores_var.iloc[j][3]) + '%'],
                                                   'Quantidade': df_menores_var.iloc[j][1].translate({ord(c): None for c in "."}),
                                                   'Setor': df_setores_menor.iloc[0][5]}, index=[0])
            df_final_pdf = df_final_pdf.append(df_setores_final)
            
    return df_final_pdf