import pandas as pd
import datetime as dt
lista = ['12/05/2022']

for i in range(len(lista)):
    Diacompleto = lista[i]
    Dia = Diacompleto[:2]
    Mes = Diacompleto[3:5]
    Ano = Diacompleto[6:]
    Hoje = str(dt.date(int(Ano),int(Mes),int(Dia)))
    url = 'http://www.b3.com.br/pt_br/market-data-e-indices/servicos-de-dados/market-data/consultas/mercado-a-vista/termo/posicoes-em-aberto/posicoes-em-aberto-8AA8D0CC77D179750177DF167F150965.htm?data=' + Diacompleto +'&f=0'

    df = pd.read_html(url)

    df_valores = df[0]
    df_valores2 = df[1]

    df_valores.to_excel('R:/Aluguel/Desenvolvimento/Rotina Relatorio/Relatorio_Diario_Termo/PosicaoTermo/posicao_termo_'+Ano +  Mes  + Dia + '.xlsx', index = False)
    df_valores2.to_excel('R:/Aluguel/Desenvolvimento/Rotina Relatorio/Relatorio_Diario_Termo/PosicaoTermo/posicao_aluguel_'+Ano +  Mes + Dia + '.xlsx', index = False)

