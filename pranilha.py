import pandas as pd
import funcao as fun
import dados as d

file = 'DadosClimaticosVicosa.xlsx'
clima = pd.read_excel(file)

#calculo da carga térmica de renovação
clima['Carga Térm. de Ren. [kW]']= pd.DataFrame(fun.carga_term_ren(d.pess,d.area,d.Fp,d.Fa,d.umd,clima['Absolute Humidity [kg/kg]'],d.temp,clima['Dry Bulb Temperature [C]']))


# calculo estatistico da carga térmica de renovação
data_est = pd.DataFrame()
data_est['Carga Térm. de Ren'] = clima['Carga Térm. de Ren. [kW]'].describe()


# Cria planilha com resultado
with pd.ExcelWriter('RESULTADO.xlsx') as writer:
    clima.to_excel(writer, sheet_name='teste1', index=False, engine='xlsxwriter')
    data_est.to_excel(writer, sheet_name='estatistica', engine='xlsxwriter')

clima['Carga Térm. de Ren. [kW]'].plot.hist(bins=15)

#==============================================================================

