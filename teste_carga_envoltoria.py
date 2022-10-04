#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Sep 25 07:21:26 2022

@author: alisson
"""

import pandas as pd
import funcao as fun
import dados as d

file = 'DadosClimaticosVicosa.xlsx'
clima = pd.read_excel(file)

#-------------------------------------------------------------------------------
materiais = pd.read_excel('Materiais.xlsx')
cons = 'Constantes.xlsx'

to_save = []
teste = pd.DataFrame()
teste['Angulo Solar[graus]'] = fun.angSol(clima['Hour'])
teste['Declinação[graus]'] = fun.declin(clima['Day'])
teste['Altitude Solar [graus]'] = fun.altSol(teste['Declinação[graus]'],d.lat,teste['Angulo Solar[graus]'])
teste['Radiação Direta [W/m2]'] = fun.radDir(teste['Altitude Solar [graus]'], cons, clima['Month'])
teste['Radiação Difusa [W/m2'] = fun.radDif(teste['Radiação Direta [W/m2]'], teste['Declinação[graus]'],d.ro,cons,teste['Altitude Solar [graus]'],clima['Month'])


for i in range(materiais.shape[0]):
    #pegando dados do material
    nm = materiais['Nome'][i]
    incli = materiais['Inclinação [graus]'][i]
    azi_sup= materiais['Azimute Solar de Superfície [graus]'][i]
    area = materiais['Área [m2]'][i]
    U = materiais['U [W/m2*K]'][i]
    alpha_Fs = materiais['Alpha/Fs'][i]
    opaca = materiais['Opaca'][i]

    teste['Cosseno theta'] = fun.cosTheta(d.lat, incli, azi_sup, teste['Declinação[graus]'], teste['Angulo Solar[graus]'])
    teste['Radiação Total na Superfície [W/m2]'] = fun.radTotSup(teste['Cosseno theta'],teste['Radiação Direta [W/m2]'],teste['Radiação Difusa [W/m2'])
    teste['Temp. Sol-Ar'] = fun.tempSolAr(clima['Dry Bulb Temperature [C]'], d.rse, alpha_Fs, teste['Radiação Total na Superfície [W/m2]'], opaca)
    teste['Temp. Sol-Ar Média 24h'] = fun.tempSolAr(clima['Mean Temperature 24h [C]'], d.rse, alpha_Fs, teste['Radiação Total na Superfície [W/m2]'], opaca)
    teste['Carga Térmica Convectiva [W]'] = fun.cTermConv(area, U, d.f, d.temp, clima['Dry Bulb Temperature [C]'],teste['Temp. Sol-Ar Média 24h'],d.lag, opaca)
    teste['Carga Térmica Radiante [W]'] = fun.cTermRad(area, alpha_Fs, teste['Radiação Total na Superfície [W/m2]'], U, clima['Dry Bulb Temperature [C]'], d.temp, opaca)
    teste['Carga Térmica Total [W]'] = teste['Carga Térmica Convectiva [W]'] + teste['Carga Térmica Radiante [W]']
    #cálculo das estatíticas
    est_carg_term = pd.DataFrame()
    est_carg_term[0] = teste['Carga Térmica Convectiva [W]'].describe()
    est_carg_term[1] = teste['Carga Térmica Radiante [W]'].describe()
    est_carg_term[2] = teste['Carga Térmica Total [W]'].describe()
    est_carg_term = est_carg_term.drop('count')
    
    #data_est['Carga Térm. de Ren'] = clima['Carga Térm. de Ren. [kW]'].describe()
    if i == 0:
        with pd.ExcelWriter('carga term envoltória teste.xlsx', engine='openpyxl') as writer:
            teste.to_excel(writer, sheet_name= nm, index=False,startrow=7)
            est_carg_term.to_excel(writer, sheet_name= nm,header=False ,startcol=8)
    else:
        with pd.ExcelWriter('carga term envoltória teste.xlsx', mode='a', engine='openpyxl',if_sheet_exists='overlay') as writer:
            teste.to_excel(writer, sheet_name= nm, index=False, startrow=7)
            est_carg_term.to_excel(writer,sheet_name=nm,header=False, startcol=8)


    
    
    
    
    



