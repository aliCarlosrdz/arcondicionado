#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct  3 14:14:08 2022

@author: alisson
"""
import pandas as pd
import funcao as fun
import dados as d


file = 'DadosClimaticosVicosa-v2.xlsx'
data = pd.read_excel(file)


#=================================CARGA TERMICA DE RENOVAÇÃO===================
#calculo da carga térmica de renovação
data['Carga Térm. de Ren. [W]']= pd.DataFrame(fun.carga_term_ren(d.pess,d.area,d.Fp,d.Fa,d.umd,data['Humidade Absoluta [kg/kg]'],d.temp,data['Temperatura de Bulbo Seco [C]']))

#calculo estatistico da carga térmica de renovação
dataEst = pd.DataFrame()
dataEst['Carga Térm. de Ren.'] = data['Carga Térm. de Ren. [W]'].describe()



data['Carga Térm. de Ren. [W]'].plot.hist(bins=15)

#=================================CARGA TERMICA DE ENVOLTÓRIA==================
materiais = pd.read_excel('Materiais.xlsx')
cons = 'Constantes.xlsx'



var2Calc = pd.DataFrame()
cargTermEnv = pd.DataFrame(columns=[[],[]])
estCargTermEnv = pd.DataFrame(columns=[[],[]])

var2Calc['Angulo Solar[graus]'] = fun.angSol(data['Hora'])
var2Calc['Declinação[graus]'] = fun.declin(data['Dia'])
var2Calc['Altitude Solar [graus]'] = fun.altSol(var2Calc['Declinação[graus]'],d.lat,var2Calc['Angulo Solar[graus]'])
var2Calc['Radiação Direta [W/m2]'] = fun.radDir(var2Calc['Altitude Solar [graus]'], cons, data['Mês'])
var2Calc['Radiação Difusa [W/m2'] = fun.radDif(var2Calc['Radiação Direta [W/m2]'], var2Calc['Declinação[graus]'],d.ro,cons,var2Calc['Altitude Solar [graus]'],data['Mês'])



for i in range(materiais.shape[0]):
    #pegando dados do material
    nm = materiais['Nome'][i]
    incli = materiais['Inclinação [graus]'][i]
    azi_sup= materiais['Azimute Solar de Superfície [graus]'][i]
    area = materiais['Área [m2]'][i]
    U = materiais['U [W/m2*K]'][i]
    alpha_Fs = materiais['Alpha/Fs'][i]
    opaca = materiais['Opaca'][i]


    var2Calc['Cosseno theta'] = fun.cosTheta(d.lat, incli, azi_sup, var2Calc['Declinação[graus]'], var2Calc['Angulo Solar[graus]'])
    cargTermEnv[nm,'Radiação Total na Superfície [W/m2]'] = fun.radTotSup(var2Calc['Cosseno theta'],var2Calc['Radiação Direta [W/m2]'],var2Calc['Radiação Difusa [W/m2'])
    cargTermEnv[nm,'Temp. Sol-Ar'] = fun.tempSolAr(data['Temperatura de Bulbo Seco [C]'], d.rse, alpha_Fs, cargTermEnv[nm,'Radiação Total na Superfície [W/m2]'], opaca)
    cargTermEnv[nm,'Temp. Sol-Ar Média 24h'] = fun.tempSolAr(data['Temperatura Média 24h [C]'], d.rse, alpha_Fs, cargTermEnv[nm,'Radiação Total na Superfície [W/m2]'], opaca)
    cargTermEnv[nm,'Carga Térmica Convectiva [W]'] = fun.cTermConv(area, U, d.f, d.temp, data['Temperatura de Bulbo Seco [C]'],cargTermEnv[nm,'Temp. Sol-Ar Média 24h'],d.lag, opaca)
    cargTermEnv[nm,'Carga Térmica Radiante [W]'] = fun.cTermRad(area, alpha_Fs, cargTermEnv[nm,'Radiação Total na Superfície [W/m2]'], U, data['Temperatura de Bulbo Seco [C]'], d.temp, opaca)
    cargTermEnv[nm,'Carga Térmica Total [W]'] = cargTermEnv[nm,'Carga Térmica Convectiva [W]'] + cargTermEnv[nm,'Carga Térmica Radiante [W]']
    #cálculo das estatíticas
   
    estCargTermEnv[nm,'Carga Térmica Convectiva [W]'] = cargTermEnv[nm,'Carga Térmica Convectiva [W]'].describe()
    estCargTermEnv[nm,'Carga Térmica Radiante [W]'] = cargTermEnv[nm,'Carga Térmica Radiante [W]'].describe()
    estCargTermEnv[nm,'Carga Térmica Total [W]'] = cargTermEnv[nm,'Carga Térmica Total [W]'].describe()
estCargTermEnv = estCargTermEnv.drop(index=('count'))
    

with pd.ExcelWriter('Carga Térmica.xlsx', engine='openpyxl') as writer:
    cargTermEnv.to_excel(writer,sheet_name='carga termica', startcol=8)
    
with pd.ExcelWriter('Carga Térmica.xlsx',mode = 'a', engine='openpyxl', if_sheet_exists='overlay') as writer:
    data.to_excel(writer,sheet_name='carga termica', index=False,startrow=1)
    # dataEst.to_excel(writer, engine='xlsxwriter')   



