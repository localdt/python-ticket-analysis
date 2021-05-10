# -*- coding: utf-8 -*-
"""
Created on Fri Nov 13 13:11:22 2020
Script para gerar base de chamados usada no gerenciamento dos chamados e análise de métricas ágeis

@author: LDT
"""
import pandas as pd


def dia(dt) :
    dt = str(dt)
    if len(dt) < 10:
        return '-'
    return dt[0:10]

dataf3 = pd.read_excel('C:\\Users\\LDT\\Desktop\\chamados-19-02-2021.xlsx')

dataf3['DATA_ABERTURA'] = dataf3.apply (lambda row: dia(row['Dia/hora da criação']), axis=1)
dataf3['DATA_FECHAMENTO'] = dataf3.apply (lambda row: dia(row['Data de fechamento']), axis=1)

aberts = dataf3['DATA_ABERTURA']
aberts = pd.DataFrame(aberts)
fecs = dataf3['DATA_FECHAMENTO']
fecs = pd.DataFrame(fecs)

aberts = aberts.rename(columns={'DATA_ABERTURA':'DATA'})
fecs = fecs.rename(columns={'DATA_FECHAMENTO':'DATA'})

dataf3r = pd.concat([aberts,fecs])
dataf3r = dataf3r.groupby(['DATA']).size().reset_index().rename(columns={0:'count'})
dataf3r = dataf3r.loc[dataf3r['DATA']!='-']

    
dataf3r = dataf3r.sort_values(by=['DATA'])

variacao_abert = []
variacao_fec = []

for c in dataf3r['DATA'].values:
    chams = dataf3.loc[dataf3['DATA_ABERTURA']==c]
    fecs2 = dataf3.loc[dataf3['DATA_FECHAMENTO'] ==c]
    variacao_abert.append(len(chams))
    variacao_fec.append(len(fecs2))
    
dataf3r['ABERTOS'] = variacao_abert
dataf3r['FECHADOS'] = variacao_fec

variacao = []
count = -1
for c in dataf3r.values:
    abert = c[2]
    fec = c[3]    
    variacao.append(abert-fec)
       

dataf3r['VARIACAO'] = variacao

acumulado = []
count = 0
for c in dataf3r['FECHADOS'].values:
    count = count + c
    acumulado.append(count)
    
dataf3r['FECHADOS_ACUMULADO'] = acumulado

acumulado = []
count = 0
for c in dataf3r['ABERTOS'].values:
    count = count + c
    acumulado.append(count)
    
dataf3r['ABERTOS_ACUMULADO'] = acumulado

acumulado = []
count = 0
for c in dataf3r['VARIACAO'].values:
    count = count + c
    acumulado.append(count)
    
dataf3r['VARIACAO_ACUMULADO'] = acumulado

dataf3r.to_excel('C:\\Users\\LDT\\Downloads\\resultados_anual-2021-2020-12-02-2021.xlsx')
