# -*- coding: utf-8 -*-
"""
Created on Sat Feb 20 16:10:38 2021

Script para comparar dois arquivos excel e encontrar divergências

-GMUD não foi informada no Topdesk quando chamado está com status "Pendente Deploy - PRD"
-ID da tarefa não informado no campo "Número externo" do Topdesk e o atendimento está sendo realizado
-Quando status entre os dois sistemas estão desatualizados

@author: LDT
"""

import pandas as pd
import numpy as np
import math

#Retorna Número do Incidente que aparece entre conchetes []
def buscar_num_incidente(titulo):
    return titulo[titulo.find('[')+1:titulo.find(']')]

#Verificar números dos status são iguais ou não
def comparar_status(status_t2s, status_btp):
    if(status_t2s == status_btp):
        return 0
    else:
        return 1 
    
#Verifica se existe a palavra GMUD dentro do campo "Ação" sempre que o status = 4 "#Selecionando quais colunas serão adicionadas no arquivo final"
def buscar_gmud(acao, status_num):
    if(acao != acao):
        return 0
    if(status_num != 4):
        return 0
    acao = acao.upper()
    if(acao.find('GMUD') == -1):
        return 1 #divergente (não encontrou o termo 'GMUD' quando a etapa era "Pendente Deploy - PRD")
    else:
        return 0

#Verifica se campo "Número externo" está preenchido
def buscar_id_t2s(num_externo):
    if(num_externo != num_externo):
        return 1
    else:
        return 0

#Lendo arquivos excel
df_btp = pd.read_excel('C:\\Users\\LDT\\Desktop\\btp.xlsx')
df_t2s = pd.read_excel('C:\\Users\\LDT\\Desktop\\t2s.xlsx')

#Criando uma cópia das colunas STATUS e ETAPA
df_btp['Status_num'] = df_btp['Status']
df_t2s['Etapa_num'] = df_t2s['Etapa']

#Atribuindo números para cada STATUS, equivalente ao campo ETAPA, para facilitar comparação
df_btp['Status_num'] = df_btp['Status_num'].replace(to_replace=['Registrado'],value=0)
df_btp['Status_num'] = df_btp['Status_num'].replace(to_replace=['Em analise'],value=1)
df_btp['Status_num'] = df_btp['Status_num'].replace(to_replace=['Alterado pelo solicitante','Em atendimento'],value=2)
df_btp['Status_num'] = df_btp['Status_num'].replace(to_replace=['Aguardando testes do solicitante','Aguardando solicitante'],value=3)
df_btp['Status_num'] = df_btp['Status_num'].replace(to_replace=['Aguardando 2º Nível'],value=4)
df_btp['Status_num'] = df_btp['Status_num'].replace(to_replace=['Aguardando fornecedor'],value=9) #usar 9 para outros status não utilizados

#Atribuindo números para cada ETAPA, equivalente ao campo STATUS, para facilitar comparação
df_t2s['Etapa_num'] = df_t2s['Etapa_num'].replace(to_replace=['0. Na fila...'],value=0)
df_t2s['Etapa_num'] = df_t2s['Etapa_num'].replace(to_replace=['2. Análise'],value=1)
df_t2s['Etapa_num'] = df_t2s['Etapa_num'].replace(to_replace=['3. Desenvolvimento','4. Testes','5. Pendente Deploy - QA'],value=2)
df_t2s['Etapa_num'] = df_t2s['Etapa_num'].replace(to_replace=['6. Homologação'],value=3)
df_t2s['Etapa_num'] = df_t2s['Etapa_num'].replace(to_replace=['7. Pendente Deploy - PRD'],value=4)

#Isolando o campo "Número do incidente", presente no campo TÍTULO, em uma coluna separada para fazer o filtro
df_t2s['Número do incidente'] = df_t2s.apply(lambda row: buscar_num_incidente(row['Título']), axis=1)

#Juntar o conteúdo dos dois arquivos excel, filtrando pelo campo "Número do incidente". Adicionar apenas se o número existir nos dois arquivos excel.
df_merge = pd.merge(df_t2s,df_btp,on='Número do incidente')

#Criando campos de controle de divergências:
#DIF_STATUS = 1 -> Quando campos status/etapa não são equivalentes
#DIF_GMUD = 1 -> Quando não tem GMUD informada nos comentários e etapa for "Pendente Deploy - PRD"
#DIF_ID = 1 -> Campo "Número externo" não está preenchido

df_merge['DIF_STATUS'] = df_merge.apply(lambda row: comparar_status(row['Etapa_num'],row['Status_num']), axis=1)
df_merge['DIF_GMUD'] = df_merge.apply(lambda row: buscar_gmud(row['Ação'],row['Status_num']), axis=1)
df_merge['DIF_ID_T2S'] = df_merge.apply(lambda row: buscar_id_t2s(row['Número externo']), axis=1)

#Filtrar apenas resultados que possuem alguma divergência
df_dif = df_merge[df_merge['DIF_STATUS'] + df_merge['DIF_GMUD'] + df_merge['DIF_ID_T2S'] >=1]

#Selecionando quais colunas serão adicionadas no arquivo final
df_dif = df_dif[['Cód','WBS ','Título','Etapa','Aguardando?','Etapa_num','Número do incidente','Nível','Breve descrição (Detalhes)','Nome do solicitante','Tipo de incidente','Status','Operador','Categoria_y','Subcategoria','Dia/hora da criação','Data de fechamento','Número externo','Pedido','Ação','Status_num','DIF_STATUS','DIF_GMUD','DIF_ID_T2S']]

#Criando arquivo final
df_dif.to_excel('C:\\Users\\LDT\\Desktop\\portal_cliente_dif.xlsx',index=False)

