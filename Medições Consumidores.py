#!/usr/bin/env python
# coding: utf-8

# In[4]:


import pandas as pd
from pandas import ExcelWriter

caminho_origem = 'S:'
nome_planilha = '\Resumo jan-23.xlsx'
nome_aba = '01.23 Resumo'
caminho_destino = 'S:\Planilhas Cons. Sem Destino'
mes_ano = '01.23 '

df=pd.read_excel(caminho_origem+nome_planilha,nome_aba,header=3)
df['Data']=df['Data'].dt.strftime("%d/%m/%Y")
df.insert(loc=4, column='Patamar de Carga', value=0)
pontos_medicao=pd.unique(df['Ponto / Grupo'])
data=pd.unique(df['Data'])
for x in pontos_medicao:
    df_individual=df[df['Ponto / Grupo']==x]
    consumo_mensal_mwm=df_individual['Ativa C (kWh)'].sum()/df_individual["Data"].count()
    df_individual=df_individual.groupby(by=['Data']).sum()
    df_individual.insert(0,'Data',data)
    df_individual.drop(columns=['Ativa G (kWh)','Reativa C (kVArh)','Reativa G (kVArh)', 'Patamar de Carga', 'Hora'],inplace=True)
    df_individual['Consumo [MWm]']=df_individual['Ativa C (kWh)']/24
    df_individual['Consumo Médio [MWm]']=consumo_mensal_mwm
    nome_aba=x.replace("(L)","")
    try:
        nome_planilha_individual=df[df['Ponto / Grupo']==x]['Nome'].iloc[0]+'.xlsx'
    except:
        nome_planilha_individual_alternativo=df[df['Ponto / Grupo']==x]['Agente'].iloc[0]+' '+df[df['Ponto / Grupo']==x]['Ponto / Grupo'].iloc[0]+'.xlsx'
    salvar=df[df['Ponto / Grupo']==x]['Destino'].iloc[0]
    try:
        with pd.ExcelWriter(salvar+mes_ano+nome_planilha_individual) as writer:
            df[df['Ponto / Grupo']==x].drop(columns=['Agente','Ponto / Grupo','Destino','Nome']).to_excel(writer,nome_aba,index=False,startrow=5)
            df_individual.to_excel(writer,'Gráfico',index=False,startcol=1,startrow=1)
    except:
        print(df[df['Ponto / Grupo']==x]['Agente'].iloc[0]+" Ponto: "+df[df['Ponto / Grupo']==x]['Ponto / Grupo'].iloc[0])


# In[ ]:





# In[ ]:




