#!/usr/bin/env python
# coding: utf-8

# In[67]:


get_ipython().system('pip install yfinance')


# In[2]:


get_ipython().system('pip install mplcyberpunk')


# In[3]:


get_ipython().system('pip install pywin32')


# In[ ]:





# In[19]:


import pandas as pd
import datetime
import yfinance as yf
from matplotlib import pyplot as plt
import mplcyberpunk
import win32com.client as win32


# In[106]:


codigos_de_negociacao = ["^BVSP", "BRL=X"]
data_hoje = datetime.datetime.now()
um_ano_atras = hoje - datetime.timedelta(days = 365)

dados_download = yf.download(codigos_de_negociacao, um_ano_atras, data_hoje)

display(dados_download)


# In[115]:


dados_fechamento = dados_download['Adj Close']
dados_fechamento.columns = ['dolar', 'ibovespa']
#dados_fechamento = dados_download.dropna()

dados_fechamento.dropna()


# In[ ]:





# In[116]:


dados_anuais = dados_fechamento.resample("Y").last()
dados_mensais = dados_fechamento.resample("M").last()


dados_mensais


# In[117]:


retorno_anual = dados_anuais.pct_change().dropna()
retorno_mensal = dados_mensais.pct_change().dropna()
retorno_diario = dados_fechamento.pct_change().dropna()

retorno_diario


# In[140]:


retorno_diario_dolar = retorno_diario.iloc[-1, 0]
retorno_diario_ibov = retorno_diario.iloc[-1, 1]

retorno_mensal_dolar = retorno_mensal.iloc[-1, 0]
retorno_mensal_ibov = retorno_mensal.iloc[-1, 1]

retorno_anual_dolar = retorno_anual.iloc[-1, 0]
retorno_anual_ibov = retorno_anual.iloc[-1, 1]


print(retorno_anual_ibov)


# In[141]:


retorno_diario_dolar = round((retorno_diario_dolar * 100), 2)
retorno_diario_ibov = round((retorno_diario_dolar * 100), 2)

retorno_mensal_dolar = round((retorno_mensal_dolar * 100), 2)
retorno_mensal_ibov = round((retorno_mensal_ibov * 100), 2)

retorno_anual_dolar = round((retorno_anual_dolar * 100), 2)
retorno_anual_ibov = round((retorno_anual_ibov * 100), 2)


# In[142]:


plt.style.use("cyberpunk")

dados_fechamento.plot(y = "ibovespa", use_index = True, legend = False)
plt.title("Ibovespa")
plt.savefig('ibovespa.png', dpi = 300)

plt.show()


# In[143]:


plt.style.use("cyberpunk")

dados_fechamento.plot(y = "dolar", use_index = True, legend = False)
plt.title("Dólar")
plt.savefig('dolar.png', dpi = 300)

plt.show()


# In[147]:


outlook = win32.Dispatch("outlook.application")

email = outlook.CreateItem(0)


# In[148]:


email.To = "brenno@varos.com.br"
email.Subject = "Relatório Diário"
email.Body = f'''Prezado Diretor, segue o relatório diário:

Bolsa:

No ano o Ibovespa está tendo uma rentabilidade de {retorno_anual_ibov}%, 
enquanto no mês a rentabilidade é de {retorno_mensal_ibov}%.

No último dia útil, o fechamento do Ibovespa foi de {retorno_diario_ibov}%.

Dólar:

No ano o Dólar está tendo uma rentabilidade de {retorno_anual_dolar}%, 
enquanto no mês a rentabilidade é de {retorno_mensal_dolar}%.

No último dia útil, o fechamento do Dólar foi de {retorno_diario_dolar}%.


Abs,

O melhor estagiário do mundo.
Geislân Gabriel, relatório certo haha

'''

anexo_ibovespa = r'C:\Users\Gabriel\ibovespa.png'
anexo_dolar = r'C:\Users\Gabriel\dolar.png'

email.Attachments.Add(anexo_ibovespa)
email.Attachments.Add(anexo_dolar)

email.Send()

