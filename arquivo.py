import pandas as pd
import dateutil.relativedelta 
import numpy as np
import openpyxl
import sys
import statsmodels.api as sm
from datetime import datetime

# Carregar os dados e  mesclar as planilhas

try:
    df_24 = pd.read_excel(r'"C:\Users\izabe\Downloads\PEDIDOS X DIAS 2024.xlsx"')
    df_25 = pd.read_excel(r'"C:\Users\izabe\Downloads\PEDIDOS X DIAS 2025.xlsx"')
except Exception as e:
    print(f'erro ao ler os arquivos de dados: {e}')
    sys.exit()

df_total = pd.concat([df_24, df_25], ignore_index=True)
print(f'total de linhas carregadas: {len(df_total)}')

df_total.columns = df_total.columns.astype(str).str.strip()

df_total =  df_total.iloc[:, :2]
df_total.columns = ['DATA', 'QNT']

df_total['DATA'] = pd.to_datetime(df_total['DATA'], dayfirst=True, errors='coerce')
df_total = df_total.dropna(subset=['DATA', 'QNT'])

df_menssal = df_total.set_index('DATA').resample('MS')['QNT'].sum()


