import pandas as pd
import dateutil.relativedelta 
import numpy as np
import openpyxl
import sys
import statsmodels.api as sm
from datetime import datetime
from statsmodels.graphics.tsaplots import plot_acf, plot_pacf
from statsmodels.tsa.statespace.sarimax import SARIMAX
import matplotlib.pyplot as plt

arquivo24 = r"C:\Users\izabe\Downloads\PEDIDOS X DIAS 2024.xlsx"
arquivo25 = r"C:\Users\izabe\Downloads\PEDIDOS X DIAS 2025.xlsx"
aba_dados_limpos = 'dados_para_analise'

# Carregar os dados e mesclar as planilhas

try:
    df_24 = pd.read_excel(arquivo24,sheet_name=aba_dados_limpos)
    df_25 = pd.read_excel(arquivo25,sheet_name=aba_dados_limpos)
except Exception as e:
    print(f'erro ao ler os arquivos de dados: {e}')
    sys.exit()

df_total = pd.concat([df_24, df_25], ignore_index=True)
print(f'total de linhas carregadas: {len(df_total)}')

# converter no formato %m %d 

df_total.columns = df_total.columns.astype(str).str.strip()
print(f'Colunas disponíveis: {df_total.columns.tolist()}')

# Converter para datetime - usar primeira coluna como data
date_col = df_total.columns[0]  # Primeira coluna
qty_col = df_total.columns[1]   # Segunda coluna

try:
    df_total[date_col] = pd.to_datetime(df_total[date_col], dayfirst=True, errors='raise')
    print(f'Data convertida com sucesso da coluna: {date_col}')
except Exception:
    try:
        df_total[date_col] = pd.to_datetime(df_total[date_col], dayfirst=True, errors='coerce')
        if df_total[date_col].isna().all():
            print(f'Erro: não foi possível converter coluna {date_col} para datas')
            print('Verifique o formato (ex: 08/2025 ou 01/01/2024)')
            sys.exit()
        else:
            print(f'Coluna {date_col} convertida com parser flexível')
    except Exception as e:
        print(f'Erro na conversão: {e}')
        sys.exit()

# Renomear colunas para padronização
df_total.columns = ['DATA', 'QNT']

# Limpar dados
df_total = df_total.sort_values(by='DATA')
df_total = df_total.dropna(subset=['DATA', 'QNT'])

print(f'Linhas após limpeza: {len(df_total)}')
print(f'Período: {df_total["DATA"].min()} a {df_total["DATA"].max()}')

# Formato personalizado
data_formatada = datetime.now().strftime('%m/%Y')
print(f'Data formatada: {data_formatada}')

# Prepara dados para o modelo sarima

df = df_total.set_index('DATA')
y = df['QNT']

#SAZONALIDADE

try:
    model=sm.tsa.statespace.SARIMAX(
        y,
        order=(1,1,1),
        seasonal_order=(1,1,1,12),
        enforce_stationarity=False,
        enforce_invertibility=False
    )
    
    results = model.fit(disp=False)
    
    print('Modelo trinado com sucesso')
except Exception as e:
    print('erro ao treinar oo modelo SARIMA')
    print('verifique que a dados suficientes para analise')
    sys.exit()


