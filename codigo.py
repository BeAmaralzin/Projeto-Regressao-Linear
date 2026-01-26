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

df_total = df_total.iloc[:,[0,1]]
df_total.columns = ['DATA', 'QNT']

df_total['DATA'] = pd.to_datetime(df_total['DATA'], dayfirst=True, errors='coerce')
df_total = df_total.dropna(subset=['DATA', 'QNT'])

# Agregar múltiplas observações no mesmo dia
df_total = df_total.sort_values('DATA')
df_total = df_total.groupby('DATA')['QNT'].sum().reset_index()

# Criar série temporal com índice de data
ts = df_total.set_index('DATA')['QNT']
# Infer frequency instead of forcing it (handles gaps in data)
ts.index.freq = pd.infer_freq(ts.index)

print(f'Série temporal: {len(ts)} observações')
print(f'Período: {ts.index.min()} a {ts.index.max()}')

# Ajustar modelo SARIMA DIARIO
try:
    model=sm.tsa.statespace.SARIMAX(
        ts, 
        order=(1,1,1), 
        seasonal_order=(1,1,1,7),
        enforce_stationarity=False,
        enforce_invertibility=False
        )
    results = model.fit(disp=False, cov_type='approx')
    print(results.summary())
    
    # Fazer previsões para os próximos 30 dias
    forecast = results.get_forecast(steps=60)
    forecast_values = forecast.predicted_mean
    
    # Criar dataframe simples com data e previsão arredondada (sem negativos)
    forecast_df = pd.DataFrame({
        'Data': forecast_values.index.strftime('%d/%m/%Y'),
        'Previsão QNT': forecast_values.round(0).clip(lower=0).astype(int)
    })
    
    print("\nPrevisão para os próximos 30 dias:")
    print(forecast_df.to_string(index=False))
    
    # Criar gráfico de linha com matplotlib
    plt.figure(figsize=(14, 6))
    forecast_df['Data'] = pd.to_datetime(forecast_df['Data'], format='%d/%m/%Y')
    plt.plot(forecast_df['Data'], forecast_df['Previsão QNT'], marker='o', linestyle='-', linewidth=2, markersize=6, color='#2E86AB')
    plt.title('Previsão de Quantidade - Próximos 60 Dias', fontsize=16, fontweight='bold')
    plt.xlabel('Data', fontsize=12, fontweight='bold')
    plt.ylabel('Quantidade (QNT)', fontsize=12, fontweight='bold')
    plt.grid(True, alpha=0.3)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(r"C:\Users\izabe\Desktop\Projeto Bernardo\previsao_30_dias.png", dpi=300, bbox_inches='tight')
    plt.show()
    
except Exception as e:
    print(f'Erro ao ajustar SARIMA: {e}')

