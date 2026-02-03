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

# configurações iniciais

nome_arquivo = r"C:\Users\izabe\Downloads\DADOS.xlsx"
aba_dados_limpos = 'dados_para_analise'
aba_original = 'Plan1'

# carregar e preparar os dados

try:
    df = pd.read_excel(nome_arquivo,sheet_name=aba_dados_limpos)
except Exception as e:
    print(f'erro ao ler aba {aba_dados_limpos} : {e}')
    print('verifique se o nome do arquivo ou da aba estão corretos')
    sys.exit()

df.columns = df.columns.astype(str).str.strip()
print(f'Colunas detectadas na aba{aba_dados_limpos}:{list(df.columns)}')

date_col = None
if 'DATA' in df.columns:
    date_col = 'DATA'
else:
    for c in df.columns:
        cl = c.lower()
        if('DATA' in cl) or ('MES' in cl) or ('MÊS' in cl) or ('MONTH' in cl) or ('DATE' in cl):
            date_col = c
            break
if not date_col:
    print(f'Erro não foi possivel encontrar uma coluna de datas na {aba_dados_limpos}')
    print('colunas encontradas', list(df.columns))
    sys.exit()

# converter no formato %m %d 
try:
    df['Data'] = pd.to_datetime(df[date_col], dayfirst=True, errors='raise')
except Exception:
    df['Data'] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    if df['Data'].isna().all():
        print(f'erro não foi possivel converter a coluna {date_col} para datas')
        print('Verifique os valores e o formato 08/2025')
        sys.exit()
    else:
        print(f'A coluna {date_col} foi convertida usando um parser flexivel')

df = df.sort_values(by='DATA')
df = df.dropna(subset=['QNT'])

# Prepara dados para o modelo sarima

df = df.set_index('DATA')
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

forecast_object = results.get_forecast(steps = 12)
previsoes_series = forecast_object.predicted_mean

previsoes_finais = []

for proxima_data,previsao_qnt in previsoes_series.items():
    previsao_qnt_arredondada = round(previsao_qnt)

    nome_mes = proxima_data.strftime('%B')

    previsoes_finais.append({
        'data' : proxima_data,
        'mes_nome': nome_mes,
        'previsao': previsao_qnt_arredondada
    })
        
    print(f"previsão para {proxima_data.strftime('%m/%Y')} : {previsao_qnt_arredondada}")

# Filtrar previsões para 2026
previsoes_2026 = [p for p in previsoes_finais if p['data'].year == 2026]

if previsoes_2026:
    # Extrair dados para o gráfico
    meses = [p['mes_nome'] for p in previsoes_2026]
    valores = [p['previsao'] for p in previsoes_2026]
    datas_str = [p['data'].strftime('%m/%Y') for p in previsoes_2026]
    
    # Criar o gráfico
    plt.figure(figsize=(12, 6))
    plt.plot(range(len(meses)), valores, marker='o', linewidth=2, markersize=8, color='#2E86AB')
    
    # Adicionar valores em cada ponto
    for i, valor in enumerate(valores):
        plt.text(i, valor + 5, str(valor), ha='center', va='bottom', fontsize=9, fontweight='bold')
    
    # Configurar rótulos e título
    plt.xlabel('Mês', fontsize=12, fontweight='bold')
    plt.ylabel('Quantidade Prevista', fontsize=12, fontweight='bold')
    plt.title('Previsões para 2026', fontsize=14, fontweight='bold')
    plt.xticks(range(len(meses)), datas_str, rotation=45, ha='right')
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(r"C:\Users\izabe\Desktop\Projeto Bernardo\previsao_anual2026.png", dpi=300, bbox_inches='tight')
    plt.show()
else:
    print("Nenhuma previsão encontrada para 2026")
    