import pandas as pd
import dateutil.relativedelta 
import numpy as np
import openpyxl
import sys
import statsmodels.api as sm
from datetime import datetime

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

# perguntar ao usuario a data que ele quer prever

entrada_usuario = input("Digite o mês e o ano que deseja prever ex: 08/2026: ")
    
try:
    data_alvo = datetime.strptime(entrada_usuario,"%m/%Y")
except ValueError:
    print("Erro: Formato invalido")
    sys.exit()

ultima_data = pd.to_datetime(y.index[-1])

num_previsoes = (data_alvo.year - ultima_data.year) * 12 + (data_alvo.month - ultima_data.month)

forecast_object = results.get_forecast(steps = num_previsoes)
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

    # Escrever os resultados na planilha

    aba_alvo = 'graficos'
    coluna_alvo = 5
    linha_inicial = 11

    try:
        # load the workbook from the configured file path
        workbook = openpyxl.load_workbook(nome_arquivo)

        if aba_alvo in workbook.sheetnames:
            sheet = workbook[aba_alvo]
        else:
            print(f"Aviso Aba {aba_alvo} não encontrada")
            sheet = workbook.create_sheet(title=aba_alvo)

        for i, previsao in enumerate(previsoes_finais):

            linha_atual = linha_inicial + i

            valor_previsao = previsao['previsao']
            mes_nome = previsao['mes_nome']
            data_str = previsao['data'].strftime('%m/%Y')

            sheet.cell(row=linha_atual, column=coluna_alvo).value = valor_previsao

       
        workbook.save(nome_arquivo)

    except Exception as e:
        print("ocorreu um erro, as planilhas não foram salvas")
