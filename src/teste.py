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

arquivo24 = r"C:\Users\izabe\Downloads\PEDIDOS X DIAS 2024 (1).xlsx"
abas = ['AAE2', 'CANT2', 'PORT2', 'SERV2']

# Dicionário para armazenar os resultados de cada aba
resultados_abas = {}

# Processar cada aba separadamente
for aba in abas:
    print(f"\n{'='*60}")
    print(f"Processando aba: {aba}")
    print(f"{'='*60}\n")
    
    try:
        df_24 = pd.read_excel(arquivo24, sheet_name=aba)
    except Exception as e:
        print(f'erro ao ler a aba {aba}: {e}')
        continue
    
    colunas_fixas = ['DATA']
    
    # Obter índice dinâmico de B1 até a primeira coluna livre - 1
    wb = openpyxl.load_workbook(arquivo24)
    ws = wb[aba]
    
    # Encontrar primeira coluna vazia a partir de B (coluna 2)
    primeira_coluna_vazia = None
    for col_idx in range(2, ws.max_column + 2):
        if ws.cell(row=1, column=col_idx).value is None:
            primeira_coluna_vazia = col_idx
            break
    
    # Criar índice de B até primeira_coluna_vazia - 1
    if primeira_coluna_vazia:
        var_name = [ws.cell(row=1, column=col).value for col in range(2, primeira_coluna_vazia)]
    else:
        var_name = [ws.cell(row=1, column=col).value for col in range(2, ws.max_column + 1)]
    
    wb.close()
    
    df_longo = df_24.melt(id_vars=colunas_fixas, value_vars=var_name, var_name='Escola', value_name='Valor')
    df_longo['Regiao'] = df_longo['Escola'].str.extract(r'/([^/]+)/')
    resultado = df_longo.groupby('Regiao')['Valor'].sum().reset_index()
    resultado.to_excel(f'resultado_por_regiao2024_{aba}.xlsx', index=False)
    print(f"Arquivo 2024 para aba {aba} gerado.")
    
    


