import pandas as pd
import numpy as np
import openpyxl
import sys
import warnings
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from prophet import Prophet

warnings.filterwarnings('ignore')

arquivo24 = r"C:\Users\izabe\Downloads\DADOS 2024.xlsx"
arquivo25 = r"C:\Users\izabe\Downloads\DADOS 2025.xlsx"
abas = ['AAE', 'CANT', 'PORT', 'SERV']
regiao = ['RG B', 'RG L', 'RG N', 'RG NO', 'RG NE', 'RG O', 'RG VN', 'RG CS','RG P']

resultados_abas = {}

def previsao_prophet(ts, steps=365):
    """Usar Prophet para previsão"""
    try:
        df_prophet = pd.DataFrame({'ds': ts.index, 'y': ts.values})
        df_prophet['ds'] = pd.to_datetime(df_prophet['ds'], errors='coerce')
        df_prophet['y'] = pd.to_numeric(df_prophet['y'], errors='coerce')
        df_prophet = df_prophet.dropna(subset=['ds', 'y'])
        df_prophet = (
            df_prophet.groupby('ds', as_index=False)['y']
            .sum()
            .sort_values('ds')
        )
        df_prophet['y'] = np.clip(df_prophet['y'], a_min=0, a_max=None)
        
        if len(df_prophet) < 30:
            print("      ✗ Dados insuficientes para Prophet")
            return None, None
        
        model = Prophet(
            yearly_seasonality=True,
            weekly_seasonality=True,
            daily_seasonality=False,
            interval_width=0.95
        )
        
        print("      → Treinando Prophet...")
        try:
            model.fit(df_prophet, algorithm='lbfgs')
        except Exception:
            print("      ⚠ lbfgs falhou, tentando newton...")
            model.fit(df_prophet, algorithm='newton')
        
        future = model.make_future_dataframe(periods=steps, freq='D')
        forecast = model.predict(future)
        
        # Retornar apenas os valores previstos futuros
        forecast_future = forecast[forecast['ds'] > ts.index.max()][['ds', 'yhat']].reset_index(drop=True)
        return forecast_future['yhat'].values, forecast_future['ds'].values
    except Exception as e:
        print(f"      ✗ Erro em Prophet: {e}")
        return None, None

def ajustar_para_fds_e_sazonalidade(forecast, dates, ts_original):
    """Ajustar previsões para fins de semana e sazonalidade mensal"""
    forecast_ajustado = forecast.copy()
    
    # Fator mensal
    ts_df = ts_original.to_frame(name='QNT')
    ts_df['weekday'] = ts_df.index.weekday
    ts_df['month'] = ts_df.index.month
    
    # Usar apenas dias úteis para calcular sazonalidade
    ts_dias_uteis = ts_df[ts_df['weekday'] < 5]
    monthly_mean = ts_dias_uteis.groupby('month')['QNT'].mean()
    overall_mean = ts_dias_uteis['QNT'].mean()
    month_factor = (monthly_mean / overall_mean).to_dict()
    
    # Aplicar ajustes
    for i, data_idx in enumerate(dates):
        dia_semana = data_idx.dayofweek
        fator_mes = month_factor.get(data_idx.month, 1.0)
        
        if dia_semana in [5, 6]:  # Sábado ou domingo
            forecast_ajustado[i] = 0
        else:
            forecast_ajustado[i] = max(0, forecast_ajustado[i] * fator_mes)
    
    return forecast_ajustado

# ==================== PROCESSAMENTO ====================

for aba in abas:
    print(f"\n{'='*70}")
    print(f"Processando aba: {aba}")
    print(f"{'='*70}\n")
    
    try:
        df_24 = pd.read_excel(arquivo24, sheet_name=aba)
        df_25 = pd.read_excel(arquivo25, sheet_name=aba)
    except Exception as e:
        print(f'Erro ao ler a aba {aba}: {e}')
        continue
    
    df_total = pd.concat([df_24, df_25], ignore_index=True)
    print(f'Total de linhas carregadas para {aba}: {len(df_total)}')

    df_total.columns = df_total.columns.astype(str).str.strip()
    df_total.rename(columns={df_total.columns[0]: 'DATA'}, inplace=True)
    df_total['DATA'] = pd.to_datetime(df_total['DATA'], dayfirst=True, errors='coerce')
    df_total = df_total.dropna(subset=['DATA'])
    df_total = df_total.sort_values('DATA')

    resultados_regioes = {}

    # Processar cada região
    for reg in regiao:
        if reg not in df_total.columns:
            print(f"  Coluna {reg} não encontrada")
            continue
        
        print(f"\n  Processando região: {reg}")
        
        ts = df_total.set_index('DATA')[reg].copy()
        ts = ts.dropna()
        
        if len(ts) < 30:
            print(f"    ⚠ Dados insuficientes ({len(ts)} observações)")
            continue
        
        print(f"    Série temporal: {len(ts)} observações")
        print(f"    Período: {ts.index.min().date()} a {ts.index.max().date()}")
        print(f"    Média diária: {ts.mean():.2f} | Máximo: {ts.max():.0f} | Mínimo: {ts.min():.0f}")
        
        # ===== GERAR PREVISÃO COM PROPHET =====
        steps = 365
        
        print("    → Executando Prophet...")
        forecast_prophet, dates_forecast = previsao_prophet(ts, steps=steps)
        
        if forecast_prophet is None:
            print(f"    ✗ Prophet falhou para {reg}")
            continue
        
        print(f"      ✓ Prophet: Média = {forecast_prophet.mean():.2f}")
        
        # ===== AJUSTAR PARA FINS DE SEMANA E SAZONALIDADE =====
        forecast_ajustado = ajustar_para_fds_e_sazonalidade(forecast_prophet, dates_forecast, ts)
        
        # ===== CRIAR DATAFRAME DE PREVISÃO =====
        forecast_df = pd.DataFrame({
            'Data': dates_forecast.strftime('%d/%m/%Y'),
            'Previsão QNT': forecast_ajustado.astype(int)
        })
        
        # Mostrar primeiras 10 previsões
        print(f"\n    Primeiras 10 previsões:")
        print(forecast_df.head(10).to_string(index=False))
        
        resultados_regioes[reg] = forecast_df
    
    resultados_abas[aba] = resultados_regioes

print(f"\n{'='*70}")
print("Processamento concluído!")
print(f"{'='*70}")

# ===== SALVAR RESULTADOS =====
arquivo_saida = r"C:\Users\izabe\Desktop\Projeto Bernardo\excel\PrevisoesPedidos_Prophet.xlsx"

with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    for aba, resultados_regioes in resultados_abas.items():
        dfs_regioes = {}
        for reg, forecast_df in resultados_regioes.items():
            dfs_regioes[reg] = forecast_df.set_index('Data')['Previsão QNT']
        
        if dfs_regioes:
            consolidated = pd.DataFrame(dfs_regioes)
            consolidated = consolidated.reset_index()
            consolidated.to_excel(writer, sheet_name=aba, index=False)
            print(f"Aba '{aba}' salva com {len(dfs_regioes)} regiões")

print(f"\n✓ Arquivo salvo em: {arquivo_saida}")
