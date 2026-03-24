import pandas as pd
import dateutil.relativedelta 
import numpy as np
import openpyxl
import sys
import warnings
import statsmodels.api as sm
from datetime import datetime
from statsmodels.graphics.tsaplots import plot_acf, plot_pacf
from statsmodels.tsa.statespace.sarimax import SARIMAX
from statsmodels.tools.sm_exceptions import ConvergenceWarning
import matplotlib.pyplot as plt

arquivo24 = r"C:\Users\izabe\Downloads\DADOS 2024.xlsx"
arquivo25 = r"C:\Users\izabe\Downloads\DADOS 2025.xlsx"
abas = ['AAE', 'CANT', 'PORT', 'SERV']
regiao = ['RG B', 'RG L', 'RG N', 'RG NO', 'RG NE', 'RG O', 'RG VN', 'RG CS','RG P']

# Dicionário para armazenar os resultados de cada aba
resultados_abas = {}

# Processar cada aba separadamente
for aba in abas:
    print(f"\n{'='*60}")
    print(f"Processando aba: {aba}")
    print(f"{'='*60}\n")
    
    try:
        df_24 = pd.read_excel(arquivo24, sheet_name=aba)
        df_25 = pd.read_excel(arquivo25, sheet_name=aba)
    except Exception as e:
        print(f'erro ao ler a aba {aba}: {e}')
        continue
    
    df_total = pd.concat([df_24, df_25], ignore_index=True)
    print(f'total de linhas carregadas para {aba}: {len(df_total)}')

    # converter no formato %m %d 

    df_total.columns = df_total.columns.astype(str).str.strip()

    # Manter DATA como primeiro índice e depois as colunas de região
    df_total.rename(columns={df_total.columns[0]: 'DATA'}, inplace=True)
    df_total['DATA'] = pd.to_datetime(df_total['DATA'], dayfirst=True, errors='coerce')
    df_total = df_total.dropna(subset=['DATA'])
    df_total = df_total.sort_values('DATA')

    # Dicionário para armazenar resultados das regiões
    resultados_regioes = {}

    # Processar cada região (coluna)
    for reg in regiao:
        if reg not in df_total.columns:
            print(f"  Coluna {reg} não encontrada")
            continue
        
        print(f"\n  Processando região: {reg}")
        
        # Criar série temporal para esta região
        ts = df_total.set_index('DATA')[reg].copy()
        ts = ts.dropna()
        
        # Infer frequency
        ts.index.freq = pd.infer_freq(ts.index)

        print(f'  Série temporal para {reg}: {len(ts)} observações')
        print(f'  Período: {ts.index.min()} a {ts.index.max()}')

        # Ajustar modelo SARIMA DIARIO
        try:
            model=sm.tsa.statespace.SARIMAX(
                ts, 
                order=(3,0,3), 
                seasonal_order=(2,0,2,7),
                enforce_stationarity=False,
                enforce_invertibility=False
                )
            results = None
            convergiu = False
            metodos_otimizacao = ['lbfgs', 'powell', 'nm']

            for metodo in metodos_otimizacao:
                with warnings.catch_warnings(record=True) as w:
                    warnings.simplefilter('always', ConvergenceWarning)
                    tentativa = model.fit(
                        method=metodo,
                        maxiter=500,
                        disp=False,
                        cov_type='approx'
                    )

                houve_warning_convergencia = any(
                    issubclass(item.category, ConvergenceWarning) for item in w
                )
                convergiu = bool(tentativa.mle_retvals.get('converged', False)) and not houve_warning_convergencia
                results = tentativa

                if convergiu:
                    print(f"  Modelo convergiu usando método: {metodo}")
                    break

            if not convergiu:
                print("  Aviso: sem convergência completa; usando melhor ajuste encontrado.")

            print(results.summary())
            
            # Fazer previsões para os próximos 60 dias com intervalo de confiança
            forecast = results.get_forecast(steps=365)
            forecast_values = forecast.predicted_mean
            forecast_ci = forecast.conf_int(alpha=0.05)  # 95% intervalo de confiança

            # Perfil mensal histórico (somente dias úteis) para manter o padrão observado
            ts_df = ts.to_frame(name='QNT')
            ts_df['weekday'] = ts_df.index.weekday
            ts_df['month'] = ts_df.index.month
            ts_weekdays = ts_df[ts_df['weekday'] < 5]
            monthly_mean = ts_weekdays.groupby('month')['QNT'].mean()
            overall_mean = ts_weekdays['QNT'].mean()
            month_factor = (monthly_mean / overall_mean).to_dict()

            # Usar previsões médias ajustadas pelo fator mensal
            forecast_com_variacao = []
            for mean, data_index in zip(forecast_values, forecast_values.index):
                dia_semana = data_index.weekday()
                fator_mes = month_factor.get(data_index.month, 1.0)
                
                if dia_semana in [5, 6]:  # Sábado ou Domingo
                    forecast_com_variacao.append(0)
                else:
                    # Dias da semana: usar previsão média ajustada pelo fator mensal
                    mean_adjusted = mean * fator_mes
                    forecast_com_variacao.append(max(0, round(mean_adjusted)))
            
            # Criar dataframe com as previsões
            forecast_df = pd.DataFrame({
                'Data': forecast_values.index.strftime('%d/%m/%Y'),
                'Previsão QNT': forecast_com_variacao
            })
            
            print(f"\n  Previsão para os próximos 60 dias - {aba} / {reg}:")
            print(forecast_df.to_string(index=False))
            
            # Armazenar resultado
            resultados_regioes[reg] = forecast_df
            
        except Exception as e:
            print(f'  Erro ao ajustar SARIMA para {reg}: {e}')
    
    # Armazenar resultados da aba
    resultados_abas[aba] = resultados_regioes

print(f"\n{'='*60}")
print("Processamento concluído!")
print(f"Total de abas processadas: {len(resultados_abas)}")
print(f"{'='*60}")

# Salvar resultados em planilha Excel
arquivo_saida = r"C:\Users\izabe\Desktop\Projeto Bernardo\excel\PrevisoesPedidos(3,0,3).xlsx"

with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    for aba, resultados_regioes in resultados_abas.items():
        # Consolidar dados de todas as regiões
        dfs_regioes = {}
        for reg, forecast_df in resultados_regioes.items():
            # Usar 'Data' como índice temporário e extrair a coluna de previsão
            dfs_regioes[reg] = forecast_df.set_index('Data')['Previsão QNT']
        
        # Combinar em um único DataFrame com datas na coluna A
        if dfs_regioes:
            consolidated = pd.DataFrame(dfs_regioes)
            consolidated = consolidated.reset_index()
            
            # Escrever na aba correspondente
            consolidated.to_excel(writer, sheet_name=aba, index=False)
            print(f"Aba '{aba}' salva com {len(dfs_regioes)} regiões")