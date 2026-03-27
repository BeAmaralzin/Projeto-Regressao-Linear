import numpy as np
import pandas as pd
from prophet import Prophet


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

    # Converte DATA de dd/mm/yyyy para datetime e padroniza visualmente como yyyy/mm/dd.
    # Prophet exige `ds` em datetime; o formato string é apenas para consistência de leitura.
    data_original = df_total['DATA']
    df_total['DATA'] = pd.to_datetime(data_original, format='%d/%m/%Y', errors='coerce')
    faltantes = df_total['DATA'].isna()
    if faltantes.any():
        # Fallback para valores que não vieram exatamente em dd/mm/yyyy (ex.: datetime do Excel).
        df_total.loc[faltantes, 'DATA'] = pd.to_datetime(data_original[faltantes], errors='coerce', dayfirst=True)
    df_total['DATA'] = pd.to_datetime(df_total['DATA'].dt.strftime('%Y/%m/%d'), format='%Y/%m/%d', errors='coerce')
    datas_invalidas = df_total['DATA'].isna().sum()
    if datas_invalidas:
        print(f"  {datas_invalidas} linhas com DATA invalida foram descartadas")
        df_total = df_total.dropna(subset=['DATA'])


    # Dicionário para armazenar resultados das regiões
    resultados_regioes = {}

    # Processar cada região (coluna)
    for reg in regiao:
        if reg not in df_total.columns:
            print(f"  Coluna {reg} não encontrada")
            continue
        
        print(f"\n  Processando região: {reg}")
        
        # Preparar dados da região atual para Prophet (y separado por cada item de `regiao`).
        df_prophet = df_total[['DATA', reg]].rename(columns={'DATA': 'ds', reg: 'y'}).copy()
        df_prophet['y'] = pd.to_numeric(df_prophet['y'], errors='coerce')
        df_prophet = df_prophet.replace([np.inf, -np.inf], np.nan).dropna(subset=['ds', 'y'])
        df_prophet = df_prophet.groupby('ds', as_index=False, sort=True)['y'].sum()

        
        # Ajustar modelo Prophet
        model = Prophet(yearly_seasonality=True, weekly_seasonality=False, daily_seasonality=False)
        try:
            model.fit(df_prophet)
        except RuntimeError as e:
            print(f"  Falha no LBFGS para {reg}: {e}")
            print("  Tentando novamente com algorithm='Newton'...")
            try:
                model.fit(df_prophet, algorithm='Newton')
            except Exception as e_newton:
                print(f"  Falha no ajuste para {reg} com Newton: {e_newton}")
                continue
        except Exception as e:
            print(f"  Erro ao ajustar Prophet para {reg}: {e}")
            continue
        
        # Criar DataFrame de datas futuras (próximos 365 dias)
        future = model.make_future_dataframe(periods=365)
        
        # Gerar previsões
        forecast = model.predict(future)
        
        # Armazenar previsões em DataFrame
        forecast_df = forecast[['ds', 'yhat']].rename(columns={'ds': 'Data', 'yhat': 'Previsao_QNT'})
        forecast_df['Data'] = pd.to_datetime(forecast_df['Data']).dt.strftime('%Y/%m/%d')
        
        resultados_regioes[reg] = forecast_df
    if resultados_regioes:
        resultados_abas[aba] = resultados_regioes
    else:
        print(f"Nenhuma regiao valida para salvar na aba '{aba}'.")

    # Salvar resultados em planilha Excel
arquivo_saida = r"C:\Users\izabe\Desktop\Projeto Bernardo\excel\PrevisoesPedidos(prophet).xlsx"
with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    for aba, resultados_regioes in resultados_abas.items():
        # Consolidar dados de todas as regiões
        dfs_regioes = {}
        for reg, forecast_df in resultados_regioes.items():
            # Usa Data como indice e extrai previsao central de cada regiao.
            dfs_regioes[reg] = forecast_df.set_index('Data')['Previsao_QNT']
        
        # Combinar em um único DataFrame com datas na coluna A
        if dfs_regioes:
            consolidated = pd.DataFrame(dfs_regioes).sort_index()
            consolidated = consolidated.reset_index()
            
            # Escrever na aba correspondente
            consolidated.to_excel(writer, sheet_name=aba, index=False)
            print(f"Aba '{aba}' salva com {len(dfs_regioes)} regiões")

print(f"Arquivo de saida gerado em: {arquivo_saida}")
            
