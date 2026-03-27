import pandas as pd
from pmdarima import auto_arima

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

    df_total.columns = df_total.columns.astype(str).str.strip()
    df_total.rename(columns={df_total.columns[0]: 'DATA'}, inplace=True)
    df_total['DATA'] = pd.to_datetime(df_total['DATA'], dayfirst=True, errors='coerce')
    df_total = df_total.dropna(subset=['DATA']).sort_values('DATA')

    resultados_abas[aba] = {}

    for reg in regiao:
        if reg not in df_total.columns:
            print(f"  Coluna {reg} não encontrada")
            continue

        ts = df_total.set_index('DATA')[reg].dropna()
        if len(ts) < 14:
            print(f"  {reg}: poucos dados ({len(ts)} observações)")
            continue

        try:
            model = auto_arima(
                ts,
                d=None,
                D=None,
                seasonal=True,
                m=7,
                trace=False,
                error_action='ignore',
                suppress_warnings=True,
                stepwise=True,
                n_jobs=1
            )

            resultados_abas[aba][reg] = {
                'order': model.order,
                'seasonal_order': model.seasonal_order,
                'aic': model.aic()
            }

            print(
                f"  {reg}: order={model.order}, "
                f"seasonal_order={model.seasonal_order}, AIC={model.aic():.2f}"
            )
        except Exception as e:
            print(f"  Erro no auto_arima para {reg}: {e}")

print(f"\n{'='*60}")
print("Resumo final")
print(f"{'='*60}")

for aba, regioes in resultados_abas.items():
    print(f"\n[{aba}]")
    for reg, params in regioes.items():
        print(
            f"{reg}: order={params['order']} | "
            f"seasonal_order={params['seasonal_order']} | AIC={params['aic']:.2f}"
        )

arquivo_saida = r"C:\Users\izabe\Desktop\Projeto Bernardo\txt\melhores_parametros_sarima.txt"

with open(arquivo_saida, 'w', encoding='utf-8') as f:
    f.write("Melhores parâmetros SARIMA por aba e região\n\n")
    for aba, regioes in resultados_abas.items():
        f.write(f"[{aba}]\n")
        for reg, params in regioes.items():
            f.write(
                f"{reg}: order={params['order']} | "
                f"seasonal_order={params['seasonal_order']} | AIC={params['aic']:.2f}\n"
            )
        f.write("\n")

print(f"\nArquivo salvo em: {arquivo_saida}")