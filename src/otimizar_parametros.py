import pandas as pd
import sys
from pmdarima import auto_arima
import warnings
warnings.filterwarnings('ignore')

arquivo24 = r"C:\Users\izabe\Downloads\DADOS 2024.xlsx"
arquivo25 = r"C:\Users\izabe\Downloads\DADOS 2025.xlsx"
abas = ['AAE', 'CANT', 'PORT', 'SERV']
regiao = ['RG B', 'RG L', 'RG N', 'RG NO', 'RG NE', 'RG O', 'RG VN', 'RG CS','RG P']

# Dicionário para armazenar os melhores parâmetros
melhores_parametros = {}

print("="*80)
print("ANÁLISE AUTOMÁTICA DE PARÂMETROS SARIMA")
print("="*80)

# Processar cada aba separadamente
for aba in abas:
    print(f"\n{'='*80}")
    print(f"ABA: {aba}")
    print(f"{'='*80}\n")
    
    try:
        df_24 = pd.read_excel(arquivo24, sheet_name=aba)
        df_25 = pd.read_excel(arquivo25, sheet_name=aba)
    except Exception as e:
        print(f'Erro ao ler a aba {aba}: {e}')
        continue
    
    df_total = pd.concat([df_24, df_25], ignore_index=True)
    
    # Preparar dados
    df_total.columns = df_total.columns.astype(str).str.strip()
    df_total.rename(columns={df_total.columns[0]: 'DATA'}, inplace=True)
    df_total['DATA'] = pd.to_datetime(df_total['DATA'], dayfirst=True, errors='coerce')
    df_total = df_total.dropna(subset=['DATA'])
    df_total = df_total.sort_values('DATA')

    # Dicionário para armazenar parâmetros desta aba
    melhores_parametros[aba] = {}

    # Processar cada região (coluna)
    for reg in regiao:
        if reg not in df_total.columns:
            print(f"  ⚠️  Coluna {reg} não encontrada")
            continue
        
        print(f"  Analisando: {reg}...", end=" ", flush=True)
        
        # Criar série temporal para esta região
        ts = df_total.set_index('DATA')[reg].copy()
        ts = ts.dropna()
        
        if len(ts) < 14:  # Mínimo para análise sazonal semanal
            print(f"❌ Poucos dados ({len(ts)} obs)")
            continue
        
        try:
            # Auto ARIMA com busca completa
            model = auto_arima(
                ts,          
                d=None,                        # deixar auto determinar diferenciaçã0
                seasonal=True,                 # habilitar sazonalidade
                m=7,                           # período sazonal (semanal)
                trace=False,                   # não mostrar todas tentativas
                error_action='ignore',
                suppress_warnings=True,
                stepwise=True,                 # busca stepwise (mais rápido)
                n_jobs=1                       # stepwise não roda em paralelo
            )
            
            order = model.order
            seasonal_order = model.seasonal_order
            aic = model.aic()
            
            # Armazenar os melhores parâmetros
            melhores_parametros[aba][reg] = {
                'order': order,
                'seasonal_order': seasonal_order,
                'AIC': aic
            }
            
            print(f"✓ order={order}, seasonal_order={seasonal_order}, AIC={aic:.2f}")
            
        except Exception as e:
            print(f"❌ Erro: {str(e)[:50]}")

print(f"\n{'='*80}")
print("RESUMO DOS MELHORES PARÂMETROS")
print(f"{'='*80}\n")

# Criar arquivo de saída com os resultados
arquivo_saida = r"C:\Users\izabe\Desktop\Projeto Bernardo\txt\parametros_otimizados.txt"

with open(arquivo_saida, 'w', encoding='utf-8') as f:
    f.write("="*80 + "\n")
    f.write("PARÂMETROS SARIMA OTIMIZADOS - AUTO ARIMA\n")
    f.write("="*80 + "\n\n")
    
    for aba, regioes in melhores_parametros.items():
        f.write(f"\n{'='*80}\n")
        f.write(f"ABA: {aba}\n")
        f.write(f"{'='*80}\n\n")
        print(f"\n{aba}:")
        
        for reg, params in regioes.items():
            linha = f"  {reg:8} → order={params['order']}, seasonal_order={params['seasonal_order']}, AIC={params['AIC']:.2f}"
            f.write(linha + "\n")
            print(linha)
    
    f.write(f"\n\n{'='*80}\n")
    f.write("CÓDIGO SUGERIDO PARA CADA ABA/REGIÃO:\n")
    f.write(f"{'='*80}\n\n")
    
    for aba, regioes in melhores_parametros.items():
        f.write(f"\n# {aba}\n")
        for reg, params in regioes.items():
            f.write(f"# {reg}: order={params['order']}, seasonal_order={params['seasonal_order']}\n")

print(f"\n{'='*80}")
print(f"✓ Resultados salvos em: {arquivo_saida}")
print(f"{'='*80}\n")

# Criar também um Excel com os parâmetros
try:
    excel_saida = r"C:\Users\izabe\Desktop\Projeto Bernardo\excel\parametros_otimizados.xlsx"
    
    dados_excel = []
    for aba, regioes in melhores_parametros.items():
        for reg, params in regioes.items():
            dados_excel.append({
                'Aba': aba,
                'Região': reg,
                'order_p': params['order'][0],
                'order_d': params['order'][1],
                'order_q': params['order'][2],
                'seasonal_P': params['seasonal_order'][0],
                'seasonal_D': params['seasonal_order'][1],
                'seasonal_Q': params['seasonal_order'][2],
                'seasonal_m': params['seasonal_order'][3],
                'AIC': round(params['AIC'], 2)
            })
    
    df_params = pd.DataFrame(dados_excel)
    df_params.to_excel(excel_saida, index=False)
    print(f"✓ Parâmetros também salvos em Excel: {excel_saida}\n")
    
except Exception as e:
    print(f"⚠️  Não foi possível salvar Excel: {e}\n")
