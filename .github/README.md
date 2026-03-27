# Projeto Bernardo - Análise de Pedidos por Região

Sistema de processamento e análise de dados de pedidos organizados por dias, com transformação de dados e agregação por região geográfica.

## 📋 Descrição

Este projeto tem como objetivo prever o numero de pedidos de substituição nas escolas de Belo horizonte, para mode lar os dados usamos as bibliotecas "pandas", "openxml", "numpy", "dateutil". O código lê dados de arquivos Excel, transforma o formato dos dados, extrai informações de região usando expressões regulares, agrega os dados por região e exporta os resultados em novos arquivos Excel.

### Funcionalidades Principais

- ✅ Leitura dinâmica de múltiplas abas do Excel
- ✅ Detecção automática de colunas de dados (B1 até primeira coluna vazia)
- ✅ Transformação de formato largo (wide) para formato longo (long)
- ✅ Extração de região usando Expressões Regulares (Regex)
- ✅ Agregação de dados por região
- ✅ Exportação automática de resultados em Excel

## 📊 Estrutura de Dados

### Entrada

- Arquivo Excel com múltiplas abas: `AAE2`, `CANT2`, `PORT2`, `SERV2`
- Coluna fixa: `DATA` ou `Data`
- Colunas de escolas no padrão: `/REGIÃO/ESCOLA/...`

Exemplo:

```
DATA        | /REGIÃO A/Escola1 | /REGIÃO B/Escola2 | ...
2024-01-01  | 10                | 15                | ...
2024-01-02  | 12                | 18                | ...
```

### Saída

- Arquivos Excel com dados agregados por região
- Formato: `resultado_por_regiao{YEAR}_{ABA}.xlsx`
- Colunas: `Regiao` e `Valor` (soma dos pedidos)

## 🚀 Como Usar

### Pré-requisitos

```bash
pip install pandas openpyxl numpy dateutil statsmodels matplotlib
```

### Execução

```bash
python src/teste.py
```

### Fluxo de Processamento

1. **Leitura**: Carrega dados de 2024 da pasta Downloads
2. **Detecção Dinâmica**: Identifica automaticamente quais colunas contêm dados
3. **Transformação**: Converte dados de formato largo para longo
4. **Extração**: Usa regex para extrair a região do nome da escola
5. **Agregação**: Agrupa por região e soma os valores
6. **Exportação**: Salva resultado em novo arquivo Excel

## 📁 Estrutura de Arquivos

```
Projeto Bernardo/
├── src/
│   ├── teste.py              # Script principal de processamento
│   ├── codigo.py             # Código auxiliar
│   └── modelo_mensal.py      # Modelo de análise mensal
├── excel/                     # Arquivos Excel processados
├── img/                       # Imagens e gráficos
├── txt/                       # Documentos de texto
└── .github/
    └── README.md             # Este arquivo
```

## 🔧 Personalização

### Alterar Abas

Modifique a lista em `teste.py`:

```python
abas = ['AAE2', 'CANT2', 'PORT2', 'SERV2']
```

### Alterar Path dos Arquivos

```python
arquivo24 = r"C:\Seu\Path\PEDIDOS X DIAS 2024.xlsx"
arquivo25 = r"C:\Seu\Path\PEDIDOS X DIAS 2025.xlsx"
```

### Alterar Coluna Fixa

```python
colunas_fixas = ['SUA_COLUNA']
```

## 📈 Próximas Melhorias

- [ ] Integração com 2025 (arquivo25)
- [ ] Análise de séries temporais (SARIMAX)
- [ ] Visualizações gráficas com matplotlib
- [ ] Dashboard interativo
- [ ] Tratamento de erros aprimorado

## 🛠️ Tecnologias

- **Python 3.x**
- **Pandas** - Manipulação de dados
- **OpenPyXL** - Leitura/escrita de Excel
- **NumPy** - Operações numéricas
- **Statsmodels** - Análise de séries temporais
- **Matplotlib** - Visualização de dados
- **DateUtil** - Manipulação de datas

## 📝 Exemplo de Saída

Arquivo gerado: `resultado_por_regiao2024_AAE2.xlsx`

| Regiao   | Valor |
| -------- | ----- |
| REGIÃO A | 1250  |
| REGIÃO B | 895   |
| REGIÃO C | 2340  |

## ⚙️ Detalhes Técnicos

### Detecção Dinâmica de Colunas

O código identifica automaticamente o intervalo de colunas válidas:

1. Começa a partir de B (coluna 2)
2. Para na primeira coluna vazia
3. Cria lista com nomes de coluna neste intervalo

### Extração de Região com Regex

Pattern: `r'/([^/]+)/'`

- Captura texto entre duas barras
- Permite nomes de região variáveis

---
