import pandas as pd

# Carregando os dados dos arquivos Excel
saldo_df = pd.read_excel('SaldoITEM.xlsx')
movimento_df = pd.read_excel('MovtoITEM.xlsx')

# Convertendo a coluna de data para o formato apropriado
movimento_df['data_lancamento'] = pd.to_datetime(movimento_df['data_lancamento'], format='%d/%m/%Y')

# Mesclando os dados com base no código do item
merged_df = pd.merge(movimento_df, saldo_df, on='item')

# Criando colunas separadas para ENTRADA e SAÍDA
merged_df['Lançamentos de ENTRADA: quantidade'] = 0.0
merged_df['Lançamentos de ENTRADA: valor'] = 0.0
merged_df['Lançamentos de SAIDA: quantidade'] = 0.0
merged_df['Lançamentos de SAIDA: valor'] = 0.0

# Separando os valores de ENTRADA e SAÍDA nas colunas apropriadas
merged_df.loc[merged_df['tipo_movimento'] == 'Ent', 'Lançamentos de ENTRADA: quantidade'] = merged_df['quantidade'].astype(float)
merged_df.loc[merged_df['tipo_movimento'] == 'Ent', 'Lançamentos de ENTRADA: valor'] = merged_df['valor'].astype(float)
merged_df.loc[merged_df['tipo_movimento'] == 'Sai', 'Lançamentos de SAIDA: quantidade'] = merged_df['quantidade'].astype(float)
merged_df.loc[merged_df['tipo_movimento'] == 'Sai', 'Lançamentos de SAIDA: valor'] = merged_df['valor'].astype(float)

# Agrupando por item, data e tipo de movimento para calcular as somas diárias
daily_summary = merged_df.groupby(['item', 'data_lancamento']).agg({
    'Lançamentos de ENTRADA: quantidade': 'sum',
    'Lançamentos de ENTRADA: valor': 'sum',
    'Lançamentos de SAIDA: quantidade': 'sum',
    'Lançamentos de SAIDA: valor': 'sum',
}).reset_index()

# Preenchendo os valores nulos com 0
daily_summary.fillna(pd.NaT, inplace=True)

# Mesclando com os dados de saldo inicial e final
final_summary = pd.merge(daily_summary, saldo_df, left_on=['item', 'data_lancamento'], right_on=['item', 'data_inicio'], how='left')
final_summary.fillna(0, inplace=True)

# Calculando o Saldo Final usando a fórmula: saldo final = saldo inicial + entrada – saída
final_summary['Saldo Final em quantidade'] = round(final_summary['qtd_inicio'] + final_summary['Lançamentos de ENTRADA: quantidade'] - final_summary['Lançamentos de SAIDA: quantidade'], 2)
final_summary['Saldo Final em valor'] = round(final_summary['valor_inicio'] + final_summary['Lançamentos de ENTRADA: valor'] - final_summary['Lançamentos de SAIDA: valor'], 2)

# Reorganizando as colunas
final_summary = final_summary[['item', 'data_lancamento', 'Lançamentos de ENTRADA: quantidade', 'Lançamentos de ENTRADA: valor', 'Lançamentos de SAIDA: quantidade', 'Lançamentos de SAIDA: valor', 'qtd_inicio', 'valor_inicio', 'Saldo Final em quantidade', 'Saldo Final em valor']]

# Renomeando as colunas
final_summary.columns = ['Item', 'Data do lançamento', 'Lançamentos de ENTRADA: quantidade', 'Lançamentos de ENTRADA: valor', 'Lançamentos de SAIDA: quantidade', 'Lançamentos de SAIDA: valor', 'Saldo Inicial em quantidade', 'Saldo Inicial em valor', 'Saldo Final em quantidade', 'Saldo Final em valor']

# Formatando a coluna 'Data do lançamento' no formato 'dd/mm/aaaa'
final_summary['Data do lançamento'] = final_summary['Data do lançamento'].dt.strftime('%d/%m/%Y')

# Salvando os dados em um arquivo CSV
final_summary.to_csv('movimentacao_diaria_10.csv', index=False)
