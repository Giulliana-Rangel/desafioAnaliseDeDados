import pandas as pd
# informações extras obtidas: https://pandas.pydata.org/docs/user_guide/index.html#user-guide

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

# Criando colunas para o saldo final
merged_df['Saldo Final em quantidade'] = 0.0
merged_df['Saldo Final em valor'] = 0.0

# Inicializando o saldo final com o saldo inicial (fonte: https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.iat.html)
merged_df['Saldo Final em quantidade'].iat[0] = merged_df['qtd_inicio'].iat[0]
merged_df['Saldo Final em valor'].iat[0] = merged_df['valor_inicio'].iat[0]

# Calculando o Saldo Final usando a fórmula: saldo final = saldo anterior + entrada – saída
# Das linhas 35 a 38 eu utilizei auxílio do chatGPT pq eu estava construindo usando o guia porém, continha erros de sintaxe.
for i in range(1, len(merged_df)):
    merged_df['Saldo Final em quantidade'].iat[i] = merged_df['Saldo Final em quantidade'].iat[i-1] + merged_df['Lançamentos de ENTRADA: quantidade'].iat[i] - merged_df['Lançamentos de SAIDA: quantidade'].iat[i]
    merged_df['Saldo Final em valor'].iat[i] = merged_df['Saldo Final em valor'].iat[i-1] + merged_df['Lançamentos de ENTRADA: valor'].iat[i] - merged_df['Lançamentos de SAIDA: valor'].iat[i]

# Reorganizando as colunas
final_summary = merged_df[['item', 'data_lancamento', 'Lançamentos de ENTRADA: quantidade', 'Lançamentos de ENTRADA: valor', 'Lançamentos de SAIDA: quantidade', 'Lançamentos de SAIDA: valor', 'qtd_inicio', 'valor_inicio', 'Saldo Final em quantidade', 'Saldo Final em valor']]

# Renomeando as colunas
final_summary.columns = ['Item', 'Data do lançamento', 'Lançamentos de ENTRADA: quantidade', 'Lançamentos de ENTRADA: valor', 'Lançamentos de SAIDA: quantidade', 'Lançamentos de SAIDA: valor', 'Saldo Inicial em quantidade', 'Saldo Inicial em valor', 'Saldo Final em quantidade', 'Saldo Final em valor']

# Formatando a coluna 'Data do lançamento' no formato 'dd/mm/aaaa'
final_summary['Data do lançamento'] = final_summary['Data do lançamento'].dt.strftime('%d/%m/%Y')

# Salvando os dados em um arquivo CSV
final_summary.to_csv('movimentacao_diaria_6.csv', index=False)

# ****************************** SEGUNDA PLANILHA *********************************************
# Planilha que possui somente os balanços finais de cada dia

final_summary['Data do lançamento'] = pd.to_datetime(final_summary['Data do lançamento'], format='%d/%m/%Y')

final_summary = final_summary.sort_values(by=['Item', 'Data do lançamento'])

last_entries_summary = final_summary.drop_duplicates(subset=['Item', 'Data do lançamento'], keep='last')

last_entries_summary.to_csv('ultimas_entradas_do_dia_agrupadas_1.csv', index=False)
