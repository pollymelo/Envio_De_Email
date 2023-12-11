import pandas as pd

# importar a base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')


# visualizar a base de dados
# pd.set_option(opção, valor)
pd.set_option('display.max_columns', None)
# mostra tudo
print(tabela_vendas)
print('-' * 50)

# visualizar tabelas
# cria uma nova tabela mostrando esses valores
# utilizado para filtrar colunas
# tabela_vendas[['ID Loja', 'Valor Final']]


# faturamento
# agrupa as lojas e soma a coluna
#groupby

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)

# quantidade de produtos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)

# ticket médio por produto em cada loja
# faturamento dividido por quantidade
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)
print('-' * 50)
# to.frame() - transformar em uma tabela, padronizar

# enviar um email com o relatório
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'pollyammelo@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''

<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att,</p>

<p>Pollyanna Melo.</p>
'''

mail.Send()
print('Email Enviado')


