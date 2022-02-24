# Importar bibliotecas
import pandas as pd
import win32com.client as win32

# Importar a base de dados do excel
tab_vendas = pd.read_excel('Vendas.xlsx')

# Vizualizar a base de dados
pd.set_option('display.max_columns', None)

# Tratar os dados
# Faturamento por loja
faturamento = tab_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja
produtos = tab_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(produtos)

# Ticket médio por produto em cada loja
# pegar o valor final e quantidade das tabelas
print('-' * 50)
ticket = faturamento['Valor Final'] / produtos['Quantidade'].to_frame()
ticket = ticket.rename(columns={0: 'Ticket Médio'})
print(ticket)

# Enviar um e-mail como relatório

outlook = win32.Dispatch('Outlook.application')
mail = outlook.CreateItem(0)
mail.to = 'cloretojannuzzi@outlook.com'
mail.subject = 'Relátorio de Vendas'
mail.HTMLBody = f'''
<p>Olá, segue abaixo o Relátório de Vendas por cada Loja:</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{produtos.to_html()}

<p>Ticket médio dos produtos:</p>
{ticket.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida só me retornar!</p>

<p>Att,</p>
<p>Cloreto Jannuzzi.</p>

'''
mail.Send()
