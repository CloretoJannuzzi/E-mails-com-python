# Importar bibliotecas
import pandas as pd
import win32com.client as win32
import time

# Importar a base de dados do excel
tab_vendas = pd.read_excel("python\windowspy\e_mail\Vendas.xlsx") #aqui você põe o caminho do arquivo em Excel, mesmo que esteja na mesma pasta

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
ticket = (faturamento['Valor Final'] / produtos['Quantidade']).to_frame()
ticket = ticket.rename(columns={0: 'Ticket Médio'})
print(ticket)


# Enviar um e-mail como relatório

outlook = win32.Dispatch('Outlook.application')
email = outlook.CreateItem(0)
email.to = 'cloretojannuzzi@outlook.com'
email.subject = 'Relátorio de Vendas'
email.HTMLBody = f'''
<p>Olá, segue abaixo o Relátório de cada Loja:</p>

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
email.display()
time.sleep(2)
email.send()
