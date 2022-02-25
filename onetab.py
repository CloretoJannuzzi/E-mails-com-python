# Importar bibliotecas
import pandas as pd
import win32com.client as win32
import time


# Importar a base de dados do excel
tab_vendas = pd.read_excel("python\windowspy\e_mail\Vendas.xlsx") # Caminho da sua base de dados

# Vizualizar a base de dados
pd.set_option('display.max_columns', None)

# Tratar os dados
# Faturamento por loja
faturamento = tab_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

# Quantidade de produtos vendidos por loja
produtos = tab_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# Ticket médio por produto em cada loja
# pegar o valor final e quantidade das tabelas
ticket = (faturamento['Valor Final'] / produtos['Quantidade']).to_frame()
ticket = ticket.rename(columns={0: 'Ticket Médio'})

# formar somente uma tabela, juntar ticket medio e as outras colunas.
tab = tab_vendas[['ID Loja', 'Valor Final',
                  'Quantidade', ]].groupby('ID Loja').sum()

fusao = pd.merge(tab, ticket,  how='inner', on='ID Loja')  # Fusão das tabelas

print(fusao)

# Enviar um e-mail como relatório
outlook = win32.Dispatch('Outlook.application')
email = outlook.CreateItem(0)
email.to = 'cloretojannuzzi@outlook.com'
email.subject = 'Relátorio de Vendas'
email.HTMLBody = f'''
    <p>Olá, segue abaixo o Relátório de cada Loja:</p>
    {fusao.to_html(formatters={'Valor final': 'R${:,.2f}'.format})} 

    <p>Qualquer dúvida só me retornar!</p>
    <p>Att,</p>
    <p>Cloreto Jannuzzi.</p>

'''
# OBS: o formatters não está funcionando.
email.display()
time.sleep(2)
email.send()
