import pandas as pd
import win32com.client as win32

arquivo = pd.read_excel('Vendas.xlsx')

faturamento = arquivo[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print()

quantidade = arquivo[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print()

ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)
print()

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'adweslly.ferreira19@gmail.com'
mail.Subject = 'Relatorio'
mail.HTMLBody = f'''
<p>Prezados, </p>
<p>Segue abaixo o relatorio de venda das lojas. </p>

<p>Faturamento por loja: </p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade de produto vendidos por loja: </p>
{quantidade.to_html()}

<p>Ticket Médio: </p>
{ticket_medio.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Att.</p>
<p>Adweslly F. Silva</p>
'''
mail.Send()
print('Email enviado com sucesso!')
