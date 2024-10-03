import pandas as pd
import win32com.client as win32

# importar base de dados
tabela_vendas = pd.read_excel('vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

print(faturamento)

# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

print(quantidade)

print('-'*50)
# ticket médio de produtos vendidos por loja
ticket_medio = faturamento['Valor Final'] / quantidade['Quantidade'].astype(float)

print(ticket_medio)

# enviar um email com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'herosmash002@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas de cada loja.</p>

<p>Faturamento:</p>

{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format })}

<p>Quantidade vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos produtos:</p>
{ticket_medio.to_frame().to_html(formatters={'Valor Final':'R${:,.2f}'.format })}

<p>Qualquer dúvida estou a disposição</p>

<p>Att.,</p>

Alan'''

mail.Send()

print('Email enviado')
