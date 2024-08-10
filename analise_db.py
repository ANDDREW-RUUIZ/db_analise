import pandas as pd
import win32com.client as win32
# Importar dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None) # pandas mostrar todos os dados da tabela sem abreviar
print('{}\n{}'.format('-=' * 30, 'vendas') )
print(tabela_vendas)
#tabela_vendas[['ID Loja', 'Valor Final']] # filtra apenas as que estao escritas 
     

# Faturamento total por loja 
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum() # faz o agrupamento de celular comm memso valor e filtra
print('{}\n{}'.format('-=' * 30, 'Faturamento') )
print(faturamento )
# tabela_vendas.groupby('ID Loja').sum() agupara todos os id's iguais e o que esta nas linhas respectivamente ira somar SUM() caso queira media usar outra função no final

# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print('{}\n{}'.format('-=' * 30, 'Quantidade') )
print(quantidade)

# Ticket medio por produto de cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame() # coloca em chave[] para filtrar a coluna exata de cada tabela, colocando em uma TABELA
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'}) # renomeia o nome da coluna (columns) que estava com nome '0', agr passa a ser : Ticket Medio


# Enviar email com relatorio
outlook = win32.Dispatch('outlook.application') #conecta no email
mail = outlook.CreateItem(0) # cria o email 
mail.To = 'andrewrb1919@gmail.com' #para quem ira enviar
mail.Subject= 'Teste' # assunto do email
mail.HTMLBody = f''' 
<p>Prezados,</p>

<p>Segue o relatorio de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket medio dos produtos em cada loja</p>
{ticket_medio.to_html(formatters={'Ticket Medio':'R${:,.2f}'.format})}

<p>Att.</p>
<p>Andrew</p>
'''
# corpo do email||| f antes de um texto ''' ''' chama f string ou seja um texto que tem chaves{} que contem variaveis 
# .to_html() formata a tabela para uma tabela mais bonitaa denytro de um corpo htmlbody
# to_html(formatters={'nome da coluna a ser formatada':'R${:,.2f}'.format})} 'R$' adiciona unidade de medida {:,.2f} 2 casas coom fponto flutuante depois da virgula exR$ 10,00

mail.Send()

print('email enviado')