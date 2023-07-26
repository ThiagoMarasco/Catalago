import pandas as pd
from livro_sp import livro
from ficha_pedido import ficha
from tabela_mensore import mensore
from tabela_z1 import z1
from farmacia import farma
import re
##### = Tabela-Estoque 190723.
caminho ='C:\VSCODE\Catalago\Catalago_excel\Tabela-Estoque 190723..xlsx'
tabela_estoque = pd.read_excel(caminho, sheet_name=None)
dados = pd.concat(tabela_estoque.values(), ignore_index=True)

dados = dados.fillna('')

dados = dados.loc[dados['STATUS SKU'] !='DESATIVADO']
dados = dados.loc[dados['DESCRIÇÃO'] =='']
dados = dados.loc[dados['ESTOQUE'] !='S/ESTOQUE']
dados = dados.loc[dados['ESTOQUE UNID.'] !='S/ESTOQUE']
dados = dados.loc[dados['EAN'] !='']

dados['ESTOQUE'] = dados['ESTOQUE'].astype(str)
dados['ESTOQUE UNID.'] = dados['ESTOQUE UNID.'].astype(str)
dados['ESTOQUE'] = dados['ESTOQUE'] + ' ' + dados['ESTOQUE UNID.']


dados['VALOR '] = dados['VALOR '].astype(str)
dados['VALOR'] = dados['VALOR'].astype(str)
dados['VALOR'] = dados['VALOR'] + ' ' + dados['VALOR ']
dados['VALOR'] = dados['VALOR'].astype(str)

dados = dados[['EAN','SKU','STATUS SKU','MARCA','SEGMENTO','ITEM',
               'CAIXA','ESTOQUE','VALOR']]


#

#LIVRO SP
livro_SP = livro()
ficha_pedido = ficha()
mensore_tabela = mensore()
tabela_z1 = z1()
farmacia = farma()
#print(livro_SP)

dados = pd.concat([mensore_tabela,dados,livro_SP,ficha_pedido,tabela_z1,farmacia], ignore_index=True)
dados['MÚLTIPLO DE'] = dados['MÚLTIPLO DE'].astype(str)
dados['CAIXA'] = dados['CAIXA'].astype(str)
dados['CAIXA'] = dados['CAIXA'] + ' ' + dados['MÚLTIPLO DE']
dados_index = ['MARCA','EAN','ITEM','VALOR','CAIXA']
dados = dados[dados_index]
#'Valor (Wabi Sabi)','Valor Total (Á vista)','Valor Total (Prazo 30 dias)',	'Valor Total (Prazo 90 dias)'

dados['EAN'] = dados['EAN'].apply(pd.to_numeric, errors='coerce')
dados['CAIXA'] = dados['CAIXA'].str.replace(r'nan', '')
# Função para limpar os valores removendo caracteres não numéricos ou não relacionados à formatação BRL
def limpar_valor(valor):
    if isinstance(valor, str):
        return re.sub(r'[^\d.]', '', valor)
    return valor

# Aplicar a limpeza na coluna 'Valores'
dados['VALOR'] = dados['VALOR'].apply(limpar_valor)
dados['CAIXA'] = dados['CAIXA'].str.replace(r'.0', '')

# Converter a coluna 'Valores' para números (usando o tipo float)
dados['VALOR'] = pd.to_numeric(dados['VALOR'], errors='coerce')

# Função para formatar em BRL
def formatar_brl(valor):
    if pd.notnull(valor):
        return '{:.2f}'.format(valor)
    else:
        return ''

# Aplicar a formatação na coluna 'Valores'
dados['VALOR'] = dados['VALOR'].map(formatar_brl)
dados['VALOR'] = pd.to_numeric(dados['VALOR'], errors='coerce')
dados['CAIXA'] = pd.to_numeric(dados['CAIXA'], errors='coerce')
dados = dados.fillna('')
dados = dados.loc[dados['EAN'] !='']

dados['Valor (Wabi Sabi)'] = dados['VALOR']* 1.4
dados['Valor (Wabi Sabi)'] = pd.to_numeric(dados['Valor (Wabi Sabi)'], errors='coerce')

dados['CAIXA'] = dados['CAIXA'].replace(['NaN', 'None',''], 1)
dados['CAIXA'] = pd.to_numeric(dados['CAIXA'], errors='coerce')
dados['Valor Total (Á vista)'] = (dados['Valor (Wabi Sabi)']*dados['CAIXA'])* 0.95
dados['Valor Total (Prazo 30 dias)'] = (dados['Valor (Wabi Sabi)']*dados['CAIXA'])
dados['Valor Total (Prazo 90 dias)'] = (dados['Valor (Wabi Sabi)']*dados['CAIXA'])*1.05

def formatar_brl_br(valor):
    if pd.notnull(valor):
        return 'R$ {:.2f}'.format(valor)
    else:
        return ''

# Aplicar a formatação na coluna 'Valores'
dados['VALOR'] = dados['VALOR'].map(formatar_brl_br)
dados['Valor (Wabi Sabi)'] = dados['Valor (Wabi Sabi)'].map(formatar_brl_br)
dados['Valor Total (Á vista)'] = dados['Valor Total (Á vista)'].map(formatar_brl_br)
dados['Valor Total (Prazo 30 dias)'] = dados['Valor Total (Prazo 30 dias)'].map(formatar_brl_br)
dados['Valor Total (Prazo 90 dias)'] = dados['Valor Total (Prazo 90 dias)'].map(formatar_brl_br)



dados.to_excel('C:\VSCODE\Catalago\Catalago_excel\Catalogo 22.07.xlsx', index=False)