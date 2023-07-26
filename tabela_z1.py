import pandas as pd

def z1():
    caminho = 'C:\VSCODE\Catalago\Catalago_excel\Tabela Z1- SÃ_O PAULO 20.03.2023.xls'
    tabela_livro = pd.read_excel(caminho, sheet_name=None, header=6)
    dados = pd.concat(tabela_livro.values(), ignore_index=True)
    dados.rename(columns={'Cód. EAN-13': 'EAN', 'Produto': 'ITEM',
                           'Z1 Física': 'VALOR','Cx': 'CAIXAS'}, inplace=True)
    
    dados = dados.fillna('')
    dados = dados.loc[dados['EAN'] !='']
    return(dados)

z1()