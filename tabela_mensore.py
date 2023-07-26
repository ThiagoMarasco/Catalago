import pandas as pd

def mensore():
    caminho = 'C:\VSCODE\Catalago\Catalago_excel\TABELA DE PREÇO Mensore.xlsm'
    tabela_livro = pd.read_excel(caminho,header=21, sheet_name=None)
    dados = pd.concat(tabela_livro.values(), ignore_index=True)
    colunas_desejadas = ['CÓDIGO','PREÇO FINAL', 'MÚLTIPLO DE', 'MARCA', 'CATEGORIA',
                          'FAMÍLIA', 'DESCRIÇÃO', 'COD. BARRAS  (EAN 13)',
                      'ORIGEM         ', 'CODIGO CEST']


    dados = dados.fillna('')
    dados = dados.loc[dados['COD. BARRAS  (EAN 13)'] !='']
    
    dados = dados[colunas_desejadas]
    dados.rename(columns={'COD. BARRAS  (EAN 13)': 'EAN', 'ORIGEM         ': 'ORIGEM',
                   'CLASSIF. FISCAL  (NCM)': 'NCM', 'DESCRIÇÃO': 'ITEM',
                     'QTD P/CX': 'CAIXA',
                   'CODIGO CEST': 'CODIGO','CATEGORIA':'SEGMENTO',
                     'PREÇO FINAL': 'VALOR'}, inplace=True)
    dados = dados.loc[dados['MARCA'] =='MARCO BONI']
    print(dados)

    return(dados)


mensore()