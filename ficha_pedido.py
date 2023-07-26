import pandas as pd

def ficha():
    caminho ='C:\VSCODE\Catalago\Catalago_excel\FICHA DE PEDIDO ATUALIZADA MARÇO23.xlsx'
    tabela_livro = pd.read_excel(caminho, sheet_name=None)
    dados = pd.concat(tabela_livro.values(), ignore_index=True)
    
    dados = dados.fillna('')

    dados = dados.loc[dados['EAN'] !='']

    colunas_desejadas = ['CÓD', 'EAN', 'DUN', 'NCM', 'DESCRIÇÃO', 'QTD P/CX',
                          'TAB CHEIA','DESC COME', 'Preço Final']
    
    dados = dados[colunas_desejadas]

    dados.rename(columns={'CÓD': 'Código', 'EAN': 'EAN', 'DUN': 'DUN',
                   'NCM': 'NCM', 'DESCRIÇÃO': 'ITEM', 'QTD P/CX': 'CAIXA',
                   'TAB CHEIA': 'Tabela Cheia', 'DESC COME': 'Desconto Comercial',
                     'Preço Final': 'VALOR'}, inplace=True)


    return(dados)

