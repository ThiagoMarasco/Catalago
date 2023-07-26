import pandas as pd

def livro():
    caminho = 'C:\VSCODE\Catalago\Catalago_excel\livro SP.xlsx'
    tabela_livro = pd.read_excel(caminho, sheet_name=None)
    dados_livro = pd.concat(tabela_livro.values(), ignore_index=True)
    dados_index = dados_livro.iloc[1].str.lstrip()
    dados_index = dados_livro.iloc[1].str.rstrip()
    dados_livro.rename(columns=dict(zip(dados_livro.columns, dados_index))
                       , inplace=True)
    
    dados_livro =  dados_livro.drop(0)
    dados_livro = dados_livro.drop(1)
    dados = dados_livro.reset_index()

    
    dados_prazo = dados['PRAZO']
    dados_FORNECEDOR = dados['FORNECEDOR']
    dados_cod = dados['CÓD.']
    dados_des = dados['DESCRIÇÃO']
    dados_ean = dados['EAN']
    dados_qcaixa = dados['QUANT.CAIXA']
    
    dados_2 = pd.DataFrame()
    
    dados_2['EAN'] = dados_ean 
    dados_2['MARCA'] = dados_FORNECEDOR
    dados_2['CODIGO'] = dados_cod 
    dados_2['ITEM'] = dados_des 
    dados_2['CAIXA'] = dados_qcaixa 
    dados_2['VALOR'] = dados_prazo
    return(dados_2)

    