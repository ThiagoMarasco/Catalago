import pandas as pd

def farma():
    # Caminho do arquivo PDF
    caminho_arquivo_pdf = 'C:\VSCODE\Catalago\Catalago_excel\Farmácia 2023.xlsx'
    
    tabela_estoque = pd.read_excel(caminho_arquivo_pdf, sheet_name=None,header=2)
    dados = pd.concat(tabela_estoque.values(), ignore_index=True)
    
    dados = dados.fillna('')

    dados['MARCA'] ='CURAPROX'
    
    dados.rename(columns={'DESCRIÇÃO COMPLETA': 'ITEM','CUSTO UNID C/ ST': 'VALOR'},
                  inplace=True)
    return(dados)
