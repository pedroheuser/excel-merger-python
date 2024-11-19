import pandas as pd
import os 

def consolidar_planilhas(diretorio):
    dataframe_finalizado = []
    colunas = {
        'Número de Inscrição do CNPJ': 'CNPJ',
        'Nome Fantasia': 'Registro RF', 
        'Municipio': 'Municipio', 
        'Atividade Turistica': 'Atividade Turistica', 
        'Validade do Certificado': 'Validade do certificado'
    }
    for arquivo in os.listdir(diretorio):
        if arquivo.endswith(('.xslx', '.xls')):
            caminho_arquivo = os.path.join(diretorio, arquivo)
            df = pd.read_excel(caminho_arquivo)
            df.rename(columns=colunas, inplace=True)
            colunas_filtradas = []
            for c in colunas.values():
                if c in df.columns:
                    colunas_filtradas.append(c)
            df_filtrado = df[colunas_filtradas]
            dataframe_finalizado.append(df_filtrado)

    df_final = pd.concat(dataframe_finalizado, ignore_index=True)
    df_final.to_excel(caminho_output, index=False)
    return df_final

diretorio_planilhas = # diretorio das planilhas 
caminho_output = # diretorio para o final
result = consolidar_planilhas(diretorio_planilhas, caminho_output)