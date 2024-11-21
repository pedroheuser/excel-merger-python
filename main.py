import pandas as pd
import os 

def consolidar_planilhas(diretorio, caminho_output):
    dataframe_finalizado = []
    colunas = {
        'Número de Inscrição do CNPJ': 'CNPJ',
        'Nome Fantasia': 'Registro RF', 
        'Município': 'Município', 
        'Atividade Turística': 'Atividade Turística', 
        'Validade do Certificado': 'Validade do Certificado'
    }
    
    print(f"Arquivos no diretório {diretorio}: {os.listdir(diretorio)}")
    
    for arquivo in os.listdir(diretorio):
        if arquivo.endswith(('.xlsx', '.xls')):
            caminho_arquivo = os.path.join(diretorio, arquivo)
            print(f"Lendo arquivo: {caminho_arquivo}")
            
            try:
                df = pd.read_excel(caminho_arquivo)
                print(f"Colunas encontradas: {df.columns}")
                df.rename(columns=colunas, inplace=True)
                colunas_filtradas = [c for c in colunas.values() if c in df.columns]
                df_filtrado = df[colunas_filtradas]
            
                dataframe_finalizado.append(df_filtrado)
            except Exception as e:
                print(f"Erro ao processar o arquivo {arquivo}: {e}")
    
    if not dataframe_finalizado:
        print("Nenhum arquivo válido encontrado para consolidar.")
        return None
    
    df_final = pd.concat(dataframe_finalizado, ignore_index=True)
    df_final.to_excel(caminho_output, index=False)
    return df_final

diretorio_planilhas = "" 
caminho_output = ""

result = consolidar_planilhas(diretorio_planilhas, caminho_output)
