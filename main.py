import pandas as pd
import os 

def consolidar_planilhas(diretorio_planilhas, caminho_output):
    print(f"Verificando o caminho: {diretorio_planilhas}")
    if not os.path.exists(diretorio_planilhas):
        print(f"O diretório {diretorio_planilhas} não existe. Verifique o caminho.")
        return None

    dataframe_finalizado = []
    colunas = {
        'Número de Inscrição do CNPJ': 'CNPJ',
        'Nome Fantasia': 'Registro RF', 
        'Município': 'Município', 
        'Atividade Turística': 'Atividade Turística', 
        'Validade do Certificado': 'Validade do Certificado'
    }
    
    print(f"Arquivos no diretório {diretorio_planilhas}: {os.listdir(diretorio_planilhas)}")
    
    for arquivo in os.listdir(diretorio_planilhas):
        if arquivo.endswith(('.xlsx', '.xls')):
            caminho_arquivo = os.path.join(diretorio_planilhas, arquivo)
            print(f"Lendo arquivo: {caminho_arquivo}")
            
            try:
                df = pd.read_excel(caminho_arquivo)
                print(f"Colunas encontradas: {df.columns}")
                df.rename(columns=colunas, inplace=True)
                
                colunas_filtradas = [c for c in colunas.values() if c in df.columns]
                df_filtrado = df[colunas_filtradas]
                
                if "CNPJ" in df_filtrado.columns:
                    df_filtrado["CNPJ"] = df_filtrado['CNPJ'].astype(str)
                    df_filtrado["CNPJ"] = df_filtrado['CNPJ'].str.zfill(14)
                
                if "Validade do Certificado" in df_filtrado.columns:
                    df_filtrado['Validade do Certificado'] = pd.to_datetime(
                        df_filtrado['Validade do Certificado']
                    ).dt.date

                dataframe_finalizado.append(df_filtrado)
            except Exception as e:
                print(f"Erro ao processar o arquivo {arquivo}: {e}")
    
    if not dataframe_finalizado:
        print("Nenhum arquivo válido encontrado para consolidar.")
        return None
    
    df_final = pd.concat(dataframe_finalizado, ignore_index=True)
    df_final = df_final.drop_duplicates(subset='CNPJ', keep='first')
    df_final.to_excel(caminho_output, index=False)
    return df_final

diretorio_planilhas = r"" 
caminho_output = r""

result = consolidar_planilhas(diretorio_planilhas, caminho_output)
