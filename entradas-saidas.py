import pandas as pd

def comparar_planilhas(planilha_base, planilha_atualizada, resultado):
    # Carregar as planilhas
    
    # Carrega planilha ANTIGA
    df1 = pd.read_excel(planilha_base, engine='openpyxl')

    # Carrega planilha NOVA
    df2 = pd.read_excel(planilha_atualizada, engine='openpyxl')

    # Mescla os dois dataframes
    merged = pd.merge(df1, df2, on=list(df1.columns), how='outer', indicator=True)
    
    # Salva a nova planilha
    merged.to_excel(resultado, index=False)

if __name__ == "__main__":
    planilha_base = "sheet1.xlsx"
    planilha_atualizada = "sheet2.xlsx"
    resultado = "differences.xlsx"

    comparar_planilhas(planilha_base, planilha_atualizada, resultado)