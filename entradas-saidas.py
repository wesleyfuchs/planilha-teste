import pandas as pd

def compare_excel_files(file1, file2, output_file):
    # Carregar as planilhas
    
    # Carrega planilha ANTIGA
    df1 = pd.read_excel(file1, engine='openpyxl')

    # Carrega planilha NOVA
    df2 = pd.read_excel(file2, engine='openpyxl')

    # Mescla os dois dataframes
    merged = pd.merge(df1, df2, on=list(df1.columns), how='outer', indicator=True)
    
    # Salva a nova planilha
    merged.to_excel(output_file, index=False)

if __name__ == "__main__":
    file1 = "sheet1.xlsx"
    file2 = "sheet2.xlsx"
    output_file = "differences.xlsx"

    compare_excel_files(file1, file2, output_file)