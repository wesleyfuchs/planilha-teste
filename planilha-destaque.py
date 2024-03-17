import pandas as pd

def compare_excel_files(file1, file2, output_file):
    # Load the first Excel file
    df1 = pd.read_excel(file1, engine='openpyxl')

    # Load the second Excel file
    df2 = pd.read_excel(file2, engine='openpyxl')

    # Calculate the differences between the two dataframes
    differences = pd.concat([df1, df2]).drop_duplicates(keep=False)

    # Save the differences to a new Excel file
    differences.to_excel(output_file, index=False)

if __name__ == "__main__":
    file1 = "sheet1.xlsx"
    file2 = "sheet2.xlsx"
    output_file = "differences.xlsx"

    compare_excel_files(file1, file2, output_file)