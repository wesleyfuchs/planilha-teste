import pandas as pd

def compare_excel_files(file1, file2, output_file):
    # Load the first Excel file
    df1 = pd.read_excel(file1, engine='openpyxl')

    # Load the second Excel file
    df2 = pd.read_excel(file2, engine='openpyxl')

    # Merge the two DataFrames on all columns
    merged = pd.merge(df1, df2, on=list(df1.columns), how='outer', indicator=True)

    # # Filter out rows that are present in both DataFrames
    # differences = merged[merged['_merge'] != 'both']

    # # Define a custom function to apply formatting
    # def format_cell(val):
    #     if val in differences:
    #         return 'color: red'
    #     else:
    #         return 'color: green'

    # # Apply the custom formatting function to the DataFrame
    # differences_styled = differences.style.applymap(format_cell)

    # # Save the styled differences to a new Excel file
    # differences_styled.to_excel(output_file, index=False)
    
    # # Save the styled differences to a new Excel file
    # differences.to_excel(output_file, index=False)
    
    # Save the styled differences to a new Excel file
    merged.to_excel(output_file, index=False)

if __name__ == "__main__":
    file1 = "sheet1.xlsx"
    file2 = "sheet2.xlsx"
    output_file = "differences.xlsx"

    compare_excel_files(file1, file2, output_file)