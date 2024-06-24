import pandas as pd
import os

def compare_excel_files(file1, file2, output_file):
    # Load the Excel files
    xl1 = pd.ExcelFile(file1)
    xl2 = pd.ExcelFile(file2)

    with open(output_file, 'w', encoding='utf-8') as f:

        # Check if both files have the same sheet names
        if xl1.sheet_names != xl2.sheet_names:
            f.write("The files have different sheet names.\n")
            f.write(f"File1 sheets: {xl1.sheet_names}\n")
            f.write(f"File2 sheets: {xl2.sheet_names}\n")
            return

        # Iterate through each sheet and compare
        all_sheets_identical = True
        for sheet_name in xl1.sheet_names:
            df1 = xl1.parse(sheet_name)
            df2 = xl2.parse(sheet_name)

            if not df1.equals(df2):
                all_sheets_identical = False
                f.write(f"表单'{sheet_name}'中存在不一致\n")
                compare_dataframes(df1, df2, file1, file2, sheet_name, f)
            else:
                f.write(f"表单'{sheet_name}'在两文件中完全一致。\n")

        if all_sheets_identical:
            f.write("两文件完全一致。\n")

def compare_dataframes(df1, df2, file1, file2, sheet_name, file):

    file_name1 = os.path.splitext(os.path.basename(file1))[0]
    file_name2 = os.path.splitext(os.path.basename(file2))[0]

    # Ensure the dataframes have the same shape
    if df1.shape != df2.shape:
        file.write(f"表单'{sheet_name}'格式不一致:\n")
        file.write(f"文件1中表单形状为: {df1.shape}, 文件2中表单形状为: {df2.shape}\n")

    # Compare the dataframes cell by cell
    for row in range(max(len(df1), len(df2))):
        for col in range(max(len(df1.columns), len(df2.columns))):
            try:
                val1 = df1.iat[row, col]
            except IndexError:
                val1 = None
            try:
                val2 = df2.iat[row, col]
            except IndexError:
                val2 = None
            if val1 != val2:
                file.write(f"表单'{sheet_name}'中第{row}行第{col}列单元格存在不一致，单元格区别为{file_name1}中值为{val1}，{file_name2}中值为{val2}\n")

if __name__ == "__main__":
    file1 = "data/test1.xlsx"
    file2 = "data/test2.xlsx"
    output_file = "result.txt"
    compare_excel_files(file1, file2, output_file)