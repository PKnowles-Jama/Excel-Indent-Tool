# Excel Indent Function

import pandas as pd
import os

def indent_function(excel_file, numbering_column, indent_column):
    # excel_file: file name to manipulate
    # numbering_column: index or name of column to use for indent information
    # indent_coumn: index or name of column to be indented

    # Create a dataframe from the excel file
    df = pd.read_excel(excel_file)

    # Identify the column to be used for number of indents
    if isinstance(numbering_column, str):
        numbering_series = df[numbering_column]
    elif isinstance(numbering_column, int):
        numbering_series = df.iloc[:, numbering_column]
    else:
        print("Error, numbering_column must be a str or int")

    # Identify the column to be indented
    if isinstance(indent_column, str):
        indent_series = df[indent_column]
    elif isinstance(indent_column, int):
        indent_series = df.iloc[:, indent_column]
    else:
        print("Error, indent_column must be a str or int")

    indented_values = []
    for i in range(len(df)):
        try:
            indent_level = int(numbering_series.iloc[i])
            indent_str = "  " * indent_level
            indented_value = f"{indent_str}{indent_series.iloc[i]}"
            indented_values.append(indented_value)
        except (ValueError, TypeError): 
            print(f"Warning: Row {i+1}: Invalid or missing value in numbering column. Skipping indentation for this row.")
            indented_values.append(indent_series.iloc[i])

    # Update the dataframe
    df[indent_column] = indented_values

    # Save the changes
    suffix = "_new"
    base_name, ext = os.path.splitext(excel_file)
    excel_file2 = os.path.join(os.path.dirname(excel_file), f"{base_name}{suffix}{ext}")
    df.to_excel(excel_file2, index=False)
    print(f"File '{excel_file2}' updated successfully.")

file = "ExcelIndentFunctionTestFile.xlsx"
numbering = 2
indent = 1
indent_function(file, numbering, indent)