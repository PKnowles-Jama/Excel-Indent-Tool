# Excel Indent Function

import pandas as pd
import os
import re
import sys

def indent_function(excel_file, heading_column, indent_column):
    """
    Reads an Excel file, applies indentation to a specified column
    based on values in another column, and saves the modified DataFrame
    to a new Excel file with an '_indented' suffix.

    Args:
        excel_file (str): The name of the Excel file to manipulate.
        heading_column (str or int): The name or index of the column
                                     containing the numerical indent information.
        indent_column (str or int): The name or index of the column
                                    whose values are to be indented.

    Returns:
        tuple: A tuple containing:
            - str: The path to the newly created indented Excel file, or an empty string if an error occurred.
            - list: A list of strings containing all output messages from the function.
    """
    output = []
    excel_file2 = "" # Initialize excel_file2 for error return

    try:
        # Create a dataframe from the excel file
        df = pd.read_excel(excel_file)
        output.append(f"Successfully read '{excel_file}'.")
    except FileNotFoundError:
        output.append(f"Error: File '{excel_file}' not found.")
        return "", output
    except Exception as e:
        output.append(f"Error reading Excel file '{excel_file}': {e}")
        return "", output

    numbering_series = None
    indent_series = None

    # Identify the column to be used for number of indents
    if isinstance(heading_column, str):
        if heading_column in df.columns:
            numbering_series = df[heading_column]
        else:
            output.append(f"Error: Heading column '{heading_column}' not found in the Excel file.")
            return "", output
    elif isinstance(heading_column, int):
        if 0 <= heading_column < len(df.columns):
            numbering_series = df.iloc[:, heading_column]
        else:
            output.append(f"Error: Heading column index {heading_column} is out of bounds.")
            return "", output
    else:
        output.append("Error: 'heading_column' must be a string (column name) or an integer (column index).")
        return "", output

    # Identify the column to be indented
    if isinstance(indent_column, str):
        if indent_column in df.columns:
            indent_series = df[indent_column]
        else:
            output.append(f"Error: Indent column '{indent_column}' not found in the Excel file.")
            return "", output
    elif isinstance(indent_column, int):
        if 0 <= indent_column < len(df.columns):
            indent_series = df.iloc[:, indent_column]
            # If indent_column is an int, we need to get the actual column name for updating df
            indent_column_name = df.columns[indent_column]
        else:
            output.append(f"Error: Indent column index {indent_column} is out of bounds.")
            return "", output
    else:
        output.append("Error: 'indent_column' must be a string (column name) or an integer (column index).")
        return "", output

    indented_values = []
    for i in range(len(df)):
        try:
            # Ensure the value is converted to a string before attempting int conversion, just in case
            indent_level = int(str(numbering_series.iloc[i]))
            indent_str = "    " * indent_level # Using 4 spaces for an indent
            indented_value = f"{indent_str}{indent_series.iloc[i]}"
            indented_values.append(indented_value)
        except (ValueError, TypeError):
            output.append(f"Warning: Row {i+1}: Invalid or missing numeric value in heading column ('{heading_column}'). Skipping indentation for this row and keeping original value.")
            indented_values.append(indent_series.iloc[i])

    # Update the dataframe
    if isinstance(indent_column, int):
        df[indent_column_name] = indented_values
    else:
        df[indent_column] = indented_values


    # Save the changes
    suffix = "_indented"
    base_name, ext = os.path.splitext(excel_file)
    excel_file2 = os.path.join(os.path.dirname(excel_file), f"{base_name}{suffix}{ext}")

    try:
        df.to_excel(excel_file2, index=False)
        output.append(f"File '{excel_file2}' updated successfully.")
        return output
    except Exception as e:
        output.append(f"Error saving Excel file '{excel_file2}': {e}")
        return "", output

def calculate_indents_and_save_new_excel(excel_file_name: str, heading_column: str = 'Heading') -> tuple:
    """
    Reads an Excel file (assumed to be in the same directory as the script),
    calculates the number of indents for each entry in a specified heading column,
    appends these indents as a new column to the DataFrame, and then saves
    the modified DataFrame to a new Excel file with a '_new' suffix.

    The indent calculation logic is as follows:
    - For numbered headings (e.g., '1. Title', '1.1 Subtitle', '1.1.1 Sub-Subtitle'):
      The indent is determined by counting the number of dots in the numbering prefix.
      Example:
        '1. Title' -> 0 indents
        '1.1 Subtitle' -> 1 indent
        '1.1.1 Sub-Subtitle' -> 2 indents
    - For non-numbered text (e.g., 'Requirement Text'):
      The indent is one more than the indent of the last encountered numbered heading.

    Args:
        excel_file_name (str): The name of the Excel file (e.g., 'my_data.xlsx').
                               It is assumed this file is in the current working directory.
        heading_column (str, optional): The name of the column in the Excel file
                                        that contains the headings. Defaults to 'Heading'.

    Returns:
        tuple: A tuple containing:
            - str: The path to the newly created Excel file, or an empty string if an error occurred.
            - list: A list of strings containing all output messages from the function.
            - int: The column index of the newly added 'Calculated Indents' column, or -1 if not created.
            - int: The column index of the 'Heading' column, or -1 if not found.
    """
    output = []
    df = None
    output_excel_file_name = ""
    new_column_index = -1
    heading_column_index = -1

    try:
        df = pd.read_excel(excel_file_name)
        output.append(f"Successfully read '{excel_file_name}'.")
    except FileNotFoundError:
        output.append(f"Error: File '{excel_file_name}' not found in the current directory.")
        return "", output, new_column_index, heading_column_index
    except Exception as e:
        output.append(f"Error reading Excel file '{excel_file_name}': {e}")
        return "", output, new_column_index, heading_column_index

    if heading_column not in df.columns:
        output.append(f"Error: '{heading_column}' column not found in the Excel file.")
        return "", output, new_column_index, heading_column_index

    calculated_indents = []
    last_numbered_heading_indent = -1

    full_numeric_prefix_pattern = re.compile(r'^(\d+(\.\d+)*)')

    for index, row in df.iterrows():
        heading = str(row[heading_column]).strip()
        full_numeric_prefix_match = full_numeric_prefix_pattern.match(heading)

        if full_numeric_prefix_match:
            prefix = full_numeric_prefix_match.group(1)
            current_indent = prefix.count('.')
            last_numbered_heading_indent = current_indent
            calculated_indents.append(current_indent)
        else:
            if last_numbered_heading_indent == -1:
                calculated_indents.append(0)
            else:
                calculated_indents.append(last_numbered_heading_indent + 1)

    df['Calculated Indents'] = calculated_indents

    try:
        new_column_index = df.columns.get_loc('Calculated Indents')
        heading_column_index = df.columns.get_loc(heading_column)
    except KeyError as e:
        output.append(f"Error getting column index: {e}. This should not happen after successful column creation/check.")


    name, ext = os.path.splitext(excel_file_name)
    output_excel_file_name = f"{name}_new{ext}"

    try:
        df.to_excel(output_excel_file_name, index=False)
        output.append(f"Successfully saved results to '{output_excel_file_name}'.")
        return output_excel_file_name, new_column_index, heading_column_index, output
    except Exception as e:
        output.append(f"Error saving Excel file '{output_excel_file_name}': {e}")
        return "", output