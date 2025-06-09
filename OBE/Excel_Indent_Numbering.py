import pandas as pd
import re
import os

def calculate_indents_from_excel(excel_file_path: str) -> pd.DataFrame:
    """
    Reads an Excel file, calculates the number of indents for each entry in the
    'Heading' column, and returns a DataFrame with the original data and
    a new 'Calculated Indents' column.

    The indent calculation logic is as follows:
    - For numbered headings (e.g., '1. Title', '1.1 Subtitle', '1.1.1 Sub-Subtitle'):
      The indent is determined by counting the number of dots in the numbering
      prefix and subtracting 1.
      Example:
        '1. Title' -> 1 dot -> 1 - 1 = 0 indents
        '1.1 Subtitle' -> 2 dots -> 2 - 1 = 1 indent
        '1.1.1 Sub-Subtitle' -> 3 dots -> 3 - 1 = 2 indents
    - For non-numbered text (e.g., 'Requirement Text'):
      The indent is one more than the indent of the last encountered numbered heading.

    Args:
        excel_file_path (str): The full path to the Excel file.

    Returns:
        pd.DataFrame: A DataFrame containing the original data and the new
                      'Calculated Indents' column. Returns an empty DataFrame
                      if the file does not exist or the 'Heading' column is missing.
    """
    # Check if the file exists
    if not os.path.exists(excel_file_path):
        print(f"Error: File not found at '{excel_file_path}'")
        return pd.DataFrame()

    try:
        df = pd.read_excel(excel_file_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()

    if 'Heading' not in df.columns:
        print("Error: 'Heading' column not found in the Excel file.")
        return pd.DataFrame()

    calculated_indents = []
    last_numbered_heading_indent = -1 # Initialize with a value that indicates no numbered heading seen yet

    # Regex to match numbered headings like "1.", "1.1", "1.1.1" etc.
    # It captures the numbering part for dot counting.
    numbered_heading_pattern = re.compile(r'^(\d+(\.\d+)*)\s.*')

    for index, row in df.iterrows():
        heading = str(row['Heading']).strip() # Ensure heading is a string and remove leading/trailing whitespace

        match = numbered_heading_pattern.match(heading)

        if match:
            # It's a numbered heading
            numbering_part = match.group(1)
            # The indent is (number of dots in numbering_part) - 1
            # Example: "1." has 1 dot -> 0 indents
            # "1.1" has 2 dots -> 1 indent
            # "1.1.1" has 3 dots -> 2 indents
            current_indent = numbering_part.count('.') - 1
            last_numbered_heading_indent = current_indent
            calculated_indents.append(current_indent)
        else:
            # It's a non-numbered text (like "Requirement Text")
            # The indent is one more than the last numbered heading's indent
            # Handle the case where the first row might not be a numbered heading
            # (though the example implies it always starts with one)
            if last_numbered_heading_indent == -1:
                # If no numbered heading has been seen yet, default to 0 or handle as an error
                # Based on the example, this case should not be hit if the first entry is numbered.
                calculated_indents.append(0) # Default to 0 if no prior numbered heading
            else:
                calculated_indents.append(last_numbered_heading_indent + 1)

    df['Calculated Indents'] = calculated_indents
    return df