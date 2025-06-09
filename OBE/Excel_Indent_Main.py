from Excel_Indent_Functions import indent_function, calculate_indents_and_save_new_excel
# from Excel_Indent_GUI import excel_indent_gui

# Need to pull these from the GUI
file = "IndentMe.xlsx"
heading = "Heading"

# Create Indent Column
file, numbering_column, heading_column = calculate_indents_and_save_new_excel(file, heading)

# Run Indenting Function
indent_function(file, numbering_column, heading_column)