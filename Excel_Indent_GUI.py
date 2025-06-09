# Excel Indent GUI

import sys
import os
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, QLineEdit, QFileDialog, QLabel, QTextEdit, QHBoxLayout)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QScreen
from Excel_Indent_Functions import indent_function, calculate_indents_and_save_new_excel

class ExcelProcessorGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Indentation Generator") # Title of app in header bar
        
        # Open app in the center of the screen & set size
        screen = QApplication.primaryScreen() # Scrape the correct screen to open on
        screen_geometry = screen.geometry() # Determine the primary screen's geometry
        window_width = 800 # Width of the app
        window_height = 100 # Height of the app
        x = (screen_geometry.width() - window_width) // 2 # Calculate the halfway width
        y = (screen_geometry.height() - window_height) // 2 # Calculate the halfway height
        self.setGeometry(x, y, window_width, window_height) # Set the app's opening location and size

        # Set the icon
        script_dir = os.path.dirname(os.path.abspath(__file__)) # Get the file path for this code
        icon_path = os.path.join(script_dir, 'jama_logo_icon.png') # Add the icon's file name to the path
        self.setWindowIcon(QIcon(icon_path)) # Add the icon to the app's header bar

        # Initialize user input variables
        self.file_path = "" # The file to be indented
        self.heading_column_name = "" # The column to base the indenting on

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()
        
        # File Selection Section
        file_selection_layout = QHBoxLayout()
        self.select_file_button = QPushButton("Select .xlsx File")
        self.select_file_button.clicked.connect(self.select_excel_file)
        file_selection_layout.addWidget(self.select_file_button)
        
        self.file_path_label = QLabel("No file selected")
        file_selection_layout.addWidget(self.file_path_label)
        main_layout.addLayout(file_selection_layout)

        # Heading Column Input Section
        heading_layout = QHBoxLayout()
        heading_label = QLabel("Heading Column Name:")
        heading_layout.addWidget(heading_label)
        self.heading_column_input = QLineEdit()
        self.heading_column_input.setPlaceholderText("e.g., 'Product Name'")
        self.heading_column_input.textChanged.connect(self.update_heading_column_name)
        heading_layout.addWidget(self.heading_column_input)
        main_layout.addLayout(heading_layout)

        # Run Button
        self.run_button = QPushButton("Run Processing")
        self.run_button.clicked.connect(self.run_processing)
        self.run_button.setEnabled(False) # Initially disabled
        main_layout.addWidget(self.run_button)

        # Output Console (initially hidden)
        self.output_console = QTextEdit()
        self.output_console.setReadOnly(True)
        self.output_console.setVisible(False) # Hide initially
        main_layout.addWidget(self.output_console)

        self.setLayout(main_layout)

    def select_excel_file(self):
        #
        # This function does the following:
        #   Open file explorer.
        #   Collect user's file name & path.
        #   Only allow user to select .xlsx files
        #   Check if all inputs have been collected.
        #
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)")
        if file_path:
            self.file_path = file_path
            self.file_path_label.setText(f"Selected: {self.file_path}")
            self.check_enable_run_button()

    def update_heading_column_name(self, text):
        # 
        # This function does the following:
        #   Collect user's heading column name.
        #   Check if all inputs have been collected.
        #
        self.heading_column_name = text
        self.check_enable_run_button()

    def check_enable_run_button(self):
        # 
        # This function does the following:
        #   Checks if both the file and heading column inputs have been provided
        #   Enables the run button if both are provided
        #
        if self.file_path and self.heading_column_name:
            self.run_button.setEnabled(True)
        else:
            self.run_button.setEnabled(False)

    def run_processing(self):
        # 
        # Execute based off of user inputs when the run button is pressed. This includes the following:
        #   Opening a console to show any error messages
        #   Calling the related functions to execute
        #
        self.output_console.clear() # Clear previous output
        self.output_console.setVisible(True) # Show the console
        self.output_console.append("--- Starting Processing ---") # Processing start message
        file, numbering_column, heading_column, output1 = calculate_indents_and_save_new_excel(self.file_path, self.heading_column_name) # Call Indent Calulator
        self.output_console.append(f"Function 1 Output:\n{output1}") # Indent Calculator Output Messages
        root, ext = os.path.splitext(self.file_path)
        new_file_path = f"{root}_new{ext}"
        output2 = indent_function(new_file_path, numbering_column, heading_column) # Call Indenting Function
        self.output_console.append(f"\nFunction 2 Output:\n{output2}") # Indenting Function Output Messages
        self.output_console.append("--- Processing Finished ---") # Processing end message
        self.adjustSize() # Adjust window size to show console

if __name__ == "__main__":
    # Run the app when running this file
    app = QApplication(sys.argv)
    window = ExcelProcessorGUI()
    window.show()
    sys.exit(app.exec())

