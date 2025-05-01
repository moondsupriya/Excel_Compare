Excel Files Comparison Tool

This Python project provides a script to compare two Excel files from a given folder and generate a new Excel file containing only the differences. If there are no differences, the script prints a message and does not create an output file. It also displays the columns with differences in the console for easy identification.
Features
- Compares two .xlsx files in a specified folder.
- Identifies cell-level differences and outputs them into a new file.
- Handles both numeric and non-numeric data.
- Outputs "0" for identical cells and the differing value for mismatched cells.
- Lists columns with differences in the console.
- Skips creating an output file if no differences are found and displays "No differences found."

Prerequisites
Before running the script, ensure you have the following installed:
Software Requirements:
- Python: Version 3.6 or higher.

Python Libraries:
- pandas: For reading and comparing Excel files.
- openpyxl: For handling .xlsx files.

To install the required libraries, run:
pip install pandas openpyxl

Usage
- Clone or download this project to your local machine.
- Place at least two .xlsx files in the folder you want to analyze.
- Run the script in a Python environment or terminal.

Steps:
- Open your terminal or Python editor.
- Run the script:python compare_excel_files.py

- Enter the folder path when prompted:Enter the folder path containing the Excel files: /path/to/your/folder

Script Workflow
- File Selection:- The script scans the specified folder for .xlsx files.
- If fewer than two files are found, it raises an error.

- Data Reading:- Reads the first two Excel files in the folder.
- Ensures both files have the same column structure.

- Comparison:- Compares each cell in the two files.
- Writes "0" for identical cells.
- Writes the differing value for mismatched cells.

- Output:- Saves the differences in a file named differences_only.xlsx.
- Lists columns with differences in the console.
- Skips creating the output file if no differences are found.


Example Console Output
Case: No Differences
Comparing files:
File 1: /path/to/your/folder/file1.xlsx
File 2: /path/to/your/folder/file2.xlsx
No differences found between the two files.

Case: Differences Found
Comparing files:
File 1: /path/to/your/folder/file1.xlsx
File 2: /path/to/your/folder/file2.xlsx
Columns with differences: Column1, Column3
The file containing all cell differences is saved as '/path/to/your/folder/differences_only.xlsx'.

Output File
If differences are found, the script generates an Excel file named differences_only.xlsx in the specified folder. This file includes:
- Column headers from the input files.
- "0" for identical cells.
- The value from the first file for mismatched cells.

Notes
- The script assumes that the first two .xlsx files in the folder are the ones to be compared.
- Only .xlsx files are supported.

License
This project is open-source and available for personal and commercial use.















