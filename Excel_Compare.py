import pandas as pd
import os

def generate_differences_for_all_cells(folder_path):
    # Find all Excel files in the folder
    all_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

    # Ensure there are at least two Excel files in the folder
    if len(all_files) < 2:
        raise ValueError("The folder must contain at least two Excel files.")

    # Use the first two files in the folder
    file1 = os.path.join(folder_path, all_files[0])
    file2 = os.path.join(folder_path, all_files[1])

    print(f"Comparing files:\nFile 1: {file1}\nFile 2: {file2}")

    # Read the data into DataFrames
    df1 = pd.read_excel(file1, dtype=str)  # Read everything as strings for uniformity
    df2 = pd.read_excel(file2, dtype=str)

    # Ensure both DataFrames have the same column names
    all_columns = df1.columns.union(df2.columns)  # Combine all columns
    df1 = df1.reindex(columns=all_columns)  # Reindex File 1
    df2 = df2.reindex(columns=all_columns)  # Reindex File 2

    # Reset index to align rows
    df1.reset_index(drop=True, inplace=True)
    df2.reset_index(drop=True, inplace=True)

    # Expand DataFrames to have the same shape by filling missing rows with empty strings
    max_rows = max(len(df1), len(df2))
    df1 = df1.reindex(index=range(max_rows), fill_value="")
    df2 = df2.reindex(index=range(max_rows), fill_value="")

    # Create a DataFrame to store differences
    differences = pd.DataFrame(0, index=range(max_rows), columns=all_columns)

    # Compare cell by cell
    for col in all_columns:
        differences[col] = df1[col].where(df1[col] != df2[col], other="0")

    # Check if there are no differences
    if differences.eq("0").all().all():
        print("No differences found between the two files.")
        return

    # List columns with differences
    columns_with_differences = [col for col in all_columns if not differences[col].eq("0").all()]
    print(f"Columns with differences: {', '.join(columns_with_differences)}")

    # Define output file name as differences_only.xlsx
    output_file_name = "differences_only.xlsx"
    output_file = os.path.join(folder_path, output_file_name)

    # Delete the file if it already exists
    if os.path.exists(output_file):
        os.remove(output_file)
        print(f"Existing file '{output_file}' deleted.")

    # Write the differences to a new Excel file
    with pd.ExcelWriter(output_file) as writer:
        differences.to_excel(writer, index=False, sheet_name="Differences")

    print(f"The file containing all cell differences is saved as '{output_file}'.")

# Example usage
folder_path = input("Enter the folder path containing the Excel files: ")
generate_differences_for_all_cells(folder_path)