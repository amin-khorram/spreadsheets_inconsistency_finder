import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import os

# Function to preprocess project numbers
def preprocess_projects(project):
    return str(project).split()[0]

# Function to process each spreadsheet separately
def process_spreadsheet(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Convert 'Ship Date' column to datetime and remove time component
    df['Ship Date'] = pd.to_datetime(df['Ship Date']).dt.date

    # Preprocess project numbers
    df['Project'] = df['Project'].apply(preprocess_projects)

    # Group by 'Project' and sum 'Units' for each project
    grouped_df = df.groupby('Project', as_index=False).agg({'Ship Date': 'first', 'Units': 'sum'})

    # Create a set of unique project numbers
    project_set = set(grouped_df['Project'])

    return grouped_df, project_set

# Function to compare ship dates and unit columns between two DataFrames
def compare_spreadsheets(df1, df2):
    # Merge DataFrames on 'Project' column
    merged_df = pd.merge(df1, df2, on='Project', how='outer')
    # Print the columns of the merged DataFrame
    print("Columns in merged DataFrame:", merged_df.columns)
    # Check for inconsistencies
    inconsistent_rows = merged_df[(merged_df['Units_x'] != merged_df['Units_y']) | (merged_df['Ship Date_x'] != merged_df['Ship Date_y'])]

    return inconsistent_rows

# Function to create a new spreadsheet with inconsistent items highlighted in red
# Function to create a new spreadsheet with inconsistent items listed one after the other
def create_output_spreadsheet(inconsistent_rows):

    # Sort inconsistent rows based on 'Ship Date_x'

    inconsistent_rows_sorted = inconsistent_rows.sort_values(by='Ship Date_x')

    # Create a new workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active

    # Write headers
    ws.append(["Project", "Ship Date", "Units"])

    # Write data
    for idx, row in inconsistent_rows_sorted.iterrows():
        ws.append([row['Project'], row['Ship Date_x'], row['Units_x']])

    # Save the workbook
    output_file = os.path.join(os.path.expanduser("~"), "Desktop", "output1.xlsx")
    wb.save(output_file)

# Main code
file_path1 = 'C:\\Users\\amink\\Documents\\rahul test.xlsx'
file_path2 = 'C:\\Users\\amink\\Documents\\spider test.xlsx'

# Process each spreadsheet separately
df1, project_set1 = process_spreadsheet(file_path1)
df2, project_set2 = process_spreadsheet(file_path2)

# Compare ship dates and unit columns between the two spreadsheets
inconsistent_rows = compare_spreadsheets(df1, df2)

# Create a new spreadsheet with inconsistent items highlighted in red
create_output_spreadsheet(inconsistent_rows)
