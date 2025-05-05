import pandas as pd

# Load the Excel file and read both sheets
file_path = 'your_excel_file.xlsx'  # Change this to your file name
sheet1 = pd.read_excel(file_path, sheet_name='Sheet1')
sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')

# Rename columns to prevent collision during merge (optional but helpful)
sheet1.columns = ['Sheet1_' + str(col) for col in sheet1.columns]
sheet2.columns = ['Sheet2_' + str(col) for col in sheet2.columns]

# Merge both sheets on the first column of each
merged = pd.merge(
    sheet1, sheet2,
    left_on=sheet1.columns[0],
    right_on=sheet2.columns[0],
    how='inner'  # Only matching rows
)

# Save result to a new Excel file
merged.to_excel('matched_rows_output.xlsx', index=False)

print("Done! Matching rows have been saved to 'matched_rows_output.xlsx'.")
