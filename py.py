import pandas as pd

# Define your input CSV files and output Excel file
csv_files = ['Fileaccessed.csv', 'pageviewed.csv']
excel_file = 'testpy.xlsx'

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
    for csv_file in csv_files:
        # Read each CSV file
        df = pd.read_csv(csv_file)
        
        # Get the sheet name by removing the .csv extension
        sheet_name = csv_file.replace('.csv', '')
        
        # Write the dataframe to the Excel file
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Successfully created {excel_file} with sheets for {', '.join(csv_files)}")
