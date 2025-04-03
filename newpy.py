import pandas as pd
from datetime import datetime
import os
import glob

# Get previous month and year for filename
previous_month = datetime.now().replace(day=1) - pd.Timedelta(days=1)
report_month_year = previous_month.strftime("%b%Y")  # e.g. "Mar2025"

# Define sites and file mappings
sites = ['site1', 'site2', 'site3', 'site4', 'site5']
file_mappings = {
    'FileViewed': ['*FileViewed*'],
    'FileAccessed': ['*FileAccessed*'], 
    'FileDownloaded': ['*FileDownloaded*']
}

for site in sites:
    # Create Excel filename
    excel_file = f"{site} {report_month_year}.xlsx"
    
    # Create Excel file with separate sheets
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        sheets_created = 0
        
        for sheet_name, patterns in file_mappings.items():
            # Find matching CSV file
            csv_file = None
            for pattern in patterns:
                full_pattern = f"{site}{pattern}.csv"
                matches = glob.glob(full_pattern)
                if matches:
                    csv_file = matches[0]
                    break
            
            if not csv_file:
                print(f"Warning: No file found for {site} {sheet_name}")
                continue
            
            try:
                # Read CSV file
                df = pd.read_csv(csv_file)
                
                # Write to Excel with exact sheet name
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                sheets_created += 1
                
                print(f"Added sheet: {sheet_name} (from {os.path.basename(csv_file)})")
            except Exception as e:
                print(f"Error processing {csv_file}: {str(e)}")
    
    if sheets_created > 0:
        print(f"\n✅ Created {excel_file} with {sheets_created} sheets\n")
    else:
        print(f"\n⚠️ No sheets created for {site}\n")
        os.remove(excel_file)  # Remove empty file

print("All processing complete!")
