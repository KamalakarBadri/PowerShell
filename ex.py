def generate_excel_reports(self):
    """Generate Excel reports from the CSV files for each site"""
    self.log("Starting Excel report generation...", Fore.CYAN)
    
    excel_files = []  # Track generated Excel files for upload
    
    for site_name in SITE_NAMES:
        # Create Excel filename with full month name
        excel_file = os.path.join(self.output_dir, f"{site_name} {self.REPORT_MONTH} {self.REPORT_YEAR}.xlsx")
        
        # Create Excel file with separate sheets
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            sheets_created = 0
            
            for operation in OPERATIONS:
                # Find matching CSV file
                pattern = os.path.join(self.output_dir, f"{site_name}_*_{operation}_*.csv")
                matches = glob.glob(pattern)
                
                if not matches:
                    self.log(f"Warning: No file found for {site_name} {operation}", Fore.YELLOW)
                    continue
                
                csv_file = matches[0]  # Take the first match
                
                try:
                    # Read CSV file
                    df = pd.read_csv(csv_file)
                    
                    # Write to Excel with operation as sheet name
                    df.to_excel(writer, sheet_name=operation, index=False)
                    
                    # Get the workbook and worksheet objects
                    workbook = writer.book
                    worksheet = writer.sheets[operation]
                    
                    # Define a regular format (no hyperlinks)
                    regular_format = workbook.add_format({
                        'font_color': 'black',
                        'underline': 0  # No underline
                    })
                    
                    # Apply regular format to all cells to remove hyperlink formatting
                    # This applies to the entire used range in the worksheet
                    if not df.empty:
                        max_row, max_col = df.shape
                        worksheet.set_column(0, max_col - 1, None, regular_format)
                    
                    sheets_created += 1
                    
                    self.log(f"Added sheet: {operation} (from {os.path.basename(csv_file)})", Fore.GREEN)
                except Exception as e:
                    self.log(f"Error processing {csv_file}: {str(e)}", Fore.RED, logging.ERROR)
        
            if sheets_created > 0:
                self.log(f"\n✅ Created {excel_file} with {sheets_created} sheets\n", Fore.GREEN)
                excel_files.append(excel_file)
            else:
                self.log(f"\n⚠️ No sheets created for {site_name}\n", Fore.YELLOW)
                try:
                    os.remove(excel_file)  # Remove empty file
                except Exception as e:
                    self.log(f"Failed to remove empty Excel file: {e}", Fore.YELLOW)

    self.log("Excel report generation complete!", Fore.GREEN)
    return excel_files
