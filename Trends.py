import pandas as pd
import numpy as np
from pathlib import Path
import re
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns
from collections import defaultdict

class SharePointAuditAnalyzer:
    def __init__(self, data_folder_path):
        """
        Initialize the analyzer with the folder containing Excel files
        
        Args:
            data_folder_path (str): Path to folder containing Excel audit log files
        """
        self.data_folder = Path(data_folder_path)
        self.audit_data = {}
        self.user_activity_summary = {}
        self.monthly_trends = {}
        
    def load_excel_files(self):
        """Load all Excel files from the data folder"""
        excel_files = list(self.data_folder.glob("*.xlsx")) + list(self.data_folder.glob("*.xls"))
        
        if not excel_files:
            raise FileNotFoundError(f"No Excel files found in {self.data_folder}")
        
        print(f"Found {len(excel_files)} Excel files")
        
        for file_path in excel_files:
            # Extract month from filename (assuming format like "SharePoint_Audit_2024_01.xlsx")
            month_match = re.search(r'(\d{4})[_-](\d{1,2})', file_path.name)
            if month_match:
                month_key = f"{month_match.group(1)}-{month_match.group(2).zfill(2)}"
            else:
                # Use filename as month key if pattern not found
                month_key = file_path.stem
            
            print(f"Processing file: {file_path.name} -> Month: {month_key}")
            
            try:
                # Load all sheets from the Excel file
                excel_data = pd.read_excel(file_path, sheet_name=None)
                self.audit_data[month_key] = excel_data
                print(f"  Loaded {len(excel_data)} sheets")
                
                # Print sheet names for reference
                for sheet_name in excel_data.keys():
                    print(f"    - {sheet_name}: {len(excel_data[sheet_name])} rows")
                    
            except Exception as e:
                print(f"Error loading {file_path.name}: {str(e)}")
                continue
    
    def identify_user_column(self, df):
        """
        Identify the user column from common SharePoint audit log column names
        
        Args:
            df (DataFrame): The dataframe to analyze
            
        Returns:
            str: The name of the user column
        """
        common_user_columns = [
            'User', 'UserName', 'User Name', 'UserId', 'User ID',
            'UserPrincipalName', 'UPN', 'Actor', 'ActorName',
            'Email', 'UserEmail', 'ModifiedBy', 'CreatedBy'
        ]
        
        df_columns = df.columns.tolist()
        
        # Check for exact matches first
        for col in common_user_columns:
            if col in df_columns:
                return col
        
        # Check for partial matches (case insensitive)
        for col in df_columns:
            for user_col in common_user_columns:
                if user_col.lower() in col.lower():
                    return col
        
        # If no match found, return the first column (assuming it might be user)
        if df_columns:
            print(f"Warning: Could not identify user column. Using first column: {df_columns[0]}")
            return df_columns[0]
        
        return None
    
    def analyze_user_activity(self):
        """Analyze user activity across all months and operations"""
        print("\n=== Analyzing User Activity ===")
        
        # Initialize data structures
        user_monthly_activity = defaultdict(lambda: defaultdict(int))
        operation_user_activity = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
        monthly_unique_users = defaultdict(set)
        
        for month, sheets_data in self.audit_data.items():
            print(f"\nProcessing month: {month}")
            
            for sheet_name, df in sheets_data.items():
                if df.empty:
                    continue
                
                # Identify user column
                user_column = self.identify_user_column(df)
                if not user_column:
                    print(f"  Skipping sheet {sheet_name} - no user column found")
                    continue
                
                print(f"  Processing sheet: {sheet_name} (User column: {user_column})")
                
                # Clean user data
                users = df[user_column].dropna().astype(str).str.strip()
                users = users[users != '']  # Remove empty strings
                
                if users.empty:
                    print(f"    No valid users found in sheet {sheet_name}")
                    continue
                
                # Count activities per user
                user_counts = users.value_counts()
                
                for user, count in user_counts.items():
                    user_monthly_activity[user][month] += count
                    operation_user_activity[sheet_name][user][month] += count
                    monthly_unique_users[month].add(user)
                
                print(f"    Found {len(user_counts)} unique users, {len(users)} total activities")
        
        # Store results
        self.user_monthly_activity = dict(user_monthly_activity)
        self.operation_user_activity = dict(operation_user_activity)
        self.monthly_unique_users = {month: list(users) for month, users in monthly_unique_users.items()}
        
        return self.user_monthly_activity
    
    def generate_summary_report(self):
        """Generate a comprehensive summary report"""
        print("\n=== Generating Summary Report ===")
        
        # Overall statistics
        all_months = sorted(self.monthly_unique_users.keys())
        total_unique_users = set()
        
        for users in self.monthly_unique_users.values():
            total_unique_users.update(users)
        
        summary = {
            'total_months_analyzed': len(all_months),
            'months_analyzed': all_months,
            'total_unique_users': len(total_unique_users),
            'monthly_stats': {}
        }
        
        for month in all_months:
            month_users = self.monthly_unique_users[month]
            total_activities = sum(
                sum(activities.get(month, 0) for activities in user_data.values())
                for user_data in self.user_monthly_activity.values()
            )
            
            summary['monthly_stats'][month] = {
                'unique_users': len(month_users),
                'total_activities': total_activities,
                'avg_activities_per_user': total_activities / len(month_users) if month_users else 0
            }
        
        self.summary_stats = summary
        return summary
    
    def create_user_trends_report(self, output_file='user_trends_report.xlsx'):
        """Create detailed user trends report in Excel format"""
        print(f"\n=== Creating User Trends Report: {output_file} ===")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            
            # 1. Monthly Summary Sheet
            months = sorted(self.monthly_unique_users.keys())
            summary_data = []
            
            for month in months:
                stats = self.summary_stats['monthly_stats'][month]
                summary_data.append({
                    'Month': month,
                    'Unique Users': stats['unique_users'],
                    'Total Activities': stats['total_activities'],
                    'Avg Activities per User': round(stats['avg_activities_per_user'], 2)
                })
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Monthly Summary', index=False)
            
            # 2. User Activity Trends Sheet
            user_trends_data = []
            
            for user, monthly_data in self.user_monthly_activity.items():
                row = {'User': user}
                total_activities = 0
                active_months = 0
                
                for month in months:
                    activities = monthly_data.get(month, 0)
                    row[f'{month} Activities'] = activities
                    total_activities += activities
                    if activities > 0:
                        active_months += 1
                
                row['Total Activities'] = total_activities
                row['Active Months'] = active_months
                row['Avg Activities per Month'] = round(total_activities / len(months), 2)
                
                user_trends_data.append(row)
            
            # Sort by total activities (descending)
            user_trends_data.sort(key=lambda x: x['Total Activities'], reverse=True)
            user_trends_df = pd.DataFrame(user_trends_data)
            user_trends_df.to_excel(writer, sheet_name='User Activity Trends', index=False)
            
            # 3. Top Users Sheet
            top_users_data = []
            for user_data in user_trends_data[:50]:  # Top 50 users
                top_users_data.append({
                    'User': user_data['User'],
                    'Total Activities': user_data['Total Activities'],
                    'Active Months': user_data['Active Months'],
                    'Avg Activities per Month': user_data['Avg Activities per Month']
                })
            
            top_users_df = pd.DataFrame(top_users_data)
            top_users_df.to_excel(writer, sheet_name='Top Users', index=False)
            
            # 4. Operation-wise Analysis
            for operation, user_data in self.operation_user_activity.items():
                operation_trends = []
                
                for user, monthly_data in user_data.items():
                    row = {'User': user}
                    total_activities = 0
                    
                    for month in months:
                        activities = monthly_data.get(month, 0)
                        row[f'{month}'] = activities
                        total_activities += activities
                    
                    row['Total'] = total_activities
                    operation_trends.append(row)
                
                # Sort by total activities
                operation_trends.sort(key=lambda x: x['Total'], reverse=True)
                operation_df = pd.DataFrame(operation_trends)
                
                # Clean sheet name for Excel
                clean_sheet_name = re.sub(r'[^\w\s-]', '', operation)[:31]
                operation_df.to_excel(writer, sheet_name=f'Op_{clean_sheet_name}', index=False)
            
            # 5. User Categorization Sheet
            user_categories = {
                'Power Users (>100 activities)': [],
                'Regular Users (20-100 activities)': [],
                'Occasional Users (5-20 activities)': [],
                'Light Users (1-5 activities)': []
            }
            
            for user_data in user_trends_data:
                total = user_data['Total Activities']
                user = user_data['User']
                
                if total > 100:
                    user_categories['Power Users (>100 activities)'].append(user_data)
                elif total >= 20:
                    user_categories['Regular Users (20-100 activities)'].append(user_data)
                elif total >= 5:
                    user_categories['Occasional Users (5-20 activities)'].append(user_data)
                else:
                    user_categories['Light Users (1-5 activities)'].append(user_data)
            
            category_summary = []
            for category, users in user_categories.items():
                category_summary.append({
                    'Category': category,
                    'User Count': len(users),
                    'Total Activities': sum(u['Total Activities'] for u in users),
                    'Avg Activities per User': round(sum(u['Total Activities'] for u in users) / len(users), 2) if users else 0
                })
            
            category_df = pd.DataFrame(category_summary)
            category_df.to_excel(writer, sheet_name='User Categories', index=False)
        
        print(f"Report saved to: {output_file}")
        return output_file
    
    def create_visualizations(self, output_folder='visualizations'):
        """Create visualizations for the analysis"""
        print(f"\n=== Creating Visualizations in {output_folder} ===")
        
        # Create output folder
        viz_folder = Path(output_folder)
        viz_folder.mkdir(exist_ok=True)
        
        # Set style
        plt.style.use('seaborn-v0_8')
        
        # 1. Monthly Activity Trends
        months = sorted(self.monthly_unique_users.keys())
        unique_users = [self.summary_stats['monthly_stats'][month]['unique_users'] for month in months]
        total_activities = [self.summary_stats['monthly_stats'][month]['total_activities'] for month in months]
        
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))
        
        # Unique users trend
        ax1.plot(months, unique_users, marker='o', linewidth=2, markersize=8)
        ax1.set_title('Monthly Unique Users Trend', fontsize=14, fontweight='bold')
        ax1.set_ylabel('Number of Unique Users')
        ax1.grid(True, alpha=0.3)
        
        # Total activities trend
        ax2.plot(months, total_activities, marker='s', linewidth=2, markersize=8, color='orange')
        ax2.set_title('Monthly Total Activities Trend', fontsize=14, fontweight='bold')
        ax2.set_ylabel('Total Activities')
        ax2.set_xlabel('Month')
        ax2.grid(True, alpha=0.3)
        
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(viz_folder / 'monthly_trends.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        # 2. Top Users Chart
        top_users = sorted(self.user_monthly_activity.items(), 
                          key=lambda x: sum(x[1].values()), reverse=True)[:15]
        
        users, activities = zip(*[(user, sum(data.values())) for user, data in top_users])
        
        plt.figure(figsize=(12, 8))
        bars = plt.barh(range(len(users)), activities)
        plt.yticks(range(len(users)), users)
        plt.xlabel('Total Activities')
        plt.title('Top 15 Most Active Users', fontsize=14, fontweight='bold')
        plt.gca().invert_yaxis()
        
        # Add value labels on bars
        for i, bar in enumerate(bars):
            width = bar.get_width()
            plt.text(width + max(activities) * 0.01, bar.get_y() + bar.get_height()/2, 
                    f'{int(width)}', ha='left', va='center')
        
        plt.tight_layout()
        plt.savefig(viz_folder / 'top_users.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        print(f"Visualizations saved in: {viz_folder}")
    
    def run_analysis(self, create_excel_report=True, create_visualizations=True):
        """Run the complete analysis pipeline"""
        print("=== SharePoint Audit Log Analysis ===")
        
        # Load data
        self.load_excel_files()
        
        # Analyze user activity
        self.analyze_user_activity()
        
        # Generate summary
        self.generate_summary_report()
        
        # Print summary to console
        print(f"\n=== Analysis Summary ===")
        print(f"Total months analyzed: {self.summary_stats['total_months_analyzed']}")
        print(f"Months: {', '.join(self.summary_stats['months_analyzed'])}")
        print(f"Total unique users: {self.summary_stats['total_unique_users']}")
        
        print(f"\nMonthly breakdown:")
        for month, stats in self.summary_stats['monthly_stats'].items():
            print(f"  {month}: {stats['unique_users']} users, {stats['total_activities']} activities")
        
        # Create reports
        if create_excel_report:
            self.create_user_trends_report()
        
        if create_visualizations:
            self.create_visualizations()
        
        print("\n=== Analysis Complete ===")
        return self.summary_stats

# Example usage
if __name__ == "__main__":
    # Initialize analyzer with your data folder path
    analyzer = SharePointAuditAnalyzer(r"C:\path\to\your\excel\files")
    
    # Run complete analysis
    try:
        summary = analyzer.run_analysis()
        print("\nAnalysis completed successfully!")
        print(f"Check the following files for results:")
        print("- user_trends_report.xlsx (detailed Excel report)")
        print("- visualizations/ folder (charts and graphs)")
        
    except Exception as e:
        print(f"Error during analysis: {str(e)}")
        print("Please check your file paths and data format.")
