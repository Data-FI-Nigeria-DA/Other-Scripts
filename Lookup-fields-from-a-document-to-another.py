import pandas as pd
import os

# path
main_report_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Centralsync Upload Monitoring/main' 
report_update_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Centralsync Upload Monitoring/other' 
output_folder = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Centralsync Upload Monitoring/lookedup'

columns_to_merge = [
    'Facility Name',
    'Sync Status',
    'File Status',
    'Last Sync Date',
    'Sync Source',
    'Generation Type',
    'Data Period'
]

# Create output folder if it doesn't exist 
os.makedirs(output_folder, exist_ok=True)

# Process files 
def process_files():
    # Get list of files in report folder
    main_report_files = [f for f in os.listdir(main_report_path) if  (f.endswith('.xlsx') or f.endswith('.csv'))]

    print(f"Found {len(main_report_files)} one report files to process.")

    for main_report_file_name in main_report_files:
        ip_prefix = main_report_file_name.split('_')[0]
        report_update_file_name = f"{ip_prefix}_.xlsx"

        main_report_file_full_path = os.path.join(main_report_path, main_report_file_name)
        report_update_file_full_path = os.path.join(report_update_path, report_update_file_name)

        print(f"\nProcessing {main_report_file_name} and {report_update_file_name}...")

        # Check if corresponding file exists
        if not os.path.exists(report_update_file_full_path):
            print(f"Warning: Corresponding Radet file '{report_update_file_name}' not found for '{main_report_file_name}'. Skipping.")
            continue

        try:
            # Read Report
            if main_report_file_name.endswith('.xlsx'):
                df_main_report = pd.read_excel(main_report_file_full_path)
            else: # Must be .csv
                df_main_report = pd.read_csv(main_report_file_full_path)

            # df_main_report.rename(columns={'Facility Name': 'Facility Name'}, inplace=True)

            # Read Radet Report
            if report_update_file_name.endswith('.xlsx'):
                df_report_update = pd.read_excel(report_update_file_full_path)
            else: # Must be .csv 
                df_report_update = pd.read_csv(report_update_file_full_path)

            # df_report_update.rename(columns={'Facility Name': 'Facility Name'}, inplace=True)

            # Ensure 'Facility Name' column exists in both dataframes
            if 'Facility Name' not in df_main_report.columns:
                print(f"Error: 'Facility Name' column not found in '{main_report_file_name}'. Skipping.")
                continue
            if 'Facility Name' not in df_report_update.columns:
                print(f"Error: 'Facility Name' column not found in '{report_update_file_name}'. Skipping.")
                continue

            # Select relevant columns from report for merging
            df_report_update_subset = df_report_update[['Facility Name'] + [col for col in columns_to_merge if col != 'Facility Name']]

            # Perform the merge (left merge to keep all report rows)
            df_merged = pd.merge(df_main_report, df_report_update_subset, on='Facility Name', how='outer')

            # Define output file path
            output_file_name = f"{ip_prefix}_Merged_Report29.xlsx" 
            output_file_full_path = os.path.join(output_folder, output_file_name)

            # Save the merged dataframe to the new folder
            df_merged.to_excel(output_file_full_path, index=False)
            print(f"Successfully merged data and saved to: {output_file_full_path}")

        except Exception as e:
            print(f"An error occurred while processing {main_report_file_name}: {e}")

# Run the processing
if __name__ == "__main__":
    process_files()
    print("\nProcessing complete!")
