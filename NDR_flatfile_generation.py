import zipfile
import os
import shutil
import pandas as pd

# paths
zip_dir = 'C:/Users/DELL/Documents/DataFi/NDR Flatfile generation/Zipped_files/FY26Q2'
output_dir = 'C:/Users/DELL/Documents/DataFi/NDR Flatfile generation/Output_Directory/new_FY26Q2'

# Initialize lists
data_groups = {
    'PVLSDen': [], 'PVLSNum': [], 'TXCURR': [], 'TXNEW': [], 
    'TXRTT': [], 'TXML_Total': [], 'TXML_IIT_LT_3': [], 'TXML_IIT_3_5': [], 
    'TXML_IIT_6': [], 'TXML_Died': [], 'TXML_Stopped': [], 
    'TXML_TransferOut': [], 'TXML_IIT_Total': []
}

if not os.path.exists(output_dir):
    os.makedirs(output_dir)

for zip_filename in os.listdir(zip_dir):
    if zip_filename.endswith('.zip'):
        zip_filepath = os.path.join(zip_dir, zip_filename)
        # Handle index errors if zip name format varies
        parts = zip_filename.split('_')
        project_name = parts[-3] if len(parts) >= 3 else "Unknown"

        with zipfile.ZipFile(zip_filepath, 'r') as zip_ref:
            temp_dir = os.path.join(output_dir, zip_filename.replace('.zip', ''))
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            os.makedirs(temp_dir, exist_ok=True)
            zip_ref.extractall(temp_dir)

            for file in os.listdir(temp_dir):
                if file.endswith('.xlsx'):
                    file_path = os.path.join(temp_dir, file)
                    df = pd.read_excel(file_path, engine='openpyxl')
                    df['ProjectName'] = project_name
                    
                    # Basic indicator routing
                    if file.startswith('PVLSDen'): data_groups['PVLSDen'].append(df)
                    elif file.startswith('PVLSNum'): data_groups['PVLSNum'].append(df)
                    elif file.startswith('TXNEW'): data_groups['TXNEW'].append(df)
                    elif file.startswith('TXRTT'): data_groups['TXRTT'].append(df)
                    elif file.startswith('TXCURR'):
                        df_filtered = df[~df['DataElementName'].str.contains('ARVDispense', na=False)].copy()
                        data_groups['TXCURR'].append(df_filtered)
                    
                    # Complex TXML logic
                    elif file.startswith('TXML'):
                        if 'TXML_IIT' in file:
                            data_groups['TXML_IIT_Total'].append(df)
                            data_groups['TXML_IIT_LT_3'].append(df[df['CategoryOptionComboName'].str.contains('<3 Months', na=False)])
                            data_groups['TXML_IIT_3_5'].append(df[df['CategoryOptionComboName'].str.contains('3-5 Months', na=False)])
                            data_groups['TXML_IIT_6'].append(df[df['CategoryOptionComboName'].str.contains('6\+ Months', na=False)])
                        else:
                            data_groups['TXML_Total'].append(df)
                            data_groups['TXML_Died'].append(df[df['CategoryOptionComboName'].str.contains('Died', na=False)])
                            data_groups['TXML_TransferOut'].append(df[df['CategoryOptionComboName'].str.contains('Transferred Out', na=False)])
                            data_groups['TXML_Stopped'].append(df[df['CategoryOptionComboName'].str.contains('Refused \(Stopped\)', na=False)])

# pivoting function to prevent the "Length Mismatch" error
def pivot_group(data_list, indicator_name):
    if not data_list:
        return pd.DataFrame()
    
    combined = pd.concat(data_list, ignore_index=True)
    if combined.empty:
        return pd.DataFrame()

    # Create pivot
    pivot = combined.pivot_table(index=['ProjectName', 'OrgUnit'], values='Value', aggfunc='sum')
    
    # Check if pivot actually has data before renaming columns
    if not pivot.empty:
        pivot.columns = [indicator_name]
    return pivot

# Process indicators
pivoted_dfs = []
for label, list_df in data_groups.items():
    res = pivot_group(list_df, label)
    if not res.empty:
        pivoted_dfs.append(res)

# Merge all into one sheet
if pivoted_dfs:
    final_df = pivoted_dfs[0]
    for next_df in pivoted_dfs[1:]:
        final_df = final_df.join(next_df, how='outer')

    final_df.reset_index(inplace=True)
    
    # Final output
    output_file = os.path.join(output_dir, 'Consolidated_Report.xlsx')
    final_df.to_excel(output_file, sheet_name='Data_by_OrgUnit', index=False)
    print(f"File created successfully: {output_file}")
else:
    print("No data was processed. Check your file naming or directory paths.")