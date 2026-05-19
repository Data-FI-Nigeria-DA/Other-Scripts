import pandas as pd
import os

def combine_excel_files(input_folder, output_folder, output_filename="Week_W19_combined_data.xlsx"):
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # list of all Excel files in the input folder
    excel_files = [os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.endswith(('.xls', '.xlsx', 'csv'))]

    if not excel_files:
        print("No Excel files found in the specified input folder.")
        return

    all_data = []

    # Loop through each file, read it into a DataFrame, and append to the list
    for file in excel_files:
        try:
            if file.endswith('.csv'):
                data = pd.read_csv(file, encoding='latin1', on_bad_lines='skip')
            elif file.endswith('.xlsx'):
                data = pd.read_excel(file, engine='openpyxl')
            else:
                continue
            all_data.append(data)
        

            data['Filename'] = os.path.basename(file)
            combined_df = pd.concat(all_data, ignore_index=True)
        except Exception as e:
            print(f"Error processing file {file}: {e}")


    combined_df['ProjectName'] = combined_df['Filename'].str.split('_').str[0]

    # Define the output path
    output_path = os.path.join(output_folder, output_filename)

    # Save the combined DataFrame to a new Excel file
    combined_df.to_excel(output_path, index=False)

    print(f"Successfully combined {len(excel_files)} files into '{output_path}'.")


if __name__ == "__main__":
    
    input_folder_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Centralsync Upload Monitoring/Week_W19'
    output_folder_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Centralsync Upload Monitoring/output'

    combine_excel_files(input_folder_path, output_folder_path)
