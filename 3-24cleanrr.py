import pandas as pd
import numpy as np

def parse_data(data, date_time):
    parsed_data = []
    for item in data:
        # Split data by comma
        parts = item.split(',')
        try:
            # Convert data to R-R interval and time sequence
            rr_interval = float(parts[0].strip())
            time_sequence = int(parts[1].strip())
            parsed_data.append((rr_interval, time_sequence, date_time))
        except ValueError as e:
            print(f"Error parsing item '{item}': {e}")
    return parsed_data

def save_to_excel(parsed_data, output_file):
    # Create DataFrame from parsed data
    df = pd.DataFrame(parsed_data, columns=['R-R Interval', 'Time Sequence', 'Date Time'])

    # Save DataFrame to Excel file
    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")

for j in range(1, 17):
    if j != 3:
        if j < 10:
            file_name = f'Subject-00{j}.xlsx'
        else:
            file_name = f'Subject-0{j}.xlsx'
        
        excel_file_path = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/' + f'{file_name}'

        # Read the Excel file
        sheet_name = '24 Hour'
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        # Get the number of rows and columns
        num_rows, num_cols = df.shape

        for i in range(0, num_rows):
            selected_cell = df.iloc[i, 11]  # Select the cell in the 12th column
            date_time = df.iloc[i, 15]


            data_list = selected_cell.strip('[]').replace('"', '').split(',')
            # Split data into pairs
            data_pairs = [data_list[i] + ',' + data_list[i+1] for i in range(0, len(data_list), 2)]



            # Parse data and include datetime
            parsed_data = parse_data(data_pairs, date_time)

            # Define new Excel file name
            output_file = f'./3-24pase/{j}-{i}parsed_data3-24.xlsx'

            # Save data to new Excel file
            save_to_excel(parsed_data, output_file)
