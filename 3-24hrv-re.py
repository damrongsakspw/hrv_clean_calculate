# # # # # import pandas as pd
# # # # # import numpy as np

# # # # # def calculate_sdnn(rr_intervals):
# # # # #     mean_rr = np.mean(rr_intervals)
# # # # #     sdnn = np.sqrt(np.mean((np.array(rr_intervals) - mean_rr) ** 2))
# # # # #     return sdnn

# # # # # # คำนวณ RMSSD
# # # # # def calculate_rmssd(rr_intervals):
# # # # #     rr_intervals = np.array(rr_intervals)
# # # # #     differences = np.diff(rr_intervals)
# # # # #     rmssd = np.sqrt(np.mean(differences ** 2))
# # # # #     return rmssd

# # # # # def process_data(data):
# # # # #     # แปลงข้อมูลเป็นลิสต์ของตัวเลข
# # # # #     data_list = [float(item) for item in data.split(',')]
    
# # # # #     # คำนวณค่าเฉลี่ยของข้อมูล
# # # # #     average = sum(data_list) / len(data_list)


# # # # #     print(f"average: {average}, 20 percents of average {0.2*average}")
    
# # # # #     # กรองข้อมูลที่เกินค่าเฉลี่ย
# # # # #     filtered_data = [item for item in data_list if item > average - 0.2*average and item < average +  0.2*average ]
    
# # # # #     return filtered_data

# # # # # def save_to_excel(data, output_file):
# # # # #     # สร้าง DataFrame จากข้อมูลที่ถูกกรอง
# # # # #     df = pd.DataFrame(data, columns=['Value'])
    
# # # # #     # บันทึก DataFrame ลงในไฟล์ Excel
# # # # #     df.to_excel(output_file, index=False)
# # # # #     print(f"Data saved to {output_file}")

# # # # # # รายการเพื่อเก็บข้อมูลทั้งหมดที่ต้องการเขียนลงในไฟล์ Excel
# # # # # results = []
# # # # # for j in range(1,17):
    
# # # # #     if j != 3:
# # # # #         if j < 10 :
# # # # #             file_name = f'Subject-00{j}.xlsx'
# # # # #         else:
# # # # #             file_name = f'Subject-0{j}.xlsx'
        
# # # # #         excel_file_path = 'C://Users/Atapy-Hardware04\Desktop/Subject-All/' + f'{file_name}'

# # # # #         # print(excel_file_path)

# # # # #         # Read the Excel file
# # # # #         sheet_name = '24 Hour'
# # # # #         df = pd.read_excel(excel_file_path,sheet_name=sheet_name)
# # # # #         # จำนวนแถวและคอลัมน์
# # # # #         num_rows, num_cols = df.shape

# # # # #         for i in range(0,num_rows):
# # # # #             file_name = f'{j}-{i}parsed_data24.xlsx'
        
# # # # #             excel_file_path = 'C://Users/Atapy-Hardware04\Desktop/Subject-All/24pase/' + f'{file_name}'

# # # # #             print(f'\nfile name:   {file_name}')

# # # # #             df = pd.read_excel(excel_file_path,sheet_name='Sheet1')

# # # # #             selected_cell = df.iloc[:,0]
# # # # #             # print(selected_cell)

# # # # #             # แปลงข้อมูลเป็นสตริงที่แยกด้วยเครื่องหมายคอมมา
# # # # #             data_as_str = ','.join(map(str, selected_cell))

# # # # #             # แปลงและกรองข้อมูล
# # # # #             filtered_data = process_data(data_as_str)
# # # # #             # print(filtered_data)

# # # # #             # คำนวณ SDNN และ RMSSD
# # # # #             sdnn = calculate_sdnn(filtered_data)
# # # # #             rmssd = calculate_rmssd(filtered_data)

# # # # #             print(f"SDNN: {sdnn}")
# # # # #             print(f"RMSSD: {rmssd}")

# # # # #             # เพิ่มข้อมูลลงในรายการผลลัพธ์
# # # # #             results.append({
# # # # #                 'Identifier': f'{j}{i+1}',
# # # # #                 'RMSSD': rmssd,
# # # # #                 'SDNN': sdnn
# # # # #             })

# # # # #             # คำนวณ SDNN



# # # # #             # # กำหนดชื่อไฟล์ Excel ใหม่
# # # # #             # output_file = '{i}-{j}filtered_data.xlsx'

# # # # #             # # บันทึกข้อมูลที่กรองแล้วลงในไฟล์ Excel ใหม่
# # # # #             # save_to_excel(filtered_data, output_file)

# # # # # # สร้าง DataFrame จากรายการผลลัพธ์
# # # # # results_df = pd.DataFrame(results)

# # # # # # บันทึก DataFrame ลงในไฟล์ Excel
# # # # # output_file = 'C://Users/Atapy-Hardware04\\Desktop/Subject-All/24pase/24summary_results.xlsx'
# # # # # results_df.to_excel(output_file, index=False)

# # # # # print(f"Summary results saved to {output_file}")
                    




# # # # import pandas as pd
# # # # import numpy as np

# # # # def calculate_sdnn(rr_intervals):
# # # #     mean_rr = np.mean(rr_intervals)
# # # #     sdnn = np.sqrt(np.mean((np.array(rr_intervals) - mean_rr) ** 2))
# # # #     return sdnn

# # # # def calculate_rmssd(rr_intervals):
# # # #     rr_intervals = np.array(rr_intervals)
# # # #     differences = np.diff(rr_intervals)
# # # #     rmssd = np.sqrt(np.mean(differences ** 2))
# # # #     return rmssd

# # # # def process_data(data):
# # # #     data_list = [float(item) for item in data.split(',')]
# # # #     average = sum(data_list) / len(data_list)
# # # #     print(f"average: {average}, 20 percents of average {0.2*average}")
# # # #     filtered_data = [item for item in data_list if item > average - 0.2*average and item < average + 0.2*average]
# # # #     return filtered_data

# # # # def save_to_excel(data, output_file):
# # # #     df = pd.DataFrame(data, columns=['Value'])
# # # #     df.to_excel(output_file, index=False)
# # # #     print(f"Data saved to {output_file}")

# # # # results = []
# # # # for j in range(1, 17):
# # # #     if j != 3:
# # # #         if j < 10:
# # # #             file_name = f'Subject-00{j}.xlsx'
# # # #         else:
# # # #             file_name = f'Subject-0{j}.xlsx'
        
# # # #         excel_file_path = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/' + f'{file_name}'

# # # #         sheet_name = '24 Hour'
# # # #         df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
# # # #         num_rows, num_cols = df.shape

# # # #         for i in range(0, num_rows):
# # # #             file_name = f'{j}-{i}parsed_data3-24.xlsx'
# # # #             excel_file_path = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/3-24pase/' + f'{file_name}'

# # # #             print(f'\nfile name: {file_name}')

# # # #             df = pd.read_excel(excel_file_path, sheet_name='Sheet1')
# # # #             selected_cell = df.iloc[:, 0]
            

# # # #             data_as_str = ','.join(map(str, selected_cell))
# # # #             filtered_data = process_data(data_as_str)

# # # #             sdnn = calculate_sdnn(filtered_data)
# # # #             rmssd = calculate_rmssd(filtered_data)

# # # #             print(f"SDNN: {sdnn}")
# # # #             print(f"RMSSD: {rmssd}")

# # # #             results.append({
# # # #                 'Identifier': f'{j}-{i+1}',
# # # #                 'Round': i+1,  # เรียงลำดับรอบ
# # # #                 'Subject': j,  # หมายเลข subject
# # # #                 'RMSSD': rmssd,
# # # #                 'SDNN': sdnn
# # # #             })

# # # # # สร้าง DataFrame จากรายการผลลัพธ์
# # # # results_df = pd.DataFrame(results)

# # # # # คำนวณค่าเฉลี่ยของ RMSSD และ SDNN สำหรับแต่ละ subject และรอบ
# # # # average_results = results_df.groupby('Subject').agg({
# # # #     'RMSSD': 'mean',
# # # #     'SDNN': 'mean'
# # # # }).reset_index()

# # # # # กำหนดชื่อไฟล์ Excel ใหม่
# # # # output_file = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/3-24pase/3-24summary_results_with_averages.xlsx'

# # # # # บันทึก DataFrame ลงในไฟล์ Excel
# # # # results_df.to_excel(output_file, sheet_name='Detailed Results', index=False)
# # # # # average_results.to_excel(output_file, sheet_name='Average Results', index=False)

# # # # print(f"Summary results with averages saved to {output_file}")

# # # import pandas as pd
# # # import numpy as np

# # # def calculate_sdnn(rr_intervals):
# # #     mean_rr = np.mean(rr_intervals)
# # #     sdnn = np.sqrt(np.mean((np.array(rr_intervals) - mean_rr) ** 2))
# # #     return sdnn

# # # def calculate_rmssd(rr_intervals):
# # #     rr_intervals = np.array(rr_intervals)
# # #     differences = np.diff(rr_intervals)
# # #     rmssd = np.sqrt(np.mean(differences ** 2))
# # #     return rmssd

# # # def process_data(data):
# # #     data_list = [float(item) for item in data.split(',')]
# # #     average = sum(data_list) / len(data_list)
# # #     print(f"average: {average}, 20 percents of average {0.2 * average}")
# # #     filtered_data = [item for item in data_list if item > average - 0.2 * average and item < average + 0.2 * average]
# # #     return filtered_data

# # # def save_to_excel(data, output_file):
# # #     df = pd.DataFrame(data, columns=['Value'])
# # #     df.to_excel(output_file, index=False)
# # #     print(f"Data saved to {output_file}")

# # # results = []
# # # for j in range(1, 17):
# # #     if j != 3:
# # #         if j < 10:
# # #             file_name = f'Subject-00{j}.xlsx'
# # #         else:
# # #             file_name = f'Subject-0{j}.xlsx'
        
# # #         excel_file_path = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/' + f'{file_name}'

# # #         sheet_name = '24 Hour'
# # #         df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
# # #         num_rows, num_cols = df.shape

# # #         for i in range(0, num_rows):
# # #             file_name = f'{j}-{i}parsed_data3-24.xlsx'
# # #             excel_file_path = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/3-24pase/' + f'{file_name}'

# # #             print(f'\nfile name: {file_name}')

# # #             df = pd.read_excel(excel_file_path, sheet_name='Sheet1')
# # #             selected_cell = df.iloc[:, 0]
# # #             date_time = df.iloc[:, 2].values[0]  # Extract the date-time value from the third column

# # #             data_as_str = ','.join(map(str, selected_cell))
# # #             filtered_data = process_data(data_as_str)

# # #             sdnn = calculate_sdnn(filtered_data)
# # #             rmssd = calculate_rmssd(filtered_data)

# # #             print(f"SDNN: {sdnn}")
# # #             print(f"RMSSD: {rmssd}")

# # #             results.append({
# # #                 'Identifier': f'{j}-{i + 1}',
# # #                 'Round': i + 1,  # Round number
# # #                 'Subject': j,  # Subject number
# # #                 'RMSSD': rmssd,
# # #                 'SDNN': sdnn,
# # #                 'DateTime': date_time  # Include the date-time information
# # #             })

# # # # Create DataFrame from results
# # # results_df = pd.DataFrame(results)

# # # # Calculate the average RMSSD and SDNN for each subject and include the date-time information
# # # average_results = results_df.groupby('Subject').agg({
# # #     'RMSSD': 'mean',
# # #     'SDNN': 'mean',
# # #     'DateTime': 'first'  # Use 'first' to include the first date-time value in the group
# # # }).reset_index()

# # # # Define new Excel file name
# # # output_file = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/3-24pase/3-24summary_results_with_averages.xlsx'

# # # # Save DataFrame to Excel file
# # # with pd.ExcelWriter(output_file) as writer:
# # #     results_df.to_excel(writer, sheet_name='Detailed Results', index=False)
# # #     average_results.to_excel(writer, sheet_name='Average Results', index=False)

# # # print(f"Summary results with averages saved to {output_file}")

# # import pandas as pd
# # import numpy as np

# # def calculate_sdnn(rr_intervals):
# #     mean_rr = np.mean(rr_intervals)
# #     sdnn = np.sqrt(np.mean((np.array(rr_intervals) - mean_rr) ** 2))
# #     return sdnn

# # def calculate_rmssd(rr_intervals):
# #     rr_intervals = np.array(rr_intervals)
# #     differences = np.diff(rr_intervals)
# #     rmssd = np.sqrt(np.mean(differences ** 2))
# #     return rmssd

# # def process_data(data):
# #     data_list = [float(item) for item in data.split(',')]
# #     average = sum(data_list) / len(data_list)
# #     print(f"average: {average}, 20 percents of average {0.2 * average}")
# #     filtered_data = [item for item in data_list if item > average - 0.2 * average and item < average + 0.2 * average]
# #     return filtered_data

# # def save_to_excel(data, output_file):
# #     df = pd.DataFrame(data, columns=['Value'])
# #     df.to_excel(output_file, index=False)
# #     print(f"Data saved to {output_file}")

# # results = []
# # for j in range(1, 17):
# #     if j != 3:
# #         if j < 10:
# #             file_name = f'Subject-00{j}.xlsx'
# #         else:
# #             file_name = f'Subject-0{j}.xlsx'
        
# #         excel_file_path = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/' + f'{file_name}'

# #         sheet_name = '24 Hour'
# #         df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
# #         num_rows, num_cols = df.shape

# #         for i in range(0, num_rows):
# #             file_name = f'{j}-{i}parsed_data3-24.xlsx'
# #             excel_file_path = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/3-24pase/' + f'{file_name}'

# #             print(f'\nfile name: {file_name}')

# #             df = pd.read_excel(excel_file_path, sheet_name='Sheet1')
# #             selected_cell = df.iloc[:, 0]
# #             date_time = pd.to_datetime(df.iloc[:, 2].values[0])  # Convert date-time to pandas datetime

# #             data_as_str = ','.join(map(str, selected_cell))
# #             filtered_data = process_data(data_as_str)

# #             sdnn = calculate_sdnn(filtered_data)
# #             rmssd = calculate_rmssd(filtered_data)

# #             print(f"SDNN: {sdnn}")
# #             print(f"RMSSD: {rmssd}")

# #             results.append({
# #                 'Identifier': f'{j}-{i + 1}',
# #                 'Round': i + 1,  # Round number
# #                 'Subject': j,  # Subject number
# #                 'RMSSD': rmssd,
# #                 'SDNN': sdnn,
# #                 'DateTime': date_time.timestamp()  # Convert datetime to timestamp
# #             })

# # # Create DataFrame from results
# # results_df = pd.DataFrame(results)

# # # Calculate the average RMSSD and SDNN for each subject and include the date-time information
# # average_results = results_df.groupby('Subject').agg({
# #     'RMSSD': 'mean',
# #     'SDNN': 'mean',
# #     'DateTime': 'first'  # Use 'first' to include the first timestamp in the group
# # }).reset_index()

# # # Define new Excel file name
# # output_file = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/3-24pase/3-24summary_results_with_averages2.xlsx'

# # # Save DataFrame to Excel file
# # with pd.ExcelWriter(output_file) as writer:
# #     results_df.to_excel(writer, sheet_name='Detailed Results', index=False)
# #     average_results.to_excel(writer, sheet_name='Average Results', index=False)

# # print(f"Summary results with averages saved to {output_file}")


# import pandas as pd
# import numpy as np

# def calculate_sdnn(rr_intervals):
#     mean_rr = np.mean(rr_intervals)
#     sdnn = np.sqrt(np.mean((np.array(rr_intervals) - mean_rr) ** 2))
#     return sdnn

# def calculate_rmssd(rr_intervals):
#     rr_intervals = np.array(rr_intervals)
#     differences = np.diff(rr_intervals)
#     rmssd = np.sqrt(np.mean(differences ** 2))
#     return rmssd

# def process_data(data):
#     data_list = [float(item) for item in data.split(',')]
#     average = sum(data_list) / len(data_list)
#     print(f"average: {average}, 20 percents of average {0.2 * average}")
#     filtered_data = [item for item in data_list if item > average - 0.2 * average and item < average + 0.2 * average]
#     return filtered_data

# def save_to_excel(data, output_file):
#     df = pd.DataFrame(data, columns=['Value'])
#     df.to_excel(output_file, index=False)
#     print(f"Data saved to {output_file}")

# results = []
# for j in range(1, 17):
#     if j != 3:
#         if j < 10:
#             file_name = f'Subject-00{j}.xlsx'
#         else:
#             file_name = f'Subject-0{j}.xlsx'
        
#         excel_file_path = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/' + f'{file_name}'

#         sheet_name = '24 Hour'
#         df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
#         num_rows, num_cols = df.shape

#         for i in range(0, num_rows):
#             file_name = f'{j}-{i}parsed_data3-24.xlsx'
#             excel_file_path = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/3-24pase/' + f'{file_name}'

#             print(f'\nfile name: {file_name}')

#             df = pd.read_excel(excel_file_path, sheet_name='Sheet1')
#             selected_cell = df.iloc[:, 0]
#             date_time = df.iloc[:, 2].values[0]  # Extract the date-time value from the third column

#             data_as_str = ','.join(map(str, selected_cell))
#             filtered_data = process_data(data_as_str)

#             sdnn = calculate_sdnn(filtered_data)
#             rmssd = calculate_rmssd(filtered_data)

#             print(f"SDNN: {sdnn}")
#             print(f"RMSSD: {rmssd}")

#             results.append({
#                 'Identifier': f'{j}-{i + 1}',
#                 'Round': i + 1,  # Round number
#                 'Subject': j,  # Subject number
#                 'RMSSD': rmssd,
#                 'SDNN': sdnn,
#                 'DateTime': date_time  # Include the date-time information
#             })

# # Create DataFrame from results
# results_df = pd.DataFrame(results)

# # Calculate the average RMSSD and SDNN for each subject and include the date-time information
# average_results = results_df.groupby('Subject').agg({
#     'RMSSD': 'mean',
#     'SDNN': 'mean',
#     'DateTime': 'first'  # Use 'first' to include the first date-time value in the group
# }).reset_index()

# # Define new Excel file name
# output_file = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/3-24pase/3-24summary_results_with_averages.xlsx'

# # Save DataFrame to Excel file
# with pd.ExcelWriter(output_file) as writer:
#     results_df.to_excel(writer, sheet_name='Detailed Results', index=False)
#     average_results.to_excel(writer, sheet_name='Average Results', index=False)

# print(f"Summary results with averages saved to {output_file}")

import pandas as pd
import numpy as np

def calculate_sdnn(rr_intervals):
    mean_rr = np.mean(rr_intervals)
    sdnn = np.sqrt(np.mean((np.array(rr_intervals) - mean_rr) ** 2))
    return sdnn

def calculate_rmssd(rr_intervals):
    rr_intervals = np.array(rr_intervals)
    differences = np.diff(rr_intervals)
    rmssd = np.sqrt(np.mean(differences ** 2))
    return rmssd

def process_data(data):
    data_list = [float(item) for item in data.split(',')]
    average = sum(data_list) / len(data_list)
    # print(f"average: {average}, 20 percents of average {0.2 * average}")
    # filtered_data = [item for item in data_list if item > average - 0.2 * average and item < average + 0.2 * average]
    return average

def save_to_excel(data, output_file):
    df = pd.DataFrame(data, columns=['Value'])
    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")

def process_date_time(date_time_str):
    # Convert DateTime with the specified format
    return pd.to_datetime(date_time_str, format='%d-%m-%YT%H:%M:%S', dayfirst=True)

results = []
for j in range(1, 17):
    if j != 3:
        if j < 10:
            file_name = f'Subject-00{j}.xlsx'
        else:
            file_name = f'Subject-0{j}.xlsx'
        
        excel_file_path = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/' + f'{file_name}'

        sheet_name = '24 Hour'
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        num_rows, num_cols = df.shape

        for i in range(0, num_rows):
            file_name = f'{j}-{i}parsed_data3-24-re.xlsx'
            excel_file_path = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/3-24pase-re/' + f'{file_name}'

            print(f'\nfile name: {file_name}')

            df = pd.read_excel(excel_file_path, sheet_name='Sheet1')
            selected_cell = df.iloc[:, 0]
            selected_cell2 = df.iloc[:, 1]
            date_time = df.iloc[:, 3].values[0]  # Extract the date-time value from the third column

            data_as_str = ','.join(map(str, selected_cell))
            data_as_str2 = ','.join(map(str, selected_cell2))
            PPG = process_data(data_as_str)
            GSR = process_data(data_as_str2)
            

            # sdnn = calculate_sdnn(filtered_data[0])
            # rmssd = calculate_rmssd(filtered_data[0])

            # avg_rr = filtered_data[1]

            # print(f"SDNN: {sdnn}")
            # print(f"RMSSD: {rmssd}")
            # print(f'avg:{avg_rr}')

            # Assuming the format is 'YYYY-MM-DD HH:MM:SS'
            date_time_format = '%Y-%m-%d %H:%M:%S'
            results.append({
                'Identifier': f'{j}-{i + 1}',
                'Round': i + 1,  # Round number
                'Subject': j,  # Subject number
                'GSR': GSR,
                'PPG': PPG,
                'DateTime': process_date_time(date_time)  # Ensure DateTime is a pandas datetime object
            })

# Create DataFrame from results
results_df = pd.DataFrame(results)

# Convert DateTime to a 3-hour period
results_df['3HourPeriod'] = results_df['DateTime'].dt.floor('3H')

# Calculate the average RMSSD and SDNN for each 3-hour period and subject
average_results = results_df.groupby(['Subject', '3HourPeriod']).agg({
    'GSR': 'mean',
    'PPG': 'mean'
}).reset_index()

# Define new Excel file name
output_file = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/3-24pase-re/3-24summary_results-re.xlsx'

# Save DataFrame to Excel file
with pd.ExcelWriter(output_file) as writer:
    results_df.to_excel(writer, sheet_name='Detailed Results', index=False)
    average_results.to_excel(writer, sheet_name='3Hour Averages', index=False)

print(f"Summary results with 3-hour averages saved to {output_file}")