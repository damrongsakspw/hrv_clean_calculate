# # import pandas as pd
# # import numpy as np

# # def parse_data(data,gsr, date_time):
# #     parsed_data = []
# #     for item,item2 in zip(data,gsr):
# #         # แยกข้อมูลด้วยเครื่องหมายคอมมา
# #         parts = item.split(',')
# #         parts2 = item2.split(',')
# #         try:
# #             # แปลงข้อมูลเป็นค่า R-R interval และลำดับเวลา
# #             ppg = float(parts[0].strip())
# #             gsr = int(parts2[0].strip())
# #             time_sequence = int(parts[1].strip())
# #             parsed_data.append((ppg,gsr, time_sequence,date_time))
# #         except ValueError as e:
# #             print(f"Error parsing item '{item}': {e}")
# #     return parsed_data

# # def save_to_excel(parsed_data, output_file):
# #     # สร้าง DataFrame จากข้อมูลที่ถูกแปลง
# #     df = pd.DataFrame(parsed_data, columns=['PPG','GSR', 'Time Sequence','Date Time'])
    
# #     # บันทึก DataFrame ลงในไฟล์ Excel
# #     df.to_excel(output_file, index=False)
# #     # print(f"Data saved to {output_file}")

# # for j in range(1,17):
# #     if j != 3:
# #         if j < 10 :
# #             file_name = f'Subject-00{j}.xlsx'
# #         else:
# #             file_name = f'Subject-0{j}.xlsx'
        
# #         excel_file_path = 'C://Users/Atapy-Hardware04\Desktop/Subject-All/' + f'{file_name}'

# #         # print(excel_file_path)

# #         # Read the Excel file
# #         sheet_name = '24 Hour'
# #         df = pd.read_excel(excel_file_path,sheet_name=sheet_name)
# #         # จำนวนแถวและคอลัมน์
# #         num_rows, num_cols = df.shape

# #         # print(f"Number of rows: {num_rows}")
# #         # print(f"Number of columns: {num_cols}")
# #         for i in range(0,num_rows):
# #             # print(i)
# #             selected_cell = df.iloc[i, 10]  # เลือกเซลล์ในแถวแรกและคอลัมน์แรก
# #             date_time = df.iloc[i, 15]
# #             selected_gsr = df.iloc[i, 9]  # เลือกเซลล์ในแถวแรกและคอลัมน์แรก

# #             # print(np.array(selected_cell))

            

# #             data_list = selected_cell.strip('[]').replace('"', '').split(',')
# #             data_gsr = selected_gsr.strip('[]').replace('"', '').split(',')
# #             print(f"Data list after split: {data_list}")
# #             # แบ่งข้อมูลที่ถูกแปลงเป็นลิสต์ของทูเพิล
# #             data_pairs = [data_list[i] + ',' + data_list[i+1] for i in range(0, len(data_list), 2)]
# #             gsr_pairs = [data_gsr[i] + ',' + data_gsr[i+1] for i in range(0, len(data_gsr), 2)]

# #             # แปลงข้อมูลเป็นลิสต์ของทูเพิล
# #             parsed_data = parse_data(data_pairs,gsr_pairs,date_time)

# #             # กำหนดชื่อไฟล์ Excel ใหม่
# #                        # Define new Excel file name
# #             output_file = f'./3-24pase-re/{j}-{i}parsed_data3-24-re.xlsx'

# #             # บันทึกข้อมูลลงในไฟล์ Excel ใหม่
# #             save_to_excel(parsed_data, output_file)
                    


# import pandas as pd
# import numpy as np

# def parse_data(data, gsr, date_time):
#     parsed_data = []
#     for item, item2 in zip(data, gsr):
#         # แยกข้อมูลด้วยเครื่องหมายคอมมา
#         parts = item.split(',')
#         parts2 = item2.split(',')
#         try:
#             # แปลงข้อมูลเป็นค่า PPG และ GSR และลำดับเวลา
#             ppg = float(parts[0].strip())
#             gsr_value = float(parts2[0].strip())
#             time_sequence = int(parts[1].strip())
#             parsed_data.append((ppg, gsr_value, time_sequence, date_time))
#         except ValueError as e:
#             print(f"Error parsing item '{item}' or '{item2}': {e}")
#     return parsed_data

# def save_to_excel(parsed_data, output_file):
#     # สร้าง DataFrame จากข้อมูลที่ถูกแปลง
#     df = pd.DataFrame(parsed_data, columns=['PPG', 'GSR', 'Time Sequence', 'Date Time'])
    
#     # บันทึก DataFrame ลงในไฟล์ Excel
#     df.to_excel(output_file, index=False)
#     print(f"Data saved to {output_file}")

# for j in range(1, 17):
#     if j != 3:
#         if j < 10:
#             file_name = f'Subject-00{j}.xlsx'
#         else:
#             file_name = f'Subject-0{j}.xlsx'
        
#         excel_file_path = f'C://Users/Atapy-Hardware04/Desktop/Subject-All/{file_name}'

#         # Read the Excel file
#         sheet_name = '24 Hour'
#         df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
#         # Get the number of rows
#         num_rows = df.shape[0]

#         for i in range(num_rows):
#             selected_cell = df.iloc[i, 10]  # Select the cell in the 11th column
#             date_time = df.iloc[i, 15]
#             selected_gsr = df.iloc[i, 9]  # Select the cell in the 10th column

#             # Process data and GSR
#             data_list = selected_cell.strip('[]').replace('"', '').split(',')
#             data_gsr = selected_gsr.strip('[]').replace('"', '').split(',')

#             # Debugging prints
#             print(f"Data list after split: {data_list}")
#             print(f"GSR list after split: {data_gsr}")

#             # Check if data lists are even-length
#             if len(data_list) % 2 != 0 or len(data_gsr) % 2 != 0:
#                 print(f"Warning: Odd length for data lists in row {i} of file {file_name}")
#                 continue

#             # Create pairs
#             data_pairs = [data_list[j] + ',' + data_list[j+1] for j in range(0, len(data_list), 2)]
#             gsr_pairs = [data_gsr[j] + ',' + data_gsr[j+1] for j in range(0, len(data_gsr), 2)]

#             # Parse data
#             parsed_data = parse_data(data_pairs, gsr_pairs, date_time)

#             # Define new Excel file name
#             output_file = f'./3-24pase-re/{j}-{i}parsed_data3-24-re.xlsx'

#             # Save data to new Excel file
#             save_to_excel(parsed_data, output_file)


import pandas as pd
import numpy as np

def parse_data(data, gsr, date_time):
    parsed_data = []
    for item, item2 in zip(data, gsr):
        # แยกข้อมูลด้วยเครื่องหมายคอมมา
        parts = item.split(',')
        parts2 = item2.split(',')
        try:
            # แปลงข้อมูลเป็นค่า PPG และ GSR และลำดับเวลา
            ppg = float(parts[0].strip())
            gsr_value = float(parts2[0].strip())
            time_sequence = int(parts[1].strip())
            parsed_data.append((ppg, gsr_value, time_sequence, date_time))
        except ValueError as e:
            print(f"Error parsing item '{item}' or '{item2}': {e}")
    return parsed_data

def save_to_excel(parsed_data, output_file):
    # สร้าง DataFrame จากข้อมูลที่ถูกแปลง
    df = pd.DataFrame(parsed_data, columns=['PPG', 'GSR', 'Time Sequence', 'Date Time'])
    
    # บันทึก DataFrame ลงในไฟล์ Excel
    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")

for j in range(1, 17):
    if j != 3:
        if j < 10:
            file_name = f'Subject-00{j}.xlsx'
        else:
            file_name = f'Subject-0{j}.xlsx'
        
        excel_file_path = f'C://Users/Atapy-Hardware04/Desktop/Subject-All/{file_name}'

        # Read the Excel file
        sheet_name = '24 Hour'
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        # Get the number of rows
        num_rows = df.shape[0]

        for i in range(num_rows):
            selected_cell = df.iloc[i, 10]  # Select the cell in the 11th column
            date_time = df.iloc[i, 15]
            selected_gsr = df.iloc[i, 9]  # Select the cell in the 10th column

            # Process data and GSR
            data_list = selected_cell.strip('[]').replace('"', '').split(',')
            data_gsr = selected_gsr.strip('[]').replace('"', '').split(',')

            # Debugging prints
            print(f"Data list after split: {data_list}")
            print(f"GSR list after split: {data_gsr}")

            # Check if data lists are even-length
            if len(data_list) % 2 != 0 or len(data_gsr) % 2 != 0:
                print(f"Warning: Odd length for data lists in row {i} of file {file_name}")

                # Handle odd-length lists
                # Skip the last item or use an appropriate strategy
                if len(data_list) % 2 != 0:
                    data_list = data_list[:-1]  # Remove last item
                if len(data_gsr) % 2 != 0:
                    data_gsr = data_gsr[:-1]  # Remove last item

            # Create pairs
            data_pairs = [data_list[j] + ',' + data_list[j+1] for j in range(0, len(data_list), 2)]
            gsr_pairs = [data_gsr[j] + ',' + data_gsr[j+1] for j in range(0, len(data_gsr), 2)]

            # Parse data
            parsed_data = parse_data(data_pairs, gsr_pairs, date_time)

            # Define new Excel file name
            output_file = f'./3-24pase-re/{j}-{i}parsed_data3-24-re.xlsx'

            # Save data to new Excel file
            save_to_excel(parsed_data, output_file)
