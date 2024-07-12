# import pandas as pd
# import numpy as np

# def calculate_sdnn(rr_intervals):
#     mean_rr = np.mean(rr_intervals)
#     sdnn = np.sqrt(np.mean((np.array(rr_intervals) - mean_rr) ** 2))
#     return sdnn

# # คำนวณ RMSSD
# def calculate_rmssd(rr_intervals):
#     rr_intervals = np.array(rr_intervals)
#     differences = np.diff(rr_intervals)
#     rmssd = np.sqrt(np.mean(differences ** 2))
#     return rmssd

# def process_data(data):
#     # แปลงข้อมูลเป็นลิสต์ของตัวเลข
#     data_list = [float(item) for item in data.split(',')]
    
#     # คำนวณค่าเฉลี่ยของข้อมูล
#     average = sum(data_list) / len(data_list)
#     print(f"average: {average}, 20 percents of average {0.2*average}")
    
#     # กรองข้อมูลที่เกินค่าเฉลี่ย
#     filtered_data = [item for item in data_list if item > average - 0.2*average and item < average +  0.2*average ]
    
#     return filtered_data

# def save_to_excel(data, output_file):
#     # สร้าง DataFrame จากข้อมูลที่ถูกกรอง
#     df = pd.DataFrame(data, columns=['Value'])
    
#     # บันทึก DataFrame ลงในไฟล์ Excel
#     df.to_excel(output_file, index=False)
#     print(f"Data saved to {output_file}")

# # รายการเพื่อเก็บข้อมูลทั้งหมดที่ต้องการเขียนลงในไฟล์ Excel
# results = []
# for j in range(1,17):
    
#     if j != 3:
#         if j < 10 :
#             file_name = f'Subject-00{j}.xlsx'
#         else:
#             file_name = f'Subject-0{j}.xlsx'
        
#         excel_file_path = 'C://Users/Atapy-Hardware04\Desktop/Subject-All/' + f'{file_name}'

#         # print(excel_file_path)

#         # Read the Excel file
#         sheet_name = '24 Hour'
#         df = pd.read_excel(excel_file_path,sheet_name=sheet_name)
#         # จำนวนแถวและคอลัมน์
#         num_rows, num_cols = df.shape

#         for i in range(0,num_rows):
#             file_name = f'{j}-{i}parsed_data24.xlsx'
        
#             excel_file_path = 'C://Users/Atapy-Hardware04\Desktop/Subject-All/24pase/' + f'{file_name}'

#             print(f'\nfile name:   {file_name}')

#             df = pd.read_excel(excel_file_path,sheet_name='Sheet1')

#             selected_cell = df.iloc[:,0]
#             # print(selected_cell)

#             # แปลงข้อมูลเป็นสตริงที่แยกด้วยเครื่องหมายคอมมา
#             data_as_str = ','.join(map(str, selected_cell))

#             # แปลงและกรองข้อมูล
#             filtered_data = process_data(data_as_str)
#             # print(filtered_data)

#             # คำนวณ SDNN และ RMSSD
#             sdnn = calculate_sdnn(filtered_data)
#             rmssd = calculate_rmssd(filtered_data)

#             print(f"SDNN: {sdnn}")
#             print(f"RMSSD: {rmssd}")

#             # เพิ่มข้อมูลลงในรายการผลลัพธ์
#             results.append({
#                 'Identifier': f'{j}{i+1}',
#                 'RMSSD': rmssd,
#                 'SDNN': sdnn
#             })

#             # คำนวณ SDNN



#             # # กำหนดชื่อไฟล์ Excel ใหม่
#             # output_file = '{i}-{j}filtered_data.xlsx'

#             # # บันทึกข้อมูลที่กรองแล้วลงในไฟล์ Excel ใหม่
#             # save_to_excel(filtered_data, output_file)

# # สร้าง DataFrame จากรายการผลลัพธ์
# results_df = pd.DataFrame(results)

# # บันทึก DataFrame ลงในไฟล์ Excel
# output_file = 'C://Users/Atapy-Hardware04\\Desktop/Subject-All/24pase/24summary_results.xlsx'
# results_df.to_excel(output_file, index=False)

# print(f"Summary results saved to {output_file}")
                    




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
    print(f"average: {average}, 20 percents of average {0.2*average}")
    filtered_data = [item for item in data_list if item > average - 0.2*average and item < average + 0.2*average]
    return filtered_data

def save_to_excel(data, output_file):
    df = pd.DataFrame(data, columns=['Value'])
    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")

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
            file_name = f'{j}-{i}parsed_data3-24.xlsx'
            excel_file_path = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/3-24pase/' + f'{file_name}'

            print(f'\nfile name: {file_name}')

            df = pd.read_excel(excel_file_path, sheet_name='Sheet1')
            selected_cell = df.iloc[:, 0]

            data_as_str = ','.join(map(str, selected_cell))
            filtered_data = process_data(data_as_str)

            sdnn = calculate_sdnn(filtered_data)
            rmssd = calculate_rmssd(filtered_data)

            print(f"SDNN: {sdnn}")
            print(f"RMSSD: {rmssd}")

            results.append({
                'Identifier': f'{j}-{i+1}',
                'Round': i+1,  # เรียงลำดับรอบ
                'Subject': j,  # หมายเลข subject
                'RMSSD': rmssd,
                'SDNN': sdnn
            })

# สร้าง DataFrame จากรายการผลลัพธ์
results_df = pd.DataFrame(results)

# คำนวณค่าเฉลี่ยของ RMSSD และ SDNN สำหรับแต่ละ subject และรอบ
average_results = results_df.groupby('Subject').agg({
    'RMSSD': 'mean',
    'SDNN': 'mean'
}).reset_index()

# กำหนดชื่อไฟล์ Excel ใหม่
output_file = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/3-24pase/3-24summary_results_with_averages.xlsx'

# บันทึก DataFrame ลงในไฟล์ Excel
results_df.to_excel(output_file, sheet_name='Detailed Results', index=False)
# average_results.to_excel(output_file, sheet_name='Average Results', index=False)

print(f"Summary results with averages saved to {output_file}")
