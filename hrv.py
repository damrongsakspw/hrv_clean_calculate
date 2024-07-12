import pandas as pd
import numpy as np

def calculate_sdnn(rr_intervals):
    mean_rr = np.mean(rr_intervals)
    sdnn = np.sqrt(np.mean((np.array(rr_intervals) - mean_rr) ** 2))
    return sdnn

# คำนวณ RMSSD
def calculate_rmssd(rr_intervals):
    rr_intervals = np.array(rr_intervals)
    differences = np.diff(rr_intervals)
    rmssd = np.sqrt(np.mean(differences ** 2))
    return rmssd

def process_data(data):
    # แปลงข้อมูลเป็นลิสต์ของตัวเลข
    data_list = [float(item) for item in data.split(',')]
    
    # คำนวณค่าเฉลี่ยของข้อมูล
    average = sum(data_list) / len(data_list)
    print(f"average: {average}, 20 percents of average {0.2*average}")
    
    # กรองข้อมูลที่เกินค่าเฉลี่ย
    filtered_data = [item for item in data_list if item > average - 0.2*average and item < average +  0.2*average ]
    
    return filtered_data

def save_to_excel(data, output_file):
    # สร้าง DataFrame จากข้อมูลที่ถูกกรอง
    df = pd.DataFrame(data, columns=['Value'])
    
    # บันทึก DataFrame ลงในไฟล์ Excel
    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")

# รายการเพื่อเก็บข้อมูลทั้งหมดที่ต้องการเขียนลงในไฟล์ Excel
results = []
for j in range(1,17):
    if j != 3:
        for i in range(0,6):
            file_name = f'{j}-{i}parsed_data.xlsx'
        
            excel_file_path = 'C://Users/Atapy-Hardware04\Desktop/Subject-All/pase/' + f'{file_name}'

            print(f'\nfile name:   {file_name}')

            df = pd.read_excel(excel_file_path,sheet_name='Sheet1')

            selected_cell = df.iloc[:,0]
            # print(selected_cell)

            # แปลงข้อมูลเป็นสตริงที่แยกด้วยเครื่องหมายคอมมา
            data_as_str = ','.join(map(str, selected_cell))

            # แปลงและกรองข้อมูล
            filtered_data = process_data(data_as_str)
            # print(filtered_data)

            # คำนวณ SDNN และ RMSSD
            sdnn = calculate_sdnn(filtered_data)
            rmssd = calculate_rmssd(filtered_data)

            print(f"SDNN: {sdnn}")
            print(f"RMSSD: {rmssd}")

            # เพิ่มข้อมูลลงในรายการผลลัพธ์
            results.append({
                'Identifier': f'{j}{i+1}',
                'RMSSD': rmssd,
                'SDNN': sdnn
            })

            # คำนวณ SDNN



            # # กำหนดชื่อไฟล์ Excel ใหม่
            # output_file = '{i}-{j}filtered_data.xlsx'

            # # บันทึกข้อมูลที่กรองแล้วลงในไฟล์ Excel ใหม่
            # save_to_excel(filtered_data, output_file)

# สร้าง DataFrame จากรายการผลลัพธ์
results_df = pd.DataFrame(results)

# บันทึก DataFrame ลงในไฟล์ Excel
output_file = 'C://Users/Atapy-Hardware04\\Desktop/Subject-All/pase/summary_results.xlsx'
results_df.to_excel(output_file, index=False)

print(f"Summary results saved to {output_file}")
                    




