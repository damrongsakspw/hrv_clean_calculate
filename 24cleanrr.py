import pandas as pd
import numpy as np

def parse_data(data):
    parsed_data = []
    for item in data:
        # แยกข้อมูลด้วยเครื่องหมายคอมมา
        parts = item.split(',')
        try:
            # แปลงข้อมูลเป็นค่า R-R interval และลำดับเวลา
            rr_interval = float(parts[0].strip())
            time_sequence = int(parts[1].strip())
            parsed_data.append((rr_interval, time_sequence))
        except ValueError as e:
            print(f"Error parsing item '{item}': {e}")
    return parsed_data

def save_to_excel(parsed_data, output_file):
    # สร้าง DataFrame จากข้อมูลที่ถูกแปลง
    df = pd.DataFrame(parsed_data, columns=['R-R Interval', 'Time Sequence'])
    
    # บันทึก DataFrame ลงในไฟล์ Excel
    df.to_excel(output_file, index=False)
    # print(f"Data saved to {output_file}")

for j in range(1,17):
    if j != 3:
        if j < 10 :
            file_name = f'Subject-00{j}.xlsx'
        else:
            file_name = f'Subject-0{j}.xlsx'
        
        excel_file_path = 'C://Users/Atapy-Hardware04\Desktop/Subject-All/' + f'{file_name}'

        # print(excel_file_path)

        # Read the Excel file
        sheet_name = '24 Hour'
        df = pd.read_excel(excel_file_path,sheet_name=sheet_name)
        # จำนวนแถวและคอลัมน์
        num_rows, num_cols = df.shape

        # print(f"Number of rows: {num_rows}")
        # print(f"Number of columns: {num_cols}")
        for i in range(0,num_rows):
            print(i)
            selected_cell = df.iloc[i, 11]  # เลือกเซลล์ในแถวแรกและคอลัมน์แรก
            # print(np.array(selected_cell))


            

            data_list = selected_cell.strip('[]').replace('"', '').split(',')
            print(f"Data list after split: {data_list}")
            # แบ่งข้อมูลที่ถูกแปลงเป็นลิสต์ของทูเพิล
            data_pairs = [data_list[i] + ',' + data_list[i+1] for i in range(0, len(data_list), 2)]

            # แปลงข้อมูลเป็นลิสต์ของทูเพิล
            parsed_data = parse_data(data_pairs)

            # กำหนดชื่อไฟล์ Excel ใหม่
            output_file = f'./24pase/{j}-{i}parsed_data24.xlsx'

            # บันทึกข้อมูลลงในไฟล์ Excel ใหม่
            save_to_excel(parsed_data, output_file)
                    




