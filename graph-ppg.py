import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from scipy.signal import find_peaks
from scipy.signal import firwin, lfilter

# Your PPG data (replace this with your actual data)
#ppg_data = np.array([725.8568325, 731.6966823, 741.4994097, 753.4205411, 765.7147667, 776.3369113, 784.0977841, 788.6888146, 791.3361767, 795.0445089, 803.6577076, 820.2171174, 847.1109634, 884.337585, 928.9320121, 976.5535012, 1021.80787, 1059.535598, 1085.642314, 1097.316207, 1094.752688, 1079.851119, 1055.408832, 1025.097765, 990.9727901, 954.9738225, 918.55497, 883.2919208, 850.9820219, 823.1552671, 800.5373676, 783.3135878, 770.9609096, 762.2993838, 756.1910837, 750.982314, 745.9210535, 739.931182, 732.5554561, 723.7244826, 714.1430894, 704.2513575, 695.2591554, 687.7102208, 682.0032751, 678.2744616, 676.5126354, 676.0580124, 676.5406634, 677.1906412, 677.602668, 678.1131746, 679.475648, 684.5025667, 696.4933106, 718.9987155, 752.488557, 796.4168637, 847.6114113, 900.5864327, 948.7752156, 986.5309981, 1010.140322, 1017.130413, 1007.836215, 984.6049085, 951.481139, 912.9204547, 872.9609869, 834.6495076, 800.1349883, 770.7785189, 746.8922381, 729.6858785, 719.363746, 716.0845278, 719.8636714, 729.7172261, 743.9912637, 760.9788672, 779.2745246, 797.1643982, 813.2011717, 827.0460059, 838.5641198, 847.4637572, 853.7620268, 858.2446149, 861.3029758, 863.3402377, 864.4889574, 864.4800671, 863.3663166, 860.9056022, 857.2545286, 851.9592831, 845.2815001, 837.894035, 830.4815689, 823.9149654, 820.4011679, 823.1060351, 834.8670456])


# file_name = input("Enter your file: ")

# #excel_file_path = 'C://Users/Atapy-Hardware04/Desktop/aoonjai project/Atapy-project-SR-Burapa/Atapy-project-SR-Burapa/excel/' + f'{file_name}.xlsx'
# #excel_file_path = './Atapy-project-SR-Burapa/Atapy-project-SR-Burapa/excel/' + 'R0' + f'{file_name}.xlsx'
# excel_file_path = './Data/Data from Burapa University/Data from Wristband/'+ 'R0' + f'{file_name}.xlsx'
# print(excel_file_path)

# event = ['Applying instrument','baseline recording','Giving instruction','Prep','Task1','Task2','STAI-S2','Recovery','STAI-S3','Debrief']
# #print(event)
# for index, data in enumerate(event):
#     print(f"event num {index+1}: {data}")
 
# event_num = int(input("Enter your're want event: "))
# print(f"Select: {event[event_num-1]}")

# #sheet_name = 'Debrief' #'Baseline recording'#'Debrief'
# # Read the Excel file
# df = pd.read_excel(excel_file_path,sheet_name=event[event_num-1])
post = 'post1'
file_name = '16-4parsed_data_graph.xlsx'
excel_file_path = 'C://Users/Atapy-Hardware04/Desktop/Subject-All/graph_pase/' + f'{file_name}'

print(f'\nfile name: {file_name}')

df = pd.read_excel(excel_file_path, sheet_name='Sheet1')
num_rows = df.shape[1]

ppg_sel = df.iloc[:, 1]
#ppg_sel = df.iloc[0:3000, 0]
print(len(ppg_sel))
ppg_data = np.array(ppg_sel)
print(ppg_data)

freq = 100
Fs  = 4000
N_samples = len(ppg_sel)
ts = np.arange(0, N_samples) / Fs

# Step 2: Create a low-pass FIR filter
N_ORDERS = 64   # Set the FIR Filter order.
Wn = 2*freq/Fs  # Set the normalized cutoff frequency.
b = firwin(N_ORDERS, Wn, window='hamming')

# Step 3: Apply the FIR filter to the input signal.
# Method 1) Use the scipy.signal.lfilter() function for filtering.
fs = 100
cutoff_freq = 0.5  # Cutoff frequency of the filter (Hz), adjust as needed
numtaps = 50  # Number of taps for the filter
lpf_coefficients = firwin(numtaps, cutoff_freq / (fs / 2), pass_zero='lowpass')
# y_output1 = lfilter(lpf_coefficients, 1, ppg_data)
y_output1 = lfilter(b, 1, ppg_data)


# b: The numerator coefficients a
# a: The denominator coefficients of the filter's transfer function. 
# For FIR filters, a=1 (no feedback).
M = len(b)
x_input_padded = np.concatenate((np.zeros(M), ppg_data))
N = len(x_input_padded)
# Method 2) Compute the output using convolution sum.
y_output2 = np.zeros(N)
for i in range(M,N):
    y_output2[i] = np.sum(b * x_input_padded[i:i-M:-1])
# Remove the first M values from the output.
y_output2 = y_output2[M:N] 
y_output = y_output1 # Select the output


peaks_avg, _ = find_peaks(y_output2, height=0)
num_peaks = len(peaks_avg)
avg_before = np.mean(y_output2[peaks_avg - 1]) if num_peaks > 0 else 0

print("Number of peaks:", num_peaks)
print("Average of peaks before:", avg_before)

#print(ppg_data)

# # Find peaks in the first derivative (these correspond to peaks in the original PPG signal)
min_height = avg_before - (avg_before *0.10)
max_height = avg_before + (avg_before *0.10)
peaks, _ = find_peaks(y_output2, height=(min_height,max_height))  # Adjust height as needed


# # Calculate R-R intervals (time between successive peaks)

rr_intervals = np.diff(peaks)
average_rr_interval = np.median(rr_intervals)
peaks_diff, _ = find_peaks(y_output2, width=average_rr_interval)

print('avg rr',average_rr_interval)

# peaks_re = peaks
# print('peaks_re:',peaks_re)
# print('Peaks:',peaks)
# for index in range(len(rr_intervals)):
#     if average_rr_interval + (average_rr_interval * 0.2) >= rr_intervals[index]:
#         peaks[index+1] = 0
#     #    peaks_re[index] = peaks[index]
#     #     #rr_intervals[index] = peaks[index]
#     # else:
#     #     peaks_re[index]

peaks_re = [0] * len(peaks)

print('peaks_re:',peaks_re)
print('Peaks:',peaks)
#peaks_re[0] = 0 
for index in range(len(rr_intervals)):
    if average_rr_interval + (average_rr_interval * 0.3) >= rr_intervals[index] and average_rr_interval - (average_rr_interval * 0.3) <= rr_intervals[index]:
        peaks_re[index] = peaks[index]
    #    peaks_re[index] = peaks[index]
    #     #rr_intervals[index] = peaks[index]
    else:
        peaks_re[index+1] = None

re_value = [-1 if x == 0 or x == None else x for x in peaks_re]
peak_filter = np.array(re_value)

# print("R-R Intervals:", rr_intervals)

# print('peaks_re:',peak_filter)
# print('Peaks:',peaks)

# print('interval : ',len(rr_intervals),'peaks',len(peaks))

# Time-domain HRV parameters
mean_rr = np.mean(rr_intervals)  # Mean RR interval
sdnn = np.std(rr_intervals)  # Standard deviation of RR intervals (SDNN)
rmssd = np.sqrt(np.mean(np.diff(rr_intervals) ** 2))  # Root mean square of successive differences (RMSSD)

# print("Mean RR Interval:", mean_rr, "ms")
# print("SDNN:", sdnn, "ms")
# print("RMSSD:", rmssd, "ms")

# # Calculate Heart rate
# average_rr_interval = np.mean(rr_intervals)
# print("avg_rr: ",average_rr_interval)
# heart_rate = 6000 / average_rr_interval
# print("Heart Rate (bpm):", heart_rate)

# Plot the PPG signal and detected peaks for visualization
# Step 4: Plot the input and filtered signals
plt.figure(figsize=(10, 4))
# plt.subplot(3, 1, 1)
plt.plot(ts, ppg_data)
plt.title(f'GSR Data {post}')
plt.xlabel('Time (ms)')
plt.ylabel('Amplitude')
plt.grid(True)
plt.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1, wspace=0.2, hspace=0.75)
# # plt.subplot(3, 1, 2)
# # plt.plot(ts, y_output1)
# # plt.title('Filtered PPG')
# # plt.xlabel('Time (ms)')
# # plt.ylabel('Amplitude')
# # plt.grid(True)
# # plt.subplot(3, 1, 3)
# # plt.plot(ts, y_output2)
# # plt.title('Filtered Output Signal (2)')
# # plt.xlabel('Time (s)')
# # plt.ylabel('Amplitude')
# # plt.grid(True)
# plt.figure(figsize=(10, 4))
# plt.title('Find peak')
# plt.plot(y_output2, label='PPG Signal')
# plt.plot(peaks, y_output2[peaks], 'r.', markersize=10, label='Detected Peaks')
# plt.plot(peak_filter, y_output2[peak_filter], 'g.', markersize=8, label='Detected Use')
# plt.xlabel('Time (ms)')
# plt.ylabel('PPG Signal')
# plt.legend()
# plt.grid(True)
# plt.subplots_adjust(left=0.1, right=0.9, top=0.9, 
#                     bottom=0.1, wspace=0.2, hspace=0.75)
plt.show()



