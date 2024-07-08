import pandas as pd
import numpy as np
import statistics
from datetime import datetime


from parameters import excel_file_path_monthly
from parameters import year_end_mon

# excel_file_path_monthly = "AutoAncillaries Monthly.xlsx"

# year_end_mon = 202405

xls = pd.ExcelFile(excel_file_path_monthly)
sheet_names = xls.sheet_names

year_end_month_values = []
for year in range(2030,2012,-1):
    for month in range(12,0,-1):
        year_end_month_values.append(int(f"{year}{month:02}"))


for sheet_name in sheet_names:
    globals()[sheet_name] = pd.read_excel(excel_file_path_monthly,sheet_name=sheet_name)

def find_latest_year_end(sheet_names):
    latest_year_end = None
    for sheet_name in sheet_names:
        df = globals()[sheet_name]
        max_year_end = df['Year & Month'].max()
        if latest_year_end is None or max_year_end > latest_year_end:
            latest_year_end = max_year_end
    return latest_year_end

if pd.isna(year_end_mon):
    year_end_mon = find_latest_year_end(sheet_names)


def adjust_year_end(year,month,offset):

    new_month = month + offset
    new_year = year

    while new_month < 1:
        new_month +=12
        new_year -=1

    while new_month > 12:
        new_month -=12
        new_year +=1 

    return new_year*100 + new_month


for sheet_name in sheet_names:
    df = globals()[sheet_name]

    year_end = year_end_mon
    available_year_end = None
    for y_end in year_end_month_values:
        if y_end in df['Year & Month'].values:
            available_year_end = y_end
            break

    if (available_year_end <= year_end_mon):
        year_end = available_year_end 

    latest_year_end = year_end

    
    latest_year = latest_year_end // 100
    latest_month = latest_year_end % 100


    latest_minus_1_year_end = adjust_year_end(latest_year,latest_month,-1)
    latest_minus_13_year_end = adjust_year_end(latest_year,latest_month,-13)
    

    closing_minus_1 = df[df['Year & Month'] == latest_minus_1_year_end]['Close'].values
    closing_minus_13 = df[df['Year & Month'] == latest_minus_13_year_end]['Close'].values
    
    twelve_month_price_momentum = (closing_minus_1/closing_minus_13)-1

    df['12 month price momentum'] = np.nan
    df.loc[df['Year & Month'] == latest_year_end,'12 month price momentum'] = twelve_month_price_momentum
    


momentum_values=[]

for sheet_name in sheet_names:
    df = globals()[sheet_name]
    year_end = year_end_mon
    available_year_end = None
    for y_end in year_end_month_values:
        if y_end in df['Year & Month'].values:
            available_year_end = y_end
            break

    if (available_year_end <= year_end_mon):
        year_end = available_year_end 

    latest_year_end = year_end

    filtered_df = df[df['Year & Month'] == latest_year_end]
    for index,row in filtered_df.iterrows():
        company_name = row['Company Name']
        momentum_value = row['12 month price momentum']
        momentum_values.append((company_name,momentum_value))


momentum_only = [ value[1] for value in momentum_values]

mean = statistics.mean(momentum_only)

std_dev = statistics.stdev(momentum_only)

ZScore_Values = []
for sheet_name in sheet_names:
    df = globals()[sheet_name]

    year_end = year_end_mon
    available_year_end = None
    for y_end in year_end_month_values:
        if y_end in df['Year & Month'].values:
            available_year_end = y_end
            break

    if (available_year_end <= year_end_mon):
        year_end = available_year_end

    latest_year_end = year_end

    Zscore = (df[df['Year & Month']== latest_year_end]['12 month price momentum'] - mean)/std_dev
    company_name = df[df['Year & Month'] == latest_year_end]['Company Name']
    ZScore_Values.append((company_name,Zscore))


# clean the ZScore_Values array
cleaned_data = [(item[0].values[0],item[1].values[0]) for item in ZScore_Values]

# make the cleaned data into a dataframe
ZScore_df_12mom  = pd.DataFrame(cleaned_data,columns=['Company Name','ZScore 12 month price momentum'])
ZScore_df_12mom['ZScore 12 month price momentum'] = ZScore_df_12mom['ZScore 12 month price momentum'].clip(lower=-4,upper=4)

# # export to excel
# timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# new_name = f'12mom_{latest_year_end}_{timestamp}'

# with pd.ExcelWriter("output.xlsx",engine="openpyxl", mode='a') as writer:
#     ZScore_df_12mom.to_excel(writer, index=False, header=True, sheet_name=new_name)


# # adjust the cell width when exporting 
#     workbook = writer.book
#     worksheet = writer.sheets[new_name]

#     for column in worksheet.columns:
#         max_length =0 
#         column_letter = column[0].column_letter

#         for cell in column:
#             try:
#                 if len(str(cell.value))> max_length:
#                     max_length = len(cell.value)
#             except:
#                 pass
#         adjusted_width = max_length + 2
#         worksheet.column_dimensions[column_letter].width = adjusted_width

print(f'12 month price momentum ZScore Calculation success for {year_end_mon}')