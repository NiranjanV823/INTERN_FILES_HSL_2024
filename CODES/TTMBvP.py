import pandas as pd
import statistics
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')

from parameters import excel_file_path_quarterly
from parameters import excel_file_path_monthly
from parameters import year_end_const
from parameters import year_end_mon

# year_end_const = 202403
# year_end_mon = 202405


# excel_file_path_quarterly = "All DataSet.xlsx"
# excel_file_path_monthly = "AutoAncillaries Monthly.xlsx"

xls_qtr = pd.ExcelFile(excel_file_path_quarterly)
sheet_names_qtr = xls_qtr.sheet_names
xls_mon = pd.ExcelFile(excel_file_path_monthly)
sheet_names_mon = xls_mon.sheet_names

for sheet_name in sheet_names_qtr:
    df_name = f'wb1_{sheet_name}'
    globals()[df_name] = pd.read_excel(excel_file_path_quarterly,sheet_name=sheet_name)

def find_latest_year_end(sheet_names):
    latest_year_end = None
    for sheet_name in sheet_names:
        df_name = f'wb1_{sheet_name}'
        df = globals()[df_name]
        max_year_end = df['Year End'].max()
        if latest_year_end is None or max_year_end > latest_year_end:
            latest_year_end = max_year_end
    return latest_year_end


if pd.isna(year_end_const):
    year_end_const = find_latest_year_end(sheet_names_qtr)



for sheet_name in sheet_names_mon:
    df_name = f'wb2_{sheet_name}'
    globals()[df_name] = pd.read_excel(excel_file_path_monthly, sheet_name= sheet_name)


def find_latest_year_end(sheet_names):
    latest_year_end = None
    for sheet_name in sheet_names:
        df_name = f'wb2_{sheet_name}'
        df = globals()[df_name]
        max_year_end = df['Year & Month'].max()
        if latest_year_end is None or max_year_end > latest_year_end:
            latest_year_end = max_year_end
    return latest_year_end

if pd.isna(year_end_mon):
    year_end_mon = find_latest_year_end(sheet_names_mon)



for sheet_name in sheet_names_qtr:
    df_name = f'wb1_{sheet_name}'
    df = globals()[df_name]
    df['Share Capital']=df['Share Capital'].fillna(method = 'ffill')
    target_column = 'Reserves & Surplus'
    other_column = 'Adjusted Profit After Extra-ordinary item'

    for i in range(1,len(df)):
        if pd.isna(df.loc[i,target_column]):
            df.loc[i,target_column] = df.loc[i-1,target_column]+df.loc[i,other_column]
    

    df['Book Value'] =df['Share Capital']+df['Reserves & Surplus']


# year end values which can exist in this space ie yearmonth 202303 type
year_end_values = [] 

for year in range(2030,2012,-1):
    for month in [12,9,6,3]:
        year_end_values.append(int(f"{year}{month:02}"))



book_values =[]

for sheet_name in sheet_names_qtr:
    df_name = f'wb1_{sheet_name}'
    df = globals()[df_name]

    year_end = year_end_const
    available_year_end = None
    for y_end in year_end_values:
        if y_end in df['Year End'].values:
            available_year_end = y_end
            break

    if (available_year_end <= year_end_const):
        year_end = available_year_end  


    filtered_df = df[df['Year End'] == year_end]
    for index,row in filtered_df.iterrows():
        company_name = row['Company Name']
        book_value = row['Book Value']
        book_values.append((company_name,book_value))


year_end_month_values = []
for year in range(2030,2012,-1):
    for month in range(12,0,-1):
        year_end_month_values.append(int(f"{year}{month:02}"))


marketcap_values=[]
for sheet_name in sheet_names_mon:
    df_name = f'wb2_{sheet_name}'
    df = globals()[df_name]

    year_end = year_end_mon
    available_year_end = None
    for y_end in year_end_month_values:
        if y_end in df['Year & Month'].values:
            available_year_end = y_end
            break

    if (available_year_end <= year_end_mon):
        year_end = available_year_end  
    filtered_df = df[df['Year & Month'] == year_end]
    for index,row in filtered_df.iterrows():
        company_name = row['Company Name']
        marketcap_value = row['Market Cap']
        marketcap_values.append((company_name,marketcap_value))



book_values_df = pd.DataFrame(book_values,columns=['Company Name','Book Value'])
marketcap_values_df  = pd.DataFrame(marketcap_values,columns=['Company Name','Market Cap'])
result_df = pd.merge(book_values_df,marketcap_values_df,on='Company Name')
result_df['Bv/P'] = result_df['Book Value']/result_df['Market Cap']

mean = statistics.mean(result_df['Bv/P'])
std_dev = statistics.stdev(result_df['Bv/P'])
result_df['ZScore Bv/P'] = (result_df['Bv/P']-mean)/std_dev

ZScore_df_TTMBvP = result_df[['Company Name','ZScore Bv/P']]
ZScore_df_TTMBvP['ZScore Bv/P'] = ZScore_df_TTMBvP['ZScore Bv/P'].clip(lower=-4,upper=4)


# # export to excel
# timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# new_name = f'TTM_Bvp_{year_end_const}_{timestamp}'

# with pd.ExcelWriter("output.xlsx",engine="openpyxl", mode='a') as writer:
#     ZScore_df_TTMBvP.to_excel(writer, index=False, header=True, sheet_name=new_name)


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


print(f'TTM Book/Price Z Score Calculation Success for {year_end_const}')