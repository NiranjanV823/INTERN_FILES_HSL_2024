import numpy as np
import pandas as pd
import statistics
from scipy.stats import linregress
from datetime import datetime


# import the parameters needed from parameters.py
from parameters import year_end_const
from parameters import LTSPS_exclude_for_mean_stddev
from parameters import how_many_year_back
from parameters import excel_file_path_quarterly


# read the dataset file
xls = pd.ExcelFile(excel_file_path_quarterly)
sheet_names = xls.sheet_names


# year end values which can exist in this space ie yearmonth 202303 type
year_end_values = [] 

for year in range(2030,2012,-1):
    for month in [12,9,6,3]:
        year_end_values.append(int(f"{year}{month:02}"))


# make the sheets into dataframes
for sheet_name in sheet_names:
    globals()[sheet_name] = pd.read_excel(excel_file_path_quarterly,sheet_name=sheet_name)


#  to find the latest year end present in the excel sheet for each company
def find_latest_year_end(sheet_names):
    latest_year_end = None
    for sheet_name in sheet_names:
        df = globals()[sheet_name]
        max_year_end = df['Year End'].max()
        if latest_year_end is None or max_year_end > latest_year_end:
            latest_year_end = max_year_end
    return latest_year_end


if pd.isna(year_end_const):
    year_end_const = find_latest_year_end(sheet_names)




# calculate the TTM Sales 
for sheet_name in sheet_names:
    df = globals()[sheet_name]
    df['TTM Sales'] = df['Net Sales'].rolling(window =4).sum()



for sheet_name in sheet_names:
    df =globals()[sheet_name]
    
    year_end = year_end_const
    available_year_end = None
    for y_end in year_end_values:
        if y_end in df['Year End'].values:
            available_year_end = y_end
            break

    if (available_year_end <= year_end_const):
        year_end = available_year_end

    # calculates the consecutive five years 
    year_end_str =str(year_end)
    years = [int(year_end_str[:4])- i for i in range(how_many_year_back)]
    months = year_end_str[4:]
    required_year_ends = [str(year) + months for year in years]

    # takes the TTM Sales for the consecutive year end calculated above
    filtered_df =df[df['Year End'].astype(str).isin(required_year_ends)]
    filtered_df = filtered_df.sort_values(by= 'Year End')
    ttm_sales_value = filtered_df['TTM Sales'].values

    # calculate the slope of ttm sales v/s 0 1 2 3 4 
    x = np.arange(len(ttm_sales_value))
    slope,intercept,r_value,p_value,std_err = linregress(x,ttm_sales_value)
    df['Slope'] = np.nan
    df.loc[df['Year End'] == year_end, 'Slope'] = slope


    ttm_avg = statistics.mean(ttm_sales_value)
    df['SPS avg'] = np.nan
    df.loc[df['Year End'] == year_end, 'SPS avg'] = ttm_avg

    df['EGRO'] = np.nan
    egro = slope/ttm_avg
    df.loc[df['Year End'] == year_end,'EGRO'] = egro



egro_values = []

for sheet_name in sheet_names:
    df = globals()[sheet_name]

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
        egro_value = row['EGRO']
        egro_values.append((company_name,egro_value))


# extract only the egro values for mean and standard deviation calculations
egro_only = [value[1] for value in egro_values if value[0] not in LTSPS_exclude_for_mean_stddev]

# calculate standard deviation for egro 
std_dev = statistics.stdev(egro_only)

# calculate mean for egro
mean = statistics.mean(egro_only)


# calculate the zscore values
ZScore_Values=[]
for sheet_name in sheet_names:
    df = globals()[sheet_name]
    year_end = year_end_const
    available_year_end = None
    for y_end in year_end_values:
        if y_end in df['Year End'].values:
            available_year_end = y_end
            break

    if (available_year_end <= year_end_const):
        year_end = available_year_end  
    Zscore = (df[df['Year End']==year_end]['EGRO'] - mean)/std_dev
    company_name = df[df['Year End'] == year_end]['Company Name']
    ZScore_Values.append((company_name,Zscore))


# clean the ZScore_Values array
cleaned_data = [(item[0].values[0],item[1].values[0]) for item in ZScore_Values]

# make the cleaned data into a dataframe
ZScore_df_LTSPS  = pd.DataFrame(cleaned_data,columns=['Company Name','ZScore LTSPS'])

ZScore_df_LTSPS['ZScore LTSPS'] = ZScore_df_LTSPS['ZScore LTSPS'].clip(lower=-4,upper=4)
# filter out the company to exclude
# filtered_ZScore_df = ZScore_df_LTSPS[~ZScore_df_LTSPS['Company Name'].isin(LTSPS_company_to_be_given_min_ZScore)]


# find the minimum ZScore of rest of the companies
# min_zscore = filtered_ZScore_df['ZScore ROE'].min()
# min_zscore = 4


# update the ZScore Values for the filtered company
# ZScore_df_LTSPS.loc[ZScore_df_LTSPS['Company Name'].isin(LTSPS_company_to_be_given_min_ZScore),'ZScore LTSPS'] = min_zscore



# # export to excel
# timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# new_name = f'LTSPS_{year_end_const}_{timestamp}'

# with pd.ExcelWriter("output.xlsx",engine="openpyxl", mode='a') as writer:
#     ZScore_df_LTSPS.to_excel(writer, index=False, header=True, sheet_name=new_name)


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


print(f'Long Term historical SPS growth ZScore calculation success for {year_end_const}')