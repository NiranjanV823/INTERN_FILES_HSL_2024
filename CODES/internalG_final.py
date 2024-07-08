import pandas as pd
import statistics
import numpy as np
import warnings 
from datetime import datetime

warnings.filterwarnings('ignore')

# import the parameters needed from the parameter.py
from parameters import year_end_const
from parameters import internalG_exclude_for_mean_stddev
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



for sheet_name in sheet_names:
    df = globals()[sheet_name]
    df['Share Capital']=df['Share Capital'].fillna(method = 'ffill')
    target_column = 'Reserves & Surplus'
    other_column = 'Adjusted Profit After Extra-ordinary item'

    for i in range(1,len(df)):
        if pd.isna(df.loc[i,target_column]):
            df.loc[i,target_column] = df.loc[i-1,target_column]+df.loc[i,other_column]
    

    df['Networth'] =df['Share Capital']+df['Reserves & Surplus']
    df['TTM PAT'] = df['Adjusted Profit After Extra-ordinary item'].rolling(window=4).sum()
    df['ROE'] = df['TTM PAT']/df['Networth']


for sheet_name in sheet_names:
    df = globals()[sheet_name]
    df['TTM EPS'] = df['Diluted EPS Adj.'].rolling(window =4).sum()


for sheet_name in sheet_names:
    df = globals()[sheet_name]
    df['Dividend Per Share'] = df['Dividend Per Share'].fillna(method ='ffill')
    df['Payout Ratio'] = df['Dividend Per Share']/df['TTM EPS']
    df['Growth'] = df['ROE']*(1-df['Payout Ratio'])


growth_values = []

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
        growth_value = row['Growth']
        growth_values.append((company_name,growth_value))
        

# extract only the egro values for mean and standard deviation calculations
growth_only = [value[1] for value in growth_values if value[0] not in internalG_exclude_for_mean_stddev]

# calculate standard deviation for egro 
std_dev = statistics.stdev(growth_only)

# calculate mean for egro
mean = statistics.mean(growth_only)


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
    Zscore = (df[df['Year End']==year_end]['Growth'] - mean)/std_dev
    company_name = df[df['Year End'] == year_end]['Company Name']
    ZScore_Values.append((company_name,Zscore))


# clean the ZScore_Values array
cleaned_data = [(item[0].values[0],item[1].values[0]) for item in ZScore_Values]

# make the cleaned data into a dataframe
ZScore_df_internalG  = pd.DataFrame(cleaned_data,columns=['Company Name','ZScore internalG'])
ZScore_df_internalG['ZScore internalG'] = ZScore_df_internalG['ZScore internalG'].clip(lower=-4,upper=4)

# filter out the company to exclude
# filtered_ZScore_df = ZScore_df_internalG[~ZScore_df_internalG['Company Name'].isin(internalG_company_to_be_given_min_ZScore)]


# find the minimum ZScore of rest of the companies
# min_zscore = filtered_ZScore_df['ZScore LTEPS'].min()
# min_zscore = 4


# update the ZScore Values for the filtered company
# ZScore_df_internalG.loc[ZScore_df_internalG['Company Name'].isin(internalG_company_to_be_given_min_ZScore),'ZScore internalG'] = min_zscore




# export to excel
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
new_name = f'internalG_{year_end_const}_{timestamp}'

with pd.ExcelWriter("output.xlsx",engine="openpyxl", mode='a') as writer:
    ZScore_df_internalG.to_excel(writer, index=False, header=True, sheet_name=new_name)


# adjust the cell width when exporting 
    workbook = writer.book
    worksheet = writer.sheets[new_name]

    for column in worksheet.columns:
        max_length =0 
        column_letter = column[0].column_letter

        for cell in column:
            try:
                if len(str(cell.value))> max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = max_length + 2
        worksheet.column_dimensions[column_letter].width = adjusted_width


print(f'Current Internal Growth Rate ZScore calculation success for {year_end_const}')