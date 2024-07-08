import pandas as pd
import warnings 
import statistics
from datetime import datetime

# import the parameters from parameters.py
from parameters import year_end_const
from parameters import DE_exclude_for_mean_stddev
from parameters import excel_file_path_quarterly


warnings.filterwarnings('ignore')

# ///////////////////////////////////////////////////////////////

# hyperparameters 

# year end where you want to calculate d/e ; here it can only be half yearly
# year_end_const = 202403

# outlier company to be excluded in mean and standard devaition
# DE_exclude_for_mean_stddev = ['Automotive Stamp']

# company which lies far beyond std dev to be given the max/min ZScore
# DE_company_to_be_given_max_ZScore = ['Ashok Leyland ','Automotive Stamp']


# ///////////////////////////////////////////////////////////////

# read the dataset of all companies 
# excel_file_path = "All DataSet.xlsx"
xls = pd.ExcelFile(excel_file_path_quarterly)
sheet_names = xls.sheet_names



# year end values which can exist in this space ie yearmonth 202303 type
year_end_values = [] 

for year in range(2030,2012,-1):
    for month in [12,9,6,3]:
        year_end_values.append(int(f"{year}{month:02}"))



# convert all the sheet into pandas dataframe by their individual name
for sheet_name in sheet_names:
    globals()[sheet_name] = pd.read_excel(excel_file_path_quarterly,sheet_name=sheet_name)

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

# perform calculations on the individual dataframes for all companies
for sheet_name in sheet_names:
    df = globals()[sheet_name]
    df['Share Capital']=df['Share Capital'].fillna(method = 'ffill')
    df['Loan Funds'] = df['Loan Funds'].fillna(method = 'ffill')
    df['Cash & Bank Balance'] = df['Cash & Bank Balance'].fillna(method = 'ffill')
    target_column = 'Reserves & Surplus'
    other_column = 'Adjusted Profit After Extra-ordinary item'

    for i in range(1,len(df)):
        if pd.isna(df.loc[i,target_column]):
            df.loc[i,target_column] = df.loc[i-1,target_column]+df.loc[i,other_column]
    
    df['Networth'] =df['Share Capital']+df['Reserves & Surplus']
    df['Net Debt'] = df['Loan Funds'] - df['Cash & Bank Balance']
    df['D/E'] = df['Net Debt']/df['Networth']
    


# extract the debt equity value with company name for particular year end and all the companies into an array de_values
de_values=[]

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
        de_value = row['D/E']
        de_values.append((company_name,de_value))


# extract the d/e value only for the mean and std dev calculations
de_only= [de for company, de in de_values if company not in DE_exclude_for_mean_stddev]


#calculate standard deviation
std_dev_de = statistics.stdev(de_only)

# calculate mean
mean_de = statistics.mean(de_only)


# calculate the ZScore
ZScore_Values_de=[]

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

    Zscore = (df[df['Year End']==year_end]['D/E'] - mean_de)/std_dev_de
    company_name = df[df['Year End'] == year_end]['Company Name']
    ZScore_Values_de.append((company_name,Zscore))




#clean the ZScore_Values_de as it also contained data type and names
cleaned_data = [(item[0].values[0],item[1].values[0]) for item in ZScore_Values_de]


#create a new dataframe with the cleaned datas
ZScore_df_DE  = pd.DataFrame(cleaned_data,columns=['Company Name','ZScore D/E'])

ZScore_df_DE['ZScore D/E'] = ZScore_df_DE['ZScore D/E'].clip(lower=-4,upper=4)


# filter the company for which ZScore has to be excluded
# filtered_ZScore_df = ZScore_df_DE[~ZScore_df_DE['Company Name'].isin(DE_company_to_be_given_max_ZScore)]

#select the max ZScore 
# max_zscore = filtered_ZScore_df['ZScore D/E'].max()
# max_zscore = 4

#give the filtered company max ZScore 
# ZScore_df_DE.loc[ZScore_df_DE['Company Name'].isin(DE_company_to_be_given_max_ZScore),'ZScore D/E'] = max_zscore



# #export to excel 
# timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# new_name = f'DE_{year_end_const}_{timestamp}'

# with pd.ExcelWriter("output.xlsx",engine="openpyxl", mode='a') as writer:
#     ZScore_df_DE.to_excel(writer, index=False, header=True, sheet_name=new_name)


# # adjusted width of cells while exporting to excel
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
    
print(f'D/E Z Score Calculation Success for {year_end_const} ')