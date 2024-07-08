import pandas as pd
import statistics
from datetime import datetime 



# import the parameters needed from parameters.py
from parameters import year_end_const
from parameters import EPSVar_exclude_for_mean_stddev
from parameters import rolling_quarters
from parameters import excel_file_path_quarterly


# year_end_const = 202403
# EPSVar_exclude_for_mean_stddev=['Varroc Engineer']
# EPSVar_company_to_be_given_min_ZScore = ['Varroc Engineer','Steel Str. Wheel']

# read the dataset file 
# excel_file_path = "All DataSet.xlsx"
xls = pd.ExcelFile(excel_file_path_quarterly)
sheet_names = xls.sheet_names

# year end values which can exist in this space ie yearmonth 202303 type
year_end_values = [] 

for year in range(2030,2012,-1):
    for month in [12,9,6,3]:
        year_end_values.append(int(f"{year}{month:02}"))


# convert all the sheets into dataframes
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


# calculate the TTM EPS
for sheet_name in sheet_names:
    df = globals()[sheet_name]
    df['TTM EPS'] = df['Diluted EPS Adj.'].rolling(window = 4).sum()


# function for calculating the year over year growth
def calculate_yoy_growth(df):
    yoy_growth = [None]*len(df)
    for i in range(len(df)-1,-1,-1):
        current_year = df.loc[i,'Year End']
        previous_year = current_year-100 
        if previous_year in df['Year End'].values:
            previous_ttm_eps= df.loc[df['Year End'] == previous_year,'TTM EPS'].values[0]
            current_ttm_eps = df.loc[df['Year End'] == current_year, 'TTM EPS'].values[0]
            if previous_ttm_eps<0 : 
               growth = -((current_ttm_eps/previous_ttm_eps)-1)
            else:
               growth = ((current_ttm_eps/previous_ttm_eps)-1)
            yoy_growth[i] = growth
            
    return yoy_growth


# calculte the yoy growth and Earnings Varibility
for sheet_name in sheet_names:
    df = globals()[sheet_name]
    df['YOY'] = calculate_yoy_growth(df)
    df['Earnings Variability'] = df['YOY'].rolling(window = rolling_quarters).std()



# extract the earnings variability values and company name for each company for a particular year end
EarningsVariability_values = []
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
        EarningsVariability_value = row['Earnings Variability']
        EarningsVariability_values.append((company_name,EarningsVariability_value))



# extract only the earnings variability values for mean and standard dev calculations
EarningsVariability_only = [value[1] for value in EarningsVariability_values if value[0] not in EPSVar_exclude_for_mean_stddev]


# calculate the standard deviation
std_dev = statistics.stdev(EarningsVariability_only)

# calculate the mean
mean = statistics.mean(EarningsVariability_only)



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
    Zscore = (df[df['Year End']==year_end]['Earnings Variability'] - mean)/std_dev
    company_name = df[df['Year End'] == year_end]['Company Name']
    ZScore_Values.append((company_name,Zscore))


# clean the ZScore_Values array
cleaned_data = [(item[0].values[0],item[1].values[0]) for item in ZScore_Values]

# make the cleaned data into a dataframe
ZScore_df_EPSVar  = pd.DataFrame(cleaned_data,columns=['Company Name','ZScore EPSVar'])

ZScore_df_EPSVar['ZScore EPSVar'] = ZScore_df_EPSVar['ZScore EPSVar'].clip(lower=-4, upper=4)

# filter out the company to exclude
# filtered_ZScore_df = ZScore_df_EPSVar[~ZScore_df_EPSVar['Company Name'].isin(EPSVar_company_to_be_given_min_ZScore)]

# find the minimum ZScore of rest of the companies
# min_zscore = filtered_ZScore_df['ZScore ROE'].min()
# min_zscore = 4

# update the ZScore Values for the filtered company
# ZScore_df_EPSVar.loc[ZScore_df_EPSVar['Company Name'].isin(EPSVar_company_to_be_given_min_ZScore),'ZScore EPSVar'] = min_zscore


# # export to excel
# timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# new_name = f'EPSVar_{year_end_const}_{timestamp}'

# with pd.ExcelWriter("output.xlsx",engine="openpyxl", mode='a') as writer:
#     ZScore_df_EPSVar.to_excel(writer, index=False, header=True, sheet_name=new_name)


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


print(f'EPS Variability Z Score Calculation Success for {year_end_const}')