import pandas as pd
from datetime import datetime


# import the ZScore values from LTEPS LTSPS and Internal Growth rate
from LTEPS_final import ZScore_df_LTEPS
from LTSPS_final import ZScore_df_LTSPS
from internalG_final import ZScore_df_internalG 

# import parameters from parameters.py
from parameters import LTEPS_ratio
from parameters import LTSPS_ratio
from parameters import internalG_ratio
from parameters import year_end_const


if 'ZScore_df_LTEPS' not in globals():
    raise NameError('ZScore_df_LTEPS not defined in LTEPS_final.py')

if 'ZScore_df_LTSPS' not in globals():
    raise NameError('ZScore_df_LTSPS not defined in LTSPS_final.py')


if 'ZScore_df_internalG' not in globals():
    raise NameError('ZScore_df_internalG not defined in internalG_final.py')


# merge LTEPS and LTSPS file 
Growth_merged_1 = pd.merge(ZScore_df_LTEPS,ZScore_df_LTSPS,on='Company Name')

# merge previous file with internalg
Growth_merged_df = pd.merge(Growth_merged_1,ZScore_df_internalG,on='Company Name')


# Calculate the Quality factor Z Score 
Growth_merged_df['Growth ZScore'] = LTEPS_ratio*Growth_merged_df['ZScore LTEPS'] + LTSPS_ratio*Growth_merged_df['ZScore LTSPS'] + internalG_ratio*Growth_merged_df['ZScore internalG']



# #export to excel 
# timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# new_name = f'Growth_{timestamp}'

# with pd.ExcelWriter("Growth_factor.xlsx",engine="openpyxl", mode='a') as writer:
#     Growth_merged_df.to_excel(writer, index=False, header=True, sheet_name=new_name)


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


print(f'Growth Factor Calculation Success')