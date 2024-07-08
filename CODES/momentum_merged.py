import pandas as pd
from datetime import datetime

from momentum_six import ZScore_df_6mom
from momentum_twelve import ZScore_df_12mom
from parameters import year_end_mon

from parameters import momentum_six_ratio
from parameters import momentum_twelve_ratio


if 'ZScore_df_6mom' not in globals():
    raise NameError('ZScore_df_6mon not defined in momentum_six')

if 'ZScore_df_12mom' not in globals():
    raise NameError('ZScore_df_12mon not defined in momentum_twelve')


Momentum_merged_df = pd.merge(ZScore_df_6mom,ZScore_df_12mom,on='Company Name')


Momentum_merged_df['Momentum ZScore'] = momentum_six_ratio*Momentum_merged_df['ZScore 6 month price momentum'] + momentum_twelve_ratio*Momentum_merged_df['ZScore 12 month price momentum']

# #export to excel 
# timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# new_name = f'Momentum_{year_end_mon}_{timestamp}'

# with pd.ExcelWriter("Momentum_factor.xlsx",engine="openpyxl", mode='a') as writer:
#     Momentum_merged_df.to_excel(writer, index=False, header=True, sheet_name=new_name)


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


print(f'Momentum Factor Calculation Success ')