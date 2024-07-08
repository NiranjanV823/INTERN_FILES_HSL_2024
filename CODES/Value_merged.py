import pandas as pd
from datetime import datetime


from TTMPE import ZScore_df_TTMPE
from TTMBvP import ZScore_df_TTMBvP
from dpsp import ZScore_df_DPSP

from parameters import year_end_const
from parameters import TTMPE_ratio
from parameters import TTMBvP_ratio
from parameters import DPSP_ratio



if 'ZScore_df_TTMPE' not in globals():
    raise NameError("ZScore_df_TTMPE not defined in TTMPE.py")

if 'ZScore_df_TTMBvP' not in globals():
    raise NameError("ZScore_df_TTMBvP not defined in TTMBvP.py")

if 'ZScore_df_DPSP' not in globals():
    raise NameError("ZScore_df_DPSP not defined in dpsp.py")


Value_merged_1 = pd.merge(ZScore_df_TTMPE,ZScore_df_TTMBvP,on='Company Name')

Value_merged_df = pd.merge(Value_merged_1,ZScore_df_DPSP,on='Company Name')


Value_merged_df['Value ZScore'] = TTMPE_ratio* Value_merged_df['ZScore TTMPE'] + TTMBvP_ratio*Value_merged_df['ZScore Bv/P'] + DPSP_ratio*Value_merged_df['ZScore DPS/Price']


# #export to excel 
# timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# new_name = f'Value_{timestamp}'

# with pd.ExcelWriter("Value_factor.xlsx",engine="openpyxl", mode='a') as writer:
#     Value_merged_df.to_excel(writer, index=False, header=True, sheet_name=new_name)


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


print(f'Value Factor Calculation Success')
