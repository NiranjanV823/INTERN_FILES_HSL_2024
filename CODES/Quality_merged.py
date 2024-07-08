import pandas as pd
from datetime import datetime



# import the ZScore Values from ROE DE and EPSVar
from ROE_final import ZScore_df_ROE
from DE_final import ZScore_df_DE
from EPSVar_final import ZScore_df_EPSVar


# import the parameters needed from the parameters.py
from parameters import ROE_ratio, DE_ratio, EPSVar_ratio, year_end_const

# CAUTION !!!!
if 'ZScore_df_ROE' not in globals():
    raise NameError("ZScore_df_ROE is not defined in ROE_final.py")

if 'ZScore_df_DE' not in globals():
    raise NameError("ZScore_df_DE is not defined in DE_final.py")

if 'ZScore_df_EPSVar' not in globals():
    raise NameError("ZScore_df_EPSVar is not defined in EPSVar_final")



# merge the roe and de file
Quality_merged_initial_df = pd.merge(ZScore_df_ROE,ZScore_df_DE,on='Company Name')

# merge the last file with epsvar file
Quality_merged_df = pd.merge(Quality_merged_initial_df,ZScore_df_EPSVar,on = 'Company Name')



# Calculate the Quality factor Z Score 
Quality_merged_df['Quality ZScore'] = ROE_ratio*Quality_merged_df['ZScore ROE'] + DE_ratio*Quality_merged_df['ZScore D/E'] + EPSVar_ratio*Quality_merged_df['ZScore EPSVar']



# #export to excel 
# timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# new_name = f'Quality_{timestamp}'

# with pd.ExcelWriter("Quality_factor.xlsx",engine="openpyxl", mode='a') as writer:
#     Quality_merged_df.to_excel(writer, index=False, header=True, sheet_name=new_name)


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


print(f'Quality Factor Calculation Success')