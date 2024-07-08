import pandas as pd
from datetime import datetime

# import individual factors 
from Quality_merged import Quality_merged_df
from Value_merged import Value_merged_df
from Growth_merged import Growth_merged_df
from momentum_merged import Momentum_merged_df

# import parameters from paramters.py
from parameters import Quality_ratio
from parameters import Value_ratio
from parameters import Growth_ratio
from parameters import Momentum_ratio


if 'Quality_merged_df' not in globals():
    raise NameError('Quality_merged_df not defined in Quality_merged')


if 'Growth_merged_df' not in globals():
    raise NameError('Growth_merged_df not defined in Growth_merged_df')


merged_1 = pd.merge(Quality_merged_df,Value_merged_df,on='Company Name')
merged_2 = pd.merge(merged_1,Growth_merged_df,on='Company Name')
merged_3 = pd.merge(merged_2,Momentum_merged_df,on='Company Name')


merged_3['Overall ZScore'] = Quality_ratio*merged_3['Quality ZScore'] + Value_ratio*merged_3['Value ZScore'] + Growth_ratio*merged_3['Growth ZScore'] + Momentum_ratio*merged_3['Momentum ZScore']



#export to excel 
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
new_name = f'MultiFactor_{timestamp}'

with pd.ExcelWriter("main.xlsx",engine="openpyxl", mode='a') as writer:
    merged_3.to_excel(writer, index=False, header=True, sheet_name=new_name)


# adjusted width of cells while exporting to excel
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



print(f'Works')