import numpy as np
# These are the parameters which can be controlled 

# The year end and month for which you want your code to execute 
# example : year_end_const = 202303
 
# If you want to run for latest year then put np.nan
# year_end_const = np.nan
year_end_const = 201706
# If you want to run for latest year then put np.nan
# year_end_mon = np.nan
year_end_mon = 201706

#Which dataset do you want to use
excel_file_path_quarterly = "All DataSet.xlsx"
excel_file_path_monthly = "AutoAncillaries Monthly.xlsx"


# change these factors if you want to calculate for past x no of years 

# for this parameter if last five years data is not present for EPS and Sales calculation, then make it 3 
how_many_year_back = 3  

# for this parameter if last five years data is present for EPSVar calculation then 17, for three years 8
rolling_quarters = 8








# Quality Factor ZScore Calculation Parameters 
# also include +/- in the parameters 
# Return on Equity Weightage
# Debt to Equity Weightage
# Earnings per Share Variability Weightage
ROE_ratio    =  0.33
DE_ratio     = -0.33
EPSVar_ratio = -0.33


# VALUE FACTOR 

# Value factor ZScore calculation parameters
# also include +/- in the parameters
# Twelve Trailing Months Price to Earnings Weightage
# Twelve Trailing Months Book Value to Price Weightage
# Dividend per Share to Price weightage
TTMPE_ratio  =-0.33
TTMBvP_ratio = 0.33
DPSP_ratio   = 0.33


# Growth Factor ZScore Calculation parameters
# also include +/- in the parameters
# Long Term historical Earnings per Share growth weightage
# Long Term historical Sales per Share weightage
# Current Internal Growth Rate Weightage
LTEPS_ratio     = 0.33
LTSPS_ratio     = 0.33
internalG_ratio = 0.33


# MOMENTUM FACTOR 

# Momentum Factor ZScore calculation parameters
# also include +/- in the parameters
# 6 month price momentum weightage
# 12 month price momentum weightage
momentum_six_ratio     = 0.5
momentum_twelve_ratio  = 0.5



# Overall ZScore calculation parameters

Quality_ratio  = 0.25
Value_ratio    = 0.25
Growth_ratio   = 0.25
Momentum_ratio = 0.25

















# QUALITY FACTOR 

# 1. ROE 

# Which companies should be excluded from the mean and standard deviation calculations 
ROE_exclude_for_mean_stddev = ['']




# 2. Debt to Equity 

# Which companies should be excluded from the mean and standard deviation calculations 
DE_exclude_for_mean_stddev = ['']




# 3. EPS Variability 

# Which Companies should be excluded from the mean and standard deviation calculations 
EPSVar_exclude_for_mean_stddev =['']




# GROWTH FACTOR 


# 1. Long Term Historical EPS Growth

# Which companies should be excluded from the mean and standard deviation calculations 
LTEPS_exclude_for_mean_stddev =['']


# 2. Long Term Historical Sales Per Share Growth

# Which companies should be excluded from the mean and standard deviation calculations 
LTSPS_exclude_for_mean_stddev =['']



# 3. Current Internal Growth Rate

# Which companies should be excluded from the mean and standard deviation calculations 
internalG_exclude_for_mean_stddev =['']












