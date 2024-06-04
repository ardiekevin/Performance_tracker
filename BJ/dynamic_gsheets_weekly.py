from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import pygsheets
import pandas as pd
import argparse
import os

client = pygsheets.authorize(service_file=r'C:\Users\ardie.asilo\Desktop\Performance_Tracker\BJ\my-test-project-416604-de901d750987.json')
parser = argparse.ArgumentParser()
parser.add_argument("--spreadsheet", help="Name of the spreadsheet")
parser.add_argument("--currency", help="Currency")
args = parser.parse_args()
spreadsheet_name = args.spreadsheet
currency = args.currency
spreadsht = client.open(spreadsheet_name)
worksht = spreadsht.worksheet("title", "Weekly") 

file_path_1 = "file/1.csv"
df1 = pd.read_csv(file_path_1)
selected_columns = ['NSU', 'FTD', 'Conversion Rate']
selected_df1 = df1[selected_columns]
selected_df1.insert(loc=2, column='Empty Column', value='')
selected_df1_transposed = selected_df1.T
#########################################################################################
file_path_2 = "file/2.csv"
df2 = pd.read_csv(file_path_2)
df2['Date'] = pd.to_datetime(df2['Date'])
df2 = df2.sort_values(by='Date', ascending=False)
selected_columns = [
    'Unique Depositors', 'Deposit Count', 'Deposit Amount', 
    'Unique Withdrawer', 'Withdrawal Count', 'Withdrawal Amount'
]
selected_df2 = df2[selected_columns]
selected_df2.insert(loc=6, column='Empty Column 1', value='')
selected_df2_transposed = selected_df2.T
###########################################################################################
file_path_3 = "file/3.csv"
df3 = pd.read_csv(file_path_3)
df3['Date'] = pd.to_datetime(df3['Date'])
df3 = df3.sort_values(by='Date', ascending=False)
selected_columns = [
    'Total Turnover', 'Profit/Loss', 'Gross Margin (%)',
    '(-) Bonus Cost', '(-) Adjustment', 'Net Gross Profit', 
    'Average Daily Turnover', 'Average Daily Active Players', 
    'Average Daily Profit/Loss'
]
selected_df3 = df3[selected_columns]
selected_df3.insert(loc=3, column='Empty Column 1', value='')
selected_df3.insert(loc=7, column='Empty Column 2', value='')
selected_df3_transposed = selected_df3.T
##############################################################################################
file_path = "file/4.csv" 
df4 = pd.read_csv(file_path)
df4['Date'] = pd.to_datetime(df4['Time Settled'])
df4 = df4.sort_values(by='Date', ascending=False)
selected_columns = [
    'Total Unique Active Players', 'Total Turnover', 
    'Profit/Loss', 'Turnover Margin (%)'
]
selected_df4 = df4[selected_columns]
selected_df4_transposed = selected_df4.T
###############################################################################################
file_path = "file/5.csv"
df5 = pd.read_csv(file_path)
if spreadsheet_name == "Baji 2024 Biz Performance Tracker - PHP":
    product_type_order = ['Sport', 'SLOT', 'CASINO', 'TABLE', 'COCK_FIGHTING',  
                        'FH', 'LOTTERY', 'ARCADE', 'CRASH', 'ESport', 
                        'CARD', 'NoValue']
    
elif spreadsheet_name == "Baji 2024 Biz Performance Tracker - VND":
    product_type_order = ['Sport', 'SLOT', 'CASINO', 'TABLE', 'P2P',  
                        'FH', 'LOTTERY', 'ARCADE', 'ESport', 
                        'CARD', 'NoValue', 'CRASH', 'COCK_FIGHTING']
    
else:
    product_type_order = ['Sport', 'SLOT', 'CASINO', 'TABLE', 'P2P',  
                        'FH', 'LOTTERY', 'ARCADE', 'ESport', 
                        'CARD', 'NoValue', 'CRASH']

missing_types = set(product_type_order) - set(df5['Product Type'])
for missing_type in missing_types:
    df5 = df5._append({'Product Type': missing_type,
                      'Number of Unique Player': 0,
                      'Total Turnover': 0,
                      'Profit/Loss': 0,
                      'Margin': 0.0}, ignore_index=True)
df5['Product Type'] = pd.Categorical(df5['Product Type'], categories=product_type_order, ordered=True)
df5 = df5.sort_values(by=['Product Type', 'settle_date_id'], ascending=[True, False])
selected_columns = [
    'Number of Unique Player', 
    'Total Turnover', 'Profit/Loss', 'Margin'
]
selected_df5 = df5[selected_columns]
selected_df5_transposed = selected_df5.T
output_df = pd.DataFrame(columns=['Index', 'Value'])
for column in selected_df5_transposed.columns:
    for index, value in selected_df5_transposed[column].items():
        output_df = output_df._append({'Index': index, 'Value': value}, ignore_index=True)
    output_df = output_df._append({'Index': '', 'Value': ''}, ignore_index=True)
    output_df = output_df._append({'Index': '', 'Value': ''}, ignore_index=True)
###############################################################################################
file_path = "file/6.csv"
df6 = pd.read_csv(file_path)
df6['create_date_id'] = pd.to_datetime(df6['create_date_id'])
df6 = df6.sort_values(by='create_date_id', ascending=False)
selected_columns = [
    'Total Bonus Cost', 'Total Claimed', 
    'Total Unique Player Claimed'
]
selected_df6 = df6[selected_columns]
selected_df6_transposed = selected_df6.T
###############################################################################################
file_path = "file/7.csv"
df7 = pd.read_csv(file_path)
desired_order = ['Acquisition', 'Retention', 'Conversion', 'Reactivation']
missing_titles = set(desired_order) - set(df7['Purpose'])
for missing_titles in missing_titles:
    df7 = df7._append({'Purpose': missing_titles,
                      'Bonus Cost': 0,
                      'Total Claimed': 0,
                      'Total Unique Player Claimed': 0,
                      'Rank_1': '',
                      'Rank_2': ''}, ignore_index=True)
df7['Purpose'] = pd.Categorical(df7['Purpose'], categories=desired_order, ordered=True)
df7 = df7.sort_values(by=['Purpose'], ascending=[True])
output_df7 = pd.DataFrame(columns=['Index', 'Value'])
df7['Date'] = pd.to_datetime(df7['Date'])
df7 = df7.sort_values(by=['Purpose', 'Date'], ascending=[True, False])
selected_columns = [    
    'Bonus Cost', 'Total Claimed', 'Total Unique Player Claimed',
    'Rank_1', 'Rank_2'
]
selected_df7 = df7[selected_columns]
for index, row in df7.iterrows():
    lowest_rank_value = None
    lowest_rank_col = None
    for i in range(2, 11):
        rank_value = row[f"Rank_{i}"]
        if isinstance(rank_value, str) and len(rank_value) > 2 and not rank_value.lower().startswith('test'):
            if lowest_rank_value is None or rank_value > lowest_rank_col:
                lowest_rank_col = f"Rank_{i}"
            lowest_rank_value = rank_value
            selected_df7.at[index, 'Rank_2'] = lowest_rank_value
transposed_df7 = selected_df7.T
for column in transposed_df7.columns:
    for index, value in transposed_df7[column].items():
        output_df7 = output_df7._append({'Index': index, 'Value': value}, ignore_index=True)
    output_df7 = output_df7._append({'Index': '', 'Value': ''}, ignore_index=True)
###############################################################################################
file_path = "file/8.csv"
if os.path.isfile(file_path):
    df8 = pd.read_csv(file_path)
    df8['create_date_id'] = pd.to_datetime(df8['create_date_id'])
    df8 = df8.sort_values(by='create_date_id', ascending=False)
    selected_columns_8 = ['Count', 'VIP Cash']
    selected_df8 = df8[selected_columns_8]
    selected_df8_transposed = selected_df8.T
else:
    selected_df8_transposed = None
###############################################################################################
file_path = "file/9.csv"
if os.path.isfile(file_path):
    df9 = pd.read_csv(file_path)
    df9['received_date_id'] = pd.to_datetime(df9['received_date_id'])
    df9 = df9.sort_values(by='received_date_id', ascending=False)
    selected_columns_9 = ['count', 'RAF Commission']
    selected_df9 = df9[selected_columns_9]
    selected_df9_transposed = selected_df9.T
else:
    selected_df9_transposed = None 
###############################################################################################
date_str = worksht.cell('E1').value
date_format = '%d/%m/%Y'
date = datetime.strptime(date_str, date_format)
worksht.insert_cols(4, number=1)
date += timedelta(days=7)
new_date_str = date.strftime(date_format)
worksht.update_value('E1', new_date_str)
if spreadsheet_name == "Baji 2024 Biz Performance Tracker - VND":
    #VND 
    worksht.set_dataframe(selected_df1_transposed, start='E2', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df2_transposed, start='E7', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df3_transposed, start='E19', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df4_transposed, start='E33', copy_index=False, copy_head=False)
    worksht.set_dataframe(output_df.iloc[:,[1]], start='E39', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df6_transposed, start='E185', copy_index=False, copy_head=False)
    worksht.set_dataframe(output_df7.iloc[:,[1]], start='E189', copy_index=False, copy_head=False)

elif (spreadsheet_name == "Baji 2024 Biz Performance Tracker - PHP" or spreadsheet_name == "Baji 2024 Biz Performance Tracker - USD" or 
    spreadsheet_name == "Baji 2024 Biz Performance Tracker - PKR" or spreadsheet_name == "Baji 2024 Biz Performance Tracker - INR" or 
    spreadsheet_name == "Baji 2024 Biz Performance Tracker - BDT"):
    # PHP, USD, PKR, INR, BDT
    worksht.set_dataframe(selected_df1_transposed, start='E2', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df2_transposed, start='E7', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df3_transposed, start='E19', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df4_transposed, start='E33', copy_index=False, copy_head=False)
    worksht.set_dataframe(output_df.iloc[:,[1]], start='E39', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df6_transposed, start='E179', copy_index=False, copy_head=False)
    worksht.set_dataframe(output_df7.iloc[:,[1]], start='E183', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df8_transposed, start='E209', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df9_transposed, start='E213', copy_index=False, copy_head=False)
else:
    #JB
    #INR, PKR, BDT
    #6s
    #BDT, INR, PKR
    worksht.set_dataframe(selected_df1_transposed, start='E2', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df2_transposed, start='E7', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df3_transposed, start='E19', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df4_transposed, start='E31', copy_index=False, copy_head=False)
    worksht.set_dataframe(output_df.iloc[:,[1]], start='E37', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df6_transposed, start='E178', copy_index=False, copy_head=False)
    worksht.set_dataframe(output_df7.iloc[:,[1]], start='E181', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df8_transposed, start='E209', copy_index=False, copy_head=False)
    worksht.set_dataframe(selected_df9_transposed, start='E213', copy_index=False, copy_head=False)

