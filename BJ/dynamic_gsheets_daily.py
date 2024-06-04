import pandas as pd
import pygsheets
from datetime import datetime, timedelta
import argparse
import numpy as np
import os

client = pygsheets.authorize(service_file=r'C:\Users\ardie.asilo\Desktop\Performance_Tracker\BJ\my-test-project-416604-de901d750987.json')

parser = argparse.ArgumentParser()
parser.add_argument("--spreadsheet", help="Name of the spreadsheet")
parser.add_argument("--currency", help="Currency")
args = parser.parse_args()

spreadsheet_name = args.spreadsheet
currency = args.currency

spreadsht = client.open(spreadsheet_name)
worksht = spreadsht.worksheet("title", "Daily")

file_path_1 = "file/1.csv"
df1 = pd.read_csv(file_path_1)

file_path_2 = "file/2.csv"
df2 = pd.read_csv(file_path_2)

file_path_3 = "file/3.csv"
df3 = pd.read_csv(file_path_3)

file_path = "file/4.csv" 
df4 = pd.read_csv(file_path)

file_path = "file/5.csv"
df5 = pd.read_csv(file_path)

file_path = "file/6.csv"
df6 = pd.read_csv(file_path)

file_path = "file/7.csv"
df7 = pd.read_csv(file_path)

####################################################################################
file_path = "file/8.csv"
if os.path.isfile(file_path):
    df8 = pd.read_csv(file_path)
    df8['Date'] = pd.to_datetime(df8['create_date_id'])
    df8 = df8.sort_values(by='Date')
    unique_dates_df8 = set(df8['Date'])
    selected_columns = [
        'Count', 'VIP Cash'
    ]
    selected_df8 = df8[selected_columns]
    selected_df8_transposed = selected_df8.T
else:
    selected_df8_transposed = None 
####################################################################################
file_path = "file/9.csv"
if os.path.isfile(file_path):
    df9 = pd.read_csv(file_path)
    df9['Date'] = pd.to_datetime(df9['received_date_id'])
    df9 = df9.sort_values(by='Date')
    unique_dates_df9 = set(df9['Date'])
    selected_columns = [
        'count', 'RAF Commission'
    ]
    selected_df9 = df9[selected_columns]
    selected_df9_transposed = selected_df9.T
else:
    selected_df9_transposed = None 
####################################################################################
def process_recursive(worksht, num_unique_dates, start_row=2):
    date_str = worksht.cell('E1').value
    date = datetime.strptime(date_str, '%d/%m/%Y')
    worksht.insert_cols(4, number=num_unique_dates)
    for i in range(num_unique_dates):
        date += timedelta(days=1)
        new_date_str = date.strftime('%d/%m/%Y')
        worksht.update_value((1, 5 + num_unique_dates- i - 1), new_date_str)
    
    for i in range(num_unique_dates):
        if spreadsheet_name == "Baji 2024 Biz Performance Tracker - VND":
            #VND
            worksht.set_dataframe(selected_df1_transposed, start='E2', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df2_transposed, start='E7', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df3_transposed, start='E19', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df4_transposed, start='E31', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df5, start='E37', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df6_transposed, start='E183', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df7, start='E187', copy_index=False, copy_head=False)

        elif spreadsheet_name == "Baji 2024 Biz Performance Tracker - PHP" or spreadsheet_name == "Baji 2024 Biz Performance Tracker - USD":
            #PHP ,USD
            worksht.set_dataframe(selected_df1_transposed, start='E2', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df2_transposed, start='E7', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df3_transposed, start='E19', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df4_transposed, start='E31', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df5, start='E37', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df6_transposed, start='E177', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df7, start='E181', copy_index=False, copy_head=False)
            
        else:
            #BAJI: BDT ,INR, PKR
            #Jeetbuzz: BDT, INR, PKR
            #6s: BDT, INR, PKR
            worksht.set_dataframe(selected_df1_transposed, start='E2', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df2_transposed, start='E7', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df3_transposed, start='E19', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df4_transposed, start='E31', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df5, start='E37', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df6_transposed, start='E179', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df7, start='E183', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df8_transposed, start='E209', copy_index=False, copy_head=False)
            worksht.set_dataframe(selected_df9_transposed, start='E213', copy_index=False, copy_head=False)

df1['Date'] = pd.to_datetime(df1['Date'])
df2['Date'] = pd.to_datetime(df2['Date'])
df3['Date'] = pd.to_datetime(df3['Date'])
df4['Date'] = pd.to_datetime(df4['Time Settled'])
df5['Date'] = pd.to_datetime(df5['settle_date_id'])
df6['Date'] = pd.to_datetime(df6['create_date_id'])
df7['Date'] = pd.to_datetime(df7['Date'])

df1 = df1.sort_values(by='Date', ascending=False)
df2 = df2.sort_values(by='Date')
df3 = df3.sort_values(by='Date')
df4 = df4.sort_values(by='Date')
df5 = df5.sort_values(by='Date')
df6 = df6.sort_values(by='Date')
df7 = df7.sort_values(by='Date')

unique_dates_df1 = set(df1['Date'])
unique_dates_df2 = set(df2['Date'])
unique_dates_df3 = set(df3['Date'])
unique_dates_df4 = set(df4['Date'])
unique_dates_df5 = set(df5['Date'])
unique_dates_df6 = set(df6['Date'])
unique_dates_df7 = set(df7['Date'])

missing_dates = unique_dates_df1 - unique_dates_df2
missing_dates_df3 = unique_dates_df1 - unique_dates_df3
missing_dates_df4 = unique_dates_df1 - unique_dates_df4
missing_dates_df6 = unique_dates_df1 - unique_dates_df6

for missing_date in missing_dates:
    row_data = {
        'Date': missing_date,
        'Unique Depositors': 0, 
        'Deposit Count': 0, 
        'Deposit Amount': 0.00,                   
        'Unique Withdrawer': 0.00, 
        'Withdrawal Count': 0.00, 
        'Withdrawal Amount': 0.00, 
        'Cash Inflow/(Outflow)': 0.00, 
        'No. Unique Players W/ Returning Deposits': 0, 
        'Returning Deposits Count': 0
    }
    df2 = df2._append(row_data, ignore_index=True)
df2 = df2.sort_values(by='Date', ascending=False)

for missing_dates_df3 in missing_dates_df3:
    row_data = {
        'Date': missing_dates_df3,
        'Total Turnover': 0.00, 
        'Profit/Loss': 0.00, 
        'Gross Margin': 0.00,
        '(-) Bonus Cost': 0.00, 
        '(-) Adjustment': 0.00, 
        'Net Gross Profit': 0.00
    }
    df3 = df3._append(row_data, ignore_index=True)
df3 = df3.sort_values(by='Date', ascending=False)

for missing_dates_df4 in missing_dates_df4:
    row_data = {
        'Date': missing_dates_df4,
        'Total Unique Active Players': 0, 
        'Total Turnover': 0.00, 
        'Profit/Loss': 0.00, 
        'Turnover Margin (%)': 0.00
    }
    df4 = df4._append(row_data, ignore_index=True)
df4 = df4.sort_values(by='Date', ascending=False)
#------------------------------------------------------------------------------------------------------------------------------5 start
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

row_data = {
    'Number of Unique Player': 0,
    'Total Turnover': 0,
    'Profit/Loss': 0,
    'Margin': 0.0
}

missing_types_df5 = set(product_type_order) - set(df5['Product Type'])
appended_rows = pd.DataFrame(columns=['Date', 'Product Type','Number of Unique Player', 'Total Turnover', 'Profit/Loss', 'Margin'])
for date in unique_dates_df1:
    if date not in unique_dates_df5:
        for product_type in product_type_order:
            new_row = {'Date': date, 'Product Type': product_type}
            new_row.update(row_data)
            appended_rows = appended_rows._append(new_row, ignore_index=True)
    else:
        for product_type in product_type_order:
            if product_type in missing_types_df5 or not ((df5['Date'] == date) & (df5['Product Type'] == product_type)).any():
                new_row = {'Date': date, 'Product Type': product_type}
                new_row.update(row_data)
                appended_rows = appended_rows._append(new_row, ignore_index=True)
output_df5 = pd.concat([df5, appended_rows])
output_df5['Product Type'] = pd.Categorical(output_df5['Product Type'], categories=product_type_order, ordered=True)
output_df5 = output_df5.sort_values(by=['Product Type', 'Date'], ascending=[True, False])
selected_columns = [
    'Number of Unique Player', 
    'Total Turnover', 
    'Profit/Loss', 
    'Margin'
]
output_df5.set_index('Date', inplace=True)
selected_df5 = output_df5.groupby('Product Type').apply(lambda x: x[selected_columns].T).reset_index(level=1, drop=True).reset_index()
selected_df5.drop(columns=['Product Type'], inplace=True)
selected_df5 = selected_df5.groupby(np.arange(len(selected_df5)) // 4).apply(lambda x: pd.concat([x, pd.DataFrame(index=range(2))])).reset_index(drop=True)
selected_df5 = selected_df5.fillna("")
#------------------------------------------------------------------------------------------------------------------------------5 end
for missing_dates_df6 in missing_dates_df6:
    row_data = {
        'Date': missing_dates_df6,
        'Total Bonus Cost': 0.00, 
        'Total Claimed': 0, 
        'Total Unique Player Claimed': 0
    }
    df6 = df6._append(row_data, ignore_index=True)
df6 = df6.sort_values(by='Date', ascending=False)
#-----------------------------------------------------------------------------------------------------------------------------1 start
unique_dates = df1['Date'].dt.date.unique()
num_unique_dates = len(unique_dates)
selected_columns = ['NSU', 'FTD', 'Conversion Rate']
selected_df1 = df1[selected_columns]
selected_df1.insert(loc=2, column='Empty Column', value='')
selected_df1_transposed = selected_df1.T
#-----------------------------------------------------------------------------------------------------------------------------1 end
#-----------------------------------------------------------------------------------------------------------------------------2 start
selected_columns = [
    'Unique Depositor', 'Deposit Count', 'Deposit Amount', 
    'Unique Withdrawer', 'Withdraw Count', 'Withdraw Amount', 
    'Cash Inflow/Outflow', 'No. Unique Players W/ Returning Deposits', 
    'Return Deposit Count'
]
selected_df2 = df2[selected_columns]
selected_df2.insert(loc=6, column='Empty Column 1', value='')
selected_df2.insert(loc=8, column='Empty Column 2', value='')
selected_df2_transposed = selected_df2.T
#-----------------------------------------------------------------------------------------------------------------------------2 end
#-----------------------------------------------------------------------------------------------------------------------------3 start
selected_columns = [
    'Total Turnover', 'Profit/Loss', 'Gross Margin',
    '(-) Bonus Cost', '(-) Adjustment', 'Net Gross Profit'
]
selected_df3 = df3[selected_columns]
selected_df3.insert(loc=3, column='', value='')
selected_df3_transposed = selected_df3.T
#-----------------------------------------------------------------------------------------------------------------------------3 end
#-----------------------------------------------------------------------------------------------------------------------------4 start
selected_columns = [
    'Total Unique Active Players', 'Total Turnover', 
    'Profit/Loss', 'Turnover Margin (%)'
]
selected_df4 = df4[selected_columns]
selected_df4_transposed = selected_df4.T
#-----------------------------------------------------------------------------------------------------------------------------4 end
#-----------------------------------------------------------------------------------------------------------------------------6 start
selected_columns = [
    'Total Bonus Cost', 
    'Total Claimed', 'Total Unique Player Claimed'
]
selected_df6 = df6[selected_columns]
selected_df6_transposed = selected_df6.T
#-----------------------------------------------------------------------------------------------------------------------------6 end
#-----------------------------------------------------------------------------------------------------------------------------7 start
desired_order = ['Acquisition', 'Retention', 'Conversion', 'Reactivation']
row_data1 = {
    'Bonus Cost': 0,
    'Total Claimed': 0,
    'Total Unique Player Claimed': 0,
    'Rank_1': '',
    'Rank_2': '',
    'Rank_3': '',
    'Rank_4': '',
    'Rank_5': '',
    'Rank_6': '',
    'Rank_7': '',
    'Rank_8': '',
    'Rank_9': '',
    'Rank_10': ''
}
missing_types_df7 = set(desired_order) - set(df7['Purpose'])
appended_rows1 = pd.DataFrame(columns=['Date', 'Purpose', 'Bonus Cost', 
                                    'Total Claimed', 'Total Unique Player Claimed', 
                                    'Rank_1', 'Rank_2', 'Rank_3', 'Rank_4', 'Rank_5', 
                                    'Rank_6', 'Rank_7', 'Rank_8', 'Rank_9', 'Rank_10'])
def RankerFunction():
    for index, row in df7.iterrows():
        lowest_rank_value = None
        lowest_rank_col = None
        for i in range(2, 11):
            rank_value = row[f"Rank_{i}"]
            if isinstance(rank_value, str) and len(rank_value) > 2 and 'test' not in rank_value.lower():
                if lowest_rank_value is None or rank_value > lowest_rank_col:
                    lowest_rank_col = f"Rank_{i}"
                lowest_rank_value = rank_value
                df7.at[index, 'Rank_2'] = lowest_rank_value
for date in unique_dates_df1:
    if date not in unique_dates_df7:
        for purpose in desired_order:
            new_row = {'Date': date, 'Purpose': purpose}
            new_row.update(row_data1)
            appended_rows1 = appended_rows1._append(new_row, ignore_index=True)
            RankerFunction()
    else:
        for purpose in desired_order:
            if purpose in missing_types_df7 or not ((df7['Date']==date)&(df7['Purpose']==purpose)).any():
                new_row = {'Date': date, 'Purpose':purpose}
                new_row.update(row_data1)
                appended_rows1 = appended_rows1._append(new_row, ignore_index=True)
                RankerFunction()

output_df7 = pd.concat([df7, appended_rows1])
output_df7['Purpose'] = pd.Categorical(output_df7['Purpose'], categories=desired_order, ordered=True)
output_df7 = output_df7.sort_values(by=['Purpose', 'Date'], ascending=[True, False])
selected_columns1 = [    
    'Bonus Cost', 
    'Total Claimed', 
    'Total Unique Player Claimed',
    'Rank_1',
    'Rank_2'
]
output_df7.set_index('Date', inplace=True)     
selected_df7 = output_df7.groupby('Purpose').apply(lambda x: x[selected_columns1].T).reset_index(level=1, drop=True).reset_index()
selected_df7.drop(columns=['Purpose'], inplace=True)
selected_df7 = selected_df7.groupby(np.arange(len(selected_df7)) // 5).apply(lambda x: pd.concat([x, pd.DataFrame(index=range(1))])).reset_index(drop=True)
selected_df7 = selected_df7.fillna("")
#-----------------------------------------------------------------------------------------------------------------------------7 end
process_recursive(worksht, num_unique_dates)