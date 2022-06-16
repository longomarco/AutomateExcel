import pandas as pd
pd.set_option('display.float_format', lambda x: '%.2f' % x)
import numpy as np
import glob

files = glob.glob('SalesData\*.csv')
df_list=[]
for f in files:
    csv = pd.read_csv(f, sep=';', decimal=',')
    df_list.append(csv)

sales = pd.concat(df_list)

#OrderDate converted in datetime
sales['OrderDate'] = pd.to_datetime(sales['OrderDate'])
#New feauture: profit
sales['Profit'] = sales['SalesAmount'] - (sales['TaxAmt'] + sales['Freight'])
#From OrderDate we extract the weekday, then converted from number to the actual day name
sales['WeekDay'] = sales['OrderDate'].dt.weekday
sales['WeekDay'] = sales['WeekDay'].map({
                    0:'Monday',
                    1:'Tuesday',
                    2:'Wednesday',
                    3:'Thursday',
                    4:'Friday',
                    5:'Saturday',
                    6:'Sunday'})
#From OrderDate we extract the month, then converted from number to the actual month name
sales['Month'] = sales['OrderDate'].dt.month
sales['Month'] = sales['Month'].map({
                    1:'Jan',
                    2:'Feb',
                    3:'Mar',
                    4:'Apr',
                    5:'May',
                    6:'Jun',
                    7:'Jul',
                    8:'Aug',
                    9:'Sep',
                    10:'Oct',
                    11:'Nov',
                    12:'Dec'})
#From OrderDate we extract the day
sales['MonthDay'] = sales['OrderDate'].dt.day
#From OrderDate we extract the quarter
sales['Quarter'] = sales['OrderDate'].dt.quarter
sales['Quarter'] = sales['Quarter'].map({
                    1:'Q1',
                    2:'Q2',
                    3:'Q3',
                    4:'Q4'})
#Use OrderDate as index for more clarity
sales.set_index(['OrderDate'], inplace=True)

#TotalSales per Product
total_sales = pd.pivot_table(sales, index='ProductName', values='SalesAmount', aggfunc=np.sum)
total_sales = total_sales.sort_values(by='SalesAmount', ascending=False)
#TotalSales per BuyerName
buyer = pd.pivot_table(sales, index='Name', values=['OrderQuantity', 'SalesAmount'], aggfunc=np.sum)
buyer = buyer.sort_values(by='SalesAmount', ascending=False)
#SalesbyCountry
country_sales = pd.pivot_table(sales, index=['ProductName', 'Country'], values=['OrderQuantity', 'SalesAmount'], aggfunc=np.sum)
country_sales = country_sales.sort_values(by=['ProductName','OrderQuantity'], ascending=False)

#TotalProfit

total_profit = pd.pivot_table(sales, index='ProductName', values='Profit', aggfunc=np.sum)

total_profit = total_profit.sort_values(by='Profit', ascending=False)

#TimeAnalysis
cats = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
sales['WeekDay'] = pd.Categorical(sales['WeekDay'], categories=cats, ordered=True)
pivot_wd = pd.pivot_table(sales, index='WeekDay', values=['OrderQuantity','SalesAmount','Profit'], aggfunc=np.sum)
cats = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
sales['Month'] = pd.Categorical(sales['Month'], categories=cats, ordered=True)
pivot_month = pd.pivot_table(sales, index='Month', values=['OrderQuantity','SalesAmount','Profit'], aggfunc=np.sum)
pivot_quarter = pd.pivot_table(sales, index='Quarter', values=['OrderQuantity','SalesAmount','Profit'], aggfunc=np.sum)

with pd.ExcelWriter('SalesReport_2018.xlsx') as writer:
    sales.to_excel(writer, sheet_name="SalesData", index=True)
    total_sales.to_excel(writer, sheet_name="TotalSales", index=True)
    total_profit.to_excel(writer, sheet_name='TotalProfit', index=True)
    buyer.to_excel(writer, sheet_name="CustomerName", index=True)
    country_sales.to_excel(writer, sheet_name='Country', index=True)
    pivot_wd.to_excel(writer, sheet_name='WeekDay', index=True)
    pivot_month.to_excel(writer, sheet_name='Month', index=True)
    pivot_quarter.to_excel(writer, sheet_name='Quarter', index=True)

print("Report successfully created")