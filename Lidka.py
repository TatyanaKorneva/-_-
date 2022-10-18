import pandas as pd
import openpyxl
from openpyxl import load_workbook
from datetime import datetime, timedelta

df = pd.read_excel('C:/Users/79042/OneDrive/Рабочий стол/Учеба/Работа с графиками/График отпусков_2023.xlsx', sheet_name='График', header=1)

# print whole sheet data
# print(excel_data_df['ФИО'])

count=0
s=[]
spisok_1=[]
spisok_2=[]
spisok_vs=[]
slovar = dict()
spisok_3=[]
for j in range(len(df)):
    # print(df.iloc[j]['ФИО'])
    for columnIndex, value in df.iloc[j].items():
        # print(columnIndex, value)
        if value==1 and columnIndex!='№':
            count += 1
            # print(columnIndex, value)
        else:
            if count!=0:
                s.append(count)
                h1=columnIndex - timedelta(days=count)
                h2=columnIndex - timedelta(days=1)
                spisok_1.append(h1.date())
                spisok_2.append(h2.date())
                count=0
                spisok_vs.append(df.iloc[j]['ФИО'])

    spisok_3 = spisok_1.copy()
    spisok_4 = spisok_2.copy()
    intermediate_dictionary = {'ФИО':spisok_vs, 'Дата начала':spisok_1, 'Дата окончания':spisok_2}
    pandas_dataframe = pd.DataFrame(intermediate_dictionary)
    # df.to_excel(path, sheet_name="sheet1")
    # df.sample(10).to_excel('C:/Users/79042/OneDrive/Рабочий стол/График отпусков_2023.xlsx', sheet_name='Sheet1')



print(pandas_dataframe)