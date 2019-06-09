# -*- coding: utf-8 -*-
"""
Created on Mon Jun  3 00:56:26 2019

@author: DELL PC
"""

import pandas as pd
import matplotlib.pyplot as plt
excel_file = "movies.xls"
movies = pd.read_excel(excel_file)
movies.head()

##reading data by sheetwise
movies_sheet1 = pd.read_excel(excel_file,sheetname=0,index_col=0)
movies_sheet1.head()
movies_sheet2 = pd.read_excel(excel_file,sheetname=1,index_col=0)
movies_sheet2.head()
movies_sheet3 = pd.read_excel(excel_file,sheetname=2,index_col=0)
movies_sheet3.head()

movies = pd.concat([movies_sheet1,movies_sheet2,movies_sheet3])
movies.shape

##using the excelfile class to read multiple sheets
xlsx = pd.ExcelFile('movies.xls')
movies_sheet = []
for sheet in xlsx.sheet_names:
    movies_sheet.append(xlsx.parse(sheet))
movies_data = pd.concat(movies_sheet)   
movies_data.shape

##plotting
sorted_by_gross = movies.sort_values(['Gross Earnings'],ascending=False)
print(sorted_by_gross['Gross Earnings'].head(10))

sorted_by_gross['Gross Earnings'].head(10).plot(kind='bar')
plt.show()


movies['IMDB Score'].plot(kind='hist')
plt.show()

##getting stastical information about the data
movies.describe()
movies['Gross Earnings'].mean()

##applying formula in columns
movies['Net Earnings'] = movies['Gross Earnings'] - movies['Budget']
sort = movies.sort_values(['Net Earnings'],ascending=False)
print(sort['Net Earnings'].head(10))

sort['Net Earnings'].head(10).plot(kind='bar')
plt.show()

##pivot table in python
movies_subset = movies[['Year','Gross Earnings']]
print(movies_subset)

pivot = movies_subset.pivot_table(index='Year')
print(pivot)

pivot.plot()
plt.show()

movies_subset = movies[['Country','Language','Gross Earnings']]
print(movies_subset)

pivot = movies_subset.pivot_table(index=['Country','Language'])
print(pivot)

pivot.head(20).plot(kind='bar',figsize=(10,15))
plt.show()

##exporting the results to excel
movies.to_excel('output.xlsx')
print(movies.head())
movies.to_excel('output.xlsx',index=False)
print(movies.head())

writer = pd.ExcelWriter('output.xlsx',engine='xlsxwriter') 
movies.to_excel(writer,index=False,sheet_name='report')
workbook = writer.book
worksheet = writer.sheets['report']
header_fmt = workbook.add_format({'bold': True})
worksheet.set_row(0, None, header_fmt)
writer.save()


















