# First exercise

import xlwings as xw
import pandas as pd
import datetime as dt
import matplotlib.pyplot as plt

# wb = xw.Book()  # this will create a new workbook
# wb = xw.Book('FileName.xlsx')  # connect to an existing file in the current working directory
# wb = xw.Book(r'C:\path\to\file.xlsx')  # on Windows: use raw strings to escape backslashes
wb = xw.Book(r'C:\Users\BC108568\Desktop\AI_Test_1.xlsx')

sht = wb.sheets['Sheet1']

sht.range('A1').value = 'Foo 1'

sht.range('B4').value = [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
# sht.range('A1').expand().value

df = pd.DataFrame([[1, 2], [3, 4]], columns=['a', 'b'])
sht.range('E10').value = df

# fig = plt.figure()
# plt.plot([0, 2, 1],[452, 462, 457])
# sht.pictures.add(fig, name='MyPlot', update=True)

sht.range('B2').value = dt.datetime(1993, 11, 12)
