# Maybe try adding something
# So logic goes as follows:
# 1) Pull from Gitkraken
# 2) Push to GitHub


# Mother (main) file

import xlwings as xw
import pandas as pd
import datetime as dt
import matplotlib.pyplot as plt

wb = xw.Book(r'C:\Users\BC108568\Desktop\Inputs.xlsx')

sht = wb.sheets[0]

