
######### Tìm ra URL
from googlesearch import search
import xlwings as xw
from time import sleep


wb = xw.Book(r'C:\Users\nglep\Downloads\New folder\Buy_Account_LMHT\DATA_T8.2021_Copy.xlsx')
sh1 = wb.sheets('XK')

start = 1284
end = 1384

List_Company_Name_Old = sh1[f'H{start}:H{end}'].value
URL = ''

for index in range(len(List_Company_Name_Old)):
    sleep(4)

    URL = search(List_Company_Name_Old[index])

    try:

        sh1[f'AE{index + start}'].value = URL[0]
    except:
        sh1[f'AE{index + start}'].value = ''

