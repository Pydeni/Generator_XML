# import pandas as pd
#
#
# excel = pd.read_excel('Координаты пунктов 100 охранных зон пунктов ГГС.xlsx', sheet_name='грунтовые ГГС')
# "Чтение строки по индексу"
# # print(excel.iloc[3])
# "Находим первую и последнюю строку в листе"
# # print('Первая и последняя строка', excel.iloc[[0,-1]])
# "Чтение конкретной ячейки в строке"
# # print(excel.iloc[1,1])
#
# # def chtenie(nomer):
# #      for i in excel.iloc[nomer]:
# #          print(i)
# # chtenie(3)
#
# print(excel.head())

import xml.etree.cElementTree as ET
from xml.dom import minidom
from openpyxl import load_workbook
from datetime import datetime

# Загружаем файл экселя
wb = load_workbook('Координаты пунктов 100 охранных зон пунктов ГГС.xlsx')
ws = wb.worksheets[0]

# Добавляем данные из книги в список
sheet = wb.active  # Активный лист в книге

# Наименования пунктов в столбце В2
name_punkt = []
for val in sheet['B'][2:]:
    name_punkt.append(val.value)
print(name_punkt)
# Номера районов в столбце С2
name_raion = []
for val in sheet['C'][2:]:
    name_raion .append(val.value)
print(name_raion)
# Точка 1, ячейка D2
name_tocha_1 = []
for val in sheet['D'][2:]:
    name_tocha_1 .append(val.value)
print(name_tocha_1)
# Точка 2, ячейка E2
name_tocha_2 = []
for val in sheet['D'][2:]:
    name_tocha_2 .append(val.value)
print(name_tocha_2)
