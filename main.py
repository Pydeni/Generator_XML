import xml.etree.ElementTree as ET
from openpyxl import load_workbook

# Загружаем файл экселя
wb = load_workbook('Координаты пунктов 100 охранных зон пунктов ГГС.xlsx')
ws = wb.worksheets[0]

# Добавляем данные из книги в список
sheet = wb.active  # Активный лист в книге

# Наименования пунктов в столбце В2
name_punkt = []
for val in sheet['B'][2:]:
    name_punkt.append(val.value)

# Номера районов в столбце С2
name_raion = []
for val in sheet['C'][2:]:
    name_raion.append(val.value)

# Точка 1x, ячейка D2
name_tocha_1x = []
for val in sheet['D'][2:]:
    name_tocha_1x.append(val.value)

# Точка 1y, ячейка E2
name_tocha_1y = []
for val in sheet['E'][2:]:
    name_tocha_1y.append(val.value)

# Точка 2x, ячейка F2
name_tocha_2x = []
for val in sheet['F'][2:]:
    name_tocha_2x.append(val.value)

# Точка 2y, ячейка G2
name_tocha_2y = []
for val in sheet['G'][2:]:
    name_tocha_2y.append(val.value)

# Точка 3x, ячейка H2
name_tocha_3x = []
for val in sheet['H'][2:]:
    name_tocha_3x.append(val.value)

# Точка 3y, ячейка I2
name_tocha_3y = []
for val in sheet['I'][2:]:
    name_tocha_3y.append(val.value)

# Точка 4x, ячейка J2
name_tocha_4x = []
for val in sheet['J'][2:]:
    name_tocha_4x.append(val.value)

# Точка 4y, ячейка K2
name_tocha_4y = []
for val in sheet['K'][2:]:
    name_tocha_4y.append(val.value)


# Объединяем все списки в один список кортежей
total_spisok = list(zip(name_punkt, name_raion, name_tocha_1x, name_tocha_1y, name_tocha_2x, name_tocha_2y,
                        name_tocha_3x, name_tocha_3y, name_tocha_4x, name_tocha_4y))



# Добавляем пространство имен, если есть

ET.register_namespace("", "urn://x-artefacts-rosreestr-ru/incoming/territory-to-gkn/1.0.4")
ET.register_namespace("p1", "http://www.w3.org/2001/XMLSchema-instance")
ET.register_namespace("Spa2", "urn://x-artefacts-rosreestr-ru/commons/complex-types/entity-spatial/2.0.1")
ET.register_namespace("CadEng4", "urn://x-artefacts-rosreestr-ru/commons/complex-types/cadastral-engineer/4.1.1")
ET.register_namespace("Doc5", "urn://x-artefacts-rosreestr-ru/commons/complex-types/document-info/5.0.1")
ET.register_namespace("tns", "urn://x-artefacts-smev-gov-ru/supplementary/commons/1.0.1")
ET.register_namespace("schemaLocation", "urn://x-artefacts-rosreestr-ru/incoming/territory-to-gkn/1.0.4 TerritoryToGKN_v01.xsd")



# Парсим хмл
tree = ET.parse('Территории.xml')
root = tree.getroot()

for i in range(1):
    b =+ i
    count = 1
    x = 2
    y = 3
    for neighbor in root.iter('{urn://x-artefacts-rosreestr-ru/commons/complex-types/entity-spatial/2.0.1}Ordinate'):
        print(neighbor.attrib)
        if count == 5:
            neighbor.attrib['X'] = total_spisok[b][2]
            neighbor.attrib['Y'] = total_spisok[b][3]
        else:
            neighbor.attrib['X'] = total_spisok[b][x]
            neighbor.attrib['Y'] = total_spisok[b][y]
            x += 2
            y += 2
            count += 1


    tree.write(f'{total_spisok[b][0]}.xml', encoding='utf-8', xml_declaration = True)
for sp in total_spisok[:1]:
    print(sp)




# Получается словарь (тег - ключ(name), attr - значение(название листа)).
# for child in root:
#     print(child.tag, child.attrib)
"""worksheet {'name': 'грунтовые ГГС'}
worksheet {'name': 'на здании ГГС'}"""

#  Получить по индексу значение ячейки (.текст -  переводит в текст)
# print(root[0][0][1].text)
"""Район"""

# Заменяем определенную колонку на нужное значение и создаем новый файл хмл
# for raion in root.iter('Column3'):
#     raion.text = '50:15'

# Создаем столько файлов, сколько кортежей в общем списке, файлы создаются по названию поселка(в данном случае)
# for i in range(len(total_spisok)):
#     b =+ i
#     tree.write(f'{total_spisok[b][0]}.xml', encoding='utf-8')
