import xml.etree.ElementTree as ET
from openpyxl import load_workbook
import uuid

# Загружаем файл экселя
wb = load_workbook('Координаты пунктов 100 охранных зон пунктов ГГС.xlsx')


# печатаем список листов
# sheets = wb.sheetnames
# for sheet in sheets:
#     print(sheet)

# Делаем первый лист активным
sheet = wb.active
# Создаем вложенный список, где каждый элемент это строка из эксель таблицы
total_sp = []
for i in range(2, sheet.max_row):
    sp = []
    for column in sheet.iter_cols(2, sheet.max_column):
        sp.append(column[i].value)
    total_sp.append(sp)

# получаем другой лист
sheet2 = wb['на здании ГГС']


# Делаем второй лист активным
wb.active = 1

for i in range(2, sheet2.max_row):
    sp = []
    for column in sheet2.iter_cols(2, sheet2.max_column):
        if column[i].value == None:
            continue
        else:
            sp.append(column[i].value)
    total_sp.append(sp)
# Удаляем последний элемент (он пустой список)
del total_sp[-1]
# В первом цикле добавляем X и Y в конец списка, 2 циклом удаляем NONE
for gg in total_sp:
    gg.append(gg[2])
    gg.append(gg[3])
for nn in total_sp:
    if nn[0] == None:
        indx = total_sp.index(nn)
        del total_sp[indx]
# for vv in total_sp:
#     print(vv)

# Генерируем рандомный GUID
guid = uuid.uuid4()
# Конвертируем в строку
str_guid = str(uuid.uuid4())

# # Добавляем пространство имен, если есть
ET.register_namespace("", "urn://x-artefacts-rosreestr-ru/incoming/territory-to-gkn/1.0.4")
ET.register_namespace("p1", "http://www.w3.org/2001/XMLSchema-instance")
ET.register_namespace("Spa2", "urn://x-artefacts-rosreestr-ru/commons/complex-types/entity-spatial/2.0.1")
ET.register_namespace("CadEng4", "urn://x-artefacts-rosreestr-ru/commons/complex-types/cadastral-engineer/4.1.1")
ET.register_namespace("Doc5", "urn://x-artefacts-rosreestr-ru/commons/complex-types/document-info/5.0.1")
ET.register_namespace("tns", "urn://x-artefacts-smev-gov-ru/supplementary/commons/1.0.1")
ET.register_namespace("schemaLocation", "urn://x-artefacts-rosreestr-ru/incoming/territory-to-gkn/1.0.4 TerritoryToGKN_v01.xsd")


#
# # Парсим хмл
tree = ET.parse('Территории.xml')
root = tree.getroot()

# Проходимся по общему списку и в каждом меняем координаты
for i in range(len(total_sp)):
    ychastok =+ i
    raion = total_sp[ychastok][1]
    print(len(total_sp[ychastok][11:]))
    x = 2
    y = 3
    count = 1
    number_tochki = 6
    number_coord = 5
    for neighbor in root.iter('{urn://x-artefacts-rosreestr-ru/commons/complex-types/entity-spatial/2.0.1}Ordinate'):
        DeltaGeopoint = str(neighbor.attrib['DeltaGeopoint'])
        GeopointOpred = str(neighbor.attrib['GeopointOpred'])
        attrib_1 = {"TypeUnit": "Точка", "SuNmb": str(number_tochki)}
        attrib_2 = {"X": str(total_sp[ychastok][x]), "Y": str(total_sp[ychastok][y]), "NumGeopoint": str(number_tochki),
                    "DeltaGeopoint": DeltaGeopoint,
                    "GeopointOpred": GeopointOpred}
        if len(total_sp[ychastok][11:]) == 1:
            neighbor.attrib['X'] = total_sp[ychastok][x]
            neighbor.attrib['Y'] = total_sp[ychastok][y]
            x += 2
            y += 2
        if len(total_sp[ychastok][11:]) > 1:
            neighbor.attrib['X'] = total_sp[ychastok][x]
            neighbor.attrib['Y'] = total_sp[ychastok][y]
            x += 2
            y += 2
            count += 1
            if count == 6: # шестерка потому что, после пятой точки , становится 6 и тогда заходит в этот if.
                root[1][0][4].attrib['SuNmb'] = '5'
                neighbor.attrib['NumGeopoint'] = "5"
                attrib_2['X'] = str(total_sp[ychastok][x])
                attrib_2['Y'] = str(total_sp[ychastok][y])
                count += 1
                x += 2
                y += 2
            if count > 6:
                konec = len(total_sp[i][12:])
                while konec != 2:
                    ET.SubElement(root[1][0],
                                  "{urn://x-artefacts-rosreestr-ru/commons/complex-types/entity-spatial/2.0.1}SpelementUnit",
                                  attrib_1)
                    ET.SubElement(root[1][0][number_coord],
                                  "{urn://x-artefacts-rosreestr-ru/commons/complex-types/entity-spatial/2.0.1}Ordinate ",
                                  attrib_2)
                    number_tochki += 1
                    attrib_1["SuNmb"] = str(number_tochki)
                    number_coord += 1
                    attrib_2['X'] = str(total_sp[ychastok][x])
                    attrib_2['Y'] = str(total_sp[ychastok][y])
                    attrib_2['NumGeopoint'] = str(number_tochki)
                    x += 2
                    y += 2
                    konec -= 2
                attrib_1["SuNmb"] = str(1)
                attrib_2['NumGeopoint'] = str(1)
                ET.SubElement(root[1][0],
                              "{urn://x-artefacts-rosreestr-ru/commons/complex-types/entity-spatial/2.0.1}SpelementUnit",
                              attrib_1)
                ET.SubElement(root[1][0][number_coord],
                              "{urn://x-artefacts-rosreestr-ru/commons/complex-types/entity-spatial/2.0.1}Ordinate ",
                              attrib_2)
    tree.write(f'.//Готовые//{total_sp[ychastok][0]} {total_sp[ychastok][1][:2]}_{total_sp[ychastok][1][3:]}.xml',
               encoding='utf-8', xml_declaration = True)

# print(root[1][0][0])


# for xx in root.iter("{urn://x-artefacts-rosreestr-ru/commons/complex-types/entity-spatial/2.0.1}SpatialElement"):
#     print(xx[0].attrib)

"""Для запоминания"""
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
