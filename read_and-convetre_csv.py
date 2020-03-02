from numpy.core import unicode, long
from openpyxl import load_workbook
import datetime

import xml.etree.cElementTree as ET

wb = load_workbook(filename='moy.xlsx')

sheet = wb.get_sheet_by_name('Output Sheet')  # sheet name

root = ET.Element("skuProducts")

i = 0
for rowOfCellObjects in sheet.rows:
    doc = ET.SubElement(root, "skuProduct" + str(i))
    j = 0

    for cellObj in rowOfCellObjects:
        print(cellObj.coordinate, cellObj.value)
        print
        type(cellObj.value)

        datatype = type(cellObj.value)
        if (datatype is unicode):
            ET.SubElement(doc, cellObj.coordinate).text = cellObj.value
        elif (datatype is long):
            print
            int(cellObj.value)
            ET.SubElement(doc, cellObj.coordinate).text = str(int(cellObj.value))
        elif isinstance(cellObj.value, datetime.date):
            # print cellObj.value.strftime('%m/%d/%Y')
            # pass
            ET.SubElement(doc, cellObj.coordinate).text = str(cellObj.value.strftime('%m/%d/%Y'))

        else:
            print
            "pass"
            pass

        j = j + 1

    i = i + 1
    print('--- END OF ROW ---')

tree = ET.ElementTree(root)
tree.write("moy.xml")
