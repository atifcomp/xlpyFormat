from workbook import FormatWriter
from worksheet import worksheet


sourceFile = "C:/Users/amomin/Desktop/Projects/Charter ISP/Days Aging.xlsx"
sourceFile2 = "C:/Users/amomin/Desktop/Projects/Charter ISP/Days Aging_2.xlsx"

wb1 = FormatWriter(sourceFile)
wb2 = FormatWriter(sourceFile2)



ws1 = worksheet(wb1.book,'Summary')
ws1.column_range_width(col_rng = 'A:ZZ',col_width = 3)
ws1.column_width(col_name='A',col_width = 12)


wb1.save_workbook()


from openpyxl.reader.excel import load_workbook
book = load_workbook(filename = sourceFile)
ws = book['Summary']
ws.column_dimensions['A'].width = 200
book.save(sourceFile)


















#ws1.column_width('A',100)
#
#
#st1 = "Atif:"
#st1.isalpha()
#
#data = 'Z'
#
#
#x = data.encode("utf-8").hex()
#x1 = int(x,16)
#print(x1)
#
#x
#x1
#
#s = "hello".encode("utf-8").hex()


#ws2 = worksheet(wb2.book,'Summary')
#
#
#ws1.set_all_borders("A1:D5")
#ws1.column_autofit()
#
#
#ws2.set_all_borders("A1:Z5")
#ws2.column_autofit()
#
#wb1.save_workbook()