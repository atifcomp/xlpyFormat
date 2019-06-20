from workbook import FormatWriter
from worksheet import Worksheet,Copyformat,DynamicCopyformat
import pandas as pd
sourceFile = "C:/Users/amomin/Documents/GitHub/xlpyFormat/data_demo.xlsx"
sourceFile2 = "C:/Users/amomin/Documents/GitHub/xlpyFormat/template.xlsx"

wb1 = FormatWriter(sourceFile)

fmt_wb1 = FormatWriter(sourceFile2)

fmt_sheet1 = DynamicCopyformat(wb1.book,'data_sheet',fmt_wb1.book,'template_sheet',static_rows = [1,2],var_row = 3)
fmt_sheet1.replicate_format()
wb1.save_workbook()


DynamicCopyformat

#fmt_sheet2 = Copyformat(wb1.book,'DEMO Summary',fmt_wb1.book,'DEMO Summary')
#fmt_sheet3 = Copyformat(wb1.book,'GOAL Summary',fmt_wb1.book,'GOAL Summary')
#fmt_sheet4 = Copyformat(wb1.book,'BUDGET Summary',fmt_wb1.book,'BUDGET Summary')
#fmt_sheet2.replicate_format()
#fmt_sheet3.replicate_format()
#fmt_sheet4.replicate_format()

#ws1.column_width(col_name='E',col_width = 12)
#ws1.set_all_borders('H3:H8','thick')
#ws1.set_all_borders('A3:F6','thin')




# In[]:



# In[]:

##set font of column
#font_format = { 'name':'Calibri',
#                'size':12,
#                'bold':True,
#                'italic':False,
#                'vertAlign':None,
#                'underline':'none',
#                'strike':False,
#                'color':'FF000000'}
#
#ws1.column_apply_font('B:B', font_format)
#
#
##column autofit example
#ws1.column_autofit()
#
##set all borders
#ws1.set_all_borders('A1:D11','thin')
#
#ws1.set_all_borders('D12:E20','thick')
#
#ws1.set_all_borders('F1:K1','thick')
#ws1.set_all_borders('F2:K15','thin')
#
##center alignment
#ws1.column_center_align('A:D')
#
##Formating Example
##ws1.column_set_format('A:A','#,##0.00')
##ws1.column_set_format('B:C','mm-dd-yy')    
##ws1.column_set_format('D:D','"$"#,##0_);[Red]("$"#,##0)')
#
##ws1.set_format('A:Z','"$"#,##0_);[Red]("$"#,##0)')               
#
##range widht and column width example
##ws1.column_range_width(col_rng = 'A:ZZ',col_width = 30)
##ws1.column_width(col_name='A',col_width = 12)
#
#wb1.save_workbook()
#
## In[]:
##multiple workbooks
#sourceFile2 = "C:/Users/amomin/Desktop/Projects/Charter ISP/Days Aging_2.xlsx"
#wb2 = FormatWriter(sourceFile2)
#
#
## In[]:
#
#
#
#
## In[]:
#from openpyxl.reader.excel import load_workbook
#from openpyxl.styles.alignment import Alignment
#book = load_workbook(filename = sourceFile)
#ws3 = book['abc']










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


#ws2 = Worksheet(wb2.book,'Summary')
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