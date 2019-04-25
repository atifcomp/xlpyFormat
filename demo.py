from workbook import FormatWriter
from worksheet import worksheet
sourceFile = "C:/Users/amomin/Desktop/Projects/Charter ISP/Days Aging.xlsx"
wb1 = FormatWriter(sourceFile)
ws1 = worksheet(wb1.book,'abc')

# In[]:


wb1.save_workbook()


# In[]:
from openpyxl.reader.excel import load_workbook
from openpyxl.styles.alignment import Alignment
book = load_workbook(filename = sourceFile)
ws3 = book['abc']


# In[]:

#set font of column
font_format = { 'name':'Calibri',
                'size':20,
                'bold':False,
                'italic':False,
                'vertAlign':None,
                'underline':'none',
                'strike':False,
                'color':'FF000000'}

ws1.column_apply_font('A:F', font_format)

#set all borders
ws1.set_all_borders('A1:C55')

#Formating Example
ws1.column_set_format('A:A','#,##0.00')
ws1.column_set_format('B:C','mm-dd-yy')    
ws1.column_set_format('D:E','"$"#,##0_);[Red]("$"#,##0)')

ws1.set_format('A:Z','"$"#,##0_);[Red]("$"#,##0)')               

#range widht and column width example
ws1.column_range_width(col_rng = 'A:ZZ',col_width = 30)
ws1.column_width(col_name='A',col_width = 12)

#center alignment
ws1.column_center_align('A:D')

#column autofit example
ws1.column_autofit()

#multiple workbooks
sourceFile2 = "C:/Users/amomin/Desktop/Projects/Charter ISP/Days Aging_2.xlsx"
wb2 = FormatWriter(sourceFile2)


# In[]:













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