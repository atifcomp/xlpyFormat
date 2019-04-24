import FormatWriter
sourceFile = "C:/Users/amomin/Desktop/Projects/Charter ISP/Days Aging.xlsx"


#Calling part starts from here
formatter = FormatWriter(sourceFile)
formatter.set_all_borders("abc","A1:Z1")
formatter.save_workbook()