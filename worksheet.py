from openpyxl.styles import Border, Side
from openpyxl.utils.cell import get_column_letter
import re
from xlsxwriter.utility import xl_cell_to_rowcol


class worksheet():
    """
    This is the class as worksheet level
    Parameters
    ----------
    wb : workbook object, required
        this is the object of work book class of xlpyformatter
    ws_name: string, required
        this is the worksheet name parameter                
    """
    
    def __init__(self,wb,ws_name):        
        self.ws = wb[ws_name]
        self.lastRow = self.ws.max_row
        self.lastCol = self.ws.max_column         
        print(self.lastRow)
        print(self.lastCol)
        
        
    def set_all_borders(self,rng=None): 
        """
        This takes sheetname and range as input and apply all borders same as all borders of
        excel
        ----------
        rng : string, required
             range of the excel where border has to be applied
        """        
        try:
            startCell,endCell = re.split(':',rng.strip())
            minRow,minCol = xl_cell_to_rowcol(startCell)
            maxRow,maxCol = xl_cell_to_rowcol(endCell)
            
            border = Border(left=Side(border_style='thin', color='000000'),
                right=Side(border_style='thin', color='000000'),
                top=Side(border_style='thin', color='000000'),
                bottom=Side(border_style='thin', color='000000'))

            rows = self.ws.iter_rows(min_row=minRow+1,max_row=maxRow+1,min_col=minCol+1,max_col=maxCol+1)
            for row in rows:
                for cell in row:                    
                    cell.border = border
        except:
            print("set_all_borders, sheetname or ranges not provided correctly")

    def column_autofit(self):
        """
        This takes auto fits the columns of the sheet
        ----------        
        """
        for col in self.ws.columns:
            max_length = 0
            column = col[0].column_letter # Get the column name
            for cell in col:
                if cell.coordinate in self.ws.merged_cells: # not check merge_cells
                    continue
                try: # Necessary to avoid error on empty cells
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass            
            adjusted_width = (max_length + 2)
            self.ws.column_dimensions[column].width = adjusted_width
    
    def column_width(self,col_name,col_width):
        self.ws.column_dimensions[col_name].width = col_width
    
    def _sequence_check(self,firstCol,lastCol):
        """
        helper function which checks if the ranges provided are in correct order, return True or False
        ----------        
        """
        firstCol = int(firstCol.encode("utf-8").hex(),16)
        lastCol = int(lastCol.encode("utf-8").hex(),16)
        if firstCol>lastCol:
            return False
        else:
            return True

    def column_range_width(self,col_rng,col_width):        
        """
        This takes column range and column widht as input and set the width of columns
        ----------
        col_rng : string, required
             eg. 'A:Z'
        col_width : integer, required
             width to be set
        """        
        try:
            firstCol,lastCol = re.split(':',col_rng.strip())
            print(type(firstCol))
            if firstCol.isalpha() and lastCol.isalpha() and self._sequence_check(firstCol,lastCol):
                _,num1 = xl_cell_to_rowcol(firstCol+'1')
                _,num2 = xl_cell_to_rowcol(lastCol+'1')
                for _i in range(num1+1,num2+2):
                    self.ws.column_dimensions[get_column_letter(_i)].width = col_width            
            else:
                raise Exception('column range is not provided correctly')
        except:
            print("set_all_borders, sheetname or ranges not provided correctly")


        