import pandas as pd
from openpyxl.styles import Border, Side
from openpyxl.reader.excel import load_workbook
import re
from xlsxwriter.utility import xl_cell_to_rowcol




class FormatWriter():
    """
    it format excel report
    Parameters
    ----------
    srcfile : string, required
        the path of the excel file to format
        
    """
    def __init__(self,srcfile):         
        self.book = load_workbook(filename = srcfile)
        self.srcfile = srcfile
    def save_workbook(self):
        self.book.save(self.srcfile)
    