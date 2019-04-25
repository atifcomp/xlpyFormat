# openpyxl.__version__ = 2.6.2
from openpyxl.reader.excel import load_workbook

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
    