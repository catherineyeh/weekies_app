import datetime
import string
from openpyxl import load_workbook

class Members:
    """information of members"""
    department = ""
    name = ""
    isPtoWfh = False
    lastDate = datetime.date(1,1,1)
    gap = 0
    count = 0
    #def __init__(self, count, department, name, lastDate, isPtoWfh):
     #   self.count = count
      #  self.department = department
       # self.name = name
        #self.lastDate = lastDate
        #self.isPtoWfh = isPtoWfh
        #return super().__init__(count, department, name, lastDate, isPtoWfh)

    def oneMember(self, sheet, rowIndex):
        for row_cells in sheet.iter_rows(min_row=rowIndex, max_row=rowIndex):
            index = 0
            while (index < len(row_cells)):
                self.count = row_cells[index].value
                index = index + 1
                self.department = row_cells[index].value
                index = index + 1
                self.name = row_cells[index].value
                index = index + 1
                self.lastDate = row_cells[index].value
        return self
    
