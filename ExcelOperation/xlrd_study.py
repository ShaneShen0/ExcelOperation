
"""
A class implementing "read" operation based on library xlrd
"""
import xlrd 
class Read_Excel():
    def __init__(self,file_name):
        self.file_name = file_name
        self.workbook = xlrd.open_workbook(self.file_name)

    def print_workbook(self):     
        print(self.workbook)

    def print_values(self):
        """ print all values in the excel file """
        for s in self.workbook.sheets():
            print('Sheet:',s.name)
            print("###########################")
            for row in range(s.nrows):
                values = []
                for col in range(s.ncols):
                    # append a value cell into the list 'values' 
                    values.append(s.cell(row,col).value)
                # print all values in the list 'values', that is to print current row 
                for value in values:
                    print(value,end=' ') 
                print()
  

    def print_first_row_col(self,sheetIndex,row,col):
        """ 
            Print the first several rows and cols

            :param row : the number of the first few rows 
            :param col : the number of the first few columns 
        """
        # if row or col are minus, then let them be 0
        if row < 0 : row = 0 
        if col < 0 : col = 0
        if sheetIndex< 0 : sheetIndex=0
        sheet = self.workbook.sheet_by_index(sheetIndex)
        for r in range(row):
            values = []
            for c in range(col):
                values.append(str(sheet.cell(r,c).value))
            for value in values:
                print(value,end=' ')
            print() 
        print(sheet.cell_value(0,0))


    def change_back_to_int(self,floatCell):
        """ 
        If the decimal part of a cell.value is 0, then let the value be a integer
        Cause xlrd will change number value to float automaticlly

        :param floatCell : the float value cell to be changed
        """
        if floatCell is not None:
            pass

    def format_to_list(self,sheetIndex=0):
        """
        Format a sheet in the workbook to a two dimension list
        :param sheetIndex : the index of the sheet which will be formatted
        """
        if sheetIndex < 0:
            return []
        sheet = self.workbook.sheet_by_index(sheetIndex)
        rows = sheet.nrows
        cols = sheet.ncols
        resultList = []
        for row in range(rows):
            rowList = []
            for col in range(cols):
                cell = sheet.cell(row,col)
                if cell.ctype == 2 and cell.value % 1 == 0.0 :
                    rowList.append(int(cell.value))
                else:
                    rowList.append(cell.value)
            resultList.append(rowList)
        return resultList
    
########## testing ###############
re = Read_Excel("pkxx.xlsx")
re.print_workbook()
# re.print_values()
# re.print_first_row_col(0,5,5)
print(re.format_to_list())