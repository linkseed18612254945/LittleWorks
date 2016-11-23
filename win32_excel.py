from win32com.client import Dispatch
import win32com.client


class easyExcel:
    """A utility to make it easier to get at Excel.    Remembering
    to save the data is your problem, as is    error handling.
    Operates on one workbook at a time."""

    def __init__(self, sheet, filename=None):
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

        self.sht = self.xlBook.Worksheets(sheet)

    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()

    def close(self):
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp

    def getCell(self, row, col):
        "Get value of one cell"
        return self.sht.Cells(row, col).Value

    def setCell(self, row, col, value):
        "set value of one cell"
        self.sht.Cells(row, col).Value = value

    def getRange(self, row1, col1, row2, col2):
        "return a 2d array (i.e. tuple of tuples)"
        return self.sht.Range(self.sht.Cells(row1, col1), self.sht.Cells(row2, col2)).Value

    def addPicture(self, sheet, pictureName, Left, Top, Width, Height):
        "Insert a picture in sheet"
        sht = self.xlBook.Worksheets(sheet)
        sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)

    def cpSheet(self, before):
        "copy sheet"
        shts = self.xlBook.Worksheets
        shts(1).Copy(None, shts(1))


if __name__ == "__main__":
    PNFILE = r'c:/screenshot.bmp'
    xls = easyExcel(r'C:\Users\51694\PycharmProjects\LittleWorks\薪酬数据\000001.xls')
    print(xls.getCell(3,1))