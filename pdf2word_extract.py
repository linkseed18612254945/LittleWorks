import win32com
from win32com.client import Dispatch, constants

w = win32com.client.Dispatch('Word.Application')
w.Visible = 0
w.DisplayAlerts = 0

path = 'C:/test/test.doc'
doc = w.Documents.Open( FileName = path )