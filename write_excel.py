import xlrd
from xlutils.copy import copy
import time
def loadCollectExcel(path):
    r = xlrd.open_workbook(path)
    w = copy(r)
    wSheet = w.get_sheet(2)
    rSheet = r.sheet_by_index(2)
    return w, wSheet, rSheet

