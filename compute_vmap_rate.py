import os
import sys
import xlrd
import time
from write_excel import loadCollectExcel


basePath = os.getcwd()
directories = os.listdir(basePath)


class ShareVMap:
    def __init__(self, path, name):
        excel = xlrd.open_workbook(path)
        self.name = name
        self.sheet = excel.sheet_by_index(0)
        self.vMap7rateList = []
        self.vMap14rateList = []
        self.vMap7TotalRate = 1
        self.vMap14TotalRate = 1
        self.vMap7Index = 0
        self.vMap14Index = 0
        self.priceIndex = 0
        self.initSheet()
        self.vMap7TotalRate, self.vMap7rateList = self.computeVmap7OrVmap14(self.vMap7Index)
        self.vMap14TotalRate, self.vMap14rateList = self.computeVmap7OrVmap14(self.vMap14Index)

    def initSheet(self):
        titles = self.sheet.row_values(0)
        for (index, title) in enumerate(titles):
            if title == 'Price':
                self.priceIndex = index
            elif title == 'VWAP 7':
                self.vMap7Index = index
            elif title == 'VWAP 14':
                self.vMap14Index = index

    def computeVmap7OrVmap14(self, vMapIndex):
        startIndex = 1
        lastIndex = self.sheet.nrows
        inNum = 0
        inPrice = 0
        outPrice = 0
        totalRate = 1
        rateList = []
        while startIndex < lastIndex:
            rowData = self.sheet.row_values(startIndex)
            if rowData[vMapIndex + 1] == '√':
                inNum = inNum + 1
                startIndex = startIndex + 1
                if inNum == 2:
                    inPrice = rowData[self.priceIndex]
            elif rowData[vMapIndex + 1] == 'X':
                startIndex = startIndex + 1
                if inNum >= 2:
                    if startIndex == lastIndex:
                        startIndex = startIndex - 1
                    rowData = self.sheet.row_values(startIndex)
                    outPrice = rowData[self.priceIndex]
                    startIndex = startIndex + 1
                inNum = 0
            else:
                startIndex = startIndex + 1
            
            if inPrice != 0 and outPrice != 0:
                rate = (outPrice - inPrice) / outPrice
                totalRate *= (rate + 1)
                rateList.append(rate)
                inPrice = 0
                outPrice = 0

        return (totalRate - 1), rateList
    
    def proFit(self):
        return list(filter(lambda x: x > 0, self.vMap7rateList)), list(filter(lambda x: x > 0, self.vMap14rateList))

if __name__ == "__main__":
    # path = f'{basePath}\\the chosen data\\AAL Historical Data.xlsx'
    # shareVMap = ShareVMap(path, 'AAL Historical Data.xlsx')
    collectionPath = 'data collection 最新版本.xlsx'
    w, wSheet, rSheet = loadCollectExcel(collectionPath)
    shareVMapDict = {}
    for d in directories:
        path = f'{basePath}\\{d}'
        if os.path.isdir(path) and d != '__pycache__':
            excelNames = os.listdir(path)
            for excelName in excelNames:
                if excelName.find('~$') != -1:
                    continue
                excelPath = f'{path}\\{excelName}'
                shareName = excelName.split(' ')[0]
                print(excelPath)
                shareVMap = ShareVMap(excelPath, shareName)
                shareVMapDict[shareName] = shareVMap

    print('------ compute finsih.start save-------' )
    for i in range(2, rSheet.nrows):
        rowData = rSheet.row_values(i)
        if rowData[0] == '' or rowData[0] not in shareVMapDict:
            continue
        shareVMap = shareVMapDict[rowData[0]]
        vMap7ProFitList, vMap14ProFitList = shareVMap.proFit()
        wSheet.write(i, 3, len(vMap7ProFitList))
        wSheet.write(i, 4, f'{round(shareVMap.vMap7TotalRate * 100, 2)}%')
        wSheet.write(i, 7, len(vMap14ProFitList))
        wSheet.write(i, 8, f'{round(shareVMap.vMap14TotalRate * 100, 2)}%')

    filename = str(time.time()) + collectionPath
    w.save(filename)
    print(f'------ save {filename} finish-------' )
    