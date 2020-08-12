import os
import sys
import xlrd
import time
from write_excel import loadCollectExcel


basePath = os.getcwd()
directories = os.listdir(basePath)


class ShareThreeTarget:
    def __init__(self, path, name):
        excel = xlrd.open_workbook(path)
        self.name = name
        self.sheet = excel.sheet_by_index(0)
        self.rateList = []
        self.qjian = []
        self.threeTargetCount = 0
        self.inTotalCount = 0
        self.totalRate = 1
        self.vMap7Index = 0
        self.RsiIndex = 0
        self.DifIndex = 0
        self.DeaIndex = 0
        self.initSheet()
        self.computeRateList()

    def initSheet(self):
        titles = self.sheet.row_values(0)
        for (index, title) in enumerate(titles):
            if title == 'Price':
                self.priceIndex = index
            elif title == 'RSI (14)':
                self.RsiIndex = index
            elif title == 'DIF':
                self.DifIndex = index
            elif title == 'DEA':
                self.DeaIndex = index
            elif title == 'VWAP 7':
                self.vMap7Index = index

    def computeRateList(self):
        startIndex = 1
        lastIndex = self.sheet.nrows
        inPrice = 0
        outPrice = 0
        totalRate = 1
        threeTargetCount = 0
        rateList = []
        isTreeTarget = False
        while startIndex < lastIndex:
            rowData = self.sheet.row_values(startIndex)
            rsi = rowData[self.RsiIndex] or 0
            dif = rowData[self.DifIndex] or 0
            dea = rowData[self.DeaIndex] or 0
            price = rowData[self.priceIndex] or 0
            vMap7 = rowData[self.vMap7Index] or 0
            # 满足3指标
            startIndex = startIndex + 1
            if startIndex == lastIndex:
                break
            if rsi < 70 and rsi > 50 and dif > dea and dea > 0 and price > vMap7:
                self.inTotalCount = self.inTotalCount + 1
                threeTargetCount = threeTargetCount + 1
                if threeTargetCount == 3:
                    self.threeTargetCount = self.threeTargetCount + 1
                if not isTreeTarget:
                    inPrice = self.sheet.row_values(startIndex)[self.priceIndex]
                    isTreeTarget = True
            else:
                threeTargetCount = 0
                if isTreeTarget:
                    isTreeTarget = False
                    outPrice = self.sheet.row_values(startIndex)[self.priceIndex]
            
            if inPrice != 0 and outPrice != 0:
                self.qjian.append([inPrice, outPrice])
                rate = (outPrice - inPrice) / outPrice
                totalRate *= (rate + 1)
                rateList.append(rate)
                inPrice = 0
                outPrice = 0
        self.totalRate = totalRate - 1
        self.rateList = rateList
    
    def proFit(self):
        return list(filter(lambda x: x > 0, self.rateList))

if __name__ == "__main__":
    # path = f'{basePath}\\the chosen data\\AAL Historical Data.xlsx'
    # shareThreeTarget = ShareThreeTarget(path, 'AAL Historical Data.xlsx')
    # print(f'区间：{shareThreeTarget.qjian}')
    # print(f'买入次数：{shareThreeTarget.inTotalCount}')
    # print(f'连续三天满足：{shareThreeTarget.threeTargetCount}')
    # print(f'三指标总收益率：{shareThreeTarget.totalRate}')
    # print(f'三指标有效盈利信号的出现次数：{len(shareThreeTarget.proFit())}')
    
    collectionPath = 'data collection 最新版本.xlsx'
    w, wSheet, rSheet = loadCollectExcel(collectionPath)
    shareThreeTargetDict = {}
    for d in directories:
        path = f'{basePath}\\{d}'
        if os.path.isdir(path) and d != '__pycache__' and d != '.git':
            excelNames = os.listdir(path)
            for excelName in excelNames:
                if excelName.find('~$') != -1:
                    continue
                excelPath = f'{path}\\{excelName}'
                shareName = excelName.split(' ')[0]
                print(excelPath)
                shareThreeTarget = ShareThreeTarget(excelPath, shareName)
                shareThreeTargetDict[shareName] = shareThreeTarget

    print('------ compute finsih.start save-------' )
    for i in range(2, rSheet.nrows):
        rowData = rSheet.row_values(i)
        if rowData[0] == '' or rowData[0] not in shareThreeTargetDict:
            continue
        shareThreeTarget = shareThreeTargetDict[rowData[0]]
        proFitList = shareThreeTarget.proFit()
        wSheet.write(i, 9, shareThreeTarget.inTotalCount)
        wSheet.write(i, 10, shareThreeTarget.threeTargetCount)
        wSheet.write(i, 11, f'{round(shareThreeTarget.totalRate * 100, 2)}%')
        wSheet.write(i, 12, len(proFitList))

    filename = str(time.time()).split('.')[0] + collectionPath
    w.save(filename)
    print(f'------ save {filename} finish-------' )
    