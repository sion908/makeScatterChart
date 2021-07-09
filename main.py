import openpyxl
import numpy as np
from openpyxl.chart import BarChart, Reference, Series,ScatterChart

def main(wb):
    X=makeXs()
    row ,ws=writeExcel(wb,X,np.sin(X),"wani","1")
    row ,ws=writeExcel(wb,None,np.cos(X),"wani","2")

    #ScatterChartオブジェクト作成
    chart = ScatterChart()

    #xの範囲を設定
    min_row = 2
    max_row = row
    x_values = Reference(ws, min_col = 1, min_row = min_row, max_row = max_row)

    #yの範囲を設定
    min_col = 2
    values = Reference(ws, min_col = min_col, min_row = min_row, max_row = max_row)
    #グラフの追加
    series = Series(values, x_values, title="ScatterChart")
    chart.series.append(series)
    ws.add_chart(chart, "D16")

#end main

def makeXs(minNum=-1,maxNum=1,count=100):

    x =np.linspace(minNum,maxNum,count+1)
    return x
#end makeXs


def writeExcel(wb,X,Y,sname,dataname):
    found = True
    for name in wb.sheetnames:
        if name == sname:
            ws = wb[name]
            found = False
            col = ws.max_column +1
            break
    if found:
        ws = wb.create_sheet(title=sname)
        col = 1
    
    
    # シートの追加
    # ws4 = wb.create_sheet(title="Sheet4")
    row=1
    # セルの指定
    # c1=ws.cell(row=1,column=1)
    print(type(X))
    if type(X)==type(None):
        ws.cell(row = row, column = col, value = dataname)
        row +=1
        for y in Y:
            #numpy.float64のままだとValueErrorが出るのでキャスト
            ws.cell(row = row, column = col, value = float(y))
            row+=1
    else:
        ws.cell(row = row, column = col, value = "X"+dataname)
        ws.cell(row = row, column = col+1, value = dataname)
        row +=1
        for x,y in zip(X,Y):
            #numpy.init32のままだとValueErrorが出るのでキャスト
            ws.cell(row = row, column = col, value = float(x))
            #numpy.float64のままだとValueErrorが出るのでキャスト
            ws.cell(row = row, column = col + 1, value = float(y))
            row+=1

    return row,ws

#end writeExxcel

def Preprocessing():
    
    wb = openpyxl.Workbook()
    # wb.remove(wb.worksheets)
    sheets =wb.worksheets
    main(wb)
    for sheet in sheets:
        wb.remove(sheet)
    wb.save("wa.xlsx")
#end Preprocessing

if __name__=="__main__":
    Preprocessing()
    # main(wb)
#end ifmain