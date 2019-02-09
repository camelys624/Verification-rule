import xlrd
import xlwt
# path
transponder='./data/transponder.xls'    # 应答器
station='./data/station.xls'    # 车站位置信息
signal='./data/signalData.xls'  # 信号机信息

ponder_wb=xlrd.open_workbook(filename=transponder)
station_wb=xlrd.open_workbook(filename=station)
signal_wb=xlrd.open_workbook(filename=signal)
ponder_ws=ponder_wb.sheet_by_name('下行')
station_ws=station_wb.sheet_by_name('Sheet1')
signal_ws=signal_wb.sheet_by_name('下行正向')

# 定义数组
ponderSet=[]
staSet=[]
signalSet=[]

# range传要循环的数组
for c in range(ponder_ws.ncols):
    # append是数组的一个方法
    ponderSet.append(ponder_ws.cell(2,c).value)

for s in range(station_ws.ncols):
    staSet.append(station_ws.cell(2,s).value)

for r in range(signal_ws.nrows):
    if(signal_ws.cell(r,2).value=='SHF'):
        for s in range(signal_ws.ncols):
            signalSet.append(signal_ws.cell(r,s).value)
from pprint import pprint

# 获取+号后面里程信息，并返回
def getLocation(location):
    sg_location=location.split('+')[1]
    sg_location=int(sg_location)
    return sg_location

# 判断里程是否正确
def judgeLocation():
    sg_location=getLocation(signalSet[3])
    ponder_location=getLocation(ponderSet[3])
    true_location=sg_location+30
    if true_location>ponder_location:
        if true_location<100 :
            true_location='0'+str(true_location)
        print('里程错误！正确的里程为:')
        print(ponderSet[3].split('+')[0]+'+'+str(true_location))
    else:
        print('里程正确!')
        return True

judgeLocation()