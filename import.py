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
ponderSet=[]    # 应答器信息
staSet=[]       # 车站位置信息
signalSet=[]    # 信号机信息

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

# 定义四个属性
B_Name = ponderSet[1]
B_Num = ponderSet[2]
B_Location = ponderSet[3]
B_Type = ponderSet[4]

# 定义其他要用到的变量
S_Location = signalSet[3]   # 信号机里程
Sta_DQu = staSet[2]          # 大区号
Sta_FQu = staSet[3]         # 分区号
Sta_CZ = staSet[4]          # 车站好

# 定义正确的值
# B_trueLocation = ''

# 获取+号后面里程信息，并返回
def getLocation(location):
    sg_location=location.split('+')[1]
    sg_location=int(sg_location)
    return sg_location

# 判断里程是否正确
def judgeLocation():
    sg_location=getLocation(S_Location)
    ponder_location=getLocation(B_Location)
    true_location=sg_location+30
    if true_location>ponder_location:
        if true_location<100 :
            true_location='0'+str(true_location)
        B_trueLocation = B_Location.split('+')[0]+'+'+str(true_location)
        print('里程错误！正确的里程为:')
        print(B_trueLocation)
        return B_trueLocation
        # print(B_Location)
    else:
        print('里程正确!')
        return B_Location

def nameIsTrue():
    B_trueLocation = judgeLocation()
    initials = B_Name[0]
    # 判断首字母是否为 'B'
    if (initials == 'B'):
        print('首字母正确')
        print(B_trueLocation)
    else:
        print('首字母错误')
    shuzi = B_trueLocation.split('K')[1]
    _distance = shuzi.replace('+', '')   # 将 '+' 消除
    distance = int(_distance[0:4])      # 取前四位
    if (distance % 2 == 0):
        trueDistance = distance + 1     # 得到正确的公里标
    else:
        trueDistance = distance
    Km_mark = int(B_Name[1:5])                 # 提取名称的公里标，ykksdadsad -> [y,k,k,s,d,a]
    if(trueDistance == Km_mark):
        print('公里标正确！')
    else:
        print('公里标错误！')
        print('正确的公里标为:' + str(trueDistance))
    # 284018 + 30 得到的是有源应答器的里程
    num = B_Name.split('-')[1]
    if(num == '1'):
        print('组内编号正确!')
    else:
        print('组内编号错误!')
    if(B_Type == '有源'):
        print('应答器类型正确！')
    else:
        print('应答器类型错误!')

print(ponderSet)
print(staSet)
# judgeLocation()
nameIsTrue()