import xlrd
import xlwt
from xlutils.copy import copy
# path 路径
transponder='./data/transponder.xls'    # 应答器
station='./data/station.xls'    # 车站位置信息
signal='./data/signalData.xls'  # 信号机信息

# 打开 excel 表
ponder_wb=xlrd.open_workbook(filename=transponder, formatting_info=True)    # formatting_info=True表示保持格式不变
station_wb=xlrd.open_workbook(filename=station, formatting_info=True)
signal_wb=xlrd.open_workbook(filename=signal, formatting_info=True)
ponder_ws=ponder_wb.sheet_by_name('下行')   # 表中的数据
station_ws=station_wb.sheet_by_name('Sheet1')
signal_ws=signal_wb.sheet_by_name('下行正向')

# 定义数组
ponderSet=[]    # 应答器信息
staSet=[]       # 车站位置信息
signalSet=[]    # 信号机信息

# range传要循环的数组
for r in range(ponder_ws.nrows):    # 2
    ponder = []
    for c in range(ponder_ws.ncols):    # 2
        # append是数组的一个方法
        ponder.append(ponder_ws.cell(r,c).value)
    ponderSet.append(ponder)

for s in range(station_ws.ncols):
    staSet.append(station_ws.cell(2,s).value)

for r in range(signal_ws.nrows):    # 这个循环使循环信号机的名称
    signal = []
    signType = signal_ws.cell(r,4).value # signType 就是信号机类型
    if(signType =='出站口' or signType == '通过信号机' or signType == '进站信号机'):
        for s in range(signal_ws.ncols): # 循环这个一行
            signal.append(signal_ws.cell(r,s).value)
        signalSet.append(signal)

# 定义其他要用到的变量
CZ_C01 = ponderSet[2:5]
CZ_C02 = ponderSet[5:8]
YG = ponderSet[8:10]
ZX = ponderSet[10:12]
FYG = ponderSet[12:14]
DW = ponderSet[14:15]
JX = ponderSet[15:18]
S_Out = signalSet[0][3]   # 信号机位置
S_Through = signalSet[1][3]  # 通过信号机位置
S_In = signalSet[2][3]      # 进站信号机位置
Sta_DQu = str(int(staSet[2]))          # 大区号 float 3.0 int -> 3
Sta_FQu = str(int(staSet[3]))         # 分区号
Sta_CZ = str(int(staSet[4]))          # 车站号
R_Reference = ''                   # 定义参照点
text = '建议修改为：'

# 获取+号后面里程信息，并返回
def getLocation(location):
    sg_location=location.split('+')[1]
    sg_location=int(sg_location)
    return sg_location

# 判断里程是否正确
def verifyLocation(reference, spacing, B_Location, index, *args):
    sg_location=getLocation(reference)
    ponder_location=getLocation(B_Location)
    true_location=sg_location+spacing
    if (true_location>ponder_location):
        if true_location<100 :
            true_location='0'+str(true_location)
        B_trueLocation = B_Location.split('+')[0]+'+'+str(true_location)
        suggest = text + B_trueLocation
        print('里程错误！正确的里程为:'+B_trueLocation)
        verify(index, B_Location, suggest)
        # print(B_Location)
    else:
        if(args[0]=='ZX'):
            if(sg_location+450 < ponder_location):
                print('里程错误！正确的里程为：' + str(sg_location+450))
                suggest = text + str(sg_location+450)
                verify(index, B_Location, suggest)
        elif(args[0]=='DW'):
            if(sg_location-250 < ponder_location ):
                print('里程错误！正确的里程为：' + str(sg_location-250))
                suggest = text + str(sg_location+450)
                verify(index, B_Location, suggest)
        else:
            print('里程正确!')
        B_trueLocation = B_Location
    return B_trueLocation

# 验证名称
def verifyName(B_trueLocation, B_Name, index):
    # B_trueLocation = judgeLocation()
    flag = True
    initials = B_Name[0]
    # 判断首字母是否为 'B'
    if (initials == 'B'):
        print('首字母正确')
    else:
        flag = False
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
        flag = False
        print('公里标错误！') 
    # 284018 + 30 得到的是有源应答器的里程
    num = B_Name.split('-')[1]
    if(num == str(index)):
        print('组内编号正确!')
    else:
        print('组内编号错误!')
        flag = False
    B_trueName = 'B'+str(trueDistance)+'-'+str(index)
    suggest = text + B_trueName
    if(flag):
        print('名称正确')
    else:
        verify(index, B_Name, suggest)

# 验证编号
def verifyNum(B_Num, index, indexNum):
    value = B_Num.split('-')    # 存放切割后的数组
    num_DQu = value[0]  # 编号的大区号
    num_FQu = value[1]  # 编号的分区号
    num_CZ = value[2]    # 编号的车站号
    num_cellNum = value[3]  # 单元编号
    num_Num =  value[4]     # 应答器组内编号
    # 未来的思路
    # 现在我们的判断，对于单元编号和组内编号使一个写死的值
    # 在将来，会结合数组来实现
    if(num_DQu == Sta_DQu and num_FQu == Sta_FQu and num_CZ == Sta_CZ and
        num_cellNum == '00'+str(indexNum) and num_Num == str(index)):
        print('应答器编号正确!')
    else:
        if (num_DQu != Sta_DQu):
            print('大区编号错误!')
        if (num_FQu != Sta_FQu):
            print('分区编号错误!')
        if (num_CZ != Sta_CZ):
            print('车站号错误!')
        if (num_cellNum != '001'):
            print('单元号错误!')
        if (num_Num != '1'):
            print('组内编号错误!')
        B_trueNum = Sta_DQu+'-'+Sta_FQu+'-'+Sta_CZ+'-00'+str(indexNum)+'-'+str(index)
        suggest = text + B_trueNum
        verify(index, B_Num, suggest)

# 验证设备类型
def verifyType(index, use, ponderType, indexNum):
    def _verifyType(ponderType,trueTpye,index):
        if(ponderType==trueTpye):
            print('设备类型正确!')
        else:
            suggest = text + trueTpye
            verify(index, ponderType, suggest)
    if(use=='CZ-C01' or use=='CZ-C02'):
        if(indexNum == 0):
            _verifyType(ponderType,'有源',index)
        else:
            _verifyType(ponderType,'无源',index)
    elif(use=='JZ'):
        if(indexNum == 2):
            _verifyType(ponderType,'有源',index)
        else:
            _verifyType(ponderType,'无源',index)
    else:
        _verifyType(ponderType,'无源',index)

# 这里用到了我们之前导入的 copy
workbook = copy(ponder_wb)
worksheet = workbook.get_sheet(0)
# 设置样式
style = xlwt.easyxf('font:name 宋体, color-index red')

def verify(index, value, suggest):
    # 根据返回值判断是否需要标红
    worksheet.write(index, index, value, style)
    worksheet.write(index, index+7, suggest, style)

# for y in range(2, len(ponderSet)-3):
#     if(y == 2):
#         # 创建有源应答器实例
#         ponder = Active_Transponder(ponderSet[y][1],ponderSet[y][2],ponderSet[y][3],ponderSet[y][4])
#         verify(S_Out,30)
#         R_Reference = ponder.B_trueLocation
#     if(y == 5):
#         ponder = Active_Transponder(ponderSet[y][1],ponderSet[y][2],ponderSet[y][3],ponderSet[y][4])
#         verify(R_Reference, 200)

workbook.save('verified.xls')