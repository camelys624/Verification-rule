import xlrd
import xlwt
from xlutils.copy import copy
import os

# path 路径
transponder = './data/transponder.xls'    # 应答器
station = './data/station.xls'    # 车站位置信息
signal = './data/signalData.xls'  # 信号机信息
grade = './data/jihen-road/grade.xlsx'  # 等级转换表

# 打开 excel 表
# formatting_info=True表示保持格式不变
ponder_wb = xlrd.open_workbook(filename=transponder, formatting_info=True)
station_wb = xlrd.open_workbook(filename=station, formatting_info=True)
signal_wb = xlrd.open_workbook(filename=signal, formatting_info=True)
ponder_ws = ponder_wb.sheet_by_name('下行')   # 表中的数据
station_ws = station_wb.sheet_by_name('Sheet1')
signal_ws = signal_wb.sheet_by_name('下行正向')

# 定义数组
ponderSet = []    # 应答器信息
staSet = []       # 车站位置信息
signalSet = []    # 信号机信息

# range传要循环的数组
for r in range(ponder_ws.nrows):    # 2
    ponder = []
    for c in range(ponder_ws.ncols):    # 2
        # append是数组的一个方法
        ponder.append(ponder_ws.cell(r, c).value)
    ponderSet.append(ponder)

for s in range(station_ws.nrows):
    station = []
    for st in range(station_ws.ncols):
        station.append(station_ws.cell(s, st).value)
    staSet.append(station)

for r in range(signal_ws.nrows):    # 这个循环使循环信号机的名称
    signal = []
    signType = signal_ws.cell(r, 4).value  # signType 就是信号机类型
    if(signType == '出站口' or signType == '通过信号机' or signType == '进站信号机'):
        for s in range(signal_ws.ncols):  # 循环这个一行
            signal.append(signal_ws.cell(r, s).value)
        signalSet.append(signal)

# 定义一个存放应答器组的数组
# ponders = []
# ponders.append(ponderSet[2:5])  # 取出第一个应答器组
# ponders.append(ponderSet[5:8])
# ponders.append(ponderSet[8:10])
# ponders.append(ponderSet[10:12])
# ponders.append(ponderSet[12:14])
# ponders.append(ponderSet[14:15])
# ponders.append(ponderSet[15:18])

# print(ponders)
# 定义一个存放用途的数组
# _use = ['CZ-C01', 'CZ-C02', 'DW,YG0/2', 'DW,ZX0/2/FZX2/0', 'DW,FYG2/0', 'DW', 'JZ']

# 定义其他要用到的变量
S_Out = signalSet[0][3]   # 信号机位置
S_Through = signalSet[1][3]  # 通过信号机位置
S_In = signalSet[2][3]      # 进站信号机位置
Sta_DQu = str(int(staSet[2][2]))          # 大区号 float 3.0 int -> 3
Sta_FQu = str(int(staSet[2][3]))         # 分区号
Sta_CZ1 = str(int(staSet[2][4]))          # 颜家垄
Sta_CZ2 = str(int(staSet[3][4]))          # 长塘线路所
R_Reference = ''                   # 定义参照点
text = '建议修改为：'


# 将里程转换为数字
def getLocNum(location):
    value = location.split('K')[1]
    _distance = value.replace('+', '')
    return int(_distance)


def getUse():
    isC0_C2 = os.path.exists(grade)  # 判断等级转换表是否存在
    if(isC0_C2):
        _use = [
            ['CZ-C01', 'CZ-C02', 'DW,YG0/2',
                'DW,ZX0/2/FZX2/0', 'DW,FYG2/0', 'DW', 'JZ'],
            [3, 3, 2, 2, 2, 1, 3]
        ]
        return _use
    else:
        return []


# 验证数据是否缺失
def isMissing(use, location, reference, index):
    reference = getLocNum(reference)
    if(use == 'CZ-C01' and index == 0):
        verifyLocation = reference + 30
    elif(use == 'DW,ZX0/2/FZX2/0'):
        verifyLocation = reference + 30
    elif(use == 'JZ' and index == 0):
        verifyLocation = reference - 40
    elif(use == 'DW'):
        verifyLocation = reference - 250
    elif(index == 0):
        verifyLocation = reference + 200
    else:
        verifyLocation = reference + 5
    
    distance = verifyLocation - getLocNum(location)
    print(distance)
    if(-30 < distance < 150):
        print('数据未缺失')
        return False
    else:
        print('数据缺失')
        return True


# 判断里程是否正确
def verifyLocation(row, reference, B_Location, *args):
    if(args[0] == 'CZ-C01' and args[1] == 0):
        spacing = 30
    elif(args[0] == 'JZ' and args[1] == 0):
        spacing = -40
    elif(args[0] == 'DW'):
        spacing = -250
    elif(args[1] == 0):
        spacing = 200
    else:
        spacing = 5

    sg_location = getLocNum(reference)
    ponder_location = getLocNum(B_Location)  # 应答器位置
    true_location = sg_location+spacing
    if (true_location > ponder_location):
        B_trueLocation = 'JHK'+str(true_location)[0:3]+'+'+str(true_location)[3:6]
        suggest = text + B_trueLocation
        print('里程错误！正确的里程为:'+B_trueLocation)
        verify(row, 3, B_Location, suggest)
        # print(B_Location)
    else:
        if(args[0] == 'DW，ZX0/2/FZX2/0'):
            if(sg_location+450 < ponder_location):
                print('里程错误！正确的里程为：' + str(sg_location+450))
                suggest = text + str(sg_location+450)
                verify(row, 3, B_Location, suggest)
        elif(args[0] == 'DW'):
            if(sg_location-250 < ponder_location):
                print('里程错误！正确的里程为：' + str(sg_location-250))
                suggest = text + str(sg_location+450)
                verify(row, 3, B_Location, suggest)
        else:
            print('里程正确!')
        B_trueLocation = B_Location
    return B_trueLocation


# 验证名称
def verifyName(row, B_trueLocation, B_Name, use, index):
    # B_trueLocation = judgeLocation()
    flag = True
    initials = B_Name[0]
    # 判断首字母是否为 'B'
    if (initials == 'B'):
        print('首字母正确')
    else:
        flag = False
        print('首字母错误')
    distance = int(str(getLocNum(B_trueLocation))[0:4])
    if (distance % 2 == 0):
        trueDistance = distance + 1     # 得到正确的公里标
    else:
        trueDistance = distance
    if(use == 'JZ'):
        print('公里标正确')
    else:
        # 提取名称的公里标，ykksdadsad -> [y,k,k,s,d,a]
        Km_mark = int(B_Name[1:5])
        if(trueDistance == Km_mark):
            print('公里标正确！')
        else:
            flag = False
            print('公里标错误！')
    # 284018 + 30 得到的是有源应答器的里程
    if(use != 'DW'):
        num = B_Name.split('-')[1]
        if(num == str(index+1)):
            print('组内编号正确!')
        else:
            print('组内编号错误!')
            flag = False
    if(use != 'DW'):
        B_trueName = 'B'+str(trueDistance)+'-'+str(index+1)
    else:
        B_trueName = 'B'+str(trueDistance)
    suggest = text + B_trueName
    if(flag):
        print('名称正确')
    else:
        verify(row, 1, B_Name, suggest)


# 验证编号,先不忙验证，有点绕
def verifyNum(row, B_Num, location, use, index):
    value = B_Num.split('-')    # 存放切割后的数组
    num_DQu = value[0]  # 编号的大区号
    num_FQu = value[1]  # 编号的分区号
    num_CZ = value[2]    # 编号的车站号
    num_cellNum = value[3]  # 单元编号
    if(use != 'DW'):
        num_Num = value[4]     # 应答器组内编号

    locaNum = getLocNum(location)
    T_locaNum = getLocNum(S_Through)

    if(locaNum < T_locaNum):
        Sta_CZ = Sta_CZ1
    else:
        Sta_CZ = Sta_CZ2

    # 未来的思路
    # 现在我们的判断，对于单元编号和组内编号使一个写死的值
    # 在将来，会结合数组来实现
    if(num_DQu == Sta_DQu and num_FQu == Sta_FQu
       and num_CZ == Sta_CZ):
        #  andnum_cellNum == '00'+str(indexNum)
        if(use != 'DW'):
            if(num_Num == str(index+1)):
                print('应答器编号正确!')
    else:
        if (num_DQu != Sta_DQu):
            print('大区编号错误!')
        elif (num_FQu != Sta_FQu):
            print('分区编号错误!')
        elif (num_CZ != Sta_CZ):
            print('车站号错误!')
        # elif (num_cellNum != '001'):
        #     print('单元号错误!')
        elif (num_Num != str(index+1)):
            print('组内编号错误!')
        if(use != 'DW'):
            B_trueNum = Sta_DQu+'-'+Sta_FQu+'-' + \
                Sta_CZ+'-'+num_cellNum+'-'+str(index+1)
        else:
            B_trueNum = Sta_DQu+'-'+Sta_FQu+'-'+Sta_CZ+'-'+num_cellNum
        suggest = text + B_trueNum
        verify(row, 2, B_Num, suggest)


# 验证设备类型
def verifyType(row, use, ponderType, index):
    # 私有方法，通过我们给定的正确类型来验证应答器是否正确
    # print('设备类型', end='')
    def _verifyType(row, ponderType, trueTpye):
        if(ponderType == trueTpye):
            print('设备类型正确!')
        else:
            suggest = text + trueTpye
            verify(row, 4, ponderType, suggest)

    # 如果应答器类型是 'CZ-C01' 或者 'CZ-C02' 那么这个应答器组第一个应答器就是有源应答器
    if(use == 'CZ-C01' or use == 'CZ-C02'):
        if(index == 0):
            _verifyType(row, ponderType, '有源')
        else:
            _verifyType(row, ponderType, '无源')
    elif(use == 'JZ'):    # 同理，这个应答器组就是最后一个是有源应答器
        if(index == 2):
            _verifyType(row, ponderType, '有源')
        else:
            _verifyType(row, ponderType, '无源')
    else:
        _verifyType(row, ponderType, '无源')


# 验证用途
def verifyUse(row, use, trueUse):
    if(use == trueUse):
        print('用途正确')
    else:
        suggest = text + trueUse
        verify(row, 5, use, suggest)


# 这里用到了我们之前导入的 copy
workbook = copy(ponder_wb)
worksheet = workbook.get_sheet(0)
# 设置样式
style = xlwt.easyxf('font:name 宋体, color-index red')


def verify(row, col, value, suggest):
    # 根据返回值判断是否需要标红
    worksheet.write(row, col, value, style)
    worksheet.write(row, col+7, suggest, style)
    worksheet.col(col+7).width = 256 * 25


# 开始验证
# 定义一个计数器,因为在 excel 中是从第二行开始的，所以定义为2
index = 2   # 数据表位置
flag = 0    # 应答器标志
reference = S_Out
use = getUse()
_reference = 0

P_use = use[0]
print(index)

while (1 < index < len(ponderSet)-3):
    end = index + use[1][flag]
    ponders = ponderSet[index:end]

    if(P_use[flag] == 'DW，ZX0/2/FZX2/0'):
        reference = S_Through
    elif(P_use[flag] == 'JZ' or P_use[flag] == 'DW'):
        reference = S_In

    for i in range(len(ponders)):
        Pname = ponders[i][1]   # 待验证的应答器名称
        Pnum = ponders[i][2]    # 待验证的应答器编号
        Plocation = ponders[i][3]   # 待验证的应答器里程
        Ptype = ponders[i][4]   # 待验证的应答器类型
        Puse = ponders[i][5]    # 待验证的应答器用途(可省略)

        # 判断数据是否缺失，如果缺失，则执行下一个应答器
        print(reference)
        if(isMissing(P_use[flag], Plocation, reference, i)):
            index += 1
            continue

        # 验证里程并得到正确的里程
        trueLocation = verifyLocation(
            index, reference, Plocation, *[P_use[flag], i])
        if(i == 0):
            # 将第一个应答器的正确位置作为应答器组的位置
            _reference = trueLocation

        # 验证名称
        verifyName(index, trueLocation, Pname, P_use[flag], i)
        
        # 验证编号
        verifyNum(index, Pnum, trueLocation, P_use[flag], i)  # 暂时不验证

        # 验证类型
        verifyType(index, P_use[flag], Ptype, i)

        # 验证用途(可省略)
        verifyUse(index, Puse, P_use[flag])
        
        reference = trueLocation
        index += 1
    flag += 1
    reference = _reference

# for i in range(2:len(ponderSet)-3):
#     Pname = ponders[i][1]
#     Pnum = ponders[i][2]
#     Plocation = ponders[i][3]
#     Ptype = ponders[i][4]
#     Puse = ponders[i][5]
#         trueLocation = verifyLocation(
#             counter, reference, Plocation, *[_use[i], j])
#         if(j == 0):
#             _reference = trueLocation
#         verifyName(counter, trueLocation, Pname, _use[i], j)
#         verifyNum(counter, Pnum, trueLocation, _use[i], j)  # 暂时不验证
#         verifyType(counter, _use[i], Ptype, j)
#         verifyUse(counter, Puse, _use[i])
#         reference = trueLocation
#         counter += 1
#         # 验证规则
#     reference = _reference

workbook.save('verified.xls')
