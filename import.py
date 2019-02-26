import xlrd
import xlwt
from xlutils.copy import copy
# path
transponder='./data/transponder.xls'    # 应答器
station='./data/station.xls'    # 车站位置信息
signal='./data/signalData.xls'  # 信号机信息

ponder_wb=xlrd.open_workbook(filename=transponder, formatting_info=True)
station_wb=xlrd.open_workbook(filename=station, formatting_info=True)
signal_wb=xlrd.open_workbook(filename=signal, formatting_info=True)
ponder_ws=ponder_wb.sheet_by_name('下行')
station_ws=station_wb.sheet_by_name('Sheet1')
signal_ws=signal_wb.sheet_by_name('下行正向')

# 定义数组
ponderSet=[]    # 应答器信息
staSet=[]       # 车站位置信息
signalSet=[]    # 信号机信息
headerSet1 = []
headerSet2 = []

# range传要循环的数组
for c in range(ponder_ws.ncols):
    # append是数组的一个方法
    headerSet1.append(ponder_ws.cell(0,c).value)
    headerSet2.append(ponder_ws.cell(1,c).value)
    ponderSet.append(ponder_ws.cell(2,c).value)

for s in range(station_ws.ncols):
    staSet.append(station_ws.cell(2,s).value)

for r in range(signal_ws.nrows):
    if(signal_ws.cell(r,2).value=='SHF'):
        for s in range(signal_ws.ncols):
            signalSet.append(signal_ws.cell(r,s).value)
from pprint import pprint

# 定义其他要用到的变量
S_Location = signalSet[3]   # 信号机里程
Sta_DQu = str(int(staSet[2]))          # 大区号 float 3.0 int -> 3
Sta_FQu = str(int(staSet[3]))         # 分区号
Sta_CZ = str(int(staSet[4]))          # 车站号

# 定义应答器父类
class Transponder(object):
    # 初始化四个属性
    def __init__(self, name, num, location, type):
        self.B_Name = name
        self.B_Num = num
        self.B_Location = location
        self.B_Type = type

    def verifyName(self, B_trueLocation):
        # B_trueLocation = judgeLocation()
        flag = True
        initials = self.B_Name[0]
        # 判断首字母是否为 'B'
        if (initials == 'B'):
            print('首字母正确')
            print(B_trueLocation)
        else:
            print('首字母错误')
            flag = False
        shuzi = B_trueLocation.split('K')[1]
        _distance = shuzi.replace('+', '')   # 将 '+' 消除
        distance = int(_distance[0:4])      # 取前四位
        if (distance % 2 == 0):
            trueDistance = distance + 1     # 得到正确的公里标
        else:
            trueDistance = distance
        Km_mark = int(self.B_Name[1:5])                 # 提取名称的公里标，ykksdadsad -> [y,k,k,s,d,a]
        if(trueDistance == Km_mark):
            print('公里标正确！')
        else:
            print('公里标错误！')
            flag = False
            print('正确的公里标为:' + str(trueDistance))
        # 284018 + 30 得到的是有源应答器的里程
        num = self.B_Name.split('-')[1]
        if(num == '1'):
            print('组内编号正确!')
        else:
            print('组内编号错误!')
            flag = False
        self.B_trueName = 'B'+str(trueDistance)+'-1'
        print(self.B_trueName)
        print(flag)
        return flag

    def verifyNum(self):
        value = self.B_Num.split('-')    # 存放切割后的数组
        num_DQu = value[0]  # 编号的大区号
        num_FQu = value[1]  # 编号的分区号
        num_CZ = value[2]    # 编号的车站号
        num_cellNum = value[3]  # 单元编号
        num_Num =  value[4]     # 应答器组内编号

        # 未来的思路
        # 现在我们的判断，对于单元编号和组内编号使一个写死的值
        # 在将来，会结合数组来实现
        if(num_DQu == Sta_DQu and num_FQu == Sta_FQu and num_CZ == Sta_CZ and
            num_cellNum == '001' and num_Num == '1'):
            self.B_trueNum = self.B_Num
            print('应答器编号正确!')
            return True
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
            self.B_trueNum = Sta_DQu+'-'+Sta_FQu+'-'+Sta_CZ+'-001'+'-1'
            print(self.B_trueNum)
            return False
        pass

# 有源应答器类继承于应答器类
# 重写了里程验证和类型验证
class Active_Transponder(Transponder):
    # 判断里程是否正确
    def verifyLocation(self):
        sg_location=getLocation(S_Location)
        ponder_location=getLocation(self.B_Location)
        true_location=sg_location+30
        if true_location>ponder_location:
            if true_location<100 :
                true_location='0'+str(true_location)
            self.B_trueLocation = self.B_Location.split('+')[0]+'+'+str(true_location)
            print('里程错误！正确的里程为:')
            print(self.B_trueLocation)
            return False
            # print(B_Location)
        else:
            print('里程正确!')
            self.B_trueLocation = self.B_Location
            return True
    def verifyType(self):
        if(self.B_Type == '有源'):
            print('应答器类型正确！')
            return True
        else:
            print('应答器类型错误!')
            return False
    pass

# 创建应答器实例
ponder = Active_Transponder(ponderSet[1],ponderSet[2],ponderSet[3],ponderSet[4])

# 定义四个属性
# ponder.B_Name = ponderSet[10]
# ponder.B_Num = ponderSet[2]
# ponder.B_Location = ponderSet[3]
# ponder.B_Type = ponderSet[4]

# 定义正确的值
# B_trueLocation = ''

# 获取+号后面里程信息，并返回
def getLocation(location):
    sg_location=location.split('+')[1]
    sg_location=int(sg_location)
    return sg_location

workbook = copy(ponder_wb)
worshell = workbook.get_sheet(0)
print(worshell)

ponder.verifyLocation()
ponder.verifyName(ponder.B_trueLocation)
ponder.verifyType()
ponder.verifyNum()