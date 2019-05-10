import xlrd
import xlwt
from xlutils.copy import copy
import os
import time
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph,SimpleDocTemplate
from reportlab.lib import  colors

pdfmetrics.registerFont(TTFont('song', './font/SimSun.ttf'))  # 导入中文字体


# 设置PDF样式
class Graphs:
    def __init__(self):
        pass

    @staticmethod
    def setHead():
        Style=getSampleStyleSheet()
        bt = Style['Normal']     #字体的样式
        bt.fontName='song'    #使用的字体
        bt.leading = 30             #该属性是设置行距
        bt.fontSize=24
        bt.alignment=1
        heading = '毕业设计——列控工程数据验证'
        head = Paragraph(heading, bt)
        return head
    @staticmethod
    def setVersion():
        Style=getSampleStyleSheet()
        bt = Style['Normal']     #字体的样式
        bt.fontName='song'    #使用的字体
        bt.wordWrap = 'CJK'    #该属性支持自动换行，'CJK'是中文模式换行，用于英文中会截断单词造成阅读困难，可改为'Normal'
        bt.leading = 30             #该属性是设置行距
        bt.fontSize=12
        bt.alignment=1
        date=str(time.localtime().tm_year)+'年'+str(time.localtime().tm_mon)+'月'+str(time.localtime().tm_mday)+'日'+str(time.localtime().tm_hour)+'时'+str(time.localtime().tm_min)+'分'
        version = '版本号:v-0.0.1       ' + date + '提交        作者：陈晨'
        versionText = Paragraph(version, bt)
        return versionText
    @staticmethod
    def setTitle():
        Style=getSampleStyleSheet()
        bt = Style['Normal']     #字体的样式
        bt.fontName='song'    #使用的字体
        bt.wordWrap = 'CJK'    #该属性支持自动换行，'CJK'是中文模式换行，用于英文中会截断单词造成阅读困难，可改为'Normal'
        bt.leading = 30             #该属性是设置行距
        bt.fontSize=18
        bt.alignment=1
        title = Paragraph('审核详情', bt)
        return title
    @staticmethod
    def setNormalText():
        Style=getSampleStyleSheet()
        ct = Style['Normal']     #字体的样式
        ct.fontName='song'    #使用的字体
        ct.wordWrap = 'CJK'    #该属性支持自动换行，'CJK'是中文模式换行，用于英文中会截断单词造成阅读困难，可改为'Normal'
        ct.leading = 30             #该属性是设置行距
        ct.fontSize=12
        ct.alignment=0
        ct.firstLineIndent = 32  #该属性支持第一行开头空格
        ct.textColor = colors.black
        return ct
    @staticmethod
    def setErrorText():
        Style=getSampleStyleSheet()
        ct = Style['Normal']     #字体的样式
        ct.fontName='song'    #使用的字体
        ct.wordWrap = 'CJK'    #该属性支持自动换行，'CJK'是中文模式换行，用于英文中会截断单词造成阅读困难，可改为'Normal'
        ct.leading = 30             #该属性是设置行距
        ct.fontSize=12
        ct.alignment=0
        ct.firstLineIndent = 32  #该属性支持第一行开头空格
        ct.textColor = colors.red
        return ct

# 定义全局变量
reportSet = []  # 所有的错误信息，存放到此数组
isTrueValue = True
strReport = ''
R_Reference = ''                   # 定义参照点
text = '建议修改为：'

# 导出PDF
# report->保存验证信息的数组，total->总共验证的数据条数，errNum->错误数据
def report(report, total, errNum):
    _report = list()
    _report.append(Graphs.setHead())
    _report.append(Graphs.setVersion())
    _report.append(Graphs.setTitle())
    line1 = '   共验证'+str(total)+'条数据，'+str(errNum)+'个数据异常'
    _report.append(Paragraph(line1, Graphs.setNormalText()))
    for i in range(len(report)):
        if(report[i][1] == 'p-normal'):
            ct = Graphs.setNormalText()
        else:
            ct = Graphs.setErrorText()
        _report.append(Paragraph(report[i][0],ct))
    pdf=SimpleDocTemplate('verifyReport.pdf', pagesize=letter)
    pdf.build(_report)


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

# 全局变量
S_Out = signalSet[0][3]   # 信号机位置
S_Through = signalSet[1][3]  # 通过信号机位置
S_In = signalSet[2][3]      # 进站信号机位置
Sta_DQu = str(int(staSet[2][2]))          # 大区号 float 3.0 int -> 3
Sta_FQu = str(int(staSet[2][3]))         # 分区号
Sta_CZ1 = str(int(staSet[2][4]))          # 颜家垄
Sta_CZ2 = str(int(staSet[3][4]))          # 长塘线路所

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
    mapDistance = getDistance(use, location, reference, index)
    print(mapDistance['distance'])
    if(-50 < mapDistance['distance'] < 150):
        print('数据未缺失')
        return False
    else:
        print('数据缺失')
        return True


# 获取距离信息
def getDistance(use, location, reference, index):
    # 因为在验证执行应答器时使用了两个参照点，所以使用了一个数组来存放
    location = getLocNum(location)
    print(reference, location)
    if(use == 'CZ-C01' and index == 0):
        verifyLocation = reference[0] + 30
    elif(use == 'DW,YG0/2' and index == 0):
        verifyLocation = reference[0] - 223
    elif(use == 'DW,ZX0/2/FZX2/0' and index == 0):
        if(reference[0] + 450 >= reference[1] + 30):
            verifyLocation = reference[0] + 450
        elif(reference[1] - 30 < reference[0] + 450 < reference[1] + 30):
            verifyLocation = reference[1] + 30
        else:
            verifyLocation = reference[1] - 30
    elif(use == 'DW,FYG2/0' and index == 0):
        verifyLocation = reference[0] + 223
    elif(use == 'JZ' and index == 0):
        verifyLocation = reference[0] - 40
    elif(use == 'DW'):
        verifyLocation = reference[0] - 250
    elif(index == 0):
        verifyLocation = reference[0] + 200
    else:
        verifyLocation = reference[0] + 5
    
    distance = verifyLocation - location
    return {
        "distance": distance,
        "location": verifyLocation
    }


# 判断里程是否正确
def verifyLocation(row, reference, B_Location, use, index):
    strReport=''
    isTrueValue=True
    # args[0] == use,args[1] == index
    if(use == 'CZ-C01' and index == 0):
        _strReport = ',【里程异常】->【该应答器应距离出站口30±0.5米】'
    elif(use == 'JZ' and index == 0):
        _strReport = '，【里程异常】->【该应答器应距离进站信号机至少40±0.5米】'
    elif(use == 'DW'):
        _strReport = '，【里程异常】->【该应答器应距离进站信号机至少250米】'
    elif(use == 'DW,YG0/2' or use == 'DW,FYG2/0' and index == 0):
         _strReport = '，【里程异常】->【该应答器应距离执行应答器组至少223米】'
    elif(use == 'DW,ZX0/2/FZX2/0' and index == 0):
        _strReport = '，【里程异常】->【该应答器应距离进站信号机至少30米而且需要距离CZ-C02应答器组至少450米】'
    elif(index == 0):
        _strReport = '，【里程异常】->【应答器组之间的距离应大于200米】'
    else:
        _strReport = '，【里程异常】->【应答器间距应为5±0.5米】'

    mapDistance = getDistance(use, B_Location, reference, index)
    if(-0.5 < mapDistance['distance'] < 0.5):
        print('里程正确！')
    else:
        strLoction = 'JHK' + str(mapDistance['location'])[0:3] + '+' + str(mapDistance['location'])[3:6]
        suggest = text + strLoction
        print('里程错误！正确的里程为:'+strLoction)
        verify(row, 3, strLoction, suggest)
        isTrueValue = False
        strReport = _strReport
    # if (true_location > ponder_location):
    #     B_trueLocation = 'JHK'+str(true_location)[0:3]+'+'+str(true_location)[3:6]
    #     suggest = text + B_trueLocation
    #     print('里程错误！正确的里程为:'+B_trueLocation)
    #     verify(row, 3, B_Location, suggest)
    #     isTrueValue = False
    #     strReport = _strReport
    #     # print(B_Location)
    # else:
    #     if(args[0] == 'DW，ZX0/2/FZX2/0'):
    #         if(sg_location+450 < ponder_location):
    #             print('里程错误！正确的里程为：' + str(sg_location+450))
    #             suggest = text + str(sg_location+450)
    #             verify(row, 3, B_Location, suggest)
    #     elif(args[0] == 'DW'):
    #         if(sg_location-250 < ponder_location):
    #             print('里程错误！正确的里程为：' + str(sg_location-250))
    #             suggest = text + str(sg_location-250)
    #             verify(row, 3, B_Location, suggest)
    #     else:
    #         print('里程正确!')
    #     B_trueLocation = B_Location
    return {
        "location":mapDistance['location'],
        "report":strReport,
        "isTrue": isTrueValue
        }


# 验证名称
def verifyName(row, B_trueLocation, B_Name, use, index):
    strReport=''
    isTrueValue=True
    reason= ''
    flag = True
    initials = B_Name[0]
    # 判断首字母是否为 'B'
    if (initials == 'B'):
        print('首字母正确')
    else:
        flag = False
        reason =reason + '名称应以B字母开头；'
        print('首字母错误')
    distance = int(str(B_trueLocation)[0:4])
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
            reason = reason + '公里标错误，正确的公里标应为里程的数字位前四位或信号机名称，上奇下偶；'
            print('公里标错误！')
    # 284018 + 30 得到的是有源应答器的里程
    if(use != 'DW'):
        num = B_Name.split('-')[1]
        if(num == str(index+1)):
            print('组内编号正确!')
        else:
            reason=reason + '组内编号错误，应为应答器排列顺序。'
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
        strReport=strReport+'，【名称错误】->【'+reason+'】'
        isTrueValue = False
    return {
        "report": strReport,
        "isTrue": isTrueValue
        }


# 验证编号,先不忙验证，有点绕
def verifyNum(row, B_Num, location, use, index):
    strReport=''
    isTrueValue=True
    reason=''
    value = B_Num.split('-')    # 存放切割后的数组
    num_DQu = value[0]  # 编号的大区号
    num_FQu = value[1]  # 编号的分区号
    num_CZ = value[2]    # 编号的车站号
    num_cellNum = value[3]  # 单元编号
    if(use != 'DW'):
        num_Num = value[4]     # 应答器组内编号

    T_locaNum = getLocNum(S_Through)

    if(location < T_locaNum):
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
            reason=reason+'大区编号应与数据表中对应；'
        elif (num_FQu != Sta_FQu):
            print('分区编号错误!')
            reason=reason+'分区编号应与数据表中对应；'
        elif (num_CZ != Sta_CZ):
            print('车站号错误!')
            reason=reason+'车站编号应与车站表车站号对应；'
        # elif (num_cellNum != '001'):
        #     print('单元号错误!')
        elif (num_Num != str(index+1)):
            print('组内编号错误!')
            reason=reason+'组内编号应为顺序编号。'
        if(use != 'DW'):
            B_trueNum = Sta_DQu+'-'+Sta_FQu+'-' + \
                Sta_CZ+'-'+num_cellNum+'-'+str(index+1)
        else:
            B_trueNum = Sta_DQu+'-'+Sta_FQu+'-'+Sta_CZ+'-'+num_cellNum
        suggest = text + B_trueNum
        verify(row, 2, B_Num, suggest)
        strReport=strReport+'，【编号错误】->【'+reason+'】'
        isTrueValue = False

    return {
        "report":strReport,
        "isTrue": isTrueValue
        }


# 验证设备类型
def verifyType(row, use, ponderType, index):
    # 私有方法，通过我们给定的正确类型来验证应答器是否正确
    # print('设备类型', end='')
    strReport=''
    isTrueValue=True
    def _verifyType(row, ponderType, trueTpye):
        global strReport,isTrueValue
        if(ponderType == trueTpye):
            print('设备类型正确!')
        else:
            suggest = text + trueTpye
            verify(row, 4, ponderType, suggest)
            strReport = strReport+'，【应答器类型错误】->【该应答器设备类型应为：'+trueTpye+'】'
            isTrueValue = False

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

    return {
        "report":strReport,
        "isTrue": isTrueValue
        }


# 验证用途
def verifyUse(row, use, trueUse):
    strReport=''
    isTrueValue=True
    if(use == trueUse):
        print('用途正确')
    else:
        suggest = text + trueUse
        verify(row, 5, use, suggest)
        strReport = strReport+'，【应答器用途错误】->【该应答器设备用途应为：'+trueUse+'】'
        isTrueValue = False
    return {
        "report":strReport,
        "isTrue": isTrueValue
        }


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


def verifyData(ponders, index, reference, P_use, flag, errNum):
    strReport = ''
    _reference = 0
    for i in range(len(ponders)):
        Pnum = ponders[i][2]    # 待验证的应答器编号
        Plocation = ponders[i][3]   # 待验证的应答器里程
        Pname = ponders[i][1]   # 待验证的应答器名称
        Ptype = ponders[i][4]   # 待验证的应答器类型
        Puse = ponders[i][5]    # 待验证的应答器用途(可省略)

        referenceSet = []
        referenceSet.append(reference)
        print(P_use[flag])
        if(P_use[flag] == 'DW,ZX0/2/FZX2/0'):
            referenceSet.append(getLocNum(S_Through))
            print(referenceSet)
        elif(P_use[flag] == 'JZ' and i == 0 or P_use[flag] == 'DW'):
            referenceSet = []
            referenceSet.append(getLocNum(S_In))

        # 判断数据是否缺失，如果缺失，则执行下一个应答器
        if(isMissing(P_use[flag], Plocation, referenceSet, i)):
            strReport='     第' + str(index-1) + '行: '+Pname+'数据缺失！'
            reportSet.append([strReport,'p-error'])
            errNum += 1
            index += 1
            continue
        # 验证里程并得到正确的里程
        mapLocation = verifyLocation(
            index, referenceSet, Plocation, P_use[flag], i)
        trueLocation = mapLocation['location']
        if(i == 0):
            # 将第一个应答器的正确位置作为应答器组的位置
            _reference = trueLocation
        # 验证名称
        mapName=verifyName(index, trueLocation, Pname, P_use[flag], i)
        # 验证编号
        mapNum=verifyNum(index, Pnum, trueLocation, P_use[flag], i)
        # 验证类型
        mapType=verifyType(index, P_use[flag], Ptype, i)
        # 验证用途(可省略)
        mapUse=verifyUse(index, Puse, P_use[flag])
        # 获取所有方法执行后的report和数据是否正确
        locationReport = mapLocation['report']
        nameReport = mapName['report']
        numReport = mapNum['report']
        typeReport = mapType['report']
        useReport = mapUse['report']
        locatioIsTrue = mapLocation['isTrue']
        nameIsTrue = mapName['isTrue']
        numIsTrue = mapNum['isTrue']
        typeIsTrue = mapType['isTrue']
        useIsTrue = mapUse['isTrue']
        if (locatioIsTrue and nameIsTrue and numIsTrue and typeIsTrue and useIsTrue) :
            strReport = '     第'+str(index-1)+'行： '+Pname+'，验证正确。'
            reportSet.append([strReport,'p-normal'])
        else:
            strReport = '     第'+str(index-1)+'行： '+Pname+locationReport+nameReport+numReport+typeReport+useReport
            reportSet.append([strReport,'p-error'])
            errNum += 1
        reference = trueLocation
        index += 1

    return {
        "reference": _reference,
        "index": index,
        "errNum": errNum
    }

# 开始验证
def main():
    # 定义其他要用到的变量

    # 定义一个计数器,因为在 excel 中是从第二行开始的，所以定义为2
    index = 2   # 数据表位置
    flag = 0    # 应答器标志
    reference = getLocNum(S_Out)
    use = getUse()
    _reference = 0
    P_use = use[0]
    total = len(ponderSet)
    errNum = 0  # 错误的数据条数

    while (1 < index < total-3):
        isTrueValue = True
        end = index + use[1][flag]
        ponders = ponderSet[index:end]
        if(P_use[flag] == 'DW,YG0/2'):  # 执行快照，将这个应答器信息保存下来
            _flag = flag
            _ponders = ponders
            _index = index
            index += 2
            flag += 1
            continue
        verifyedData = verifyData(ponders, index, reference, P_use, flag, errNum)

        if(P_use[flag] == 'DW,ZX0/2/FZX2/0'):
            verifyData(_ponders, _index, verifyedData['reference'], P_use, _flag, verifyedData['errNum'])

        flag += 1
        reference = verifyedData["reference"]
        index = verifyedData["index"]
        errNum = verifyedData["errNum"]

    # 保存excel文件
    workbook.save('verified.xls')

    # 导出报表信息
    report(reportSet, total-5, errNum)


if __name__ == "__main__":
    main()