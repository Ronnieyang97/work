import xlrd
import xlwt
from xlutils.copy import copy
import datetime

namedata = [["Ruby Lu", "13701619050"], ["John Hore", "15201723478"], ["Yan Zhuang", "18301903536"],
            ["John Hore 2", "13816857312"], ["For UK marketing", "15802198903"], ["Qi Li", "15000905387"],
            ["WIFI in car", "14782702489"], ["WIFI in car", "14782702495"], ["Kevin Kang", "13482415729"],
            ["Matthew Crabbe", "15821106759"], ["Elsa Wang", "13482417259"], ["Elsa Wang", "13818321319"],
            ["Benson Chen", "17802102349"], ["Benson Chen", "13482437049"], ["Amanda Yang", "18221051728"],
            ["Qi Li", "13817455491"], ["Teresa Wang", "15900865357"], ["Teresa Wang", "13817626549"],
            ["Sandy Chang", "13918921491"], ["Estelle Xu", "1064853917025"], ["Aileen Chen", "1064853865995"],
            ["Jenny Zhong", "1064878996586"], ["Lee Ann Fagan", "1064878996587"], ["Shane Wang", "1064853865996"],
            ["Vincent Qin", "1064722349718"], ["Lee Ann Fagan", "15021706843"], ["Jenny Zhong", "15021730284"],
            ["Della He", "15821571067"], ["Shane Wang", "15021241924"], ["Matthew Nelson", "15821580674"],
            ["Jess Zhang", "1064722349723"], ["Jess Zhang", "13761066414"], ["Linna Zhang", "1064722349720"],
            ["Insight team", "1064722349719"], ["Linna Zhang", "18721069534"], ["Amanda Yang", "15021236491"],
            ["Maggie Huang", "15002146701"], ["Christiana Leng", "13916049174"], ["Alice Jiang", "13701923153"],
            ["Wei Wen", "13482426475"], ["Kenny Li", "13817341276"], ["Katrina Wu", "13482464801"],
            ["Sandy Chang", "13916445914"], ["Bora Kim", "15001798257"], ["Sierra Sheng", "13916134384"]]

data = xlrd.open_workbook("C:\\Users\\Ronnie\\Downloads\\telelist.xls")
table = data.sheets()[1]
rows = table.nrows
x = 0
temp = []

for i in range(rows):   #读取旧表信息
    if table.row(i)[0].value == "小计":
        if table.row(i)[3].value:
            temp.append([table.row(i)[2].value, temp1, table.row(i)[3].value])      #temp中数据依次为[话费，基础话费，电话]
        if i < rows-1:
            temp1 = table.row(i + 1)[2].value
            if float(temp1) < 10:
                temp1 = table.row(i + 2)[2].value

workbook = copy(data)
output = workbook.add_sheet("result")

year = datetime.datetime.now().year     #处理日期格式
month = datetime.datetime.now().month
if month < 10:
    date = str(year) + "-0" + str(month) + "-01"
else:
    date = str(year) + "-" + str(month) + "-01"

lenth = len(temp)       #写入新表
for i in range(lenth):      #写入新表格的数据为[电话，日期，基础话费，总话费]
    output.write(i, 0, temp[i][2])
    output.write(i, 1, date)
    output.write(i, 2, temp[i][1])
    output.write(i, 3, temp[i][0])



workbook.save("C:\\Users\\Ronnie\\Downloads\\telelist.xls")     #保存数据