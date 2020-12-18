# author:zhujianhua
# contact: 996376627@qq.com
# datetime:2020/12/12
# software: PyCharm
from openpyxl import load_workbook

#从汇新云获取所需的数据


from Data import weburl
from Public.GetDataFromHuixinyunToExcel.GetDataFromHuixinyunToExcel import Getalldata, Gotonextpage, GetData, save, \
    Gotonextpage1
from Data.Xpath.PublicMethod import OPenbrowser
import datetime

def startGetdata():
    begin=datetime.datetime.now()
    OPenbrowser(weburl.huixin)
    GetData()
#控制从多少页开始取数据
    for num1 in range(1,140):
        if num1 <=94:
            Gotonextpage()
        else:
            #控制每页取多少条
            for num in range(0, 15):

                if num<=14:
                    Getalldata(num,num1)
                else:
                    #跳转到下一页
                    Gotonextpage1()
            Gotonextpage1()
    save()

    last=datetime.datetime.now()
    print('获取数据用时:',last-begin,)


if __name__ == '__main__':
    startGetdata()