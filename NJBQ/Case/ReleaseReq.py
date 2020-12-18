# author:zhujianhua
# contact: 99376627@qq.com
# datetime:2020/12/16 14:40
# software: PyCharm

#发布所需
import random
from Data import weburl, user
from Data.Xpath.PublicMethod import OPenbrowser, loginTianfkjy, releaseReq


def DataFromExcelToPage():
    OPenbrowser(weburl.tianfu)
    loginTianfkjy(username=user.user1[1],pwd=user.user1[2])
    for i in range(50):
        releaseReq()
        print(i+1,'条所需已经发布！！！********************************************************************************')








if __name__ == '__main__':
    DataFromExcelToPage()