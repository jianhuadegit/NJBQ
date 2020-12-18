# author:zhujianhua
# contact: 99376627@qq.com
# datetime:2020/12/16 15:48
# software: PyCharm
from Data import user, weburl
from Data.Xpath import can_xpath
from Data.Xpath.PublicMethod import getdatafromExcel, OPenbrowser, loginTianfkjy, getregion
from Public.CanMethod import releasecan, upload, save, clickreleasecan


def can():
    OPenbrowser(weburl.tianfu)
    loginTianfkjy(username=user.user1[1], pwd=user.user1[2])
    clickreleasecan()
    for i in range(50):

        releasecan()



if __name__ == '__main__':
    can()