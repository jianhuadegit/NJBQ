# author:zhujianhua
# contact: 99376627@qq.com
# datetime:2020/12/17 9:56
# software: PyCharm
import random
import time

import win32con
import win32gui
from selenium.webdriver import ActionChains

from Data.Xpath import can_xpath
from Data.Xpath.PublicMethod import wd, getdatafromExcel, addrole


def releasecan():

    #鼠标悬浮在发布所能上
    release = wd.find_element_by_css_selector(can_xpath.canicon)
    ActionChains(wd).move_to_element(release).perform()
    #提供服务
    provide_severce_class = wd.find_elements_by_class_name(can_xpath.canclass)
    provide_severce_class[0].click()
    hands = wd.window_handles
    hand = list(hands)
    wd.switch_to_window(hand[len(hand) - 1])
    wd.find_element_by_css_selector(can_xpath.icanicon).click()
    # time.sleep(3)
#   #选择电子信息
    time.sleep(1)
    information = wd.find_element_by_css_selector(can_xpath.information)
    ActionChains(wd).move_to_element(information).perform()
    ActionChains(wd).double_click(information).perform()
    #输入标题
    wd.find_element_by_css_selector(can_xpath.logo).send_keys(getdatafromExcel()[0])
    #选择标签
    wd.find_element_by_css_selector(can_xpath.select).click()
    time.sleep(1)
    lll = wd.find_element_by_css_selector(can_xpath.development)
    time.sleep(1.5)
    lll.click()
    #输入期望金额
    wd.find_element_by_css_selector(can_xpath.Expectedamount).send_keys(getdamount())
    #选择元/项
    wd.find_element_by_css_selector(can_xpath.yuan).click()
    time.sleep(1)
    a=wd.find_elements_by_class_name(can_xpath.meixiang)
    a[3].click()

    def getimg():
        a = random.randint(1, 16)
        return a

    imgfilepath = 'E:\\NJBQ\Img\\' + str(getimg()) + '.jpg'
    #上传图片
    wd.find_element_by_css_selector(can_xpath.updateimg).click()
    time.sleep(1)
    upload(filePath=imgfilepath)
    save()
    #输入需求详情
    wd.find_element_by_css_selector(can_xpath.reqdetiles).send_keys(addrole()+getdatafromExcel()[1])
    #点击已阅读
    wd.find_element_by_css_selector(can_xpath.red).click()
    #发布
    time.sleep(4)
    wd.find_element_by_css_selector(can_xpath.releasebutton).click()
    #点击确定
    wd.find_element_by_css_selector(can_xpath.yes).click()





def clickreleasecan():
    #点击卖家中心
    wd.find_element_by_css_selector(can_xpath.SellerCenter).click()


def upload(filePath, browser_type="chrome"):
    '''
    通过pywin32模块实现文件上传的操作
    :param filePath: 文件的绝对路径
    :param browser_type: 浏览器类型（默认值为chrome）
    :return:
    '''
    if browser_type.lower() == "chrome":
        title = "打开"
    elif browser_type.lower() == "firefox":
        title = "文件上传"
    elif browser_type.lower() == "ie":
        title = "选择要加载的文件"
    else:
        title = ""  # 这里根据其它不同浏览器类型来修改

    # 找元素
    # 一级窗口"#32770","打开"
    dialog = win32gui.FindWindow("#32770", title)
    # 向下传递
    print(title)
    ComboBoxEx32 = win32gui.FindWindowEx(dialog, 0, "ComboBoxEx32", None)  # 二级
    comboBox = win32gui.FindWindowEx(ComboBoxEx32, 0, "ComboBox", None)   # 三级
    # 编辑按钮
    edit = win32gui.FindWindowEx(comboBox, 0, 'Edit', None)  # 四级
    # 打开按钮
    print(edit)
    button = win32gui.FindWindowEx(dialog, 0, 'Button', "打开(&O)")  # 二级

    # 输入文件的绝对路径，点击“打开”按钮
    win32gui.SendMessage(edit, win32con.WM_SETTEXT, None, filePath)  # 发送文件路径
    win32gui.SendMessage(dialog, win32con.WM_COMMAND, 1, button)  # 点击打开按钮
    time.sleep(1)

def save():
    wd.find_element_by_css_selector(can_xpath.savebutton).click()
#获取期望金额
def getdamount():
    i=random.randint(20,400)
    damount=100*i
    return damount



