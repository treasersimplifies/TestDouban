# -*- coding: utf-8 -*-
import requests, json, re
import time, datetime, os, sys
import getpass
from selenium import webdriver
# from selenium.webdriver.chrome.options import Options
from halo import Halo
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import random
from Excel import Excel

filePath = "./TestCases.xlsx"
fileSavePath = "./TestResults.xls"

WAIT_TIME = 1
# class TestCase(object):
#     def __init__(self, testItem, testSubItem, testStep):


class SysTest(object):
    def __init__(self, username, password):
        self.PASS = 0.0
        self.FAIL = 0.0
        self.username = username
        self.password = password
        # chrome_options = Options()
        # chrome_options.add_argument('--headless')
        # self.driver = webdriver.Chrome('./chromedriver', chrome_options=chrome_options)
        self.driver = self._set_driver()
        self.driver.set_window_size(1450,700)
        self.driver.set_window_position(0,190)
        # self.csdn = "https://i.csdn.net/#/uc/profile"
        self.douban = "https://accounts.douban.com/passport/login?source=movie"
        self.excel_wb = Excel(filePath, fileSavePath)
        # self.sess = requests.Session()
        
    def _set_driver(self):
        """Set driver according to the os system"""
        if sys.platform == "win32":
            phantomjs_path = "./phantomjs.exe"
        elif sys.platform == "darwin":
            phantomjs_path = "./phantomjs-mac"
        else:
            phantomjs_path = "./phantomjs-linux"
        return webdriver.Chrome() # webdriver.PhantomJS(phantomjs_path)

    def login(self):
        driver = self.driver
        driver.get(self.douban)
        
        time.sleep(WAIT_TIME)
        # douban Login
        driver.find_element_by_xpath("//*[@id=\"account\"]/div[2]/div[2]/div/div[1]/ul[1]/li[2]").click()
        driver.find_element_by_xpath("//*[@id=\"username\"]").send_keys(self.username)
        driver.find_element_by_xpath("//*[@id=\"password\"]").send_keys(self.password)
        driver.find_element_by_xpath("//*[@id=\"account\"]/div[2]/div[2]/div/div[2]/div[1]/div[4]/a").click()
        time.sleep(3*WAIT_TIME)

        self.cookies = driver.get_cookies()
        cookie = [item["name"] + "=" + item["value"] for item in self.cookies ]
        
        print("\n***********************************************************\n")
        print("cookie:",cookie)
        print("\n***********************************************************\n")

        self.cookiestr = '; '.join(item for item in cookie)
        # driver.close()
        return self.cookiestr

    def readCases(self, n0=1, n1=None):
        driver = self.driver
        excel_wb = self.excel_wb

        # 读取要测试的 大功能:
        data = excel_wb.readColN(0)[n0:n1]
        self.testItem = data
        # print("大功能:", self.testItem, "\n")

        # 读取要测试的 小功能:
        data = excel_wb.readColN(1)[n0:n1]
        self.testSubItem = data
        # print("小功能:", self.testSubItem, "\n")

        # 读取要测试的 小功能 的 测试步骤:
        data = excel_wb.readColN(2)[n0:n1]
        self.testStep = data
        # print("测试步骤:", self.testStep, "\n")

        # 读取 测试步骤 的 操作元素:
        self.opElement = excel_wb.readColN(3)[n0:n1]
        # print("操作元素:", self.opElement, "\n")
        # 读取 操作元素 的 操作类型:
        self.opType = excel_wb.readColN(4)[n0:n1]
        # print("操作类型:", self.opType, "\n")
        # 读取 操作元素 的 输入值:
        self.opInput = excel_wb.readColN(5)[n0:n1]
        # print("输入值:", self.opInput, "\n")

        # 读取 测试步骤 的 预期元素:
        self.expElement = excel_wb.readColN(6)[n0:n1]
        # print("预期元素:", self.expElement, "\n")
        # 读取 预期元素 的 输出值:
        self.expOutput = excel_wb.readColN(7)[n0:n1]
        # print("输出值:", self.expOutput, "\n")

        # 读取 测试步骤 的 说明:
        self.opInfo = excel_wb.readColN(8)[n0:n1]

    def test(self, n0=1, n1=None):
        driver = self.driver
        print("\n 本轮有", len(self.testItem), "个测试用例，开始测试。")
        for i in range(0, len(self.testItem)):
            if self.testItem[i] != '':
                print(self.testItem[i])
                time.sleep(WAIT_TIME)
            if self.testSubItem[i] != '':
                print("   |- ", self.testSubItem[i])
                time.sleep(WAIT_TIME)
            if self.testStep[i] != '':
                print("       |- ", self.testStep[i])
                time.sleep(WAIT_TIME)
            print("正在进行的测试用例: ", self.opInfo[i])

            if self.opType[i] == 'click':
                try:
                    element = driver.find_element_by_xpath(self.opElement[i])
                except Exception as e:
                    print("Network not good, begin to waitX1 ...")
                    time.sleep(2*WAIT_TIME)
                    try:
                        element = driver.find_element_by_xpath(self.opElement[i])
                    except Exception as e:
                        print("Network not good, begin to waitX2 ...")
                        time.sleep(4*WAIT_TIME)
                        element = driver.find_element_by_xpath(self.opElement[i])
                element.click()
            elif self.opType[i] == 'input':
                try:
                    element = driver.find_element_by_xpath(self.opElement[i])
                except Exception as e:
                    print("Network not good, begin to waitX1 ...")
                    time.sleep(2*WAIT_TIME)
                    try:
                        element = driver.find_element_by_xpath(self.opElement[i])
                    except Exception as e:
                        print("Network not good, begin to waitX2 ...")
                        time.sleep(4*WAIT_TIME)
                        element = driver.find_element_by_xpath(self.opElement[i])
                if self.opInput[i] == "Keys.ENTER":
                    # print("press ENTER")
                    element.send_keys(Keys.ENTER)    
                else:
                    element.send_keys(self.opInput[i])
            elif self.opType[i] == 'move_to_element':
                # ele = driver.find_element_by_xpath(self.opElement[i])
                try:
                    element = driver.find_element_by_xpath(self.opElement[i])
                except Exception as e:
                    print("Network not good, begin to waitX1 ...")
                    time.sleep(2*WAIT_TIME)
                    try:
                        element = driver.find_element_by_xpath(self.opElement[i])
                    except Exception as e:
                        print("Network not good, begin to waitX2 ...")
                        time.sleep(4*WAIT_TIME)
                        element = driver.find_element_by_xpath(self.opElement[i])
                ActionChains(driver).move_to_element(element).perform()
                print("move_to_element 等待 2*WAIT_TIME")
                time.sleep(WAIT_TIME)
            elif self.opType[i] == 'accept':
                driver.switch_to_alert().accept()
            elif self.opType[i] == 'switch_to':
                driver.switch_to.window(driver.window_handles[int(self.opElement[i])])
                time.sleep(WAIT_TIME)
            elif self.opType[i] == 'back':
                times = int(self.opElement[i])
                for j in range(0,times):
                    driver.back()
            elif self.opType[i] == 'scroll':
                times = int(self.opElement[i])
                for j in range(0,times):
                    driver.execute_script("window.scrollBy(0, 200);")
                    time.sleep(WAIT_TIME/2)
            time.sleep(WAIT_TIME)
            if self.expElement[i] != '':
                try: 
                    actElement = driver.find_element_by_xpath(self.expElement[i])
                    print("预期值：", str(self.expOutput[i]), "实际值：", actElement.text, "Bool:", self.expOutput[i]==actElement.text)
                    if str(self.expOutput[i])==actElement.text:
                        print("OK.")
                        self.excel_wb.writeResults(n0+i, 9, "PASS")
                        self.PASS += 1
                        # self.excel_wb.saveResult()
                    else:
                        print("FAIL.")
                        self.excel_wb.writeResults(n0+i, 9, "NOT")
                        self.FAIL += 1
                        # self.excel_wb.saveResult()
                except Exception as e:
                    print("FAIL.")
                    print("XXX Exception:", e)
                    self.excel_wb.writeResults(n0+i, 9, "NOT")
                    self.FAIL += 1
                    # self.excel_wb.saveResult()
            else:
                print("OK.")
                self.excel_wb.writeResults(n0+i, 9, "PASS")
                self.PASS += 1
                # self.excel_wb.saveResult()
        self.testItem = None
        self.testSubItem = None
        self.testStep = None
        self.opElement = None
        self.opType = None
        self.opInput = None
        self.expElement = None
        self.expOutput = None
        self.opInfo = None
        self.excel_wb.saveResult()

    # def writeResult(self):
        # excel_wb.writeResult(3,3,"3-3")
        # excel_wb.saveResult()

    def finish(self):
        time.sleep(2*WAIT_TIME)
        self.excel_wb.writeResults(0, 10, "测试通过率%")
        self.excel_wb.writeResults(1, 10, self.PASS*100/(self.FAIL+self.PASS))
        print("测试【失败】用例:", self.FAIL, "测试【成功】用例:", self.PASS)
        self.excel_wb.saveResult()
        self.driver.close()

def main():  #  username, password
    if os.path.exists('./config.json'):
        configs = json.loads(open('./config.json', 'r').read())
        username = configs["username"]
        password = configs["password"]
        
    else:
        username = input("👤 网站账户名: ")
        password = getpass.getpass('🔑 网站密码: ')

    print("\n 获取用户账号、密码成功!")
    print("\n 🚀 Selenium 测试启动！\n")
    spinner = Halo(text='Loading', spinner='dots')

    spinner.start('启动浏览器中 ...\n')
    web = SysTest(username, password)
    spinner.succeed('已启动浏览器')

    spinner.start(text='测试功能一: 豆瓣读书  ...')
    web.login()
    spinner.succeed('功能一测试成功！！')

    # spinner.start(text='测试功能二: 豆瓣电影 ✨  ...')
    print("⠹ 测试功能二: 豆瓣电影 ✨  ... \n")

    web.readCases(1,3) # 测试 2.1.0
    web.test(1,3)
    # web.readCases(3,3) # 测试 2.1.1
    # web.test()
    # web.readCases(3,16) # 测试 2.1.2
    # web.test()
    # web.readCases(16,29) # 测试 2.1.3
    # web.test()
    # web.readCases(29,42) # 测试 2.1.4
    # web.test()
    # web.readCases(42,44) # 测试 2.1.5
    # web.test()
    # web.readCases(44,46) # 测试 2.1.6
    # web.test()

    web.readCases(3,48) # 测试 2.1.1-2.1.6
    web.test(3,48)

    web.readCases(48,60) # 测试 2.2
    web.test(48,60)

    web.readCases(60,75) # 测试 2.3
    web.test(60,75)

    web.readCases(75,90) # 测试 2.4
    web.test(75,90)
    # web.finish()
    
    # web = SysTest(username, password)
    # web.login()
    # web.readCases(1,3) # 测试 2.1.0
    # web.test()

    web.readCases(90,149) # 测试 2.5
    web.test(90,149)

    # web.readCases(90,92) # 测试 2.5
    # web.test()
    # web.readCases(126,152) # 测试 2.5
    # web.test()

    # spinner.succeed('功能二测试成功！！')
    print("✔ 功能二测试完成！！\n")
    
    # spinner.start(text='测试功能三: 豆瓣同城  ...')
    # spinner.succeed('功能三测试完成！！')

    # spinner.start(text='测试功能四: 豆瓣小组  ...')
    # spinner.succeed('功能四测试完成！！')

    # spinner.start(text='测试功能五: 豆瓣？？  ...')
    # spinner.succeed('功能五测试完成！！')

    # spinner.start(text='测试功能六: 豆瓣？？  ...')
    # spinner.succeed('功能六测试完成！！')

    web.finish()

if __name__=="__main__":
    main() # username, password

    