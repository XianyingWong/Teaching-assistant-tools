# 开发时间2021/11/16 21:31
import time

import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")#要在cmd中执行命令：chrome.exe --remote-debugging-port=9222 --user-data-dir="临时创建的Chrome的配置文件夹路径"
chrome_driver = "D:\driver\chromedriver.exe"
ser=Service("D:\\driver\\chromedriver.exe")
driver = webdriver.Chrome(service=ser, options=chrome_options)

workbook = openpyxl.load_workbook(r'C:\Users\86157\Desktop\大学计算机－实验三周四-实验3完整成绩.xlsx')
sheet = workbook.worksheets[0]

driver.implicitly_wait(10)

times=102#循环次数
for i in range(times):
    print(i)
    #获取姓名
    elements=driver.find_elements(By.CSS_SELECTOR,'div.students-pager > h3 > span')
    tem=elements[2].get_attribute('innerText')
    index1=tem.find(' ')
    index2=tem.find(' ',index1+1)
    name=tem[index1+1:index2]

    #在Excel中查找姓名对应的总分和评价
    for j in range(2,len(sheet['B'])+1):
        if name==sheet['B'+str(j)].value:
            # 填入总分
            grade = driver.find_element(By.ID, 'currentAttempt_grade')
            grade.click()
            grade.send_keys(sheet['Q' + str(j)].value)
            time.sleep(1)

            # 填入评价
            driver.switch_to.frame("feedbacktext_ifr")
            driver.find_element('tag name', 'body > p').send_keys(sheet['P' + str(j)].value)
            driver.switch_to.default_content()
            time.sleep(0.5)
            break

    #点击提交
    submit=driver.find_element(By.ID,'currentAttempt_submitButton')
    submit.click()
    time.sleep(0.5)