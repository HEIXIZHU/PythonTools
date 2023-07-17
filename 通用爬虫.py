#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# 通用爬虫
# 需求：有部分手工整理信息的繁琐，并且需要整理到特定格式文档；虽说不要重复造轮子，但是总有特殊场景，
# 上网搜爬虫、改爬虫又要花费不少时间
# 功能：1.能够所见即所得，通过简单输入，快速明确一级页面和二级页面选中哪些信息，无需重复修改代码；
# 2.能够自由的对爬取的信息赋予变量名，方便存储到excel中；
# 3.能够自由的把变量组合成excel的索引。
# 4.能够保存专用模板。


# In[ ]:


# 模块测试
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from urllib.parse import urljoin
import warnings
warnings.filterwarnings('ignore')

# 加载模块
def loading():
    # get直接返回，不再等待界面加载完成
    desired_capabilities = DesiredCapabilities.EDGE
    desired_capabilities["pageLoadStrategy"] = "none"
    # 设置谷歌驱动器的环境
    options = webdriver.EdgeOptions()
    # 设置chrome不加载图片，提高速度
    #options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
    # # 设置不显示窗口
    # options.add_argument('--headless')
    # 创建一个谷歌驱动器
    global driver
    driver = webdriver.Edge(options=options)
    require()
    crawl()


# 下次开发：建立一个元素可用性检查系统，先检查xpath是否正确，包括文本获取与点击操作
# 下次开发：翻页按钮、二级爬取、文档写入
# 下下次维护：检查爬取完成还是爬取失败/要输入验证码，以决定下一步操作；网页模板存取；尽量封装为函数
# 下下下次：研究一下cookies
# 下下下下次维护：点选按钮

# 请求参数模块
def require():
    ask = eval(input('是否启用固定模板？'))
    if ask == 1:
        #with open():
        savelists = ''
        choice = input('模板选择：'+ savelists)
    else:
        pass
    # 翻页按钮
    global pagenext
    global nextpage
    pagenext = eval(input('是否需要翻页？1/2'))
    if pagenext == 1:
        nextpage = input('请输入翻页按钮xpath：')
    else:
        pass
    link_i = input('输入需要爬取的界面：')
    signin = eval(input('是否需要登录：1/2'))
    driver.get(link_i)
    time.sleep(2)
    if signin == 1:
        wait = input('登录完成后按enter')
    else:
        pass
    time.sleep(2)
    global levels
    levels = eval(input('是否有二级界面：1/2'))
    # parts存储目标字段名，lst存储目标字段名对应的xpath
    global parts
    parts = input('请用英文输入一级界面字段名，并以逗号隔开：')
    parts = parts.split(',')
    global lst1
    lst1 = []
    for i in range(1,len(parts)+1):
        a = f'element{i}XP'
        lst1.append(a)
    for i in lst1:
        name22 = parts[lst1.index(i)]
        globals()[i] = input(f'一级界面字段{name22}位置，变化部分用term占位：') 
    if levels == 1:
        global parts2
        parts2 = input('请用英文输入二级界面的字段名，并以逗号隔开：')
        parts2 = parts2.split(',')
        global lst2
        lst2 = []
        for i in range(1,len(parts2)+1):
            a = f'L2_element{i}XP'
            lst2.append(a)
        for i in lst2:
            name23 = parts2[lst2.index(i)]
            globals()[i] = input(f'二级界面字段{name23}位置，变化部分用term2占位：') 

    

def crawl():
    df = pd.DataFrame(columns= parts)
    print('开始爬取信息……')
    # 等待加载完全，休眠3S
    time.sleep(3)
    #title_list = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "fz14")))
    # 循环网页一页中的条目
    locc = 0 # locc用于确定，该次循环处理第几条字段
    #对每条字段进行遍历处理
    for i in parts:
        globals()[i] = [] 
        term = 1
        while True:
            #try:
                target = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, globals()[lst1[locc]].replace('term',str(term))))).text
                globals()[i].append(target)
                print(globals()[i])
                if levels == 1:
                    # 获取driver的句柄
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, globals()[lst1[locc]].replace('term',str(term))))).click()
                    # 假设打开了窗口后，当前窗口的句柄是handle1
                    handle1 = driver.current_window_handle

                    # 切换到其他窗口
                    handles = driver.window_handles
                    for handle2 in handles:
                        if handle2 != handle1:
                            driver.switch_to.window(handle2)
                            break
                    #n = driver.window_handles
                    # driver切换至最新生产的页面
                    #driver.switch_to.window(n[-1])
                    time.sleep(3)
                    locc2 = 0
                    for k in parts2:
                        term2 = 1
                        while True:
                            #try:
                                globals()[k] = []
                                target2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, globals()[lst2[locc2]].replace('term2',str(term2))))).get_attribute("innerText")
                                globals()[k].append(target2)
                                term2 += 1
                            #except Exception as e:
                             #   print(e)
                              #  break
                        print(globals()[k])     
                        locc2+=1
                n2 = driver.window_handles
                if len(n2) > 1:
                        driver.close()
                        driver.switch_to.window(n2[0])
                term += 1
            #except Exception as e:
             #   print(e)
              #  if pagenext == 1:
               #     WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, nextpage))).click()
                #    time.sleep(2)
               # else:
                #    print(f'字段{i}爬取完成。')
                 #   break
        print(globals()[i])
        locc += 1
    saveask = input('是否将上述参数保留为固定模板？')
    driver.close()
    
loading()

