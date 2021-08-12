from re import L
import requests
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from lxml import etree
import os
from time import sleep
import xlrd
import win32api
import win32con
import xlsxwriter as xw
from openpyxl import load_workbook
# #实现无可视化界面
# from selenium.webdriver.chrome.options import Options
# #实现规避检测
# from selenium.webdriver import ChromeOptions
# #实现无可视化界面的操作
# chrome_options = Options()
# chrome_options.add_argument('--headless')
# chrome_options.add_argument('--disable-gpu')
# #实现规避检测
# option = ChromeOptions()
# option.add_experimental_option('excludeSwitches', ['enable-automation'])
# #如何实现让selenium规避被检测到的风险
# brs = webdriver.Chrome(executable_path='./selenium/chromedriver',chrome_options=chrome_options,options=option)
def Load_ExcelDone(list,n=0):
    data=xlrd.open_workbook('./paper/paper.xlsx')
    table=data.sheets()[0]
    nrows=table.nrows
    for i in range(nrows):
        try:list.append(int(float(table.row_values(i)[n])))
        except:
            continue
    print(list)

def add_information_excel(filename,paper,list_done):
    workbook=load_workbook(filename+'.xlsx')
    wb=workbook.active
    for p in paper:
        column_n='A'+str(p.n)
        column_name='B'+str(p.n)
        column_wos='C'+str(p.n)
        column_url='D'+str(p.n)
        wb[column_n]=p.n
        wb[column_name]=p.name
        wb[column_wos]=p.wos
        wb[column_url]=p.url
    workbook.save(filename+'.xlsx')

def savepage_pywin32():
    win32api.keybd_event(17, 0, 0, 0)           # 按下ctrl
    win32api.keybd_event(83, 0, 0, 0)           # 按下s
    win32api.keybd_event(83, 0, win32con.KEYEVENTF_KEYUP, 0)    # 释放s
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)    # 释放ctrl
    sleep(1)
    win32api.keybd_event(13, 0, 0, 0)           # 按下enter
    win32api.keybd_event(13, 0, win32con.KEYEVENTF_KEYUP, 0)    # 释放enter

def search(kw):
    #seach in sci
    search_input=brs.find_element_by_xpath('//input[@data-ta="search-criteria-input"]')    
    try:  
        wind=brs.find_element_by_id('pendo-close-guide-8fdced48')
        wind.click()                     #find serch box
        search_input.click()
        search_input.clear()
        search_input.send_keys(kw)    
        search_input.send_keys(Keys.ENTER)                                          #input keywords
        butn=brs.find_element_by_xpath('//span[@class="mat-button-wrapper"]')   #find search button
        butn.click()          
    except:
        search_input.click()
        search_input.clear()
        search_input.send_keys(kw)    
        search_input.send_keys(Keys.ENTER)                                          #input keywords
        butn=brs.find_element_by_xpath('//span[@class="mat-button-wrapper"]')   #find search button
        butn.click()                                                            #click serch button and serch            

def getsource(url,broser):
    broser.get(url)
    page_text=broser.page_source
    print('page_souce load successful')
    return page_text

def SaveHtml(HTML,Filename):
    if not os.path.exists('./paper/HTML'):
        os.makedirs('./paper/HTML')
    Filename=Filename+'.html'
    with open('./paper/HTML/'+Filename,'wb',encoding='utf-8') as fp:
        fp.write(HTML)
    return 'save successful'

def closewind(s):
    brs.find_element_by_id(s).click

def getpaper_wos_url(source):
    tree=etree.HTML(source)
    download=tree.xpath('//app-records-list//a[@data-ta="summary-record-title-link"]/@href')[0]
    Download='https://www.webofscience.com'+download
    return Download

def load_BasicExcel(lists,number,n=7):
    data=xlrd.open_workbook('./paper/gao.xlsx')
    table=data.sheets()[0]
    nrows=table.nrows
    if nrows>number:
        nrows=number
    for i in range(nrows):
        if i<=1:
            continue
        elif table.row_values(i)[n]==table.row_values(i-1)[n]:
            continue
        else :
            s=paper(i,table.row_values(i)[n],0,0)
            lists.append(s)

def getHTML(url,headers = { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36'}):
    c=requests.get(url=url,headers=headers).content
    return c

def getfulltext_url(page):
    tree=etree.HTML(page)
    c=tree.xpath('//app-full-record-links//a/@href')[0]
    return c



class paper:
    def __init__(self,n,name,wos,url):
        self.n=n
        self.name=name
        self.wos=wos
        self.url=url

brs = webdriver.Chrome(executable_path='./selenium/chromedriver')
url='https://www.webofscience.com/wos/woscc/basic-search'
#create list
Paper_list=[]           #save paper (url,name,i,)
Paper_list_result=[]
# url_fulltest_list=[]    #save url
Paper_list_failed=[]    #save paper_url_failed
Paper_list_success_n=[]
Load_ExcelDone(Paper_list_success_n)
Paper_list_done=Paper_list_success_n
load_BasicExcel(Paper_list,50) # load data

#Seach and find paper_url
for kw in Paper_list:
    if kw.n not in Paper_list_success_n:
        try:  
            brs.get(url)
            sleep(6)
            search(kw.name)
            sleep(2)
            wos=getpaper_wos_url(brs.page_source)
            sleep(3)
            brs.get(wos)
            sleep(2)
            url_fulltest=getfulltext_url(brs.page_source)
            # #save page by ctrl+s - enter 
            # brs.get(url_fulltest)
            # sleep(6)
            # savepage_pywin32()
            # sleep(20)
            #save in list
            kw.wos=wos
            kw.url=url_fulltest
            Paper_list_result.append(kw)
            # url_fulltest_list.append(kw.url)

            #save name of success
            # Paper_list_success.append(kw.name)
            Paper_list_success_n.append(kw.n)
            print(kw.n,'load susccessful')
        except:
            #save name of failed
            print(kw.n,'load failed')
            Paper_list_failed.append(kw)

brs.quit()

add_information_excel('./paper/paper',Paper_list_result,Paper_list_done)
add_information_excel('./paper/paper_failed',Paper_list_failed,[])
