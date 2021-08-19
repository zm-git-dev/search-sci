from sys import path
from time import sleep
import pyperclip
import defination as df
from defination import judge_filename
from lxml import etree
import os

def load_pdf_done(path,list_filename):
    for root,dirs,files in os.walk (path):
        for filename in files:
            filename=filename.replace('.pdf','')
            filename=int(float(filename))
            list_filename.append(filename)
    return(list_filename)

def load_Content_Excel(path,lists,n=7):
    data=df.xlrd.open_workbook(path)
    table=data.sheets()[0]
    nrows=table.nrows
    for i in range(nrows):
        if i<=1:
            continue
        elif table.row_values(i)[n]==table.row_values(i-1)[n]:
            continue
        else :
            try:a=int(float(table.row_values(i)[n]))
            except:
                continue
            a=int(float(table.row_values(i)[n]))
            s=df.paper(a,table.row_values(i)[n+1],table.row_values(i)[n+2],table.row_values(i)[n+3])
            lists.append(s)

list_paper=[]
list_paper_done=[]
load_pdf_done('paper_pdf',list_paper_done)
print(list_paper_done)
load_Content_Excel('./excel/paper.xlsx',list_paper,n=0)
for u in list_paper:
    try:
        name=judge_filename(u.n)
        path_dir='./paper_pdf/'+name
        if not os.path.exists(path_dir):
            os.makedirs(path_dir)
        if u.n in list_paper_done:
            continue
        df.brs.get(u.url)
        sleep(10)
        tree=etree.HTML(df.brs.page_source)
        pdf_url=tree.xpath('//div[@class="toolbar-buttons content-box"]/ul/li[@class="PrimaryCtaButton"]/a/@href')[0]
        pdf_url='https://www.sciencedirect.com'+pdf_url
        df.brs.get(pdf_url)
        sleep(10)
        save_path=r'C:\Users\28606\Documents\vscode\python\paper\paper_pdf'+'\\'+ str(u.n)
        pyperclip.copy(save_path)
        df.savepage_pywin32()
        sleep(5)
        print('-----------------------------',u.n,'download success----------------------------------')
    except:
        
        print('===================================',u.n,'failed ============================')
df.brs.quit()
