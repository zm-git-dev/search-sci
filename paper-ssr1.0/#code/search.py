import defination as  df
from time import sleep
#create list
Paper_list=[]           #save paper (url,name,i,)
Paper_list_result=[]
# url_fulltest_list=[]    #save url
Paper_list_failed=[]    #save paper_url_failed
Paper_list_success_n=[]
df.Load_number_ExcelDone('./excel/paper.xlsx',Paper_list_success_n)
Paper_list_done=Paper_list_success_n
df.load_Content_Excel('./excel/paper_failed.xlsx',Paper_list,number=200,n=1) # load data

#Seach and find paper_url
for kw in Paper_list:
    if kw.n not in Paper_list_success_n:
        try:  
            df.brs.get(df.url)
            sleep(6)
            df.search(kw.name)
            sleep(2)
            wos=df.getpaper_wos_url(df.brs.page_source)
            sleep(3)
            df.brs.get(wos)
            sleep(2)
            url_fulltest=df.getfulltext_url(df.brs.page_source)
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

df.brs.quit()
df.add_information_excel('./excel/paper',Paper_list_result)
df.add_information_excel('./excel/paper_failed',Paper_list_failed)
