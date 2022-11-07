# 1.連結至玩股網,取得股市交易前200大股票資訊
# 2.利用twstock模組取得股票上市上櫃,類別資訊
# 3.寫檔
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import codecs
#from selenium.webdriver.chrome.options import Options
#import os
#import time

#options = Options()
#options.add_argument("--disable-notifications")

#set chromedriver.exe path
#driver = webdriver.Chrome(executable_path="C:\chromedriver.exe")
#driver.implicitly_wait(0.5)

#chrome = webdriver.Chrome('./chromedriver', chrome_options=options)
#chrome.implicitly_wait(5)
# =======讀Html檔_Start===========

filename='玩股網Top200'
full_path_html_file="d:\\temp\\" + filename + ".html"
full_path_xlsx_file="d:\\temp\\" + filename + ".xlsx"

chrome = webdriver.Chrome()

#launch URL
chrome.get("https://www.wantgoo.com/stock/ranking/turnover")

try:
    element = WebDriverWait(chrome, 20).until(
        #    EC.presence_of_element_located((By.ID, "ottv32v7nduo1667562890485"))
        #    element_p=EC.presence_of_element_located((By.CLASS_NAME, 'pager'))
        #    EC.title_contains("chat")
        #    EC.presence_of_element_located((By.CLASS_NAME, "tawk-min-chat-icon"))
        #    EC.presence_of_element_located((By.CLASS_NAME, "pager"))
        #    EC.title_is("chat widget")
        # <div id="chat-bubble"></div>  #右下角聊天室的 id
        EC.presence_of_element_located((By.ID, "chat-bubble"))
    , "No search element")
finally:
    #time.sleep(1)
    #get file path to save page
    #filename=os.path.join(r"d:\temp","玩股網Top200.html")
    #open file in write mode with encoding
    file = codecs.open(full_path_html_file, "w", "utf−8")
    #obtain page source
    html = chrome.page_source

    #print(html)
    #write page source content to file
    file.write(html)
    file.close()

    #close browser
    chrome.quit()

# =======讀Html檔_End===========

#maximize browser
#driver.maximize_window()




import pandas as pd
#讀出Html檔
df = pd.read_html(full_path_html_file)
df2=df[0]              #讀出來進DataFrame

import twstock
df3 = pd.DataFrame()   #建立空的DataFrame
index=0
for row in df2["代碼"]:
    df3.loc[index,"market"]=twstock.codes[row].market   #抓market
    df3.loc[index,"group"]=twstock.codes[row].group     #抓分類
    index+=1


df_Total=pd.concat([df2,df3],axis=1)  #以烈為基底來做合併

df_Total.to_excel(full_path_xlsx_file) #寫至xlsx檔案