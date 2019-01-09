from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import unittest, time, re, os, mmap, shutil, glob, time, logging
import pymssql # http://pymssql.org/en/latest/index.html 
import os,csv,linecache
import codecs
from DB import ExecuteQuery as exeQry
from bs4 import BeautifulSoup 
from Setting import Config as conf
from clsAirQulityByPymssql import AirQulity as AQ
from datetime import datetime
import datetime


def WriteLog(strMsg,level="info"):
    print(strMsg)
    if (level == "debug"): 
        logging.debug("," + GetDateTime() + "," + strMsg)    
    elif (level == "info"):
        logging.info("," + GetDateTime() + "," + strMsg)
    elif (level == "warning"):
        logging.warning("," + GetDateTime() + "," + strMsg)
    elif (level == "error"):
        logging.error("," + GetDateTime() + "," + strMsg)

def GetDateTime():
    localtime = time.localtime(time.time())
    strDateTime = str(localtime[0]) + "-" + str(localtime[1]).zfill(2) + "-" + str(localtime[2]).zfill(2) + " " + str(localtime[3]).zfill(2) + ":" + str(localtime[4]).zfill(2) + ":" + str(localtime[5]).zfill(2) #'2009-01-05 22:14:39'

    return strDateTime

def GetDate():
    localtime = time.localtime(time.time())
    strDate = str(localtime[0]) + "-" + str(localtime[1]).zfill(2) + "-" + str(localtime[2]).zfill(2) #'2009-01-05'

    return strDate

def LogInit():
    LogPath =  conf.Value('Log','LogPath')  

    if (LogPath == ""):
       CurrentPyFilePath = os.path.dirname(os.path.realpath(__file__))
       LogPath = CurrentPyFilePath + "\\Log"

    MkDirectory(LogPath)   
    logfilename = LogPath + "\\" + GetDate() + ".log"  

    LogLevel = conf.Value('Log','LogLevel')  
    if (LogLevel == "DEBUG"):    
        logging.basicConfig(filename=logfilename,level=logging.DEBUG)         
    elif (LogLevel == "INFO"): 
        logging.basicConfig(filename=logfilename,level=logging.INFO)
    elif (LogLevel == "WARNING"):
        logging.basicConfig(filename=logfilename,level=logging.WARNING)
    elif (LogLevel == "ERROR"): 
        logging.basicConfig(filename=logfilename,level=logging.ERROR)

def MkDirectory(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except:
        os.mkdir(directory) 

def Diff_Dates(d1, d2):
    return abs((d2 - d1).days)

def ConfigInit():
    global ServerIP
    global User
    global Password 
    global DBName
    global CSVFileName

    ServerIP = conf.Value('DB','ServerIP')
    User = conf.Value('DB','User')
    Password = conf.Value('DB','Password')   
    DBName = conf.Value('DB','DBName')  


def GetDBConnection_CallByRef(): 
    try:
        ## get DB connection information, http://pymssql.org/en/latest/pymssql_examples.html
        # Server,User,Password,DBName are global variables
        mConn = pymssql.connect(ServerIP, User, Password, DBName)
        return mConn

    except Exception as e:
        ErrMsg = str(e)
        WriteLog(ErrMsg,level="error")

def combineFilesContent(CombineCSVFileName,StartLine=0):
    funName = "combineFilesContent(FullCombineCSVFileName,StartLine)"
    try: 
        LogInit()

        if StartLine == 0:
            StartLine = int(conf.Value('CSV','ReadCombineFileStartLine'))

        idx = int(StartLine)-1 # 第一行 idx為0

        #預設在python執行檔目錄下
        #scriptpath = os.path.realpath(__file__) # 取得python程式目錄
        #app_folder = scriptpath[:scriptpath.rfind("\\")]
        #outF = codecs.open(app_folder + "\\" + "CombineFullContent.csv", "w", 'utf-8') #寫入 UTF8編碼 檔案 # https://stackoverflow.com/questions/19591458/python-reading-from-a-file-and-saving-to-utf-8/19591815

        OutputCSVFolder = conf.Value('CSV','CombineCSVFolder')

        outF = codecs.open(FullCombineCSVFileName, "w", 'utf-8') #寫入 UTF8編碼 檔案 # https://stackoverflow.com/questions/19591458/python-reading-from-a-file-and-saving-to-utf-8/19591815
        #f = linecache.getlines(OutputCSVFolder + "\\" + CombineCSVFileName)[idx:] # http://blog.51cto.com/wangwei007/1246214
        #reader = csv.reader(f)
        #for row in reader:
        #    NewLine = ','.join(row).strip()
        #    print('new line:' + NewLine)

        #    outF.write(NewLine)
        #    outF.write("\n")
        #    print('writing new line to file.')

        outF.write(NewLine)
        outF.close()
        EndLine = StartLine + 1
        return EndLine

    except Exception as e:
        ErrMsg = "(" + __file__ + "-" + funName + "):  " + str(e)
        print(ErrMsg)
        WriteLog(ErrMsg,level="error")

    finally:
        print(GetDateTime())


def GetTccgAIR_History_WebPage():
    try:
        WaitTime = int(conf.Value('Wait','WaitSecondTime'))        

        WebPage_URL = conf.Value('WebBrowser','HistoryData_URL')    

        #chrome_path = "E:\WEB\TccgAIR\PyApp\GetCityAQIForecast(AutoRun)\chromedriver.exe" 
        # 如果可以開啟模擬瀏覽器但沒有執行頁面抓取動作時可能是chromedriver.exe版本過舊，需要重新抓取新版即可
        # 新版下載網址  https://sites.google.com/a/chromium.org/chromedriver/downloads
        chrome_path = conf.Value('WebBrowser','ChromeBrowserExec')    
        web = webdriver.Chrome(chrome_path)
        web.get(WebPage_URL) 

        # 取得測站清單代碼
        SiteCodeList = conf.Value('AirQualityInfo','SiteCodeList')    
        lstSiteCode = SiteCodeList.split('|')

        # 取得測站監測項目代碼
        MonitorItemCodeList = conf.Value('AirQualityInfo','MonitorItemCodeList')    
        lstMonitorItemCode = MonitorItemCodeList.split('|')

        today = datetime.date.today()
        first = today.replace(day=1)
        lastMonth = first - datetime.timedelta(days=1)
        PreviousMonth_YYYMM = lastMonth.strftime("%Y%m") ## 201901

        for SiteCode in lstSiteCode:
            for MonitorItemCode in lstMonitorItemCode:
                web.find_element_by_id("DropDownList1").click()
                Select(web.find_element_by_id("DropDownList1")).select_by_value(SiteCode)
                web.find_element_by_id("DropDownList1").click()
                web.find_element_by_id("DropDownList2").click()
                Select(web.find_element_by_id("DropDownList2")).select_by_value(MonitorItemCode)
                web.find_element_by_id("DropDownList2").click()
                web.find_element_by_id("DropDownList3").click()
                Select(web.find_element_by_id("DropDownList3")).select_by_value(PreviousMonth_YYYMM)
                web.find_element_by_id("DropDownList3").click()
                web.find_element_by_id("Button1").click() ## 查詢 
                #web.find_element_by_id("Button2").click() ## 下載

                html_gridview_table = web.find_element_by_id('GridView1').find_elements_by_tag_name('tr')

                OutputCSVFolder = conf.Value('CSV','CombineCSVFolder')
                FullCombineCSVFileName = OutputCSVFolder + "\\" + SiteCode + "_" + PreviousMonth_YYYMM + "_" +  MonitorItemCode + ".csv"

                outF = codecs.open(FullCombineCSVFileName, "w", 'utf-8') #寫入 UTF8編碼 檔案 # https://stackoverflow.com/questions/19591458/python-reading-from-a-file-and-saving-to-utf-8/19591815
                ##  https://blog.kelu.org/tech/2017/07/22/python-selenium-sample-for-v2ex-login.html，一行一行讀取
                for index, tr_elem in enumerate(html_gridview_table):
                    row = tr_elem.text
                    NewLine = str.replace(row,' ',',')
                    print('new line:' + NewLine)

                    outF.write(NewLine)
                    outF.write("\n")
                    print('writing new line to file.')

                outF.close()

        web.close
        web.quit() #關閉瀏覽器視窗

    except Exception as e:
        ErrMsg = str(e)
        WriteLog(ErrMsg,level="error")

def main():
    try:   
        LogInit()
        ConfigInit()
        print(GetDateTime())
        GetTccgAIR_History_WebPage()

    except Exception as e:
        ErrMsg = str(e)
        WriteLog(ErrMsg,level="error")

    finally:
        print(GetDateTime())


if __name__ == '__main__':
    # execute only if run as the entry point into the program
    main()