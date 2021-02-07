# coding=utf-8

from selenium import webdriver
from selenium.webdriver.ie.service import Service
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.select import Select
from captcha import VerificationCode
import time
import datetime
import os
import pandas
import sys
import re
import traceback
import logging
import zipfile
import GeYeMerge
import mingxi1
import mingxi2
import pandas as pd
import openpyxl
import threading
from Logger import Logger
#import PyMail
from pdExcelWriter import pdExcelWriter


def sendEmail(self, reciver, subject, body, *attachmentFilePaths):
    '''
    发送邮件
    :param reciver:
    :param subject:
    :param body:
    :param attachmentFilePaths:
    :return:
    '''
    sml = PyMail.SendMailDealer()
    sml.setMailInfo(reciver, subject, body, 'plain', *attachmentFilePaths)
    sml.sendMail()
    print('邮件发送成功')
    sml.__del__()


def resource_path(relative_path):
    if getattr(sys, 'frozen', False): #是否Bundle Resource
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class NotYetLoggingException(Exception):
    def __init__(self,value):
        self.value = value

    def __str__(self):
        return str(self.value)

class NoDataException(Exception):
    def __init__(self,value):
        self.value = value

    def __str__(self):
        return str(self.value)

class WrongAuthenticationException(Exception):
    def __init__(self,value):
        self.value = value

    def __str__(self):
        return str(self.value)

class UnknownReasonException(Exception):
    def __init__(self,value):
        self.value = value

    def __str__(self):
        return str(self.value)

#登录
def sfLogin(drive):

    while True:
        try:
            username=drive.find_element_by_name('username')
            username.click()
            username.send_keys("zhengzeyuan")
            password=drive.find_element_by_name('password')
            password.click()
            password.send_keys("Zhen@2020")
            ca=VerificationCode(drive,'validationCode')
            while True:
                captcha_str = ca.image_str()
                if(len(captcha_str)==4):
                    print(captcha_str)
                    drive.find_element_by_name('validationCode').send_keys(captcha_str)                        
                    drive.execute_script("$(arguments[0]).click()", drive.find_element_by_xpath('/html/body/div[1]/form/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr[1]/td/table/tbody/tr[5]/td[2]/input'))
                    time.sleep(1)
                    break
                else:
                    drive.find_element_by_id('validationCode').click()   
                    time.sleep(1)                                             
                    continue
            if(drive.page_source.find('验证码错误！')!=-1):
                continue
            else:
                break
        except Exception as e:
            print(traceback.print_exc())                
            time.sleep(5)
            continue

#1.来单兑现率2020-12-21_2020-12-21orderTicket.csv
def sfDuiXianLv(drive,strTime):
        while True:
            try:
                iframe = drive.find_element_by_name("topFrame")
                break
            except:
                time.sleep(5)
                continue

        drive.switch_to.frame(iframe)
        my_jobs_button = drive.find_element_by_link_text('查询统计')
        my_jobs_button.click()

        drive.switch_to.parent_frame()
        iframe=drive.find_element_by_name("mainFrame")
        drive.switch_to.frame(iframe)

        iframe=drive.find_element_by_name("leftFrame")
        drive.switch_to.frame(iframe)
        drive.find_element_by_link_text('家客质量管控').click()
        drive.find_element_by_link_text('480管控报表').click()



        drive.switch_to.parent_frame()
        iframe=drive.find_element_by_name("contFrame")
        drive.switch_to.frame(iframe)

        textinput = drive.find_element_by_id('actions')
        textinput.click()
        textinput.send_keys('业务开通')
        textinput = drive.find_element_by_id('serviceType')
        textinput.click()
        textinput.send_keys('家客开通')


        textinput=drive.find_element_by_id('startDate')
        textinput.clear()
        textinput.click()
        textinput.send_keys(strTime)

        textinput=drive.find_element_by_id('endDate')
        textinput.click()
        textinput.send_keys(strTime)

        textinput=drive.find_element_by_id('condition_region')
        textinput.click()
        textinput.send_keys('佛山')
        textinput.send_keys(Keys.ENTER)

        drive.find_element_by_id('ext-gen100').click()
        time.sleep(10)
        drive.find_element_by_id('ext-gen120').click()
        windows = drive.window_handles 
        drive.switch_to.window(windows[-1])
        drive.execute_script('toDownload()')
        time.sleep(2)
        drive.close()
        drive.switch_to.window(windows[0])


        time.sleep(2)

#2. 家客在途单 家客在途单_佛山_2020-12-21（.zip）
def sfZaiTuTuiDan(drive,strTime):

        drive.switch_to.default_content()
        iframe=drive.find_element_by_name("mainFrame")
        drive.switch_to.frame(iframe)

        iframe=drive.find_element_by_name("leftFrame")
        drive.switch_to.frame(iframe)
        drive.find_element_by_id('sd163').click()
        


        drive.switch_to.parent_frame()
        iframe=drive.find_element_by_name("contFrame")
        drive.switch_to.frame(iframe)

        #操作类型
        textinput = drive.find_element_by_id('actionType')
        drive.execute_script("$(arguments[0]).removeAttr('readOnly')", textinput)
        textinput.click()
        textinput.clear()
        textinput.send_keys('业务开通')


        textinput=drive.find_element_by_id('startDate')
        textinput.clear()
        textinput.click()
        textinput.send_keys(strTime)

        textinput=drive.find_element_by_id('endDate')
        textinput.clear()
        textinput.click()
        textinput.send_keys(strTime)

        textinput=drive.find_element_by_id('condition_region')
        textinput.click()
        textinput.send_keys('佛山')
        textinput.send_keys(Keys.ENTER)

        drive.find_element_by_id('ext-gen128').click()
        time.sleep(20)
        drive.find_element_by_id('ext-gen258').click()
        time.sleep(5)
        



#3. 退单在途工单 隔夜存量在途退单清单导出_佛山_2020-12-22
def sfTuiDanZaiTuGongDan(drive,strTime):
        '''
        drive.switch_to.default_content()
        iframe = drive.find_element_by_name("topFrame")
        drive.switch_to.frame(iframe)
        my_jobs_button = drive.find_element_by_link_text('查询统计')
        my_jobs_button.click()
        '''
        drive.switch_to.default_content()
        iframe=drive.find_element_by_name("mainFrame")
        drive.switch_to.frame(iframe)

        iframe=drive.find_element_by_name("leftFrame")
        drive.switch_to.frame(iframe)
        drive.find_element_by_link_text('退单审核管控').click()
        


        drive.switch_to.parent_frame()
        iframe=drive.find_element_by_name("contFrame")
        drive.switch_to.frame(iframe)


        textinput=drive.find_element_by_id('countTime')
        textinput.clear()
        textinput.click()
        textinput.send_keys(strTime)

        textinput=drive.find_element_by_id('startDate')
        textinput.clear()
        textinput.click()
        textinput.send_keys(strTime)


        drive.find_element_by_id('ext-gen64').click()
        time.sleep(15)
        drive.find_element_by_id('ext-gen85').click()
        time.sleep(5)

#4. 催装报表 ReminderOrderTicket（.csv）
def sfCuiZhuangBaobiao(drive,startTime,endTime):

        drive.switch_to.default_content()
        while True:
            try:
                iframe = drive.find_element_by_name("topFrame")
                break
            except:
                time.sleep(5)
                continue
        drive.switch_to.frame(iframe)
        my_jobs_button = drive.find_element_by_link_text('调度管理')
        my_jobs_button.click()

        drive.switch_to.parent_frame()
        iframe=drive.find_element_by_name("mainFrame")
        drive.switch_to.frame(iframe)

        iframe=drive.find_element_by_name("leftFrame")
        drive.switch_to.frame(iframe)
        drive.find_element_by_link_text('督办').click()
        drive.find_element_by_link_text('催办单管理').click()
        drive.find_element_by_link_text('催办查询').click()


        drive.switch_to.parent_frame()
        iframe=drive.find_element_by_name("contFrame")
        drive.switch_to.frame(iframe)

        iframe=drive.find_element_by_name("topFrame")
        drive.switch_to.frame(iframe)

        textinput=drive.find_element_by_id('requestFirstDate_start')
        textinput.clear()
        textinput.click()
        textinput.send_keys(startTime+ ' 00:00:00')

        textinput=drive.find_element_by_id('requestFirstDate_end')
        textinput.clear()
        textinput.click()
        textinput.send_keys(endTime+ ' 00:00:00')


        drive.execute_script('exportToCSV()')
        t=drive.switch_to.alert
        t.accept()
        time.sleep(15)
        windows = drive.window_handles 
        drive.switch_to.window(windows[-1])
        drive.execute_script('toDownload()')
        time.sleep(10)
        drive.close()
        drive.switch_to.window(windows[0])

#5. 隔夜工单数据导出家客工单导出_佛山_2020-12-22-2020-12-22（.zip）
def sfGeYeGongDanDaoChu(drive,strTime):
        '''
        while True:
            try:
                iframe = drive.find_element_by_name("topFrame")
                break
            except:
                time.sleep(5)
                continue
        '''

        iframe = drive.find_element_by_name("topFrame")
        drive.switch_to.frame(iframe)
        my_jobs_button = drive.find_element_by_link_text('查询统计')
        my_jobs_button.click()

        drive.switch_to.parent_frame()
        iframe=drive.find_element_by_name("mainFrame")
        drive.switch_to.frame(iframe)

        iframe=drive.find_element_by_name("leftFrame")
        drive.switch_to.frame(iframe)
        drive.find_element_by_link_text('查询').click()
        drive.find_element_by_link_text('隔夜工单数据导出').click()



        drive.switch_to.parent_frame()
        iframe=drive.find_element_by_name("contFrame")
        drive.switch_to.frame(iframe)




        textinput=drive.find_element_by_name('startDate')
        textinput.clear()
        textinput.click()
        textinput.send_keys(strTime)
        textinput.send_keys(Keys.ENTER)

        textinput=drive.find_element_by_name('endDate')
        textinput.click()
        textinput.send_keys(strTime)
        textinput.send_keys(Keys.ENTER)

        drive.find_element_by_id('_isJiaKe').click()

        '''
        drive.execute_script('loadSgOrgination1sTree()')

        windows = drive.window_handles 
        drive.switch_to.window(windows[-1])

        drive.find_element_by_id('cd0').click()
        drive.find_element_by_id('cd1').click()
        drive.find_element_by_id('cd2').click()
        drive.find_element_by_id('cd3').click()
        drive.find_element_by_id('cd4').click()
        drive.find_element_by_id('cd5').click()
        drive.execute_script('sel()')

        drive.switch_to.window(windows[0])
        '''
        textinput=drive.find_element_by_name('organization')
        drive.execute_script("$(arguments[0]).removeAttr('readOnly')", textinput)
        textinput.clear()
        textinput.send_keys('佛山_顺德区,佛山_禅城区,佛山_高明区,佛山_南海区,佛山_三水区,佛山_佛山公司')
        drive.execute_script('doDownloadAction()')
        time.sleep(15)
#解压其中的ZIP文件
def sfUnzip(output_folder_path,strTime,strYesterday):
    file_name = os.path.join(output_folder_path,'家客工单导出_佛山_'+strTime+'-' + strTime + '.zip')
    zip_file = zipfile.ZipFile(file_name)
    for names in zip_file.namelist():
        try:
            name1 = names.encode('cp437').decode('gbk')
        except:
            name1 = names.encode('utf-8').decode('utf-8')
        if not os.path.exists(os.path.join(output_folder_path,name1)):
            zip_file.extract(names,output_folder_path)
            os.rename(os.path.join(output_folder_path,names),os.path.join(output_folder_path,name1))
    zip_file.close()

    file_name = os.path.join(output_folder_path,'家客在途单_佛山_'+strYesterday + '.zip')
    zip_file = zipfile.ZipFile(file_name)
    for names in zip_file.namelist():
        try:
            name1 = names.encode('cp437').decode('gbk')
        except:
            name1 = names.encode('utf-8').decode('utf-8')
        if not os.path.exists(os.path.join(output_folder_path,name1)):
            zip_file.extract(names,output_folder_path)
            os.rename(os.path.join(output_folder_path,names),os.path.join(output_folder_path,name1))
    zip_file.close()

def sfwriteExcelOptions(pdew,strTime):
    currentday = datetime.datetime.strptime(strTime, '%Y-%m-%d')
    daybefore=currentday - datetime.timedelta(days=1)
    daysfrommonday = daybefore.weekday()
    daysfromfirstdayofthemonth = 7
   


    dateformonday_str=(currentday - datetime.timedelta(days = daysfromfirstdayofthemonth)).strftime("%Y-%m-%d")
    firstdayofthismonth_str = (daybefore - datetime.timedelta(days = daysfromfirstdayofthemonth - 1 )).strftime("%Y-%m-%d")
    daybefore_str=daybefore.strftime('%Y-%m-%d')
    
    pdoptions=pd.DataFrame({'name':['currentdate','dateformonday','daysfrommonday','firstdayofthismonth','daybefore','daysfromfirstdayofthemonth'], \
        'value':[strTime[:10],dateformonday_str,daysfrommonday,firstdayofthismonth_str,daybefore_str,daysfromfirstdayofthemonth], \
        'description':['当前日期','本周周一日期','周一到昨天的天数','每月1号','前一天日期','当月1号到昨天的天数'] })
    pdew.df2Sheet('options',pdoptions)

def sfBeiZengDataExport(output_folder_path,_signal=None):

    # logging.basicConfig(filename=os.path.join('.', 'error_log.log'), level=logging.INFO, \
    #                     format="%(asctime)s %(name)s %(levelname)s %(message)s", datefmt='%Y-%m-%d  %H:%M:%S %a')
    # logger_name = 'sfTuiDan'
    # logger = logging.getLogger(logger_name)

    TIME_STR_FORMAT = '%#Y{y}%#m{m}%#d{d}%#H{h}%#M{f}%#S{s}'
    posfix = time.strftime(TIME_STR_FORMAT, time.localtime()).format(y='年', m='月', d='日', h='时', f='分', s='秒提交')
    result_file_path = os.path.join(output_folder_path, '退单执行结果-' + posfix + '.xlsx')


    
    #获取当日字符串: 
    today_strTime = (datetime.datetime.now()).strftime("%Y-%m-%d") 
    
    #获取昨日字符串: 

    yesterday_strTime = (datetime.datetime.now() - datetime.timedelta(days = 1)) .strftime("%Y-%m-%d") 

    try:

#        data_excel_query_info = pandas.read_excel(inputFile, sheet_name='Sheet1', dtype=str)
#        dataFrame_query_info = pandas.DataFrame(data_excel_query_info)
#        dataFrame_query_info = dataFrame_query_info.applymap(lambda x: x if str(x) != 'nan' else '')

#        row_num = dataFrame_query_info.shape[0]


        
        
        output_folder_path=os.path.join(output_folder_path,today_strTime)
        options = webdriver.ChromeOptions()
        

        prefs = { 'profile.default_content_settings.popups':0 ,'download.default_directory': output_folder_path,"profile.content_settings.exceptions.automatic_downloads.*.setting":1}

                    #设置为0表示禁止弹出窗口，                     #设置文件下载路径

        starttime = datetime.datetime.now()
        print('开始时间：'+starttime.strftime("%Y-%m-%d %H-%M-%S"))
        dfa=None

        options.add_experimental_option('prefs',prefs)
        options.add_argument("--no-sandbox")
        
        drive = webdriver.Chrome(chrome_options=options)
        
        drive.implicitly_wait(20)
        drive.maximize_window()

        lg = Logger(os.path.join(output_folder_path,starttime.strftime("%Y-%m-%d %H-%M-%S")+'.log'))
        sys.stdout = lg
        
        sys.stderr = lg
        
        drive.get("http://sf.gmcc.net:7001/mobilesg/checkUsergd.action")

        
        #登录
        print('登录')
        sfLogin(drive)

        #1.来单兑现率2020-12-21_2020-12-21orderTicket.csv
        if not os.path.exists(os.path.join(output_folder_path,yesterday_strTime+'_'+yesterday_strTime+'orderTicket.csv')):
            print('下载来单兑现率'+yesterday_strTime+'_'+yesterday_strTime+'orderTicket.csv')
            sfDuiXianLv(drive,yesterday_strTime)

        #5. 隔夜工单数据导出家客工单导出_佛山_2020-12-22-2020-12-22（.zip）
        if not os.path.exists(os.path.join(output_folder_path,'家客工单导出_佛山_'+today_strTime+'-'+ today_strTime+'.zip')):
            print('隔夜工单数据导出家客工单导出_佛山_'+today_strTime+'-'+ today_strTime+'.zip')
            sfGeYeGongDanDaoChu(drive,today_strTime)
        
        #2. 家客在途单 家客在途单_佛山_2020-12-21（.zip）
        if not os.path.exists(os.path.join(output_folder_path,'家客在途单_佛山_'+yesterday_strTime+'.zip')):
            print('家客在途单 家客在途单_佛山_'+yesterday_strTime+'（.zip）')
            sfZaiTuTuiDan(drive,yesterday_strTime)

        #3. 退单在途工单 隔夜存量在途退单清单导出_佛山_2020-12-22        
        if not os.path.exists(os.path.join(output_folder_path,'隔夜存量在途退单清单导出_佛山_'+today_strTime+'.csv')):
            print('退单在途工单 隔夜存量在途退单清单导出_佛山_'+today_strTime)
            sfTuiDanZaiTuGongDan(drive,today_strTime)

        #4. 催装报表 ReminderOrderTicket（.csv）
        if not os.path.exists(os.path.join(output_folder_path,'ReminderOrderTicket.csv')):
            print('催装报表 ReminderOrderTicket.csv')
            sfCuiZhuangBaobiao(drive,yesterday_strTime,today_strTime)

        drive.close()



        #5. 解压
        print('解压')
        sfUnzip(output_folder_path, today_strTime,yesterday_strTime)
        

        if (not os.path.exists(os.path.join(output_folder_path,'服开数据更新'+yesterday_strTime[5:]+'.csv'))) or (not os.path.exists(os.path.join(output_folder_path,'当月指标'+yesterday_strTime[5:]+'.xlsx'))):
            #6.整合服开数据更新12-22.csv
            print('#6.整合服开数据更新'+yesterday_strTime[5:]+'.csv')
            dfa=GeYeMerge.dataMerge(today_strTime)
        
        

        
            midtime = datetime.datetime.now()
            print('开始创建pdew')
            pdew=pdExcelWriter(os.path.join('templates','倍增模板.xlsx'))
            sfwriteExcelOptions(pdew,today_strTime)     
    
            endtime = datetime.datetime.now()
            duringtime = endtime -  midtime                
            print('完成创建pdew for ' + str(duringtime.seconds) + '秒')        
            

            #7.mingxi1
            print('mingxi1')        
            #writer=pd.ExcelWriter(os.path.join(output_folder_path,"当月指标.xlsx"),engine='openpyxl')
            writer=None
            mingxi1.mingxi1(output_folder_path,today_strTime,yesterday_strTime,writer,dfa,pdew)
            
            #8.mingxi2
            print('mingxi2')
            mingxi2.mingxi2(output_folder_path,today_strTime,yesterday_strTime,writer,dfa,pdew)
            #writer.save()
            #writer.close()
            
            pdew.saveWorkbook(os.path.join(output_folder_path,'当月指标'+yesterday_strTime[5:]+'.xlsx'))

        endtime = datetime.datetime.now()
        duringtime = endtime -  starttime
        print("all for" +str(duringtime.seconds)+'seconds')
        
        #sendEmail('13923110117@139.com', '倍增指标'+today_strTime[5:0], '倍增指标'+today_strTime[5:0],*[os.path.join(output_folder_path,'当月指标'+today_strTime[5:]+'.xlsx')])
        time.sleep(5)
        

    except Exception as e:
        # print(e)
        print(traceback.print_exc())
        raise e



if __name__ == '__main__':
    sfBeiZengDataExport(resource_path('output'))


