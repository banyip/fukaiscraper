# -*- coding: utf-8 -*-
"""
Created on Tue Dec 15 15:23:06 2020

@author: Administrator
"""

import pandas as pd
import datetime
import time
import os
import threading

#文件服开数据更新处理

def dataMerge(strTime):
    starttime = datetime.datetime.now()
    currentday = datetime.datetime.strptime(strTime, '%Y-%m-%d')
    yesterday_strTime=(currentday - datetime.timedelta(days = 1)).strftime("%Y-%m-%d") 
    if (currentday.weekday()>0):
        daybefore= currentday - datetime.timedelta(days = 1)    
    else:
        daybefore= currentday - datetime.timedelta(days = 3)

    if (currentday.weekday()==1):
        daybefore1=daybefore - datetime.timedelta(days = 3)
    else:
        daybefore1=daybefore - datetime.timedelta(days = 1)
    daybefore_strTime = daybefore.strftime("%Y-%m-%d") 
    daybefore1_strTime = daybefore1.strftime("%Y-%m-%d") 
    today_path = os.path.join('output',strTime)
    yesterday_path = os.path.join('output',daybefore_strTime)
    

    dtype={
        #'五级地址ID':str,
        #'CRM业务流水号':str,
        '工单号':str,
        '客户原因缓装申请时间': 'category',
        'CRM业务流水号': 'object',
        'CRM工单号': 'object',
        'IVR拨打成功的时间': 'category',
        'OBD端口名称': 'object',
        'ONU厂家': 'category',
        'ONU型号': 'category',
        'ONU端口端口号': 'object',
        'SN码': 'category',
        '一级场景属性': 'category',
        '上门测试到单时间': 'category',
        '上门测试结束时间': 'category',
        '下单地市': 'category',
        '下单渠道名称': 'category',
        '下单渠道编码': 'category',
        '业务受理地市': 'category',
        '中高端客户': 'category',
        '二级地址名称': 'category',
        '二级场景属性': 'category',
        '五级地址ID': 'category',
        '五级地址名称': 'category',
        '产品业务属性': 'category',
        '产品包ID': 'category',
        '产品包名称': 'category',
        '产品名称': 'category',
        '产品实例编码': 'object',
        '光功率达标情况': 'category',
        '前端首次催办时间': 'category',
        '勘察单到单时间': 'category',
        '勘察单归档时间': 'category',
        '勘察单现场勘查到单时间': 'category',
        '勘察单现场勘查完成时间': 'category',
        '勘察单调度中心到单时间': 'category',
        '勘察单调度中心完成时间': 'category',
        '勘察单预约上门到单时间': 'category',
        '勘察单预约上门完成时间': 'category',
        '原标准地址': 'category',
        '原标准地址1': 'category',
        '处理渠道': 'category',
        '外线施工完成时间': 'category',
        '外线施工开始时间': 'category',
        '多媒体家庭电话': 'category',
        '客户上门时限要求': 'category',
        '客户原因缓装原因': 'category',
        '客户原因缓装状态': 'category',
        '客户原因缓装结束时间': 'category',
        '客户同意缓装申请时间': 'category',
        '客户同意缓装结束时间': 'category',
        '客户类型': 'category',
        '客户需求带宽': 'category',
        '宽带帐号': 'object',
        '宽带提供方': 'category',
        '小区地址': 'object',
        #'工单号': 'object',
        '工单来源': 'category',
        '工单状态': 'category',
        '工单质检到单时间': 'category',
        '工单质检完成时间': 'category',
        '当前环节': 'category',
        '微区名称': 'category',
        '批次号': 'category',
        '支付情况': 'category',
        '支付类型': 'category',
        '改址类型': 'category',
        '改约原因': 'category',
        '改约时间': 'category',
        '数据制作到单时间': 'category',
        '数据制作结束时间': 'category',
        '新阶段回复': 'category',
        '旧ONU SN码': 'category',
        '是否480小区': 'category',
        '是否上门收费': 'category',
        '是否催装': 'category',
        '是否建装一体化': 'category',
        '是否异地': 'category',
        '是否施工标识': 'category',
        '是否派发给装维': 'category',
        '是否需要后台协助': 'category',
        '更换ONU是否成功': 'category',
        '最新一次催办来源': 'category',
        '最近一次的拨号开始时间': 'category',
        '最近一次的拨号结束时间': 'category',
        '服务厅': 'category',
        '未预约原因': 'category',
        '标准地址': 'object',
        '派单日期': 'object',
        '测速方式': 'category',
        '测速时间': 'category',
        '渠道类型': 'category',
        '用户班/装维组': 'category',
        '申请人': 'category',
        '申请退单备注': 'category',
        '第三方复核到单时间': 'category',
        '第三方复核完成时间': 'category',
        '管线资源配置环节到达时间': 'category',
        '终端来源': 'category',
        '结单时间': 'category',
        '综合网格': 'category',
        '缓装原因': 'category',
        '缓装开始时间': 'category',
        '缓装激活时间': 'category',
        '网管和管线数据对比审核到单时间': 'category',
        '装维上门时间': 'category',
        '装维人员/用户班人员': 'category',
        '装维人员登录账号': 'category',
        '装维人员账号': 'category',
        '装维修改五级地址时间': 'category',
        '装维备注': 'category',
        '装维组/用户班': 'category',
        '设备型号': 'category',
        '设备码': 'category',
        '设备类型': 'category',
        '调度中心到单时间': 'object',
        '调度中心结束时间': 'object',
        '质检整改到单时间': 'category',
        '质检整改完成时间': 'category',
        '资源类型': 'category',
        '通过管线资源配置完成时间': 'category',
        '速率变更单_前端已撤单': 'category',
        '速率测试是否达标': 'category',
        '速率测试达标情况': 'category',
        '阶段回复': 'category',
        '集中预约到单时间': 'category',
        '集中预约结束时间': 'category',
        '集团客户批次': 'category',
        '预约上门时间': 'category',
        '预约人': 'category',
        '预约完成时间': 'category',
        '预约状态': 'category',
        '首次催办时间': 'category',
        '首次催装环节': 'category',
        '首次拨号时间': 'category',
        '首次拨号时间.1': 'category',
        '首次第三方复核结果': 'category',
        '首次质检结果': 'category',
        '首次预约上门时间': 'category',
    }
    
    
    print('开始读入服开数据更新'+os.path.join(yesterday_path,'服开数据更新'+daybefore1_strTime[5:]+'.csv'))
    #dfa=pd.read_excel(os.path.join(yesterday_path,'服开数据更新'+daybefore1_strTime[5:]+'.xlsx'),dtype=dtype)
    dfa=pd.read_csv(os.path.join(yesterday_path,'服开数据更新'+daybefore1_strTime[5:]+'.csv'),engine='python',dtype=dtype,encoding='utf-8_sig')
    endtime = datetime.datetime.now()
    duringtime = endtime -  starttime
    starttime = endtime
    print("read 服开数据更新"+daybefore1_strTime[5:]+'.csv for' +str(duringtime.seconds)+'seconds')
    df1=pd.read_csv(os.path.join(today_path,"家客工单导出_佛山(" + strTime + ").csv"),engine='python',dtype=dtype)#,encoding='utf-8_sig')




    df1['工单号']=df1['工单号'].map(lambda x: str(x[1:]))
    #dfa['工单号']=dfa['工单号'].map(lambda x: "'" + str(x))
    df1.insert(17,'区域','')
    df1.insert(17,'代维','')
    df1.insert(27,'归档时长','')
    df1.insert(36,'预约时长（IVR）','')
    df1.insert(36,'预约时长（旧）','')
    df1.insert(36,'派单时间次日8点','')
    df1.insert(36,'派单时间段','')
    df1.insert(36,'勘查到单时间段','')
    df1.insert(36,'勘查到单时间次日8点','')
    def getNext8(dt):
        result=None
        if not pd.isna(dt):
            result=datetime.datetime(int(dt.year),int(dt.month),int(dt.day),8,0)
        return result
    df1['勘察单到单时间']=pd.to_datetime(df1['勘察单到单时间'],errors='coerce')
    df1['勘查到单时间次日8点']=df1['勘察单到单时间']
    df1['勘查到单时间次日8点']=df1['勘查到单时间次日8点'] + pd.offsets.Day(1)
    df1['勘查到单时间次日8点']=df1['勘查到单时间次日8点'].map(getNext8)

    df1['勘查到单时间段'] = df1['勘察单到单时间'].dt.hour     
    df1['勘查到单时间段']=df1['勘查到单时间段'].fillna(0)
    df1['勘查到单时间段']=df1['勘查到单时间段'].astype('int') 

    df1['派单日期']=pd.to_datetime(df1['派单日期'],errors='coerce')
    df1['派单时间次日8点']=df1['派单日期']
    df1['派单时间次日8点']=df1['派单时间次日8点'] + pd.offsets.Day(1)
    df1['派单时间次日8点']=df1['派单时间次日8点'].map(getNext8)

    df1['派单时间段'] = df1['派单日期'].dt.hour     
    df1['派单时间段']=df1['派单时间段'].fillna(0)
    df1['派单时间段']=df1['派单时间段'].astype('int') 


    

    

    dfa=pd.concat([df1,dfa],sort=False)
    #dfa.to_csv(today_path+'afterconcat.csv',encoding='utf-8_sig',index=False)
    #2、删除重复项（工单号）
    dfa.drop_duplicates(subset=['工单号'],keep='first',inplace=True)
    #dfa.to_csv(today_path+'afterdropduplicate.csv',encoding='utf-8_sig',index=False)

    '''3、[操作类型]：业务拆除、业务开通、业务移机、预勘查
        4、[是否施工标识]：普通开通施工
        5、[资源类型]：自建宽带、空白
        6、筛选完后保存一份，命名为服开数据更新-×××
    '''
    dfa=dfa[dfa['操作类型'].isin(['业务拆除','业务开通','业务移机','预勘查'])&(dfa['是否施工标识']=='普通开通施工')&((dfa['资源类型'].isin(['自建宽带','第三方（客服服务）']))|dfa['资源类型'].isnull())]
    #dfa.to_excel(writer,sheet_name='afterfilter')
    #dfa.to_csv(today_path + "服开数据更新"+strTime[5:]+".csv",encoding='utf-8_sig')


    #7、[资源类型]前插入2列：区域、代维，标准地址筛选，先填代维“中鹏粤=宝洪=丽普盾、祥昊、广俊、铁通只填五区即可，再填区域，填区名不加区字，（如区域：禅城、代维：禅城）空白的再看装维组列。三水、高明只有铁通，没有丽普盾。
    def getDaiWei(df):        
        strAddress=str(df['标准地址'])
        result=''
        if((strAddress.find('中鹏粤')!=-1)|(strAddress.find('宝洪')!=-1)|(strAddress.find('丽普盾')!=-1)):
            result='丽普盾'
        elif(strAddress.find('祥昊')!=-1):        
            result='祥昊'
        elif(strAddress.find('广俊')!=-1):
            result='广俊'
        elif(strAddress.find('顺德')!=-1):
            result='顺德'
        elif(strAddress.find('南海')!=-1):
            result='南海'
        elif(strAddress.find('禅城')!=-1):
            result='禅城'
        elif(strAddress.find('三水')!=-1):
            result='三水'
        elif(strAddress.find('高明')!=-1):
            result='高明'
        return result

    def getDistrict(df):        
        strAddress=str(df['标准地址'])
        result=''
        if(strAddress.find('顺德')!=-1):
            result='顺德'
        elif(strAddress.find('南海')!=-1):
            result='南海'
        elif(strAddress.find('禅城')!=-1):
            result='禅城'
        elif(strAddress.find('三水')!=-1):
            result='三水'
        elif(strAddress.find('高明')!=-1):
            result='高明'
        return result
    
    dfa['代维'] = dfa.apply(getDaiWei,axis=1)
    dfa['区域'] = dfa.apply(getDistrict,axis=1)

    #8、[归档时间]右插入1列“归档时长”，复制公式进去。
    

    dfa['勘察单调度中心到单时间']=pd.to_datetime(dfa['勘察单调度中心到单时间'],errors='coerce')
    dfa['结单时间']=pd.to_datetime(dfa['结单时间'],errors='coerce')
    dfa['派单日期']=pd.to_datetime(dfa['派单日期'],errors='coerce')

    def getGuiDangDuration(df):
        result=None
        if((df['结单时间'] is None)|(df['结单时间'] is pd.NaT)):
            result=None
        elif((df['勘察单调度中心到单时间'] is not None)&(df['勘察单调度中心到单时间'] is not pd.NaT)):
                result = df['结单时间'].__sub__(df['勘察单调度中心到单时间']).days * 24
        elif((df['派单日期'] is not None)&(df['派单日期'] is not pd.NaT)):
            result = df['结单时间'].__sub__(df['派单日期']).days * 24
        else:
            result=None        
        return result

    dfa['归档时长']=dfa.apply(getGuiDangDuration,axis=1)

    
    
    '''
    多线程(未完成)
    def writeExcel(dfa):
        starttime=datetime.datetime.now()
        #writer=pd.ExcelWriter(os.path.join(today_path , "服开数据更新"+strTime[5:]+".xlsx"))
        #dfa.to_excel(excel_writer=writer,sheet_name='sheet1') 
        dfa.to_csv(os.path.join(today_path , "服开数据更新"+daybefore_strTime[5:]+".csv"),encoding='utf-8_sig',index=False)
        #writer.save()
        #writer.close()        
        endtime = datetime.datetime.now()
        duringtime = endtime -  starttime
        print("write 服开数据更新"+daybefore_strTime[5:]+".csv for" +str(duringtime.seconds)+'seconds')
    
    
    
    
    t1 = threading.Thread(target=dfa.tocsv,args=(file_path,encoding='utf-8_sig'))
    threads.append(t1)
    t1.setDaemon(True)
    t1.start()
    '''
    try:
        dfa.drop(['工单号1'],axis=1,inplace=True)
    except Exception as e:
        print(e)
    dfa['工单号1']=dfa['工单号'].map(lambda x: "'" + str(x))

    print("开始写入服开数据更新"+strTime[5:]+".xlsx")
    starttime=datetime.datetime.now()
    dfa.to_csv(os.path.join(today_path , "服开数据更新"+yesterday_strTime[5:]+".csv"),encoding='utf-8_sig',index=False)
    endtime = datetime.datetime.now()
    duringtime = endtime -  starttime
    print("write 服开数据更新"+daybefore_strTime[5:]+".csv for" +str(duringtime.seconds)+'seconds')
    
    dfa.CRM业务流水号=dfa.CRM业务流水号.apply(lambda x:x[1:]).astype('str')
    '''
    dfb=dfa.loc[dfa['操作类型']=='预勘察']
    dfb.to_excel(excel_writer=writer,sheet_name='服开预勘查工单明细')    
    
    
    dfc=dfa.loc[dfa['操作类型'].isin(['业务拆除','业务开通','业务移机'])]
    '''
    return dfa

    
if __name__ == '__main__':
    threads = []
    dataMerge('2021-01-11')
    for t in threads:        
        t.join()