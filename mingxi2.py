# -*- coding: utf-8 -*-
"""
Created on Mon Dec  7 15:26:34 2020

@author: Administrator
"""
import pandas as pd
import datetime
import os
#import time

def mingxi2(file_path,strToday,strYesterday,writer,df5):
    starttime = datetime.datetime.now()
    yesterday_path = os.path.join('output',strYesterday)    

    df3=pd.DataFrame(pd.read_csv(os.path.join(file_path,strYesterday+'_'+strYesterday+'orderTicket.csv'), engine='python'))
    #df3.开通工单号=df3.开通工单号.apply(lambda x:x[1:]).astype('str')
    df3.五级地址ID=df3.五级地址ID.astype('str')
    df3.drop(['CRM业务流水号'],axis=1,inplace=True)

    df4=pd.DataFrame(pd.read_csv(os.path.join(file_path,'ReminderOrderTicket.csv'), engine='python',dtype={'五级地址ID': str,}))#催单
    df4.CRM业务流水号=df4.CRM业务流水号.apply(lambda x:x[1:]).astype('str')
    df4.产品名称=df4.产品名称.apply(lambda x:x[1:]).astype('str')
    df4.客户类型=df4.客户类型.apply(lambda x:x[1:]).astype('str')
    #df4.工单号=df4.工单号.apply(lambda x:x[1:]).astype('str')
    df4.五级地址ID=df4.五级地址ID.astype('str')

    df4=df4[(df4['产品施工属性'] !='存量业务')]
            
    '''        
    dtype={
        '五级地址ID':str,
        'CRM业务流水号':str,
        }
    '''
    '''
    dtype={
        #'五级地址ID':str,
        #'CRM业务流水号':str,
        #'工单号':str,
        ' 客户原因缓装申请时间': 'category',
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
        '工单号': 'object',
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
    
    df5=pd.DataFrame(pd.read_csv(os.path.join(file_path,'服开数据更新'+strYesterday[5:]+'.csv'),engine='python',encoding='utf-8_sig',dtype=dtype))
    df5.CRM业务流水号=df5.CRM业务流水号.apply(lambda x:x[1:]).astype('str')
    '''
    dtype={
        '地址ID':str,
        }

    df6=pd.DataFrame(pd.read_excel(os.path.join('output','小区和网格关联关系.xls'),sheet_name='第1页',dtype=dtype))

    
    #读取前一天的当月指标，如果当天是周一，读上周五
    currentday = datetime.datetime.strptime(strToday, '%Y-%m-%d')
    if (currentday.weekday()>0):
        daybefore= currentday - datetime.timedelta(days = 1)    
    else:
        daybefore= currentday - datetime.timedelta(days = 3)

    daybefore_strTime = daybefore.strftime("%Y-%m-%d") 
    
    file_path=os.path.join('output\\'+daybefore_strTime,'当月指标.xlsx')
    df7=pd.DataFrame(pd.read_excel(file_path,sheet_name='市场兑换率（归档数）'))

    df8=pd.DataFrame(pd.read_excel(file_path,sheet_name='催装'))


    col_name = df3.columns.tolist()          # 将数据框的列名全部提取出来存放在列表里
    col_name.insert(3, '工单时长（24小时）')  
    col_name.insert(4, '工单时长（24小时）（8级加权）')           
    col_name.insert(5, '（8级加权）') 
    df3 = df3.reindex(columns=col_name)       # 整列都是NaN

    da6=df3[( df3['操作类型'] =='业务开通')&(df3['产品名称']=='家客开通')]
    da6.reset_index(drop=True,inplace=True)
    da6['五级地址ID']=da6['五级地址ID'].astype('str')


    #替换为
    right=df5.loc[:,['工单号','区域','代维']]
    da6=pd.merge(da6,right,left_on='开通工单号',right_on='工单号',how='left')
    da6.drop(['工单号'],axis=1,inplace=True)

    right=df6.loc[:,['地址ID','网格']]
    da6=pd.merge(da6,right,left_on='五级地址ID',right_on='地址ID',how='left')
    da6.drop(['地址ID'],axis=1,inplace=True)


    right=df6.loc[:,['用户班名称','网格']]
    right.rename(columns={'网格':'网格1'},inplace=True)
    right.drop_duplicates(subset=['用户班名称'],inplace=True)
    right.reset_index(drop=True,inplace=True)
    da6=pd.merge(da6,right,left_on='装维组/用户班',right_on='用户班名称',how='left')
    def getCell(df):
        if(pd.isnull(df['网格'])):
            result=df['网格1']
        else:
            result=df['网格']
        return result
    da6['网格']=da6.apply(getCell,axis=1)

    #调整列顺序
    da6_f=da6.区域
    da6_s=da6.代维
    da6_t=da6.网格
    da6.drop(['区域','代维','网格','用户班名称','网格1'],axis=1,inplace=True)
    da6.insert(3,'区域',da6_f)
    da6.insert(4,'代维',da6_s)
    da6.insert(5,'网格',da6_t)
    #替换结束


    '''需要替换
    left=da6.loc[:,['开通工单号']]
    right=df5.loc[:,['工单号','区域','代维']]
    pin=pd.merge(left,right,left_on='开通工单号',right_on='工单号',how='inner')
    for i in range(len(pin)):
        da6.loc[i,'区域']=pin.loc[i,'区域']
        da6.loc[i,'代维']=pin.loc[i,'代维']

        
    left=da6.loc[:,['五级地址ID']]       
    right=df6.loc[:,['地址ID','网格']]
    pin=pd.merge(left,right,left_on='五级地址ID',right_on='地址ID',how='inner')
    for i in range(len(da6)):
        for j in range(len(pin)):
            if da6.loc[i,'五级地址ID']==pin.loc[j,'地址ID']:
            da6.loc[i,'网格']=pin.loc[j,'网格']
            else:
                continue
        
    left=da6.loc[:,['装维组/用户班']]       
    right=df6.loc[:,['用户班名称','网格']]
    pin=pd.merge(left,right,left_on='装维组/用户班',right_on='用户班名称',how='inner')
    pin.drop_duplicates(subset=None, keep='first', inplace=True)
    pin.reset_index(drop=True,inplace=True)
    for i in range(len(da6)):
        for j in range(len(pin)):
            if (pd.isnull(da6.loc[i,'网格'])==True)&(da6.loc[i,'装维组/用户班']==pin.loc[j,'用户班名称']):
            da6.loc[i,'网格']=pin.loc[j,'网格']
            else:
                continue
        
    '''


    for i in range(len(da6)):
        da6.loc[i,'工单时长（24小时）']=(pd.to_datetime(da6.loc[i,'归档时间'])-pd.to_datetime(da6.loc[i,'派单时间'])).days*24+(pd.to_datetime(da6.loc[i,'归档时间'])-pd.to_datetime(da6.loc[i,'派单时间'])).seconds/3600   
        da6.loc[i,7]=((da6.loc[i,'工单时长（24小时）']-24)/24)
        if (da6.loc[i,7]>=2.0) & (da6.loc[i,7]<15.0):
            da6.loc[i,'（8级加权）']=1
        elif (da6.loc[i,7]>=15.0) & (da6.loc[i,7]<16.0):
            da6.loc[i,'（8级加权）']=4
        elif da6.loc[i,7]>=16.0: 
            da6.loc[i,'（8级加权）']=5
        else:
            continue
        
    #da=df7.append(da6,ignore_index=True,sort=False)
    #da.reset_index(drop=True,inplace=True)

    frames=[df7,da6]
    da=pd.concat(frames,sort=False)
    da.reset_index(drop=True,inplace=True)

    for i in range(len(da)):
        if(pd.isnull(da.loc[i,'24小时及时率'])==True):
            if (pd.isnull(da.loc[i,'首次预约上门时间'])==True):
                da.loc[i,'24小时及时率']=(pd.to_datetime(da.loc[i,'归档时间'])-pd.to_datetime(da.loc[i,'派单时间'])).days*24+(pd.to_datetime(da.loc[i,'归档时间'])-pd.to_datetime(da.loc[i,'派单时间'])).seconds/3600  
            else:
                da.loc[i,'24小时及时率']=(pd.to_datetime(da.loc[i,'归档时间'])-pd.to_datetime(da.loc[i,'首次预约上门时间'])).days*24+(pd.to_datetime(da.loc[i,'归档时间'])-pd.to_datetime(da.loc[i,'首次预约上门时间'])).seconds/3600  
        else:
            continue
            
    da.to_excel(excel_writer=writer,sheet_name='市场兑换率（归档数）', index=False)


    #催装
    right=df5.loc[:,['工单号','CRM业务流水号']]
    df3=pd.merge(df3,right,left_on='开通工单号',right_on='工单号',how='left')
    df3.drop(['工单号'],axis=1,inplace=True)

    df3.rename(columns={'CRM业务流水号':'crm业务流水号'},inplace=True) 


    '''需要替换
    col_name = df3.columns.tolist()          # 将数据框的列名全部提取出来存放在列表里
    col_name.insert(2, 'crm业务流水号')   
    df3 = df3.reindex(columns=col_name)       # 整列都是NaN


    left=df3.loc[:,['开通工单号']]
    right=df5.loc[:,['工单号','CRM业务流水号']]
    pin=pd.merge(left,right,left_on='开通工单号',right_on='工单号',how='inner')
    for i in range(len(pin)):
        df3.loc[i,'crm业务流水号']=pin.loc[i,'CRM业务流水号']
    '''  
        
    col_name = df4.columns.tolist()   # 将数据框的列名全部提取出来存放在列表里
    '''  需要替换
    col_name.insert(5, '区域')  
    col_name.insert(6, '网格')            
    col_name.insert(7, '结单时间')
    ''' 
    col_name.insert(5, '结单-催装')  
    col_name.insert(6, '预约上门时间')           
    col_name.insert(7, '预约-催单')  
    col_name.insert(15, '催办-派单')  
    df4 = df4.reindex(columns=col_name)       # 整列都是NaN


    right=df5.loc[:,['工单号','区域']]
    df4= pd.merge(df4,right,on=['工单号'],how='left')

    right=df6.loc[:,['地址ID','网格']]
    df4=pd.merge(df4,right,left_on='五级地址ID',right_on='地址ID',how='left')
    df4.drop(['地址ID'],axis=1,inplace=True)


    #调整列顺序
    df4_f=df4.区域
    df4_s=df4.网格
    df4.drop(['区域','网格'],axis=1,inplace=True)
    df4.insert(5,'区域',df4_f)
    df4.insert(6,'网格',df4_s)

    frames=[df8,df4]
    dc=pd.concat(frames,sort=False)
    dc.reset_index(drop=True,inplace=True)

    df3['crm业务流水号']= df3['crm业务流水号'].astype(str)
    dc['CRM业务流水号']= dc['CRM业务流水号'].astype(str)

    '''
    需要替换
    left=dc.loc[:,['CRM业务流水号']]       
    right=df3.loc[:,['crm业务流水号','归档时间']]
    pin=pd.merge(left,right,left_on='CRM业务流水号',right_on='crm业务流水号',how='inner')
    pin.drop_duplicates(subset=None, keep='first', inplace=True)
    pin.reset_index(drop=True,inplace=True)
    for i in range(len(dc)):
        for j in range(len(pin)):
            if (pd.isnull(dc.loc[i,'结单时间']))&(dc.loc[i,'CRM业务流水号']==pin.loc[j,'crm业务流水号']):
                dc.loc[i,'结单时间']=pin.loc[j,'归档时间']
            else:
                continue
    '''

    #替换为

    dc = pd.merge(dc,df3.loc[:,['crm业务流水号','归档时间']],left_on='CRM业务流水号',right_on='crm业务流水号',how='left')       
    def getJieDanShiJian(df):
        if(pd.isnull(df['结单时间'])):
            result=df['归档时间']
        else:
            result=df['结单时间']
        return result
    dc['结单时间']=dc.apply(getJieDanShiJian,axis=1)
    
    dc.drop(['crm业务流水号','归档时间'],axis=1,inplace=True)
    
    #替换结束
    
    



    '''  需要替换
    left=dc.loc[:,['工单号']]
    right=df5.loc[:,['工单号','区域']]
    pin=pd.merge(left,right,left_on='工单号',right_on='工单号',how='inner')
    for i in range(len(pin)):
        dc.loc[i,'区域']=pin.loc[i,'区域']


    left=dc.loc[:,['五级地址ID']]       
    right=df6.loc[:,['地址ID','网格']]
    pin=pd.merge(left,right,left_on='五级地址ID',right_on='地址ID',how='inner')
    for i in range(len(dc)):
        for j in range(len(pin)):
            if dc.loc[i,'五级地址ID']==pin.loc[j,'地址ID']:
            dc.loc[i,'网格']=pin.loc[j,'网格']
            else:
                continue
            
            
        
    left=dc.loc[:,['CRM业务流水号']]       
    right=df3.loc[:,['crm业务流水号','归档时间']]
    pin=pd.merge(left,right,left_on='CRM业务流水号',right_on='crm业务流水号',how='inner')
    for i in range(len(dc)):
        for j in range(len(pin)):
            if (pd.isnull(dc.loc[i,'结单时间'])==True)&(dc.loc[i,'CRM业务流水号']==pin.loc[j,'crm业务流水号']):
            dc.loc[i,'结单时间']=pin.loc[j,'归档时间']
            else:
                continue
            
    '''

    for i in range(len(dc)):
        dc.loc[i,'结单-催装']=(pd.to_datetime(dc.loc[i,'结单时间'])-pd.to_datetime(dc.loc[i,'催办发起时间'])).days*24+(pd.to_datetime(dc.loc[i,'结单时间'])-pd.to_datetime(dc.loc[i,'催办发起时间'])).seconds/3600
    k=len(df8)
    for k in range(len(dc)):      
        dc.loc[k,'催办-派单']=(pd.to_datetime(dc.loc[k,'催办发起时间'])-pd.to_datetime(dc.loc[k,'派单时间'])).days*24+(pd.to_datetime(dc.loc[k,'催办发起时间'])-pd.to_datetime(dc.loc[k,'派单时间'])).seconds/3600
            
        
    data=dc[dc['催办-派单'] >=12]
    data.reset_index(drop=True,inplace=True)
    data.to_excel(excel_writer=writer,sheet_name='催装', index=False)


    print('end')
    endtime = datetime.datetime.now()
    duringtime = endtime -  starttime
    print(duringtime.seconds)