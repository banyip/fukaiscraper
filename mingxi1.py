# -*- coding: utf-8 -*-
"""
Created on Tue Dec 15 15:23:06 2020

@author: Administrator
"""

import pandas as pd
import datetime
import time
import os

def mingxi1(file_path,strToday,strYesterday,writer,df5,pdew):
    starttime = datetime.datetime.now()
    

    df1=pd.DataFrame(pd.read_csv(os.path.join(file_path,'隔夜存量在途退单清单导出_佛山_'+strToday+'.csv'), engine='python'))
    #df1.工单号=df1.工单号.apply(lambda x:x[1:]).astype('str')
    df1['区域']=df1['区域'].map(lambda x: str(x)[:-1])

    df2=pd.DataFrame(pd.read_csv(os.path.join(file_path,'家客在途单_佛山_'+strYesterday+'.csv'), engine='python'))
    #df2.工单号=df2.工单号.apply(lambda x:x[1:]).astype('str')
    df2=df2.rename(columns={'区县':'区域'})
    df2['区域']=df2['区域'].map(lambda x: str(x)[:-1])

    '''
    dtype={
      '五级地址ID':str,
      'CRM业务流水号':str,
       '工单号':str,
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

    df5=pd.DataFrame(pd.read_csv(os.path.join(file_path,'服开数据更新'+strYesterday[5:]+'.csv'),dtype=dtype))
    

    df5.CRM业务流水号=df5.CRM业务流水号.apply(lambda x:x[1:]).astype('str')
    '''
    #df5.宽带帐号=df5.宽带帐号.astype('str')
    
    dtype={
        '地址ID':str,
        }

    df6=pd.DataFrame(pd.read_excel(os.path.join('output','小区和网格关联关系.xls'),sheet_name='第1页',dtype=dtype))

    #col_name = df1.columns.tolist()          # 将数据框的列名全部提取出来存放在列表里
    ''' 以下列通过关联获取，不需新增
    col_name.insert(7, '五级地址')  
    col_name.insert(8, '网格')           
    col_name.insert(9, '七级地址') 
    '''
    #df1 = df1.reindex(columns=col_name)       # 整列都是NaN


    da1=df1[(( df1['操作类型'] =='业务开通')|(df1['操作类型'] =='预勘查'))&(df1['产品名称']=='家客开通')&(df1['退单状态'] =='家客退单审批-容量审核')&(df1['派单时间'] >='2020-01-01 00:00:00')]
    da1.reset_index(drop=True,inplace=True)


    ''' 需要优化
    left=da1.loc[:,['工单号']]
    right=df5.loc[:,['工单号','五级地址ID','标准地址','五级地址名称']]
    pin=pd.merge(left,right,left_on='工单号',right_on='工单号',how='inner')
    for i in range(len(pin)):
        da1.loc[i,'七级地址']=pin.loc[i,'标准地址']
        da1.loc[i,'五级地址']=pin.loc[i,'五级地址ID']
        if pd.isnull(da1.loc[i,'五级地址名称'])==True:     
        da1.loc[i,'五级地址名称']=pin.loc[i,'五级地址名称']
    '''
    #以上语句为获得七级地址、五级地址和五级地址名称，其中七级地址和五级地址可以直接通过MERGE获得，五级地址名称通过merge之后apply获得：
    right=df5.loc[:,['工单号','五级地址ID','标准地址','五级地址名称']]
    right.rename(columns={'五级地址名称':'五级地址名称1'},inplace=True) 
    da1.drop(['五级地址ID'],axis=1,inplace=True)
    da1 = pd.merge(da1,right,on=['工单号'],how='left')
    da1.rename(columns={'五级地址ID':'五级地址','标准地址':'七级地址'},inplace=True) 
    def get5AddressName(df):
        if(pd.isnull(df['五级地址名称'])):
            result=df['五级地址名称1']
        else:
            result=df['五级地址名称']
        return result
    da1['五级地址名称']=da1.apply(get5AddressName,axis=1)
    da1.drop(['五级地址名称1'],axis=1,inplace=True)
    #替换结束




    '''需要优化
    left=da1.loc[:,['五级地址']]       
    right=df6.loc[:,['地址ID','网格']]
    pin=pd.merge(left,right,left_on='五级地址',right_on='地址ID',how='inner')
    for i in range(len(da1)):
        for j in range(len(pin)):
            if da1.loc[i,'五级地址']==pin.loc[j,'地址ID']:
            da1.loc[i,'网格']=pin.loc[j,'网格']
            else:
                continue
    '''

    #替换为
    right=df6.loc[:,['地址ID','网格']]
    da1=pd.merge(da1,right,left_on='五级地址',right_on='地址ID',how='left')
    da1.drop(['地址ID'],axis=1,inplace=True)



    '''
    left=da1.loc[:,['装维组名称']]       
    right=df6.loc[:,['用户班名称','网格']]
    pin=pd.merge(left,right,left_on='装维组名称',right_on='用户班名称',how='inner')
    for i in range(len(da1)):
        for j in range(len(pin)):
            if (pd.isnull(da1.loc[i,'网格'])==True)&(da1.loc[i,'装维组名称']==pin.loc[j,'用户班名称']):
            da1.loc[i,'网格']=pin.loc[j,'网格']
            else:
                continue
    '''    

    #替换为
    right=df6.loc[:,['用户班名称','网格']]
    right.rename(columns={'网格':'网格1'},inplace=True)
    right.drop_duplicates(subset=['用户班名称'],inplace=True)
    right.reset_index(drop=True,inplace=True)
    da1=pd.merge(da1,right,left_on='装维组名称',right_on='用户班名称',how='left')
    def getCell(df):
        if(pd.isnull(df['网格'])):
            result=df['网格1']
        else:
            result=df['网格']
        return result
    da1['网格']=da1.apply(getCell,axis=1)

    #调整列顺序
    da1_f=da1.五级地址
    da1_s=da1.网格
    da1_t=da1.七级地址
    da1.drop(['五级地址','网格','七级地址','用户班名称','网格1'],axis=1,inplace=True)
    da1.insert(7,'五级地址',da1_f)
    da1.insert(8,'网格',da1_s)
    da1.insert(9,'七级地址',da1_t)
    #替换结束
    for i in range(len(da1)):
        if (pd.isnull(da1.loc[i,'区域'])):
            da1.loc[i,'区域']=da1.loc[i,'七级地址'][3:5]
        else:
            continue
        
    for i in range(len(da1)):
        if (pd.isnull(da1.loc[i,'网格'])):
            da1.loc[i,'网格']=da1.loc[i,'七级地址'][8:10]
        else:
            continue
    #da1.to_excel(excel_writer=writer,sheet_name='扩容单量',index=False)         
    pdew.df2Sheet('扩容单量',da1)   
        
    da2=df1[(( df1['操作类型'] =='业务开通')|(df1['操作类型'] =='预勘查'))&(df1['产品名称']=='家客开通')&((df1['退单状态'] =='家客退单审批-攻坚岗')|(df1['退单状态'] =='家客退单审批-物业协调岗'))&(df1['派单时间'] >='2020-01-01 00:00:00')]
    da2.reset_index(drop=True,inplace=True)

    '''需要优化
    left=da2.loc[:,['工单号']]
    right=df5.loc[:,['工单号','五级地址ID','标准地址','五级地址名称']]
    pin=pd.merge(left,right,left_on='工单号',right_on='工单号',how='inner')
    for i in range(len(pin)):
        da2.loc[i,'七级地址']=pin.loc[i,'标准地址']
        da2.loc[i,'五级地址']=pin.loc[i,'五级地址ID']
        if pd.isnull(da2.loc[i,'五级地址名称'])==True:     
        da2.loc[i,'五级地址名称']=pin.loc[i,'五级地址名称']
    '''
    #替换为
    right=df5.loc[:,['工单号','五级地址ID','标准地址','五级地址名称']]
    right.rename(columns={'五级地址名称':'五级地址名称1'},inplace=True) 
    da2.drop(['五级地址ID'],axis=1,inplace=True)
    da2 = pd.merge(da2,right,on=['工单号'],how='left')
    da2.rename(columns={'五级地址ID':'五级地址','标准地址':'七级地址'},inplace=True) 
    da2['五级地址名称']=da2.apply(get5AddressName,axis=1)
    da2.drop(['五级地址名称1'],axis=1,inplace=True)
    #替换结束



    '''需要优化
    left=da2.loc[:,['五级地址']]       
    right=df6.loc[:,['地址ID','网格']]
    pin=pd.merge(left,right,left_on='五级地址',right_on='地址ID',how='inner')
    for i in range(len(da2)):
        for j in range(len(pin)):
            if da2.loc[i,'五级地址']==pin.loc[j,'地址ID']:
            da2.loc[i,'网格']=pin.loc[j,'网格']
            else:
                continue
    '''
    #替换为
    right=df6.loc[:,['地址ID','网格']]
    da2=pd.merge(da2,right,left_on='五级地址',right_on='地址ID',how='left')
    da2.drop(['地址ID'],axis=1,inplace=True)

    #替换结束




    '''需要优化
    left=da2.loc[:,['装维组名称']]       
    right=df6.loc[:,['用户班名称','网格']]
    pin=pd.merge(left,right,left_on='装维组名称',right_on='用户班名称',how='inner')
    for i in range(len(da2)):
        for j in range(len(pin)):
            if (pd.isnull(da2.loc[i,'网格'])==True)&(da2.loc[i,'装维组名称']==pin.loc[j,'用户班名称']):
            da2.loc[i,'网格']=pin.loc[j,'网格']
            else:
                continue       
    '''
    #替换为
    right=df6.loc[:,['用户班名称','网格']]
    right.rename(columns={'网格':'网格1'},inplace=True)
    right.drop_duplicates(subset=['用户班名称'],inplace=True)
    right.reset_index(drop=True,inplace=True)
    da2=pd.merge(da2,right,left_on='装维组名称',right_on='用户班名称',how='left')
    da2['网格']=da2.apply(getCell,axis=1)

    #调整列顺序
    da2_f=da2.五级地址
    da2_s=da2.网格
    da2_t=da2.七级地址
    da2.drop(['五级地址','网格','七级地址','用户班名称','网格1'],axis=1,inplace=True)
    da2.insert(7,'五级地址',da2_f)
    da2.insert(8,'网格',da2_s)
    da2.insert(9,'七级地址',da2_t)
    
    for i in range(len(da2)):
        if (pd.isnull(da2.loc[i,'区域'])):
            da2.loc[i,'区域']=da2.loc[i,'七级地址'][3:5]
        else:
            continue
        
    for i in range(len(da2)):
        if (pd.isnull(da2.loc[i,'网格'])):
            da2.loc[i,'网格']=da2.loc[i,'七级地址'][8:10]
        else:
            continue
    

    #da2.to_excel(excel_writer=writer,sheet_name='攻坚单量',index=False)       
    pdew.df2Sheet('攻坚单量',da2)       
        

    now_time = time.strftime('%Y-%m-%d',time.localtime(time.time()))+' 00:00:00' #今天的日期凌晨
    now=datetime.datetime.now()
    dayOfWeek = datetime.datetime.now().isoweekday() 
    days=datetime.timedelta(days=dayOfWeek-1) 
    oneday=(now-days).strftime("%Y-%m-%d")  +' 00:00:00' #周一凌晨

    #时间修改为整个月在
    da3=df1[(( df1['操作类型'] =='业务开通')|(df1['操作类型'] =='预勘查'))&(df1['产品名称']=='家客开通')&(df1['退单状态'] !='家客退单审批')&(df1['退单状态'] !='部分退单审批')&(df1['派单时间'] >= '2021-01-01')&(df1['派单时间'] < '2021-01-31' )]
    da3.reset_index(drop=True,inplace=True)
    #da3=da3.drop(['七级地址'],axis=1)

    '''需要优化
    left=da3.loc[:,['工单号']]
    right=df5.loc[:,['工单号','五级地址ID','五级地址名称']]
    pin=pd.merge(left,right,left_on='工单号',right_on='工单号',how='inner')
    for i in range(len(pin)):
        da3.loc[i,'五级地址']=pin.loc[i,'五级地址ID']
        if pd.isnull(da3.loc[i,'五级地址名称'])==True:     
        da3.loc[i,'五级地址名称']=pin.loc[i,'五级地址名称']
        
    left=da3.loc[:,['五级地址']]       
    right=df6.loc[:,['地址ID','网格']]
    pin=pd.merge(left,right,left_on='五级地址',right_on='地址ID',how='inner')
    for i in range(len(da3)):
        for j in range(len(pin)):
            if da3.loc[i,'五级地址']==pin.loc[j,'地址ID']:
            da3.loc[i,'网格']=pin.loc[j,'网格']
            else:
                continue

    left=da3.loc[:,['装维组名称']]       
    right=df6.loc[:,['用户班名称','网格']]
    pin=pd.merge(left,right,left_on='装维组名称',right_on='用户班名称',how='inner')
    for i in range(len(da3)):
        for j in range(len(pin)):
            if (pd.isnull(da3.loc[i,'网格'])==True)&(da3.loc[i,'装维组名称']==pin.loc[j,'用户班名称']):
            da3.loc[i,'网格']=pin.loc[j,'网格']
            else:
                continue        
    '''

    #替换为
    right=df5.loc[:,['工单号','五级地址ID','标准地址','五级地址名称']]
    right.rename(columns={'五级地址名称':'五级地址名称1'},inplace=True) 
    da3.drop(['五级地址ID'],axis=1,inplace=True)
    da3 = pd.merge(da3,right,on=['工单号'],how='left')
    da3.rename(columns={'五级地址ID':'五级地址','标准地址':'七级地址'},inplace=True) 
    da3['五级地址名称']=da3.apply(get5AddressName,axis=1)
    da3.drop(['五级地址名称1'],axis=1,inplace=True)


    right=df6.loc[:,['地址ID','网格']]
    da3=pd.merge(da3,right,left_on='五级地址',right_on='地址ID',how='left')
    da3.drop(['地址ID'],axis=1,inplace=True)


    right=df6.loc[:,['用户班名称','网格']]
    right.rename(columns={'网格':'网格1'},inplace=True)
    right.drop_duplicates(subset=['用户班名称'],inplace=True)
    right.reset_index(drop=True,inplace=True)
    da3=pd.merge(da3,right,left_on='装维组名称',right_on='用户班名称',how='left')
    da3['网格']=da3.apply(getCell,axis=1)

    #调整列顺序
    da3_f=da3.五级地址
    da3_s=da3.网格
    da3_t=da3.七级地址
    da3.drop(['五级地址','网格','七级地址','用户班名称','网格1'],axis=1,inplace=True)
    da3.insert(7,'五级地址',da3_f)
    da3.insert(8,'网格',da3_s)
    da3.insert(9,'七级地址',da3_t)
    #替换结束
    for i in range(len(da3)):
        if (pd.isnull(da3.loc[i,'区域'])):
            da3.loc[i,'区域']=da3.loc[i,'七级地址'][3:5]
        else:
            continue
        
    for i in range(len(da3)):
        if (pd.isnull(da3.loc[i,'网格'])):
            da3.loc[i,'网格']=da3.loc[i,'七级地址'][8:10]
        else:
            continue

    #da3.to_excel(excel_writer=writer,sheet_name='退单在途工单', index=False)      
    pdew.df2Sheet('退单在途工单',da3)   

        
    #col_name = df2.columns.tolist()          # 将数据框的列名全部提取出来存放在列表里
    '''
    col_name.insert(9, '五级地址')  
    col_name.insert(10, '网格')           
    '''
    #df2 = df2.reindex(columns=col_name)       # 整列都是NaN

    da4=df2[(( df2['操作类型'] =='开通')|(df2['操作类型'] =='预勘查'))&(df2['产品类型']=='家客开通')&(df2['订单状态'] =='服务开通中')]
    da4.reset_index(drop=True,inplace=True)

    '''需要替换
    left=da4.loc[:,['工单号']]
    right=df5.loc[:,['工单号','五级地址ID']]
    pin=pd.merge(left,right,left_on='工单号',right_on='工单号',how='inner')
    for i in range(len(pin)):
        da4.loc[i,'五级地址']=pin.loc[i,'五级地址ID']
        
    left=da4.loc[:,['五级地址']]       
    right=df6.loc[:,['地址ID','网格']]
    pin=pd.merge(left,right,left_on='五级地址',right_on='地址ID',how='inner')
    for i in range(len(da4)):
        for j in range(len(pin)):
            if da4.loc[i,'五级地址']==pin.loc[j,'地址ID']:
            da4.loc[i,'网格']=pin.loc[j,'网格']
            else:
                continue
            
    left=da4.loc[:,['装维组']]       
    right=df6.loc[:,['用户班名称','网格']]
    pin=pd.merge(left,right,left_on='装维组',right_on='用户班名称',how='inner')
    for i in range(len(da4)):
        for j in range(len(pin)):
            if (pd.isnull(da4.loc[i,'网格'])==True)&(da4.loc[i,'装维组']==pin.loc[j,'用户班名称']):
            da4.loc[i,'网格']=pin.loc[j,'网格']
            else:
                continue
    '''
    #替换为
    right=df5.loc[:,['工单号','五级地址ID']]
    da4.drop(['五级地址ID'],axis=1,inplace=True)
    da4 = pd.merge(da4,right,on=['工单号'],how='left')
    da4.rename(columns={'五级地址ID':'五级地址'},inplace=True) 


    right=df6.loc[:,['地址ID','网格']]
    da4=pd.merge(da4,right,left_on='五级地址',right_on='地址ID',how='left')
    da4.drop(['地址ID'],axis=1,inplace=True)


    right=df6.loc[:,['用户班名称','网格']]
    right.rename(columns={'网格':'网格1'},inplace=True)
    right.drop_duplicates(subset=['用户班名称'],inplace=True)
    right.reset_index(drop=True,inplace=True)
    da4=pd.merge(da4,right,left_on='装维组',right_on='用户班名称',how='left')
    da4['网格']=da4.apply(getCell,axis=1)

    #调整列顺序
    da4_f=da4.五级地址
    da4_s=da4.网格
    da4.drop(['五级地址','网格','用户班名称','网格1'],axis=1,inplace=True)
    da4.insert(9,'五级地址',da4_f)
    da4.insert(10,'网格',da4_s)
    #替换结束



    #da4.to_excel(excel_writer=writer,sheet_name='正常在途数', index=False)   
    pdew.df2Sheet('正常在途数',da4)
    
    da5=da4
    col_name = da5.columns.tolist()          # 将数据框的列名全部提取出来存放在列表里
    col_name.insert(6, '当天时间')  
    col_name.insert(7, '超7天')       
    da5 = da5.reindex(columns=col_name)       # 整列都是NaN
    
    right=df5.loc[:,['工单号','宽带帐号']]
    da5 = pd.merge(da5,right,on=['工单号'],how='left')
   
        #调整列顺序
    da5_f=da5.宽带帐号
    da5.drop(['宽带帐号'],axis=1,inplace=True)
    da5.insert(14,'宽带帐号',da5_f)
    #替换结束

    
    '''
    for i in range(len(da5)):
        da5.loc[i,'当天时间']=now_time
        pdsj = datetime.datetime.strptime(da5.loc[i,'派单时间'], "%Y-%m-%d %H:%M:%S") 
        dt = datetime.datetime.strptime(da5.loc[i,'当天时间'], "%Y-%m-%d %H:%M:%S") 
        da5.loc[i,'超7天']=(dt-pdsj).days
    '''
    def getOver7Days(df):
        now=datetime.datetime.now() #今天的日期凌晨        
        if(not pd.isnull(df['派单时间'])):
            result=(now-df['派单时间']).days
        else:
            result=''
        return result
    da5['超7天']=da5.apply(getOver7Days,axis=1)
    #da5.to_excel(excel_writer=writer,sheet_name='超长在途数', index=False)    
    pdew.df2Sheet('超长在途数',da5)




    print('end')
    endtime = datetime.datetime.now()
    duringtime = endtime -  starttime
    print(duringtime.seconds)