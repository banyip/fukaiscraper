#date:xxxx-xx-xx
#function：合并指标数据和EXCEL公示模板
#######################################################
import xlrd
import xlutils.copy
import os
import xlwt
def excelMerge(output_folder_path):
    #打开一个workbook
    rb = xlrd.open_workbook('templates\\倍增模板.xlsx') 
    rb1 = xlrd.open_workbook(os.path.join(output_folder_path,'当月指标2.xlsx'))     
    wb = xlutils.copy.copy(rb)
    wb1 = xlutils.copy.copy(rb1)
    
    '''
    s_ws = rb.sheet_by_index(0)
    t_ws = wb1.add_sheet(s_ws.name)
    numRow = s_ws.nrows 
    numCol = s_ws.ncols 
    for col in range(numCol): 
        for row in range(numRow):
            cellObj = s_ws.cell(row,col)
            t_ws.write(row, col, xlwt.Formula("Sheet2!A1"))
    '''
    wb.save(os.path.join(output_folder_path,'倍增指标1.xls'))

    #利用保存时同名覆盖达到修改excel文件的目的,注意未被修改的内容保持不变
    

if __name__ == '__main__':
    excelMerge('output\\2021-01-11')