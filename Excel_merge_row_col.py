# -*- coding: utf-8 -*- 
'''
Created on 2016年1月2日

@author: lushangqi
'''


import xlrd,xlwt
Debug = 1
Default = 0

#行模式        
def Union_Excel_Row(filename_main = 'file1.xlsx',filename_fu = 'file2.xlsx', sheet_main = 0, sheet_fu = 0):
    '''
    行模式：
    filename_main:主表
    filename_fu:附表
    sheet_main：主表的工作表索行下标
    sheet_fu：附表的工作表索引行下标
          行模式下将两张表合并成一个新表
        1、将第二张表中没有在第一张表的那一行直接加进去
        2、将第二张表中在第一张表中存在的那一行加到主表中
                            每次操作前先询问是否相加
        
    ps: 默认两张excel是相同的
        默认两张表的第一行为行关键字，不予合并
    '''
    global Debug
    global Default
    #打开第一张表，默认获得该表的第一张工作簿
    data_main = xlrd.open_workbook(filename_main)
    table_main = data_main.sheets()[sheet_main]
    #打开第二张表，默认获得该表的第一张工作簿
    data_fu = xlrd.open_workbook(filename_fu)
    table_fu = data_fu.sheets()[sheet_fu]
    #第一张表的第一列（主键）
    file_main_head = table_main.col_values(0)
    #第二张表的第一列（主键）
    file_fu_head = table_fu.col_values(0)
    
    if Debug:
        print u'主表第一列： ',file_main_head
        print u'附表第一列： ',file_fu_head
    
    newf = xlwt.Workbook() #创建新表
    new_sheet = newf.add_sheet(u'Sheet1', cell_overwrite_ok = True)#创建sheet
    #处理第一张表
    row_cnt = 0 #新表的行数
    
    #找出在工作表2，不在工作表1中的行
    for i in range(len(file_main_head)):
        #ok=1,在第二张表中没有第一张表的主键
        ok = 1
        for j in range(len(file_fu_head)):#说明附表中存在相同主键的行
            if file_main_head[i] == file_fu_head[j]:
                ok = 0
                break
        if i == 0: #默认第一行为列关键字，不相加
            ok = 1
        if ok == 1: #说明只有主表存在该行，直接将该行写入到新表中
            tmp_row = table_main.row_values(i)
            for c in range(len(tmp_row)):
                new_sheet.write(row_cnt, c, tmp_row[c])
        else:
            ss = ""
            if Default: #功能3：默认所用主键相同的行都相加，不再第二次询问
                ss = "y"
            else:       #功能2：询问主键相同的行是否相加
                print u'是否将主表第%d行与附表第%d行相加？ '%(i+1, j+1);
                print u'该行表头名称：%s ' %file_main_head[i]            
                ss = raw_input('请输入y/n:')
            
            if ss[0] == 'y':
                print u'You choosed yes， 行相加'
                #将相加之后的结果放到列表中
                tmp_row=[]
                #每一行第一个元素为主键，不相加
                first = 1
                for c1 in range(len(table_main.row_values(i))):
                    if first == 0:
                        sum = 0.0
                        #sum = (table_main.row_values(i)[c1])+(table_fu.row_values(j)[c1])
                        #print sum
                        data1 = table_main.row_values(i)[c1]
                        data2 = table_fu.row_values(j)[c1]
                        #如果是数字则相加
                        if isinstance(data1, float):
                            sum = data1 + data2
                            #去掉小数点后的0
                            gsum = "%g" %(sum) 
                            tmp_row.append(str(gsum))
                        #否则将主表中的字符串写入新表
                        else:
                            #print data1
                            tmp_row.append((data1))
                    else:
                        tmp_row.append(file_main_head[i])
                    first = 0
                #将该列表写入到新表中
                for c in range(len(tmp_row)):
                    new_sheet.write(row_cnt, c, tmp_row[c])
            else:
                print u'将只写入主表的行数据, 附表的该行数据将被舍弃'
                tmp_row = table_main.row_values(i)
                for c in range(len(tmp_row)):    
                    new_sheet.write(row_cnt, c, tmp_row[c])
        #新表中新增一行
        row_cnt = row_cnt + 1

    #将出现在第二张表不在主表的行数据写入的新表中,将被添加到新表的后面
    for i in range(len(file_fu_head)):
        ok = 1
        for j in range(len(file_main_head)):
            if file_fu_head[i] == file_main_head[j]:
                ok = 0
                break;
        if ok == 1:
            tmp_row = table_fu.row_values(i)
            for c in range(len(tmp_row)):
                new_sheet.write(row_cnt, c, tmp_row[c])
            row_cnt = row_cnt + 1
    #保存新表
    newf.save('demo_row.xls')
    
def Union_Excel_Col(filename_main = 'file1.xlsx',filename_fu = 'file2.xlsx', sheet_main = 0, sheet_fu = 0):
    '''
    列模式：
    filename_main:主表
    filename_fu:附表
    sheet_main：主表的工作表索引下标
    sheet_fu：附表的工作表索引下标
        列模式下将两张表合并成一个新表
        1、将第二张表中没有在第一张表的直接加进去
        2、将第二张表中在第一张表中存在的那一列加到主表中
                            每次操作前先询问是否相加
        
    ps: 默认两张excel的行是相同的
        默认两张表的第一列为行关键字，不予合并
    
    '''
    global Debug
    global Default
    #打开第一张表，默认获得该表的第一张工作簿
    data_main = xlrd.open_workbook(filename_main)
    table_main = data_main.sheets()[sheet_main]
    #打开第二张表，默认获得该表的第一张工作簿
    data_fu = xlrd.open_workbook(filename_fu)
    table_fu = data_fu.sheets()[sheet_fu]
    #第一张表的第一行（主键）
    file_main_head = table_main.row_values(0)
    #第二张表的第一行（主键）
    file_fu_head = table_fu.row_values(0)
    
    if Debug:
        print u'主表第一行： ',file_main_head
        print u'附表第一行： ',file_fu_head
    
    newf = xlwt.Workbook() #创建新表
    
    new_sheet = newf.add_sheet(u'Sheet1', cell_overwrite_ok = True)#创建sheet
    #处理第一张表
    col_cnt = 0#新表的列数
    
    #找出在工作表2，不在工作表1中的行
    for i in range(len(file_main_head)):
        #ok=1,在第二张表中没有第一张表的主键
        ok = 1
        for j in range(len(file_fu_head)):#说明附表中存在相同主键的行
            if file_main_head[i] == file_fu_head[j]:
                ok = 0
                break
        if i == 0: #默认第一列为行关键字，不相加
            ok = 1
        if ok == 1:#说明只有主表存在该行，直接将该行写入到新表中
            tmp_col = table_main.col_values(i)
            for r in range(len(tmp_col)):
                new_sheet.write(r,col_cnt,tmp_col[r])
        else:
            ss = ""
            if Default: #功能3：默认所用主键相同的行都相加，不再第二次询问
                ss = "y"
            else:       #功能2：询问主键相同的行是否相加
                print u'是否将主表第%d行与附表第%d行相加？ '%(i+1, j+1);
                print u'该行表头名称：%s ' %file_main_head[i]            
                ss = raw_input('请输入y/n:')#功能2、
            if ss[0] == 'y':
                print u'You choosed yes， 列相加'
                #将相加之后的结果放到列表中
                tmp_col=[]
                #每一列第一个元素为主键，不相加
                first = 1
                for c1 in range(len(table_main.col_values(i))):
                    if first == 0:
                        sum = 0.0
                        #sum = (table_main.col_values(i)[c1])+(table_fu.col_values(j)[c1])
                        data1 = table_main.col_values(i)[c1]
                        data2 = table_fu.col_values(j)[c1]
                        #如果是数字则相加
                        if isinstance(data1, float):
                            sum = data1 + data2
                            #去掉小数点后的0
                            gsum = "%g" %(sum) 
                            tmp_col.append(str(gsum))
                        #否则将主表中的数字写入新表
                        else:
                            #print data1
                            tmp_col.append((data1))
                    else:
                        tmp_col.append(file_main_head[i])
                    first = 0
                for r in range(len(tmp_col)):
                    new_sheet.write(r,col_cnt,tmp_col[r])
            else:
                print u'将只写入主表的行数据, 附表的该行数据将被舍弃'
                tmp_col = table_main.col_values(i)
                for r in range(len(tmp_col)):
                    new_sheet.write(r,col_cnt,tmp_col[r])
        col_cnt = col_cnt + 1

              
    #将出现在第二张表不在主表的行数据写入的新表中,将被添加到新表的后面
    for i in range(len(file_fu_head)):
        ok = 1
        for j in range(len(file_main_head)):
            if file_fu_head[i] == file_main_head[j]:
                ok = 0
                break;
        if ok == 1:
            tmp_col = table_fu.col_values(i)
            for r in range(len(tmp_col)):
                new_sheet.write(r,col_cnt,tmp_col[r])
            col_cnt = col_cnt + 1
    #保存新表        
    newf.save('demo_col.xls')
    
if __name__ == '__main__':
    print u"合并EXCEL"
    print u"模式1："
    print u"行模式：将主表和附表中的行合并（必须列相主键（第一行）相等）"
    print u"     功能1：将主表不存在的行加到主表中"
    print u"     功能2：将两表表同时存在的行相加（每行执行前 先询问）"
    print u"     功能3：将两表将两表所有同时存在的行相加（只询问第一次）"
    
    print u"模式2："
    print u"列模式：将主表和附表中的列合并（必须行主键（第一列）相等）"
    print u"     功能1：将主表不存在的列加到主表中"
    print u"     功能2：将两表表同时存在的列相加（每行执行前 先询问）"
    print u"     功能3：将两表将两表所有同时存在的列相加（只询问第一次）"

    ans = raw_input(u"请选择模式：1/2:")
    if ans == "1":
        print u"行模式"
        print u"请输入主表文件名和附表文件名:"
        f_m = raw_input()
        f_t = raw_input()
        print u"是否默认将所有主键相等的行相加？"
        print u"y:之后默认相加"
        print u"n:之后 询问相加"
        ans_defaut = raw_input(u"请输入:y/n：")
        if ans_defaut == "y":
            Default = 1
        else:
            Default = 0
        Union_Excel_Row(f_m, f_t)
        print u"行模式结束"
    elif ans == "2":
        print u"列模式"
        print u"请输入主表文件名和附表文件名"
        f_m = raw_input()
        f_t = raw_input()
        print u"是否默认将所有主键相等的行相加？"
        print u"y:之后默认相加"
        print u"n:之后 询问相加"
        ans_defaut = raw_input(u"请输入:y/n：")
        if ans_defaut == "y":
            Default = 1
        else:
            Default = 0
        Union_Excel_Col(f_m, f_t)
        print u"列模式结束"
    
