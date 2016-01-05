#!/usr/bin/python
# -*- coding: utf-8 -*-
'''
Created on 2016年1月5日

@author: lushangqi
'''

import xlrd,xlwt
import types
import os
import fnmatch


Debug = 0
Default = 0
def add_list(list1, list2, list1_head, list2_head):
    '''
    将两个列表相加
    1、如果列表表头含有关键字u'姓', u'名', u'次', u'号', u'分', u'卡'
    则不相加
    2、如果列表中的内容不是数字则不予相加，只有是数字 才相加
    '''
    ans = []
    ans.append(list1[0])
    if Debug:
        print 'list1'
        print list1
        print 'list2'
        print list2
    #如果是不可加的关键字
    word = [u'姓', u'名', u'次', u'号', u'分', u'卡']
    #如果那一行的内容是不可加的关键字，则返回list1列表
    for Str in word:
        if list1[0].find(Str) != -1:
            #print Str
            return list1
    
    for i in range(len(list1_head)):
        for j in range(len(list2_head)):
            if i == 0 or j == 0:
                continue
            if list1_head[i] == list2_head[j]:
                data1 = list1[i]
                data2 = list2[j]
                #print type(data1)
                if isinstance(list1[i], float) or isinstance(list1[i], int):
                    sum = 0.0
                    sum = data1+data2
                    #print sum
                    #gsum = "%g" %(sum)
                    #print type(gsum)
                    ans.append(sum)
                else:
                    ans.append(data1)
    return ans

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
    
    if not os.path.isfile(filename_main):
        print filename_main + "不存在！"
        return 
    if not os.path.isfile(filename_fu):
        print filename_fu + "不存在！"
        return 
    
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
                print '请输入y/n:'
                ss = raw_input()
            
            if ss[0] == 'y':
                if not Default:
                    print u'You choosed yes， 列相加'
                #将相加之后的结果放到列表中
                tmp_row=[]
                tmp_row= add_list(table_main.row_values(i), table_fu.row_values(j),\
                table_main.row_values(0), table_fu.row_values(0))
                if Debug:
                    print tmp_row
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
                                
def Union_Excel_Col(filename_main,filename_fu, \
                    newname = 'demo_col.xls', sheet_main = 0, sheet_fu = 0):
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
    
    if not os.path.isfile(filename_main):
        print filename_main + "不存在！"
        return 
    if not os.path.isfile(filename_fu):
        print filename_fu + "不存在！"
        return 
    
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
                print '请输入y/n:'
                ss = raw_input()#功能2、
            if ss[0] == 'y':
                if not Default:
                    print u'You choosed yes， 列相加'
                #将相加之后的结果放到列表中
                tmp_col= add_list(table_main.col_values(i), table_fu.col_values(j),\
                table_main.col_values(0), table_fu.col_values(0))
                if Debug:
                    print tmp_col
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

    newf.save(newname)       

def is_Excel(file):
    if not fnmatch.fnmatch(file, '*.xls') and not fnmatch.fnmatch(file, '*.xlsx'):
        return False
    else:
        return True 

def Multy_Excel_Merge(path):
    '''
    将一个文件夹下所有EXcel相加
    根据实际 需求实在列模式下进行
    默认列关键字 相等，即第一列的行主关键字相等
    '''
    all_File_num = 0
    global Debug
    if Debug:
        print path
    #该目录下所有文件的名字
    files = os.listdir(path)
    pre_file_name = ''
    pre_rm_name = ''
    newname = ''
    excel_cnt = 0
    for i in range(len(files)):
        if not is_Excel(files[i]):
            continue
        if fnmatch.fnmatch(files[i], 'demo*.xls'):
            continue
        if Debug:
            print files[i]
            print pre_file_name + '+' + files[i]
        excel_cnt = excel_cnt + 1
        if excel_cnt ==  1:
            pre_file_name = files[i]
            continue
        newname = "demo" + str(i) + ".xls"
        
        Union_Excel_Col(pre_file_name, files[i], newname)
        if i == len(files)-1:
            break;
        
        if excel_cnt >= 3:
            os.remove(os.path.join(path, pre_file_name))
        pre_file_name = newname
    #发现改名字会发生错误，所以还是不改了。。。
    #os.rename(os.path.join(path,newname), os.path.join(path, 'demo_col.xls'))
             

def Default_control():
    print u"是否默认将所有主键相等的行相加？"
    print u"y:之后默认相加"
    print u"n:之后 询问相加"
    global Default
    print u"请输入:y/n："
    ans_defaut = raw_input()
    if ans_defaut == "y":
        Default = 1
    else:
        Default = 0
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
    
    print u"模式3："
    print u"列模式：将一个文件夹里的所有Excel合并 （必须行主键（第一列）相等）"
    print u"     功能类似模式1和模式2"
    print u"可以选择 是否每行执行前 先询问，或者询问第一次后不再询问 "

    print u"请选择模式：1, 2, 3:"
    ans = raw_input()
    if ans[0] == "1":
        print u"行模式"
        print u"请输入主表文件名"
        f_m = raw_input()
        print u"请输入附表文件名:"
        f_t = raw_input()
        #询问第一次
        Default_control()
        Union_Excel_Row(f_m, f_t)
        print u"行模式结束"
    elif ans[0] == "2":
        print u"列模式"
        print u"请输入主表文件名"
        f_m = raw_input()
        print u"请输入附表文件名:"
        f_t = raw_input()
        #询问第一次
        Default_control()
        Union_Excel_Col(f_m, f_t)
        print u"列模式结束"
    elif ans[0] == '3':
        print u'''
        将一个目录下所有的EXcel合并
        生成demo_col.xls文件 
        默认将所有行主关键字相同对应的那一列数字相加
        如果列表表头含有关键字'姓', '名', '次', '号', '分', '卡'
    则不相加
    在命令行下只需将文件夹拖入黑框或直接输入
    '''
        print u'在命令行下将文件夹拖入黑框或直接输入文件夹名字'
        print 'Please Enter the name:'
        mypath = raw_input()
        #询问第一次
        Default_control()
        
        Multy_Excel_Merge(mypath)
        print "多Excel合并结束"
    
    

