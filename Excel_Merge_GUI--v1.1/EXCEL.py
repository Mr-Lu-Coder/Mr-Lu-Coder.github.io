#!/usr/bin/python
# -*- coding: utf-8 -*-
'''
Created on 2016年1月22日

@author: lushangqi
'''

import xlrd,xlwt
import os
import fnmatch
import CONFIGURE
import re
from PyQt4 import QtCore,  QtGui
import ConfigParser, codecs

#open_workbook
#Returns:
#An instance of the Book class

#返回该表有多少个工作表
def sheets_numbers(filename):
    ss = filename
    #转化成unicode编码
    #print type(ss)
    if not isinstance(filename,  unicode):
        ss = filename.decode('utf-8')
    #print type(ss)
    #print type(ss),ss
    if not os.path.isfile(ss):
        #print 'error'
        return CONFIGURE.ERROR
    #print ss
    book = xlrd.open_workbook(ss)
    return len(book.sheet_names())

#将主表的键值展示出来，如果是行模式，则返回第一行，如果是列模式，则返回第一列
#将该行或该列返回，如果失败则返回-1
def present_comboBox(filename, mode, sheet_index = 0):
    ss = filename
    #转化成unicode编码
    #print type(ss)
    if not isinstance(filename,  unicode):
        ss = filename.decode('utf-8')
    #print type(ss)
    #print type(ss),ss
    if not os.path.isfile(ss):
        return CONFIGURE.ERROR
    book = xlrd.open_workbook(filename)
    table = book.sheet_by_index(sheet_index)
    #如果是行模式返回第一行
    if CONFIGURE.DEBUG:
        print filename,  mode,  sheet_index
    if mode == 0:
        return table.row_values(0)
    #如果是列模式返回第一列
    elif mode == 1:
        return table.col_values(0)
        


def get_list_by_config():
    config = ConfigParser.ConfigParser()
    with codecs.open('WORD.ini', mode='r', encoding='utf-8-sig') as fp:
        config.readfp(fp)
        word_list = config.get('Exclude',  'word')
        #print word_list
        ans = []
        left = 0
        tmp = u''
        for i in word_list:
            if left == 1:
                if i == '\'' or i == '\"':
                    left = 0
                    ans.append(tmp)
                    tmp = u''
                    continue
                else:
                    tmp = tmp+i
            elif i == '\''or i == '\"':
                left = 1
                continue
            else:
                continue
                
        return ans
        


def add_list(list1, list2, list1_head, list2_head, main_key = 0):
    '''
    将两个列表相加
    1、如果列表表头含有关键字u'姓', u'名', u'次', u'号', u'分', u'卡'
    则不相加
    2、如果列表中的内容不是数字则不予相加，只有是数字 才相加
    '''
    ans = []###
    #ans.append(list1[0])
#    print 'list1'
#    print list1
#    print 'list2'
#    print list2
        
    #如果是不可加的关键字
    #word = [u'姓', u'名', u'次', u'号', u'分', u'卡']
    #如果那一行的内容是不可加的关键字，则返回list1列表
    #print WORD.word
    word_list = get_list_by_config()
    for Str in word_list:
        print Str
        if list1[main_key].find(Str) != -1:
            print 'find', list1[main_key],  Str
#            print main_key
            return list1
    
#    print 'main_key:',  main_key
    for i in range(len(list1_head)):
        if i == main_key:
            ans.append(list1[i])
            continue
        for j in range(len(list2_head)):
            if list1_head[i] == list2_head[j]:
#                print 'i', i,  'j',  j
                data1 = list1[i]
                data2 = list2[j]
                #print type(data1)
                if isinstance(list1[i], float) or isinstance(list1[i], int):#####！！！！
                    sum = 0.0
                    sum = data1+data2
                    #print sum
                    #gsum = "%g" %(sum)
                    #print type(gsum)
                    ans.append(sum)
#                    print 'sum',  sum, len(ans)
                else:
#                    print 'data1', data1,  len(ans)
                    ans.append(data1)
#    print 'ans',  ans,  len(ans)
    return ans

def Union_Excel_Row(filename_main = 'file1.xlsx',filename_fu = 'file2.xlsx', 
    main_key = 0,sheet_main = 0, sheet_fu = 0,  default = 0):
    '''
    行模式：
    filename_main:主表
    filename_fu:附表
    sheet_main：主表的工作表索行下标
    sheet_fu：附表的工作表索引行下标
    main_key:主键所在第一行、第几列，从0开始
          行模式下将两张表合并成一个新表
        1、将第二张表中没有在第一张表的那一行直接加进去
        2、将第二张表中在第一张表中存在的那一行加到主表中
                            每次操作前先询问是否相加
        
    ps: 默认两张excel的每一行的单元格数是相同的
        默认两张表的第一行为列关键字，不予合并
    '''
    
    #打开第一张表，默认获得该表的第一张工作簿
    data_main = xlrd.open_workbook(filename_main)
    table_main = data_main.sheets()[sheet_main]
    #打开第二张表，默认获得该表的第一张工作簿
    data_fu = xlrd.open_workbook(filename_fu)
    table_fu = data_fu.sheets()[sheet_fu]####判空
    
    
    #第一张表的第main_key列（主键）
    file_main_head = table_main.col_values(main_key)
    #第二张表的第main_key列（主键）
    file_fu_head = table_fu.col_values(main_key)
    
    if CONFIGURE.DEBUG:
        print u'主表各行主键： ',file_main_head
        print u'附表各行主键： ',file_fu_head
    
    newf = xlwt.Workbook() #创建新表，，，失败
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
            if default: #功能3：默认所用主键相同的行都相加，不再第二次询问
                ss = "y"
            else:    
                   #功能2：询问主键相同的行是否相加
                Strtext = u'是否将主表第%d行与附表第%d行相加？ '%(i+1, j+1)
                Strinforma = u'该行表头名称：%s ' %file_main_head[i]
                #apps = QtGui.QApplication(sys.argv)
                msgBox = QtGui.QMessageBox()
                msgBox.setWindowTitle(u'主键相同')
                msgBox.setText(Strtext)
                msgBox.setInformativeText(Strinforma)
                
                msgBox.setStandardButtons(QtGui.QMessageBox.Yes|QtGui.QMessageBox.No)
                msgBox.setDefaultButton(QtGui.QMessageBox.No)
                reply = msgBox.exec_()
                
                if reply == QtGui.QMessageBox.Yes:
                    ss = 'y'
                else:
                    ss = 'n'
            
            if ss[0] == 'y':
                #if not default:
                 #   print u'You choosed yes， 列相加'
                #将相加之后的结果放到列表中
                tmp_row=[]
                tmp_row= add_list(table_main.row_values(i), table_fu.row_values(j),\
                table_main.row_values(0), table_fu.row_values(0),  main_key)
                if CONFIGURE.DEBUG:
                    print tmp_row
                #将该列表写入到新表中
                for c in range(len(tmp_row)):
                    new_sheet.write(row_cnt, c, tmp_row[c])
            else:
                #print u'将只写入主表的行数据, 附表的该行数据将被舍弃'
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
    #print 'save'
    return newf


def Union_Excel_Col(filename_main, filename_fu, 
    main_key = 0, sheet_main = 0, sheet_fu = 0, default = 0):
    '''
    列模式：
    filename_main:主表
    filename_fu:附表
    sheet_main：主表的工作表索行下标
    sheet_fu：附表的工作表索引行下标
    main_key:主键所在第几行
          行模式下将两张表合并成一个新表
        1、将第二张表中没有在第一张表的那一列直接加进去
        2、将第二张表中在第一张表中存在的那一列加到主表中
                            每次操作前先询问是否相加
        
    ps: 默认两张excel的每一列的单元格数是相同的
        默认两张表的第一列为列关键字，不予合并
    '''
    
    #打开第一张表，默认获得该表的第一张工作簿
    data_main = xlrd.open_workbook(filename_main)
    table_main = data_main.sheets()[sheet_main]
    #打开第二张表，默认获得该表的第一张工作簿
    data_fu = xlrd.open_workbook(filename_fu)
    table_fu = data_fu.sheets()[sheet_fu]####判空
    
    
    #第一张表的第main_key列（主键）
    file_main_head = table_main.row_values(main_key)
    #第二张表的第main_key列（主键）
    file_fu_head = table_fu.row_values(main_key)
    
    if CONFIGURE.DEBUG:
        print u'主表各行主键： ',file_main_head
        print u'附表各行主键： ',file_fu_head
    
    newf = xlwt.Workbook() #创建新表，，，失败
    new_sheet = newf.add_sheet(u'Sheet1', cell_overwrite_ok = True)#创建sheet
    #处理第一张表
    col_cnt = 0 #新表的行数
    
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
        if ok == 1: #说明只有主表存在该列，直接将该列写入到新表中
            tmp_col = table_main.col_values(i)
            for r in range(len(tmp_col)):
                new_sheet.write(r,  col_cnt, tmp_col[r])
        else:
            ss = ""
            if default: #功能3：默认所用主键相同的行都相加，不再第二次询问
                ss = "y"
            else:    
                   #功能2：询问主键相同的行是否相加
                Strtext = u'是否将主表第%d列与附表第%d列相加？ '%(i+1, j+1)
                Strinforma = u'该列表头名称：%s ' %file_main_head[i]
                #apps = QtGui.QApplication(sys.argv)
                msgBox = QtGui.QMessageBox()
                msgBox.setWindowTitle(u'主键相同')
                msgBox.setText(Strtext)
                msgBox.setInformativeText(Strinforma)
                msgBox.setStandardButtons(QtGui.QMessageBox.Yes|QtGui.QMessageBox.No)
                msgBox.setDefaultButton(QtGui.QMessageBox.No)
                reply = msgBox.exec_()
                
                if reply == QtGui.QMessageBox.Yes:
                    ss = 'y'
                else:
                    ss = 'n'
            
            if ss[0] == 'y':
                #if not default:
                #   print u'You choosed yes， 列相加'
                #将相加之后的结果放到列表中
                tmp_col=[]
                tmp_col= add_list(table_main.col_values(i), table_fu.col_values(j),\
                table_main.col_values(0), table_fu.col_values(0),  main_key)
                if CONFIGURE.DEBUG:
                    print tmp_col
                #将该列表写入到新表中
                for r in range(len(tmp_col)):
                    new_sheet.write(r,  col_cnt, tmp_col[r])
            else:
                #print u'将只写入主表的行数据, 附表的该行数据将被舍弃'
                tmp_col = table_main.col_values(i)
                for r in range(len(tmp_col)):    
                    new_sheet.write(r,  col_cnt, tmp_col[r])
        #新表中新增一行
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
                new_sheet.write(r,  col_cnt,  tmp_col[r])
            col_cnt = col_cnt + 1
    #保存新表
    #print 'save'
    return newf

#返回该文件夹下excel的数量
def excel_numbers(path):
    files = os.listdir(path)
    
    number = 0
    for f in files:
        if re.match('~$[^~$]*', f):
            continue
        if re.match('[^.]*.xlsx*$', f):
            number = number + 1
    return number
def Union_Excel_directory_col(path,  default = 0,  main_key = 0):
    
    #该目录下所有文件的名字
    os.chdir(path)
    files = os.listdir(path)
    pre_file_name = ''
    excel_number = excel_numbers(path)
    if CONFIGURE.DEBUG:
        print 'excel_number:',  excel_number
    newname = ''
    excel_cnt = 0
    for i in range(len(files)):
        
        '不是excel文件'
        if not re.match('[^.]*.xlsx*$', files[i]):
            continue
        if fnmatch.fnmatch(files[i], 'demo*.xls'):
            continue
            
        if CONFIGURE.DEBUG:
            print files[i]
            print pre_file_name + '+' + files[i]
        #ss = raw_input()
        excel_cnt = excel_cnt + 1
        if excel_cnt ==  1:
            pre_file_name = files[i]
            continue
        newname = "demo" + str(i) + ".xls"
        
        
        
        newf = Union_Excel_Col(pre_file_name, files[i], 0, 0, 0, default)
        if excel_number == 2:
            return newf
        if excel_cnt >= 3:
            #print 'erase: ',  pre_file_name
            os.remove(os.path.join(path, pre_file_name))
        pre_file_name = newname
        
        if excel_cnt < excel_number:
            newf.save(newname)
        else:
            if CONFIGURE.DEBUG:
                print 'yes return!'
            return newf
            
            
def Union_Excel_directory_row(path,  default = 0,  main_key = 0):
    
    #该目录下所有文件的名字
    os.chdir(path)
    files = os.listdir(path)
    pre_file_name = ''
    excel_number = excel_numbers(path)
    if CONFIGURE.DEBUG:
        print 'excel_number:',  excel_number
    newname = ''
    excel_cnt = 0
    for i in range(len(files)):
        
        '不是excel文件'
        if not re.match('[^.]*.xlsx*$', files[i]):
            continue
        if fnmatch.fnmatch(files[i], 'demo*.xls'):
            continue
            
            
        if CONFIGURE.DEBUG:
            print files[i]
            print pre_file_name + '+' + files[i]
        #ss = raw_input()
        excel_cnt = excel_cnt + 1
        if excel_cnt ==  1:
            pre_file_name = files[i]
            continue
        newname = "demo" + str(i) + ".xls"
        
        
        
        newf = Union_Excel_Row(pre_file_name, files[i], 0, 0, 0, default)
        if excel_number == 2:
            return newf
        if excel_cnt >= 3:
            #print 'erase: ',  pre_file_name
            os.remove(os.path.join(path, pre_file_name))
        pre_file_name = newname
        
        if excel_cnt < excel_number:
            newf.save(newname)
        else:
            if CONFIGURE.DEBUG:
                print 'yes return!'
            return newf
        
        
