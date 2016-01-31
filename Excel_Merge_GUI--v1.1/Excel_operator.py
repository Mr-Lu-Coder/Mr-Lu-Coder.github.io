# -*- coding: utf-8 -*-

"""
Module implementing Excel_operator.
"""
import sys, os
from PyQt4 import QtCore,  QtGui
import xlrd,xlwt
from Ui_mainwindow import Ui_MainWindow
import EXCEL
import CONFIGURE

class Excel_operator(QtGui.QMainWindow, Ui_MainWindow):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """
        Constructor        
        @param parent reference to the parent widget
        @type QWidget

        """
        QtGui.QMainWindow.__init__(self, parent)
        self.setupUi(self)
        
        self.mode = -1
        self.label_mode.setText(u'模式：未选择')
        #文件1和2的初始化
        self.filename1 = ''
        self.filename2 = ''
        self.directory = ''
        self.Str_directoryname = u''
        #选项是否未完成
        self.mode_row_complete = 0
        self.mode_col_complete = 0
        self.mode_directory_complete = 0
        #切换至勾选
        self.checkBox.toggle()
        #默认是只询问第一次
        self.default = 1
        #设置图标
        self.setWindowIcon(QtGui.QIcon('icons/operator.png'))
        self.action_R.setIcon(QtGui.QIcon('icons/row.png'))
        self.action_L.setIcon(QtGui.QIcon('icons/column.png'))
        self.action_F.setIcon(QtGui.QIcon('icons/directory.png'))
        self.action_helpdoc.setIcon(QtGui.QIcon('icons/helpdoc.png'))
        self.action_about.setIcon(QtGui.QIcon('icons/about.png'))

        #默认第一个界面
        self.stackedWidget.setCurrentIndex(0)
        self.sheet_main = 0
        self.main_key = 0
        self.sheet_fu = 0
        #当前程序所在文件
        self.curdirectory = os.getcwd()
        
        #第二个界面
        #勾选文件模式行相等
        
        self.checkBox_row.toggle()
        #0为行相等，1为列相等
        self.Dmode = 1
        self.row_toggle = 1
        self.col_toggle = 0
        
        
        
        self.connect(self.checkBox, QtCore.SIGNAL('stateChanged(int)'),
            self.changestate)    
        
        self.connect(self.checkBox_row, QtCore.SIGNAL('stateChanged(int)'),
            self.changestate_row)
           
        self.connect(self.checkBox_col, QtCore.SIGNAL('stateChanged(int)'),
            self.changestate_col) 
        
        self.connect(self.comboBox_sheet_main_F, QtCore.SIGNAL('activated(QString)'),
            self.sheet_main_F)
        
        self.connect(self.comboBox_sheet_fu_F, QtCore.SIGNAL('activated(QString)'),
            self.sheet_fu_F)
            
        self.connect(self.comboBox_main_key_F, QtCore.SIGNAL('activated(QString)'),
            self.main_key_F)

         
      
        
    def sheet_main_F(self):
        self.sheet_main = self.comboBox_sheet_main_F.currentIndex()
        
        if self.mode == -1:
            self.textBrowser.append(u'您还没有选择模式！')
            return
            
        if not os.path.isfile(self.filename1.__str__()):
            self.textBrowser.append(u'请选择正确的主表文件！')
            return
        if CONFIGURE.DEBUG:
            print 'sheet_main:', self.sheet_main
        self.label_sheet_main_F.setText(u'主表第' + str(self.sheet_main+1) + u'张(默认第1张)')
        
        
        #更新主键
        #将主表的键值展示出来，如果是行模式，则加入第一行，如果是列模式，则加入第一列
        List = EXCEL.present_comboBox(self.filename1.__str__(),  self.mode,  self.sheet_main)
        if CONFIGURE.DEBUG:
                print List
        #首先清空
        self.comboBox_main_key_F.clear()
        #如果该列表元素是float型，则把小数点后末尾的0去掉
        #如果List为空，直接返回
        if len(List) == 0:
            self.textBrowser.append(u'您选择的表为空！')
            return
        import re
        for con in List:
            if CONFIGURE.DEBUG:
                print con, type(con)
            if isinstance(con,  float) and re.match(r'\d*\.0*$',  str(con)):
                gstr = re.match(r'\d*',  str(con)).group()
                self.comboBox_main_key_F.addItem(gstr)
                
            elif isinstance(con,  float):
                self.comboBox_main_key_F.addItem(str(con))
            else:
                self.comboBox_main_key_F.addItem((con))
        
        self.label_key_F.setText(u'您选择的主键是:'+self.comboBox_main_key_F.currentText().__str__() + u'(默认是第1个)')
        
            
    def sheet_fu_F(self):
        
        if self.mode == -1:
            self.textBrowser.append(u'您还没有选择模式！')
            return
        
        self.sheet_fu = self.comboBox_sheet_fu_F.currentIndex()
        if CONFIGURE.DEBUG:
            print 'sheet_fu:', self.sheet_fu
        self.label_sheet_fu_F.setText(u'副表第' + str(self.sheet_fu+1) + u'张(默认第1张)')
    
    def main_key_F(self):
        
        if self.mode == -1:
            self.textBrowser.append(u'您还没有选择模式！')
            return
        
        self.main_key = self.comboBox_main_key_F.currentIndex()
        if CONFIGURE.DEBUG:
            print 'main_key:', self.main_key
            print self.comboBox_main_key_F.currentText().__str__()
            #print self.comboBox_main_key_F.Text().__str__()
        if self.mode == 0:
            self.label_key_F.setText(u'您选择的主键是:'+self.comboBox_main_key_F.currentText().__str__() + u'所在的列')
        elif self.mode == 1:
            self.label_key_F.setText(u'您选择的主键是:'+self.comboBox_main_key_F.currentText().__str__() + u'所在的行')
            
            
    def changestate(self):
        if self.default == 0:
            self.default = 1
        else:
            self.default = 0
        if self.default == 0:
            self.textBrowser.append(u'每次相加都询问')
        elif self.default == 1:
            self.textBrowser.append(u'不再询问')
    def changestate_row(self):
        self.Dmode = self.Dmode^1
        if self.Dmode&1 and self.Dmode &2:
            self.textBrowser.append(u'行列模式请二选一！')
        elif self.Dmode&1:
            self.textBrowser.append(u'行模式！')
        elif self.Dmode&2:
            self.textBrowser.append(u'列模式！')
        else:
            self.textBrowser.append(u'行列模式请二选一！')
  
    def changestate_col(self):
        self.Dmode = self.Dmode^2
        if self.Dmode&1 and self.Dmode &2:
            self.textBrowser.append(u'行列模式请二选一！')
        elif self.Dmode&1:
            self.textBrowser.append(u'行模式！')
        elif self.Dmode&2:
            self.textBrowser.append(u'列模式！')
        else:
            self.textBrowser.append(u'行列模式请二选一！')
        
    def check_box_state_present(self):
        if self.Dmode == 0:
            print self.checkBox_row.checkState()
        else:
            print self.checkBox_col.checkState()
    
    @QtCore.pyqtSignature("")
    def on_bt_file1_clicked(self):
        """
        Slot documentation goes here.
        """
        if self.mode == -1:
            self.textBrowser.append(u'您还没有选择模式！')
            return
        self.filename1 = QtGui.QFileDialog.getOpenFileName(self,
            u"选择文件1", "/home", str("Excel Files (*.xlsx *.xls)"))
           
        #如果文件不存在，返回
        if not os.path.isfile(self.filename1.__str__()):
            self.textBrowser.append(u'请选择正确的主表文件！')
            return
        self.label_1.setText(u'您选择的文件：'+self.filename1)
        
        #得到主表有几个sheet，如果为0则返回为空
        num_main = EXCEL.sheets_numbers(self.filename1.__str__())
        if num_main == 0:
            self.textBrowser.append(u'您选择的表为空！')
            return
        
        
        print 'num_main:', num_main
        #self.textBrowser.append('sheet numbers:' + str(num_main))
        
        self.textBrowser.append(u'filename1:' + self.filename1)
        self.comboBox_sheet_main_F.clear()
        
        #主表第几张sheet，默认是第一张
        i = 1
        while i <= num_main:
            self.comboBox_sheet_main_F.addItem(str(i))
            i = i+1
        #将主表的键值展示出来，如果是行模式，则加入第一行，如果是列模式，则加入第一列
        List = EXCEL.present_comboBox(self.filename1.__str__(),  self.mode,  self.sheet_main)
        if CONFIGURE.DEBUG:
                print List
        print List
        #首先清空
        self.comboBox_main_key_F.clear()
        #如果该列表元素是float型，则把小数点后末尾的0去掉
        #如果List为空，直接返回
        if len(List) == 0:
            self.textBrowser.append(u'您选择的主表为空！')
            return
        import re
        for con in List:
            if CONFIGURE.DEBUG:
                print con, type(con)
            if isinstance(con,  float) and re.match(r'\d*\.0*$',  str(con)):
                gstr = re.match(r'\d*',  str(con)).group()
                self.comboBox_main_key_F.addItem(gstr)
                
            elif isinstance(con,  float):
                self.comboBox_main_key_F.addItem(str(con))
            else:
                self.comboBox_main_key_F.addItem((con))
        
        self.label_key_F.setText(u'您选择的主键是:'+self.comboBox_main_key_F.currentText().__str__() + u'(默认是第1个)')
        
        #第一个文件
        if self.mode == 0:
            self.mode_row_complete = self.mode_row_complete | 1
        elif self.mode == 1:
            self.mode_col_complete = self.mode_col_complete  | 1
        
    
    @QtCore.pyqtSignature("")
    def on_bt_file2_clicked(self):
        """
        Slot documentation goes here.
        """
        if self.mode == -1:
            self.textBrowser.append(u'您还没有选择模式！')
            return
        self.filename2 = QtGui.QFileDialog.getOpenFileName(self,
            u"选择文件2", "/home", str("Excel Files (*.xlsx *.xls)"))
           
        #如果文件不存在，返回
        if not os.path.isfile(self.filename2.__str__()):
            self.textBrowser.append(u'请选择正确的副表文件！')
            return
        self.label_2.setText(u'您选择的文件：'+self.filename2)
        
        #得到副表有几个sheet，如果为0则返回为空
        num_fu = EXCEL.sheets_numbers(self.filename2.__str__())
        if num_fu == 0:
            self.textBrowser.append(u'您选择的表为空！')
            return
        
        if CONFIGURE.DEBUG:
            print 'num_fu:', num_fu
        #self.textBrowser.append('sheet numbers:' + str(num_main))
        
        self.textBrowser.append(u'filename2:' + self.filename2)
        self.comboBox_sheet_fu_F.clear()
        
        #主表第几张sheet，默认是第一张
        i = 1
        while i <= num_fu:
            self.comboBox_sheet_fu_F.addItem(str(i))
            
            i = i+1
        #将副表的键值展示出来，如果是行模式，则加入第一行，如果是列模式，则加入第一列
#        List = EXCEL.present_comboBox(self.filename2.__str__(),  self.mode,  self.sheet_fu)
#      if CONFIGURE.DEBUG:
#                print List
                
                #不需要再选副表主键
#        #首先清空
#        self.comboBox_main_key_F.clear()
#        #如果该列表元素是float型，则把小数点后末尾的0去掉
#        #如果List为空，直接返回
#        if len(List) == 0:
#            self.textBrowser.append(u'您选择的表为空！')
#            return
#        import re
#        for con in List:
#            if CONFIGURE.DEBUG:
#                print con, type(con)
#            if isinstance(con,  float) and re.match(r'\d*\.0*$',  str(con)):
#                gstr = re.match(r'\d*',  str(con)).group()
#                self.comboBox_main_key_F.addItem(gstr)
#                
#            elif isinstance(con,  float):
#                self.comboBox_main_key_F.addItem(str(con))
#            else:
#                self.comboBox_main_key_F.addItem((con))
#        
#        self.label_key_F.setText(u'您选择的主键是:'+self.comboBox_main_key_F.currentText().__str__() + u'(默认是第1个)')
        
        #第二个文件
        if self.mode == 0:
            self.mode_row_complete = self.mode_row_complete | 2
        elif self.mode == 1:
            self.mode_col_complete = self.mode_col_complete | 2
        
    
    
    @QtCore.pyqtSignature("")
    def on_bt_directory_clicked(self):
        """
        Slot documentation goes here.
        """
        if self.mode == -1:
            self.textBrowser.append(u'您还没有选择模式！')
            return
        
        dialog = QtGui.QFileDialog()
        dialog.setFileMode(4) #仅显示目录
        dialog.setDirectory('/')
        dialog.setViewMode(QtGui.QFileDialog.List)
        ok = 0
        if (dialog.exec_()):
            filename = dialog.selectedFiles()
            ok = 1
        if ok and len(filename):
            self.directory = filename.first()
            if os.path.isdir(self.directory.__str__()):
                self.label_directory.setText(u'您选择的文件夹：'+self.directory.__str__())
                self.mode_directory_complete = 1
                self.Str_directoryname = self.directory.__str__()
                if CONFIGURE.DEBUG:
                    print 'yes'
                    print self.directory.__str__(),  type(self.directory.__str__())
            else:
                self.textBrowser.append(u'请选择正确的文件夹！')
    
    
    
    @QtCore.pyqtSignature("")
    def on_bt_func_R_clicked(self):
        """
        Slot documentation goes here.
        """
        self.mode_row_init()
    
    @QtCore.pyqtSignature("")
    def on_bt_func_L_clicked(self):
        """
        Slot documentation goes here.
        """
        self.mode_col_init()
    
    @QtCore.pyqtSignature("")
    def on_bt_func_D_clicked(self):
        """
        Slot documentation goes here.
        """
        self.mode_directory_init()
    
    
    @QtCore.pyqtSignature("")
    def on_bt_start_clicked(self):
        """
        Slot documentation goes here.
        """
        print 'mode', self.mode
        if self.mode == -1:
            self.textBrowser.append(u'您还没有选择模式！')
            return
        #如果条件不够则不能执行，并返回给必要的信息
        #print 'checksb', self.mode_directory_complete
        #print self.Str_directoryname
        if not self.check_can_start():
            print 'no'
            return
             
        if self.mode == 0:
            newf = EXCEL.Union_Excel_Row(self.filename1.__str__(),  self.filename2.__str__(),  self.main_key,  self.sheet_main, 
            self.sheet_fu,  self.default)
             
        elif self.mode == 1:
            newf = EXCEL.Union_Excel_Col(self.filename1.__str__(),  self.filename2.__str__(),  self.main_key,  self.sheet_main, 
            self.sheet_fu,  self.default)
        elif self.mode == 2:
            if CONFIGURE.DEBUG:
                print 'begin:default:',self.default
            if self.Dmode == 2:
                newf = EXCEL.Union_Excel_directory_col(self.Str_directoryname,  self.default)
            else:
                newf = EXCEL.Union_Excel_directory_row(self.Str_directoryname,  self.default)
        if not newf:
            self.textBrowser.append(u'合并失败！')
            return   
        fileName = QtGui.QFileDialog.getSaveFileName(self, u'保存路径', '/union_excel', selectedFilter='*.xls')
        if fileName:
            import re
            if not re.match('[^.]*.xls$', fileName.__str__()):
                self.textBrowser.append(u'只能保存为.xls格式的文件!')
                self.textBrowser.append(u'合并失败！')
            newf.save(fileName.__str__())
            self.textBrowser.append(u'合并成功！')
            self.textBrowser.append(u'已保存为' + fileName.__str__())
                 
                 
        print 'end'
        
        
    @QtCore.pyqtSignature("")
    def on_action_set_triggered(self):
        """
        Slot documentation goes here.
        """
        self.textBrowser.append(u'修改配置文件！保存后生效')
        import subprocess    
        filepath = os.path.join(self.curdirectory,  'WORD.ini')
        subprocess.Popen('notepad ' + filepath, shell=True)
    
    @QtCore.pyqtSignature("")
    def on_action_R_triggered(self):
        """
        Slot documentation goes here.
        """
        self.mode_row_init()
    
    @QtCore.pyqtSignature("")
    def on_action_L_triggered(self):
        """
        Slot documentation goes here.
        """
        self.mode_col_init()
    
    @QtCore.pyqtSignature("")
    def on_action_F_triggered(self):
        """
        Slot documentation goes here.
        """
        self.mode_directory_init()
    
    @QtCore.pyqtSignature("")
    def on_action_about_triggered(self):
        """
        Slot documentation goes here.
        """
        msgBox = QtGui.QMessageBox() 
        msgBox.setWindowTitle(u'关于')
        ss =u'''
Excel Operator
版本： 1.1
@Develpoed By Lushangqi
    '''
        msgBox.setText(ss)
        msgBox.exec_()
        
    @QtCore.pyqtSignature("")
    def on_action_helpdoc_triggered(self):
        """
        Slot documentation goes here.
        """
        msgBox = QtGui.QMessageBox() 
        SS = u'''
Excel Operator version 1.1
合并EXCEL
    行模式：
        将主表和附表中的行合并（必须列主键（第一行）相等）
        由用户选择哪一列中的每一个元素作为每一行的主键
        功能1：将主表不存在的行加到主表中
        功能2：将两表同时存在的行相加（每行执行前 先询问或者只询问第一次）
    
    列模式：
        将主表和附表中的列合并（必须行主键（第一列）相等）
        由用户选择哪一行中的每一个元素作为每一列的主键
        功能1：将主表不存在的列加到主表中
        功能2：将两表同时存在的列相加（每行执行前 先询问者只询问第一次）
    
    文件模式：将一个文件夹里的所有Excel合并 （必须行主键（第一列）相等）
        功能类似模式1和模式2
        
    tips：
        默认相加时，用户可以配置哪些主键不相加，比如含'姓'、'名'等字符，
        使程序在默认相加时对该行或该列之间不相加，修改时只需用户单击
        菜单栏‘配置修改’即可
    '''
        msgBox.setWindowTitle(u'帮助文档')
        msgBox.setText(SS)
        msgBox.exec_()
    
    #函数功能：检查是否可以开始运行程序
    #返回值：1,：可以运行  0：不可以运行
    def check_can_start(self):
        
        if self.mode == 0 or self.mode == 1:
            if not os.path.isfile(self.filename1.__str__()):
                self.textBrowser.append(u'请选择正确的主表文件！')
                return CONFIGURE.NO
            if not os.path.isfile(self.filename2.__str__()):
                self.textBrowser.append(u'请选择正确的副表文件！')
                return CONFIGURE.NO
            
        if self.mode == 0:
            if self.mode_row_complete == 3:
                return CONFIGURE.OK
            else:
                if self.mode_row_complete == 0:
                    self.textBrowser.append(u'请选择正确的文件！')
                elif self.mode_row_complete == 1:
                    self.textBrowser.append(u'请选择正确的副表文件！')
                elif self.mode_row_complete == 2:
                    self.textBrowser.append(u'请选择正确的主表文件！')
                return CONFIGURE.NO
        elif self.mode == 1:
            if self.mode_col_complete == 3:
                return CONFIGURE.OK
            else:
                if self.mode_col_complete == 0:
                    self.textBrowser.append(u'请选择正确的文件！')
                elif self.mode_col_complete == 1:
                    self.textBrowser.append(u'请选择正确的副表文件！')
                elif self.mode_col_complete == 2:
                    self.textBrowser.append(u'请选择正确的主表文件！')
                return CONFIGURE.NO
                
        elif self.mode == 2:
            if self.Dmode == 0 or self.Dmode == 3:
                self.textBrowser.append(u'行列模式请二选一！')
                return CONFIGURE.NO
            if not os.path.isdir(self.Str_directoryname):
                self.textBrowser.append(u'请选择正确的文件夹！')
                return CONFIGURE.NO
            number =  EXCEL.excel_numbers(self.Str_directoryname)
            #print 'number:', number
            if number == 0:
                self.textBrowser.append(u'文件夹下没有.xls或.xlsx文件！')
                return CONFIGURE.NO
            elif number == 1:
                self.textBrowser.append(u'文件夹下只有1个.xls或.xlsx文件！')
                return CONFIGURE.NO
            return CONFIGURE.OK
            
            
        
        
        
    def mode_row_init(self):
        self.mode = 0
        self.label_mode.setText(u'模式：行模式')
        self.textBrowser.clear()
        #切换到第一个界面
        self.stackedWidget.setCurrentIndex(0)
        self.filename1 = ''
        self.label_1.setText(u'您选择的文件：未选择')
        self.filename2 = ''
        self.label_2.setText(u'您选择的文件：未选择')
        
        self.comboBox_main_key_F.clear()
        self.label_key_F.setText(u'主键：第1行|列(默认第一个)')
        self.comboBox_sheet_main_F.clear()
        self.label_sheet_main_F.setText(u'主表第1张(默认第1张)')
        self.comboBox_sheet_fu_F.clear()
        self.label_sheet_fu_F.setText(u'副表第1张(默认第1张)')
        #选项未完成
        self.mode_row_complete = 0
        
        self.sheet_main = 0
        self.main_key = 0
        self.sheet_fu = 0
        if CONFIGURE.DEBUG:
            print 'mode:',  self.mode
            
        #！！！！！！
        self.set_comboBox_main_key_mode_init()
    
    def mode_col_init(self):
        self.mode = 1
        self.label_mode.setText(u'模式：列模式')
        self.textBrowser.clear()
        #切换到第一个界面
        self.stackedWidget.setCurrentIndex(0)
        
        self.filename1 = ''
        self.label_1.setText(u'您选择的文件：未选择')
        self.filename2 = ''
        self.label_2.setText(u'您选择的文件：未选择')
        
        self.comboBox_main_key_F.clear()
        self.label_key_F.setText(u'主键：第2行|列(默认第一个)')
        self.comboBox_sheet_main_F.clear()
        self.label_sheet_main_F.setText(u'主表第1张(默认第1张)')
        self.comboBox_sheet_fu_F.clear()
        self.label_sheet_fu_F.setText(u'副表第1张(默认第1张)')
        
        #选项未完成
        self.mode_col_complete = 0
        
        self.sheet_main = 0
        self.main_key = 0
        self.sheet_fu = 0
        if CONFIGURE.DEBUG:
            print 'mode:',  self.mode
        self.set_comboBox_main_key_mode_init()
        
    def mode_directory_init(self):
        self.textBrowser.clear()
        self.mode = 2
        self.label_mode.setText(u'模式：文件夹模式')
        self.mode_directory_complete = 0
        #切换到第二个界面
        self.stackedWidget.setCurrentIndex(1)
        self.directory = ''
        self.label_directory.setText(u'您选择的文件夹：未选择')
    
    def set_comboBox_main_key_mode_init(self):
        if os.path.isfile(self.filename1.__str__()):
            #将主表的键值展示出来，如果是行模式，则加入第一行，如果是列模式，则加入第一列
            List = EXCEL.present_comboBox(self.filename1.__str__(),  self.mode,  self.sheet_main)
            if CONFIGURE.DEBUG:
                print List
            #首先清空
            self.comboBox_main_key_F.clear()
            #如果该列表元素是float型，则把小数点后末尾的0去掉
            #如果List为空，直接返回
            if len(List) == 0:
                #self.textBrowser.append(u'您选择的主表为空！')
                return
            import re
            for con in List:
                if CONFIGURE.DEBUG:
                    print con, type(con)
                if isinstance(con,  float) and re.match(r'\d*\.0*$',  str(con)):
                    gstr = re.match(r'\d*',  str(con)).group()
                    self.comboBox_main_key_F.addItem(gstr)
                
                elif isinstance(con,  float):
                    self.comboBox_main_key_F.addItem(str(con))
                else:
                    self.comboBox_main_key_F.addItem((con))
        
            self.label_key_F.setText(u'您选择的主键是:'+self.comboBox_main_key_F.currentText().__str__() + u'(默认是第1个)')
    

if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    myapp = Excel_operator()
    myapp.show()
    sys.exit(app.exec_())


