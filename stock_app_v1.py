#!/usr/bin/python3
# -*- coding: utf-8 -*-

import sys
import pandas as pd
# to desable warnings
# pd.options.mode.chained_assignment = None  # default='warn'

from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QComboBox, QPushButton, QVBoxLayout, QHBoxLayout, QSizePolicy
from PyQt5.QtWidgets import QLineEdit, QTextEdit, QInputDialog, QSpinBox, QCheckBox, QButtonGroup, QProgressBar
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QFileDialog, QAction, QMainWindow#QToolBar #QStatusBar, QMenuBar
from PyQt5.QtCore import Qt #QCoreApplication QButtonGroup
from PyQt5.Qt import QIcon

import os.path

from math import fabs

# file with names of products and dicts of moduls
# FILE_OF_PRODUCTS = 'data.xlsx'
FILE_OF_PRODUCTS = 'data.csv'

# columns in this file = columns in data frame
COLUMN_PRODUCT_NAMES = 'наименование блока'
COLUMN_DICT_OF_MODULS = 'словарь модулей'


FILE_STOCK = 'Z:/Склад/Склад 14.01.16.xlsx'
# FILE_OF_SP_PLAT # Склад 14.01.16


#--------------------------------------------------------------
# dict prepearing for report (input dict from comboBox_list)

class DictMaker:
    def __init__(self, combobox_dict):
        self.combobox_dict = combobox_dict
        
    def _moduls_in_block(self,moduls_dict, q_blocks):
        return {k: v*q_blocks for k,v in moduls_dict.items()}        

    def _moduls_in_all_block(self,all_block):
        moduls_dict = {} 
        for d in all_block:
            for k,v in d.items():
                moduls_dict.update({k:v}) if k not in moduls_dict.keys() else moduls_dict.update(
                    {k:v + moduls_dict.get(k)})
        return moduls_dict
    
    def makeReportDict(self):
        all_block = [self._moduls_in_block(v[0], v[1]) for v in self.combobox_dict.values()] # list of moduls dicts and q-ties
        dict_of_moduls_for_report = self._moduls_in_all_block(all_block)
        return dict(sorted(dict_of_moduls_for_report.items()))
    
#--------------------------------------------------------------    

class ReportMaker:
    def __init__(self, report_name, report_option, report_dict, modul_df):
        self.report_name = report_name # self.checkBox_report_1 (names written near ceck boxes)
        self.report_option = report_option # self.checkBox_group.checkedButton() (selected option)
        
        # sorted dict from combo box
        self.report_dict = report_dict
        
        # modul_df = pd.read_excel('Z:/Склад/Склад 14.01.16.xlsx', sheet_name='Склад модулей(узлов)', usecols='C,F,G')
        self.modul_df = modul_df  
        
        
    def _makeReport_1(self):
#         if not self.modul_stock_isRead:            
#             self.readModulStock()
        if self.modul_df is not None:
            print('In progress...')
            filtered_modul_df = self.modul_df.loc[self.modul_df['Артикул'].isin(self.report_dict.keys())]
            print(self.report_dict.values())
            filtered_modul_df['moduls_in_order'] = self.report_dict.values()

            filtered_modul_df = filtered_modul_df.fillna(0)
            filtered_modul_df['q-ty of orders from moduls'] = filtered_modul_df[
                'Количество (в примечаниях история приходов и уходов)']//filtered_modul_df['moduls_in_order']
            filtered_modul_df['balance'] = filtered_modul_df[
                'Количество (в примечаниях история приходов и уходов)'] - filtered_modul_df['moduls_in_order']
            bad_balance_df = filtered_modul_df[filtered_modul_df['balance'] < 0]
            bad_balance_dict = {}
            for kv in bad_balance_df[['Артикул', 'balance']].values:    
                bad_balance_dict.update({int(kv[0]):fabs(kv[1])})
            
            print(f'bad_balance for {bad_balance_dict}') ################
#             self.progress_bar.setValue(90)
            return bad_balance_dict, bad_balance_df[['Артикул', 'balance']]
        else:
            print('File not found')
            return None , None       
        
        
    def identReport(self):
        if self.report_option == self.report_name:
            res = self._makeReport_1()
            print(f'RESULT_DICT: {res[0]}')
            print(f'RESULT_DF: {res[1]}')
        else:
            res = None, None
            print('Отчет не выбран') 
        return res

            
#--------------------------------------------------

# class ReportWindow(QWidget):
class ReportWindow(QMainWindow):
    def __init__(self, report_name, report_df):
        super().__init__() 
        
        self.report_df = report_df
        
        self.setWindowTitle(report_name)
#         self.label = QLabel(report_name)     
        self.table = QTableWidget()
        self.setCentralWidget(self.table)
        
# def fillTable
        cols = self.report_df.shape[1]
        rows = self.report_df.shape[0]

        self.table.setColumnCount(cols)
        self.table.setRowCount(rows) 
        
        
        [self.table.setHorizontalHeaderItem(i, QTableWidgetItem(report_df.columns[i])) 
         for i in range(cols)]
        
        [[self.table.setItem(i, j, QTableWidgetItem( str(report_df.iloc[i,j]) ))
         for j in range(cols)]
        for i in range(rows)]

        self.statusBar()
        
        saveFile = QAction(QIcon('save.png'),'Save As...', self)
        saveFile.setShortcut('Ctrl+Shift+S')
        saveFile.setStatusTip('Save data to excel file')
        saveFile.triggered.connect(self.showDialog)
        
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&Save data to yuor Excel File')
        fileMenu.addAction(saveFile)
#         self.toolbar = QToolBar('Save data to yuor Excel File')
#         self.addToolBar(self.toolbar)

        

#         layout = QVBoxLayout()
# #         layout.addWidget(self.label)
#         layout.addWidget(self.table)
#         self.setLayout(layout)
    
#     def closeEvent(self, event):
#         self.parent().child_closed()
#         event.accept()

    def showDialog(self):
        fname = QFileDialog.getSaveFileName(self, 'Open file', '/home')[0]
        print(fname)
        self.report_df.to_excel(f'{fname}.xlsx', index=False)


#--------------------------------------------------



class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        
        self.w = None # No external window yet
        
        self.valueLabel = QLabel()
        self.valueLabel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
      
        # Load data from Excel file (data is data about blocks)
        try:
# нужно задать путь для чтения, читать только 2 правильных столбца, отсортировать
#             self.cond = (by=COLUMN_PRODUCT_NAMES, ascending=False)
            self.data = pd.read_csv(FILE_OF_PRODUCTS).sort_values(by=COLUMN_PRODUCT_NAMES, ascending=False) 
            self.data_search = self.data
            self.valueLabel.setStyleSheet('color:blue;')            
            self.valueLabel.setText(f"Количество найденных результатов: {self.data_search.shape[0]}")        
        except:
            self.data = None
            self.data_search = None
            self.valueLabel.setStyleSheet('color:red;')
            self.valueLabel.setText("Файл не найден")

            
        self.block_list_dict = {} #comboBox dict
        self.msg = ''
        
        self.modul_stock_isRead = False
        self.modul_df = None
        
        self.initUI()


    def initUI(self):
        # Set window properties
        self.setWindowTitle("Приложение СКЛАД")
        self.setMinimumWidth(500)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Create widgets for loading data
        
#         # search
#         self.label_search = QLabel('Поиск по списку из файла', self)
#         self.textbox_search = QLineEdit(self)  
#         self.textbox_search.setPlaceholderText("Enter your text")
        
#         # add SEARCH BOTTON
#         self.searchButton = QPushButton("Найти")
#         self.searchButton.clicked.connect(self.searchBlock)
#         self.searchButton.clicked.connect(self.refresh_comboBox)        
        

        # combo box for loaded dicts of products
        self.label = QLabel("Выберите изделие:")
        self.comboBox = QComboBox()
        self.comboBox.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.comboBox.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.comboBox.setMinimumContentsLength(15)
        self.comboBox.setEditable(True)  
        
        # Add items to combo box
        self._refresh_comboBox()

        # CHOICE AND SEARCH
        self.comboBox.activated.connect(self.updateValue)


        # Create widgets for inserting data        - ПОКА НЕКРАСИВЫЕ 
        self.label1 = QLabel('наименование блока:', self)
        self.textbox1 = QLineEdit(self)
        
        self.label2 = QLabel('словарь модулей:', self)
#         self.textbox2 = QLineEdit(self)
        self.textbox2 = QTextEdit(self)

        self.label3 = QLabel('количество блоков:', self)
#         self.textbox3 = QLineEdit(self)
        self.textbox3 = QSpinBox()
        self.textbox3.setMinimum(1)

        

        # ADD BOTTON
        self.addButton = QPushButton("Добавить в список")
        self.addButton.clicked.connect(self.addBlock)      #(self.textbox1, self.textbox2, self.textbox3)
        
        self.to_fileButton = QPushButton("Добавить в файл")
        self.to_fileButton.clicked.connect(self.to_fileBlock)      #(self.textbox1, self.textbox2, self.textbox3)        
        

        # combo box with selected products and quantities        
        self.label_list = QLabel("Выбранные блоки:")
        self.comboBox_list = QComboBox()
        self.comboBox_list.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.comboBox_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.comboBox_list.setMinimumContentsLength(15)
#         self.comboBox_list.setEditable(True)
        self.comboBox_list.addItem('') # add empty string 
        # CHOICE
        self.comboBox_list.activated.connect(self.editBlock)        

#         # EDIT BOTTON
#         self.editButton = QPushButton("Редактировать выбранное")
#         self.editButton.clicked.connect(self.editBlock)          
        
        # REMOVE BOTTON    
        self.removeButton = QPushButton("Убрать из списка")
        self.removeButton.clicked.connect(self.removeBlock)         

        # chekBoxes to choice the report
        
        # Report_1 получить словарь модулей, которые надо изготовить {арт модуля:кол-во}
        self.checkBox_report_1 = QCheckBox('Report_1') 
        self.checkBox_report_2 = QCheckBox('Report_2')
        
        self.checkBox_group = QButtonGroup()
#         self.checkBox_group.setExclusive(True)
        self.checkBox_group.addButton(self.checkBox_report_1)
        self.checkBox_group.addButton(self.checkBox_report_2)
        
        
        # GET REPORT BOTTON    
        self.reportButton = QPushButton("Получить отчет")
        self.reportButton.clicked.connect(self.getReport)      
        
        # PROGRESS BAR
        self.progress_bar = QProgressBar(self)
        
        
        # EXIT BOTTON
        self.exitButton = QPushButton("Выйти")
        self.exitButton.clicked.connect(self.exitApp)
#         self.exitButton.clicked.connect(QCoreApplication.instance().quit)



        # Create layout
#         hbox_search = QHBoxLayout()
#         hbox_search.addWidget(self.label_search)
#         hbox_search.addWidget(self.textbox_search)        
    
        hbox = QHBoxLayout()        
        hbox.addWidget(self.label)
        hbox.addWidget(self.comboBox)
        
        hbox1 = QHBoxLayout()
        hbox1.addWidget(self.label1)
        hbox1.addWidget(self.textbox1)
        
        hbox2 = QHBoxLayout()
        hbox2.addWidget(self.label2)
        hbox2.addWidget(self.textbox2)
        
        hbox3 = QHBoxLayout()
        hbox3.addWidget(self.label3)
        hbox3.addWidget(self.textbox3)
        
        hbox_list = QHBoxLayout()
        hbox_list.addWidget(self.label_list)
        hbox_list.addWidget(self.comboBox_list)        
        
        vbox = QVBoxLayout()
#         vbox.addWidget(self.label_search)
#         vbox.addWidget(self.textbox_search)
#         vbox.addWidget(self.searchButton)
        vbox.addWidget(self.valueLabel)        # widget with selected data_name
        vbox.addLayout(hbox)
        vbox.addLayout(hbox1)
        vbox.addLayout(hbox2)
        vbox.addLayout(hbox3)
        vbox.addWidget(self.addButton)
        vbox.addWidget(self.to_fileButton)
        vbox.addLayout(hbox_list)              # widget with list selected data_names (blocks)
#         vbox.addWidget(self.editButton)        
        vbox.addWidget(self.removeButton)
    
        vbox.addWidget(self.checkBox_report_1)
        vbox.addWidget(self.checkBox_report_2)  
        
        vbox.addWidget(self.reportButton)
        
        vbox.addWidget(self.progress_bar)
        
        vbox.addWidget(self.exitButton)
        

                

        # Set layout
        self.setLayout(vbox)
   
        
        self.show()

        
        
        # ----------------------------------------------------------------
        # functions for loading data   
        

    def _refresh_comboBox(self): # refresh loaded list of products under SEARCH        
        self.comboBox.clear()
        self.comboBox.addItem('Обновить поиск')
        if self.data_search is not None and self.data_search.shape[0]==1:
            self.data_search = self.data
        if self.data_search is not None and not self.data_search.empty:
            for row in range(len(self.data_search.index)):
                self.comboBox.addItem(str(self.data_search.iloc[row, 0]))  
                

    def _searchBlock(self):
        if self.data_search is not None:
            self.data_search = self.data.loc[
#                 self.data[COLUMN_PRODUCT_NAMES].str.find(str(self.textbox_search.displayText())) != -1]
               self.data[COLUMN_PRODUCT_NAMES].str.find(str(self.comboBox.currentText())) != -1]                
            self.valueLabel.setStyleSheet('color:blue;')
            self.msg = f"Количество найденных результатов: {self.data_search.shape[0]}"
            return self.data.loc[self.data[COLUMN_PRODUCT_NAMES] == self.comboBox.currentText()] 
        else:
            return None
            # good info print
#         print(self.data_search.shape, self.data_search.shape[0], len(self.data_search.index), self.data_search.index)


    def _get_value(self, combo_idx, search_result):
        
        if self.data_search is None:
            self.valueLabel.setStyleSheet('color:red;')
            return "Файл не найден или пуст"

        if self.data_search.shape[0]:
            self.textbox1.setText('')
            self.textbox2.setText('')
#             res = self.msg
            if not search_result.empty: # key_string in list of products
                self.valueLabel.setStyleSheet('color:green;')
                self.textbox1.setText(search_result.iloc[0][COLUMN_PRODUCT_NAMES])
                self.textbox2.setText(search_result.iloc[0][COLUMN_DICT_OF_MODULS])
#                 if search_result.shape[0] == 1:
#                     self.msg = search_result.iloc[0][COLUMN_PRODUCT_NAMES]
            res = self.msg
            self._refresh_comboBox()
            return res  
        
        # 0 results were found
        else:
            res = 'Поиск обновлен' if combo_idx == 0 else "Выбранная строка не найдена"
            self.textbox1.setText('')
            self.textbox2.setText('')
            self.data_search = self.data
            self._refresh_comboBox()      
            return res              
 
    
    def updateValue(self): 
        
        combo_idx = self.comboBox.currentIndex()
        search_result = self._searchBlock()
        search_result
        self._refresh_comboBox()
        value = self._get_value(combo_idx, search_result)
        self.valueLabel.setText(value)
        

    # ADD product to list of products (dict with indexes for combobox) and add to comboBox_list
    def addBlock(self): # self.block_list_dict = {}
        
        self.msg = ''

        if self.moduls_dict_validation():
            self.block_name_validation() # creates name, if name is None or ''
            i = 0   # index
            self.block_list_dict[self.textbox1.displayText()] = [
#                     eval(self.textbox2.displayText()), 
                    eval(self.textbox2.toPlainText()),
                    self.textbox3.value(),
                    i]
            self.comboBox_list.clear()      
            # rewrite comboBox_list by values from block_list_dict
            for k, v in self.block_list_dict.items():
                self.comboBox_list.addItem(f'{v[1]}  {k}')
                self.block_list_dict[k] = [v[0], v[1], i]  # update dict with indexes from combobox to have points to connection
                i+=1
                self.valueLabel.setStyleSheet('color:green;')
                self.valueLabel.setText(self.msg)
        else:
            self.valueLabel.setStyleSheet('color:red;')
            self.valueLabel.setText(self.msg)                    

         
            
            
    def to_fileBlock(self):           
        self.msg = ''
        
        if self.moduls_dict_validation():
            self.block_name_validation()    
#             self.valueLabel.setStyleSheet('color:green;')

            add_df = pd.DataFrame({
                COLUMN_PRODUCT_NAMES: [self.textbox1.displayText()],
                COLUMN_DICT_OF_MODULS: [self.textbox2.toPlainText()]
            })
    
            if os.path.exists(FILE_OF_PRODUCTS):                
                if (
                    self.data_search is not None and 
                    self.textbox1.displayText() not in self.data_search[COLUMN_PRODUCT_NAMES].to_list()
                ):
                    add_df.to_csv(FILE_OF_PRODUCTS, mode='a', index = False, header=None)
                    self.data=pd.concat([self.data, add_df], ignore_index=True).sort_values(by=COLUMN_PRODUCT_NAMES, ascending=False)
                    self.msg = 'Добавлено!'
                else: 
                    self.msg = 'Уже есть. Нужно изменить название'
            else:
                add_df.to_csv(FILE_OF_PRODUCTS, mode='a', index = False)
                self.data = add_df
                self.msg = f'Файл {FILE_OF_PRODUCTS} создан. Запись добавлена!'
                
            self.valueLabel.setStyleSheet('color:green;')    
            self.valueLabel.setText(self.msg)                
                
        else:
            self.valueLabel.setStyleSheet('color:red;')
            self.valueLabel.setText(self.msg)          
            
            
    def editBlock(self):
        combo_idx = self.comboBox_list.currentIndex()
        for k, v in self.block_list_dict.items():
            if v[2] == combo_idx:                
                self.textbox1.setText(k)
                self.textbox2.setText(f'{v[0]}') # set str-type, not dict-type
                self.textbox3.setValue(v[1])
        
    def removeBlock(self):
        combo_idx = self.comboBox_list.currentIndex()
        for k, v in self.block_list_dict.items():
            if v[2] == combo_idx:        
                self.block_list_dict.pop(k)
                break
        
        self.comboBox_list.clear()
        i=0
        for k1, v1 in self.block_list_dict.items():
            self.comboBox_list.addItem(f'{v1[1]}  {k1}')
            self.block_list_dict[k1] = [v1[0], v1[1], i]  # update dict with indexes from combobox to have points to connection
            i+=1

      
    def moduls_dict_validation(self):
        
#         toPlainText()
#         string_moduls_dict = self.textbox2.displayText()
        string_moduls_dict = self.textbox2.toPlainText()
#         print(string_moduls_dict.startswith('{'))
        if not string_moduls_dict.startswith('{'):
            string_moduls_dict = '{' + string_moduls_dict
#         print(string_moduls_dict.startswith('{'))
        if not string_moduls_dict.endswith('}'):
            string_moduls_dict = string_moduls_dict +'}'
            
        try:
            moduls_dict = eval(string_moduls_dict)
            if (isinstance(moduls_dict, dict) 
                    and all(isinstance(k, int) 
                            and (isinstance(v, int) or isinstance(v, float)) for k, v in moduls_dict.items())):
                self.msg += " moduls_dict_validation OK."
                self.textbox2.setText(f'{moduls_dict}')
                return True
        except:
            self.msg += 'Приведите словарь к виду: {1: 5, 2: 6} или 1:5,2:6' 

            
    def block_name_validation(self):
        if self.textbox1.displayText() is None or self.textbox1.displayText() == '': 
#             moduls_dict = eval(self.textbox2.displayText())
            moduls_dict = eval(self.textbox2.toPlainText())
            name_str = [str(k) for k,v in moduls_dict.items()]
            name_str = '_'.join(name_str)
            self.textbox1.setText(name_str)
            
            
#--------------------

    def makeReport_dict(self):
        self.progress_bar.setValue(5)
#         print(self.block_list_dict)
        return DictMaker(self.block_list_dict).makeReportDict()


    def readModulStock(self):
        self.modul_stock_isRead = True
        try:
            print('--1')
            self.modul_df = pd.read_excel(FILE_STOCK, sheet_name='Склад модулей(узлов)', usecols='C,F,G')
            print('--2')
            self.modul_df = self.modul_df.iloc[2:]
            self.progress_bar.setValue(10)
            print('File OK')
        except:
#             self.modul_stock_isRead
            print('File not found')
            pass

    
#     def makeReport_1(self, report_dict):
#         if self.modul_df is not None:
#             print('In progress...')
#             filtered_modul_df = self.modul_df.loc[self.modul_df['Артикул'].isin(report_dict.keys())]
#             print(report_dict.values())
#             filtered_modul_df['moduls_in_order'] = report_dict.values()

#             filtered_modul_df = filtered_modul_df.fillna(0)
#             filtered_modul_df['q-ty of orders from moduls'] = filtered_modul_df[
#                 'Количество (в примечаниях история приходов и уходов)']//filtered_modul_df['moduls_in_order']
#             filtered_modul_df['balance'] = filtered_modul_df[
#                 'Количество (в примечаниях история приходов и уходов)'] - filtered_modul_df['moduls_in_order']
#             bad_balance_df = filtered_modul_df[filtered_modul_df['balance'] < 0]
#             bad_balance_dict = {}
#             for kv in bad_balance_df[['Артикул', 'balance']].values:    
#                 bad_balance_dict.update({int(kv[0]):fabs(kv[1])})
            
#             print(f'bad_balance for {bad_balance_dict}') ################
#             self.progress_bar.setValue(90)
#             return bad_balance_dict
#         else:
#             print('File not found')
#             return None

           
#     def getReport(self, checked):
    def getReport(self):
        
        if self.w is not None:
            self.w.close()
            self.w = None # Discard reference to ReportWindow            
        self.progress_bar.reset()    
#         self.progress_bar.setValue(0)

#         print('--progress_bar.reset')
        
        if not self.modul_stock_isRead:            
            self.readModulStock()  
                
        report_dict = self.makeReport_dict()
        print(f'report_dict is: {report_dict}')

#         ReportMaker !!!!!!!!!!!! тут еще доделать - есть траблы с разными репортами - как их подавать
# в вариантах сравнение репортов проводить тут, а в класс подавать уже извесный репорт, а не определять его там
        report_dict_res, report_dict_df = ReportMaker(
            self.checkBox_report_1, 
            self.checkBox_group.checkedButton(), 
            report_dict, 
            self.modul_df).identReport()
        print(f'report_dict_res is : {report_dict_res}')
        report_name = 'Report_1' #self.checkBox_report_1
#         if self.checkBox_group.checkedButton() == self.checkBox_report_1:
#             res = self.makeReport_1(report_dict)
#             print(f'RESULT: {res}')
#         else:
#             res = None
#             print('Отчет не выбран')
        self.progress_bar.setValue(100)
        print('Finish')

        # REPORT WINDOW
        if report_dict_df is not None:
            self.w = ReportWindow(report_name, report_dict_df)#, parent=self)
            self.w.show()       
        
#         if self.w is None:
#             self.w = ReportWindow(report_name, report_dict_df)#, parent=self)
#             self.w.show()
#         else:
#             self.w.close()
#             self.w = None # Discard reference
            
            
#     def child_closed(self):
#         self.w = None
#         self.progress_bar.reset()
            
        
    def exitApp(self):
        sys.exit()

#-------------------------------------------------------------------------        
        # !!!!!!!   RUN  !!!!!!
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
#     ex.show()
    sys.exit(app.exec_())