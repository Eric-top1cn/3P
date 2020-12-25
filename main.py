# -*- coding: utf-8-sig -*-
"""
File Name ：    main
Author :        Eric
Create date ：  2020/10/28
"""
from bin.combo_split import ComboSplit
from ui.ui import Ui_MainWindow
from PyQt5 import QtCore, QtGui, QtWidgets,Qt
from PyQt5.QtGui import QIcon
from requests import HTTPError
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow,QMessageBox,QInputDialog,QLineEdit,QFileDialog,QAction,QInputDialog,QButtonGroup
from settings import settings
import pandas as pd
import re
import os
import numpy as np
import datetime
import warnings
from xlrd.biffh import XLRDError
from gevent import monkey
monkey.patch_all()
import gevent
import time



class ModelThrP(ComboSplit,Ui_MainWindow,QMainWindow):
    def __init__(self):
        super(Ui_MainWindow,self).__init__()
        super().__init__()
        self.setupUi(self)
        self.select_radio_button()
        self.open_file_path = settings.origin_file_path
        try:self.filter_data_frame = pd.read_excel(os.path.join(settings.out_put_file_path,settings.out_put_file['FilterFile']))
        except:self.filter_data_frame = pd.DataFrame()
        try:
            self.sum_result_frame = pd.read_excel(os.path.join(settings.out_put_file_path,settings.out_put_file['SumResult']))
        except:
            self.filter_data_frame = pd.DataFrame()
        self.file_path = settings.origin_file_path

    def select_radio_button(self):
        self.function_group = QButtonGroup()
        self.function_group.addButton(self.data_filter_button, id=11)
        self.function_group.addButton(self.inv_download_button, id=12)
        self.function_group.addButton(self.inv_data_process_button, id=13)
        self.function_group.addButton(self.combo_split_button, id=14)
        self.function_group.addButton(self.func_merge_button,id=15)

        self.info = ''
        self.function_group.buttonClicked.connect(self.func_group_radio_click)
        self.func_select_certain_button.clicked.connect(self.submit)

    def func_group_radio_click(self):
        sender = self.sender()
        if not sender == self.function_group: return

        if self.function_group.checkedId() == 11: self.info = '11'
        elif self.function_group.checkedId() == 12: self.info = '12'
        elif self.function_group.checkedId() == 13: self.info = '13'
        elif self.function_group.checkedId() == 14: self.info = '14'
        elif self.function_group.checkedId() == 15: self.info = '15'

        else: self.info = ''


    def submit(self):
        if self.info == '':
            QMessageBox.information(self,'Surprise','别皮！！！')
            return
        if self.info == '11':
            try:
                signal = self.main_origin_data_process()
                if signal:
                    QMessageBox.information(self,'Notice','原始数据处理完成',QMessageBox.Ok)
            except XLRDError: QMessageBox.information(self,'Warning','文件类型错误，请确认',QMessageBox.Ok)
            except: QMessageBox.information(self,'Warning','发生错误，请重试',QMessageBox.Ok)
            return
        elif self.info == '12':
            try:
                signal = self.main_get_inv_data()
                if signal: QMessageBox.information(self, 'Notice', '库存下载完成', QMessageBox.Ok)
            except HTTPError: QMessageBox.information(self,'Warning','网络错误，请重试',QMessageBox.Ok)
            except:QMessageBox.information(self, 'Warning', '发生错误，请重试', QMessageBox.Ok)
            return
        elif self.info == '13':
            try:
                signal = self.main_inv_data_process()
                if signal: QMessageBox.information(self, 'Notice', '库存数据处理完成', QMessageBox.Ok)
            except:QMessageBox.information(self, 'Warning', '发生错误，请重试', QMessageBox.Ok)
            return

        elif self.info == '14':
            try:
                signal = self.main_combo_split()
                if signal: QMessageBox.information(self, 'Notice', '处理完成，文件已保存', QMessageBox.Ok)
            except:QMessageBox.information(self, 'Warning', '发生错误，请重试', QMessageBox.Ok)
            return

        elif self.info == '15':
            try:
                self.main_origin_data_process()
                self.main_get_inv_data()
                self.main_inv_data_process()
                signal = self.main_combo_split()
                if signal: QMessageBox.information(self, 'Notice', '处理完成', QMessageBox.Ok)
            except:QMessageBox.information(self, 'Warning', '发生错误，请重试', QMessageBox.Ok)

    def search_file(self,filetype,):
        QMessageBox.information(self,'Warning','%s未找到，请指定'%filetype,QMessageBox.Ok)
        file_name,ok = QFileDialog.getOpenFileName(self,'选择%s'%filetype,self.file_path,'Excel (*.xls *.xlsx *.xltx);')
        if ok:
            return file_name


    def main_origin_data_process(self):
        QMessageBox.information(self, 'Notice', '选择原始数据所在文件夹', QMessageBox.Ok)
        self.file_path = QFileDialog.getExistingDirectory(self,'选择原始数据所在文件夹',settings.origin_file_path)

        origin_data_file = os.path.join(self.file_path, settings.origin_file['BasicData'])
        sku_relation_file = os.path.join(self.file_path, settings.origin_file['SKURalationFile'])
        avc_file = os.path.join(self.file_path, settings.origin_file['CancelFile'])
        if not os.path.exists(origin_data_file):
            origin_data_file = self.search_file('原始表')
        if not os.path.exists(sku_relation_file):
            sku_relation_file = self.search_file('关系表')
        if not os.path.exists(avc_file):
            sku_relation_file = self.search_file('Cancellation表')
        QMessageBox.information(self, 'Notice', '开始处理原始数据', QMessageBox.Ok)
        origin_data_frame = self.read_origin_data(origin_data_file)
        origin_data_frame = self.relation_judgement(sku_relation_file, origin_data_frame)
        origin_data_frame = self.avc_judgement(avc_file, origin_data_frame)
        self.filter_data_frame = origin_data_frame.drop('condition', axis=1)  # 记录筛选完成后的原始数据
        self.sum_result_frame = self.data_aggr(origin_data_frame)  # 按Model Num与Vendor Code汇总后的结果表

        self.sum_result_frame.to_excel(os.path.join(settings.out_put_file_path,settings.out_put_file['SumResult']),index=False) # 临时结果文件，预防断点
        self.filter_data_frame.to_excel(os.path.join(settings.out_put_file_path,settings.out_put_file['FilterFile']),index=False)
        return 1

    def main_get_inv_data(self):

        self.sum_result_frame = pd.read_excel(os.path.join(settings.out_put_file_path,settings.out_put_file['SumResult']))
        sku_list = self.sum_result_frame['Model Number'].tolist()
        self.start_spider(sku_list)

        return 1

    def main_inv_data_process(self):
        inv_result_file = os.path.join(settings.out_put_file_path,settings.out_put_file['InvFile'])
        sum_agg_result_file = os.path.join(settings.out_put_file_path,settings.out_put_file['SumResult'])
        if not os.path.exists(sum_agg_result_file):
            QMessageBox.information(self,'Warning','请按顺序执行',QMessageBox.Ok)
            return

        self.data_merge(sum_agg_result_file)
        self.vendor_sheet_data_process(inv_result_file)
        self.status_modify(inv_result_file)
        return 1


    def main_combo_split(self):
        vendor_order_file = os.path.join(self.file_path, settings.origin_file['ComboRelationFile'])
        if not os.path.exists(vendor_order_file):
            vendor_order_file = self.search_file("Vendor Ordering")

        QMessageBox.information(self,'Notice','选择要保存的结果文件夹',QMessageBox.Ok)
        result_path = QFileDialog.getExistingDirectory(self,'选择要保存的文件夹',self.file_path)

        self.get_case_frame(vendor_order_file)
        self.get_combo_splitted_frame(vendor_order_file)
        # 保存文件到结果文件夹
        self.inv_result_file = os.path.join(settings.out_put_file_path,settings.out_put_file['InvFile'])
        if not os.path.exists(self.inv_result_file):
            self.inv_result_file = self.search_file('库存表')

        excel_file = pd.ExcelFile(self.inv_result_file)
        for sheet in settings.combo_split_sheet_list:
            if not sheet in excel_file.sheet_names: continue  # 跳过不包含combo的sheet
            combo_split_result_file = os.path.join(result_path, f'{sheet}_GL52.xlsx')
            self.data_aggregate(sheet, combo_split_result_file)
        # 原始表、库存表保存到指定文件夹
        inv_target_file = os.path.join(result_path,settings.out_put_file['InvFile'])
        agg_target_file = os.path.join(result_path,settings.out_put_file['SumResult']) # 分类汇总后的数据
        origin_target_file = os.path.join(result_path, settings.out_put_file['FilterFile'])
        [os.remove(file) for file in [inv_target_file,agg_target_file,origin_target_file]  if os.path.exists(file)]

        os.rename(os.path.join(settings.out_put_file_path,settings.out_put_file['InvFile']),inv_target_file)
        os.rename(os.path.join(settings.out_put_file_path,settings.out_put_file['SumResult']),agg_target_file)
        os.rename(os.path.join(settings.out_put_file_path,settings.out_put_file['FilterFile']),origin_target_file)
        return 1

    def start_spider(self,sku_list):
        # 删除已存在的sku
        date = self.get_current_date()
        mongo = self.hs_mongo_config(f'inv_data_{date}')
        data_frame = pd.DataFrame(mongo.find())
        if not data_frame.empty:
            sku_list_got = data_frame['ItemNum'].tolist()
            for i,sku in enumerate(sku_list.copy()):
                if sku in sku_list_got: sku_list.remove(sku)

        task_list = []
        for sku in sku_list: #self.get_data(sku)
            task = gevent.spawn(self.get_data,sku)
            task_list.append(task)
        gevent.joinall(task_list)

        self.check_lost_data(sku_list)


    def get_data(self,sku):
        # main函数
        date = self.get_current_date()
        mongo = self.hs_mongo_config(f'inv_data_{date}')
        data = self.get_api_data(sku)
        # self.data_to_file(data)
        self.save_to_mongo(mongo,data)
        time.sleep(2)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainwindow = ModelThrP()
    mainwindow.show()
    sys.exit(app.exec_())