# -*- coding: utf-8-sig -*-
"""
File Name ：    main
Author :        Eric
Create date ：  2020/10/28
------------------------------------------------------ 20.12.23 更新----------------------------------------------------
修改Mongo Collection Name后缀为年月日
新增一键删除3P库中所有Collection功能
设置爬虫请求超时时间为5分钟
新增自动添加不存在的文件及文件及
------------------------------------------------------ 21.02.03 更新----------------------------------------------------
* 添加判断订单是否全部取消，接收返回信号停止程序运行
* 当订单取消时新增弹窗提醒
* 修改库存查询方式为批量查询
* 删除mongo信息，将保存结果存储到本地json文件
* 新增缓存文件，记录上次打开文件夹为本次选择的原始数据所在文件夹路径打开
* 添加函数判断文件夹是否存在，当文件不存在时自动建立
* 添加条件判断原始数据表是否为空，并给出提示
* 当使用实时API计算库存后，使用快照API批量获取SKU的Location信息，并合并到结果dataframe中
* 关闭清空在线数据库信息API接口
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
import os
import warnings
from xlrd.biffh import XLRDError
import requests
import shutil
warnings.filterwarnings('ignore')


class ModelThrP(ComboSplit,Ui_MainWindow,QMainWindow):
    def __init__(self):
        super(Ui_MainWindow,self).__init__()
        super().__init__()
        self.setupUi(self)
        self.select_radio_button()
        self.open_file_path = settings.origin_file_path
        self.cache_file = settings.cache_file
        self.cache_result = settings.result_cache_path_file
        if not os.path.exists(self.open_file_path):os.mkdir(self.open_file_path)
        if not os.path.exists(settings.out_put_file_path): os.makedirs(settings.out_put_file_path)
        try:self.filter_data_frame = pd.read_excel(os.path.join(settings.out_put_file_path,settings.out_put_file['FilterFile']))
        except:self.filter_data_frame = pd.DataFrame()
        try:
            self.sum_result_frame = pd.read_excel(os.path.join(settings.out_put_file_path,settings.out_put_file['SumResult']))
        except:
            self.filter_data_frame = pd.DataFrame()
        self.file_path = settings.origin_file_path
        self.actionDeleteCollections.triggered.connect(self.truncate_database)

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
        signal = -1
        if self.info == '':
            QMessageBox.information(self,'Surprise','别皮！！！')
            return
        if self.info == '11':
            try:
                signal = self.main_origin_data_process()
                if signal == 1:
                    QMessageBox.information(self,'Notice','原始数据处理完成',QMessageBox.Ok)
                elif signal == 4: return
            except XLRDError: QMessageBox.information(self,'Warning','文件类型错误，请确认',QMessageBox.Ok)
            except: QMessageBox.information(self,'Warning','发生错误，请重试',QMessageBox.Ok)
            return
        elif self.info == '12':
            if signal == 4: return
            try:
                signal = self.main_get_inv_data()
                if signal == 1:
                    QMessageBox.information(self, 'Notice', '库存下载完成', QMessageBox.Ok)
            except HTTPError:
                QMessageBox.information(self, 'Warning', '网络错误，请检查网络连接', QMessageBox.Ok)
            except requests.exceptions.ConnectionError:
                QMessageBox.information(self, 'Warning', '与服务器连接出现问题，请稍后重试', QMessageBox.Ok)
            except:
                QMessageBox.information(self, 'Warning', '发生错误，请重试', QMessageBox.Ok)
            return
        elif self.info == '13':
            # if signal == 4: return
            # try:
                signal = self.main_inv_data_process()
                if signal: QMessageBox.information(self, 'Notice', '库存数据处理完成', QMessageBox.Ok)
            # except:QMessageBox.information(self, 'Warning', '发生错误，请重试', QMessageBox.Ok)
            # return

        elif self.info == '14':
            if signal == 4: return
            try:
                signal = self.main_combo_split()
                if signal: QMessageBox.information(self, 'Notice', '处理完成，文件已保存', QMessageBox.Ok)
            except:QMessageBox.information(self, 'Warning', '发生错误，请重试', QMessageBox.Ok)
            return

        elif self.info == '15':
            try:
                signal = self.main_origin_data_process()
                if signal == 4: return
                self.main_get_inv_data()
                self.main_inv_data_process()
                signal = self.main_combo_split()
                if signal: QMessageBox.information(self, 'Notice', '处理完成', QMessageBox.Ok)
            except:QMessageBox.information(self, 'Warning', '发生错误，请重试', QMessageBox.Ok)

    def search_file(self,filetypename,):
        '''查找指定用途的excel文件'''
        QMessageBox.information(self,'Warning','%s未找到，请指定'%filetypename,QMessageBox.Ok)
        file_name,ok = QFileDialog.getOpenFileName(self,'选择%s'%filetypename,self.file_path,'Excel (*.xls *.xlsx *.xltx);')
        if ok:
            return file_name

    def write_cache_path(self,path):
        '''保存打开的文件夹路径'''
        with open(self.cache_file,'w') as fp:
            fp.write(path)

    def read_cache_path(self):
        '''读取上次打开的文件夹路径'''
        with open(self.cache_file, 'r') as fp:
            path = fp.readlines()[-1].strip()
            return path

    def main_origin_data_process(self):
        '''原始数据处理'''
        QMessageBox.information(self, 'Notice', '选择原始数据所在文件夹', QMessageBox.Ok)
        self.file_path = settings.origin_file_path
        if os.path.exists(self.cache_file): self.file_path = self.read_cache_path() # 读上次打开文件夹路径
        if not os.path.exists(os.path.dirname(self.cache_file)): os.makedirs(os.path.dirname(self.cache_file))
        self.file_path = QFileDialog.getExistingDirectory(self,'选择原始数据所在文件夹',self.file_path)

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
        if origin_data_frame.empty:
            QMessageBox.information(self,'Warning!','原始表为空表，请注意！！！',QMessageBox.Ok)
            return 4
        origin_data_frame = self.relation_judgement(sku_relation_file, origin_data_frame)
        origin_data_frame = self.avc_judgement(avc_file, origin_data_frame)
        self.filter_data_frame = origin_data_frame.drop('condition', axis=1)  # 记录筛选完成后的原始数据
        self.sum_result_frame = self.data_aggr(origin_data_frame)  # 按Model Num与Vendor Code汇总后的结果表
        self.sum_result_frame.to_excel(os.path.join(settings.out_put_file_path,settings.out_put_file['SumResult']),index=False) # 临时结果文件，预防断点
        self.filter_data_frame.to_excel(os.path.join(settings.out_put_file_path,settings.out_put_file['FilterFile']),index=False)
        self.write_cache_path(self.file_path)  # 保存本次读取文件所在文件夹路径
        if self.filter_data_frame.empty:
            QMessageBox.information(self,'Notice','筛选完成，本次无订单',QMessageBox.Ok)
            return 4
        return 1

    def main_get_inv_data(self):
        '''获取库存数据'''
        self.sum_result_frame = pd.read_excel(os.path.join(settings.out_put_file_path,settings.out_put_file['SumResult']))
        sku_list = self.sum_result_frame['Model Number'].tolist()
        if len(sku_list)==0:
            QMessageBox.information(self, 'Notice', '本次无订单信息', QMessageBox.Ok)
            return 4
        self.get_inv_num(sku_list)
        return 1

    def main_inv_data_process(self):
        '''库存数据处理'''
        inv_result_file = os.path.join(settings.out_put_file_path,settings.out_put_file['InvFile'])
        sum_agg_result_file = os.path.join(settings.out_put_file_path,settings.out_put_file['SumResult'])
        if not os.path.exists(sum_agg_result_file):
            QMessageBox.information(self,'Warning','请按顺序执行',QMessageBox.Ok)
            return

        self.data_merge(sum_agg_result_file)
        self.vendor_sheet_data_process(inv_result_file)
        self.status_modify(inv_result_file)
        return 1

    def write_result_path(self, path):
        '''保存打开的文件夹路径'''
        with open(self.cache_result, 'w') as fp:
            fp.write(path)

    def read_result_path(self):
        with open(self.cache_result,'r') as fp:
            path = fp.readlines()[-1].strip()
            return path

    def main_combo_split(self):
        '''将combo拆分为对应单品信息，并结合库存给出建议下单数量'''
        vendor_order_file = os.path.join(self.file_path, settings.origin_file['ComboRelationFile'])
        if not os.path.exists(vendor_order_file):
            vendor_order_file = self.search_file("Vendor Ordering")
        result_path = self.file_path
        if not os.path.exists(os.path.dirname(self.cache_file)): os.makedirs(os.path.dirname(self.cache_file))
        if os.path.exists(self.cache_result): result_path = self.read_result_path()
        QMessageBox.information(self,'Notice','选择要保存的结果文件夹',QMessageBox.Ok)
        result_path = QFileDialog.getExistingDirectory(self,'选择要保存的文件夹',result_path)

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
        try:
            os.rename(os.path.join(settings.out_put_file_path,settings.out_put_file['InvFile']),inv_target_file)
            os.rename(os.path.join(settings.out_put_file_path,settings.out_put_file['SumResult']),agg_target_file)
            os.rename(os.path.join(settings.out_put_file_path,settings.out_put_file['FilterFile']),origin_target_file)
            self.write_result_path(result_path)
        except OSError:
            shutil.copy(os.path.join(settings.out_put_file_path, settings.out_put_file['InvFile']), inv_target_file)
            shutil.copy(os.path.join(settings.out_put_file_path, settings.out_put_file['SumResult']), agg_target_file)
            shutil.copy(os.path.join(settings.out_put_file_path, settings.out_put_file['FilterFile']), origin_target_file)
            os.remove(os.path.join(settings.out_put_file_path, settings.out_put_file['InvFile']))
            os.remove(os.path.join(settings.out_put_file_path, settings.out_put_file['SumResult']))
            os.remove(os.path.join(settings.out_put_file_path, settings.out_put_file['FilterFile']))
        return 1

    def truncate_database(self):
        QMessageBox.information(self,'Notice','此功能已停止使用，数据错误请联系管理员',QMessageBox.Ok)
        return
        ok = QMessageBox.information(self,'Warnings','确定删除所有数据表吗？？？',QMessageBox.Yes|QMessageBox.No)
        if not ok == 16384: return None
        ok = QMessageBox.information(self, 'Warnings', '考虑好了吗？？？', QMessageBox.Yes | QMessageBox.No)
        if not ok == 16384: return None
        ok = QMessageBox.information(self, 'Warnings', '再考虑一下吧！！！', QMessageBox.Yes | QMessageBox.No)
        if not ok == 16384: return None
        QMessageBox.information(self,'Notice','开始删除数据表',QMessageBox.Ok)
        self.delete_mongo_collections()
        QMessageBox.information(self,'Notice','删库完成，准备跑路吧！',QMessageBox.Ok)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainwindow = ModelThrP()
    mainwindow.show()
    sys.exit(app.exec_())