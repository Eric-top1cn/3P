# -*- coding: utf-8-sig -*-
"""
File Name ：    inv_data_process
Author :        Eric
Create date ：  2020/11/11
"""

import numpy as np
import pandas as pd
import re
import os
from bin.get_api_data import GetApiInvData
from bin.data_filter import DataFilter
from settings import settings
from PyQt5.QtWidgets import QMessageBox

class InvDataProcess(GetApiInvData,DataFilter):
    def __init__(self):
        super().__init__()

    def inv_process(self):
        '''
        处理库存数据，inv_result_file为要输出的库存结果保存目录
        '''

        self.inv_result_file = os.path.join(self.result_path,'inventory.xlsx')
        sku_list = self.sum_result_frame['Model Number'].tolist()
        self.start_spider(sku_list)
        self.data_merge()
        self.vendor_sheet_data_process(self.inv_result_file)
        self.status_modify(self.inv_result_file)


    def vendor_sheet_data_process(self,target_file):
        '''
        按汇总后的Vendor处理数据，并将不同Vendor保存到不同sheet中
        '''
        if os.path.exists(target_file): os.remove(target_file)
        excel_writer = pd.ExcelWriter(target_file)
        sum_frame = self.sum_result_frame.copy()
        for sheet in sum_frame['Vendor Code'].drop_duplicates():
            sheet_frame = sum_frame[sum_frame['Vendor Code'] == sheet]
            sheet_frame = sheet_frame.sort_values('Model Number', ascending=False)

            sheet_frame.loc[sheet_frame['ActualAvailableQty'] <= 0, 'Warehouse'] = np.nan
            sheet_frame.loc[sheet_frame['ActualAvailableQty'] <= 0, 'ActualAvailableQty'] = np.nan
            sheet_frame.to_excel(excel_writer,sheet_name=sheet, index=False)
        excel_writer.save()


    def status_modify(self,target_file):
        '''
        修改原始表数据中不接单的状态，并将dj0jz中无库存的产品全部改为不接单
        target file为汇总数据与库存合并后的存储文件
        '''

        try:
            djz_frame = pd.read_excel(target_file,sheet_name='DJ0JZ')
            djz_frame['ActualAvailableQty'].fillna(0,inplace=True)
        except:
            djz_frame = pd.DataFrame()
        filter_frame = self.filter_data_frame.copy()

        # 修改状态
        for i,model_number in enumerate(filter_frame['Model Number']):
            model_number = filter_frame['Model Number'].iloc[i]
            vendor = filter_frame['Vendor Code'].iloc[i]
            if not vendor == 'DJ0JZ':
                continue
            status = filter_frame['Availability Status'].iloc[i]  # DJ0JZ未接单
            if not status == 'AC - Accepted': continue
            # 跳过BOP中查不到的库存信息
            if len(djz_frame[djz_frame['Model Number'] == model_number]) < 1:
                continue

            if int(djz_frame[djz_frame['Model Number'] == model_number]['ActualAvailableQty']) <= 0:
                filter_frame['Availability Status'].iloc[i] = 'OS - Cancelled: Out of stock'

        filter_frame.loc[filter_frame['Availability Status'] != 'AC - Accepted', 'Quantity Confirmed'] = 0
        self.filter_data_frame = filter_frame.copy()


    def data_merge(self,sum_agg_file):
        '''
        将汇总后的原始数据与库存数据合并
        '''

        sum_frame = pd.read_excel(sum_agg_file)
        inv_mongo = self.hs_mongo_config('inv_data_' + self.get_current_date())
        inv_frame = pd.DataFrame(inv_mongo.find())
        # inv_frame = pd.read_csv(self.inv_save_file,sep='\t')
        inv_frame.drop_duplicates(['ItemNum'],inplace=True)
        inv_frame = inv_frame[['ItemNum', 'ActualAvailableQty', 'Warehouse']]
        for column in ['ActualAvailableQty', 'Warehouse', 'filter']: sum_frame[column] = np.nan


        for i,sku in enumerate(sum_frame['Model Number']):
            sku = str(sum_frame['Model Number'].iloc[i])
            vendor_code = str(sum_frame['Vendor Code'].iloc[i])
            try:
                qty = int(inv_frame[inv_frame['ItemNum'] == sku]['ActualAvailableQty'])
                location = inv_frame[inv_frame['ItemNum'] == sku]['Warehouse'].to_list()[0]
            except:
                qty = ''
                location = ''
            sum_frame['ActualAvailableQty'].iloc[i] = qty
            sum_frame['Warehouse'].iloc[i] = location
            if vendor_code.upper() == 'DJ0JZ':
                sum_frame['filter'].iloc[i] = '只走库存'
            if re.search(r'-(16|17)$', sku):  # 是否结尾待确认
                sum_frame['filter'].iloc[i] = '不走库存'
            if re.search(r'-15$', sku):
                requirement_num = int(sum_frame['total'].iloc[i])
                inventory_num = int(sum_frame['ActualAvailableQty'].iloc[i])
                if requirement_num <= inventory_num // 2:
                    sum_frame['filter'].iloc[i] = '走库存'
                else:
                    sum_frame['filter'].iloc[i] = '不走库存'

        sum_frame.drop('ASIN', axis=1, inplace=True)
        sum_frame = sum_frame[
            ['Vendor Code', 'Model Number', 'Title', 'total', 'ActualAvailableQty', 'Warehouse', 'filter']]
        sum_frame['Model Number'].fillna('空', inplace=True)
        sum_frame.loc[sum_frame['ActualAvailableQty'] == '', 'ActualAvailableQty'] = np.nan  # 将空串替换为Nan
        sum_frame['ActualAvailableQty'].fillna(0, inplace=True)  # 将Nan值填充为0
        sum_frame.sort_values('ActualAvailableQty', ascending=False, inplace=True)
        self.sum_result_frame = sum_frame

if __name__ == '__main__':
    inv = InvDataProcess()
    # inv.del_inv_file()
    file = r'C:\Users\Administrator\data\3p'
    inv.data_filter(file)
    # inv.inv_process()
    inv.get_api_data('GL557502')