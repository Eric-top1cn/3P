# -*- coding: utf-8-sig -*-
"""
File Name ：    data_filter
Author :        Eric
Create date ：  2020/11/11
"""
from settings import settings
import pandas as pd
import numpy as np
import re
import os
import warnings
warnings.filterwarnings('ignore')

class DataFilter:

    def data_filter(self,file_path):
        '''main 函数'''
        self.file_path = file_path
        self.result_path = os.path.join(self.file_path,'result')
        if not os.path.exists(self.result_path): os.mkdir(self.result_path)

        origin_data_file = os.path.join(file_path,settings.origin_file['BasicData'])
        sku_relation_file = os.path.join(file_path,settings.origin_file['SKURalationFile'])
        avc_file = os.path.join(file_path,settings.origin_file['CancelFile'])

        origin_data_frame = self.read_origin_data(origin_data_file)
        origin_data_frame = self.relation_judgement(sku_relation_file,origin_data_frame)
        origin_data_frame = self.avc_judgement(avc_file,origin_data_frame)
        self.filter_data_frame = origin_data_frame.drop('condition', axis=1) # 记录筛选完成后的原始数据
        self.sum_result_frame = self.data_aggr(origin_data_frame) # 按Model Num与Vendor Code汇总后的结果表

    def read_origin_data(self,basic_file):
        '''
        读给出的原始文件中数据，并进行简单处理
        basic_file 为记录原始数据的excel表格
        '''
        result = pd.DataFrame(pd.read_excel(basic_file))  # 源数据
        if result.empty: return result
        result.drop([0, 1], inplace=True)
        result.columns = result.iloc[0]
        result.reset_index(inplace=True, drop=True)
        result.drop(0, inplace=True)
        result.reset_index(inplace=True, drop=True)
        result['condition'] = np.nan
        return result

    def relation_judgement(self,relation_file,data_frame):
        '''
        根据给定的关系表，判断指定的Vendor Code中是否包含指定的Model Num
        relation file 为记录Vendor及ModelNum的Excel表格
        data frame 为原始数据处理完成后的结果
        '''
        relationship = pd.DataFrame(pd.read_excel(relation_file))  # 读关系表
        cancel_words = settings.cancel_words  # cancel words


        for i in range(len(data_frame['Vendor Code'])):
            code = str(data_frame['Vendor Code'].iloc[i])
            model_num = str(data_frame['Model Number'].iloc[i])

            if model_num[:4] in relationship[relationship['Vendor Code'] == code]['Model Num'].to_list():
                data_frame['condition'].iloc[i] = 1
            else:
                continue
            for word in cancel_words:  # cancel words
                if re.search(word + '$', model_num, re.I):
                    data_frame['condition'].iloc[i] = 0
        return data_frame


    def avc_judgement(self,avc_file,data_frame):
        '''
        根据avc cancellation中记录的ASIN判断Model Num是否满足下单需求
        avc_file 指定AVC Cancellation Excel表格路径
        data_frame 待处理的数组
        '''
        cancel_sheets = settings.dependant_sku_sheets
        # 将指定的sheet中要取消的ASIN合并，并将Combo asin添加到每个code中
        # cancel asin
        asin_frame = pd.DataFrame()
        for sheet in cancel_sheets:
            try:
                item = pd.DataFrame(
                    pd.read_excel(
                        avc_file,
                        sheet_name=sheet,
                        usecols=[1]))
            except:
                continue
            item['Vendor Code'] = sheet
            item.dropna(subset=['ASIN'], inplace=True)
            asin_frame = pd.concat([asin_frame, item], ignore_index=True)
        item = pd.DataFrame()
        for sheet in settings.sku_cancel_sheets:
            item = pd.concat([item, pd.DataFrame(pd.read_excel(
                avc_file, sheet_name=sheet, usecols=[1]))], ignore_index=True)

        for sheet in cancel_sheets:
            item['Vendor Code'] = sheet
            item.dropna(subset=['ASIN'], inplace=True)
            asin_frame = pd.concat([asin_frame, item], ignore_index=True)
            # 判断asin是否可接单
            data_frame['condition'].fillna(0, inplace=True)
            for i in range(len(data_frame['Vendor Code'])):
                if not int(data_frame['condition'].iloc[i]) == 1:
                    continue
                code = data_frame['Vendor Code'].iloc[i]
                asin = data_frame['ASIN'].iloc[i]
                if asin in asin_frame[asin_frame['Vendor Code'] == code]['ASIN'].to_list():
                    data_frame['condition'].iloc[i] = 0
        # 修改Availability Status
        data_frame.loc[~ (data_frame['condition'] == 1.0),
                       'Availability Status'] = 'CB - Cancelled: Not our publication'
        return data_frame

    def data_aggr(self,data_frame):
        '''
        将选出的可接单的产品按Model Num分组求和
        '''
        sum_frame = data_frame[data_frame['condition'] == 1][['Model Number', 'ASIN',
                                    'Title', 'Vendor Code']].drop_duplicates(['Model Number', 'Vendor Code'])
        sum_frame['total'] = 0
        for i in range(len(sum_frame['total'])):
            model_num = sum_frame['Model Number'].iloc[i]
            vendor_code = sum_frame['Vendor Code'].iloc[i]
            sum_frame['total'].iloc[i] = data_frame[data_frame['Model Number'] ==
                                        model_num][data_frame['Vendor Code'] == vendor_code]['Quantity Ordered'].sum()
        return sum_frame

if __name__ == '__main__':
    f = DataFilter()

    f.data_filter(r'C:\Users\Administrator\data\3p')