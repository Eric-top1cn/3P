# -*- coding: utf-8-sig -*-
"""
File Name ：    combo_split
Author :        Eric
Create date ：  2020/11/11
"""
import os
import pandas as pd
import re
from bin.inv_data_process import InvDataProcess
from settings import settings
import numpy as np


class ComboSplit(InvDataProcess):
    def __init__(self):
        super().__init__()

    def combo_process(self):
        vendor_order_file = os.path.join(self.file_path,settings.origin_file['ComboRelationFile'])
        print(vendor_order_file)
        self.get_case_frame(vendor_order_file)
        self.get_combo_splitted_frame(vendor_order_file)
        excel_file = pd.ExcelFile(self.inv_result_file)
        for sheet in settings.combo_split_sheet_list:
            if not sheet in excel_file.sheet_names: continue # 跳过不包含combo的sheet
            combo_split_result_file = os.path.join(self.result_path, f'{sheet}_GL52.xlsx')
            self.data_aggregate(sheet,combo_split_result_file)



    def data_aggregate(self,sheet, target_file):
        '''
        将原始数据与combo进行合并，并计算要下单的batch size与总价
        origin file为库存处理完成后要拆分combo的文件
        sheet为文件中要拆分combo的sheet
        target file为要保存的结果文件
        '''
        origin_file = self.inv_result_file
        excel_file = pd.ExcelFile(origin_file)
        if not sheet in excel_file.sheet_names:
            print(r'%s 无GL52订单' %sheet)
            return
        # 读关系表
        combo_split_frmae = self.new_combo_frame
        case_frame = self.case_frame

        # 读数据文件
        target_frame = pd.read_excel(origin_file)
        # 选出GL52产品
        target_frame = target_frame[
            target_frame['Model Number'].apply(lambda x: True if re.match('GL52', x) else False)]
        # 判断是否有GL52产品
        if target_frame.empty:
            print(r'%s 无GL52订单' % origin_file.split('\\')[-1].split('.')[0])
            return

        # Combo与单品结合
        target_frame = target_frame.merge(combo_split_frmae, on='Model Number', how='left')
        # 将合并后的GL52单品 sku 与 qty 改为GL52的SKU及1.0
        target_frame.loc[target_frame['sku'].isnull(), 'sku'] = target_frame.loc[
            target_frame['sku'].isnull(), 'Model Number']
        target_frame['qty'].fillna(1.0, inplace=True)
        # 计算要下单的单品数量 Combo数量 * weight
        target_frame['SingleNum'] = target_frame['total'] * target_frame['qty']

        # -15结尾的产品需要满足库存量为订单数2倍，否则不走库存
        target_frame.loc[
            (target_frame['total'] * 2 >= target_frame['ActualAvailableQty']) & target_frame['sku'].str.contains(
                '\-15$'), 'Location'] = np.nan
        target_frame.loc[
            (target_frame['total'] * 2 >= target_frame['ActualAvailableQty']) & target_frame['sku'].str.contains(
                '\-15$'), 'ActualAvailableQty'] = np.nan

        # 以-16 -17结尾的产品全部不走库存
        target_frame.loc[target_frame['sku'].apply(
            lambda x: True if re.search('(\-16|\-17)$', x) else False), 'ActualAvailableQty'] = np.nan
        target_frame.loc[
            target_frame['sku'].apply(lambda x: True if re.search('(\-16|\-17)$', x) else False), 'Location'] = np.nan

        # 将单品信息添加到目标矩阵中
        target_frame = target_frame.merge(case_frame, on='sku', how='left')
        # 计算每个case要下的订单数
        # (单品数量 - (库存数量 // CaseSize * CaseSize)) // CaseSize
        target_frame['SuggestOrderNum'] = (target_frame['SingleNum'] - (
                target_frame['ActualAvailableQty'].fillna(0) // target_frame['CaseSize'] * target_frame['CaseSize'])) // \
                                          target_frame['CaseSize']

        # 按单品SKU分类汇总
        total_frame = target_frame.groupby('sku').sum()

        # 删除不需合并的列
        total_frame.drop(['Univalence', 'CaseSize'], axis=1, inplace=True)
        # 将汇总后矩阵与单品信息合并
        total_frame = total_frame.merge(case_frame, on='sku', how='left')
        # 列重命名
        total_frame.rename(columns={'sku': 'SingleSKU', 'qty': 'weight', 'total': 'ComboOrderNum'}, inplace=True)
        # 保留目标列
        total_frame = total_frame[
            ['SingleSKU', 'Description', 'CaseSize', 'SingleNum', 'SuggestOrderNum', 'Univalence']]
        # 计算总价
        total_price = (total_frame['SuggestOrderNum'] * total_frame['CaseSize'] * total_frame['Univalence']).sum()
        # 将总价添加到汇总后的矩阵中
        total_frame = total_frame.append({'SingleSKU': '总价', 'Univalence': total_price}, ignore_index=True)

        # 列重命名并保留目标列
        target_frame.rename(
            columns={'Model Number': 'ComboSKU', 'total': 'ComboOrderNum', 'sku': 'SingleSKU', 'qty': 'weight'},
            inplace=True)
        target_frame = target_frame[
            ['Vendor Code', 'Title', 'ComboSKU', 'CaseSize', 'SingleSKU', 'Description', 'ActualAvailableQty',
             'SingleNum', 'SuggestOrderNum', 'Univalence']]

        # 保存文件
        writer = pd.ExcelWriter(target_file)
        target_frame.to_excel(writer, sheet_name='SplitSheet', index=False)
        total_frame.to_excel(writer, sheet_name='AggregationSheet', index=False)
        writer.save()


    def get_case_frame(self,case_file):
        '''
        读取combo与单品数量及单价对应关系表，保留'sku', 'Description', 'CaseSize', 'Univalence'列
        case file 为Vendor Order Excel表路径，其sheet为表中记录此关系的sheet名，在setting中修改
        '''
        # 判断文件是否存在
        if not os.path.exists(case_file):
            print(r'%s 文件不存在' % case_file)
            exit()
        case_frame = pd.read_excel(case_file, sheet_name=settings.casesize_sheet)
        # 保留可用的列并重命名
        case_frame = case_frame[case_frame.columns.tolist()[:4]]
        case_frame.columns = ['sku', 'Description', 'CaseSize', 'Univalence']
        self.case_frame = case_frame


    def get_combo_splitted_frame(self,combo_file):
        '''
        将Combo拆分成对应的单品并去重，删除重复值
        combo file为 Vendor Order Excel表路径，其sheet为表中记录combo关系的sheet名，在setting中修改
        '''
        if not os.path.exists(combo_file):
            print(f'%s 文件不存在' % (combo_file))
            exit()

        combo_compose_info_frame = pd.read_excel(combo_file,
                                                     sheet_name=settings.combo_sheet)
        # 保留Combo中不为空的行
        combo_compose_info_frame = combo_compose_info_frame[~combo_compose_info_frame[
            'Vendor SKU'].isnull()]
        # 删除ASIN列
        combo_compose_info_frame.drop('ASIN', axis=1, inplace=True)

        # 将Combo中1-4个单品合并为新表，存为ComboSKU（ModelNumber），SKU，QTY
        new_combo_frame = pd.DataFrame()
        for i in range(1, settings.max_combo_product_num + 1):
            iter_combo_frame = combo_compose_info_frame[['Vendor SKU', 'sku' + str(i), 'qty' + str(i)]]
            iter_combo_frame.columns = ['Model Number', 'sku', 'qty']
            iter_combo_frame.dropna(inplace=True)
            #     iter_combo_frame.dropna()
            new_combo_frame = new_combo_frame.append(iter_combo_frame, ignore_index=True)

        new_combo_frame.drop_duplicates(inplace=True)  # 去重
        self.new_combo_frame = new_combo_frame



if __name__ == '__main__':
    c = ComboSplit()
