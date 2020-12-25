# -*- coding: utf-8-sig -*-
"""
File Name ：    settings
Author :        Eric
Create date ：  2020/10/23
"""
import os
import copy

# 原始数据文件目录
origin_file_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
# 输出文件目录
out_put_file_path = os.path.join(origin_file_path, 'results')
# 验证输出文件目录是否改变
origin_file_path_check = copy.deepcopy(origin_file_path)

origin_file = {
    'BasicData': 'basic_data.xlsx',  # 原始表
    'SKURalationFile': '关系表.xlsx',  # Vendor 对应 SKU 开头的关系表

    'CancelFile': 'AVC Cancellation List.xlsx',  # 记录无效ASIN的文件
    'ComboRelationFile': 'Amazon Vendor Ordering sheet.xlsx',  # Combo对应单品关系表，及CaseSize关系
}

out_put_file ={
    'FilterFile':'filter.xlsx', # 原始表处理结果
    'SumResult' : 'sum_result.xlsx', # 库存整理结果
    'InvFile': 'inventory.xlsx',  # 库存表
}

cancel_words = ['moq', 'ob', 'dsc'] # 不接单的后缀词
sku_cancel_sheets = ['Packing Cetificate'] # 与vendor无关的 记录要取消的产品编号的sheet
dependant_sku_sheets = ['IPOWX','IPOWP','IPOY3','IPOY4','DJ0JZ'] # 每个vendor对应的要取消的独立的产品编号

combo_sheet = 'Combo'  # 订单信息表中记录Combo及 单品关系表的 sheet名
casesize_sheet = 'AN.D'  # 订单信息表中记录单品Case Size的sheet名
max_combo_product_num = 4  # 组成combo的最大单品数量
combo_split_file_list = ['IPOWX.xlsx', 'IPOY4.xlsx'] # 要拆分的文件列表
combo_split_sheet_list = ['IPOWX', 'IPOY4']  # 要拆分的sheet列表

no_inventory_product_pattern = '\-(15|16|17)$' # 不走库存的产品序列
half_inventory_product_pattern = '' # 当库存超过订单两倍时参与计算的产品编号序列


if __name__ == '__main__':
    print(out_put_file_path)