# -*- coding: utf-8-sig -*-
"""
File Name ：    get_api_data
Author :        Eric
Create date ：  2020/11/11
"""
import requests
from multiprocessing import Pool
import os
import pandas as pd
from pymongo import MongoClient
import datetime
# from gevent import monkey
# monkey.patch_all()
# import gevent

class GetApiInvData:
    def __init__(self):
        self.inv_save_file = r'inv.csv'


    def get_current_date(self):
        today = datetime.datetime.today()
        date = str(today.month).rjust(2, '0') + str(today.day).rjust(2, '0')
        return date


    def hs_mongo_config(self,col_name, db_name='3P'):
        host = ''
        port = ''
        user = ''
        pwd = ''
        client = MongoClient(host, port, username=user, password=pwd)[db_name]
        col = client[col_name]
        return col

    def get_api_data(self,sku,location='LA',retry=0):
        '''
        以post方式获取指定仓库指定SKU的产品库存信息
        '''
        uri =  f'http://vpn.bizrightllc.com:8111/TPA/API/InventoryData/getInventoryDataDailySnapShot?pageNo=1&pageSize=10000&ItemNum={sku}&Warehouse={location}'
        print(f'开始获取{sku}数据')
        if retry >= 3: return

        try:
            response = requests.post(uri)
            if response.status_code == 200:
                print(f'{sku}获取成功')
                return response.json()['result']['dataResult'][0]
            else:
                print(sku, ' 获取失败， ', response.status_code, retry + 1)
            return self.get_api_data(sku, location,retry+1)
        except:
            print('其他错误',sku,retry)
            return self.get_api_data(sku, location, retry + 1)

    def save_to_mongo(self,db,info):
        if not info:return
        if db.insert(info.copy()):
            print(f'插入成功{info}')
        else:
            print(f'插入失败{info}')


    def check_lost_data(self,sku_list):
        '''
        获取api数据完成后，对比查找缺失的sku数据，重新获取并保存
        '''
        date = self.get_current_date()
        mongo = self.hs_mongo_config(f'inv_data_{date}')

        target_frame = pd.DataFrame(mongo.find())
        target_list = target_frame['ItemNum'].drop_duplicates().tolist()
        for  sku in sku_list:
            if not sku in target_list:
                self.get_data(sku)

    def del_inv_file(self):
        '''
        处理完成后删除inv文件
        '''
        os.remove(self.inv_save_file)