# -*- coding: utf-8-sig -*-
"""
File Name ：    get_api_data
Author :        Eric
Create date ：  2020/11/11
"""
import requests
import re
import json
from settings import settings
class GetApiInvData:
    def __init__(self):
        self.inv_save_file = settings.inv_cache_file

    def get_inv_num(self,sku_list):
        url = input('请输入BOP接口url')  # 实时数据接口
        
        data = input('请补充api参数')
        res = requests.post(url, data=data)
        if res.status_code == 200:
            if not res.json()['code'] == 200:  # 正常响应，返回错误
                print(res.json())
                # 从SKU列表中删除不存在信息，并重新发起请求
                sku_list.remove(re.search('ItemNum(.+?)不存在', res.json()['message']).group(1).strip())
                return self.get_inv_num(sku_list)
            json.dump(self.parse_json(res), open(self.inv_save_file, 'w'))
            self.get_location_data(sku_list)
            # json.dump(res.json()['result']['dataResult'], open(self.inv_save_file, 'w'))
        else:
            return self.get_inv_num(sku_list)

    def parse_json(self, response):
        '''解析库存结果，选出结果中Warehouse为LA的库存结果，并逐条dict删除StoreLocationLines键及其结果'''
        result = []
        js = response.json()['result']['dataResult']
        for item in js:
            item = item['InventoryModel']['InventoryDataTypes']
            item = [info for info in item if info['Warehouse'] == 'LA'][0]
            item = {key: item[key] for key in item.keys() if not key == 'StoreLocationLines'}
            result.append(item)
        return result

    def get_location_data(self,sku_list):
        '''通过快照API，获取给定SKU列表中有库存产品的Location信息'''
        url = input('请输入BOP备份接口url')  # 备份数据接口
        data =  input('请补充api参数')
        res = requests.post(url, data=data)
        if res.status_code == 200:
            if not res.json()['code'] == 200:  # 正常响应，返回错误
                print(res.json())
                # 从SKU列表中删除不存在信息，并重新发起请求
                sku_list.remove(re.search('ItemNum(.+?)不存在', res.json()['message']).group(1).strip())
                return self.get_inv_num(sku_list)
            json.dump(res.json()['result']['dataResult'], open(settings.location_data_file, 'w'))
        else:
            return self.get_inv_num(sku_list)

if __name__ == '__main__':
    # t = GetApiInvData().get_current_date()
    # print(t)
    pass