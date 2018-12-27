# coding:utf-8
import urllib.request
import re
from html.parser import HTMLParser
from html.entities import name2codepoint
import pandas as pd
import time
import datetime
import threading
import schedule
import threading
import random
import json

'''
功能：从天天基金网拉取每日基金净值预测数据

数据包括：产品代码、产品名称、预测净值

配置：将下面outputFile的值修改为你希望保存的位置

使用方法：直接运行
'''

outputFile = r'C:\Users\tangheng\Desktop'

class FundObj:
    '''
    单个基金对象
    '''
    def __init__(self,code,name,estimate):
        self.code = code
        self.name = name
        self.estimate = estimate

class WebHandle():
    '''
    数据爬取
    '''
    def __init__(self):
        url = r'http://api.fund.eastmoney.com/FundGuZhi/GetFundGZList?type=1&sort=3&orderType=desc&canbuy=0&pageIndex=1&pageSize=7000&_='
        k = str(int(round(time.time() * 1000)))
        url = url + k

        ua_list = [
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.6; rv2.0.1) Gecko/20100101 Firefox/4.0.1",
            "Mozilla/5.0 (Windows NT 6.1; rv2.0.1) Gecko/20100101 Firefox/4.0.1",
            "Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; en) Presto/2.8.131 Version/11.11",
            "Opera/9.80 (Windows NT 6.1; U; en) Presto/2.8.131 Version/11.11",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_0) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.56 Safari/535.11"]
        user_agent = random.choice(ua_list)

        request = urllib.request.Request(url = url)
        request.add_header('Accept','*/*')
        request.add_header('Accept-Encoding','gzip, deflate')
        request.add_header('Accept-Language','zh-CN,zh;q=0.9')
        request.add_header('Connection','keep-alive')
        # request.add_header('Cookie','st_pvi=61931004199536; qgqp_b_id=fcbd44a1159f30cf6ef5b68d9f332b5d; st_si=46241426393240')
        request.add_header('Host','api.fund.eastmoney.com')
        request.add_header('Referer','http://fund.eastmoney.com/fundguzhi.html')
        request.add_header('User-Agent',user_agent)

        page = urllib.request.urlopen(request)
        html = page.read()
        buffer0 = html.decode('utf8')
        obj = json.loads(buffer0)
        totalCount = obj['TotalCount']
        pageSize = obj['PageSize']
        datas = obj['Data']
        self.fundlist = datas['list']
        self.kk = list()

    # 读取数据，并保存为list<FundObj>
    def process(self):
        length = len(self.fundlist)

        for i in range(length):
            fund = self.fundlist[i]
            self.kk.append(FundObj(fund['bzdm'],fund['jjjc'],fund['gsz']))
        return self.kk


def mission():
    '''
    每天需要进行的任务
    '''
    handle = WebHandle()
    # 从天天基金网拉取数据
    data = handle.process()
    
    # 组装DataFrame
    codeList = []
    nameList = []
    estimateList = []
    for x in data:
        codeList.append(x.code)
        nameList.append(x.name)
        estimateList.append(x.estimate)
    
    result = pd.DataFrame({'code':codeList,'name':nameList,'estimate':estimateList})
    
    # 输出至Excel
    fileLocation = outputFile + '\天天基金网数据_'+ time.strftime('%Y-%m-%d',time.localtime(time.time())) + '.xlsx'
    time.strftime('%Y-%m-%d',time.localtime(time.time()))
    with pd.ExcelWriter(fileLocation) as writer:
        result.to_excel(writer,sheet_name = 'data')

    # 计算明天的执行时间
    now_time = datetime.datetime.now()
    if now_time.hour >= 9:
        now_time = now_time + datetime.timedelta(days=1)
    now_year = now_time.year
    now_month = now_time.month
    now_day = now_time.day
    next_time = str(now_year)+"-"+str(now_month)+"-"+str(now_day)+" 09:00:00"

    # 输出执行情况
    print('fetching complete')
    print('next fetching will be excute at :'+next_time)

def mission_thread():
    '''
    包装成线程
    '''
    threading.Thread(target=mission).start()
    

if __name__ == '__main__':
    mission_thread()
    schedule.every().day.at("9:00").do(mission_thread)

    while True:
        schedule.run_pending()
        time.sleep(1)