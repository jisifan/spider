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


class MyHTMLParser(HTMLParser):
    '''
    处理html对象
    '''
    def __init__(self):
        HTMLParser.__init__(self)
        self.data = []

    def handle_starttag(self, tag, attrs):
        # if self.lasttag == 'input':
        #     for attr in attrs:
        #         print("     attr:", attr)
        pass

    def handle_endtag(self, tag):
        # print("End tag  :", tag)
        pass

    def handle_data(self, data):
        if self.lasttag == 'td' or self.lasttag == 'a':
            self.data.append(data)

    def get_data(self):
        return FundObj(self.data[1],self.data[2],self.data[6])

    # 每次调用push_data，self.data都会被清空
    # 返回产品代码，产品名称，预测净值
    def push_data(self):
        fundobject = FundObj(self.data[1],self.data[2],self.data[6])
        self.data = []
        return fundobject

class WebHandle():
    '''
    数据爬取
    '''
    def __init__(self):
        self.url = 'http://fund.eastmoney.com/fundguzhi.html'
        self.kk = list()

    # 读取数据，并保存为list<FundObj>
    def process(self):
        page = urllib.request.urlopen(self.url)
        html = page.read()
        buffer0 = html.decode('gbk')
        # 获取包含基金列表的 <table></table>
        reg = r'<table id="oTable" class="dbtable">(.+)</table>'
        reg_fund = re.compile(reg)
        temp = reg_fund.findall(buffer0)#进行匹配
        buffer1 = temp[0]

        # 获取包含基金列表的 <tbody></tbody>
        reg2 = r'<tbody>(.+)</tbody>'
        reg_fund2 = re.compile(reg2)
        temp2 = reg_fund2.findall(buffer1)
        buffer2 = temp2[0]

        # 获取list<tr>，一个<tr>包含一支基金
        reg3 = r'<tr>(.+?)</tr>'
        reg_fund3 = re.compile(reg3)
        temp3 = reg_fund3.findall(buffer2)

        reg4 = r'<tr class="ev">(.+?)</tr>'
        reg_fund4 = re.compile(reg4)
        temp4 = reg_fund4.findall(buffer2)

        parser1 = MyHTMLParser()

        length = max(len(temp3),len(temp4))

        for i in range(length):
            if(i<len(temp3)):
                parser1.feed(temp3[i])
                self.kk.append(parser1.push_data())
            if(i<len(temp4)):
                parser1.feed(temp4[i])
                self.kk.append(parser1.push_data())

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