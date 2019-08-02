#!/usr/bin/python
# -*- coding: utf-8 -*-
# Author: wwcheng
import csv
import os
import random
import time
from urllib import request
import logging
import logging.handlers
import prettytable as pt
import requests
from fake_useragent import UserAgent
from lxml import etree


class QiChaMao(object):

    def __init__(self):
        self.start_time = time.time()
        self.timestr = ''
        self.timestr1 = str(int(time.time() * 1000))
        self.nowpath = os.path.abspath(os.curdir)
        self.search_key = ''
        self.company_id = ''
        self.company_name = ''
        self.info = []
        self.company_list = []
        self.base_headers = {
            'Host': 'www.qichamao.com',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': UserAgent().random,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'zh-CN,zh;q=0.9',
        }
        self.__logging()

    def __cookies_load(self):
        '''内部函数：加载cookies，切换时间戳'''
        filepath= self.nowpath+r'\cookies.txt'
        cookies = (open(filepath, 'r').read()).strip()
        if len(cookies)<50:
            self.logger.info('请先在cookies.txt中填写cookies信息')
            return None
        else:
            cookies = cookies[0:-10]+self.timestr
            return cookies

    def __logging(self):
        # 内部函数；用于记录日志
        logging.basicConfig(filename= self.nowpath+r"\report.log",
                            filemode="w",
                            format="%(asctime)s %(name)s:%(levelname)s:%(message)s",
                            datefmt="%d-%M-%Y %H:%M:%S",
                            level=logging.INFO
                            )
        self.logger = logging.getLogger("15Scends")
        handler = logging.StreamHandler()
        handler.setLevel(logging.INFO)
        formatter = logging.Formatter(
            "%(asctime)s %(name)s %(levelname)s %(message)s")
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)

    def WelCome(self):
        # 打印程序描述
        tb = pt.PrettyTable()
        tb.field_names = ['名称', '企业工商信息批量查询']
        tb.add_row(['作者', 'wwcheng'])
        tb.add_row(['微信公众号', '15Sceconds'])
        tb.add_row(['GitHub项目地址', 'https://github.com/nigo81/tools-for-auditor'])
        tb.add_row(['数据来源', 'www.qichamao.com'])
        print(tb)

    def load_company(self):
        '''加载公司列表'''
        input_csv = self.nowpath+r'\company_input.csv'
        reader = csv.reader(open(input_csv, 'r', encoding='utf-8-sig'))
        
        self.company_list = [r[1] for r in reader]
        if len(self.company_list)==1:
            self.logger.info('company_input.csv中未加载到任何公司名称')
        else:
            self.company_list = self.company_list[1:]
            self.logger.info('已成功加载{}个待查询公司名称'.format(len(self.company_list)))

    def get_companyid(self):
        ''''获取公司对应的链接'''
        search_key = request.quote(self.search_key)
        self.search_url = 'https://www.qichamao.com/search/all/{}?o=0&area=0&p=1'.format(
            search_key)
        headers1 = {'Cache-Control': 'max-age=0',
                    'Cookie': self.__cookies_load(), }
        headers1.update(self.base_headers)
        r = requests.get(self.search_url, headers=headers1)
        text = r.text
        html = etree.HTML(text)
        self.company_id = html.xpath('//a[@class="listsec_tit"]/@href')[0]

    def get_basicinfo(self):
        '''获取公司基本信息'''
        headers2 = {'Referer': self.search_url,
                    'Cookie': self.__cookies_load(), }
        headers2.update(self.base_headers)
        url2 = 'https://www.qichamao.com{}'.format(self.company_id)
        r = requests.get(url2, headers=headers2)
        text = r.text
        html = etree.HTML(text)
        self.company_name = html.xpath('//div[@class="arthd_tit"]/h1/text()')[0]
        if self.search_key == self.company_name:
            match = '一致'
        else:
            match = '名称不一致'
        info1 = ['无', '无']
        info2 = ['无']
        try:
            info1 = html.xpath('//div[@class="arthd_info"]/span/text()')[:2]
            info1 = [(i.split('：'))[1] for i in info1]
            info2 = html.xpath('//section[@class="pb-d2"]/ul[1]/li/span[@class="info"]/text()')
            info2 = [(i.replace('\n', '')).replace('--', '无') for i in info2]
            info2 = info2[:4]+info2[-10:]# 剔除名称
        except:
            self.logger.exception("正在记录发生的错误")
        self.info = [self.search_key, self.company_name, match]+info1+info2

    def Output_csv(self, content=''):
        '''输出文件'''
        output_path = self.nowpath + \
            r'\company_output_{}.csv'.format(self.timestr1)
        with open(output_path, 'a', newline='', encoding='utf-8-sig')as f:
            writer = csv.writer(f)
            writer.writerow(content)

    def main(self):
        '''主体循环'''
        self.WelCome()
        self.load_company()
        title = ['公司名称', '搜索结果', '名称是否一致', '企业电话', '企业邮箱', '统一社会信用代码',
                 '纳税人识别号', '注册号', '机构代码', '企业类型', '经营状态', '注册资本',
                 '成立日期', '登记机关', '经营期限', '所属地区', '核准日期', '企业地址', '经营范围']
        self.Output_csv(title)
        for i in range(len(self.company_list)):
            self.search_key = self.company_list[i]
            self.timestr = str(int(time.time() * 1000))
            self.logger.info('当前进度：{}/{} {}'.format(i+1, len(self.company_list),self.search_key))
            try:
                self.get_companyid()
                self.get_basicinfo()
                self.Output_csv(self.info)
                sleep_time = round(random.uniform(0.2, 0.8), 2)
                time.sleep(sleep_time)
            except Exception:
                self.logger.info('获取{}的信息时出错'.format(i))
                self.logger.exception("正在记录发生的错误")
            if i+1!=1 and (i+1) % 30 == 0:
                self.logger.info('为减小网站压力，每获取{}个信息，休息1分钟再继续'.format(str(i+1)))
                time.sleep(60)
        stamp = time.time() - self.start_time
        self.logger.info('查询总用时：'+str(stamp))
        
    def run(self):
        if self.__cookies_load():
            self.main()
            time.sleep(4)

if __name__ == "__main__":

    QiChaMao = QiChaMao()
    QiChaMao.run()
