#!/usr/bin/python
# -*- coding: utf-8 -*-
# Author: wwcheng

import csv
import datetime
import logging
import logging.handlers
import os
import random
import shutil
import time
from urllib.request import urlretrieve

import pandas as pd
import prettytable as pt
from PIL import Image, ImageDraw, ImageFont
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait


class Sf_Express(object):

    def __init__(self):
        self.WelCome()
        self.start_time = time.time()
        self.query_time = datetime.datetime.strftime(
            datetime.datetime.now(), '%Y-%m-%d %H-%M-%S')
        self.nowpath = os.path.abspath(os.curdir)
        self.font = self.nowpath+r'\Fonts\simkai.ttf'
        self.url = r'http://www.sf-express.com/cn/sc/dynamic_function/waybill/#search/bill-number/'

        self.bill_number_list = []
        self.remarks_list = []
        self.bill_number_list1 = []
        self.remarks_list1 = []
        self.route_list = []
        self.dzcg_list = []

        self.bill_number = ''
        self.remarks = ''
        self.route = ''
        self.dzcg = ''

        self.incomplete_img = ''
        self.complete_img = ''
        self.screenshot_img1 = ''
        self.screenshot_img2 = ''

        self.initial_pos = 24
        self.mov_pos = 0
        self.forward_tracks = []
        self.backward_tracks = []

        self.load_status = True
        self.verif_retries = 1
        self.error_retries = 0
        self.num = 1

        if self.__Is64Windows:
            self.executable_path = self.nowpath+r'\Driver\geckodriver-win64.exe'
        else:
            self.executable_path = self.nowpath+r'\Driver\geckodriver-win32.exe'

        self.ceart_enviroment()
        self.__logging()
        self.load_input_xls()

        if self.load_status:
            firefox_options = webdriver.FirefoxOptions()
            firefox_options.add_argument('--headless')
            self.browser = webdriver.Firefox(
                executable_path=self.executable_path,
                firefox_options=firefox_options,
                service_log_path=self.log_path+r'\geckodriver.log'
            )
            self.open_sf()
            self.scan_login()

    def __Is64Windows(self):
        # 内部函数：用于判断电脑是否为64位系统，是则返回TRUE
        return 'PROGRAMFILES(X86)' in os.environ

    def __mkdir(self, path):
        # 类内部函数；用于创建文件夹
        folder = os.path.exists(path)
        if not folder:
            os.makedirs(path)

    def __logging(self):
        # 内部函数；用于记录日志
        logging.basicConfig(filename=self.log_path+r"\sf_express.log",
                            filemode="w",
                            format="%(asctime)s %(name)s:%(levelname)s:%(message)s",
                            datefmt="%d-%M-%Y %H:%M:%S",
                            level=logging.INFO
                            )
        self.logger0 = logging.getLogger("sf_express_error")
        self.logger = logging.getLogger("15Scends")
        handler = logging.StreamHandler()
        handler.setLevel(logging.INFO)
        formatter = logging.Formatter(
            "%(asctime)s %(name)s %(levelname)s %(message)s")
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)

    def __wait(self):
        # 内部函数；随机等待
        sleep_time = round(random.uniform(0.5, 2), 2)
        time.sleep(sleep_time)

    def __is_visible(self, locator, timeout=60):
        # 内部函数；判断网页元素是否出现
        try:
            WebDriverWait(self.browser, timeout).until(
                EC.visibility_of_element_located((By.XPATH, locator)))
            return True
        except TimeoutException:
            return False

    def __is_not_visible(self, locator, timeout=60):
        # 内部函数；判断网页元素是否消失
        try:
            WebDriverWait(self.browser, timeout).until_not(
                EC.visibility_of_element_located((By.XPATH, locator)))
            return True
        except TimeoutException:
            return False

    def __watermark(self, imageFile):
        # 内部函数；用于给单张图片添加水印
        font = ImageFont.truetype(self.font, 24)
        time_str = datetime.datetime.strftime(
            datetime.datetime.now(), '%Y-%m-%d %H:%M:%S')
        watermark = '备注：' + self.remarks+'\n' + \
            '截图时间：'+str(time_str)+'\n微信公众号：15Seconds'
        im = Image.open(imageFile)
        im1 = im.crop((im.size[1]/3, 0, im.size[0]-im.size[1]/3, im.size[1]))
        draw = ImageDraw.Draw(im1)
        # 设置文字位置/内容/颜色/字体
        draw.text((im1.size[0]*3/5, im1.size[1]-100),
                  watermark, (255, 0, 0), font=font)
        draw = ImageDraw.Draw(im1)
        im1.save(imageFile)

    def __is_pixel_equal(self, bg_image, fullbg_image, x, y):
        # 内部函数：用于判断两张图片的像素点差异
        # 获取缺口图片的像素点(按照RGB格式)
        bg_pixel = bg_image.load()[x, y]
        # 获取完整图片的像素点(按照RGB格式)
        fullbg_pixel = fullbg_image.load()[x, y]
        # 设置一个判定值，像素值之差超过判定值则认为该像素不相同
        threshold = 50
        # 判断像素的各个颜色之差，abs()用于取绝对值
        se1 = abs(bg_pixel[0] - fullbg_pixel[0])
        se2 = abs(bg_pixel[1] - fullbg_pixel[1])
        se3 = abs(bg_pixel[2] - fullbg_pixel[2])
        if se1 < threshold and se2 < threshold and se3 < threshold:
            return True
        else:
            return False

    def __reload_verify_img(self):
        # 内部函数，用于刷新验证码
        self.logger.info('尝试刷新验证码')
        try:
            self.verif_retries += 1
            self.browser.find_element_by_xpath(
                '//div[@class="tcaptcha-action-icon show-reload"]').click()
            self.__wait()
            self.get_verify_img()
            self.get_distance()
            self.get_tracks()
            self.move_slider()
        except Exception as e:
            self.logger0.exception("正在记录发生的可预见错误")

    def __output_xls(self):
        # 内部函数；用于输出快递信息表格
        self.logger.info('将物流节点和电子存根信息写入csv\n')
        path = self.xls_path+r'\sf_express_route.xlsx'
        writer = pd.ExcelWriter(path)
        df1 = pd.DataFrame(
            {'物流单号': self.bill_number_list1, '备注': self.remarks_list1, '物流节点': self.route_list, '电子存根': self.dzcg_list})
        df1.to_excel(writer)
        writer.save()

    def WelCome(self):
        # 打印程序描述
        tb = pt.PrettyTable()
        tb.field_names = ['名称', '顺丰速运批量查询截图工具']
        tb.add_row(['作者', '王文铖'])
        tb.add_row(['微信公众号', '15Sceconds'])
        tb.add_row(
            ['GitHub项目地址', 'https://github.com/nigo81/tools-for-auditor'])
        tb.add_row(['运行环境', 'Python3'])
        tb.add_row(['', '需要先安装Driver文件夹中的对应版本的Firefox浏览器'])
        tb.add_row(['使用方法', '然后再Input文件中填入顺丰单号和备注，备注将会成为截图的水印'])
        tb.add_row(['', '以上两步完成后双击 sf_express_15Seconds.exe 程序，开始运行'])
        print(tb)

    def load_input_xls(self):
        # 加载快递提单号并剔除非顺丰单号
        input_xls = self.nowpath+r'\Input\bill_number.xlsx'
        d = pd.read_excel(input_xls)
        bills = []
        for i in d.index.values:
            info = d.ix[i, ['快递单号', '备注']].to_list()
            bills.append(info)
        self.logger.info('部分数据预览如下：')
        self.logger.info(bills[:3])
        try:
            for bill in bills:
                bi = str(bill[0]).strip()
                if len(bi) == 12 or len(bi) == 15:
                    self.bill_number_list.append(bill[0])
                    self.remarks_list.append(bill[1])
                else:
                    self.logger.info('%s 不是顺丰单号' % bi)
        except Exception as e:
            self.logger0.exception("正在记录发生的可预见错误")
        if len(self.bill_number_list) > 0:
            self.logger.info('已成功加载%s个顺丰快递单号' % len(self.bill_number_list))
        else:
            self.logger.info('bill_number.xlsx 中未发现顺丰快递单号（12或15字节）')
            self.load_status = False
        self.logger.info('请在60秒内用手机完成微信扫码，超时后会视同不扫码继续运行')
        self.logger.info('可以不使用对应手机号的微信扫码，但是将无法获得电子存根信息')
        self.logger.info('成功扫码之后可以将二维码图片关闭')
        self.logger.info('正在打开微信扫码界面···\n')

    def ceart_enviroment(self):
        # 创建本次输出的文件环境
        result_path = self.nowpath+'\\Output\\' + self.query_time
        self.img_path = result_path+r'\img'
        self.xls_path = result_path+r'\xls'
        self.log_path = result_path+r'\log'
        self.temp_path = result_path+r'\temp'
        self.__mkdir(result_path)
        self.__mkdir(self.img_path)
        self.__mkdir(self.xls_path)
        self.__mkdir(self.log_path)
        self.__mkdir(self.temp_path)

    def open_sf(self):
        # 打开浏览器
        try:
            self.browser.get(self.url)
            self.__wait()
            self.browser.maximize_window()
        except Exception as e:
            self.logger0.exception("正在记录发生的可预见错误")
            self.logger.info('打开浏览器出错，检查火狐浏览器是否安装')

    def scan_login(self):
        # 扫码登陆
        self.__wait()
        self.browser.find_element_by_xpath(
            '//span[@class="agreeCookie"]').click()
        self.__wait()
        self.browser.find_element_by_xpath(
            '//a[@class="topa maidian"]').click()
        self.__is_visible('//img[@class="scan-img"]')
        self.__wait()
        self.scan_img = self.temp_path+r'\scan_img.png'
        self.browser.save_screenshot(self.scan_img)
        self.__wait()
        self.__watermark(self.scan_img)
        img = Image.open(self.scan_img)
        img.show()
        scan = self.__is_not_visible('//div[@class="layui-layer-shade"]')
        if scan == False:
            self.logger.info('超出60秒未进行微信扫码，直接进行查询')
            try:
                self.browser.find_element_by_xpath(
                    '//span[@class="layui-layer-setwin"]/a').click()
                self.__wait()
            except Exception as e:
                self.logger0.exception("正在记录发生的可预见错误")
                self.logger.info('没有查询到相关物流信息')
        else:
            pass

    def input_bill_number(self):
        # 输入快递单号开始查询
        self.logger.info('正在输入单号进行查询')
        self.browser.find_element_by_xpath(
            '//input[@class="token-input"]').send_keys(self.bill_number)
        self.__wait()
        self.browser.find_element_by_xpath('//span[@id="queryBill"]').click()
        self.__wait()

    def switch_to_iframe(self):
        # 切换至验证码弹窗网页框架
        self.__is_visible('//iframe')
        # 找到“嵌套”的iframe
        iframe = self.browser.find_element_by_xpath('//iframe')
        self.browser.switch_to.frame(iframe)
        self.__wait()

    def get_verify_img(self):
        # 获取验证码图片
        try:
            self.logger.info('尝试解析滑动验证码')
            self.incomplete_img = self.temp_path+r'\%s_img_incomplete.jpg' % self.num
            self.complete_img = self.temp_path+r'\%s_img_complete.jpg' % self.num
            # 保存带缺口的滑动图片
            self.__wait()
            img = self.browser.find_element_by_xpath('//img[@id="slideBkg"]')
            incomplete_img_url = img.get_attribute('src')
            urlretrieve(url=incomplete_img_url, filename=self.incomplete_img)
            # 保存完整的滑动图片
            complete_img_url = incomplete_img_url[:-1]+'0'
            urlretrieve(url=complete_img_url, filename=self.complete_img)
        except Exception as e:
            self.logger0.exception("正在记录发生的可预见错误")

    def get_distance(self):
        # 获取滑动验证码需要滑动的距离
        try:
            fullbg_image = Image.open(self.complete_img)
            bg_image = Image.open(self.incomplete_img)
            for i in range(self.initial_pos, fullbg_image.size[0]):
                # 遍历像素点纵坐标
                for j in range(fullbg_image.size[1]):
                    # 如果不是相同像素
                    if not self.__is_pixel_equal(fullbg_image, bg_image, i, j):
                        # 返回此时横轴坐标就是滑块需要移动的距离
                        self.mov_pos = i/2-self.initial_pos
                        return i
        except Exception as e:
            self.logger0.exception("正在记录发生的可预见错误")
            self.__wait()
            self.__reload_verify_img()
            self.get_verify_img()
            self.get_distance()

    def get_tracks(self):
        # 获取滑动路径列表
        self.forward_tracks.clear()
        self.mov_pos += 6  # 要回退的像素
        v0, s, t = 0, 0, 0.4  # 初速度为v0，s是已经走的路程，t是时间
        mid = self.mov_pos*3/5  # mid是进行减速的路程
        while s < self.mov_pos:
            if s < mid:  # 加速区
                a = 5
            else:  # 减速区
                a = -3
            v = v0
            tance = v*t+0.5*a*(t**2)
            tance = round(tance)
            s += tance
            v0 = v+a*t
            self.forward_tracks.append(tance)
        # 因为回退20像素，所以可以手动打出，只要和为20即可
        bk_tracks = [-1, -2, -1, -2]
        random.shuffle(bk_tracks)  # 洗牌
        self.backward_tracks = bk_tracks

    def move_slider(self):
        # 开始模拟滑动验证码，先前进再后退
        self.logger.info('解析成功，正在模拟滑动验证码···')
        # 得到滑块标签
        slider = self.browser.find_element_by_xpath(
            '//div[@id="tcaptcha_drag_button"]')
        # 使用click_and_hold()方法悬停在滑块上，perform()方法用于执行
        ActionChains(self.browser).click_and_hold(slider).perform()
        # 使用move_by_offset()方法拖动滑块，perform()方法用于执行
        for forword_track in self.forward_tracks:
            ActionChains(self.browser).move_by_offset(
                xoffset=forword_track, yoffset=0).perform()
        self.__wait()
        for backward_tracks in self.backward_tracks:
            ActionChains(self.browser).move_by_offset(
                xoffset=backward_tracks, yoffset=0).perform()
        # 模拟人类对准时间
        self.__wait()
        # 释放滑块
        ActionChains(self.browser).release().perform()
        try:
            if self.__is_visible('//p[@class="tcaptcha-title"]', 10):
                self.__reload_verify_img()
                return False
            else:
                return True
        except Exception as e:
            self.logger0.exception("正在记录发生的可预见错误")

    def switch_to_default(self):
        # 切换至主页面
        self.browser.switch_to.default_content()
        self.__wait()

    def get_route_info(self):
        # 获取物流节点信息和截图
        self.route = ''
        self.logger.info('获取物流节点信息和截图')
        self.screenshot_img1 = self.img_path + \
            r'\%s-wljd-%s.png' % (self.num, self.bill_number)
        target = self.browser.find_element_by_xpath(
            '//input[@type="checkbox"]')
        self.browser.execute_script("arguments[0].scrollIntoView();", target)
        target.click()
        target.click()
        self.__wait()
        self.browser.save_screenshot(self.screenshot_img1)
        try:
            routes = self.browser.find_elements_by_xpath(
                '//div[@class="route-list"]/ul')
            for rou in routes:
                r = rou.text.replace('\n', ' ')
                self.logger.info(r)
                self.route = self.route+'\n'+r
        except Exception as e:
            self.logger0.exception("正在记录发生的可预见错误")
            self.logger.info('没有查询到相关物流信息')

    def get_dzcg_info(self):
        # 获取电子存根信息和截图
        self.logger.info('正在获取电子存根信息和截图')
        self.screenshot_img2 = self.img_path + \
            r'\%s-dzcg-%s.png' % (self.num, self.bill_number)
        self.dzcg = ''
        try:
            self.browser.find_element_by_xpath(
                '//div[@class="operates-wrapper"]/a').click()
            self.__wait()
            target = self.browser.find_element_by_xpath(
                '//a[@aria-controls="electronicStub"]')
            target.click()
            self.browser.execute_script(
                "arguments[0].scrollIntoView();", target)
            self.__wait()
            self.browser.save_screenshot(self.screenshot_img2)
            dzcgs = self.browser.find_elements_by_xpath(
                '//table[@class="borderrb"]/tbody/tr')
            for dzcg in dzcgs:
                r = dzcg.text.replace('\n', ' ')
                self.logger.info(r)
                self.dzcg = self.dzcg+'\n'+r
        except Exception as e:
            self.logger0.exception("正在记录发生的可预见错误")
            self.logger.info('该单号对应的手机号与扫码的微信号可能不符，无法获取电子存根截图')

    def add_watermark(self):
        # 给截图添加水印
        self.logger.info('正在为截图添加水印')
        self.__watermark(self.screenshot_img1)
        try:
            self.__watermark(self.screenshot_img2)
        except Exception as e:
            self.logger0.exception("正在记录发生的可预见错误")

    def next_bill_number(self):
        # 关闭电子存根界面，开始查下一个单号
        try:
            self.browser.find_element_by_xpath(
                '//button[@class="close"]').click()
            self.__wait()
        except Exception as e:
            self.logger0.exception("正在记录发生的可预见错误")
        self.browser.find_element_by_xpath(
            '//div[@class="tokenfield form-control"]/div[1]/a').click()
        self.__wait()

    def main(self):
        # 主体程序
        self.input_bill_number()
        self.switch_to_iframe()
        self.get_verify_img()
        self.get_distance()
        self.get_tracks()
        self.move_slider()
        self.switch_to_default()
        self.get_route_info()
        self.get_dzcg_info()
        self.add_watermark()
        self.next_bill_number()

    def start_single(self, billnumer):
        # 查询单个物流单号
        st = time.time()
        self.logger.info('正在查询的物流单号：%s' % billnumer)
        self.verif_retries = 0
        try:
            self.main()
        except:
            self.browser.close()
            self.__wait()
            self.open_sf()
            self.scan_login()
            self.__wait()
            self.main()
        finally:
            self.logger.info('单号 '+billnumer+'  用时 ' +
                             str(time.time() - st)+'秒')
            # 准备写入excel中
            self.bill_number_list1.append(self.bill_number)
            self.remarks_list1.append(self.remarks)
            self.dzcg_list.append(self.dzcg)
            self.route_list.append(self.route)
            self.__output_xls()

    def start_all(self):
        for self.num in range(self.num, len(self.bill_number_list)+1):
            self.logger.info('当前进度：%s/%s' %
                             (self.num, len(self.bill_number_list)))
            self.bill_number = str(self.bill_number_list[self.num-1])
            self.remarks = str(self.remarks_list[self.num-1])
            self.start_single(self.bill_number)

            if self.num!=1 and (self.num-1) % 20 == 0:
                self.logger.info('已经获取%s个物流信息，休息2分钟再继续' % str(self.num-1))
                time.sleep(120)
            self.error_retries == 1

    def run(self):
        # while self.error_retries <= 3 and self.num+1 > len(self.bill_number_list):
        try:
            self.start_all()
        except Exception as e:
            self.logger0.exception("正在记录发生的可预见错误")
            # self.error_retries += 1
            # self.logger.info('浏览器出现错误，正在尝试第%s次重启' % self.error_retries)
            # self.run()
        finally:
            stamp = time.time() - self.start_time
            self.logger.info('程序总用时：'+str(stamp))

    def __del__(self):
        if self.load_status:
            self.browser.quit()


if __name__ == "__main__":
    sf = Sf_Express()
    temp_path = sf.temp_path
    sf.run()
    del sf
    shutil.rmtree(temp_path)
