import json
import os
import keyboard
import pyperclip
import requests
from PIL import Image, ImageGrab
import prettytable as pt
import datetime

#打印程序介绍
def WelCome():
    tb = pt.PrettyTable()
    tb.field_names = ['名称', 'Sina图片外链生成工具']
    tb.add_row(['作者', 'wwcheng'])
    tb.add_row(['微信公众号', '15Scecond'])
    tb.add_row(['GitHub项目地址', 'https://github.com/nigo81/tools-for-auditor'])
    tb.add_row(['运行环境', 'Python3'])
    tb.add_row(['推荐图片格式', 'PNG'])
    tb.add_row(['',''])
    tb.add_row(['\n使用方法', 
    '截图后 使用组合键CTRL+ALT来上传\n上传成功后 外链将直接拷贝到剪贴板中\n截图和外链文本将会保存在TEMP文件夹中'])
    print(tb)

#从剪切板保存图片至本地
def clipboard_save(pic_path):
    pic = True
    im = ImageGrab.grabclipboard()
    if im != None:
        im.save(pic_path)
    else:
        print('2. 剪切板上未发现图片')
        pic = False
    return pic

#将保存的图片上传至Sina服务器，并返回外链
def Sina_upload(pic_path):
    try:
        url = 'https://api.top15.cn/picbed/picApi.php?type=multipart'
        files = {'file': ('imageGrab.png', open(pic_path, 'rb'))}
        data = requests.post(url, files=files)
        data = json.loads(data.text)
        img_url = data['url']
        print("3. 生成的图片外链：{}".format(img_url))
        return img_url
    except Exception as err:
        print("3. 图片上传失败：{}".format(err))
        return None

#主体逻辑
def main():
    print('\n1. 截图后 按组合键CTRL+ALT开始上传')
    if keyboard.wait(hotkey='ctrl+alt') == None:
        now_path = os.path.abspath(os.curdir)
        time_str = datetime.datetime.strftime(datetime.datetime.now(),'%Y%m%d%H%M%S')
        if not os.path.exists(now_path+r'\temp'):
            os.mkdir(now_path+r'\temp')
        pic_path = now_path+r'\temp\imageGrab_%s.png'%time_str
        if clipboard_save(pic_path):
            print('2. 正在上传图片')
            url_str = Sina_upload(pic_path)
            #url_str = r'"<img src=""%s"">"'%Sina_upload(pic_path)
            if url_str != None:
                f = open(now_path+r'\temp\imageUrl.txt', 'a+')
                f.write(url_str+'\n')
                pyperclip.copy(url_str)
                pyperclip.paste()
                print('4. 外链已经保存并拷贝到剪贴板中')

#程序入口
if __name__ == "__main__":
    WelCome()
    while True:
        main()
