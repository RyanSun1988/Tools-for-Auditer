## 审计实用工具——顺丰速运批量查询并截图

作者 / wwcheng

公众号 / 15Seconds

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5fexmuuivj31hc0p3jxs.jpg)

一、引言

当我们需要查询某个顺丰单号的物流节点和电子存根信息时，需要进入顺丰速运官网，用与单号签收人或发件人手机号对应的微信进行扫码登陆，每次查询都会有滑动验证码，如果需要集中查询多个单号的话，需要花费较多的人工和时间。

最近在学习python中“类”和滑动验证码的解析，所以通过顺丰速运来进行尝试。

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5feyjeb23j30xf0lq41w.jpg)

二、简要流程

1 下载程序

2 安装Firefox浏览器

3 在模板中输入单号

4 运行程序

5 得到物流信息和截图

三、学习内容

1 滑动验证码的解析——PIL

2 自动化测试的应用——Selenium

3 日志模块——logging

四、程序具体使用方法

1 下载程序

从文末的下载链接中下载 sf_express，程序已经封装成EXE文件，不需要任何编程环境，内容如下

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5ff1f0lmtj30lg088q3b.jpg)

2 安装Firefox浏览器

从 Driver 文件夹中安装 Firefox浏览器，要根据自己电脑的系统版本来安装，如64位的系统就要安装 Firefox-v68.0.1-win64.exe

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5ff26p2agj30oe06uaak.jpg)

3 在模板中输入单号

在 Input 文件夹中的 bill_number.xlsx 表格中输入快递单号和备注，备注内容会成为截图的水印

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5ff3lnj4aj30oa04gwek.jpg)

4 运行程序

双击 sf_express_15Seconds.exe 开始运行程序

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5ff46yxtfj30o709zq3j.jpg)

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5ff908gruj30yf0hzgo7.jpg)

5 用微信扫码

等待程序运行一会后，会用电脑的默认图片查看工具来打开扫码界面，然后用手机微信进行扫码，用对应的微信号扫码可以同时看到物流节点和电子存根信息，也可以使用其他微信号扫码，但是只能获取物流节点，无法获取电子存根

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5ffb4ti2ij30xp0jk3zn.jpg)

6 得到物流信息和截图

运行结束后，会在 Output 文件夹中生成结果，结果包括截图、表格、日志文件

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5ff5tzfmzj30o005aaa8.jpg)

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5ff6nsb57j30if061q3o.jpg)

![img](https://tva2.sinaimg.cn/large/005PdFYUly1g5ff6z7ivmj30me04c0ss.jpg)