# 系统测试——豆瓣
## 概述

本项目用于对豆瓣网站进行系统测试。测试用例以excel的形式存在，Python程序读取excel中的测试用例，然后去执行，最后输出测试结果到新的excel文件中。

使用到的框架或者模块：```selenium```、```xlrd```、```xlwt```、```xlutils```

## 文件说明

* Excel.py: 使用了```xlrd```模块读取excel文件，```xlutils```和```xlwt```模块写excel文件。
* TestDouban.py: 使用了```selenium```来操纵浏览器进行网站的各种测试。
* TestCases.xlsx: 写明了测试用例，皆以```xpath```的形式。
* TestResults.xls: 生成的测试结果。xls格式而不是xlsx格式。
* config-template.json: 读取用户的账号密码。写好账号密码后，将其重命名为```config.json```，否则需要手动输入账号密码。

## 安装依赖

```shell
$ python3 -m pip install [模块名] # 手动去安装一下
```

## 运行

```shell
$ python3 TestDouban.py
```

## 演示

https://www.bilibili.com/video/av89356377
