# -*- coding: utf-8 -*-
import unittest
import HTMLTestRunner, time, os


def interface():
    DIR = os.path.dirname(__file__)	# 获取工程文件根目录路径
    if os.path.isdir(DIR + u'/test_report') is False:
        os.mkdir(DIR + u'/test_report')
    newtime = time.strftime('%Y-%m-%d %H-%M-%S', time.localtime()) # 获取系统时间
    test_dir = DIR + u'/test_case'
    discover = unittest.defaultTestLoader.discover(test_dir, pattern='testCase.py') # 打开test_case里的testCase.py
    filePath = DIR + u'/test_report/' + newtime + u'接口自动化测试报告.html' # 生成测试报告的路径
    fp = open(filePath, 'wb')	# 打开文件流
    runner = HTMLTestRunner.HTMLTestRunner(stream=fp, title=u'接口自动化测试报告', description='商品搜索接口&商圈搜索接口') #写入测试数据
    runner.run(discover) # 执行测试


if __name__ == '__main__':
    interface()
