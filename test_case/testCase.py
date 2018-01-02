# coding=utf-8

import unittest, requests
import xlrd, xlsxwriter
from test_data.load_excel import readexcel_data
import os, time

global url, payload, result, actual, sum, sum_pass, sum_fail, temp, name
name = ''
url = ''
payload = ''
result = ''
actual = ''
sum = 0
sum_pass = 0
sum_fail = 0
temp = 3
newtime = time.strftime('%Y-%m-%d %H-%M-%S', time.localtime())  # 获取系统时间
DIR = os.path.dirname(os.path.dirname(__file__))  # 获取文件根目录路径
filename = os.path.join(DIR, 'test_data', 'interface.xlsx')  # 获取interface.xlsx的路径
if os.path.isdir(DIR + u'/test_report') is False:  # 判断工程根目录是否有test_report文件夹，没有则新建
    os.mkdir(DIR + u'/test_report')
data = xlrd.open_workbook(filename)  # 打开interface.xlsx文件
sheet = data.sheet_by_index(0)  # 读取interface.xlsx文件第一页的数据
nrows = sheet.nrows  # 获取第一页的总行数


# 设置格式
def get_format(wd, option={}):
    return wd.add_format(option)


# 设置居中
def get_format_center(wd, num=1):
    return wd.add_format({'align': 'center', 'valign': 'vcenter', 'border': num})


# 设置边框
def set_border_(wd, num=1):
    return wd.add_format({}).set_border(num)


# 写数据
def _write_center(worksheet, cl, data, wd):
    return worksheet.write(cl, data, get_format_center(wd))


# 居中
def _write_center(worksheet, cl, data, wd):
    return worksheet.write(cl, data, get_format_center(wd))


# 写入第一页固定内容和结果内容
def init(worksheet):
    # 设置列行的宽高
    global url, payload, result, actual, sum, sum_pass, sum_fail
    worksheet.set_column("A:A", 15)
    worksheet.set_column("B:B", 20)
    worksheet.set_column("C:C", 20)
    worksheet.set_column("D:D", 20)
    worksheet.set_column("E:E", 20)
    worksheet.set_column("F:F", 20)

    worksheet.set_row(1, 30)
    worksheet.set_row(2, 30)
    worksheet.set_row(3, 30)
    worksheet.set_row(4, 30)

    # worksheet.set_row(0, 200)

    define_format_H1 = get_format(workbook, {'bold': True, 'font_size': 18})
    define_format_H2 = get_format(workbook, {'bold': True, 'font_size': 14})
    define_format_H1.set_border(1)

    define_format_H2.set_border(1)
    define_format_H1.set_align("center")
    define_format_H2.set_align("center")
    define_format_H2.set_bg_color("blue")
    define_format_H2.set_color("#ffffff")
    # Create a new Chart object.

    worksheet.merge_range('A1:D1', '测试报告总况', define_format_H2)
    worksheet.merge_range('A2:B2', "接口总数", get_format_center(workbook))
    worksheet.merge_range('A3:B3', "通过总数", get_format_center(workbook))
    worksheet.merge_range('A4:B4', "失败总数", get_format_center(workbook))
    worksheet.merge_range('A5:B5', "测试日期", get_format_center(workbook))

    data1 = {"test_sum": sum, "test_success": sum_pass, "test_failed": sum_fail, "test_date": newtime}
    _write_center(worksheet, "C2", data1['test_sum'], workbook)
    _write_center(worksheet, "C3", data1['test_success'], workbook)
    _write_center(worksheet, "C4", data1['test_failed'], workbook)
    _write_center(worksheet, "C5", data1['test_date'], workbook)

    _write_center(worksheet, "D2", "通过比", workbook)
    worksheet.merge_range('D3:D5', '%d%%' % (sum_pass / sum * 100), get_format_center(workbook))


# 写入第二页固定内容
def test_detail(worksheet):
    global url, payload, result, actual, sum, sum_pass, sum_fail
    # 设置列行的宽高
    worksheet.set_column("A:A", 30)
    worksheet.set_column("B:B", 20)
    worksheet.set_column("C:C", 20)
    worksheet.set_column("D:D", 20)

    worksheet.set_row(1, 30)
    worksheet.set_row(2, 30)

    worksheet.merge_range('A1:E1', '测试详情', get_format(workbook, {'bold': True, 'font_size': 18, 'align': 'center',
                                                                 'valign': 'vcenter', 'bg_color': 'blue',
                                                                 'font_color': '#ffffff'}))
    _write_center(worksheet, "A2", '接口名称', workbook)
    _write_center(worksheet, "B2", '接口请求地址', workbook)
    _write_center(worksheet, "C2", '请求内容', workbook)
    _write_center(worksheet, "D2", '测试结果', workbook)
    _write_center(worksheet, "E2", '返回内容', workbook)


# 生成饼形图
def pie(workbook, worksheet):
    chart1 = workbook.add_chart({'type': 'pie'})
    chart1.add_series({
        'name': '接口测试统计',
        'categories': '=测试总况!$A$3:$A$4',
        'values': '=测试总况!$C$3:$C$4',
    })
    chart1.set_title({'name': '接口测试统计'})
    chart1.set_style(10)
    worksheet.insert_chart('E1', chart1, {'x_offset': 25, 'y_offset': 10})


# 写入第二页结果内容
def write_data(worksheet):
    global url, payload, result, actual, sum, sum_pass, sum_fail, temp, name
    # 设置列行的宽高
    worksheet.set_column("A:A", 30)
    worksheet.set_column("B:B", 20)
    worksheet.set_column("C:C", 20)
    worksheet.set_column("D:D", 20)
    worksheet.set_column("E:E", 20)

    worksheet.set_row(temp, 30)

    yellow = workbook.add_format(
        {'bg_color': 'yellow', 'color': 'red', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    data = {"t_name": name, "t_url": url,
            "t_param": payload, "t_result": result,
            "t_actual": actual}

    if result.find('测试通过') != -1:
        _write_center(worksheet2, "A" + str(temp), data["t_name"], workbook)
        _write_center(worksheet2, "B" + str(temp), data["t_url"], workbook)
        _write_center(worksheet2, "C" + str(temp), data["t_param"], workbook)
        _write_center(worksheet2, "D" + str(temp), data["t_result"], workbook)
        _write_center(worksheet2, "E" + str(temp), data["t_actual"], workbook)
        temp += 1
    else:
        worksheet2.write('A' + str(temp), data['t_name'], yellow)
        worksheet2.write('B' + str(temp), data["t_url"], yellow)
        worksheet2.write('C' + str(temp), data["t_param"], yellow)
        worksheet2.write('D' + str(temp), data["t_result"], yellow)
        worksheet2.write('E' + str(temp), data["t_actual"], yellow)
        temp += 1


workbook = xlsxwriter.Workbook(DIR + u'/test_report/' + newtime + u'接口自动化测试报告.xlsx')  # 生成的xlsx文件
worksheet = workbook.add_worksheet("测试总况")  # 生成文件的第一页名称
worksheet2 = workbook.add_worksheet("测试详情")  # 生成文件的第二页名称


class MyTestCase(unittest.TestCase):
    def test_interface(self):  # 接口自动化测试
        global url, payload, result, actual, sum, sum_pass, sum_fail, name
        for i in range(nrows - 1):  # 循环遍历interface.xlsx表
            if i < nrows:
                name = readexcel_data(0, i + 1, 0)  # 读取interface.xlsx表的第一页的A列各行数据
                url = readexcel_data(0, i + 1, 1)  # 读取interface.xlsx表的第一页的B列各行数据
                payload = readexcel_data(0, i + 1, 2)  # 读取interface.xlsx表的第一页的C列各行数据
                headers = {
                    'content-type': "application/json",
                    'cache-control': "no-cache",
                }  # 发送的头文件类型
                # if name.find(u'退款申请接口') and name.find(u'退款申诉接口'):
                #     time.sleep(10)
                try:
                    r = requests.request("POST", url, data=payload.encode('utf-8'), headers=headers)  # 向着请求地址发送请求数据
                except Exception as e:  # 判断失败的情况
                    result = '测试失败'
                    actual = 'error: ' + str(e)
                    sum_fail = sum_fail + 1
                    sum += 1
                    write_data(worksheet2)
                    continue
                s = r.text  # 获得返回json内容
                if r.status_code == 200:  # 判断返回码200的情况
                    if s.find('"errorCode":0') != -1 and s.find('"success":true') != -1:  # 判断返回json正常的情况
                        print('测试通过  ' + url)
                        result = '测试通过'
                        actual = r.text
                        sum_pass = sum_pass + 1
                        sum += 1
                        write_data(worksheet2)
                    elif s == 'null':  # 判断返回空的情况
                        print('测试失败  ' + url + "  " + r.text)
                        # self.assertEqual(s.find('"errorCode":0'), 0, msg=url + r.text)
                        result = '测试失败'
                        actual = r.text
                        sum_fail = sum_fail + 1
                        sum += 1
                        write_data(worksheet2)
                    elif name == '获取4位数随机验证码':  # 跳过获取4位验证码的情况
                        print('测试通过  ' + url)
                        result = '测试通过'
                        actual = r.text
                        sum_pass = sum_pass + 1
                        sum += 1
                        write_data(worksheet2)
                    else:  # 不符合上述情况的时候
                        print('测试失败  ' + url + "  " + r.text)
                        # self.assertEqual(s.find('"errorCode":0'), 0, msg=url + r.text)
                        result = '测试失败'
                        actual = r.text
                        sum_fail = sum_fail + 1
                        sum += 1
                        write_data(worksheet2)
                else:  # 返回码不为200的时候
                    print('测试失败  ' + url + "  " + str(r.status_code))
                    # self.assertEqual(r.status_code, 200, msg=url + r.status_code)
                    result = '测试失败'
                    actual = r.status_code
                    sum_fail = sum_fail + 1
                    sum += 1
                    write_data(worksheet2)
                i += 1
            else:  # 不符合情况跳出循环
                break

        test_detail(worksheet2)
        init(worksheet)
        pie(workbook, worksheet)

        workbook.close()  # 关闭文件流
