import time
import json

import xlrd
import xlsxwriter

import requests


class STSpider():
    def __init__(self, name,name1):
        self.params={'':''}
        self.name = name
        self.name1=name1
        self.workbook = xlsxwriter.Workbook('./file/{}.xlsx'.format(name1))
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36',
            # 'Referer': 'http://webstads.sciinfo.cn/exportController.do?toExpert&sendId=30',
            'Cookie': 'JSESSIONID=34CA3F08814C5C4453F2309A7FCE3080; BROWSER_TYPE=Netscape; JEECGINDEXSTYLE=hplus; ZINDEXNUMBER=1990'
        }


    def generate_excel(self, rec_data, name, field):

        worksheet = self.workbook.add_worksheet(name)

        # bold_format = self.workbook.add_format({'bold': True})
        # 将二行二列设置宽度为15(从0开始)
        worksheet.set_column(1, 1, 15)

        # 用符号标记位置，例如：A列1行
        worksheet.write('A1', field[0])
        worksheet.write('B1', field[1])
        worksheet.write('C1', field[2])
        worksheet.write('D1', field[3])
        worksheet.write('E1', field[4])
        worksheet.write('F1', field[5])
        # worksheet.write('F1', 'wid', bold_format)
        # worksheet.write('F1', 'id', bold_format)

        row = 1
        col = 0
        for item in (rec_data):
            # 使用write_string方法，指定数据格式写入数据
            try:
                if item[field[0]]:
                    worksheet.write_string(row, col, item[field[0]])
            except:
                item[field[0]] = ''
                worksheet.write_string(row, col, item[field[0]])
            try:
                if item[field[1]]:
                    worksheet.write_string(row, col + 1, item[field[1]])
            except:
                item[field[1]] = ''
                worksheet.write_string(row, col + 1, item[field[1]])
            try:
                if item[field[2]]:
                    worksheet.write_string(row, col + 2, item[field[2]])
            except:
                item[field[2]] = ''
                worksheet.write_string(row, col + 2, item[field[2]])
            try:
                if item[field[3]]:
                    worksheet.write_string(row, col + 3, item[field[3]])
            except:
                item[field[3]] = ''
                worksheet.write_string(row, col + 3, item[field[3]])
            try:
                if item[field[4]]:
                    worksheet.write_string(row, col + 4, item[field[4]])
            except:
                item[field[4]] = ''
                worksheet.write_string(row, col + 4, item[field[4]])
            try:
                if item[field[5]]:
                    worksheet.write_string(row, col + 5, item[field[5]])
            except:
                item[field[5]] = ''
                worksheet.write_string(row, col + 5, item[field[5]])
            # worksheet.write_string(row, col + 6, item['wid'])
            # worksheet.write_string(row, col + 7, item['id'])
            row += 1
        # self.workbook.close()

    # 研发合作
    def cooperation(self):
        result = []
        field = ['date', 'title', 'author', 'org', 'ckey', 'abstract']
        url = 'http://webstads.sciinfo.cn/exportController.do?getCooperateList'
        page = 0
        while True:
            resp = requests.get(url=url + self.name.format(page), headers=self.headers, params=self.params)
            for i in range(3):
                try:
                    resp = resp.json()
                    print(resp)
                    break
                except:
                    print(resp.text)
                    continue
            resp = resp.replace('%', ';')
            resp = eval(resp)
            if not resp['result']:
                break
            for i in resp['result']:
                result.append(i)
            page += 1
            # time.sleep(0.5)
        self.generate_excel(result, '研发合作', field)

    # 科研成果
    def technology(self):
        result = []
        field = ['dop', 'title', 'au', 'orgc', 'fitclass', 'pte']
        url = 'http://webstads.sciinfo.cn/exportController.do?getExpertAchInfo'
        page = 0
        while True:
            resp = requests.get(url=url+self.name.format( page), headers=self.headers,params=self.params)
            for i in range(3):
                try:
                    resp = resp.json()
                    break
                except:
                    print(resp.text)
                    continue
            resp = resp.replace('%', ';')
            resp = eval(resp)
            if not resp['result']:
                break
            for i in resp['result']:
                result.append(i)
            page += 1
            # time.sleep(0.5)
        self.generate_excel(result, '科技成果', field)

    # 专利产出
    def patent(self):
        result = []
        field = ['reqno', 'annodate', 'title', 'reqpep', 'au', 'patt']
        url = 'http://webstads.sciinfo.cn/exportController.do?getExpertPatentInfo'
        page = 0
        while True:
            resp = requests.get(url=url+self.name.format( page), headers=self.headers,params=self.params)
            for i in range(3):
                try:
                    resp = resp.json()
                    print(resp)
                    break
                except:
                    print(resp.text)
                    continue
            resp = resp.replace('%', ';')
            resp = eval(resp)
            if not resp['result']:
                break
            for i in resp['result']:
                result.append(i)
            page += 1
            # time.sleep(0.5)
        self.generate_excel(result, '专利产出', field)

    # 人才培养
    def student(self):
        result = []
        field = ['date', 'author', 'title', 'degree', '', '']
        url = 'http://webstads.sciinfo.cn/exportController.do?getXueweiInfo'
        page = 0
        while True:
            resp = requests.get(url=url + self.name.format(page), headers=self.headers, params=self.params)
            for i in range(3):
                try:
                    resp = resp.json()
                    print(resp)
                    break
                except:
                    print(resp.text)
                    continue
            resp = resp.replace('%', ';')
            resp = eval(resp)
            if not resp['result']:
                break
            for i in resp['result']:
                result.append(i)
            page += 1
            # time.sleep(0.5)
        print(result)
        self.generate_excel(result, '人才培养', field)

    # 国内基础研究
    def research(self):
        result = []
        field = ['year', 'au', 'title', 'joucn', 'pg', 'per']
        url = 'http://webstads.sciinfo.cn/exportController.do?getOutInfo'
        page = 0
        while True:
            resp = requests.get(url=url + self.name.format(page), headers=self.headers, params=self.params)
            for i in range(3):
                try:
                    resp = resp.json()
                    print(resp)
                    break
                except:
                    print(resp.text)
                    continue
            resp = resp.replace('%', ';')
            resp = eval(resp)
            if not resp['result']:
                break
            for i in resp['result']:
                result.append(i)
            page += 1
            # time.sleep(0.5)
        self.generate_excel(result, '国内基础研究', field)

    def run(self):
        print('开始抓取专利产出')
        self.patent()
        # time.sleep(2)
        print('开始抓取研究成果')
        self.technology()
        # time.sleep(2)
        print('开始抓取研究合作')
        self.cooperation()
        # time.sleep(2)
        print('开始抓取人才培养')
        self.student()
        # time.sleep(2)
        print('开始抓取国内基础研究')
        self.research()
        self.workbook.close()


def read_xlrd():
    data = xlrd.open_workbook('机械运载学部2019版.xlsx')
    table = data.sheet_by_index(1)
    dataFile = []
    for rowNum in range(table.nrows):
        # if 去掉表头
        if rowNum > 0:
            dataFile.append(table.row_values(rowNum))

    return dataFile


if __name__ == '__main__':
    # excelFile = read_xlrd()
    # for name in excelFile:
    #     print(name[1])
        s = STSpider('&name=%25E9%25A9%25AC%25E4%25BC%259F%25E6%25B0%2591%2520OR%2520ma%2520weimin%2520OR%2520weimin%2520ma%2520OR%2520ma%2520w.m.%2520OR%2520w.m.%2520ma&auIds=L27904684&infoType=a&pageNo={}&pageSize=10&org=%25E4%25B8%25AD%25E5%259B%25BD%25E6%25B5%25B7%25E6%25B4%258B%25E5%25A4%25A7%25E5%25AD%25A6&ckeys=&type=',"马伟民")
        s.run()
