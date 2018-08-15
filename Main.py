import datetime
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

file_name = 'data.xlsx'  # 存储数据文件名
today = datetime.date.today()  # 启动date


# html转换
def html_to_soup(url):
    r = requests.get(url)
    html = r.content
    return BeautifulSoup(html, 'xml')


# url构造器 获得爬取链接
def url_constructor(page_index, _type):
    base_url = 'https://invest.ppdai.com/loan/' \
               'listnew?LoanCategoryId=' + str(_type) + \
               '&PageIndex=' + str(page_index) + \
               '&SortType=0&MinAmount=0&MaxAmount=0'
    aim_url = base_url
    return aim_url


# 获取链接页面所包含的详情页链接并封装到List
def details_url_list_getter(url):
    soup = html_to_soup(url)

    details_urls = []

    # TODO 获取列表页内容

    return details_urls


# 详情页信息提取
def details_info_getter(details_url):
    soup = html_to_soup(details_url)

    # TODO 从详情页面提取信息

    # 将信息放入字典中
    result_dic = {}  # 返回爬取结果字典
    return result_dic


# 获取总页数 [用于修正]
def total_page_getter(url):
    print('获取总页数')


# 爬取逻辑
def data_spider(total_page=100):
    print('获取页码,循环构造页面链接')


# 输出数据到excel
def data_output_xls(data_list):
    print('数据输出开始....')
    wb = Workbook()
    title = "Hentai_" + str(today)
    # 标题行
    work_sheet = wb.create_sheet(title=title)
    _ = work_sheet.cell(column=1, row=1, value="%s" % '风险等级')

    row = 2
    # 数据行
    for it in data_list:
        investor_list_size = len(it['investor_list'])  # 投资人信息数量
        for i in range(investor_list_size):
            _ = work_sheet.cell(column=1, row=row, value="%s" % it['risk_level'])  # 风险等级
            row += 1
    wb.save(filename=file_name)
    print('数据输出完成....')


# Main method
if __name__ == '__main__':
    data_spider()
