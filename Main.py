import time
import datetime
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

TYPE_KPT = 4  # LoanCategoryId 4:平衡型,8:保守型,5:进取型
file_name = 'data.xlsx'  # 存储数据文件名
today = datetime.date.today()  # 启动date
now = datetime.datetime.now()  # 启动datetime
TYPE_KPT_MAP = {4: '平衡型', 8: '保守型', 5: '进取型'}  # 类型映射Map


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
    # 整个信息列表所在的div
    div_outer_borrow_list = soup.find('div', attrs={'class': 'outerBorrowList'})
    if '很抱歉，热门列表已经被抢空啦' in str(div_outer_borrow_list):
        return []
    else:
        # 详情信息所在的a标签
        a_s = div_outer_borrow_list.find_all('a', attrs={'class': 'title ell'})
        details_urls = []
        for a in a_s:
            details_urls.append(a['href'])
        return details_urls


# 详情页信息提取
def details_info_getter(details_url):
    soup = html_to_soup(details_url)

    # 整个信息列表所在的div
    wrap_new_lend_detail = soup.find('div', attrs={'class': 'wrapNewLendDetailInfoLeft'})

    # 包含信用等级的a [信用等级详细值在其中span标签的class属性中]
    credit_level = wrap_new_lend_detail.find('a', attrs={'class': 'altQust'}).find('span')['class'] \
        .replace('creditRating ', '')
    # 用户名
    user_name = wrap_new_lend_detail.find('span', attrs={'class': 'username'}).get_text()

    # 赔标/信用标
    pei_i = soup.find('i', attrs={'class': 'pei'})

    # 完成度
    progress = soup.find('div', attrs={'class': 'newLendDetailRefundLeft'}) \
        .find('div', attrs={'class': 'part clearfix'}) \
        .find('div').get_text().replace('进度条：', '').strip()

    # 借款的详细信息 [借款金额 , 协议利率 , 期限]
    new_lend_detail_money_list = soup.find('div', attrs={'class': 'newLendDetailMoneyLeft'}).find_all('dd')

    # 人员详细信息 [性别 , 年龄 , 注册时间 , 文化程度 , 毕业院校 , 学习形式 , 借款用途 , 还款来源 , 工作信息 , 收入情况]
    lender_info_list = soup.find('div', attrs={'class': 'lender-info'}).findAll('p', attrs={'class': 'ex col-1'})
    # 信息可能不完全存在 , 需要特殊处理
    sex = '--'  # 性别
    age = '--'  # 年龄
    regist_date = '--'  # 注册时间
    edu_bg = '--'  # 文化程度
    graduate = '--'  # 毕业院校
    learn_way = '--'  # 学习形式
    lend_purpose = '--'  # 借款用途
    payment = '--'  # 还款来源
    work_info = '--'  # 工作信息
    income_info = '--'  # 收入情况
    for it in lender_info_list:
        if '性别' in str(it):
            sex = it.find('span').get_text()
        elif '年龄' in str(it):
            age = it.find('span').get_text()
        elif '注册时间' in str(it):
            regist_date = it.find('span').get_text()
        elif '文化程度' in str(it):
            edu_bg = it.find('span').get_text()
        elif '毕业院校' in str(it):
            graduate = it.find('span').get_text()
        elif '学习形式' in str(it):
            learn_way = it.find('span').get_text()
        elif '借款用途' in str(it):
            lend_purpose = it.find('span').get_text()
        elif '还款来源' in str(it):
            payment = it.find('span').get_text()
        elif '工作信息' in str(it):
            work_info = it.find('span').get_text()
        elif '收入情况' in str(it):
            income_info = it.find('span').get_text()

    # 认证信息
    record_info_list = soup.find('ul', attrs={'class', 'record-info'}).findAll('li')
    verfied_info = ''
    for i in range(len(record_info_list)):
        verfied_info += (record_info_list[i].get_text()) + ' '
    verfied_info = verfied_info.strip()

    # 总体div
    tab_contain_divs = soup.find("div", attrs={'class': 'lendDetailTab_tabContent w1000center'}) \
        .find_all('div', attrs={'class': 'tab-contain'})

    # 统计信息
    statistics_info_list = tab_contain_divs[2].find_all('span', attrs={'class', 'num'})

    # 投资人情况List
    ol_list_parent = soup.find('div', attrs={'class': 'scroll-area'})
    ol_list = []
    if ol_list_parent is not None:
        ol_list = ol_list_parent.find_all('ol')
    else:
        print('当前页面暂无投资人信息')
    # 投资人信息列表
    investor_list = []
    # 投资人单条信息
    for ol in ol_list:
        # 单条投资人信息
        investor_info = {}
        li_list = ol.find_all('li')
        investor_info['investor_id'] = li_list[0].find('span', attrs={'class': 'listname'}).get_text()  # 投资人id
        investment_way_tag = li_list[0].find('a', attrs={'target': '_blank'})  # 投资方式 [根据图标区分 , 可能没有]
        if 'title' in str(investment_way_tag):
            investor_info['investment_way'] = investment_way_tag['title']
        else:
            investor_info['investment_way'] = '未知'

        investor_info['rate'] = li_list[1].get_text()  # 利率
        investor_info['term'] = li_list[2].get_text()  # 期限
        investor_info['valid_amount'] = li_list[3].get_text().replace('¥', '')  # 有效金额
        investor_info['investment_date'] = li_list[4].get_text()  # 投标时间
        investor_list.append(investor_info)

    # 将信息放入字典中
    result_dic = {}  # 返回爬取结果字典
    result_dic['risk_level'] = TYPE_KPT_MAP[TYPE_KPT]  # 风险等级
    if pei_i is not None:  # 赔标 [存在 `赔`图标 即视为`赔标`]
        result_dic['pei'] = '赔标'
    else:
        result_dic['pei'] = '--'
    result_dic['progress'] = progress  # 完成度
    result_dic['user_name'] = user_name  # 用户名
    result_dic['credit_level'] = credit_level  # 信用等级
    result_dic['amount'] = new_lend_detail_money_list[0].get_text()  # 贷款金额
    result_dic['rate'] = new_lend_detail_money_list[1].get_text()  # 协议利率
    result_dic['term'] = new_lend_detail_money_list[2].get_text()  # 期限
    result_dic['sex'] = sex  # 性别
    result_dic['age'] = age  # 年龄
    result_dic['regist_date'] = regist_date  # 注册时间
    result_dic['edu_bg'] = edu_bg  # 文化程度
    result_dic['graduate'] = graduate  # 毕业院校
    result_dic['learn_way'] = learn_way  # 学习形式
    result_dic['lend_purpose'] = lend_purpose  # 借款用途
    result_dic['payment'] = payment  # 还款来源
    result_dic['work_info'] = work_info  # 工作信息
    result_dic['income_info'] = income_info  # 收入情况
    result_dic['verfied_info'] = verfied_info  # 认证信息
    result_dic['sucuess_cnt'] = statistics_info_list[0].get_text().strip()  # 成功借款次数
    result_dic['history_info'] = statistics_info_list[1].get_text().strip()  # 历史记录
    result_dic['sucuess_repayment_cnt'] = statistics_info_list[2].get_text().strip()  # 成功还款次数
    result_dic['normal_repayment_cnt'] = statistics_info_list[3].get_text().strip()  # 正常还清次数
    result_dic['delay_lt15_repayment_cnt'] = statistics_info_list[4].get_text().strip()  # 逾期(0-15天)还清次数
    result_dic['delay_gt15_repayment_cnt'] = statistics_info_list[5].get_text().strip()  # 逾期(15天以上)还清次数
    result_dic['amount_sum'] = statistics_info_list[6].get_text().strip().replace('¥', '')  # 累计借款金额
    result_dic['unreturned_amount'] = statistics_info_list[7].get_text().strip().replace('¥', '')  # 待还金额
    result_dic['unreceived_amount'] = statistics_info_list[8].get_text().strip().replace('¥', '')  # 待收金额
    result_dic['biggest_lend_amount'] = statistics_info_list[9].get_text().strip().replace('¥', '')  # 单笔最高借款金额
    result_dic['biggest_debt_amount'] = statistics_info_list[10].get_text().strip().replace('¥', '')  # 历史最高负债
    result_dic['investor_list'] = investor_list  # 投资信息列表
    return result_dic


# 获取总页数 [用于修正]
def total_page_getter(url):
    soup = html_to_soup(url)
    page_info = soup.find('span', attrs={'class': 'pagerstatus'}).replace('共', '').replace('页', '').strip()
    if page_info is None:  # 可能出现不存在的情况
        total_page = -1
    else:
        total_page = int(page_info, base=10)
    return total_page


# 爬取逻辑
def data_spider(total_page=100):
    current_page = 1
    data_list = []
    while current_page < total_page:
        url = url_constructor(current_page, TYPE_KPT)
        details_url_list = details_url_list_getter(url)
        if len(details_url_list) > 0:
            print('爬取第' + str(current_page) + '页数据。')
            for it in details_url_list:
                url = 'https:' + it
                print('爬取链接:' + url)
                data_list.append(details_info_getter(url))
            data_output_xls(data_list)  # 输出数据
            time.sleep(1)
        else:
            print('Error:未找到数据!\n尝试重新获取...')
            current_page = 1
            total_page = total_page_getter(url)
            if total_page == -1:  # 如果返回-1 则表示当前没有数据,休眠后继续尝试
                time.sleep(10)
                total_page = 100
        current_page += 1


# 输出数据到excel
def data_output_xls(data_list):
    print('数据输出开始....')
    wb = Workbook()
    title = "拍拍贷数据" + str(today)
    # 标题行
    work_sheet = wb.create_sheet(title=title)
    _ = work_sheet.cell(column=1, row=1, value="%s" % '风险等级')
    _ = work_sheet.cell(column=2, row=1, value="%s" % '赔标/信用标')
    _ = work_sheet.cell(column=3, row=1, value="%s" % '完成度')
    _ = work_sheet.cell(column=4, row=1, value="%s" % '用户名')
    _ = work_sheet.cell(column=5, row=1, value="%s" % '信用等级')
    _ = work_sheet.cell(column=6, row=1, value="%s" % '贷款金额')
    _ = work_sheet.cell(column=7, row=1, value="%s" % '利率')
    _ = work_sheet.cell(column=8, row=1, value="%s" % '还款期限')
    _ = work_sheet.cell(column=9, row=1, value="%s" % '性别')
    _ = work_sheet.cell(column=10, row=1, value="%s" % '年龄')
    _ = work_sheet.cell(column=11, row=1, value="%s" % '文化程度')
    _ = work_sheet.cell(column=12, row=1, value="%s" % '学习形式')
    _ = work_sheet.cell(column=13, row=1, value="%s" % '借款用途')
    _ = work_sheet.cell(column=14, row=1, value="%s" % '还款来源')
    _ = work_sheet.cell(column=15, row=1, value="%s" % '工作信息')
    _ = work_sheet.cell(column=16, row=1, value="%s" % '收入情况')
    _ = work_sheet.cell(column=17, row=1, value="%s" % '认证状况')
    _ = work_sheet.cell(column=18, row=1, value="%s" % '成功借款次数')
    _ = work_sheet.cell(column=19, row=1, value="%s" % '历史记录')
    _ = work_sheet.cell(column=20, row=1, value="%s" % '逾期(0-15天)还清次数')
    _ = work_sheet.cell(column=21, row=1, value="%s" % '逾期(15天以上)还清次数')
    _ = work_sheet.cell(column=22, row=1, value="%s" % '累计借贷金额')
    _ = work_sheet.cell(column=23, row=1, value="%s" % '待还金额')
    _ = work_sheet.cell(column=24, row=1, value="%s" % '待收金额')
    _ = work_sheet.cell(column=25, row=1, value="%s" % '历史最高负债')
    _ = work_sheet.cell(column=26, row=1, value="%s" % '单笔最高借款金额')
    _ = work_sheet.cell(column=27, row=1, value="%s" % '投标人名称')
    _ = work_sheet.cell(column=28, row=1, value="%s" % '投资方式')
    _ = work_sheet.cell(column=29, row=1, value="%s" % '有效投标金额')
    _ = work_sheet.cell(column=30, row=1, value="%s" % ' 投标日期')
    row = 2
    # 数据行
    for it in data_list:
        investor_list_size = len(it['investor_list'])  # 投资人信息数量
        investor_list_output_flag = True
        if investor_list_size == 0:
            investor_list_size = 1
            investor_list_output_flag = False
        for i in range(investor_list_size):
            _ = work_sheet.cell(column=1, row=row, value="%s" % it['risk_level'])  # 风险等级
            _ = work_sheet.cell(column=2, row=row, value="%s" % it['pei'])  # 赔标/信用标
            _ = work_sheet.cell(column=3, row=row, value="%s" % it['progress'])  # 完成度
            _ = work_sheet.cell(column=4, row=row, value="%s" % it['user_name'])  # 用户名
            _ = work_sheet.cell(column=5, row=row, value="%s" % it['credit_level'])  # 信用等级
            _ = work_sheet.cell(column=6, row=row, value="%s" % it['amount'])  # 贷款金额
            _ = work_sheet.cell(column=7, row=row, value="%s" % it['rate'])  # 利率
            _ = work_sheet.cell(column=8, row=row, value="%s" % it['term'])  # 还款期限
            _ = work_sheet.cell(column=9, row=row, value="%s" % it['sex'])  # 性别
            _ = work_sheet.cell(column=10, row=row, value="%s" % it['age'])  # 年龄
            _ = work_sheet.cell(column=11, row=row, value="%s" % it['edu_bg'])  # 文化程度
            _ = work_sheet.cell(column=12, row=row, value="%s" % it['learn_way'])  # 学习形式
            _ = work_sheet.cell(column=13, row=row, value="%s" % it['lend_purpose'])  # 借款用途
            _ = work_sheet.cell(column=14, row=row, value="%s" % it['payment'])  # 还款来源
            _ = work_sheet.cell(column=15, row=row, value="%s" % it['work_info'])  # 工作信息
            _ = work_sheet.cell(column=16, row=row, value="%s" % it['income_info'])  # 收入情况
            _ = work_sheet.cell(column=17, row=row, value="%s" % it['verfied_info'])  # 认证状况
            _ = work_sheet.cell(column=18, row=row, value="%s" % it['sucuess_cnt'])  # 成功借款次数
            _ = work_sheet.cell(column=19, row=row, value="%s" % it['history_info'])  # 历史记录
            _ = work_sheet.cell(column=20, row=row, value="%s" % it['delay_lt15_repayment_cnt'])  # 逾期(0-15天)还清次数
            _ = work_sheet.cell(column=21, row=row, value="%s" % it['delay_gt15_repayment_cnt'])  # 逾期(15天以上)还清次数
            _ = work_sheet.cell(column=22, row=row, value="%s" % it['amount_sum'])  # 累计借贷金额
            _ = work_sheet.cell(column=23, row=row, value="%s" % it['unreturned_amount'])  # 待还金额
            _ = work_sheet.cell(column=24, row=row, value="%s" % it['unreceived_amount'])  # 待收金额
            _ = work_sheet.cell(column=25, row=row, value="%s" % it['biggest_lend_amount'])  # 历史最高负债
            _ = work_sheet.cell(column=26, row=row, value="%s" % it['biggest_debt_amount'])  # 单笔最高借款金额
            if investor_list_output_flag:
                _ = work_sheet.cell(column=27, row=row, value="%s" % it['investor_list'][i]['investor_id'])  # 投标人名称
                _ = work_sheet.cell(column=28, row=row, value="%s" % it['investor_list'][i]['investment_way'])  # 投资方式
                _ = work_sheet.cell(column=29, row=row, value="%s" % it['investor_list'][i]['valid_amount'])  # 有效投标金额
                _ = work_sheet.cell(column=30, row=row, value="%s" % it['investor_list'][i]['investment_date'])  # 投标日期
            row += 1
    wb.save(filename=file_name)
    print('数据输出完成....')


# Main method
if __name__ == '__main__':
    data_spider()
