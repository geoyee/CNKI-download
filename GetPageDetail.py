"""
-------------------------------------------------
   File Name：     GetPageDetail.py
   Description :   获取文献摘要等信息存储至excel
   Author :        Cyrus_Ren
   date：          2018/12/12
-------------------------------------------------
   Change Activity:
                   
-------------------------------------------------
"""
__author__ = 'Cyrus_Ren'

import xlwt
from bs4 import BeautifulSoup
from GetConfig import config
import re
import math, random
from GetConfig import config

HEADER = config.crawl_headers


class PageDetail(object):
    def __init__(self):
        # count用于计数excel行
        self.excel = xlwt.Workbook(encoding='utf8')
        self.sheet = self.excel.add_sheet('文献列表', True)
        self.set_style()
        self.sheet.write(0, 0, '序号', self.basic_style)
        self.sheet.write(0, 1, '题名', self.basic_style)
        self.sheet.write(0, 2, '摘要', self.basic_style)
        if config.crawl_isDownLoadLink == '1':
            self.sheet.write(0, 3, '下载地址', self.basic_style)

        # 生成userKey,服务器不做验证
        self.cnkiUserKey = self.set_new_guid()

    def get_detail_page(self, session, result_url, page_url,
                        single_refence_list, download_url):
        '''
        发送三次请求
        前两次服务器注册 最后一次正式跳转
        '''
        # 这个header必须设置
        HEADER['Referer'] = result_url
        self.single_refence_list = single_refence_list
        self.session = session
        self.session.cookies.set('cnkiUserKey', self.cnkiUserKey)
        self.download_url = download_url
        cur_url_pattern_compile = re.compile(
            r'.*?FileName=(.*?)&.*?DbCode=(.*?)&')
        cur_url_set = re.search(cur_url_pattern_compile, page_url)
        # 前两次请求需要的验证参数
        params = {
            'curUrl': 'detail.aspx?dbCode=' + cur_url_set.group(2) + '&fileName=' + cur_url_set.group(1),
            'referUrl': result_url + '#J_ORDER&',
            'cnkiUserKey': self.session.cookies['cnkiUserKey'],
            'action': 'file',
            'userName': '',
            'td': '1544605318654'
        }
        # 首先向服务器发送两次预请求
        self.session.get(
            'https://i.shufang.cnki.net/KRS/KRSWriteHandler.ashx',
            headers=HEADER,
            params=params,
            verify=False)
        self.session.get(
            'https://kns.cnki.net/KRS/KRSWriteHandler.ashx',
            headers=HEADER,
            params=params,
            verify=False)
        page_url = 'https://kns.cnki.net' + page_url
        get_res = self.session.get(page_url, headers=HEADER)
        self.pars_page(get_res.text)
        self.excel.save('data/Reference_detail.xls')

    def pars_page(self, detail_page):
        '''
        解析页面信息
        '''
        soup = BeautifulSoup(detail_page, 'lxml')
        # 获取摘要
        if soup.find(name='span', id='ChDivSummary'):
            abstract_list = soup.find(name='span', id='ChDivSummary').strings
        else:
            abstract_list = '无摘要'
        self.abstract = ''
        for a in abstract_list:
            self.abstract += a
        self.wtire_excel()

    def create_list(self):
        '''
        整理excel每一行的数据
        序号 题名 摘要
        '''
        self.reference_list = []
        self.reference_list.append(self.single_refence_list[0])  # 编号
        self.reference_list.append(self.single_refence_list[1].replace("网络首发", ""))  # 书名
        self.reference_list.append(self.abstract)  # 摘要
        if config.crawl_isDownLoadLink == '1':
            self.reference_list.append(self.download_url)

    def wtire_excel(self):
        '''
        将获得的数据写入到excel
        '''
        self.create_list()
        if config.crawl_isDownLoadLink == '1':
            for i in range(0, 4):
                self.sheet.write(int(self.reference_list[0]), i, self.reference_list[i], self.basic_style)
        else:
            for i in range(0, 3):
                self.sheet.write(int(self.reference_list[0]), i, self.reference_list[i], self.basic_style)

    def set_style(self):
        '''
        设置excel样式
        '''
        self.sheet.col(1).width = 256 * 30
        self.sheet.col(2).width = 256 * 15
        self.sheet.col(3).width = 256 * 60
        self.sheet.col(4).width = 256 * 15
        self.sheet.row(0).height_mismatch = True
        self.sheet.row(0).height = 20 * 20
        self.basic_style = xlwt.XFStyle()
        al = xlwt.Alignment()
        # 垂直对齐
        al.horz = al.HORZ_CENTER
        # 水平对齐
        al.vert = al.VERT_CENTER
        # 换行
        al.wrap = al.WRAP_AT_RIGHT
        # 设置边框
        borders = xlwt.Borders()
        borders.left = 6
        borders.right = 6
        borders.top = 6
        borders.bottom = 6

        self.basic_style.alignment = al
        self.basic_style.borders = borders

    def set_new_guid(self):
        '''
        生成用户秘钥
        '''
        guid = ''
        for i in range(1, 32):
            n = str(format(math.floor(random.random() * 16.0), 'x'))
            guid += n
            if (i == 8) or (i == 12) or (i == 16) or (i == 20):
                guid += "-"
        return guid


# 实例化
page_detail = PageDetail()
