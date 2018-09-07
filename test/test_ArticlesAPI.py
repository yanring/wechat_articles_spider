# coding: utf-8
import os
import sys
from pprint import pprint

import time
import xlrd as xlrd

sys.path.append(os.getcwd())
from wechatarticles.ReadOutfile import Reader
from wechatarticles import ArticlesAPI, tools, ArticlesUrls

if __name__ == '__main__':
    official_cookie_1 = 1
    token_1 = 1
    official_cookie_2 = 1
    token_2 = "*"
    cookie_list = [[official_cookie_1, token_1], [official_cookie_2, token_2]]
    # excel
    data = xlrd.open_workbook(r'/home/yanring/Project/wechat_articles_spider/wechat.xlsx')
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    # time.sleep(600)
    for row in range(18,nrows):
        if 1:
            test = ArticlesUrls(cookie=cookie_list[1][0], token=cookie_list[1][1])
        else:
            test = ArticlesUrls(cookie=cookie_list[0][0], token=cookie_list[0][1])

        wechat_name = table.row(row)[0].value
        articles_sum,first_publish_time = test.articles_total_nums(wechat_name)
        # table.put_cell(row, 5, 2, articles_sum, 0)
        # table.put_cell(row, 6, 1, first_publish_time, 0)
        with open(r'/home/yanring/Project/wechat_articles_spider/data.txt','a') as f:
            f.write('%d,%s,%s,\n'%(articles_sum,first_publish_time,wechat_name))
        print(wechat_name,' ',articles_sum)
    # 实例化爬取对象
    # 账号密码自动获取cookie和token
    # test = ArticlesUrls(username=username, password=password)
    # 手动输入账号密码

    # test = ArticlesUrls(cookie=official_cookie, token=token)

    # 输入公众号名称，获取公众号文章总数
    # articles_sum = test.articles_total_nums('闵行档案')
    # print(articles_sum)
