#!/usr/bin/env python
# -*- coding: utf-8 -*-

import json
import os
from _md5 import md5

import requests
import re
from requests.exceptions import RequestException
from pyquery import PyQuery as pq
import pymongo

from config import *

import xlwt

wb = xlwt.Workbook()
sh = wb.add_sheet('豆瓣top250')
sh.write(0, 0, '排名')
sh.write(0, 1, '电影名')
sh.write(0, 2, '评分')
sh.write(0, 3, '下载地址')
sh.write(0, 4, '介绍')

LINE = 1

client = pymongo.MongoClient(MONGO_URL, connect=False)
db = client[MONGO_BD]


def get_one_page(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            return response.text
        return None
    except RequestException:
        return None


def parse_one_page(html):
    global LINE

    a = 0
    doc = pq(html)
    # 排名，海报，标题（主标题&连接，副标题），介绍，评分
    length = doc('body > div.container > div.container-fluid > div.row').length
    lists = doc('body > div.container > div.container-fluid > div.row').items()
    while a < length:
        for list in lists:
            rank = list('div.row > div:nth-child(1)').text()
            image = list('div.row > div.col-xs-8 > div > div.col-xs-2 > img').attr.src
            download_image(image)
            main_title = list('div.row > div.col-xs-8 > div > div.col-xs-9 > h4 > a').text()
            title_link = 'http://www.id97.com' + list('div.row > div.col-xs-8 > div > div.col-xs-9 > h4 > a').attr.href
            subhead = list('div.row > div.col-xs-8 > div > div.col-xs-9 > h4').text()[len(main_title) + 1:]
            introduce = list('div.row > div.col-xs-8 > div > div.col-xs-9 > p').text()
            grade = list('div.row > div:nth-child(3)').text()
            info = {
                'rank': rank,
                'image': image,
                'title': [main_title, title_link, subhead],
                'introduce': introduce,
                'grade': grade
            }
            save_info_mongo(info)

            sh.write(LINE, 0, info['rank'] or '')
            sh.write(LINE, 1, info['title'][0] or '')
            sh.write(LINE, 2, info['grade'] or '')
            sh.write(LINE, 3, info['title'][1] or '')
            sh.write(LINE, 4, info['introduce'] or '')
            LINE = LINE + 1
        a += 1


def save_info_mongo(result):
    if db[MONGO_TABLE].insert(result):
        print('存储到mongoDB成功', result)
        return True
    return False


def download_image(url):
    print('正在下载:' + url)
    try:
        response = requests.get(url)
        if response.status_code == 200:
            save_image(response.content)
        return None
    except RequestException:
        print("请求图片出错", url)
        return None


def save_image(content):
    file_path = '{0}/{1}.{2}'.format(os.getcwd() + IMAGE_PATH, md5(content).hexdigest(), 'jpg')
    if not os.path.exists(file_path):
        with open(file_path, 'wb') as f:
            f.write(content)
            f.close()


def main(page):
    if page == 1:
        url = 'http://www.id97.com/movie/top250_douban'
    else:
        url = 'http://www.id97.com/movie/top250_douban?page=' + str(page)
    html = get_one_page(url)
    if html:
        parse_one_page(html)


if __name__ == '__main__':
    for x in range(1, 11):
        main(x)
    wb.save('豆瓣电影top250.xls')
