# @Time    : 2019/11/25 13:19
# @Author  : Libuda
# @FileName: test_spider.py
# @Software: PyCharm
import pandas
from configparser import ConfigParser
from selenium import webdriver
import time
import base64
import threading
import requests
from queue import Queue
from lxml import etree
config_parser = ConfigParser()
config_parser.read('config.cfg')
config = config_parser['default']

class Spider:
    def __init__(self):
        self.pd = pandas.DataFrame()
        self.base_url = 'https://cn.bing.com/search?q={}&ensearch=1&first={}'
        # self.keywords = pandas.read_csv(r"C:\Users\lenovo\PycharmProjects\Spider\biying_data.xls",encoding="utf8")
        # self.biying_data_csv = pandas.read_csv(config['biying_datas'],encoding="utf-8")
        self.keyword_queue = Queue()
        self.key_urls = Queue()

    def gen_key_queue(self):
        self.keyword_queue.put("aa")

    def get_domain(self,url):
        """
        解析主域名
        :param url:
        :return:
        """
        try:
            tmp = url[0] + "//" + url[2]
        except Exception as e:
            tmp = url[0] + "//" + url[2]

        return tmp
    def gen_key_urls(self):
        """
        根据关键词生产url
        :return:
        """
        while not self.keyword_queue.empty():
            key = self.keyword_queue.get()
            num =1
            flag = True
            while flag:
                if num>100:
                    flag=False
                else:
                    self.key_urls.put(self.base_url.format(key,num))
                    num+=14

    def get_html_from_url(self):
        """
        从url中解析出内容
        :return:
        """
        while not self.key_urls.empty():
            url = self.key_urls.get()
            print("当前Url",url)
            response = requests.get(url)
            print(response.text)
            html = etree.HTML(response.text)
            url = html.xpath('//*[@id="b_results"]/li/h2/a//@href')
            for one in url:
                print(one)

    def main(self,thread_num=1):
        threads = []
        S.gen_key_queue()
        S.gen_key_urls()
        for i in range(thread_num):
            threads.append(threading.Thread(self.get_html_from_url))
        for one in threads:
            one.start()
        for one in threads:
            one.join()

if __name__ == '__main__':
    S = Spider()
    S.gen_key_queue()
    S.gen_key_urls()
    S.get_html_from_url()