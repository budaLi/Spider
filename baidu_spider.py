# @Time    : 2020/2/24 17:29
# @Author  : Libuda
# @FileName: baidu_spider.py
# @Software: PyCharm
from configparser import ConfigParser
from selenium import webdriver
import time
import xlrd
from queue import Queue
from xlutils.copy import copy  # 写入Excel

config_parser = ConfigParser()
config_parser.read('config.cfg', encoding="utf-8-sig")
config = config_parser['default']
browser = webdriver.PhantomJS(executable_path=config['executable_path'])
guolv_uls = ['edu.cn','edu.com','gov.com','alibaba','1688','www.1688.com','1688.com','job','job.com','www.sz.gov.cn','.ORG',
       '.EDU','GOV','ALIBABA','EC21','B2B','.PDF','ARTICLE','dictionary',
    'shop','company.look','51job.com','gongcha','b2b.','www.zhipin.com']

from operationExcel import OperationExcel

res_count = 0


class Spider():
    def __init__(self):
        self.opExcel = OperationExcel(config['keywords_excel_path'], 0)
        self.file_path = config['baidu_datas']
        # self.pass_key_excel = OperationExcel(config['pass_key_path'],0)
        self.dataExcel = OperationExcel(self.file_path, 0)
        self.keywords_queue = Queue()
        self.res = set()



    def get_keywords_data(self, row):
        """
        获取关键词数据
        :param row:
        :return:
        """
        actual_data = OperationExcel(config['keywords_excel_path'], 0).get_cel_value(row, 0)
        return actual_data

    def write_to_excel(self, file_path, sheet_id, row, col, value):
        """
        写入Excel
        :param sheet_id:
        :param row:
        :param col:
        :param value:
        :return:
        """
        work_book = xlrd.open_workbook(file_path, formatting_info=False)
        # 先通过xlutils.copy下copy复制Excel
        write_to_work = copy(work_book)
        # 通过sheet_by_index没有write方法 而get_sheet有write方法
        sheet_data = write_to_work.get_sheet(sheet_id)
        sheet_data.write(row, col, str(value))
        # 这里要注意保存 可是会将原来的Excel覆盖 样式消失
        write_to_work.save(file_path)

    def main(self):
        global res_count
        test_count = int(config['max_test_count'])
        last_count = 0
        count = self.dataExcel.tables.nrows
        print("当前已有url数量：", count)
        key_len = self.opExcel.get_nrows()
        print("关键词总数：", key_len)

        for index in range(1, key_len):

            key = self.get_keywords_data(index)

            try:
                print("启动中。。。。，如果20s内没有启动 请重新启动本软件")
                browser.get("https://www.baidu.com/")
                browser.find_element_by_css_selector("#kw").send_keys(key)
                browser.find_element_by_css_selector("#su").click()

            except Exception as e:
                print("a",e)

            page = 1
            current_url_set = set()
            flag = True
            while flag:
                try:
                    print("当前正在采集第 {} 个关键词:{}，采集的页数为 :{} ".format((index + 1), key, page))
                    print("当前url", browser.current_url)

                    title = browser.find_elements_by_css_selector('.t')
                    url = browser.find_elements_by_css_selector('.f13')

                    lenght = min(len(title),len(url))

                    for i in range(lenght):
                        tmp = url[i].text.split(" ")[0]
                        if not tmp.startswith("http") and not tmp.startswith("www"):
                            continue
                        #过滤 www.innuo-instruments....类似
                        elif tmp.endswith("."):
                            continue
                        elif tmp.startswith("http") :
                            tmp = tmp.split("/")
                            try:
                                tmp = tmp[0] + "//" + tmp[2]
                            except Exception as e:
                                tmp = tmp[0] + "//" + tmp[2]
                        elif tmp.startswith("www"):
                            tmp = tmp.split("/")
                            try:
                                tmp = tmp[0]
                            except Exception as e:
                                print(e)
                                pass

                        else:
                            pass
                            # print(tmp,"lllllllll")



                        #过滤指定url
                        try:
                            for one in guolv_uls:
                                if one in tmp:
                                    tmp= None
                        except Exception as e :
                            pass
                            # print("12",e)

                        if tmp and tmp not in self.res:
                            self.res.add(tmp)
                            try:
                                self.write_to_excel(self.file_path, -1, count, 0, title[i].text)
                                self.write_to_excel(self.file_path, -1, count, 1, tmp)
                                print("{}:{},{}".format(count,title[i].text,tmp))
                                count += 1
                                res_count += 1
                            except Exception as e:
                                pass

                    try:
                        time.sleep(5)
                        next_page = browser.find_element_by_xpath("//*[contains(text(), '下一页')]")
                        next_page.click()
                        page += 1
                    except Exception as e:
                        time.sleep(5)
                        try:
                            next_page = browser.find_element_by_xpath("//*[contains(text(), '下一页')]")
                            next_page.click()
                            page += 1
                        except Exception as e:
                            time.sleep(5)
                            try:
                                next_page = browser.find_element_by_xpath("//*[contains(text(), '下一页')]")
                                next_page.click()
                                page += 1
                            except Exception as e:
                                flag = False
                except Exception as e:
                    # print("123",e)
                    pass

            try:

                print("当前关键词 ：{} 爬取完毕 已爬取数据 ：{}".format(key, res_count - last_count))
            except Exception as e:
                pass

            print("本次采集已获取url总数为：", str(res_count))
            last_count = res_count
        print("关键词搜索完毕，谢谢使用!")
        while 1:
            pass

if __name__ == "__main__":
    print("欢迎使用 （百度客户搜索系统）")
    spider = Spider()
    spider.main()

