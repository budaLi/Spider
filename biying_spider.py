# @Time    : 2019/11/14 8:37
# @Author  : Libuda
# @FileName: google_spider.py
# @Software: PyCharm

from configparser import ConfigParser
from selenium import webdriver

config_parser = ConfigParser()
config_parser.read('config.cfg')
config = config_parser['default']
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('window-size=1200,1100')
browser = webdriver.Chrome(chrome_options=chrome_options, executable_path=config['executable_path'])
import xlrd
from queue import Queue
from xlutils.copy import copy  # 写入Excel

file_path = config['biying_datas']
from operationExcel import OperationExcel

res_count = 0


class Spider():
    def __init__(self):
        self.opExcel = OperationExcel(config['keywords_excel_path'], 0)
        self.keywords_queue = Queue()

    def get_keywords_data(self, row):
        """
        获取关键词数据
        :param row:
        :return:
        """
        actual_data = OperationExcel(config['keywords_excel_path'], 0).get_cel_value(row, 0)
        return actual_data

    def gen_keywords_queue(self):
        """
        关键词队列
        :return:
        """
        key_len = self.opExcel.get_nrows()
        print(key_len)
        for index in range(key_len):
            self.keywords_queue.put(self.get_keywords_data(index))
        print("关键词队列生成完毕")

    def write_to_excel(self, sheet_id, row, col, value):
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
        # 打开登录页
        count = 0

        while not self.keywords_queue.empty():
            key = self.keywords_queue.get()

            print("当前关键词", key)

            try:
                browser.get("https://cn.bing.com/?FORM=BEHPTB&ensearch=1")
                browser.find_element_by_css_selector("#sb_form_q").send_keys(key)
                browser.find_element_by_css_selector("#sb_form_go").click()
            except Exception as e:
                print(e)
                pass
            res = set()
            current_url_set = set()
            global df
            flag = True
            test_count = 3
            while flag:
                try:
                    if browser.current_url in current_url_set:
                        if test_count<0:
                            print("no next")
                            flag = False
                        else:
                            test_count-=1
                        continue
                    else:
                        current_url_set.add(browser.current_url)

                    title = browser.find_elements_by_css_selector("#b_results > li > h2")
                    url = browser.find_elements_by_css_selector('#b_results > li> h2 > a')
                    res_dic = {}

                    for i in range(len(url)):

                        s = url[i].get_attribute("href").split("/")
                        try:
                            tmp = s[0] + "//" + s[2]
                        except Exception as e:
                            print(e)
                            tmp = s[0] + "//" + s[2]
                        if tmp not in res:
                            res.add(tmp)
                            res_count += 1

                            self.write_to_excel(-1, count, 0, title[i].text)
                            self.write_to_excel(-1, count, 1, tmp)
                            print(count, title[i].text, tmp)
                            count += 1



                    next_paget = browser.find_element_by_css_selector(
                        "#b_results > li.b_pag > nav > ul > li:nth-child(7) > a")
                    next_paget.click()

                except Exception as e:
                    print(e)
                    try:
                        browser.find_element_by_css_selector("#b_results > li.b_pag > nav > ul > li:nth-child(7) > a")
                        continue
                    except Exception as e:
                        print("no next")
                        flag = False
            print("已获取：", str(res_count))


if __name__ == "__main__":
    spider = Spider()
    spider.gen_keywords_queue()
    spider.main()
