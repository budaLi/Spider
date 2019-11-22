# @Time    : 2019/11/14 8:37
# @Author  : Libuda
# @FileName: google_spider.py
# @Software: PyCharm

from configparser import ConfigParser
from selenium import webdriver
import time
import base64
import xlrd
from queue import Queue
from xlutils.copy import copy  # 写入Excel


config_parser = ConfigParser()
config_parser.read('config.cfg')
config = config_parser['default']
chrome_options = webdriver.ChromeOptions()
browser=webdriver.Chrome(executable_path=config['executable_path'])
if config['is_have_chrome'] == "0":
    print("正在使用无界面爬取。。。")
    prefs = {"profile.managed_default_content_settings.images": 2}  #禁止图片加载
    chrome_options.add_experimental_option("prefs", prefs)
    #chrome_options.add_argument('user-agent="Mozilla/5.0 (iPhone; CPU iPhone OS 9_1 like Mac OS X) AppleWebKit/601.1.46 (KHTML, like Gecko) Version/9.0 Mobile/13B143 Safari/601.1"')
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('window-size=1200,1100')
    browser = webdriver.Chrome(chrome_options=chrome_options, executable_path=config['executable_path'])


file_path = config['biying_datas']
from operationExcel import OperationExcel

res_count = 0


class Spider():
    def __init__(self):
        self.opExcel = OperationExcel(config['keywords_excel_path'], 0)
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


    def write_to_excel(self, sheet_id, row, col, value):
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
        global res_count,browser
        count = 0
        key_len = self.opExcel.get_nrows()
        print("关键词总数：",key_len)
        for index in range(key_len):
            key = self.get_keywords_data(index)
            print("当前为第 {} 个关键词:{}".format(index+1, key))
            try:
                print("启动中。。。。，如果20s内没有启动 请重新启动本软件")
                browser.get("https://cn.bing.com/?FORM=BEHPTB&ensearch=1")
                browser.find_element_by_css_selector("#sb_form_q").send_keys(key)
                browser.find_element_by_css_selector("#sb_form_go").click()

                for i in range(20):
                    if browser.current_url !="https://cn.bing.com/?FORM=BEHPTB&ensearch=1":
                        continue
                    else:
                        print(20-i)
                        time.sleep(1)
                        print("正在第{}次尝试自动启动。。。。。".format(i+1))
                        browser.get("https://cn.bing.com/?FORM=BEHPTB&ensearch=1")
                        browser.find_element_by_css_selector("#sb_form_q").send_keys(key)
                        browser.find_element_by_css_selector("#sb_form_go").click()
            except Exception as e:
                print("正在尝试自动启动。。。。。")
                browser.get("https://cn.bing.com/?FORM=BEHPTB&ensearch=1")
                browser.find_element_by_css_selector("#sb_form_q").send_keys(key)
                browser.find_element_by_css_selector("#sb_form_go").click()

            current_url_set = set()
            flag = True
            while flag:
                try:
                    if browser.current_url in current_url_set:
                        if config["max_test_count"]<0:
                            print("no next")
                            flag = False
                        else:
                            print("当前url {} 可能为最后一页,进行第{}次测试".format(browser.current_url,config["max_test_count"]))
                            config["max_test_count"]-=1
                        continue
                    else:
                        current_url_set.add(browser.current_url)
						
                    title = browser.find_elements_by_css_selector("#b_results > li > h2")
                    url = browser.find_elements_by_css_selector('#b_results > li> h2 > a')
                    for i in range(len(url)):

                        s = url[i].get_attribute("href").split("/")
                        try:
                            tmp = s[0] + "//" + s[2]
                        except Exception as e:
                            print(e)
                            tmp = s[0] + "//" + s[2]
                        if tmp not in self.res:
                            self.res.add(tmp)
                            try:
                                self.write_to_excel(-1, count, 0, title[i].text)
                                self.write_to_excel(-1, count, 1, tmp)
                                print(count, title[i].text, tmp)
                                count += 1
                                res_count += 1
                            except Exception:
                                print("请关闭Excel 否则10秒后本条数据将不再写入")
                                for i in range(10):
                                    print(10 - i)
                                    time.sleep(1)
                                try:
                                    self.write_to_excel(-1, count, 0, title[i].text)
                                    self.write_to_excel(-1, count, 1, tmp)
                                    print(count, title[i].text, tmp,browser.current_url)
                                except Exception:
                                    print("已漏掉数据...{}  {}".format(title[i].text,tmp))

                    try:
                        next_paget = browser.find_element_by_css_selector("#b_results > li.b_pag > nav > ul > li:nth-child(9) > a")
                        next_paget.click()
                    except Exception:
                        next_paget = browser.find_element_by_css_selector("#b_results > li.b_pag > nav > ul > li:nth-child(7) > a")
                        next_paget.click()

                except Exception:
                    try:
                        try:
                            next_paget = browser.find_element_by_css_selector(
                                "#b_results > li.b_pag > nav > ul > li:nth-child(9) > a")
                            next_paget.click()
                        except Exception:
                            next_paget = browser.find_element_by_css_selector(
                                "#b_results > li.b_pag > nav > ul > li:nth-child(7) > a")
                            next_paget.click()
                            print("找不到下一页呢")
                            time.sleep(5)
                            flag= False
                    except Exception:
                        print("可能是最后一页了呢 当前url为{}".format(browser.current_url))
                        time.sleep(5)
                        flag= False
            print("已获取：", str(res_count))


if __name__ == "__main__":
    try:
        code = config['code']
        now_time = int(time.time())
        s = str(base64.b64decode(code), "utf-8")
        s=  time.strptime(s, "%Y-%m-%d %H:%M:%S")
        time_sti = int(time.mktime(s)) #时间戳
        if now_time>time_sti:
            print("您的注册码已过期 请联系管理员")
        else:
            print("======欢迎使用 必应爬虫======")
            spider = Spider()
            spider.main()
    except Exception as e:
        print(e)
        print("您的注册码格式有误")
