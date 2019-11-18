# @Time    : 2019/11/14 8:37
# @Author  : Libuda
# @FileName: google_spider.py
# @Software: PyCharm
import time
from configparser import ConfigParser
from selenium import webdriver
config_parser = ConfigParser()
config_parser.read('config.cfg')
config = config_parser['default']

browser = webdriver.Chrome(executable_path=config['executable_path'])
import xlrd
from xlutils.copy import copy  # 写入Excel
file_path = config['biying_datas']
from operationExcel import OperationExcel
class Spider():
    def __init__(self):
        self.opExcel = OperationExcel(config['keywords_excel_path'],0)
    def get_keywords_data(self, row):
        actual_data = OperationExcel(config['keywords_excel_path'],0).get_cel_value(row, 0)
        return actual_data

    def write_to_excel(self,sheet_id,row, col,value):
        work_book = xlrd.open_workbook(file_path, formatting_info=False)
        # 先通过xlutils.copy下copy复制Excel
        write_to_work = copy(work_book)
        # 通过sheet_by_index没有write方法 而get_sheet有write方法
        sheet_data = write_to_work.get_sheet(sheet_id)
        sheet_data.write(row, col, str(value))
        # 这里要注意保存 可是会将原来的Excel覆盖 样式消失
        write_to_work.save(file_path)

    def main(self):
        #打开登录页
        key_len  = self.opExcel.get_nrows()

        for index in range(key_len):
            count = 0
            key = self.get_keywords_data(index)
            print(key)
            try:
                if index==0:
                    pass
                else:
                    self.opExcel = OperationExcel(config['biying_datas'],0)
                    self.opExcel.create_sheet(key)
            except Exception as e:
                print("已有该Excel")

            try:
                browser.get("https://cn.bing.com/?FORM=BEHPTB&ensearch=1")

                browser.find_element_by_css_selector("#sb_form_q").send_keys(key)
                time.sleep(5)
                browser.find_element_by_css_selector("#sb_form_go").click()
            except Exception as e:
                print(e)
                pass
            while 1:
                res_set = set()
                try:
                    print("开始解析")
                    title = browser.find_elements_by_css_selector("#b_results > li > h2")
                    url = browser.find_elements_by_css_selector('#b_results > li> h2 > a')

                    for i in range(len(url)):
                        s = url[i].get_attribute("href").split("/")
                        try:
                            tmp =s[0]+"//"+s[2]
                        except Exception as e:
                            print(e)
                            tmp= s[0]+"//"+s[2]

                        if tmp not in res_set and title[i].text!="":
                            res_set.add(tmp)
                            print(title[i].text,tmp)
                            self.write_to_excel(index,count, 0, title[i].text)
                            self.write_to_excel(index,count, 1,tmp)
                            count+=1
                    time.sleep(5)
                    next_paget = browser.find_element_by_css_selector("#b_results > li.b_pag > nav > ul > li:nth-child(7) > a")
                    next_paget.click()
                    time.sleep(1)
                except Exception as e:
                    print(e)
                    try:
                        browser.find_element_by_css_selector("#b_results > li.b_pag > nav > ul > li:nth-child(7) > a")
                        continue
                    except Exception as e:
                        print("no next")
                        break


if __name__=="__main__":
    spider = Spider()
    spider.main()

