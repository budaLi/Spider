# @Time    : 2019/11/14 8:37
# @Author  : Libuda
# @FileName: main3.py
# @Software: PyCharm

from selenium import webdriver
browser = webdriver.Chrome(executable_path=r"C:\Users\lenovo\PycharmProjects\Spider\chromedriver.exe")
import xlrd
from xlutils.copy import copy  # 写入Excel
file_path = r"C:\Users\lenovo\PycharmProjects\Spider\data.xls"
from operationExcel import OperationExcel
class Spider():
    def __init__(self):
        self.opExcel = OperationExcel(r"C:\Users\lenovo\PycharmProjects\Spider\keywords.xls",0)
    def get_keywords_data(self, row):
        actual_data = OperationExcel(r"C:\Users\lenovo\PycharmProjects\Spider\keywords.xls",0).get_cel_value(row, 0)
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
                    self.opExcel = OperationExcel(r"C:\Users\lenovo\PycharmProjects\Spider\data.xls",0)
                    self.opExcel.create_sheet(key)
            except Exception as e:
                print("已有该Excel")

            try:
                browser.get("https://www.google.com/")

                browser.find_element_by_css_selector("#tsf > div:nth-child(2) > div.A8SBwf > div.RNNXgb > div > div.a4bIc > input").send_keys(key)

                browser.find_element_by_css_selector("#tsf > div:nth-child(2) > div.A8SBwf > div.FPdoLc.VlcLAe > center > input.gNO89b").click()
            except Exception as e:
                pass
            while 1:
                res_set = set()
                try:
                    title = browser.find_elements_by_css_selector(".S3Uucc")
                    url = browser.find_elements_by_xpath('//*[@class="r"]/a')
                    for i in range(len(url)):
                        s = url[i].get_attribute("href").split("/")
                        tmp =s[0]+"//"+s[1]+s[2]
                        if tmp not in res_set:
                            res_set.add(tmp)
                            # print(title[i].text,tmp)
                            self.write_to_excel(index,count, 0, title[i].text)
                            self.write_to_excel(index,count, 1,tmp)
                            count+=1
                    next_paget = browser.find_element_by_css_selector("#pnnext > span:nth-child(2)")
                    next_paget.click()
                except Exception as e:
                    test_count =3
                    while test_count>0:
                        try:
                            next_paget = browser.find_element_by_css_selector("#pnnext > span:nth-child(2)")
                            next_paget.click()
                        except Exception as e:
                            test_count-=1
                            continue
                    break


if __name__=="__main__":
    spider = Spider()
    spider.main()

