# @Time    : 2019/10/18 8:25
# @Author  : Libuda
# @FileName: main2.py
# @Software: PyCharm
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as Ec
from selenium.webdriver.common.by import By
import xlrd
from xlutils.copy import copy  # 写入Excel
browser = webdriver.Chrome(executable_path=r"C:\Users\lenovo\PycharmProjects\Spider\chromedriver.exe")
file_path = r"C:\Users\lenovo\PycharmProjects\Spider\data.xls"

def main():
    res = []
    browser.get("https://passport.zhaopin.com/org/login")
    while 1:
        current_url = browser.current_url
        # 首先判断当前url是否为主url
        if current_url=="":
            # 判断勾选简历是否完成
            try:
                WebDriverWait(browser,10).until(Ec.element_to_be_selected((By.LINK_TEXT, 'CSDN')))
            except Exception as e:
                print("请勾选全部简历")
                continue
            #开始逐个点击显示简历进行页面解析
            print("开始点击简历")
            browser.find_element_by_css_selector(
                "#apply-resume-list > div.fixable-list__footer-wrapper > div.fixable-list__footer.affix-bottom > div > div.resume-action > div.resume-action__buttons > a:nth-child(2)").click()
            #切换窗口
            current_handle = browser.current_window_handle
            for one in browser.window_handles:
                if one != current_handle:
                    browser.switch_to.window(one)
                    print("解析数据")
                    # 姓名
                    name = browser.find_element_by_css_selector(
                        "#resume-detail-wrapper > div.resume-content.is-mb-0 > div.resume-content__section > div > div.resume-content__candidate-basic > div.resume-content__candidate-header > span.resume-content__candidate-name.resume-tomb").text
                    # 性别
                    sex = browser.find_element_by_css_selector(
                        "#resume-detail-wrapper > div.resume-content.is-mb-0 > div.resume-content__section > div > div.resume-content__candidate-basic > p.resume-content__labels > span:nth-child(1)").text
                    # 岗位
                    gangwei = browser.find_element_by_css_selector(
                        "#root > div.app-container.resume-detail.resume-detail--medium > div.resume-detail__main.resume-detail__structure > div.resume-content > dl > dd").text
                    # 手机号
                    phone = browser.find_element_by_css_selector(
                        "#resume-detail-wrapper > div.resume-content.is-mb-0 > div.resume-content__section > div > div.resume-content__candidate-basic > div.resume-content__status-box > a.resume-content__button.is-primary.is-static > p.resume-content__mobile-phone > span").text
                    # 邮箱
                    email = browser.find_element_by_css_selector(
                        "#resume-detail-wrapper > div.resume-content.is-mb-0 > div.resume-content__section > div > div.resume-content__candidate-basic > div.resume-content__status-box > a.resume-content__button.is-primary.is-static > p.resume-content__email > span").text
                    # 关闭当前窗口切回主窗口
                    tem = [name, sex, gangwei, phone, email]
                    res.append(tem)

                    return res

def write_to_excel(row, col,value):
    work_book = xlrd.open_workbook(file_path, formatting_info=False)
    # 先通过xlutils.copy下copy复制Excel
    write_to_work = copy(work_book)
    # 通过sheet_by_index没有write方法 而get_sheet有write方法
    sheet_data = write_to_work.get_sheet(0)
    sheet_data.write(row, col, str(value))
    # 这里要注意保存 可是会将原来的Excel覆盖 样式消失
    write_to_work.save(file_path)


if __name__=="__main__":
    res = main()
