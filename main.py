# @Time    : 2019/10/16 17:09
# @Author  : Libuda
# @FileName: main.py
# @Software: PyCharm
from selenium import webdriver
import time
import xlrd
from xlutils.copy import copy  # 写入Excel
file_path = r"C:\Users\lenovo\PycharmProjects\Spider\data.xls"


def write_to_excel(row, col,value):
    work_book = xlrd.open_workbook(file_path, formatting_info=False)
    # 先通过xlutils.copy下copy复制Excel
    write_to_work = copy(work_book)
    # 通过sheet_by_index没有write方法 而get_sheet有write方法
    sheet_data = write_to_work.get_sheet(0)
    sheet_data.write(row, col, str(value))
    # 这里要注意保存 可是会将原来的Excel覆盖 样式消失
    write_to_work.save(file_path)

def main():
    res= []
    browser = webdriver.Chrome(executable_path=r"C:\Users\lenovo\PycharmProjects\Spider\chromedriver.exe")

    #打开登录页
    browser.get("https://passport.zhaopin.com/org/login")
    #点击短信登录
    browser.find_elements_by_css_selector(".k-tabs__item")[1].click()
    #输入手机号
    browser.find_element_by_css_selector(".zp-passport-widget-b-login-sms__number").send_keys(
        "18234116562")
    #点击发送验证码
    browser.find_element_by_css_selector(".zp-passport-widget-b-login-sms__send-code").click()
    #等待手动输入
    time.sleep(20)

    try:
        # 点击登录
        browser.find_element_by_css_selector(".zp-passport-widget-b-login-sms__submit").click()
    except Exception:
        time.sleep(20)
        browser.find_element_by_css_selector(".zp-passport-widget-b-login-sms__submit").click()

    print("等待点击")
    time.sleep(10)
    #点击全部投递简历
    browser.find_element_by_css_selector("body > div.rd55-header > div.rd55-header__nav > div > ul:nth-child(2) > li:nth-child(3) > ul > li:nth-child(2) > a").click()
    #点击有意向
    browser.find_element_by_css_selector("#root > div.app-container.resume-apply > div.k-tabs.k-tabs--border-card.resume-tabs > div.k-tabs__header > div > div.k-tabs__nav > div.k-tabs__item.is-active").click()
    #当前主窗口
    current_handle = browser.current_window_handle

    #点击用户名 多个
    for one in browser.find_elements_by_css_selector(".user-name"):
        one.click()
        for handle in browser.window_handles:
            if handle!=current_handle:
                browser.switch_to_window(handle)
                #姓名
                name = browser.find_element_by_css_selector("").text
                #性别
                sex = browser.find_element_by_css_selector("#resume-detail-wrapper > div.resume-content.is-mb-0 > div.resume-content__section > div > div.resume-content__candidate-basic > p.resume-content__labels > span:nth-child(1)").text
                #岗位
                gangwei = browser.find_element_by_css_selector("#root > div.app-container.resume-detail.resume-detail--medium > div.resume-detail__main.resume-detail__structure > div.resume-content > dl > dd").text
                #手机号
                phone = browser.find_element_by_css_selector("#resume-detail-wrapper > div.resume-content.is-mb-0 > div.resume-content__section > div > div.resume-content__candidate-basic > div.resume-content__status-box > a.resume-content__button.is-primary.is-static > p.resume-content__mobile-phone > span").text
                #邮箱
                email = browser.find_element_by_css_selector("#resume-detail-wrapper > div.resume-content.is-mb-0 > div.resume-content__section > div > div.resume-content__candidate-basic > div.resume-content__status-box > a.resume-content__button.is-primary.is-static > p.resume-content__email > span").text
                #关闭当前窗口切回主窗口
                tem = [name,sex,gangwei,phone,email]
                res.append(tem)
                browser.close()
                browser.switch_to_window(current_handle)
    return res


if __name__=="__main__":
    res= main()
    #将结果写入数据
    # res=[["吉广阔",'男','测试','17512514571',' xiaokuo21@163.com'],["吉广阔",'男','测试','17512514571',' xiaokuo21@163.com']]
    for index,datas in enumerate(res):
        for j,data in enumerate(datas):
            write_to_excel(index,j,data)