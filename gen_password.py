# @Time    : 2019/11/22 11:57
# @Author  : Libuda
# @FileName: gen_password.py
# @Software: PyCharm

import time
import base64

if __name__ == '__main__':
    try:
        end_date = input("请输入过期时间 如 2019-11-22 12:00:00:")
        # end_date = "2019-11-22 12:00:00"
        start_date = str(int(time.mktime(time.strptime(end_date,"%Y-%m-%d %H:%M:%S"))))
        #加密结束时间
        code = str(base64.b64encode(end_date.encode("utf-8")),"utf-8")
        print("注册码为:",code)
        for i in range(10):
            print("10s内此窗口关闭")
            time.sleep(1)
    except Exception as e:
        print("时间格式有误")