# coding = utf-8

import datetime
import execjs
import os


def encrypt(text):
    os.environ["EXECJS_RUNTIME"] = "PhantomJS"
    node = execjs.get()
    # return execjs.compile(open(r"security.js", encoding='utf-8').read()).call('cmdEncrypt', text)
    ctx = node.compile(open("security.js", encoding='utf-8').read())
    js = 'cmdEncrypt("{0}")'.format(text)
    return ctx.eval(js)


if __name__ == '__main__':
    print(encrypt('chenman'))
    # current = datetime.datetime.now()
    # print(current.strftime('%Y-%m-%d %H:%M:%S'))
    # current_day = datetime.datetime(current.year, current.month, current.day)
    # print(current_day.strftime('%Y-%m-%d %H:%M:%S'))
    # # td = datetime.datetime.strptime(stime[:8], "%Y%m%d")  # 当天
    # month_fd = datetime.datetime(current.year, current.month, 1)  # 月初
    #
    # lst_month_ld = month_fd - datetime.timedelta(days=1)
    # lst_month_fd = datetime.datetime(lst_month_ld.year, lst_month_ld.month, 1)  # 月初
    # print(lst_month_fd.strftime('%Y-%m-%d %H:%M:%S'))
    # # print(urllib.parse.unquote_plus('functionXml: %3C%3Fxml+version%3D%21.0%22+encoding%3D%22gb2312%22%3F%3E%3Crequestdata%3E%3Cparameter+name%3D%22reportCode%22%3EorderNotRealTime%3C%2Fparameter%3E%3Cparameter+name%3D%22ttOrgId%22%3E10000000%3C%2Fparameter%3E%3Cparameter+name%3D%22priv%22%3EinfoHide%3C%2Fparameter%3E%3Cparameter+name%3D%22areaIds%22%3E4%2C102%2C41%2C42%2C43%2C44%2C45%3C%2Fparameter%3E%3Cparameter+name%3D%22serviceIds%22%3E220323%2C220372%3C%2Fparameter%3E%3Cparameter+name%3D%22woStates%22%3E101%2C102%2C103%2C104%2C301%2C303%3C%2Fparameter%3E%3Cparameter+name%3D%22omStates%22%3E10H%3C%2Fparameter%3E%3Cparameter+name%3D%22searchType%22%3EOrdMoni%3C%2Fparameter%3E%3Cparameter+name%3D%22startDate%22%3E2020-02-25+00%3A00%3A00%3C%2Fparameter%3E%3Cparameter+name%3D%22endDate%22%3E2020-03-04+18%3A57%3A19%3C%2Fparameter%3E%3Cparameter+name%3D%22DateType%22%3E3%3C%2Fparameter%3E%3Cparameter+name%3D%22searchFlag%22%3Etrue%3C%2Fparameter%3E%3Cparameter+name%3D%22staffId%22%3E223214%3C%2Fparameter%3E%3Cparameter+name%3D%22pageIndex%22%3E1%3C%2Fparameter%3E%3Cparameter+name%3D%22pageSize%22%3E10000%3C%2Fparameter%3E%3C%2Frequestdata%3E'))