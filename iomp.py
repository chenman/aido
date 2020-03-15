# coding = utf-8

import urllib.request
import urllib.parse
import http.cookiejar
import pandas as pd
import datetime
import execjs
import os


def encrypt(text):
    """
    用户名、密码加密
    :param text: 明文
    :return: 经加密后的密文
    """
    os.environ["EXECJS_RUNTIME"] = "PhantomJS"
    node = execjs.get()
    # return execjs.compile(open(r"security.js", encoding='utf-8').read()).call('cmdEncrypt', text)
    ctx = node.compile(open("security.js", encoding='utf-8').read())
    js = 'cmdEncrypt("{0}")'.format(text)
    return ctx.eval(js)


def gen_function_xml(areaIds, serviceIds, woStates, omStates, startDate, endDate, DateType,
                     reportCode='orderNotRealTime', ttOrgId='10008663', staffId=223214, priv='infoHide',
                     searchType='OrdMoni', searchFlag='true', pageIndex=1, pageSize=10000):
    xml = '<?xml version="1.0" encoding="gb2312"?><requestdata>'
    xml += '<parameter name="reportCode">%s</parameter>' % reportCode
    xml += '<parameter name="ttOrgId">%s</parameter>' % ttOrgId
    xml += '<parameter name="staffId">%s</parameter>' % staffId
    xml += '<parameter name="priv">%s</parameter>' % priv
    xml += '<parameter name="searchType">%s</parameter>' % searchType
    xml += '<parameter name="searchFlag">%s</parameter>' % searchFlag
    xml += '<parameter name="pageIndex">%s</parameter>' % pageIndex
    xml += '<parameter name="pageSize">%s</parameter>' % pageSize
    xml += '<parameter name="areaIds">%s</parameter>' % areaIds
    xml += '<parameter name="serviceIds">%s</parameter>' % serviceIds
    if woStates is not None:
        xml += '<parameter name="woStates">%s</parameter>' % woStates
    xml += '<parameter name="omStates">%s</parameter>' % omStates
    xml += '<parameter name="startDate">%s</parameter>' % startDate
    xml += '<parameter name="endDate">%s</parameter>' % endDate
    xml += '<parameter name="DateType">%s</parameter>' % DateType
    xml += '</requestdata>'
    return xml


def login(user_name, password):
    headers = {
        'Accept': '*/*',
        'Accept-Language': 'zh-CN',
        'Referer': 'http://10.53.160.88',
        'Accept-Encoding': 'gzip, deflate',
        'User-Agent': 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; InfoPath.3; .NET4.0E; .NET4.0C)',
        'Host': '10.53.160.88:9001',
        'Connection': 'Keep-Alive',
        'Pragma': 'no-cache',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept-Encoding': 'identity'
    }
    url = 'http://10.53.160.88:9001/IOMPROJ/logonin'
    login_data = '''<?xml version="1.0"?>
        <Function name="login" serviceName="com.zterc.web.WebLoginManager" userTransaction="true">
        <Param type="s">{R}%s</Param>
        <Param type="s">{R}%s</Param>
        <Param type="s">zh_cn</Param></Function>''' % (encrypt(user_name), encrypt(password))

    request = urllib.request.Request(url, data=login_data.encode('utf8'), headers=headers)
    filename = 'cookie.txt'
    # 使用http.cookiejar.CookieJar()创建CookieJar对象
    cjar = http.cookiejar.CookieJar()
    # 使用HTTPCookieProcessor创建cookie处理器，并以其为参数构建opener对象
    cookie = urllib.request.HTTPCookieProcessor(cjar)
    opener = urllib.request.build_opener(cookie)
    # 将opener安装为全局
    urllib.request.install_opener(opener)

    try:
        response = urllib.request.urlopen(request)
    except urllib.error.HTTPError as e:
        print(e.code)
        print(e.reason)


def search_download(areaIds, serviceIds, omStates, startDate, endDate, DateType, fileName, woStates=None):
    url = 'http://10.53.160.88:9001/IOMPROJ/FaultExportForwardServlet'
    headers = {
        'Accept': 'image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, '
                  'application/x-ms-xbap, */*',
        'Accept-Language': 'zh-CN',
        'Referer': 'http://10.53.160.88:9001/IOMPROJ/order/orderNotRealTimeSearch.jsp',
        'Accept-Encoding': 'gzip, deflate',
        'User-Agent': 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/7.0; SLCC2; .NET CLR '
                      '2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; InfoPath.3; .NET4.0E; .NET4.0C)',
        'Host': '10.53.160.88:9001',
        'Connection': 'Keep-Alive',
        'Pragma': 'no-cache',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept-Encoding': 'gzip, deflate'
    }

    xml = gen_function_xml(areaIds, serviceIds, woStates, omStates, startDate, endDate, DateType)
    # print(xml)
    post_data = ('functionXml=%s&forwardServlet=1' % urllib.parse.quote_plus(xml)).encode('utf-8')
    request = urllib.request.Request(url=url, data=post_data, headers=headers)

    try:
        response = urllib.request.urlopen(request)
        # print(response.read().decode('GBK'))
        out = response.read()
        with open(fileName, "wb") as code:
            code.write(out)
    except urllib.error.HTTPError as e:
        print(e.code)
        print(e.reason)


def get_current_finish():
    """
    查询当前竣工量及待办量
    :return:
    """
    current = datetime.datetime.now()
    endDate = current.strftime('%Y-%m-%d %H:%M:%S')

    current_day = datetime.datetime(current.year, current.month, current.day)
    startDate = current_day.strftime('%Y-%m-%d %H:%M:%S')

    fileName = '%s-finish.xlsx' % (current.strftime('%Y%m%d%H%M%S'))

    # <parameter name="areaIds">4,102,41,42,43,44,45</parameter>
    search_download(areaIds='43', serviceIds='220323', omStates='10F',
                    startDate=startDate, endDate=endDate, DateType=2, fileName=fileName)
    df = pd.read_excel(fileName, sheet_name='全量工单信息')
    print('\n%s - %s 竣工量：' % (startDate, endDate))
    print(df['工单状态'].groupby(df['区县']).count())
    print(df[['区县', '所属网格', '工单状态']].groupby(['区县', '所属网格']).count().sort_values(by=['区县', '所属网格'],
                                                                                 ascending=False).reset_index())

    startDate = (current_day - datetime.timedelta(days=60)).strftime('%Y-%m-%d %H:%M:%S')

    fileName = '%s-todo.xlsx' % (current.strftime('%Y%m%d%H%M%S'))
    search_download(areaIds='43', serviceIds='220323', woStates='101,102,103,104,301,303', omStates='10H',
                    startDate=startDate, endDate=endDate, DateType=1, fileName=fileName)
    df = pd.read_excel(fileName, sheet_name='全量工单信息')
    print('\n%s - %s 待办量：' % (startDate, endDate))
    print(df['工单状态'].groupby(df['区县']).count())
    print(df[['区县', '所属网格', '工单状态']].groupby(['区县', '所属网格']).count().sort_values(by=['区县', '所属网格'],
                                                                                 ascending=False).reset_index())


def get_yesterday_finish():
    """
    查询昨日竣工量
    :return:
    """
    current = datetime.datetime.now()
    current_day = datetime.datetime(current.year, current.month, current.day)
    startDate = (current_day - datetime.timedelta(days=1)).strftime('%Y-%m-%d %H:%M:%S')
    endDate = (current_day - datetime.timedelta(seconds=1)).strftime('%Y-%m-%d %H:%M:%S')

    fileName = '%s-finish.xlsx' % ((current_day - datetime.timedelta(seconds=1)).strftime('%Y%m%d%H%M%S'))

    # <parameter name="areaIds">4,102,41,42,43,44,45</parameter>
    search_download(areaIds='43', serviceIds='220323', omStates='10F',
                    startDate=startDate, endDate=endDate, DateType=2, fileName=fileName)
    df = pd.read_excel(fileName, sheet_name='全量工单信息')
    print('\n%s - %s 竣工量：' % (startDate, endDate))
    print(df['工单状态'].groupby(df['区县']).count())
    print(df[['区县', '所属网格', '工单状态']].groupby(['区县', '所属网格']).count().sort_values(by=['区县', '所属网格'],
                                                                                 ascending=False).reset_index())


if __name__ == '__main__':
    # 模拟登录
    login('chenman', 'HB6q$*2y')
    # get_current_finish()
    get_yesterday_finish()
