# coding = utf-8

import urllib.request
import urllib.parse
import http.cookiejar
import pandas as pd


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


def login():
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
        <Param type="s">{R}158a10da9f94b367ceb870b92ce6e69a58918e69dde9c17cc861d6970ee245034b62d2e16cc347fea184b33af36e06fdc7b63f8c9317f5d1d420892e4918f3ad2bb0ff57fb7d9ed7b51078c76c708f6b6baf5e165e5812d7181c553984ef44bbf1a117567b74a610a5071dbb68b7143b73552f145627ef55ee37b25a70ac0542</Param>
        <Param type="s">{R}1983d032f55a206359067368a0fd8507ceaf0b00a7233b04cd6916e8bd9a09377c2cbce5ed78a182bc9c4851fafb8581d20da80d84622e516b345201124a6229e3b3b98ba16fd5ee554b7327adb7ed0d6c2d282978d31dafd76fa3310123f144ca8d2fb3037aa930f06b0caa8c0242724287acd0c7eedfaeb57cb3f6d6bfc401</Param>
        <Param type="s">zh_cn</Param></Function>'''

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


def search_download(areaIds, serviceIds, omStates, startDate, endDate, DateType, fileName, woStates = None):
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


if __name__ == '__main__':

    # 模拟登录
    login()

    fileName = '20200303-finish.xlsx'

    # <parameter name="areaIds">4,102,41,42,43,44,45</parameter>
    search_download(areaIds='43', serviceIds='220323', omStates='10F',
                    startDate='2020-03-03 00:00:00', endDate='2020-03-04 00:00:00', DateType=2, fileName=fileName)
    df = pd.read_excel(fileName, sheet_name='全量工单信息')
    print(df['工单状态'].groupby(df['区县']).count())
    print(df[['区县', '所属网格', '工单状态']].groupby(['区县', '所属网格']).count().reset_index())

    fileName = '20200303-todo.xlsx'
    search_download(areaIds='43', serviceIds='220323', woStates='101,102,103,104,301,303', omStates='10H',
                    startDate='2020-01-01 00:00:00', endDate='2020-03-04 00:00:00', DateType=1, fileName=fileName)
    df = pd.read_excel(fileName, sheet_name='全量工单信息')
    print(df['工单状态'].groupby(df['区县']).count())
    print(df[['区县', '所属网格', '工单状态']].groupby(['区县', '所属网格']).count().reset_index())
