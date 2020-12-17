# 首先得登录上，记住请求的时候用post，然后拿到cookies
import requests
import time
import html
import json
from bs4 import BeautifulSoup
from xlrd import open_workbook
from xlutils.copy import copy
import xlsxwriter
from urllib import request
from http import cookiejar

# 国抽
from requests_html import HTMLSession

session = HTMLSession()

# 获取获取检验数据列表  需在网页登陆后获取
pageUrl = r'http://spcjinsp.gsxt.gov.cn/test_platform/api/food/getFood?order=desc&offset=0&limit=10000&dataType=5&startDate=2020-01-01&endDate=2020-12-17&taskFrom=&samplingUnit=&testUnit=&enterprise=&sampledUnit=&foodName=&province=&reportNo=&bsfla=&bsflb=&sampleNo=&foodType1=&foodType4=&sampleNo_index=0&_=1608190838853'
# 获取普通检验数据详情
infoUrl = ''
infoUrl1 = r'http://spcjinsp.gsxt.gov.cn/test_platform/foodTest/foodDetail/%s'
# 获取农产品检验数据详情
infoUrl2 = r'http://spcjinsp.gsxt.gov.cn/test_platform/agricultureTest/agricultureDetail/%s'

infoUrl1_1=r'http://spcjinsp.gsxt.gov.cn/test_platform/api/food/getTestInfo'

infoUrl2_1=r'http://spcjinsp.gsxt.gov.cn/test_platform/api/agriculture/getTestInfo'
# 请求头信息 需在网页登陆后获取 修改Cookie即可
headers_1 = {

    'Cookie': 'JSESSIONID=C78B1243DCA3CE69C5C8CA90744518B5-n3; sod=x8O9MNQd6rzatjS4+oPQyFDlCEzR8iTja0D5Qs11+0uKhDmGbf0wW4FT2eCxk661EA4iDNnPfh8='
}
excel_name = ''
colNum = 0
rowNum = 0


def getData(infoUrl):
    # 有密码的请求一定要用post()
    response = session.get(infoUrl, headers=headers_1)
    html_doc = response.text
    soup = BeautifulSoup(html_doc, "lxml")
    # print(soup.prettify())
    divs = soup.find_all('div', class_="ibox float-e-margins")
    # print(len(divs))
    address = []
    for div in divs:
        if div.text.find('抽检场所信息') >= 0 and div.text.find('所在地：') >= 0:
            textStr = div.text
            startNum = textStr.find("所在地：")
            endNum = textStr.find("区域类型：")
            address.append(textStr[startNum + 4:endNum].strip())
        elif div.text.find('生产企业信息') >= 0 and div.text.find('所在地：') >= 0:
            textStr = div.text
            startNum = textStr.find("所在地：")
            endNum = textStr.find("企业地址：")
            address.append(textStr[startNum + 4:endNum].strip())
    return address


def getAllData(infoUrl):
    # 有密码的请求一定要用post()
    # print(infoUrl)
    response = session.get(infoUrl, headers=headers_1)
    html_doc = response.text
    soup = BeautifulSoup(html_doc, "lxml")
    # print(soup.prettify())
    # form = soup.find('form', id="testform")
    sd = soup.find('input', id="sd")["value"]
    # print(sd)
    data={'sd':sd};
    if infoUrl.find("api/agriculture/getAgriculture") > -1:
        productInfoUrl = infoUrl2_1
    else:
        productInfoUrl = infoUrl1_1

    # rowlist=[]

    # print(getProductInfo(productInfoUrl, data)["rows"])
    return  getProductInfo(productInfoUrl, data)["rows"]


#
def getAllPageData(pageUrl):
    response = session.get(pageUrl, headers=headers_1)
    infoJson = json.loads(response.text)  # 先把字典转成json
    infoList = []
    num = 0

    for info in infoJson['rows']:
        # oneList = [info['id'], info['sp_s_16']]
        # print(info)
        oneList = [info['sp_s_16']]
        oneList.append(getAllData(infoUrl % info['id']))
        # print(oneList)
        infoList.append(oneList)

    return infoList


def getProductInfo(url,data):
    response = session.post(url, data,headers=headers_1)
    infoJson = json.loads(response.text)  # 先把字典转成json
    # print(infoJson)
    return infoJson


def writeContent2Excel(sheet, infoAllData):
    colNum = 1
    for infoData in infoAllData:
        for info in infoData[1]:
            # print(info)
            sheet.write(colNum, 0, infoData[0])
            sheet.write(colNum, 1, info["spdata_0"])
            sheet.write(colNum, 2, info["spdata_1"])
            sheet.write(colNum, 3,info["spdata_18"])
            sheet.write(colNum, 4,info["spdata_2"])
            sheet.write(colNum, 5, info["spdata_3"])
            sheet.write(colNum, 6, info["spdata_4"])
            sheet.write(colNum, 7, info["spdata_11"])
            sheet.write(colNum, 8, info["spdata_15"])
            sheet.write(colNum, 9, info["spdata_16"])
            sheet.write(colNum, 10, info["spdata_7"])
            sheet.write(colNum, 11, info["spdata_20"])
            sheet.write(colNum, 12, info["spdata_17"])
            colNum += 1



'''获取标题头名称
'''
def getMaxKey(type, infoList, max):
    retInfoKey = set([])
    for info in infoList:
        for info1 in info[1][type].keys():
            retInfoKey.add(info1)
    return list(retInfoKey)

def getdate():
    return time.strftime("%Y-%m-%d", time.localtime())

if __name__ == "__main__":
    # getCookie()
    print("======执行中，请等待======")
    if pageUrl.find("api/agriculture/getAgriculture") > -1:
        infoUrl = infoUrl2
        excel_name = getdate() + '市县级农产品检测信息检验结果统计'
    else:
        infoUrl = infoUrl1
        excel_name = getdate() + '普通食品检测信息检验结果统计'
    book = xlsxwriter.Workbook(excel_name+'.xlsx')
    sheet = book.add_worksheet('sheet1')

    sheet.write(0, 0, '单号')
    sheet.write(0, 1, '检验项目')
    sheet.write(0, 2, '检验结果')
    sheet.write(0, 3, '结果单位')
    sheet.write(0, 4, '结果判定')
    sheet.write(0, 5, '检验依据')
    sheet.write(0, 6, '判定依据')
    sheet.write(0, 7, '最小允许限')
    sheet.write(0, 8, '最大允许限')
    sheet.write(0, 9, '允许限单位')
    sheet.write(0, 10, '方法检出限')
    sheet.write(0, 11, '检出限单位')
    sheet.write(0, 12, '备注')
    infoList = getAllPageData(pageUrl)
    # print(infoList)
    writeContent2Excel(sheet,infoList)
    book.close()

    # print(getAllData(infoUrl))
    print("======执行完毕，请查看文件======")
