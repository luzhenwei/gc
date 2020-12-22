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
pageUrl = r'http://spcjinsp.gsxt.gov.cn/test_platform/api/agriculture/getAgriculture?order=desc&offset=0&limit=10000&dataType=5&startDate=2020-01-01&endDate=2020-12-18&taskFrom=&samplingUnit=&testUnit=&enterprise=&sampledUnit=&foodName=&province=&reportNo=&bsfla=&bsflb=&sampleNo=&foodType1=&foodType4=&_=1608282850126'
# 获取普通检验数据详情
infoUrl = ''
infoUrl1 = r'http://spcjinsp.gsxt.gov.cn/test_platform/foodTest/foodDetail/%s'
# 获取农产品检验数据详情
infoUrl2 = r'http://spcjinsp.gsxt.gov.cn/test_platform/agricultureTest/agricultureDetail/%s'

infoUrl1_1=r'http://spcjinsp.gsxt.gov.cn/test_platform/api/food/getTestInfo'

infoUrl2_1=r'http://spcjinsp.gsxt.gov.cn/test_platform/api/agriculture/getTestInfo'
# 请求头信息 需在网页登陆后获取 修改Cookie即可
headers_1 = {

    'Cookie': 'JSESSIONID=01B3402C9605B60E809386F66CDC046B-n3; sod=XnkPPCM729n8tyxhjeremyQroXNlUq+AynWFYOSir25lGD6YmA3Fw12W5AhGaAMcYyQiN2K2lAdwciPoqCw='
}
excel_name = ''
colNum = 0
rowNum = 0


def getAllData(infoUrl,number):
    # 有密码的请求一定要用post()
    # print(infoUrl)
    response = session.get(infoUrl, headers=headers_1)
    html_doc = response.text
    soup = BeautifulSoup(html_doc, "lxml")
    # print(soup.prettify())
    # form = soup.find('form', id="testform")
    sd = soup.find('input', id="sd")["value"]
    type1 = soup.find('div', id="type1").text.strip()
    type2 = soup.find('div', id="type2").text.strip()
    type3 = soup.find('div', id="type3").text.strip()
    type4 = soup.find('div', id="type4").text.strip()
    divs = soup.find_all('div', class_="ibox float-e-margins")
    # print(len(divs))
    # print(type4)
    sampleName = ""
    for div in divs:
        if div.text.find('抽检样品信息') >= 0:
            div2 = div.find_all('div', class_="row form-group")
            div3 = div2[2].find_all('div', class_="col-sm-4")
            div4 = div3[0].find_all('div', class_="col-sm-8")
            sampleName =  div4[0].text.strip()



    data={'sd':sd};
    if infoUrl.find("test_platform/agricultureTest/agricultureDetail") > -1:
        productInfoUrl = infoUrl2_1
    else:
        productInfoUrl = infoUrl1_1

    rowlist=[]
    rowlist.append(getProductInfo(productInfoUrl, data)["rows"])
    rowlist.append(number)
    rowlist.append(type1)
    rowlist.append(type2)
    rowlist.append(type3)
    rowlist.append(type4)
    rowlist.append(sampleName)
    # print(getProductInfo(productInfoUrl, data)["rows"])
    return  rowlist


#
def getAllPageData(pageUrl):
    response = session.get(pageUrl, headers=headers_1)
    infoJson = json.loads(response.text)  # 先把字典转成json
    infoList = []
    num = 0

    for info in infoJson['rows']:
        # oneList = [info['id'], info['sp_s_16']]
        # print(info)
        # print(oneList)
        infoList.append(getAllData(infoUrl % info['id'],info['sp_s_16']))

    return infoList


def getProductInfo(url,data):
    response = session.post(url, data,headers=headers_1)
    infoJson = json.loads(response.text)  # 先把字典转成json
    # print(infoJson)
    return infoJson

def writeContent2Excel(sheet, infoAllData):
    colNum = 1
    for infoData in infoAllData:
        infoRow = 0
        for info1 in infoData[0]:
            sheet.write(colNum, 0, infoData[1])
            sheet.write(colNum, 1, infoData[2])
            sheet.write(colNum, 2, infoData[3])
            sheet.write(colNum, 3, infoData[4])
            sheet.write(colNum, 4, infoData[5])
            sheet.write(colNum, 5, infoData[6])
            sheet.write(colNum, 6, info1["spdata_0"])
            sheet.write(colNum, 7, info1["spdata_1"])
            sheet.write(colNum, 8, info1["spdata_18"])
            sheet.write(colNum, 9, info1["spdata_2"])
            sheet.write(colNum, 10, info1["spdata_3"])
            sheet.write(colNum, 11, info1["spdata_4"])
            sheet.write(colNum, 12, info1["spdata_11"])
            sheet.write(colNum, 13, info1["spdata_15"])
            sheet.write(colNum, 14, info1["spdata_16"])
            sheet.write(colNum, 15, info1["spdata_7"])
            sheet.write(colNum, 16, info1["spdata_8"])
            sheet.write(colNum, 17, info1["spdata_20"])
            sheet.write(colNum, 18, info1["spdata_17"])
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
    sheet.write(0, 1, '食品大类')
    sheet.write(0, 2, '食品亚类')
    sheet.write(0, 3, '食品次亚类')
    sheet.write(0, 4, '食品细类')
    sheet.write(0, 5, '样品名称')
    sheet.write(0, 6, '检验项目')
    sheet.write(0, 7, '检验结果')
    sheet.write(0, 8, '结果单位')
    sheet.write(0, 9, '结果判定')
    sheet.write(0, 10, '检验依据')
    sheet.write(0, 11, '判定依据')
    sheet.write(0, 12, '最小允许限')
    sheet.write(0, 13, '最大允许限')
    sheet.write(0, 14, '允许限单位')
    sheet.write(0, 15, '方法检出限')
    sheet.write(0, 16, '检出限单位')
    sheet.write(0, 17, '备注')
    sheet.write(0, 18, '说明')
    infoList = getAllPageData(pageUrl)
    print(len(infoList))
    writeContent2Excel(sheet,infoList)
    book.close()

    # print(getAllData(infoUrl))
    print("======执行完毕，请查看文件======")
