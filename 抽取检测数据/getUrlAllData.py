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
pageUrl = r'http://spcjinsp.gsxt.gov.cn/test_platform/api/food/getFood?order=desc&offset=0&limit=10000&dataType=8&startDate=2020-04-01&endDate=2020-12-15&taskFrom=&samplingUnit=&testUnit=&enterprise=&sampledUnit=&foodName=&province=&reportNo=&bsfla=&bsflb=&sampleNo=&foodType1=&foodType4=&sampleNo_index=0&_=1607999260523'
# 获取普通检验数据详情
infoUrl = ''
infoUrl1 = r'http://spcjinsp.gsxt.gov.cn/test_platform/foodTest/foodDetail/%s'
# 获取农产品检验数据详情
infoUrl2 = r'http://spcjinsp.gsxt.gov.cn/test_platform/agricultureTest/agricultureDetail/%s'
# 请求头信息 需在网页登陆后获取 修改Cookie即可
headers_1 = {

    'Cookie': 'JSESSIONID=4E9AB10ADDDD69810DCEC4BDE2711DCC-n3; sod=bX5LETFu3xUGsPj4aGSM6C6DEjml87uImYSyIrpJOkefMGhHQIaRlA1jhL66X4R3vOcnAsAIfAE='
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
    print(infoUrl)
    response = session.get(infoUrl, headers=headers_1)
    html_doc = response.text
    soup = BeautifulSoup(html_doc, "lxml")
    # print(soup.prettify())
    divs = soup.find_all('div', class_="ibox float-e-margins")
    # print(len(divs))
    infoAll = []
    infoAllData = {}
    for div1 in divs:
        divName = div1.find('h2').text
        # print(divName)
        div2Dict = {}
        divFormDict = {}
        if divName == '照片信息':
            continue
        elif divName == '检验信息':
            testform = div1.find('form', id="testform")
            if testform  is None:
                continue
            divForm = testform.find_all('div', class_="row")

            for divForm1 in divForm:
                divForm2 = divForm1.find_all('label', class_="control-label col-sm-5")
                divForm3 = divForm1.find_all('div', class_="col-sm-7")
                divForm2List = []
                divForm3List = []
                for divForm20 in divForm2:
                    divForm2List.append(divForm20.text.strip())
                for divForm30 in divForm3:
                    divForm3List.append(divForm30.text.strip())

                divFormLen = len(divForm2List)
                if divFormLen > 0:
                    i = 0
                    while i < divFormLen:
                        a = {divForm2List[i]: divForm3List[i]}
                        divFormDict.update(a)
                        i += 1
            # print(divFormDict)
            aDict = {divName: divFormDict}
            infoAllData.update(aDict)
        else:

            div2 = div1.find_all('div', class_="row form-group")
            for div3 in div2:
                div40List = []
                div50List = []
                div4 = div3.find_all('label', class_="control-label col-sm-4")
                div5 = div3.find_all('div', class_="col-sm-8")
                for div40 in div4:
                    div40List.append(div40.text.strip())
                for div50 in div5:
                    div50List.append(div50.text.strip())
                if len(div40List) > 0:
                    i = 0
                    div4len = len(div40List)
                    while i < div4len:
                        b = {div40List[i]: div50List[i]}
                        div2Dict.update(b)
                        i += 1
            # print(div2Dict)
            bDict = {divName: div2Dict}
            infoAllData.update(bDict)
    # print(infoAllData)
    dictValue = infoAllData['生产企业信息']
    del (infoAllData['生产企业信息'])
    newDict = {'生产企业信息': dictValue}
    infoAllData.update(newDict)
    return infoAllData

def getAllDataDtb(infoUrl):
    # 有密码的请求一定要用post()
    response = session.get(infoUrl, headers=headers_1)
    html_doc = response.text
    soup = BeautifulSoup(html_doc, "lxml")
    # print(soup.prettify())
    # divOne = soup.find(id="collapseOne")
    # print(divOne)
    # print(len(divOne))
    divs = soup.find_all('div', class_="ibox float-e-margins")

    infoAllData = {}
    for div1 in divs:
        divName = div1.find('h2').text
        # print(divName)
        div2Dict = {}
        divFormDict = {}
        if divName == '照片信息':
            continue
        elif divName == '检验信息':
            # testform = div1.find('form', id="testform")
            divForm = div1.find_all('div', class_="row")
            for divForm1 in divForm:
                divForm2 = divForm1.find_all('label', class_="control-label col-sm-5")
                divForm3 = divForm1.find_all('div', class_="col-sm-7")
                divForm2List = []
                divForm3List = []
                for divForm20 in divForm2:
                    divForm2List.append(divForm20.text.strip())
                for divForm30 in divForm3:
                    divForm3List.append(divForm30.text.strip())

                divFormLen = len(divForm2List)
                if divFormLen > 0:
                    i = 0
                    while i < divFormLen:
                        a = {divForm2List[i]: divForm3List[i]}
                        divFormDict.update(a)
                        i += 1
            # print(divFormDict)
            aDict = {divName: divFormDict}
            infoAllData.update(aDict)
        else:

            div2 = div1.find_all('div', class_="row form-group")
            for div3 in div2:
                div40List = []
                div50List = []
                div4 = div3.find_all('label', class_="control-label col-sm-4")
                div5 = div3.find_all('div', class_="col-sm-8")
                for div40 in div4:
                    div40List.append(div40.text.strip())
                for div50 in div5:
                    div50List.append(div50.text.strip())
                if len(div40List) > 0:
                    i = 0
                    div4len = len(div40List)
                    while i < div4len:
                        b = {div40List[i]: div50List[i]}
                        div2Dict.update(b)
                        i += 1
            # print(div2Dict)
            bDict = {divName: div2Dict}
            infoAllData.update(bDict)
    # print(infoAllData)
    dictValue = infoAllData['生产企业信息']
    del (infoAllData['生产企业信息'])
    newDict = {'生产企业信息': dictValue}
    infoAllData.update(newDict)
    return infoAllData


#
def getAllPageData(pageUrl):
    response = session.get(pageUrl, headers=headers_1)
    infoJson = json.loads(response.text)  # 先把字典转成json
    infoList = []
    num = 0

    for info in infoJson['rows']:
        # oneList = [info['id'], info['sp_s_16']]

        # print(info['id'])
        # print(info['sp_s_16'])
        # 筛选状态是 完全提交的
        # sp_i_state == 9
        # 完全提交
        # sp_i_state == 2 | | sp_i_state == 0) {
        # if (row.sp_i_jgback == 1) {
        # 退修待填报
        # }
        # 待填报
        # row.sp_i_state == 12
        # 待签章
        # sp_i_state == 4
        # 待审核
        # sp_i_state == 5
        # 待批准
        # sp_i_state == 7
        # 待发送
        # sp_i_state == 1
        # 已退修
        if info['sp_i_state'] == 9 or  info['sp_i_state'] == 5:
            oneList = [info['sp_s_16']]
            oneList.append(getAllData(infoUrl % info['id']))
            num += 1
            # print(oneList)
            # updateExcel(sheet, num, oneList)
            # num += 1
            infoList.append(oneList)
        elif  info['sp_i_state'] == 2:
            oneList = [info['sp_s_16']]
            oneList.append(getAllDataDtb(infoUrl % info['id']))
            num += 1
            # print(oneList)
            # updateExcel(sheet, num, oneList)
            # num += 1
            infoList.append(oneList)
    print("未统计数据条数：" + str(len(infoJson['rows']) - num))
    return infoList


def updateExcel(sheet, num, oneList):
    if num == 1:
        writeTital2Excel(sheet, oneList[1])
    writeContent2Excel(sheet, oneList[1], num)


def updateAllExcel(sheet, num, newDict):
    if num == 1:
        writeTital2Excel(sheet, newDict)
    writeContent2Excel(sheet, newDict, num)


def writeTital2Excel(sheet, infoAllData):
    colNum = 0
    rowNum = 0
    for infoDataKey in infoAllData:
        # print(infoDataKey)
        for infoKey in infoAllData[infoDataKey]:
            sheet.write(rowNum, colNum, infoDataKey)
            colNum += 1
    rowNum += 1
    colNum = 0
    for infoDataKey in infoAllData:
        for infoKey in infoAllData[infoDataKey]:
            sheet.write(rowNum, colNum, infoKey.replace('：', ''))
            colNum += 1
    rowNum += 1


def writeContent2Excel(sheet, infoAllData, rowNum, newDict):
    colNum = 0
    for infoDataKey in infoAllData:
        for infoKey in newDict[infoDataKey]:
            sheet.write(rowNum + 1, colNum, infoAllData[infoDataKey].get(infoKey, ''))
            colNum += 1


def getCookie():
    # 声明一个CookieJar对象实例来保存cookie
    cookie = cookiejar.CookieJar()
    # 利用urllib.request库的HTTPCookieProcessor对象来创建cookie处理器,也就CookieHandler
    handler = request.HTTPCookieProcessor(cookie)
    # 通过CookieHandler创建opener
    opener = request.build_opener(handler)
    # 此处的open方法打开网页
    response = opener.open('http://spcj.gsxt.gov.cn/login')
    print(cookie)
    # 打印cookie信息
    for item in cookie:
        print('Name = %s' % item.name)
        print('Value = %s' % item.value)


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
        excel_name = getdate() + '市县级农产品检测信息统计'
    else:
        infoUrl = infoUrl1
        excel_name = getdate() + '普通食品检测信息统计'
    book = xlsxwriter.Workbook(excel_name+'.xlsx')
    sheet = book.add_worksheet('sheet1')

    infoList = getAllPageData(pageUrl)
    dictLen2 = 0
    newDict = {}
    type_list = list(infoList[0][1].keys())

    for infoKey in type_list:
        dictLen1 = 0
        max = 0
        retInfoKey = []
        newDict.update({infoKey: getMaxKey(infoKey, infoList, max)})
        # print(newDict)
    num = 1
    # print(newDict)
    # print(info[1])
    for info in infoList:
        if num == 1:
            writeTital2Excel(sheet, newDict)
        writeContent2Excel(sheet, info[1], num, newDict)
        num += 1
    book.close()

    # print(getAllData(infoUrl))
    print("======执行完毕，请查看文件======")
