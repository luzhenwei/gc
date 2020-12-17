# 首先得登录上，记住请求的时候用post，然后拿到cookies
import time
import json
from bs4 import BeautifulSoup
from urllib import request
import os
import urllib.request
from tqdm import tqdm
# 国抽
from requests_html import HTMLSession

session = HTMLSession()

# 获取获取检验数据列表  需在网页登陆后获取
pageUrl = r'http://spcjinsp.gsxt.gov.cn/test_platform/api/food/getFood?order=desc&offset=0&limit=100&dataType=1&startDate=2020-05-30&endDate=2020-08-30&taskFrom=&samplingUnit=&testUnit=&enterprise=&sampledUnit=&foodName=&province=&reportNo=&bsfla=&bsflb=&sampleNo=&foodType1=&foodType4=&sampleNo_index=0&_=1598763659956'
# 获取普通检验数据详情
infoUrl = ''
infoUrl1 = r'http://spcjinsp.gsxt.gov.cn/test_platform/foodTest/foodDetail/%s'
# infoUrl1 = r'http://spcjinsp.gsxt.gov.cn/test_platform/foodTest/foodDetail/5643269'
# 获取农产品检验数据详情
infoUrl2 = r'http://spcjinsp.gsxt.gov.cn/test_platform/agricultureTest/agricultureDetail/%s'
# 请求头信息 需在网页登陆后获取 修改Cookie即可
headers_1 = {

    'Cookie': 'JSESSIONID=028218106A08F8E54A6F1FB6D34344E8-n3; sod=Pf7YHr8dKq0uxyAoyxhjeremyEYhQIyxhjeremyhr1TmyZmA9A4tkQfkfdX8pFtyxhjeremyXP2arvHByxhjeremy1O2qAWMWezLAR32k+g='
}

new_dic = ""
file_path = ""

def getData(infoUrl, doc):
    num = 1
    # 有密码的请求一定要用post()
    response = session.get(infoUrl, headers=headers_1)
    html_doc = response.text
    soup = BeautifulSoup(html_doc, "lxml")
    # print(soup.prettify())
    imgDivs = soup.find('ul', id="dowebok").find_all("div")
    # print(len(imgDivs))

    for img in imgDivs:
        imgUrl = img.find("img").attrs["src"]
        # print(imgUrl)
        file_suffix = os.path.splitext(imgUrl)[1]
        filename = '{}{}'.format(file_path + "/" + doc + "——" + str(num), file_suffix)
        if not os.path.exists(filename):
            urllib.request.urlretrieve(imgUrl, filename=filename)
        num = num + 1


#
def getAllPageData(pageUrl):
    response = session.get(pageUrl, headers=headers_1)
    infoJson = json.loads(response.text)  # 先把字典转成json
    infoList = []
    num = 0
    with tqdm(total=100) as pbar:
        for info in infoJson['rows']:
            # oneList = [info['id'], info['sp_s_16']]

            # print(info['id']) #id
            # print(info['sp_s_16'])#单号
            # 筛选状态是 完全提交的
            if info['sp_i_state'] == 2:
                oneList = [info['sp_s_16']]
                oneList.append(getData(infoUrl % info['id'], info['sp_s_16']))
            time.sleep(0.1)
            pbar.update(100 / len(infoJson['rows']))
    return infoList


def getdate():
    return time.strftime("%Y-%m-%d", time.localtime())


def createDic(file_path):
    return file_path + getdate();


if __name__ == "__main__":
    print("=====运行开始====")
    new_dic = './images'

    file_path = createDic(new_dic);
    if not os.path.exists(file_path):
        os.makedirs(file_path)
    if pageUrl.find("api/agriculture/getAgriculture") > -1:
        infoUrl = infoUrl2
    else:
        infoUrl = infoUrl1
    infoList = getAllPageData(pageUrl)
    print("=====运行结束====")
