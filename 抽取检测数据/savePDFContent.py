import fitz
import sys, time
import os
from PyPDF2 import PdfFileWriter, PdfFileReader
import pdfplumber
import re
import xlsxwriter
import pandas as pd
from PIL import Image

"""
获取指定路径下的文件名
"""


def all_files_path(rootDir):
    for root, dirs, files in os.walk(rootDir):  # 分别代表根目录、文件夹、文件
        for file in files:  # 遍历文件
            file_path = os.path.join(root, file)  # 获取文件绝对路径
            filepaths.append(file_path)  # 将文件路径添加进列表
        for dir in dirs:  # 遍历目录下的子目录
            dir_path = os.path.join(root, dir)  # 获取子目录路径
            all_files_path(dir_path)  # 递归调用


"""
获取pdf指定页
"""


def getPage(dict, pageNum):
    output = PdfFileWriter()
    pdf_file = PdfFileReader(open(dict, "rb"))
    # 保存input.pdf中的1-5页到output.pdf
    return pdf_file.getPage(pageNum)


def getPageContent(filepath):
    with pdfplumber.open(filepath) as pdf:
        str = pdf.pages[2].extract_text()
        # print(str)
        startNum = str.find("抽样单编号")
        # endNum = str.find("检查封样人员")
        # print(startNum)
        result = str[startNum:startNum + 26].replace("抽样单编号", '').strip()
    return result


def get_filePath_fileName_fileExt(filepath):
    """
    获取文件路径， 文件名， 后缀名
    :param fileUrl:
    :return:
    """
    filepath, tmpfilename = os.path.split(filepath)
    shotname, extension = os.path.splitext(tmpfilename)
    startNum = shotname.find('_')
    fileNo = shotname[0:startNum]
    return fileNo


"""
抓取图片
"""


def pdf2pic(filepath, pic_path):
    checkXO = r"/Type(?= */XObject)"  # 使用正则表达式来查找图片
    checkIM = r"/Subtype(?= */Image)"
    doc = fitz.open(filepath)  # 打开pdf文件
    imgcount = 0  # 图片计数
    lenXREF = doc._getXrefLength()  # 获取对象数量长度
    imageList = []
    imageList2 = []
    # 遍历每一个对象
    for i in range(1, lenXREF):
        text = doc._getXrefString(i)  # 定义对象字符串
        isXObject = re.search(checkXO, text)  # 使用正则表达式查看是否是对象
        isImage = re.search(checkIM, text)  # 使用正则表达式查看是否是图片
        if not isXObject or not isImage:  # 如果不是对象也不是图片，则continue
            continue
        imgcount += 1
        # if imgcount != 12:
        #     continue
        pix = fitz.Pixmap(doc, i)  # 生成图像对象
        new_name = "图片{}.png".format(time.time())  # 生成图片的名称
        imageList.append(new_name)
        if pix.n < 5:  # 如果pix.n<5,可以直接存为PNG
            pix.writePNG(os.path.join(pic_path, new_name))
        else:  # 否则先转换CMYK
            pix0 = fitz.Pixmap(fitz.csRGB, pix)
            pix0.writePNG(os.path.join(pic_path, new_name))
            pix0 = None
        pix = None  # 释放资源
        time.sleep(0.1)
    time.sleep(1)
    imagePath = imageList[-2]
    imageList2.append(pic_path + '\\' + imagePath)
    imageList.pop(len(imageList) - 2)
    imagePath2 = imageList[0]
    imageList.pop(0)
    imageList2.append(pic_path + '\\' + imagePath2)
    # print(imageList)
    for image in imageList:
        os.remove(pic_path + '\\' + image)
    return imageList2


"""
写内容到excel
"""


def moveImage2Exl(imagePathStrList, pdfInfoStr, num, sheet):
    sheet.insert_image('C' + str(num + 1), imagePathStrList[0])
    sheet.write('B' + str(num + 1), pdfInfoStr)
    sheet.write('A' + str(num + 1), str(num))

    img = Image.open(imagePathStrList[1])
    h = img.height
    hasImage = "否"
    if h == 350:
        hasImage = "是"

    sheet.write('D' + str(num + 1), hasImage)


def del_file(path):
    ls = os.listdir(path)
    for i in ls:
        c_path = os.path.join(path, i)
        if os.path.isdir(c_path):
            del_file(c_path)
        else:
            os.remove(c_path)

if __name__ == "__main__":
    filepaths = []  # 初始化列表用来
    pageNum = 2
    # 放提取图片的文件夹
    pic_path = r"G:\pythonTmp1"
    if not os.path.exists(pic_path):
        os.makedirs(pic_path)

    # 放pdf的文件夹
    office_file_path = r'G:\office1'
    if not os.path.exists(office_file_path):
        print("请创建放置pdf的文件夹")
    all_files_path(office_file_path)
    infoDict = {}
    num = 1
    # 本程序同级目录下新建工作簿1.xlsx
    book = xlsxwriter.Workbook('工作簿1.xlsx')
    sheet = book.add_worksheet('sheet1')
    sheet.write('A' + str(num), '序号')
    sheet.write('B' + str(num), '抽样单编号')
    sheet.write('C' + str(num), '出具报告日期')
    sheet.write('D' + str(num), '是否有二维码')
    print(len(filepaths))
    for filepath in filepaths:
        # print(filepath)
        imagePathStrList = pdf2pic(filepath, pic_path)
        pdfInfoStr = get_filePath_fileName_fileExt(filepath)
        moveImage2Exl(imagePathStrList, pdfInfoStr, num, sheet)
        sys.stdout.write('\r======程序运行======[%s' % (round(num / len(filepaths) * 100, 2)) + '%]\n')
        num += 1
        # if num >10:
        #     break
    book.close()
    del_file(office_file_path) #清空文件夹
    print("======执行完毕，请查看文件======\n")
