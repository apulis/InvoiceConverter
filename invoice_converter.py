#!/usr/bin/env Python
# coding=utf-8

import sys, datetime, fitz, os, codecs
from pptx import Presentation
from pptx.util import Cm
import comtypes.client
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger

invoicePath = './inputs'
tempImagePath = './temp/images'
tempPptxPath = './temp/pptx'
templatePptxPath = './凭证粘贴模板.pptx'
outPath = './outputs'
gdict = {0:u'零',1:u'壹',2:u'贰',3:u'叁',4:u'肆',5:u'伍',6:u'陆',7:u'柒',8:u'捌',9:u'玖',10:u'拾'}
#总页数
totalPage = None
#总金额
totalAmount = None
#经办人
name = None
#是否自动计算总页数、总金额
skip = False

def pyMuPDF_fitz(pdfPath, imagePath):
    # startTime_pdf2img = datetime.datetime.now()#开始时间

    # print("imagePath="+imagePath)
    baseName = os.path.basename(pdfPath)
    pdfName = os.path.splitext(baseName)[0]
    pdfDoc = fitz.open(pdfPath)

    page = pdfDoc[0]
    rotate = int(0)
    # 每个尺寸的缩放系数为1.3，这将为我们生成分辨率提高2.6的图像。
    # 此处若是不做设置，默认图片大小为：792X612, dpi=96
    zoom_x = 2.5 #(1.33333333-->1056x816)   (2-->1584x1224)
    zoom_y = 2.5
    mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
    pix = page.getPixmap(matrix=mat, alpha=False)

    if not os.path.exists(imagePath):#判断存放图片的文件夹是否存在
        os.makedirs(imagePath) # 若图片文件夹不存在就创建

    pix.writePNG('%s/%s.png' % (imagePath, pdfName))#将图片写入指定的文件夹内

def batchPdf2Png(invoicePath, outPngPath):
    files = os.listdir(invoicePath)
    pdfFiles = [f for f in files if f.endswith((".pdf"))]
    
    global totalPage, totalAmount

    #计算总页数
    if not skip :
        totalPage = len(pdfFiles)
    
    #计算总金额
    if not skip :
        totalAmount = 0

    for pdfFile in pdfFiles:
        fullpath = os.path.join(invoicePath, pdfFile)
        pyMuPDF_fitz(fullpath, outPngPath)

        if not skip :
            totalAmount +=  float(pdfFile[:-4])

def insertPngInSlide(path_to_presentation, img_path):
    prs = Presentation(path_to_presentation)

    slide = prs.slides[0]
    left, top, width, height= Cm(4.28), Cm(2.79), Cm(25.58), Cm(16.59)
    pic = slide.shapes.add_picture(img_path, left, top, height=height, width=width)
    prs.save('test.pptx')

def numToCN(num):
    cn = ''

    if num < 10:
        cn = gdict[num]
    elif num > 10 and num < 100 :
        if num % 10 == 0:
            cn = '{0}拾'.format(gdict[num // 10])
        else:
            cn = '{0}拾{1}'.format(gdict[num // 10], gdict[num % 10])
    return cn

def fillTextInSlide(slide, curPage, totalPage, totalAmount, curAmount):
    global name

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.text == '凭证总张数：':
                    run.text = '凭证总张数：{0}张'.format(numToCN(totalPage))
                if run.text == '本页张数：':
                    run.text = '本页张数：' + '壹张'
                if run.text == '凭证总金额：':
                    run.text = '凭证总金额：¥{total:.2f}'.format(total=totalAmount)
                if run.text == '本页金额：':
                    run.text = '本页金额：¥{cur:.2f}'.format(cur=curAmount)
                if run.text == '经办人：':
                    run.text = '经办人：{name}'.format(name=name)
                if run.text == '第      页        共      页':
                    run.text = '第  {0}  页        共  {1}  页'.format(curPage, totalPage)

def batchInsertPngInSlide(path_to_tmpl_presentation, imgsPath):
    left, top, width, height= Cm(4.28), Cm(2.79), Cm(25.58), Cm(16.59)

    files = os.listdir(imgsPath)
    pngfiles = [f for f in files if f.endswith((".png"))]

    for index, pngfile in enumerate(pngfiles):
        fullpath = os.path.join(imgsPath, pngfile)

        prs = Presentation(path_to_tmpl_presentation)
        slide = prs.slides[0]        
        pic = slide.shapes.add_picture(fullpath, left, top, height=height, width=width)

        # 填入文字内容
        curAmount = float(pngfile[:-4])
        fillTextInSlide(slide, index+1, totalPage, totalAmount, curAmount)

        pptxPath = os.path.join(tempPptxPath, os.path.splitext(pngfile)[0]+'.pptx')
        prs.save(pptxPath)

def init_powerpoint():
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    return powerpoint

def ppt_to_pdf(powerpoint, inputFileName, outputFileName, formatType = 32):
    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName.replace(".pptx","").replace(".ppt","") + ".pdf"
    # print(inputFileName)
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
    deck.Close()
    # print('转换%s文件完成'%outputFileName)

def convert_files_in_folder(powerpoint, folder, outPath):
    files = os.listdir(folder)
    pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
    for pptfile in pptfiles:
        fullpath = os.path.join(folder, pptfile)
        pdfpath = os.path.join(outPath, os.path.splitext(pptfile)[0]+'.pdf')
        ppt_to_pdf(powerpoint, fullpath, pdfpath)

def del_file(path_data):
    for i in os.listdir(path_data) :# os.listdir(path_data)#返回一个列表，里面是当前目录下面的所有东西的相对路径
        file_data = path_data + "\\" + i#当前文件夹的下面的所有东西的绝对路径
        if os.path.isfile(file_data) == True:#os.path.isfile判断是否为文件,如果是文件,就删除.如果是文件夹.递归给del_file.
            os.remove(file_data)
        else:
            del_file(file_data)

def getfilenames(filepath='',filelist_out=[],file_ext='all'):
    # 遍历filepath下的所有文件，包括子目录下的文件
    for fpath, dirs, fs in os.walk(filepath):
        for f in fs:
            fi_d = os.path.join(fpath, f)
            if  file_ext == 'all':
                filelist_out.append(fi_d)
            elif os.path.splitext(fi_d)[1] == file_ext:
                filelist_out.append(fi_d)
            else:
                pass
    return filelist_out

def mergefiles(path, output_filename, import_bookmarks=False):
    # 遍历目录下的所有pdf将其合并输出到一个pdf文件中，输出的pdf文件默认带书签，书签名为之前的文件名
    # 默认情况下原始文件的书签不会导入，使用import_bookmarks=True可以将原文件所带的书签也导入到输出的pdf文件中
    merger = PdfFileMerger()
    filelist = getfilenames(filepath=path, file_ext='.pdf')
    if len(filelist) == 0:
        print("当前目录及子目录下不存在pdf文件")
        sys.exit()
    for filename in filelist:
        f = codecs.open(filename, 'rb')
        file_rd = PdfFileReader(f)
        short_filename = os.path.basename(os.path.splitext(filename)[0])
        if file_rd.isEncrypted == True:
            print('不支持的加密文件：%s'%(filename))
            continue
        merger.append(file_rd, bookmark=short_filename, import_bookmarks=import_bookmarks)
        print('合并文件：%s'%(filename))
        f.close()
    out_filename=os.path.join(os.path.abspath(path), output_filename)
    merger.write(out_filename)
    print('合并后的输出文件：%s'%(out_filename))
    merger.close()

def excetue():
    if len(sys.argv) != 2 and len(sys.argv) != 4:
        print('参数个数不对')
        return

    global name, totalPage, totalAmount, skip

    name = str(sys.argv[1])

    if len(sys.argv) == 4:
        totalPage = int(sys.argv[2])
        totalAmount = float(sys.argv[3])
        skip = True

    # pdf转图片
    batchPdf2Png(invoicePath, tempImagePath)
    # batchPdf2Png(invoicePath, tempImagePath, totalPage, totalAmount)
    # 图片插入pptx模板
    # batchInsertPngInSlide(templatePptxPath, tempImagePath, totalPage, totalAmount)
    batchInsertPngInSlide(templatePptxPath, tempImagePath)
    # pptx导出pdf
    powerpoint = init_powerpoint()
    absPptxPath = os.path.abspath(tempPptxPath)
    absOutPath = os.path.abspath(outPath)
    convert_files_in_folder(powerpoint, absPptxPath, absOutPath)
    powerpoint.Quit()
    # 清理临时文件
    del_file(tempImagePath)
    del_file(tempPptxPath)
    # 合并pdf文件
    mergefiles(outPath, 'All.pdf')

def printText():
    fillTextInSlide(templatePptxPath, 10, 21, 300,100)

if __name__ == '__main__':
    excetue()
    # printText()
