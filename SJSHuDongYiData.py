
# -*- coding: cp936 -*-
import mechanize
from bs4 import BeautifulSoup
import time
import re
import urllib2
import os

#doc
from docx import Document
from docx.shared import Inches
from win32com import client

#pdf
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import *

#br = mechanize.Browser()
#response = br.open('http://irm.cninfo.com.cn/ircs/interaction/irmInformationList.do?pageNo=1&stkcode=&beginDate=2015-10-30&endDate=2015-11-30&keyStr=&irmType=251314')
#print response.read()



#soup = BeautifulSoup(open(r'C:\Users\zcj\Desktop\text2.html'))


#from docx import Document
#from docx.shared import Inches
#from PyPDF2 import PdfFileWriter, PdfFileReader
#import pdfminer
#document = Document(r'c:\1.docx')
#tables = document.tables

#for table in tables:
#    for row in table.rows:
#        for cell in row.cells:
#            print(cell.text)
 #           print('++++++++++++++++')


#input1 = PdfFileReader(open(r"c:\1.pdf", "rb"))
#print input1.getDocumentInfo()
#f = open("c:\\2.txt","w")
#input1.read(f)
#print "document1.pdf has %d pages." % input1.getNumPages()
#page = input1.getPage(1)
#print page['/Contents']
#print page.getContents()['/Filter']

def CreateDirs(dirs):
    if not os.path.exists(dirs):
        print u'创建目录:'+ dirs
        os.makedirs(dirs)

class InvestmentInfo(object):
    def __init__(self):
        self.code = ""
        self.name = ""
        self.uploadDate = ""
        self.recordData = ""
        self.recordFileAddr =""
        self.fileExt = ""
        self.localAddr = ""
        self.company = []

class InvestmentInfoTable(object):
    def __init__(self):
        self._addr = r'http://irm.cninfo.com.cn/ircs/interaction/irmInformationList.do?pageNo=1&stkcode=&beginDate=%s&endDate=%s&keyStr=&irmType=251314'

    def GetInvestmentInfo(self,investmentInfoVec,beginDate,endDate):
        '''
        获取记录信息
        :param investmentInfoVec: 获取的信息结构体
        :param beginDate: 查询的开始时间
        :param endDate: 查询的结束时间
        :return:
        '''
        soup = BeautifulSoup(self._GetWebPageInfo(beginDate,endDate))
        for tr in soup.table.find_all('tr'):
            tagAs = tr.find_all('a')
            if len(tagAs) == 2:
                print tagAs[0]
                print tagAs[1].decode()
                info = InvestmentInfo()
                self._GetCode(tagAs[0],info)
                self._GetOtherInfo(tagAs[1],info)
                investmentInfoVec.append(info)

            else:
                print 'TagA num error!'

    def GetRecordFile(self,investmentInfoVec,downloadDir):
        '''
        根据记录信息下载文件并且填充investmentInfo结构体的本地文件信息
        :param investmentInfoVec: investmentInfo信息结构体
        :param downloadDir: 需要把文件下载的本地目录
        :return:
        '''
        for info in investmentInfoVec:
            filedir = downloadDir + info.uploadDate
            CreateDirs(filedir)
            filename = re.sub(r'[/\\*?<>:|]',"",info.name)
            filename = filedir+'\\'+info.code+filename+info.recordData+'.'+info.fileExt
            f = urllib2.urlopen(info.recordFileAddr)
            with open(filename, "wb") as code:
                code.write(f.read())


    def _GetWebPageInfo(self,beginDate,endDate):
        pageAddr = self._addr % (beginDate,endDate)
        br = mechanize.Browser()
        response = br.open(pageAddr)
        return response.read()

    def _GetCode(self,tagA,info):
        info.code = unicode(tagA.string).strip()

    def _GetOtherInfo(self,tagA,info):
        info.recordFileAddr = tagA['href']
        #解析记录日期和文件后缀名
        m = re.match('.+/([0-9]{4}-[0-9]{1,2}-[0-9]{1,2})/.+\.(.+)\?.+',info.recordFileAddr)
        if m is None:
            print '获取上传日期及扩展名错误'
        info.uploadDate =m.group(1)
        info.fileExt = m.group(2).lower()
        title = tagA['title']
        m = re.match(u'(.+) *： *([0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日).*',tagA['title'])
        if m is None:
            print '获取记录日期及名称错误'
        info.name = m.group(1)
        info.recordData = m.group(2)


class InvestigateInfo(object):
    def GetInvestigateCompanyAndPeople(self,file):
        if file.endswith(r'.doc'):
            return self.GetFromDoc(file)
        elif file.endswith(r'.docx'):
            return self.GetFromDocx(file)
        elif file.endswith(r'.pdf'):
            return self.GetFromPdf(file)

    def GetFromPdf(self,pdf):
        pass

    def GetFromDoc(self,wordDoc):
        word = client.Dispatch('Word.Application')
        doc = word.Documents.Open(wordDoc)
        return doc.Tables[0].Rows[1].Cells[1].Range.Text

    def GetFromDocx(self,wordDocx):
        document = Document(wordDocx)
        return document.tables[0].cell(1,1).text




if __name__ == "__main__":

    #read = InvestigateInfo()
    #s = read.GetInvestigateCompanyAndPeople(u'D:\\test\\2.docx')
    #print s
    #curTime = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    #web = InvestmentInfoTable()
    #infoVec = []
    #web.GetInvestmentInfo(infoVec,curTime,curTime)
    #print infoVec

    #web.GetRecordFile(infoVec,'D:\\test\\')


    #pdf
    from pdfminer.layout import LAParams
    from pdfminer.converter import PDFPageAggregator
    from pdfminer.pdfparser import PDFParser
    from pdfminer.pdfdocument import PDFDocument
    from pdfminer.pdfpage import PDFPage
    from pdfminer.pdfpage import PDFTextExtractionNotAllowed
    from pdfminer.pdfinterp import PDFResourceManager
    from pdfminer.pdfinterp import PDFPageInterpreter
    from pdfminer.pdfdevice import PDFDevice
    fp = open(r'C:\Users\zcj\Desktop\py\1.PDF', 'rb')
    #用文件对象来创建一个pdf文档分析器
    parser = PDFParser(fp)
    # 创建一个  PDF 文档
    doc = PDFDocument(parser)
    # 连接分析器 与文档对象
    parser.set_document(doc)
    # 检测文档是否提供txt转换，不提供就忽略
    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed

    # Get the outlines of the document.
    #outlines = doc.get_outlines()
    #for (level,title,dest,a,se) in outlines:
    #    print (level, title)
    # 创建PDf 资源管理器 来管理共享资源
    rsrcmgr = PDFResourceManager()
    # 创建一个PDF设备对象
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    for page in PDFPage.create_pages(doc):
        interpreter.process_page(page)
        # receive the LTPage object for the page.
        layout = device.get_result()
        for x in layout:
            if(not isinstance(x, LTTextBox)):
                print x.get_text()

    # 处理文档对象中每一页的内容
    # doc.get_pages() 获取page列表
    # 循环遍历列表，每次处理一个page的内容
    # 这里layout是一个LTPage对象 里面存放着 这个page解析出的各种对象 一般包括LTTextBox, LTFigure, LTImage, LTTextBoxHorizontal 等等 想要获取文本就获得对象的text属性，
    for i, page in enumerate(doc.get_pages()):
        interpreter.process_page(page)
        layout = device.get_result()
        for x in layout:
            if(isinstance(x, LTTextBoxHorizontal)):
                if(len(x.text) > 100):
                    string = x.text.replace('/n', ' ')
                    print string
        print '/n/n/n/n'
    #http://www.unixuser.org/~euske/python/pdfminer/programming.html