
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
from pdfminer.layout import LTTextContainer
from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice

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

        #互动易记录文件网络存储路径
        self.recordFileAddr =""

        self.fileExt = ""

        #互动易记录表本地存储路径
        self.localAddr = ""

        #调研机构名称
        self.company = []

class InvestmentInfoTable(object):
    def __init__(self):
        self._addr = r'http://irm.cninfo.com.cn/ircs/interaction/irmInformationList.do?pageNo=1&stkcode=&beginDate=%s&endDate=%s&keyStr=&irmType=251314'

    def GetInvestmentInfo(self,investmentInfoVec,beginDate,endDate):
        '''
        获取记录信息
        :param investmentInfoVec: 获取的信息结构体
        :param beginDate: 查询的开始时间，比如2015-12-01
        :param endDate: 查询的结束时间，比如2015-12-01
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
        '''
        :param file: 投资者互动易文档路径，支持pdf，doc， docx
        :return:返回调研机构名，是一个列表
        '''
        companyInfo = ''
        if file.endswith(r'.doc'):
            companyInfo = self._GetFromDoc(file)
        elif file.endswith(r'.docx'):
            companyInfo = self._GetFromDocx(file)
        elif file.endswith(r'.pdf'):
            companyInfo = self._GetFromPdf(file)

        return self._FormatCompanyInfo(companyInfo)


    def _FormatCompanyInfo(self,compInfo):
        compInfo = compInfo.replace(u'；',r'\r')
        return [x.strip() for x in compInfo.split(r'\r') if len(x.strip())!=0]


    def _GetFromPdf(self,pdf):
        '''
        参考文档http://www.unixuser.org/~euske/python/pdfminer/programming.html
        '''
        pass
        fp = open(pdf, 'rb')
        #用文件对象来创建一个pdf文档分析器
        parser = PDFParser(fp)
        # 创建一个  PDF 文档
        doc = PDFDocument(parser)
        # 连接分析器 与文档对象
        parser.set_document(doc)
        # 检测文档是否提供txt转换，不提供就忽略
        if not doc.is_extractable:
            raise PDFTextExtractionNotAllowed

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
                if(isinstance(x, LTTextContainer)):
                    print x.get_text()

        pass

    def _GetFromDoc(self,wordDoc):
        word = client.Dispatch('Word.Application')
        doc = word.Documents.Open(wordDoc)
        return doc.Tables[0].Rows[1].Cells[1].Range.Text

    def _GetFromDocx(self,wordDocx):
        document = Document(wordDocx)
        return document.tables[0].cell(1,1).text




if __name__ == "__main__":

    #curTime = '2015-12-21'
    curTime = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    web = InvestmentInfoTable()
    infoVec = []
    web.GetInvestmentInfo(infoVec,curTime,curTime)
    #print infoVec

    #web.GetRecordFile(infoVec,'D:\\test\\')

    #打印code到文本
    f = open('D:\\test\\1.txt','w')
    codevec = []
    for info in infoVec:
        codevec.append(info.code)
    f.write('\n'.join(codevec))
    f.close()


    read = InvestigateInfo()
    s = read.GetInvestigateCompanyAndPeople(u'D:\\test\\2.pdf')
    for li in s:
        print li






