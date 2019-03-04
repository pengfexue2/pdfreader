from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal, LAParams
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed
from mailmerge import MailMerge
import os

# text_path = r'photo-words.pdf'
def parse(text_path):
    '''解析PDF文本，并保存到TXT文件中'''
    with open(text_path, 'rb') as fp:
        # 用文件对象创建一个PDF文档分析器
        parser = PDFParser(fp)

    # 创建一个PDF文档
    doc = PDFDocument()

    # 连接分析器，与文档对象
    parser.set_document(doc)
    doc.set_parser(parser)

    # 提供初始化密码，如果没有密码，就创建一个空的字符串
    doc.initialize()

    # 检测文档是否提供txt转换，不提供就忽略
    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:

        # 创建PDF，资源管理器，来共享资源
        rsrcmgr = PDFResourceManager()

        # 创建一个PDF设备对象
        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)

        # 创建一个PDF解释其对象
        interpreter = PDFPageInterpreter(rsrcmgr, device)

        #content 用于保存文本信息
        content = []
        #content_test 用于标记文本信息方便查找
        content_test=[]
        # 循环遍历列表，每次处理一个page内容
        # doc.get_pages() 获取page列表
        for page in doc.get_pages():
            index = 0

            #单页文本信息列表
            singlepage = []
            singletest = []

            interpreter.process_page(page)

            # 接受该页面的LTPage对象
            layout = device.get_result()
            # 这里layout是一个LTPage对象 里面存放着 这个page解析出的各种对象
            # 一般包括LTTextBox, LTFigure, LTImage, LTTextBoxHorizontal 等等
            # 想要获取文本就获得对象的text属性，
            for x in layout:
                if (isinstance(x, LTTextBoxHorizontal)):
                    results = x.get_text()
                    singlepage.append(results.replace("\n",'').replace(u'\u3000',u'').replace(" ",''))
                    singletest.append(f"{index}:{results}")
                    index += 1
            content.append(singlepage)
            content_test.append(singletest)
    return content,content_test

def work(filename):
    content,content_test = parse(filename)
    print("检测PDF读取第五页第1段和第4段：")
    print("第一段: ",content[4][1])
    print("第四段: ",content[4][4])

    template = "../读书笔记.docx"
    document = MailMerge(template)
    document.merge(
        firstTED = content[4][1],
        fourthTED = content[4][4],
    )
    document.write("读书笔记更新.docx")


if __name__ == '__main__':
    folder = "test"
    os.chdir(folder)
    pdflist = os.listdir()
    print(pdflist)
    faillist=[]
    for item in pdflist:
        if item[-4:] in [".pdf",".PDF"] :
            work(item)
            # try:
            #     print(f"开始解析《{item}》。。。")
            #     work(item)
            #     print(f"若无警告信息，成功搞定；若有警告，请检查。。。\n")
            # except:
            #     print(f"解析{item}失败失败失败失败失败\n")
            #     faillist.append(item)
    # print(f"以下pdf文件解析失败：{faillist}")

