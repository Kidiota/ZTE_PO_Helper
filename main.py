import pdfplumber
import PyPDF2
import os

#修复CropBox问题
def fix_cropbox(pdf_path, output_path):
    #"""修复PDF文件的CropBox问题，并生成新的PDF文件."""
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        writer = PyPDF2.PdfWriter()

        for page in reader.pages:
            if "/CropBox" not in page:
                page.cropbox = page.mediabox
            writer.add_page(page)

    with open(output_path, "wb") as output_file:
        writer.write(output_file)

#读取input文件夹内文件的文件名
folder_path = "input"
filesName = os.listdir(folder_path)             #以列表方式记录所有文件名

#循环修复所有文件
i = 0
while i < len(filesName):
    fileFullName = "input\\" + filesName[i]
    fixedFileName = "fixed\\fixed_" + filesName[i]
    fix_cropbox(fileFullName, fixedFileName)
    i = i + 1


filesName = os.listdir("fixed")

#循环输出PO号和金额
i = 0
while i < len(filesName):
    fileFullName = "fixed\\" + filesName[i]
    #提取pdf文字
    with pdfplumber.open(fileFullName) as pdf:
        firstPage = pdf.pages[0]                    #指定页码
        firstPageText = firstPage.extract_text()    #提取文本

        #分离PO_Number
        poNum = firstPageText[firstPageText.find("PO Number : ") + 12 : firstPageText.find('PO Number : ') + 22]       #默认PO Number 为十位数字
        #print(poNum)

        #分离 TWENTY-TWO AND SEVENTY-SIX CENT ONLY 
        endPage = pdf.pages[-1]                     #指定页码
        endPageText = endPage.extract_text()        #提取文本

        #计算总价开始位置
        totalAmountStart = endPageText.find("ENT ONLY ") + 9

        #计算总价结束位置
        totalAmountEnd = endPageText.find("Should the Sales Tax/Service")


        #分离总价
        totalAmount = endPageText[totalAmountStart : totalAmountEnd]
        
    #显示结果
    print(fileFullName, "    ", poNum, "    ", totalAmount)
    i = i + 1
