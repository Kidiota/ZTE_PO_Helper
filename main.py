import pdfplumber
import PyPDF2
import os
import shutil

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

#循环输出PO号和金额
def get_defult_info():
    i = 0
    #创建总输出数组
    outputInfo = [len(filesName)]
    while i < len(filesName):
        fixedFileFullName = "fixed\\" + filesName[i]
        #提取pdf文字
        text = ""
        with pdfplumber.open(fixedFileFullName) as pdf:
            for page in pdf.pages:
                text = text + page.extract_text()

            #分离PO_Number
            poNum = text[text.find("PO Number : ") + 12 : text.find('PO Number : ') + 22]       #默认PO Number 为十位数字

            #计算总价开始位置
            totalAmountStart = text.find(" ONLY ") + 6

            #计算总价结束位置
            totalAmountEnd = text.find("\nShould the Sales Tax/Service")

            #判断SUBRACK是否存在
            subrack = text.find("SUBRACK")
            
            if subrack >= 0:
                subrack = "Has SUBRACK      "
                #分离SUBRACK数量
                sub_num_start = text.find("SUBRACK") + 19
                sub_num_end = text.find("SUBRACK") + 20
                sub_num = text[sub_num_start : sub_num_end]
                #分离SUBRACK型号
                sub_type_start = text.find("SUBRACK") - 5
                sub_type_end = text.find("SUBRACK") -1
                sub_type = text[sub_type_start : sub_type_end]
            else:
                subrack = "Don't has SUBRACK"
                sub_num = "-"
                sub_type = "-"
                
            
            
            #分离总价
            totalAmount = text[totalAmountStart : totalAmountEnd]
        outputInfo += [[fixedFileFullName[12 : -4], poNum, totalAmount, subrack, sub_num, sub_type]]
        print('\r' + '已提取' + str(i) + '/' + str(len(filesName)), end='', flush=True)
        i = i + 1
    print('\n提取完毕')
    return outputInfo

os.mkdir('fixed')


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
    print('\r' + '预处理文件' + str(i) + '/' + str(len(filesName)), end='', flush=True)
print('\n处理完毕，开始提取数据')


filesName = os.listdir("fixed")

outputInfo = get_defult_info()
i = 0
while i < outputInfo[0]:
    i = i + 1
    print("文件名：", outputInfo[i][0], "    PO号：", outputInfo[i][1], "    项目总价：", outputInfo[i][2], "    ", outputInfo[i][3], "    SUBRACK数量：", outputInfo[i][4], "    SUBRACK型号", outputInfo[i][5])
    
shutil.rmtree('fixed')