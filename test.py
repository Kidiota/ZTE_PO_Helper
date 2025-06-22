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
def get_raw_info():
    i = 0
    #创建总输出数组

    rawInfo = []
    outputInfo = [len(filesName)]
    while i < len(filesName):
        fixedFileFullName = "fixed\\" + "fixed_4902261116 NGN DWDM EQ.pdf"#filesName[i]
        print(fixedFileFullName,len(filesName), i, "\n\n")
        with pdfplumber.open(fixedFileFullName) as pdf:
            oneRawInfo = []
            table_settings = {
                "vertical_strategy": "lines", 
                "horizontal_strategy": "text",
                }
            for page in pdf.pages:
                print(page)
                tables = page.extract_tables(table_settings)
                for table in tables:
                    for row in table:
                        if row != ['', '', '', '', '', '', '', '', '', '', '', '']:
                            if row != ['', '', '', '', '', '', '', '', '', '', '']:
                                if row != ['', '', '', '', '', '', '', '', '', '']:
                                    if row != ['', '', '', '', '', '', '', '', '']:
                                        if row != ['', '', '', '', '', '', '', '']:
                                            if row != ['', '', '', '', '', '', '']:
                                                if row != ['', '', '', '', '', '']:
                                                    if row != ['', '', '', '', '']:
                                                        if row != ['', '', '', '']:
                                                            if row != ['', '', '']:
                                                                if row != ['', '']:
                                                                    if row != ['']:
                                                                        #print(row)
                                                                        infoInLine = [row]
                                                                        oneRawInfo += infoInLine
        rawInfo += [oneRawInfo]
        i = i + len(filesName)
    return(rawInfo)

                
        

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


filesName = os.listdir("fixed")

#读取表格

raw_data = get_raw_info()

print(len(raw_data[0]))

i = 0
while i < len(raw_data[0]):
    print(i, " ", raw_data[0][i])
    i = i + 1



shutil.rmtree('fixed')