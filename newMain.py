import pdfplumber
import PyPDF2
import os
import shutil

#修复CropBox问题
def fix_cropbox(pdf_path, output_path):
    #修复PDF文件的CropBox问题，并生成新的PDF文件.
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        writer = PyPDF2.PdfWriter()

        for page in reader.pages:
            if "/CropBox" not in page:
                page.cropbox = page.mediabox
            writer.add_page(page)

    with open(output_path, "wb") as output_file:
        writer.write(output_file)

#获取原始数据
def get_raw_info():
    i = 0
    rawInfo = []
    outputInfo = [len(filesName)]
    while i < len(filesName):
        fixedFileFullName = "fixed\\" + filesName[i]
        print(fixedFileFullName,i,"/",len(filesName),"\n\n")
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
        i = i + 1
    return(rawInfo)

def Preprocessing_Data(rawData):
    infoInLine = []
    Processed_Data = []
    i = 0
    #下面这个循环一次读取一个PDF文件的信息
    while i < len(rawData):
        one_PDF_data = rawData[i]
        #处理简单的信息
        PO_No = one_PDF_data[0][0][12:]
        PO_Date = one_PDF_data[1][0][10:]
        Contract_No = one_PDF_data[2][0][15:]
        Payment_Terms = one_PDF_data[3][0][16:]
        Project_Cost_Center = one_PDF_data[5][0][22:] + " " + one_PDF_data[6][0]
        Tracking_No = one_PDF_data[7][0][14:]

        #关键信息提取
        Form_Data = [] # 0: Service No  1: Description  2: Delivery Date  3: Order Qty  4: UoM  5: Unit Price  6: Total Price



        infoInLine = [PO_No, PO_Date, Contract_No, Payment_Terms, Project_Cost_Center, Tracking_No] 
        Processed_Data += [infoInLine]
        i = i + 1

    i = 0
    while i < len(Processed_Data):
        print(Processed_Data[i])
        i = i + 1

        
                
        

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


#上面那一坨就是从文件读取数据到内存，能运行，不需要知道怎么运行的，别碰就行了







raw_data = get_raw_info()

Preprocessing_Data(raw_data)







shutil.rmtree('fixed')