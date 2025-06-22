import pdfplumber
import PyPDF2
import os
import shutil
import time
import xlsxwriter

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
    print("开始提取数据")
    while i < len(filesName):
        fixedFileFullName = "fixed\\" + filesName[i]
        with pdfplumber.open(fixedFileFullName) as pdf:
            oneRawInfo = []
            table_settings = {
                "vertical_strategy": "lines", 
                "horizontal_strategy": "text",
                }
            for page in pdf.pages:
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
        print('\r' + '已提取' + str(i) + '/' + str(len(filesName)), end='', flush=True)
    print('\n提取完毕')
    return(rawInfo)

def Preprocessing_Data(rawData):
    print("开始处理数据")
    infoInLine = []
    Processed_Data = []
    i = 0
    #下面这个循环一次读取一个PDF文件的信息
    while i < len(rawData):
        one_PDF_data = rawData[i]

        #关键信息提取
        Form_Data = [] # 0: Service No  1: Description  2: Delivery Date  3: Order Qty  4: UoM  5: Unit Price  6: Total Price

        if one_PDF_data[0][0] == 'To:':
            POSN = one_PDF_data.index(['No','Service No', '', 'Date', 'Qty', '', '', ''])
            PO_No = one_PDF_data[0][1][12:]
            PO_Date = one_PDF_data[1][1][10:]
            Contract_No = one_PDF_data[2][1][15:]
            Payment_Terms = one_PDF_data[3][1][16:]
            Project_Cost_Center = one_PDF_data[5][1][22:] + " "
            if one_PDF_data[6][0] == 'MALAYSIA':
                Project_Cost_Center = one_PDF_data[5][1][22:] + " " + one_PDF_data[6][1]
            else:
                Project_Cost_Center = one_PDF_data[5][1][22:] + " " + one_PDF_data[6][0]
            Tracking_No = one_PDF_data[7][1][14:]
        else:
            POSN = one_PDF_data.index(['Service No', '', 'Date', 'Qty', '', '', ''])
            PO_No = one_PDF_data[0][0][12:]
            PO_Date = one_PDF_data[1][0][10:]
            Contract_No = one_PDF_data[2][0][15:]
            Payment_Terms = one_PDF_data[3][0][16:]
            Project_Cost_Center = one_PDF_data[5][0][22:] + " "
            if one_PDF_data[6][0] == 'MALAYSIA':
                Project_Cost_Center = one_PDF_data[5][0][22:] + " " + one_PDF_data[6][1]
            else:
                Project_Cost_Center = one_PDF_data[5][0][22:] + " " + one_PDF_data[6][0]
            Tracking_No = one_PDF_data[7][0][14:]
        
        Service_No = []

        Description = []
        pod = ''

        Delivery_Date = []

        Order_Qty = []

        UoM = []

        Unit_Price = []

        Total_Price = []


        while POSN < len(one_PDF_data):
            
            if len(one_PDF_data[POSN]) > 1:
                #不是最后一页，表格不封闭情况下找PO号
                if one_PDF_data[POSN][0][:3] == '000':
                    Service_No += [one_PDF_data[POSN][0]]   #找PO号
                    
                    #找描述
                    pod = one_PDF_data[POSN][1]   
                    a = 1
                    
                    """
                    while one_PDF_data[POSN + a][1] != '' and one_PDF_data[POSN + a][1] != 'Non-SST Registered Supplier Purchases 0%' and one_PDF_data[POSN + a + 1][0] != '':
                        pod += one_PDF_data[POSN + a][1]
                        print(pod)
                        a = a + 1
                    """

                    while one_PDF_data[POSN + a + 1][0] != '':
                        if len(one_PDF_data[POSN + a]) > 1:
                            if one_PDF_data[POSN + a][1] != 'Non-SST Registered Supplier Purchases 0%':
                                if one_PDF_data[POSN + a][1] != '':
                                    pod += one_PDF_data[POSN + a][1]
                        a = a + 1


                    Description += [pod]

                    #找Delivery Date
                    Delivery_Date += [one_PDF_data[POSN][2]]

                    #找Order_Qty
                    Order_Qty += [one_PDF_data[POSN][3]]

                    #找UoM
                    UoM += [one_PDF_data[POSN][4]]

                    #找Unit_Price
                    Unit_Price += [one_PDF_data[POSN][5]]

                    #找Total_Price
                    Total_Price += [one_PDF_data[POSN + 1][len(one_PDF_data[POSN + 1]) - 1]]






                #最后一页，表格封闭情况下找PO号
                elif one_PDF_data[POSN][1] != None:
                    if one_PDF_data[POSN][1][:3] == '000':
                        Service_No += [one_PDF_data[POSN][1]]

                        #找描述
                        pod = one_PDF_data[POSN][2]   
                        a = 1
                        while one_PDF_data[POSN + a][2] != '' and one_PDF_data[POSN + a][2] != 'Non-SST Registered Supplier Purchases 0%':
                            p = ' ' + one_PDF_data[POSN + a][2]
                            pod += p
                            a = a + 1
                        Description += [pod]

                        #找Delivery Date
                        Delivery_Date += [one_PDF_data[POSN][3]]

                        #找Order_Qty
                        Order_Qty += [one_PDF_data[POSN][4]]

                        #找UoM
                        UoM += [one_PDF_data[POSN][5]]

                        #找Unit_Price
                        Unit_Price += [one_PDF_data[POSN][6]]

                        #找Total_Price
                        Total_Price += [one_PDF_data[POSN + 1][len(one_PDF_data[POSN + 1]) - 1]]



            POSN = POSN + 1

        Form_Data += [Service_No, Description, Delivery_Date, Order_Qty, UoM, Unit_Price, Total_Price]
        infoInLine = [PO_No, PO_Date, Contract_No, Payment_Terms, Project_Cost_Center, Tracking_No, Form_Data] 
        Processed_Data += [infoInLine]
        i = i + 1

        print('\r' + '已处理' + str(i) + '/' + str(len(filesName)), end='', flush=True)
    print('\n处理完毕')
    return(Processed_Data)


        
"""------------打这儿起，底下就不让放函数了啊！------------"""
        

os.mkdir('fixed')

#读取input文件夹内文件的文件名
print("正在读取input文件夹")
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


#上面那一坨就是从文件读取数据到内存，能运行，不需要知道怎么运行的，别碰就行了





raw_data = get_raw_info()

all_data = Preprocessing_Data(raw_data)


shutil.rmtree('fixed')


print("开始生成.xlsx文件")
#用当前时间做文件名
xlsxName = time.strftime('%Y%m%d%H%M%S', time.localtime()) + ".xlsx"

#生成空文件
workbook = xlsxwriter.Workbook(xlsxName)
#生成空工作表
worksheet = workbook.add_worksheet('PO Info')
worksheet.set_column('A:A', 12)
worksheet.set_column('B:B', 12)
worksheet.set_column('C:C', 12)
worksheet.set_column('D:D', 25)
worksheet.set_column('E:E', 55)
worksheet.set_column('F:F', 12)
worksheet.set_column('G:G', 12)
worksheet.set_column('H:H', 60)
worksheet.set_column('I:I', 15)
worksheet.set_column('J:J', 10)
worksheet.set_column('K:K', 4)
worksheet.set_column('L:L', 12)
worksheet.set_column('M:M', 12)

#向工作表输入数据
worksheet.write(0,0,"PO Number")
worksheet.write(0,1,"PO Date")
worksheet.write(0,2,"Contract No")
worksheet.write(0,3,"Payment Terms")
worksheet.write(0,4,"Project/Cost Center")
worksheet.write(0,5,"Tracking No")
worksheet.write(0,6,"Service No")
worksheet.write(0,7,"Description")
worksheet.write(0,8,"Delivery Date")
worksheet.write(0,9,"Order Qty")
worksheet.write(0,10,"UoM")
worksheet.write(0,11,"Unit Price")
worksheet.write(0,12,"Total Price")

i = 1
a = 0
while a < len(all_data):
    worksheet.write(i,0,all_data[a][0])
    worksheet.write(i,1,all_data[a][1])
    worksheet.write(i,2,all_data[a][2])
    worksheet.write(i,3,all_data[a][3])
    worksheet.write(i,4,all_data[a][4])
    worksheet.write(i,5,all_data[a][5])
    q = 0
    c = i
    while q < len(all_data[a][6][0]):
        worksheet.write(c,6,all_data[a][6][0][q])
        worksheet.write(c,7,all_data[a][6][1][q])
        worksheet.write(c,8,all_data[a][6][2][q])
        worksheet.write(c,9,all_data[a][6][3][q])
        worksheet.write(c,10,all_data[a][6][4][q])
        worksheet.write(c,11,all_data[a][6][5][q])
        worksheet.write(c,12,all_data[a][6][6][q])
        q = q + 1
        c = c + 1
    i = c



    print('\r' + '正在写入数据' + str(a + 1) + '/' + str(len(filesName)), end='', flush=True)
    i = i + 1
    a = a + 1

print('输出文件: ', xlsxName)

workbook.close()