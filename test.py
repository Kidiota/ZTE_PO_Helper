# 提取pdf文字
import pdfplumber
#with pdfplumber.open("c:\\Users\\DuangRui\\Documents\\GitHubProjects\\ZTE_PO_Helper\\pdfs\\01.pdf") as pdf:
with pdfplumber.open("pdfs\\16.pdf") as pdf:
    page01 = pdf.pages[-1] #指定页码
    text = page01.extract_text()#提取文本
    print(text)
    print("data type: ", type(text))
    



